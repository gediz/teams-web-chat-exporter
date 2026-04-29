import { defineBackground } from 'wxt/sandbox';
import { createBadgeManager } from '../utils/badge';
import { isTeamsUrl } from '../utils/teams-urls';
import {
  buildAndDownload,
  buildAndDownloadBundle,
  buildAndDownloadBundlesZip,
  buildAndDownloadZip,
  buildOneChatForBundle,
  pickBundleFolderName,
  type BundleEntry,
  type BundleEmpty,
  type BundleFailure,
} from '../background/download';
import { revokeDownloadUrl, textToDownloadUrl } from '../background/builders';
import { ACTIVE_EXPORTS_STORAGE_KEY, FIRST_INSTALL_STORAGE_KEY, appendHistoryEntry } from '../utils/options';
import type { BackgroundIncomingMessage } from '../types/messaging';
import type {
  ActiveExportInfo,
  BuildOptions,
  ExportMessage,
  ExportMeta,
  ExportStatusPayload,
  HistoryEntry,
  ScrapeOptions,
  ScrapeResult,
} from '../types/shared';

/* eslint-disable @typescript-eslint/no-explicit-any */
// Typed globals for Firefox builds
declare const browser: typeof chrome | undefined;

// ===== service-worker.js (WXT version) =====
export default defineBackground(() => {
// Browser API compatibility for Firefox
const runtime = typeof browser !== 'undefined' ? browser.runtime : chrome.runtime;
const tabs = typeof browser !== 'undefined' ? browser.tabs : chrome.tabs;
// Firefox MV2 uses browserAction, Chrome MV3 uses action
const action = typeof browser !== 'undefined'
    ? (browser.action || browser.browserAction)
    : chrome.action;
const downloads = typeof browser !== 'undefined' ? browser.downloads : chrome.downloads;
const scripting = typeof browser !== 'undefined' ? browser.scripting : chrome.scripting;
const storage = typeof browser !== 'undefined' ? browser.storage : chrome.storage;
const badge = createBadgeManager(action);
const { reset: resetBadge, updateForStatus: updateBadgeForStatus, updateForProgress: updateBadgeForProgress } = badge;
const isFirefox = typeof browser !== 'undefined' && navigator.userAgent.includes('Firefox');

function log(...a: unknown[]) { try { console.log("[Teams Exporter SW]", ...a) } catch { } }
log("boot");

// Ask the content script for the tab's current Teams conversation id. The
// content script uses IndexedDB/DOM/URL to resolve it — the address-bar
// URL alone omits the id in Teams v2. Returns undefined when the page
// isn't on a conversation, when the content script hasn't loaded yet, or
// when the extraction fails.
async function getConvIdForTab(tabId: number): Promise<string | undefined> {
    try {
        const resp = await sendMessageToTab(tabId, { type: 'GET_CONV_ID' });
        return resp?.convId || undefined;
    } catch {
        return undefined;
    }
}

// Append a row to the persisted export history. The popup reads this on
// open to render the History page and to compute whether to show the
// "new entry" dot on the history icon. Writing from the background (not
// the popup) means the entry is recorded even when the popup was closed
// during the export.
async function persistHistoryEntry(entry: HistoryEntry): Promise<void> {
    await appendHistoryEntry(storage, entry);
}

// Best-effort UUID. Service worker has crypto.randomUUID() in MV3 / Firefox.
function makeEntryId(): string {
    try {
        // crypto.randomUUID is available in modern Chromium SW + Firefox WebExt.
        const c = (globalThis as { crypto?: { randomUUID?: () => string } }).crypto;
        if (c?.randomUUID) return c.randomUUID();
    } catch { /* fall through */ }
    // Fallback: timestamp + random. Collision-prone only under absurd load.
    return `${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
}

// downloads.download() returns the item's ID before the file is actually
// written to disk. Calling downloads.open() in that window would fail for
// larger exports (blob-URL writes can take a few hundred ms for 50MB+).
// Wait up to `timeoutMs` for the item to reach state='complete'. Returns
// true when ready, false on timeout or interruption.
function waitForDownloadComplete(id: number, timeoutMs = 10_000): Promise<boolean> {
    return new Promise(resolve => {
        let done = false;
        const finish = (ok: boolean) => {
            if (done) return;
            done = true;
            try { downloads.onChanged.removeListener(onChange); } catch { /* noop */ }
            clearTimeout(timer);
            resolve(ok);
        };
        const onChange = (delta: chrome.downloads.DownloadDelta) => {
            if (delta.id !== id) return;
            if (delta.state?.current === 'complete') finish(true);
            else if (delta.state?.current === 'interrupted') finish(false);
        };
        const timer = setTimeout(() => finish(false), timeoutMs);
        try { downloads.onChanged.addListener(onChange); } catch { /* noop */ }
        // Fast path — the item may already be complete by the time we check.
        Promise.resolve(downloads.search({ id }))
            .then(results => {
                const item = Array.isArray(results) ? results[0] : undefined;
                if (item?.state === 'complete') finish(true);
                else if (item?.state === 'interrupted') finish(false);
            })
            .catch(() => { /* leave the listener to resolve it */ });
    });
}

runtime.onInstalled.addListener((details) => {
    log("onInstalled", details?.reason);
    resetBadge();
    // Stamp the first-install timestamp if not already present — regardless
    // of reason. We deliberately do NOT restrict to reason='install':
    // users who had the extension before this feature shipped see reason=
    // 'update' and would otherwise never get a stamp, which would keep the
    // review prompt permanently invisible for them. Stamping on first-seen
    // undercounts their real install age by whatever time has elapsed
    // since their actual first install, but that's fine — the 7-day gate
    // becomes "7 days since this feature reached them" which is still a
    // reasonable heuristic for "has had time to form an opinion".
    storage.local.get(FIRST_INSTALL_STORAGE_KEY).then((stored) => {
        if (!stored?.[FIRST_INSTALL_STORAGE_KEY]) {
            return storage.local.set({ [FIRST_INSTALL_STORAGE_KEY]: Date.now() });
        }
    }).catch(() => { /* nothing critical depends on this */ });
});
runtime.onStartup?.addListener(() => {
    log("onStartup");
    resetBadge();
});

tabs.onUpdated.addListener((tabId: number, changeInfo: chrome.tabs.TabChangeInfo, tab?: chrome.tabs.Tab) => {
    const nextUrl = changeInfo.url ?? tab?.url;
    if (changeInfo.status === 'loading' && isTeamsUrl(nextUrl)) {
        activeExports.delete(tabId); void persistActiveExports();
        resetBadge();
    }
});

const activeExports = new Map<number, ActiveExportInfo>(); // tabId -> { startedAt, lastStatus }
// TERMINAL_PHASES: 'complete' = success, 'error' = failure, 'empty' = no data found (not a failure),
// 'cancelled' = user-stopped (not a failure)
const TERMINAL_PHASES = new Set(['complete', 'error', 'empty', 'cancelled']);

function updateActiveExport(tabId: number, patch: Partial<ActiveExportInfo> = {}) {
    if (tabId == null) return;
    const prev = activeExports.get(tabId) || {};
    const next: ActiveExportInfo = { ...prev, ...patch };
    activeExports.set(tabId, next);
    void persistActiveExports();
    return next;
}

// Snapshot the in-memory activeExports Map to chrome.storage.local so
// the popup can pre-hydrate the export-button state on mount. Writes
// are fire-and-forget — storage errors aren't actionable and the
// worst case is the popup falls back to the async GET_EXPORT_STATUS
// round-trip. Debouncing isn't worth it: updates are driven by user
// export start/stop and progress events that already throttle.
async function persistActiveExports() {
    try {
        const snapshot: Record<string, ActiveExportInfo> = {};
        for (const [tabId, info] of activeExports) snapshot[String(tabId)] = info;
        await storage.local.set({ [ACTIVE_EXPORTS_STORAGE_KEY]: snapshot });
    } catch { /* ignore */ }
}

const sendMessageToTab = (tabId: number, msg: unknown) => new Promise<any>((resolve, reject) => {
    tabs.sendMessage(tabId, msg, (resp) => {
        const err = runtime.lastError;
        if (err) {
            reject(new Error(err.message || 'Failed to reach tab context'));
            return;
        }
        resolve(resp);
    });
});

async function ensureContentScript(tabId: number) {
    try {
        const pong = await sendMessageToTab(tabId, { type: 'PING' });
        if (pong?.ok) return;
    } catch (_) {
        // fallback to injection
    }
    // WXT bundles the content script at content-scripts/content.js
    // (matches the manifest's content_scripts[0].js path).
    type FrameResult = chrome.scripting.InjectionResult & { error?: { message?: string } };
    let injectResult: FrameResult[] | undefined;
    try {
        injectResult = (await scripting.executeScript({ target: { tabId, allFrames: true }, files: ['content-scripts/content.js'] })) as FrameResult[];
    } catch (e) {
        log('ensureContentScript: executeScript threw:', e);
        throw new Error(`Could not inject content script: ${(e as Error)?.message || String(e)}`);
    }
    // executeScript returned a per-frame result. Surface frame-level
    // errors when the top frame failed to load the script. Common
    // causes: sandboxed iframe layout, page-CSP block, or a URL that
    // slipped past host_permissions. The `error` field is present
    // on recent Chrome but missing from @types/chrome (hence cast).
    const topFrameError = injectResult?.find(r => r.frameId === 0)?.error;
    if (topFrameError) {
        log('ensureContentScript: top-frame injection error:', topFrameError);
        throw new Error(`Content script injection failed in top frame: ${topFrameError.message || String(topFrameError)}`);
    }
    // Retry the post-injection PING a few times. Listener registration
    // in the freshly-injected content script should be synchronous,
    // but some Firefox + cold-tab paths take a tick or two longer
    // than executeScript's resolve promise. 5 × 50 ms = 250 ms
    // ceiling. Cheap insurance against a transient race.
    let lastErr: unknown = null;
    for (let attempt = 0; attempt < 5; attempt++) {
        if (attempt > 0) await new Promise(r => setTimeout(r, 50));
        try {
            const pong = await sendMessageToTab(tabId, { type: 'PING' });
            if (pong?.ok) return;
        } catch (e) {
            lastErr = e;
        }
    }
    const tail = lastErr ? `: ${(lastErr as Error)?.message || String(lastErr)}` : '';
    throw new Error(`Content script did not respond after injection${tail}`);
}

// Port-based streaming: content script sends scrape results in chunks to bypass 64MiB limit
const pendingScrapes = new Map<string, {
    resolve: (result: ScrapeResult) => void;
    reject: (err: Error) => void;
}>();

runtime.onConnect.addListener((port: chrome.runtime.Port) => {
    if (!port.name.startsWith('scrape-result:')) return;
    const requestId = port.name.slice('scrape-result:'.length);
    const pending = pendingScrapes.get(requestId);
    if (!pending) { log('port connected but no pending scrape for', requestId); port.disconnect(); return; }

    log('streaming port connected for', requestId);
    const messages: ExportMessage[] = [];
    let meta: ExportMeta = {};
    let batches = 0;

    port.onMessage.addListener((msg: any) => {
        if (msg.type === 'meta') {
            meta = msg.meta || {};
            log('received meta:', { title: meta.title, avatars: Object.keys(meta.avatars || {}).length });
        } else if (msg.type === 'messages') {
            const batch = Array.isArray(msg.messages) ? msg.messages : [];
            messages.push(...batch);
            batches++;
            log(`received batch ${batches}: ${batch.length} messages (total: ${messages.length})`);
        } else if (msg.type === 'done') {
            log(`streaming complete: ${messages.length} messages in ${batches} batches`);
            pendingScrapes.delete(requestId);
            pending.resolve({ messages, meta });
        } else if (msg.type === 'error') {
            log('streaming error:', msg.error);
            pendingScrapes.delete(requestId);
            pending.reject(new Error(msg.error || 'Streaming error'));
        }
    });

    port.onDisconnect.addListener(() => {
        if (pendingScrapes.has(requestId)) {
            log('port disconnected unexpectedly for', requestId, `(received ${messages.length} messages so far)`);
            pendingScrapes.delete(requestId);
            pending.reject(new Error('Content script disconnected unexpectedly'));
        }
    });
});

async function requestScrape(tabId: number, options: ScrapeOptions): Promise<ScrapeResult> {
    const requestId = `${tabId}-${Date.now()}`;

    // Register pending entry BEFORE sending message to avoid race condition:
    // the port may connect before sendMessageToTab's promise resolves.
    let streamResolve: (r: ScrapeResult) => void;
    let streamReject: (e: Error) => void;
    const streamPromise = new Promise<ScrapeResult>((resolve, reject) => {
        streamResolve = resolve;
        streamReject = reject;
    });
    // Attach a no-op catch so an early rejection (e.g. STOP_EXPORT firing
    // before the content script even responds, while we're still in the
    // non-streaming branch and never await streamPromise) doesn't surface as
    // an "Uncaught (in promise)" in the service worker. The downstream `await`
    // in the streaming branch still observes the rejection normally.
    streamPromise.catch(() => { /* noop */ });
    pendingScrapes.set(requestId, { resolve: streamResolve!, reject: streamReject! });

    try {
        const res = await sendMessageToTab(tabId, { type: 'SCRAPE_TEAMS', options, requestId });
        if (!res) {
            pendingScrapes.delete(requestId);
            throw new Error('No response from content script');
        }
        if (res.error) {
            pendingScrapes.delete(requestId);
            throw new Error(res.error);
        }
        // Content script aborted before it had results to stream. Surface this
        // as a 'cancelled' rejection so handleExportWithScrape's catch can
        // distinguish it from a normal "0 messages found" outcome (which would
        // wrongly show the EMPTY_RESULTS error in the popup).
        if (res.cancelled) {
            pendingScrapes.delete(requestId);
            throw new Error('cancelled');
        }

        if (res.streaming) {
            log('awaiting streamed scrape results for request', requestId);
            // Port should connect within seconds; 30s safety net
            setTimeout(() => {
                if (pendingScrapes.has(requestId)) {
                    pendingScrapes.delete(requestId);
                    streamReject!(new Error('Timed out waiting for streamed scrape results'));
                }
            }, 30_000);
            return streamPromise;
        }

        // Legacy non-streaming fallback
        pendingScrapes.delete(requestId);
        return res;
    } catch (err) {
        pendingScrapes.delete(requestId);
        throw err;
    }
}

async function checkContext(tabId: number, options: ScrapeOptions) {
    const target = options?.exportTarget === 'team' ? 'team' : 'chat';
    return sendMessageToTab(tabId, { type: 'CHECK_CHAT_CONTEXT', target });
}

function defaultContextError(options: ScrapeOptions) {
    return options?.exportTarget === 'team'
        ? 'Open a team channel before exporting.'
        : 'Open a chat conversation before exporting.';
}

// Debug: track broadcastStatus call rate so we can correlate the popup
// white-screen symptom with broadcast-storm load. Logs every 2s while
// activity is non-zero. Counter resets after each report.
let __broadcastCount = 0;
let __broadcastWindowStart = Date.now();
setInterval(() => {
    if (__broadcastCount === 0) return;
    const elapsedSec = (Date.now() - __broadcastWindowStart) / 1000;
    const rate = (__broadcastCount / elapsedSec).toFixed(1);
    log(`broadcastStatus rate ${rate}/s (${__broadcastCount} in ${elapsedSec.toFixed(1)}s)`);
    __broadcastCount = 0;
    __broadcastWindowStart = Date.now();
}, 2000);

function broadcastStatus(payload: ExportStatusPayload) {
    __broadcastCount += 1;
    let enriched = { ...payload };
    const tabId = payload?.tabId;
    if (tabId != null) {
        const phase = payload?.phase;
        let info;
        if (phase) {
            const record = { ...payload };
            if (TERMINAL_PHASES.has(phase)) {
                info = updateActiveExport(tabId, { lastStatus: record, phase, completedAt: Date.now() });
            } else {
                info = updateActiveExport(tabId, { lastStatus: record, phase });
            }
        } else {
            info = updateActiveExport(tabId, { lastStatus: { ...payload } });
        }
        const startedAt = info?.startedAt;
        if (startedAt && enriched.startedAt == null) {
            enriched = { ...enriched, startedAt };
        }
    }
    // Firefox compatibility: wrap in try-catch since sendMessage may not return a Promise
    try {
        const msgPromise = runtime.sendMessage({ type: 'EXPORT_STATUS', ...enriched });
        if (msgPromise && msgPromise.catch) msgPromise.catch(() => { });
    } catch (e) {
        // Ignore errors when popup is closed
    }
    updateBadgeForStatus(payload);
}

function handleBuildAndDownloadMessage(msg: any, sendResponse: (res: any) => void) {
    (async () => {
        try {
            const result = await buildAndDownload({ downloads, isFirefox }, msg.data || {});
            sendResponse(result);
        } catch (err: any) {
            sendResponse({ error: err?.message || String(err) });
        }
    })();
}

type BuildStep = (
    scrapeRes: ScrapeResult,
    buildOptions: BuildOptions,
    tabId: number,
) => Promise<{ filename?: string; id?: number }>;

function handleExportWithScrape(
    msg: any,
    sendResponse: (res: any) => void,
    buildStep: BuildStep,
) {
    const data = msg.data || {};
    const tabId = data.tabId;
    if (typeof tabId !== 'number') {
        sendResponse({ error: 'Missing tabId for export request' });
        return;
    }
    // activeExports is dual-purpose: it tracks in-flight exports AND the
    // last-status cache (so GET_EXPORT_STATUS can answer after completion).
    // After cancel/complete/error, broadcastStatus re-inserts the tab with
    // the terminal phase, so plain .has() would wrongly block a retry.
    // Guard against *in-flight* only by checking the phase.
    const existing = activeExports.get(tabId);
    if (existing?.phase && !TERMINAL_PHASES.has(existing.phase)) {
        sendResponse({ error: 'An export is already running for this tab' });
        return;
    }

    const scrapeOptions = data.scrapeOptions || {};
    const buildOptions = data.buildOptions || {};

    (async () => {
        let startedAt;
        try {
            await ensureContentScript(tabId);
            const ctx = await checkContext(tabId, scrapeOptions);
            if (!ctx?.ok) {
                const message = ctx?.reason || defaultContextError(scrapeOptions);
                sendResponse({ error: message });
                return;
            }

            startedAt = Date.now();
            updateActiveExport(tabId, { startedAt, phase: 'starting', lastStatus: undefined });
            broadcastStatus({ tabId, phase: 'starting', startedAt });

            broadcastStatus({ tabId, phase: 'scrape:start' });
            const scrapeRes = await requestScrape(tabId, scrapeOptions);
            const totalMessages = Array.isArray(scrapeRes.messages) ? scrapeRes.messages.length : 0;
            broadcastStatus({ tabId, phase: 'scrape:complete', messages: totalMessages });

            if (totalMessages === 0) {
                const message = 'No messages found for the selected range.';
                broadcastStatus({ tabId, phase: 'empty', message });
                sendResponse({ error: message, code: 'EMPTY_RESULTS' });
                return;
            }

            const buildRes = await buildStep(scrapeRes, buildOptions, tabId);

            // Wait until the download has actually settled on disk.
            // chrome.downloads.download() resolves with the id as soon as
            // the download is queued — the file may still be writing, OR
            // the user may have hit Cancel in the Save As dialog (in which
            // case the download immediately enters state='interrupted'
            // without any file ever being written). If we broadcast
            // 'complete' and persist a history entry before checking, the
            // popup ends up showing a row that points at a non-existent
            // file, and Open fails with "An unexpected error occurred."
            const downloadOk = buildRes.id != null
                ? await waitForDownloadComplete(buildRes.id, 30_000)
                : true;
            if (!downloadOk) {
                log('export download did not complete (cancelled or interrupted) for id', buildRes.id);
                broadcastStatus({ tabId, phase: 'cancelled' });
                sendResponse({ cancelled: true });
                return;
            }

            broadcastStatus({
                tabId,
                phase: 'complete',
                filename: buildRes.filename,
                downloadId: buildRes.id,
                // Forward the setting so the popup can try auto-open from
                // its own context (see auto-action comment below).
                afterExport: buildOptions?.afterExport,
            });

            // Append a row to the persisted export history. Reading happens
            // from the popup's HistoryPage; the entry is written here (not
            // popup-side) so it gets recorded even when the popup was
            // closed during the export.
            const completeStartedAt = activeExports.get(tabId)?.startedAt;
            const completeElapsedMs = completeStartedAt
                ? Math.max(0, Date.now() - completeStartedAt)
                : 0;
            const completeFname = buildRes.filename || '';
            // Prefer the convId the content script already resolved during
            // the scrape; only re-query the tab if it's missing.
            const metaConvId = typeof scrapeRes.meta?.conversationId === 'string'
                ? scrapeRes.meta.conversationId
                : undefined;
            const completeConvId = metaConvId ?? await getConvIdForTab(tabId);
            const completeTitle = typeof scrapeRes.meta?.title === 'string'
                ? scrapeRes.meta.title
                : undefined;
            // Write both `formats` (canonical, multi) and `format` (singular,
            // back-compat for code paths still reading the old field). The
            // singular value reflects "what the badge should show" — for
            // bundle exports it's left undefined so the History badge can
            // pick the bundle treatment instead.
            const completeFormats = Array.isArray(buildOptions?.formats) ? buildOptions.formats : undefined;
            const singleFormat = completeFormats && completeFormats.length === 1 ? completeFormats[0] : undefined;
            // Promote 'success' to 'partial' when the scrape signalled an
            // incomplete-data condition. The file IS on disk and Open /
            // Show still work, but History renders an amber badge so the
            // user can tell this row apart from a clean export.
            const completePartial = (scrapeRes.meta as { partial?: { reason: 'network' | 'truncation' } } | undefined)?.partial;
            await persistHistoryEntry({
                id: makeEntryId(),
                tabId,
                kind: completePartial ? 'partial' : 'success',
                partialReason: completePartial?.reason,
                convId: completeConvId,
                downloadId: buildRes.id,
                filename: completeFname || undefined,
                title: completeTitle,
                formats: completeFormats,
                format: singleFormat,
                isZip: completeFname.toLowerCase().endsWith('.zip'),
                messageCount: totalMessages,
                elapsedMs: completeElapsedMs,
                savedAt: Date.now(),
            });

            // Auto-action dispatch:
            //   'show' — downloads.show() doesn't need a user gesture, so
            //            the SW handles it here. We don't have to wait for
            //            the download to be on disk because the
            //            waitForDownloadComplete above already gated us.
            //   'manual' — no auto-action; the user clicks the Open / Show
            //              buttons on the History page when ready.
            const after = buildOptions?.afterExport;
            log('after-export setting is', after, 'downloadId is', buildRes.id);
            if (buildRes.id != null && after === 'show') {
                try {
                    await downloads.show(buildRes.id);
                    log('auto-show: downloads.show resolved for id', buildRes.id);
                } catch (e: any) {
                    log('auto-show failed — error:', e?.message || String(e));
                }
            }

            sendResponse({ ok: true, filename: buildRes.filename, downloadId: buildRes.id });
        } catch (err: any) {
            const message = err?.message || String(err);
            // Cancellation flows here as a rejection of the streaming promise
            // with message 'cancelled'. Treat it as a successful stop so the
            // popup doesn't show a red error banner.
            if (message === 'cancelled') {
                sendResponse({ cancelled: true });
            } else {
                broadcastStatus({ tabId, phase: 'error', error: message });
                sendResponse({ error: message });
            }
        } finally {
            activeExports.delete(tabId); void persistActiveExports();
        }
    })();
}

function handleStartExportMessage(msg: any, sendResponse: (res: any) => void) {
    handleExportWithScrape(msg, sendResponse, async (scrapeRes, buildOptions, tabId) => {
        // Sanitize the formats array. The popup always sends one, but be
        // defensive — a corrupt/legacy payload reaching the SW would
        // otherwise crash the build switch below.
        const validFormats = ['json', 'csv', 'html', 'txt', 'pdf'] as const;
        const formats = (buildOptions.formats || []).filter((f): f is typeof validFormats[number] =>
            (validFormats as readonly string[]).includes(f),
        );
        if (!formats.length) formats.push('json');
        const downloadImages = Boolean(buildOptions.downloadImages);
        const deps = {
            downloads,
            isFirefox,
            onStatus: (payload: Record<string, unknown>) => broadcastStatus({ ...payload, tabId }),
        };
        // PDF knobs — plumb through to every build path that might
        // produce a PDF. Safe to pass even when PDF isn't selected;
        // the builders ignore them.
        const pdfKnobs = {
            pdfPageSize: buildOptions.pdfPageSize,
            pdfBodyFontSize: buildOptions.pdfBodyFontSize,
            pdfShowPageNumbers: buildOptions.pdfShowPageNumbers,
            pdfIncludeAvatars: buildOptions.pdfIncludeAvatars,
        };
        const commonOpts = {
            messages: scrapeRes.messages || [],
            meta: scrapeRes.meta || {},
            embedAvatars: Boolean(buildOptions.embedAvatars),
            downloadImages,
            ...pdfKnobs,
        };
        // 2+ formats -> always bundle.zip. The bundle path doesn't honor
        // saveAs because zips force a Save As anyway via downloads.show
        // semantics (the file would conflict otherwise).
        const avatarMode = buildOptions.avatarMode === 'files' ? 'files' : 'inline';
        if (formats.length >= 2) {
            return buildAndDownloadBundle(deps, { ...commonOpts, formats, avatarMode });
        }
        const format = formats[0];
        // HTML goes into a .zip when EITHER inline images are on OR the
        // user chose 'files' avatar mode (both need the zip's folder
        // structure). Without either, single HTML stays inline.
        const wantFiles = avatarMode === 'files' && Boolean(buildOptions.embedAvatars);
        if (format === 'html' && (downloadImages || wantFiles)) {
            return buildAndDownloadZip(deps, { ...commonOpts, avatarMode });
        }
        return buildAndDownload(deps, {
            ...commonOpts,
            format,
            saveAs: buildOptions.saveAs !== false,
        });
    });
}

// Multi-chat bundle export. Loops over the requested conversation ids
// serially (parallel scrapes risk Teams throttling), runs the existing
// per-chat scrape pipeline, and packs everything into one outer zip
// with FAILURES.txt for any chat that errored. The loop is interruptible:
// STOP_EXPORT both aborts the current scrape and short-circuits the
// remaining iterations via bundleStops.
function handleStartBundleExportMessage(msg: any, sendResponse: (res: any) => void) {
    const data = msg.data || {};
    const tabId = data.tabId;
    const list: Array<{ id: string; title: string }> = Array.isArray(data.conversations) ? data.conversations : [];
    const buildOptions: BuildOptions = data.buildOptions || {};
    const baseScrapeOptions: ScrapeOptions = data.scrapeOptions || {};

    if (typeof tabId !== 'number') {
        sendResponse({ error: 'Missing tabId for bundle export request' });
        return;
    }
    if (!list.length) {
        sendResponse({ error: 'No conversations selected for bundle export' });
        return;
    }
    const existing = activeExports.get(tabId);
    if (existing?.phase && !TERMINAL_PHASES.has(existing.phase)) {
        sendResponse({ error: 'An export is already running for this tab' });
        return;
    }

    const validFormats = ['json', 'csv', 'html', 'txt', 'pdf'] as const;
    const formats = (buildOptions.formats || []).filter((f): f is typeof validFormats[number] =>
        (validFormats as readonly string[]).includes(f),
    );
    if (!formats.length) formats.push('json');
    const downloadImages = Boolean(buildOptions.downloadImages);
    const embedAvatars = Boolean(buildOptions.embedAvatars);
    const avatarMode: 'inline' | 'files' = buildOptions.avatarMode === 'files' ? 'files' : 'inline';

    bundleStops.delete(tabId);

    (async () => {
        const startedAt = Date.now();
        updateActiveExport(tabId, { startedAt, phase: 'starting', lastStatus: undefined });
        broadcastStatus({ tabId, phase: 'starting', startedAt });

        try {
            await ensureContentScript(tabId);
        } catch (e: any) {
            const message = e?.message || String(e);
            broadcastStatus({ tabId, phase: 'error', error: message });
            sendResponse({ error: message });
            activeExports.delete(tabId); void persistActiveExports();
            return;
        }

        const totalChats = list.length;
        const usedFolderNames = new Set<string>();
        const entries: BundleEntry[] = [];
        const failures: BundleFailure[] = [];
        const noHistory: BundleEmpty[] = [];

        for (let i = 0; i < list.length; i++) {
            if (bundleStops.has(tabId)) break;

            const conv = list[i];
            const folderName = pickBundleFolderName(conv.title || conv.id, usedFolderNames);
            const currentChat = i + 1;

            const bundleCtx = {
                bundleCurrentChat: currentChat,
                bundleTotalChats: totalChats,
                bundleChatName: folderName,
                bundleSuccessCount: entries.length,
                bundleFailedCount: failures.length,
            };

            broadcastStatus({ tabId, phase: 'scrape:start', ...bundleCtx });

            const perChatScrape: ScrapeOptions = {
                ...baseScrapeOptions,
                conversationId: conv.id,
                conversationTitle: conv.title || null,
                noDomFallback: true,
            };

            let scrapeRes: ScrapeResult;
            try {
                scrapeRes = await requestScrape(tabId, perChatScrape);
            } catch (e: any) {
                const reason = e?.message || String(e);
                if (reason === 'cancelled') {
                    bundleStops.add(tabId);
                    break;
                }
                failures.push({ folderName, conversationId: conv.id, reason });
                continue;
            }

            const totalMessages = Array.isArray(scrapeRes.messages) ? scrapeRes.messages.length : 0;
            broadcastStatus({ tabId, phase: 'scrape:complete', messages: totalMessages, ...bundleCtx });

            // 0-message API responses go to NO_HISTORY.txt at bundle root
            // rather than producing a per-chat folder of empty files. The
            // common case is Teams Free legacy Skype-imported 1:1s where
            // Microsoft never migrated the history into the consumer chat
            // backend — there is genuinely nothing to render, and a 7 MB
            // empty PDF per chat is just noise. Distinct from FAILURES.txt
            // because the API call succeeded; the chat is simply empty
            // server-side.
            if (totalMessages === 0) {
                noHistory.push({ folderName, conversationId: conv.id });
                continue;
            }

            try {
                broadcastStatus({ tabId, phase: 'build', messages: totalMessages, ...bundleCtx });
                const files = await buildOneChatForBundle({
                    messages: scrapeRes.messages || [],
                    meta: scrapeRes.meta || {},
                    formats,
                    embedAvatars,
                    downloadImages,
                    avatarMode,
                    pdfPageSize: buildOptions.pdfPageSize,
                    pdfBodyFontSize: buildOptions.pdfBodyFontSize,
                    pdfShowPageNumbers: buildOptions.pdfShowPageNumbers,
                    pdfIncludeAvatars: buildOptions.pdfIncludeAvatars,
                    onPdfProgress: (done, total) => {
                        broadcastStatus({
                            tabId,
                            phase: 'build',
                            messages: totalMessages,
                            messagesBuilt: done,
                            messagesTotal: total,
                            ...bundleCtx,
                            bundleSuccessCount: entries.length,
                            bundleFailedCount: failures.length,
                        });
                    },
                });
                entries.push({ folderName, files });
            } catch (e: any) {
                failures.push({ folderName, conversationId: conv.id, reason: e?.message || String(e) });
            }
        }

        const cancelled = bundleStops.has(tabId);
        bundleStops.delete(tabId);

        // Cancellation is a hard abort: no partial zip, no history row.
        // The user said stop; producing half a bundle would be confusing
        // and risks shipping output the user didn't sanity-check.
        if (cancelled) {
            broadcastStatus({ tabId, phase: 'cancelled' });
            sendResponse({ cancelled: true });
            activeExports.delete(tabId); void persistActiveExports();
            return;
        }

        if (!entries.length && !failures.length && !noHistory.length) {
            const message = 'No conversations produced output.';
            broadcastStatus({ tabId, phase: 'empty', message });
            sendResponse({ error: message, code: 'EMPTY_RESULTS' });
            activeExports.delete(tabId); void persistActiveExports();
            return;
        }

        // All-empty path: every chat the user picked returned 0 messages
        // (and nothing actually failed). Common case is bundling a set of
        // legacy Skype-imported 1:1s where Microsoft never migrated the
        // history. A zip-wrapped NO_HISTORY.txt is just an extra extract
        // step for a single text file, so save it directly. Marked as a
        // normal completed export (not a failure) since the API ran fine.
        if (entries.length === 0 && failures.length === 0 && noHistory.length > 0) {
            console.log(`[bundle-allempty] saving NO_HISTORY.txt directly (no zip): noHistory=${noHistory.length}`);
            const stamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19).replace('T', '_');
            const filename = `TeamsExport_bundle_${stamp}_NO_HISTORY.txt`;
            const lines = noHistory.map(e => `${e.folderName}\t${e.conversationId}`);
            const header = '# Chats with no retrievable message history. The API succeeded\n'
                + '# but returned 0 messages — most often legacy Skype-imported 1:1\n'
                + '# chats where Microsoft did not migrate the message history into\n'
                + '# the Teams Free chat backend. Not a failure; nothing to export.\n'
                + '# Columns: folder\tconversationId';
            const body = `${header}\n${lines.join('\n')}\n`;
            const url = textToDownloadUrl(body, 'text/plain');

            let downloadId: number | undefined;
            try {
                downloadId = await downloads.download({ url, filename, saveAs: true });
                setTimeout(() => revokeDownloadUrl(url), 60_000);
            } catch (e: any) {
                revokeDownloadUrl(url);
                const message = e?.message || String(e);
                broadcastStatus({ tabId, phase: 'error', error: message });
                sendResponse({ error: message });
                activeExports.delete(tabId); void persistActiveExports();
                return;
            }

            broadcastStatus({
                tabId,
                phase: 'complete',
                filename,
                downloadId,
                afterExport: buildOptions?.afterExport,
                bundleTotalChats: totalChats,
                bundleSuccessCount: 0,
                bundleFailedCount: 0,
            });

            const elapsedMs = Math.max(0, Date.now() - startedAt);
            await persistHistoryEntry({
                id: makeEntryId(),
                tabId,
                kind: 'success',
                downloadId,
                filename,
                title: `0 chats (${noHistory.length} empty)`,
                isZip: false,
                elapsedMs,
                savedAt: Date.now(),
            });

            const after = buildOptions?.afterExport;
            if (downloadId != null && after === 'show') {
                try { await downloads.show(downloadId); } catch { /* noop */ }
            }

            sendResponse({
                ok: true,
                filename,
                downloadId,
                totalChats,
                successChats: 0,
                failedChats: 0,
            });
            activeExports.delete(tabId); void persistActiveExports();
            return;
        }

        // Bundle where every chat failed: skip the .zip wrapper. We'd
        // be packaging a single ~1 KB FAILURES.txt inside an otherwise
        // empty zip — pointless extraction step for the user. Save the
        // text file directly and mark the history row as 'failed' so
        // it's visually distinct from a real bundle.
        if (entries.length === 0 && failures.length > 0) {
            console.log(`[bundle-allfail] saving FAILURES.txt directly (no zip): failures=${failures.length}`);
            const stamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19).replace('T', '_');
            const filename = `TeamsExport_bundle_${stamp}_FAILURES.txt`;
            const lines = failures.map(f => `${f.folderName}\t${f.conversationId}\t${f.reason}`);
            const header = '# Chats that failed to export. Columns: folder\tconversationId\treason';
            const body = `${header}\n${lines.join('\n')}\n`;
            const url = textToDownloadUrl(body, 'text/plain');

            let downloadId: number | undefined;
            try {
                downloadId = await downloads.download({ url, filename, saveAs: true });
                console.log(`[bundle-allfail] downloads.download resolved: id=${downloadId} filename=${filename}`);
                setTimeout(() => revokeDownloadUrl(url), 60_000);
            } catch (e: any) {
                console.log(`[bundle-allfail] downloads.download rejected: ${e?.message || String(e)}`);
                revokeDownloadUrl(url);
                const message = e?.message || String(e);
                broadcastStatus({ tabId, phase: 'error', error: message });
                sendResponse({ error: message });
                activeExports.delete(tabId); void persistActiveExports();
                return;
            }

            broadcastStatus({
                tabId,
                phase: 'complete',
                filename,
                downloadId,
                afterExport: buildOptions?.afterExport,
                bundleTotalChats: totalChats,
                bundleSuccessCount: 0,
                bundleFailedCount: failures.length,
            });

            const elapsedMs = Math.max(0, Date.now() - startedAt);
            await persistHistoryEntry({
                id: makeEntryId(),
                tabId,
                kind: 'failed',
                downloadId,
                filename,
                title: `0 chats (${failures.length} failed)`,
                isZip: false,
                elapsedMs,
                savedAt: Date.now(),
            });

            const after = buildOptions?.afterExport;
            if (downloadId != null && after === 'show') {
                try { await downloads.show(downloadId); } catch { /* noop */ }
            }

            sendResponse({
                ok: true,
                filename,
                downloadId,
                totalChats,
                successChats: 0,
                failedChats: failures.length,
            });
            activeExports.delete(tabId); void persistActiveExports();
            return;
        }

        try {
            broadcastStatus({
                tabId,
                phase: 'build',
                messages: 0,
                bundleCurrentChat: totalChats,
                bundleTotalChats: totalChats,
                bundleSuccessCount: entries.length,
                bundleFailedCount: failures.length,
            });
            const buildRes = await buildAndDownloadBundlesZip({ downloads, isFirefox }, entries, failures, noHistory);
            const downloadOk = buildRes.id != null
                ? await waitForDownloadComplete(buildRes.id, 60_000)
                : true;
            if (!downloadOk) {
                broadcastStatus({ tabId, phase: 'cancelled' });
                sendResponse({ cancelled: true });
                return;
            }

            broadcastStatus({
                tabId,
                phase: 'complete',
                filename: buildRes.filename,
                downloadId: buildRes.id,
                afterExport: buildOptions?.afterExport,
                bundleTotalChats: totalChats,
                bundleSuccessCount: entries.length,
                bundleFailedCount: failures.length,
            });

            const elapsedMs = Math.max(0, Date.now() - startedAt);
            await persistHistoryEntry({
                id: makeEntryId(),
                tabId,
                kind: 'success',
                downloadId: buildRes.id,
                filename: buildRes.filename,
                // Title carries chat count + failure marker. messageCount is
                // omitted because the per-chat counts don't aggregate into a
                // single meaningful number for a bundle row.
                title: `${entries.length} chats${failures.length ? ` (${failures.length} failed)` : ''}`,
                isZip: true,
                elapsedMs,
                savedAt: Date.now(),
            });

            const after = buildOptions?.afterExport;
            if (buildRes.id != null && after === 'show') {
                try { await downloads.show(buildRes.id); } catch { /* noop */ }
            }

            sendResponse({
                ok: true,
                filename: buildRes.filename,
                downloadId: buildRes.id,
                totalChats,
                successChats: entries.length,
                failedChats: failures.length,
            });
        } catch (e: any) {
            const message = e?.message || String(e);
            broadcastStatus({ tabId, phase: 'error', error: message });
            sendResponse({ error: message });
        } finally {
            activeExports.delete(tabId); void persistActiveExports();
        }
    })();
}

// Circuit breaker for sustained rate limits: after several consecutive exhausted
// retries, skip further retries for a cool-off window. Prevents a rate-limited
// export from stalling for minutes as each concurrent image burns its full retry
// budget. Resets on any successful fetch.
let rateLimitStreak = 0;
let rateLimitCoolOffUntil = 0;
const RATE_LIMIT_STREAK_THRESHOLD = 3;
const RATE_LIMIT_COOL_OFF_MS = 20_000;

// In-flight FETCH_BLOB controllers grouped by tabId. STOP_EXPORT aborts every
// fetch in the group so a stop frees connections + memory immediately rather
// than waiting up to ~30s per pending image for the request to time out on
// its own.
const inFlightFetchesByTab = new Map<number, Set<AbortController>>();

// Tabs whose export was cancelled. Used to drop late SCRAPE_PROGRESS events
// (e.g. the result of an image fetch that was already in flight when the
// user pressed stop) so they don't repaint the badge after we cleared it.
const cancelledTabs = new Set<number>();

// Tabs running a multi-chat bundle export whose loop should stop before
// the next iteration. STOP_EXPORT writes to this set as well as aborting
// the in-flight scrape; the bundle loop checks it between chats.
const bundleStops = new Set<number>();

function trackFetch(tabId: number | undefined, controller: AbortController) {
    if (tabId == null) return;
    let set = inFlightFetchesByTab.get(tabId);
    if (!set) { set = new Set(); inFlightFetchesByTab.set(tabId, set); }
    set.add(controller);
}

function untrackFetch(tabId: number | undefined, controller: AbortController) {
    if (tabId == null) return;
    const set = inFlightFetchesByTab.get(tabId);
    if (!set) return;
    set.delete(controller);
    if (set.size === 0) inFlightFetchesByTab.delete(tabId);
}

function abortFetchesForTab(tabId: number): number {
    const set = inFlightFetchesByTab.get(tabId);
    if (!set) return 0;
    let n = 0;
    for (const c of set) {
        try { c.abort(); n++; } catch { /* noop */ }
    }
    inFlightFetchesByTab.delete(tabId);
    return n;
}

// Background-mediated fetch for credentialed cross-origin requests (Firefox content
// scripts can't send page cookies cross-origin, but background scripts can with
// matching host_permissions). Retries transient errors (network, 429, 5xx) with
// exponential backoff; treats 4xx (except 429) and size-rejects as permanent.
async function handleFetchBlobMessage(
    msg: { url: string; bearerToken?: string; maxBytes?: number; minBytes?: number },
    sendResponse: (resp: unknown) => void,
    tabId?: number,
) {
    const maxBytes = msg.maxBytes ?? 5 * 1024 * 1024;
    const minBytes = msg.minBytes ?? 100;
    const MAX_ATTEMPTS = 3;
    const sleep = (ms: number) => new Promise(r => setTimeout(r, ms));
    const backoffMs = (attempt: number) => Math.min(500 * 2 ** attempt, 8_000);
    // One controller per FETCH_BLOB call. Registered with the tab's group so
    // STOP_EXPORT can abort every in-flight fetch for that tab at once.
    const controller = new AbortController();
    trackFetch(tabId, controller);
    const init: RequestInit = msg.bearerToken
        ? { headers: { 'Authorization': `Bearer ${msg.bearerToken}` }, signal: controller.signal }
        : { credentials: 'include', signal: controller.signal };
    const inCoolOff = Date.now() < rateLimitCoolOffUntil;
    const maxAttempts = inCoolOff ? 1 : MAX_ATTEMPTS;

    let lastStatus: number | undefined;
    let lastStatusText: string | undefined;
    let lastError: string | undefined;

    try {
        for (let attempt = 0; attempt < maxAttempts; attempt++) {
            if (controller.signal.aborted) {
                sendResponse({ ok: false, cancelled: true });
                return;
            }
            try {
                const resp = await fetch(msg.url, init);
                if (resp.ok) {
                    const blob = await resp.blob();
                    if (blob.size > maxBytes || blob.size < minBytes) {
                        sendResponse({ ok: false, sizeReason: blob.size });
                        return;
                    }
                    const dataUrl = await new Promise<string>((resolve, reject) => {
                        const reader = new FileReader();
                        reader.onloadend = () => resolve(reader.result as string);
                        reader.onerror = () => reject(reader.error);
                        reader.readAsDataURL(blob);
                    });
                    rateLimitStreak = 0;
                    rateLimitCoolOffUntil = 0;
                    sendResponse({ ok: true, dataUrl, size: blob.size });
                    return;
                }
                lastStatus = resp.status;
                lastStatusText = resp.statusText;
                // Retry 429 (rate-limit), 408 (timeout), 410 (the Teams URL-image
                // proxy returns 410 for transient upstream failures, not just
                // permanent gone), and any 5xx. All other 4xx stay permanent.
                const transient = resp.status === 429 || resp.status === 408
                    || resp.status === 410 || resp.status >= 500;
                if (!transient) {
                    sendResponse({ ok: false, status: resp.status, statusText: resp.statusText });
                    return;
                }
                if (attempt < maxAttempts - 1) {
                    const retryAfter = resp.headers.get('Retry-After');
                    const waitMs = retryAfter ? Math.min(parseInt(retryAfter, 10) * 1000, 30_000) : backoffMs(attempt);
                    await sleep(waitMs);
                }
            } catch (e) {
                if (controller.signal.aborted) {
                    sendResponse({ ok: false, cancelled: true });
                    return;
                }
                lastError = String(e);
                if (attempt < maxAttempts - 1) await sleep(backoffMs(attempt));
            }
        }
        if (lastStatus === 429) {
            rateLimitStreak++;
            if (rateLimitStreak >= RATE_LIMIT_STREAK_THRESHOLD) {
                rateLimitCoolOffUntil = Date.now() + RATE_LIMIT_COOL_OFF_MS;
            }
        }
        sendResponse({ ok: false, status: lastStatus, statusText: lastStatusText, error: lastError });
    } finally {
        untrackFetch(tabId, controller);
    }
}

resetBadge();

runtime.onMessage.addListener((msg: BackgroundIncomingMessage, sender, sendResponse) => {
    if (!msg || !msg.type) return;

    if (msg.type === 'PING_SW') {
        sendResponse({ ok: true, now: Date.now() });
        return;
    }

    if (msg.type === 'BUILD_AND_DOWNLOAD') {
        handleBuildAndDownloadMessage(msg, sendResponse);
        return true;
    }

    if (msg.type === 'START_EXPORT') {
        handleStartExportMessage(msg, sendResponse);
        return true;
    }

    if (msg.type === 'START_BUNDLE_EXPORT') {
        handleStartBundleExportMessage(msg, sendResponse);
        return true;
    }

    if (msg.type === 'EXPORT_STATUS_UPDATE') {
        const payload = msg.payload || {};
        broadcastStatus(payload);
        if (payload?.tabId != null && TERMINAL_PHASES.has(payload.phase || '')) {
            activeExports.delete(payload.tabId); void persistActiveExports();
        }
        sendResponse({ ok: true });
        return true;
    }

    if (msg.type === 'SCRAPE_PROGRESS') {
        const payload = msg.payload || msg;
        const senderTabId = sender?.tab?.id;
        // Drop late progress from a tab whose export was just cancelled — the
        // event was already in flight when stop ran and would otherwise repaint
        // the badge / popup with stale numbers.
        if (typeof senderTabId === 'number' && cancelledTabs.has(senderTabId)) {
            return;
        }
        updateBadgeForProgress(payload);
        // Forward progress to popup for detailed status display
        try {
            const fwd = runtime.sendMessage({ type: 'EXPORT_PROGRESS', ...payload });
            if (fwd && fwd.catch) fwd.catch(() => {});
        } catch { /* popup may be closed */ }
        return;
    }

    if (msg.type === 'FETCH_BLOB') {
        // sender.tab.id groups in-flight fetches under the originating tab so
        // STOP_EXPORT can abort them all at once.
        handleFetchBlobMessage(msg, sendResponse, sender?.tab?.id);
        return true;
    }

    // No SHOW_DOWNLOAD / OPEN_DOWNLOAD handlers: the popup calls
    // chrome.downloads.show / .open directly from its click handler so the
    // user-activation check on .open() sees a live gesture. Routing through
    // the service worker loses the activation and fails with "User gesture
    // required." See openSavedDownload in App.svelte.

    if (msg.type === 'STOP_EXPORT') {
        const tabId = typeof msg.tabId === 'number' ? msg.tabId : sender?.tab?.id;
        if (typeof tabId !== 'number') {
            sendResponse({ ok: false, error: 'no-tab-id' });
            return;
        }
        handleStopExportMessage(tabId, sendResponse);
        return true;
    }

    if (msg.type === 'GET_EXPORT_STATUS') {
        const tabId = typeof msg.tabId === 'number' ? msg.tabId : sender?.tab?.id;
        if (typeof tabId !== 'number') {
            sendResponse({ active: false });
            return;
        }
        const info = activeExports.get(tabId) || null;
        sendResponse({ active: Boolean(info), info });
        return;
    }

    if (msg.type === 'LIST_CONVERSATIONS_QUICK') {
        // Fast first-paint: IDB-only read in the content script,
        // no Graph or roster fetches. Returns ~instantly.
        const tabId = typeof msg.tabId === 'number' ? msg.tabId : sender?.tab?.id;
        if (typeof tabId !== 'number') {
            sendResponse({ ok: false, error: 'missing-tab-id' });
            return;
        }
        (async () => {
            try {
                await ensureContentScript(tabId);
                const resp = await sendMessageToTab(tabId, { type: 'LIST_CONVERSATIONS_QUICK' });
                sendResponse(resp ?? { ok: false, error: 'no-response' });
            } catch (e: any) {
                sendResponse({ ok: false, error: String(e?.message || e) });
            }
        })();
        return true;
    }

    if (msg.type === 'LIST_CONVERSATIONS') {
        // Popup sends this to fetch the user's chat list. We ensure the
        // content script is injected, then forward the request to it —
        // only the content script has access to the Teams page's MSAL
        // tokens and the chat service.
        const tabId = typeof msg.tabId === 'number' ? msg.tabId : sender?.tab?.id;
        if (typeof tabId !== 'number') {
            sendResponse({ ok: false, error: 'missing-tab-id' });
            return;
        }
        (async () => {
            try {
                await ensureContentScript(tabId);
                const resp = await sendMessageToTab(tabId, { type: 'LIST_CONVERSATIONS' });
                sendResponse(resp ?? { ok: false, error: 'no-response' });
            } catch (e: any) {
                sendResponse({ ok: false, error: String(e?.message || e) });
            }
        })();
        return true; // async response
    }
});

async function handleStopExportMessage(tabId: number, sendResponse: (resp: unknown) => void) {
    log('STOP_EXPORT for tab', tabId);
    const existing = activeExports.get(tabId);

    // Race guard: small/fast exports can finish saving in the gap between
    // the user pressing Stop and this handler running. If the export
    // already reached phase='complete', a success row was just persisted —
    // we must not also persist a cancelled row, and we shouldn't broadcast
    // 'cancelled' to the popup either (the user just got their file).
    if (existing?.phase === 'complete') {
        log('STOP_EXPORT ignored — export already completed');
        sendResponse({ ok: true, alreadyComplete: true });
        return;
    }

    // Capture startedAt before we delete the active-export entry so the
    // persisted snapshot can include how long the user had run the export.
    const stopStartedAt = existing?.startedAt;

    // Tell the popup right away — the actual teardown takes a beat as the
    // content script unwinds, but the user gets feedback immediately.
    broadcastStatus({ tabId, phase: 'cancelling' });

    // Multi-chat bundle: signal the loop to stop before the next chat.
    // Combined with the requestScrape rejection below this aborts the
    // current chat AND skips the remaining ones.
    bundleStops.add(tabId);

    // Forward to the content script so it aborts its IC3 pagination loop
    // (or whichever phase it's in) and discards collected data.
    try {
        await sendMessageToTab(tabId, { type: 'STOP_SCRAPE' });
    } catch (e) {
        log('failed to forward STOP_SCRAPE:', e);
    }

    // Cancel any in-flight FETCH_BLOB calls for this tab — without this,
    // up to ~30s of pending image fetches can keep sockets and buffers alive
    // after the user already pressed Stop.
    const aborted = abortFetchesForTab(tabId);
    if (aborted > 0) log(`aborted ${aborted} in-flight fetches for tab ${tabId}`);

    // Reject any pending streaming-scrape promises for this tab and clear the
    // buffered messages array so the data doesn't sit in memory waiting for
    // the closure to be GC'd.
    for (const [requestId, pending] of pendingScrapes) {
        if (requestId.startsWith(`${tabId}-`)) {
            pendingScrapes.delete(requestId);
            pending.reject(new Error('cancelled'));
        }
    }

    activeExports.delete(tabId); void persistActiveExports();
    // Mark the tab as cancelled so any late SCRAPE_PROGRESS events that were
    // already in flight don't repaint the badge with stale numbers. The mark
    // is cleared after a few seconds — long enough for the in-flight events
    // to drain, short enough not to leak across a future export.
    cancelledTabs.add(tabId);
    setTimeout(() => cancelledTabs.delete(tabId), 5_000);
    // Clear the toolbar badge so the partial message count doesn't linger
    // after the user pressed Stop.
    resetBadge();

    // Append a cancelled row to the persisted history. The popup's History
    // page renders these as muted rows so the user has an audit trail of
    // "I tried this and stopped" — useful when revisiting context later.
    const stopElapsedMs = stopStartedAt ? Math.max(0, Date.now() - stopStartedAt) : 0;
    const stopConvId = await getConvIdForTab(tabId);
    await persistHistoryEntry({
        id: makeEntryId(),
        tabId,
        kind: 'cancelled',
        convId: stopConvId,
        elapsedMs: stopElapsedMs,
        savedAt: Date.now(),
    });

    broadcastStatus({ tabId, phase: 'cancelled' });
    sendResponse({ ok: true });
}

}); // End of defineBackground
