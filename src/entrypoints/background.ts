import { defineBackground } from 'wxt/sandbox';
import { createBadgeManager } from '../utils/badge';
import { isTeamsUrl } from '../utils/teams-urls';
import { buildAndDownload, buildAndDownloadZip } from '../background/download';
import { formatDayLabelForExport, parseTimeStamp } from '../utils/time';
import type { BackgroundIncomingMessage } from '../types/messaging';
import type {
  ActiveExportInfo,
  BuildOptions,
  ExportMessage,
  ExportMeta,
  ExportStatusPayload,
  Reaction,
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
const badge = createBadgeManager(action);
const { set: setBadge, reset: resetBadge, clearSoon: clearBadgeSoon, updateForStatus: updateBadgeForStatus, updateForProgress: updateBadgeForProgress } = badge;
const isFirefox = typeof browser !== 'undefined' && navigator.userAgent.includes('Firefox');

function log(...a: unknown[]) { try { console.log("[Teams Exporter SW]", ...a) } catch { } }
log("boot");

runtime.onInstalled.addListener(() => {
    log("onInstalled");
    resetBadge();
});
runtime.onStartup?.addListener(() => {
    log("onStartup");
    resetBadge();
});

tabs.onUpdated.addListener((tabId: number, changeInfo: chrome.tabs.TabChangeInfo, tab?: chrome.tabs.Tab) => {
    const nextUrl = changeInfo.url ?? tab?.url;
    if (changeInfo.status === 'loading' && isTeamsUrl(nextUrl)) {
        activeExports.delete(tabId);
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
    return next;
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
    await scripting.executeScript({ target: { tabId, allFrames: true }, files: ['content.js'] });
    const pong2 = await sendMessageToTab(tabId, { type: 'PING' });
    if (!pong2?.ok) throw new Error('Content script did not respond after injection');
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

function broadcastStatus(payload: ExportStatusPayload) {
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
    if (activeExports.has(tabId)) {
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

            broadcastStatus({ tabId, phase: 'complete', filename: buildRes.filename });
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
            activeExports.delete(tabId);
        }
    })();
}

function handleStartExportMessage(msg: any, sendResponse: (res: any) => void) {
    handleExportWithScrape(msg, sendResponse, async (scrapeRes, buildOptions, tabId) => {
        const format = buildOptions.format || 'json';
        const downloadImages = Boolean(buildOptions.downloadImages);
        const deps = {
            downloads,
            isFirefox,
            onStatus: (payload: Record<string, unknown>) => broadcastStatus({ ...payload, tabId }),
        };
        const commonOpts = {
            messages: scrapeRes.messages || [],
            meta: scrapeRes.meta || {},
            embedAvatars: Boolean(buildOptions.embedAvatars),
            downloadImages,
        };
        if (format === 'html' && downloadImages) {
            return buildAndDownloadZip(deps, commonOpts);
        }
        return buildAndDownload(deps, {
            ...commonOpts,
            format,
            saveAs: buildOptions.saveAs !== false,
        });
    });
}

function handleStartExportZipMessage(msg: any, sendResponse: (res: any) => void) {
    handleExportWithScrape(msg, sendResponse, async (scrapeRes, buildOptions, tabId) => {
        return buildAndDownloadZip(
            {
                downloads,
                isFirefox,
                onStatus: (payload: Record<string, unknown>) => broadcastStatus({ ...payload, tabId }),
            },
            {
                messages: scrapeRes.messages || [],
                meta: scrapeRes.meta || {},
                embedAvatars: Boolean(buildOptions.embedAvatars),
                downloadImages: Boolean(buildOptions.downloadImages),
            },
        );
    });
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

    if (msg.type === 'START_EXPORT_ZIP') {
        handleStartExportZipMessage(msg, sendResponse);
        return true;
    }

    if (msg.type === 'EXPORT_STATUS_UPDATE') {
        const payload = msg.payload || {};
        broadcastStatus(payload);
        if (payload?.tabId != null && TERMINAL_PHASES.has(payload.phase || '')) {
            activeExports.delete(payload.tabId);
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
});

async function handleStopExportMessage(tabId: number, sendResponse: (resp: unknown) => void) {
    log('STOP_EXPORT for tab', tabId);
    // Tell the popup right away — the actual teardown takes a beat as the
    // content script unwinds, but the user gets feedback immediately.
    broadcastStatus({ tabId, phase: 'cancelling' });

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

    activeExports.delete(tabId);
    // Mark the tab as cancelled so any late SCRAPE_PROGRESS events that were
    // already in flight don't repaint the badge with stale numbers. The mark
    // is cleared after a few seconds — long enough for the in-flight events
    // to drain, short enough not to leak across a future export.
    cancelledTabs.add(tabId);
    setTimeout(() => cancelledTabs.delete(tabId), 5_000);
    // Clear the toolbar badge so the partial message count doesn't linger
    // after the user pressed Stop.
    resetBadge();
    broadcastStatus({ tabId, phase: 'cancelled' });
    sendResponse({ ok: true });
}

}); // End of defineBackground
