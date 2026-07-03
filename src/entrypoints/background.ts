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
  type BundleEmpty,
  type BundleFailure,
} from '../background/download';
import {
  collectDocumentAttachments,
  downloadAttachments,
  downloadChatAttachments,
  toFilesSummaryWire,
  verifyDownloadOnChanged,
  type AttachmentCandidate,
  type AttachmentDownloadSummary,
} from '../background/attachment-download';
import { waitForDownloadSettled, type SettledOutcome } from '../background/download-wait';
import { revokeDownloadUrl, textToDownloadUrl } from '../background/builders';
import { createZipStream, type ZipStream } from '../background/zip';
import { beginExportStats, endExportStats, getExportStats } from '../background/export-stats';
import { clearTransferBlobs } from '../utils/blob-transfer';
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

// Diagnostic log tail.
//
// BG owns the master log buffer. It contains:
//   - BG's own console captures (src: 'bg'), wrapped here at module load
//   - lines forwarded from content scripts (src: 'content'), received
//     via DIAG_LOG_FORWARD messages
//
// Persistence is opt-in (Options.diagLogPersist, default false). When
// off, the buffer lives in memory only and is lost on SW eviction.
// When on, BG flushes a debounced copy of the buffer to
// chrome.storage.local under DIAG_LOG_STORAGE_KEY. On SW startup the
// stored entries are restored (prepended to whatever lines the wrap
// has captured during boot).
//
// Trim is byte-based: the buffer cannot exceed DIAG_LOG_BYTE_CAP when
// serialized. Oldest entries are dropped first. This keeps the disk
// footprint bounded while letting the buffer hold thousands of typical
// log lines.
type DiagLogEntry = { src: 'bg' | 'content'; ts: number; level: string; line: string };
const DIAG_LOG_BYTE_CAP = 8 * 1024 * 1024;
const DIAG_LOG_STORAGE_KEY = 'teamsExporterDiagLog';
const DIAG_FLUSH_DEBOUNCE_MS = 500;
const DIAG_PREFIX_RE = /^\[[A-Za-z][A-Za-z0-9 _\-:]{0,40}\]/;
const diagLogBuffer: DiagLogEntry[] = [];
let diagLogPersistEnabled = false;
// Opt-in (Options.diagVerboseStats, default false). When on, the
// [export-stats] block includes per-chat detail (titles, convIds, per-chat
// scrape-stage split); off, only the ID-free aggregates are logged. Toggled
// from the Diagnostics page. Independent of log persistence.
let diagVerboseStatsEnabled = false;
let diagLogFlushTimer: ReturnType<typeof setTimeout> | null = null;
// Tracks the most recent in-flight storage write so toggle-off can
// await it before removing the storage key. Without this, an in-flight
// set() can land AFTER remove() and silently resurrect the key.
let diagInflightFlush: Promise<void> | null = null;
// Most recent storage write failure (null if none, or if the last
// write succeeded). The popup surfaces this so a user with
// persistence on but zero bytes on disk sees a clear "writes are
// failing" hint instead of an ambiguous 0.
let diagLastFlushError: { ts: number; reason: string } | null = null;
// Running byte count of `diagLogBuffer`'s serialized form. Maintained
// incrementally to avoid an O(n) JSON.stringify per push (which would
// stack up to hundreds of ms on a near-cap buffer during a verbose
// export). The estimate is intentionally conservative (over-counts a
// little) so the cap stays a safe floor for actual on-disk size.
let diagBufferBytes = 0;
function estimateEntryBytes(e: DiagLogEntry): number {
    // JSON overhead per entry: 4 keys (`src`,`ts`,`level`,`line`),
    // each with quotes + colon + comma, plus the bracketed wrapper.
    // ~60 bytes is a generous round number that covers the structure.
    return 60 + (e.line?.length || 0) + (e.level?.length || 0);
}
function trimDiagBuffer() {
    // Drop oldest entries until under cap. Small floor (50) keeps the
    // buffer from dropping to zero on a pathological single mega-entry.
    while (diagLogBuffer.length > 50 && diagBufferBytes > DIAG_LOG_BYTE_CAP) {
        const removed = diagLogBuffer.shift();
        if (removed) diagBufferBytes -= estimateEntryBytes(removed);
    }
}
function pushDiagEntry(entry: DiagLogEntry) {
    diagLogBuffer.push(entry);
    diagBufferBytes += estimateEntryBytes(entry);
    diagBufferDirty = true;
    trimDiagBuffer();
    if (diagLogPersistEnabled) scheduleDiagFlush();
}
function recomputeBufferBytes() {
    diagBufferBytes = diagLogBuffer.reduce((sum, e) => sum + estimateEntryBytes(e), 0);
}
// True between a push and the next successful flush. Lets the
// post-flush finalizer decide whether to re-arm a flush (catches
// pushes that arrived while the previous write was in flight).
let diagBufferDirty = false;
function scheduleDiagFlush() {
    // Single-writer guard: drop the call if another flush is already
    // scheduled OR mid-write. The in-flight finalizer below catches up
    // by re-arming when it sees dirty=true at finish.
    if (diagLogFlushTimer || diagInflightFlush) return;
    diagLogFlushTimer = setTimeout(() => {
        diagLogFlushTimer = null;
        if (!diagLogPersistEnabled) return;
        // Snapshot dirty=false now; any pushes that arrive during the
        // write below will flip it back to true and the finally clause
        // re-arms scheduleDiagFlush.
        diagBufferDirty = false;
        diagInflightFlush = (async () => {
            try {
                await chrome.storage.local.set({ [DIAG_LOG_STORAGE_KEY]: diagLogBuffer.slice() });
                diagLastFlushError = null;
            } catch (e: any) {
                const msg = e?.message || String(e);
                diagLastFlushError = { ts: Date.now(), reason: msg };
                // Quota recovery: drop oldest half so the next flush
                // attempt fits. Without this, a stuck-quota state
                // captures lastFlushError once and then keeps failing
                // identically forever. Halving is conservative.
                if (/quota/i.test(msg)) {
                    const half = Math.floor(diagLogBuffer.length / 2);
                    if (half > 0) {
                        diagLogBuffer.splice(0, half);
                        recomputeBufferBytes();
                    }
                }
                // Ensure the retry happens regardless of whether new
                // pushes came in during the in-flight write.
                diagBufferDirty = true;
            } finally {
                diagInflightFlush = null;
                if (diagLogPersistEnabled && diagBufferDirty) {
                    scheduleDiagFlush();
                }
            }
        })();
    }, DIAG_FLUSH_DEBOUNCE_MS);
}
async function restoreDiagBuffer() {
    try {
        const r = await chrome.storage.local.get([DIAG_LOG_STORAGE_KEY, 'teamsExporterOptions']);
        const opts = r['teamsExporterOptions'];
        diagLogPersistEnabled = !!opts?.diagLogPersist;
        // Sync the verbose-stats flag regardless of persistence (the early
        // return below only concerns the log buffer restore).
        diagVerboseStatsEnabled = !!opts?.diagVerboseStats;
        if (!diagLogPersistEnabled) return;
        const stored = r[DIAG_LOG_STORAGE_KEY];
        if (!Array.isArray(stored) || stored.length === 0) return;
        // Prepend restored entries: they are older than anything the
        // boot-time console wrap might have already captured.
        const inMem = diagLogBuffer.splice(0);
        diagLogBuffer.push(...(stored as DiagLogEntry[]), ...inMem);
        recomputeBufferBytes();
        trimDiagBuffer();
    } catch { /* fall through; restore is best-effort */ }
}
// Idempotent guard: a second module load (HMR, recovery) skips the
// re-wrap so we don't double-capture.
if (!(console as unknown as { __tceDiagWrapped?: boolean }).__tceDiagWrapped) {
    (console as unknown as { __tceDiagWrapped: boolean }).__tceDiagWrapped = true;
    const origLog = console.log.bind(console);
    const origWarn = console.warn.bind(console);
    const origError = console.error.bind(console);
    const origInfo = console.info ? console.info.bind(console) : origLog;
    const origDebug = console.debug ? console.debug.bind(console) : origLog;
    const formatArg = (a: unknown): string => {
        if (typeof a === 'string') return a;
        try { return JSON.stringify(a); } catch { return String(a); }
    };
    const capture = (level: string, args: unknown[]) => {
        try {
            // Cheap gate before formatting: skip the JSON.stringify
            // pass entirely on lines that clearly aren't ours. Errors
            // bypass the prefix filter so real exceptions still land
            // in the buffer even without a `[tag]` prefix.
            if (level !== 'error') {
                const first = args[0];
                if (typeof first !== 'string' || !DIAG_PREFIX_RE.test(first)) return;
            }
            const line = args.map(formatArg).join(' ');
            pushDiagEntry({ src: 'bg', ts: Date.now(), level, line });
        } catch { /* never break logging */ }
    };
    console.log = (...args: unknown[]) => { capture('log', args); origLog(...args); };
    console.warn = (...args: unknown[]) => { capture('warn', args); origWarn(...args); };
    console.error = (...args: unknown[]) => { capture('error', args); origError(...args); };
    console.info = (...args: unknown[]) => { capture('info', args); origInfo(...args); };
    console.debug = (...args: unknown[]) => { capture('debug', args); origDebug(...args); };
}
// Kick off restore. Async; lines captured during this window will be
// preserved (restore prepends, not overwrites).
void restoreDiagBuffer();

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
// Per-build identifier (git hash + build time), injected at build time. Logged
// at boot and at each export start/end so any SW console / export can be tied to
// the exact build — no more guessing whether a stale service worker is running.
const BUILD_STAMP = __BUILD_STAMP__;
log(`boot build=${BUILD_STAMP} [files=settled-package; markup=uniqueid; verify=onchanged-mime; dl-complete=enum+settle]`);

// Post-download request-access cleanup for the Files toggle. Registered at the
// service-worker top level so a download completing after the export's await
// chain has ended (or after the worker was evicted and re-spawned) still wakes
// the worker and reaches the verifier. A non-markup attachment that lands as
// text/html is a "request access" page; verifyDownloadOnChanged deletes it and
// records the file in FAILURES.txt. See attachment-download.ts.
try {
    downloads.onChanged.addListener(d => { void verifyDownloadOnChanged({ downloads, log }, d); });
} catch { /* noop: downloads API unavailable */ }

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

// True when a DownloadItem's bytes are fully on disk, even if we never observed
// the 'complete' state transition. Used as a last-chance check on timeout to
// distinguish a real stall from a missed 'complete' signal. Both engines
// serialize `exists`; the ===false guard rejects files already removed from
// disk. (An earlier comment here claimed Firefox omits the field — wrong.)
function downloadIsOnDisk(item: chrome.downloads.DownloadItem): boolean {
    if ((item as { exists?: boolean }).exists === false) return false;
    if (item.state === 'interrupted') return false;
    // Empty filename = target still undecided (a Save As dialog, or the
    // browser's ask-where-to-save chooser, is open). Chrome spools the blob's
    // bytes to a temp file DURING the dialog, so bytesReceived can equal
    // totalBytes while the user has neither confirmed nor cancelled — that is
    // an undecided save, never a finished one.
    if (!item.filename) return false;
    const received = item.bytesReceived || 0;
    if (received <= 0) return false;
    const total = item.totalBytes;
    // total <= 0 means the size was never reported (a data: URL can read -1);
    // after the full timeout, a non-interrupted item with bytes received is a
    // finished write, not a partial one.
    return total <= 0 || received >= total;
}

// The InterruptReason (e.g. FILE_NO_SPACE, USER_CANCELED) of a download, or
// undefined when it can't be read. Firefox encodes a manual cancel as an
// EMPTY error where Chrome uses USER_CANCELED; callers treat both as a cancel.
async function downloadErrorReason(id: number): Promise<string | undefined> {
    try {
        const r = await downloads.search({ id });
        const item = Array.isArray(r) ? r[0] : undefined;
        return item?.error || undefined;
    } catch { return undefined; }
}

// downloads.download() returns the item's ID before the file is actually
// written to disk. Calling downloads.open() in that window would fail for
// larger exports (blob-URL writes can take a few hundred ms for 50MB+). Watch
// the item for up to `timeoutMs` and report which terminal state it reached:
//   'complete'    — reached state='complete'.
//   'interrupted' — reached state='interrupted' (a real cancel/failure, e.g.
//                   the user hit Cancel in the Save As dialog).
//   'timeout'     — the window elapsed without observing either. A data:-scheme
//                   download settles synchronously and can reach 'complete' in
//                   the gap between download() resolving and our listener
//                   attaching, so the one fast-path search can miss it; the
//                   timeout does a final on-disk recheck and upgrades to
//                   'complete' when the file is actually there, rather than
//                   mislabeling a saved export as cancelled.
type DownloadOutcome = 'complete' | 'interrupted' | 'timeout';

function waitForDownloadComplete(id: number, timeoutMs = 10_000): Promise<DownloadOutcome> {
    return new Promise(resolve => {
        let done = false;
        const finish = (outcome: DownloadOutcome) => {
            if (done) return;
            done = true;
            try { downloads.onChanged.removeListener(onChange); } catch { /* noop */ }
            clearTimeout(timer);
            resolve(outcome);
        };
        const onChange = (delta: chrome.downloads.DownloadDelta) => {
            if (delta.id !== id) return;
            if (delta.state?.current === 'complete') finish('complete');
            else if (delta.state?.current === 'interrupted') finish('interrupted');
        };
        const onTimeout = () => {
            // Hard backstop: a hung downloads.search() must never leave this
            // promise unresolved (that would freeze the whole export). Force
            // 'timeout' after a short grace if the recheck doesn't settle.
            const hard = setTimeout(() => finish('timeout'), 2_000);
            // Last chance: the 'complete' we missed may already be on disk.
            Promise.resolve(downloads.search({ id }))
                .then(results => {
                    clearTimeout(hard);
                    const item = Array.isArray(results) ? results[0] : undefined;
                    if (item && (item.state === 'complete' || downloadIsOnDisk(item))) finish('complete');
                    else if (item?.state === 'interrupted') finish('interrupted');
                    else finish('timeout');
                })
                .catch(() => { clearTimeout(hard); finish('timeout'); });
        };
        const timer = setTimeout(onTimeout, timeoutMs);
        try { downloads.onChanged.addListener(onChange); } catch { /* noop */ }
        // Fast path — the item may already be complete by the time we check.
        Promise.resolve(downloads.search({ id }))
            .then(results => {
                const item = Array.isArray(results) ? results[0] : undefined;
                if (item?.state === 'complete') finish('complete');
                else if (item?.state === 'interrupted') finish('interrupted');
            })
            .catch(() => { /* leave the listener / timeout to resolve it */ });
    });
}

// Drop any orphaned download-transfer Blob left by a worker that died
// mid-handoff. MV3 only (the offscreen blob-URL path is the only producer);
// guarded on chrome.offscreen so Firefox MV2 never opens the IDB store.
function cleanupTransferBlobs() {
    try {
        if (typeof chrome !== 'undefined' && chrome.offscreen) {
            void clearTransferBlobs().catch(() => { /* best-effort */ });
        }
    } catch { /* noop */ }
}

runtime.onInstalled.addListener((details) => {
    log("onInstalled", details?.reason);
    resetBadge();
    cleanupTransferBlobs();
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
    cleanupTransferBlobs();
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
    // Oversized-image dataUrls peeled off single messages that were too big for
    // one port message; reattached by message + attachment index on 'done'.
    const chunks: Array<{ mi: number; ai: number; dataUrl: string }> = [];

    port.onMessage.addListener((msg: any) => {
        if (msg.type === 'meta') {
            meta = msg.meta || {};
            log('received meta:', { title: meta.title, avatars: Object.keys(meta.avatars || {}).length });
        } else if (msg.type === 'messages') {
            const batch = Array.isArray(msg.messages) ? msg.messages : [];
            messages.push(...batch);
            batches++;
            log(`received batch ${batches}: ${batch.length} messages (total: ${messages.length})`);
        } else if (msg.type === 'attachment-chunk') {
            chunks.push({ mi: msg.mi, ai: msg.ai, dataUrl: msg.dataUrl });
        } else if (msg.type === 'done') {
            // Reattach any peeled oversized-image dataUrls now that every message
            // has arrived (chunks are always sent after all message batches).
            for (const c of chunks) {
                const att = messages[c.mi]?.attachments?.[c.ai];
                if (att) att.dataUrl = c.dataUrl;
            }
            log(`streaming complete: ${messages.length} messages in ${batches} batches${chunks.length ? `, ${chunks.length} reattached image chunk(s)` : ''}`);
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
    // Feed the build phase into the stage timeline (the fine fetch phases
    // arrive via SCRAPE_PROGRESS instead; 'build' only flows through here).
    if (tabId != null) getExportStats(tabId)?.markPhase(payload?.phase, Date.now());
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
    // `folder` is set only in package mode (Files toggle on): the Downloads-
    // relative folder the export was saved into, which the attachments phase
    // must reuse verbatim so both land in the same place.
) => Promise<{ filename?: string; id?: number; folder?: string }>;

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
        // Console-only diagnostics. statsOutcome is set at each terminal
        // branch; the finally block logs the collected block and clears it.
        let statsOutcome = 'error';
        let statsFilename: string | undefined;
        try {
            await ensureContentScript(tabId);
            const ctx = await checkContext(tabId, scrapeOptions);
            if (!ctx?.ok) {
                const message = ctx?.reason || defaultContextError(scrapeOptions);
                sendResponse({ error: message });
                return;
            }

            startedAt = Date.now();
            beginExportStats(tabId, 'single', startedAt);
            updateActiveExport(tabId, { startedAt, phase: 'starting', lastStatus: undefined });
            // Await the FIRST snapshot write specifically. The popup's
            // instant-paint reads this key on mount; if the user reopens the
            // popup before this lands they fall through to the slow
            // GET_EXPORT_STATUS round-trip (worsened by service-worker cold
            // start), which shows "Checking status…" for a beat. Progress
            // writes below stay fire-and-forget; only this start write gates
            // the reopen reveal.
            await persistActiveExports();
            broadcastStatus({ tabId, phase: 'starting', startedAt });
            log(`export start build=${BUILD_STAMP} tab=${tabId}`);

            broadcastStatus({ tabId, phase: 'scrape:start' });
            const scrapeStartedAt = Date.now();
            getExportStats(tabId)?.beginChat(scrapeStartedAt);
            const scrapeRes = await requestScrape(tabId, scrapeOptions);
            const totalMessages = Array.isArray(scrapeRes.messages) ? scrapeRes.messages.length : 0;
            broadcastStatus({ tabId, phase: 'scrape:complete', messages: totalMessages });
            // One chat per single export. Title/convId come from the scrape
            // meta when present (no extra tab round-trip on the stats path).
            const scrapeDoneAt = Date.now();
            getExportStats(tabId)?.addChat({
                title: typeof scrapeRes.meta?.title === 'string' ? scrapeRes.meta.title : undefined,
                convId: typeof scrapeRes.meta?.conversationId === 'string' ? scrapeRes.meta.conversationId : undefined,
                messages: totalMessages,
                scrapeMs: scrapeDoneAt - scrapeStartedAt,
            }, scrapeDoneAt);

            if (totalMessages === 0) {
                statsOutcome = 'empty';
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
            log('awaiting zip download completion id=', buildRes.id);
            const packageMode = Boolean(buildOptions?.downloadFiles);
            let downloadStatus: 'complete' | 'interrupted' | 'timeout';
            if (packageMode && buildRes.id != null) {
                // Package mode (Files on): the attachments phase must never run
                // against an export that isn't actually on disk, so there is no
                // proceed-on-timeout here — only a confirmed terminal state. The
                // wait is event-driven with a keep-alive poll; a genuine stall or
                // a vanished download is an error, not a silent success. STOP
                // aborts the wait via the tab's tracked controllers.
                const waitAbort = new AbortController();
                trackFetch(tabId, waitAbort);
                let settled: SettledOutcome;
                try {
                    settled = await waitForDownloadSettled(
                        { downloads, onChanged: downloads.onChanged },
                        buildRes.id,
                        { signal: waitAbort.signal },
                    );
                } finally {
                    untrackFetch(tabId, waitAbort);
                }
                log('export download settle ->', settled, 'id=', buildRes.id);
                if (settled === 'aborted') {
                    // STOP while the artifact was still writing: the file is not
                    // (fully) on disk — cancel it and report cancelled. The STOP
                    // handler owns the cancelled history row.
                    try { await downloads.cancel(buildRes.id); } catch { /* already terminal */ }
                    statsOutcome = 'cancelled';
                    broadcastStatus({ tabId, phase: 'cancelled' });
                    sendResponse({ cancelled: true });
                    return;
                }
                if (settled === 'interrupted') {
                    // With saveAs:false there is no Save As dialog, so an
                    // interrupt is a real failure (disk full, blocked by
                    // policy...) unless the user cancelled it from the shelf.
                    const reason = await downloadErrorReason(buildRes.id);
                    if (reason && reason !== 'USER_CANCELED') {
                        throw new Error(`Export download failed (${reason})`);
                    }
                    statsOutcome = 'cancelled';
                    log('export download interrupted (cancelled) for id', buildRes.id);
                    broadcastStatus({ tabId, phase: 'cancelled' });
                    sendResponse({ cancelled: true });
                    return;
                }
                if (settled === 'stalled') throw new Error('Export download did not finish (no progress)');
                if (settled === 'missing') throw new Error('Export download disappeared before completing');
                downloadStatus = 'complete';
            } else {
                downloadStatus = buildRes.id != null
                    ? await waitForDownloadComplete(buildRes.id, 30_000)
                    : 'complete';
            }
            log('zip download wait ->', downloadStatus, 'id=', buildRes.id);
            // Only a real interruption (e.g. Cancel in the Save As dialog, which
            // surfaces as state='interrupted') counts as cancelled. A 'timeout'
            // means we could not confirm completion in the window — common for
            // small data:-scheme exports that settle synchronously — but the file
            // is not interrupted, so we treat it as saved and continue. (Package
            // mode never reaches here with 'timeout'; its wait above only yields
            // confirmed outcomes.)
            if (downloadStatus === 'interrupted') {
                statsOutcome = 'cancelled';
                log('export download interrupted (cancelled) for id', buildRes.id);
                broadcastStatus({ tabId, phase: 'cancelled' });
                sendResponse({ cancelled: true });
                return;
            }
            if (downloadStatus === 'timeout') {
                log('export download completion unconfirmed; proceeding for id', buildRes.id);
            }

            statsOutcome = 'complete';
            statsFilename = buildRes.filename;

            // Stream document attachments to disk (the "Files" toggle). Runs
            // after the main export is confirmed saved and only when
            // downloadFiles is on, so a normal export is byte-identical. In
            // package mode the builders return the folder the export was saved
            // INTO ('<base>/<base>.<ext>'), and the attachments/ tree goes in
            // the same folder. The legacy extension-strip stays as the
            // fallback (recomputing the base would mint a fresh, mismatched
            // timestamp). The AbortController is registered with the tab so
            // STOP_EXPORT (abortFetchesForTab) stops the phase.
            let filesSummary: AttachmentDownloadSummary | undefined;
            if (buildOptions?.downloadFiles && buildRes.filename) {
                log('Files phase: starting for', (scrapeRes.messages || []).length, 'messages');
                const exportFolder = buildRes.folder ?? buildRes.filename.replace(/\.[^.\\/]+$/, '');
                const filesAbort = new AbortController();
                trackFetch(tabId, filesAbort);
                try {
                    filesSummary = await downloadChatAttachments(
                        scrapeRes.messages || [],
                        exportFolder,
                        {
                            downloads,
                            log,
                            onProgress: (filesDone, filesTotal) =>
                                broadcastStatus({ tabId, phase: 'downloading-files', filesDone, filesTotal }),
                        },
                        filesAbort.signal,
                        // Firefox does not cap concurrent downloads; bound them so a
                        // large set doesn't burst SharePoint into throttling. Chrome
                        // caps per host, so it stays fire-and-forget (undefined).
                        { maxConcurrent: isFirefox ? 6 : undefined },
                    );
                } finally {
                    untrackFetch(tabId, filesAbort);
                }
            }

            broadcastStatus({
                tabId,
                phase: 'complete',
                filename: buildRes.filename,
                downloadId: buildRes.id,
                // Forward the setting so the popup can try auto-open from
                // its own context (see auto-action comment below).
                afterExport: buildOptions?.afterExport,
                filesSummary: filesSummary && toFilesSummaryWire(filesSummary),
            });

            // Append a row to the persisted export history. Reading happens
            // from the popup's HistoryPage; the entry is written here (not
            // popup-side) so it gets recorded even when the popup was
            // closed during the export.
            // Local startedAt, not the activeExports entry: a STOP during the
            // Files phase deletes that entry while this flow keeps running.
            const completeElapsedMs = startedAt
                ? Math.max(0, Date.now() - startedAt)
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
            //            the SW handles it here. A real interruption already
            //            returned above; we're here on a confirmed (or
            //            on-disk-but-unconfirmed) download, and show() only
            //            reveals the file's folder, so this is safe either way.
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
                statsOutcome = 'cancelled';
                sendResponse({ cancelled: true });
            } else {
                // Surface the error message in the SW console so a red-badge
                // failure is debuggable. Without this, broadcastStatus only
                // updates state — the actual reason never reaches the log.
                statsOutcome = 'error';
                log('export failed:', message);
                if (err?.stack) log('export stack:', err.stack);
                broadcastStatus({ tabId, phase: 'error', error: message });
                sendResponse({ error: message });
            }
        } finally {
            // Emit the console-only stats block, then drop the collector. Runs
            // for every outcome (complete / empty / cancelled / error).
            log(`export done build=${BUILD_STAMP} tab=${tabId} outcome=${statsOutcome}`);
            getExportStats(tabId)?.log(Date.now(), statsOutcome, { filename: statsFilename, verbose: diagVerboseStatsEnabled, buildStamp: BUILD_STAMP });
            endExportStats(tabId);
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
            // Per-format output sizes for the console-only stats block.
            onArtifact: (a: { format: string; bytes: number }) => getExportStats(tabId)?.addFormat(a.format, a.bytes),
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
            imageFilenameDate: Boolean(buildOptions.imageFilenameDate),
            imageModifiedDate: Boolean(buildOptions.imageModifiedDate),
            ...pdfKnobs,
        };
        // 2+ formats -> always bundle.zip. The bundle path doesn't honor
        // saveAs because zips force a Save As anyway via downloads.show
        // semantics (the file would conflict otherwise).
        const avatarMode = buildOptions.avatarMode === 'files' ? 'files' : 'inline';
        // Files toggle on -> package mode: the export saves INSIDE its base
        // folder (saveAs:false) so the attachments/ tree lands beside it.
        const packageFolder = Boolean(buildOptions.downloadFiles);
        if (formats.length >= 2) {
            return buildAndDownloadBundle(deps, { ...commonOpts, formats, avatarMode, packageFolder });
        }
        const format = formats[0];
        // HTML goes into a .zip when EITHER inline images are on OR the
        // user chose 'files' avatar mode (both need the zip's folder
        // structure). Without either, single HTML stays inline.
        const wantFiles = avatarMode === 'files' && Boolean(buildOptions.embedAvatars);
        if (format === 'html' && (downloadImages || wantFiles)) {
            return buildAndDownloadZip(deps, { ...commonOpts, avatarMode, packageFolder });
        }
        return buildAndDownload(deps, {
            ...commonOpts,
            format,
            saveAs: buildOptions.saveAs !== false,
            packageFolder,
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
        beginExportStats(tabId, 'bundle', startedAt);
        updateActiveExport(tabId, { startedAt, phase: 'starting', lastStatus: undefined });
        // Await the first snapshot write so a popup reopened right after the
        // bundle starts can instant-paint the busy state instead of waiting on
        // the GET_EXPORT_STATUS round-trip. See the single-export path for the
        // full rationale. Later progress writes stay fire-and-forget.
        await persistActiveExports();
        broadcastStatus({ tabId, phase: 'starting', startedAt });

        try {
            await ensureContentScript(tabId);
        } catch (e: any) {
            const message = e?.message || String(e);
            broadcastStatus({ tabId, phase: 'error', error: message });
            sendResponse({ error: message });
            endExportStats(tabId);
            activeExports.delete(tabId); void persistActiveExports();
            return;
        }

        const totalChats = list.length;
        const usedFolderNames = new Set<string>();
        // Successful chats are streamed straight into bundleZip as they build
        // (so their uncompressed bytes are freed immediately); `entries` keeps
        // only the folder name for the success count / history. bundleZip is
        // created lazily on the first successful chat, so the empty / all-empty
        // / all-fail paths (which never zip) never allocate one.
        const entries: { folderName: string }[] = [];
        let bundleZip: ZipStream | null = null;
        // Set if streaming a chat into the zip throws: a zip-stream error poisons
        // the whole archive, so it is a fatal bundle error, not one chat's failure.
        let bundleZipFatal: unknown = null;
        const failures: BundleFailure[] = [];
        const noHistory: BundleEmpty[] = [];
        // Bundle-level partial accumulator. Each entry tracks a chat
        // whose per-chat scrape signalled meta.partial. We surface this
        // at the bundle level rather than tagging individual chat
        // folders inside the zip — simpler to implement, and the
        // bundle-root PARTIAL.txt + outer-zip filename suffix are the
        // signals we want users to see anyway.
        const partials: { folderName: string; conversationId: string; reason: 'network' | 'truncation' }[] = [];
        // Document attachments (the "Files" toggle) collected per successful
        // chat. Resolved + streamed to disk after the outer zip is saved (the
        // bundle folder name isn't known until then). Each item carries its
        // per-chat folder so files land under <bundleBase>/<chat>/attachments/.
        const bundleAttachItems: AttachmentCandidate[] = [];
        // Set true at either cancel point so the stats finally below can
        // label the outcome without threading it through every return.
        let bundleCancelled = false;

        // Precompute every chat's folder name up front, IN ORDER, so the dedup
        // numbering is byte-identical to the serial version regardless of when
        // each chat is actually processed. The prefetch pipeline below can have
        // chat i+1's scrape running while chat i builds, so the names must not
        // depend on processing timing.
        const plan = list.map((conv) => ({
            conv,
            folderName: pickBundleFolderName(conv.title || conv.id, usedFolderNames),
        }));

        // Launch the scrape for plan[idx], returning the in-flight promise plus
        // its real start time, or null past the end. Errors are deliberately
        // NOT caught here — the awaiting site associates them with the right
        // chat. SAFETY: requestScrape sends exactly one SCRAPE_TEAMS, and we
        // only ever start the next scrape after the previous one has fully
        // resolved (its results streamed, its run-state reset). So the content
        // script — which has no concurrent-run guard — never runs two scrapes
        // at once; only the build (worker context) overlaps a scrape (page
        // context), and those don't compete.
        const launchScrape = (idx: number): { promise: Promise<ScrapeResult>; startedAt: number } | null => {
            if (idx < 0 || idx >= plan.length) return null;
            const promise = requestScrape(tabId, {
                ...baseScrapeOptions,
                conversationId: plan[idx].conv.id,
                conversationTitle: plan[idx].conv.title || null,
                noDomFallback: true,
            });
            // Mark the OUTER promise handled at the instant it is created. A
            // prefetched scrape sits un-awaited across the next chat's build
            // (the 1-deep window); if STOP rejects it there, this no-op catch
            // stops an "Uncaught (in promise) cancelled" in the worker console.
            // Attaching a catch does NOT consume the rejection for the real
            // await site below — that still observes the value or the error.
            // (requestScrape's own catch at the streamPromise guards a
            // different promise object than the async-fn promise stored here.)
            promise.catch(() => { /* may be abandoned by a mid-build cancel */ });
            return { promise, startedAt: Date.now() };
        };

        // Prefetch the first chat. From then on each iteration starts the NEXT
        // chat's scrape before building the current one, so paging (network,
        // 429-limited) overlaps building (CPU). Only one scrape is ever in
        // flight (1-deep), bounding the extra peak memory to a single chat's
        // data. Same request cadence as the serial version, so no extra 429
        // pressure.
        let pending = launchScrape(0);

        // Outer try/finally: the bundle has many terminal returns (cancelled,
        // empty, all-fail, success). Wrapping from here lets ONE finally emit
        // the console-only stats block and clear the collector for every path.
        // The accumulators above stay in scope so the outcome is derived, not
        // threaded.
        try {
        for (let i = 0; i < plan.length; i++) {
            if (bundleStops.has(tabId)) break;

            const { conv, folderName } = plan[i];
            const currentChat = i + 1;

            const bundleCtx = {
                bundleCurrentChat: currentChat,
                bundleTotalChats: totalChats,
                bundleChatName: folderName,
                bundleSuccessCount: entries.length,
                bundleFailedCount: failures.length,
            };

            broadcastStatus({ tabId, phase: 'scrape:start', ...bundleCtx });

            // `pending` holds chat i's scrape, prefetched last iteration (or
            // before the loop for i=0). startedAt is the real launch time, which
            // may predate this iteration because it overlapped chat i-1's build.
            // The `pending ? … : requestScrape(…)` fallback is defensive — at
            // i=0 the list is non-empty so launchScrape(0) is non-null, and
            // every later iteration sets `pending` before reaching here.
            const chatScrapeStartedAt = pending ? pending.startedAt : Date.now();
            getExportStats(tabId)?.beginChat(chatScrapeStartedAt);

            let scrapeRes: ScrapeResult;
            try {
                scrapeRes = await (pending
                    ? pending.promise
                    : requestScrape(tabId, {
                        ...baseScrapeOptions,
                        conversationId: conv.id,
                        conversationTitle: conv.title || null,
                        noDomFallback: true,
                    }));
            } catch (e: any) {
                const reason = e?.message || String(e);
                if (reason === 'cancelled') {
                    bundleStops.add(tabId);
                    pending = null;
                    break;
                }
                failures.push({ folderName, conversationId: conv.id, reason });
                // Keep the pipeline full: start the next chat's scrape before
                // moving on, unless a stop arrived while we were awaiting.
                pending = bundleStops.has(tabId) ? null : launchScrape(i + 1);
                continue;
            }

            // Chat i scraped OK. Start chat i+1's scrape NOW so it pages while
            // we build chat i. Skip if a stop landed between the await and here.
            pending = bundleStops.has(tabId) ? null : launchScrape(i + 1);

            const totalMessages = Array.isArray(scrapeRes.messages) ? scrapeRes.messages.length : 0;
            broadcastStatus({ tabId, phase: 'scrape:complete', messages: totalMessages, ...bundleCtx });
            // Record every scraped chat (including 0-message ones) for the
            // console-only stats block.
            const chatScrapeDoneAt = Date.now();
            getExportStats(tabId)?.addChat({
                title: conv.title || undefined,
                convId: conv.id,
                messages: totalMessages,
                scrapeMs: chatScrapeDoneAt - chatScrapeStartedAt,
            }, chatScrapeDoneAt);

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

            // Collect this chat's document attachments for the post-zip Files
            // pass (cheap: just hrefs + names, not the message bodies).
            if (buildOptions?.downloadFiles) {
                for (const c of collectDocumentAttachments(scrapeRes.messages || [])) {
                    bundleAttachItems.push({ ...c, chatFolder: folderName });
                }
            }

            try {
                broadcastStatus({ tabId, phase: 'build', messages: totalMessages, ...bundleCtx });
                const { files, formatFailures } = await buildOneChatForBundle({
                    messages: scrapeRes.messages || [],
                    meta: scrapeRes.meta || {},
                    formats,
                    embedAvatars,
                    downloadImages,
                    imageFilenameDate: Boolean(buildOptions.imageFilenameDate),
                    imageModifiedDate: Boolean(buildOptions.imageModifiedDate),
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
                // Per-format output sizes for the stats block. Each content
                // file is named `messages.<ext>`, so the extension IS the
                // format. Everything else (avatars/, images/) is summed into a
                // distinct 'assets' line so the size breakdown is complete.
                const chatStats = getExportStats(tabId);
                if (chatStats) {
                    let assetBytes = 0;
                    for (const f of files) {
                        const m = f.relativePath.match(/^messages\.([a-z]+)$/);
                        if (m) chatStats.addFormat(m[1], f.data.length);
                        else assetBytes += f.data.length;
                    }
                    if (assetBytes > 0) chatStats.addFormat('assets', assetBytes);
                }
                // Per-format resilience: keep the chat (with the formats that
                // built) when only some formats failed; drop it only if every
                // format failed. Each failed format is still listed in
                // FAILURES.txt so the gap is visible.
                const hasContent = files.some(f => f.relativePath.startsWith('messages.'));
                if (hasContent) {
                    // Stream this chat's files straight into the outer zip and
                    // let `files` go out of scope, instead of accumulating every
                    // chat's bytes in entries[] until the end. This drops the
                    // separate entries[] copy (the old ~GB held at ~300 chats);
                    // images/PDF then retain nothing (ZipPassThrough), though the
                    // deflated text formats stay in fflate's buffers until finish.
                    // Created lazily on the first success.
                    bundleZip = bundleZip ?? createZipStream('zip-outer-bundle');
                    try {
                        for (const f of files) {
                            await bundleZip.add(`${folderName}/${f.relativePath}`, f.data, f.mtime);
                        }
                    } catch (zipErr) {
                        // A zip-stream error poisons the archive — fatal for the
                        // whole bundle, not this one chat. Stop and surface it
                        // instead of silently downgrading to a no-zip FAILURES.txt.
                        bundleZipFatal = zipErr;
                        break;
                    }
                    entries.push({ folderName });
                }
                for (const ff of formatFailures) {
                    failures.push({ folderName, conversationId: conv.id, reason: ff.error });
                }
                // Record partial signal for bundle-level reporting. The
                // chat's individual files inside its folder already carry
                // their own banners + -PARTIAL filename suffix from the
                // single-chat code path; this accumulation drives the
                // bundle-root PARTIAL.txt and the outer zip's suffix.
                const perChatPartial = (scrapeRes.meta as { partial?: { reason: 'network' | 'truncation' } } | undefined)?.partial;
                if (perChatPartial) {
                    partials.push({ folderName, conversationId: conv.id, reason: perChatPartial.reason });
                }
            } catch (e: any) {
                failures.push({ folderName, conversationId: conv.id, reason: e?.message || String(e) });
            }
        }

        // A zip-stream error during the loop is fatal for the whole bundle (the
        // archive is unusable). Surface it as an error rather than a partial save.
        if (bundleZipFatal) {
            const message = `bundle zip failed: ${(bundleZipFatal as Error)?.message || String(bundleZipFatal)}`;
            broadcastStatus({ tabId, phase: 'error', error: message });
            sendResponse({ error: message });
            activeExports.delete(tabId); void persistActiveExports();
            return;
        }

        const cancelled = bundleStops.has(tabId);
        bundleStops.delete(tabId);

        // Cancellation is a hard abort: no partial zip, no history row.
        // The user said stop; producing half a bundle would be confusing
        // and risks shipping output the user didn't sanity-check.
        if (cancelled) {
            bundleCancelled = true;
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
            // entries.length > 0 here (the entries===0 cases returned above),
            // and every entries.push is preceded by creating bundleZip, so it
            // is non-null. Guard defensively anyway.
            if (!bundleZip) throw new Error('bundle zip stream missing despite built chats');
            const bundlePackageMode = Boolean(buildOptions?.downloadFiles);
            const buildRes = await buildAndDownloadBundlesZip(
                { downloads, isFirefox }, bundleZip, entries.length, failures, noHistory, partials,
                { packageFolder: bundlePackageMode },
            );
            if (bundlePackageMode && buildRes.id != null) {
                // Package mode: same confirmed-terminal wait as the single-chat
                // path — the Files phase must not run against an unsaved zip,
                // and there is no proceed-on-timeout.
                const waitAbort = new AbortController();
                trackFetch(tabId, waitAbort);
                let settled: SettledOutcome;
                try {
                    settled = await waitForDownloadSettled(
                        { downloads, onChanged: downloads.onChanged },
                        buildRes.id,
                        { signal: waitAbort.signal },
                    );
                } finally {
                    untrackFetch(tabId, waitAbort);
                }
                log('bundle export download settle ->', settled, 'id=', buildRes.id);
                if (settled === 'aborted') {
                    try { await downloads.cancel(buildRes.id); } catch { /* already terminal */ }
                    bundleCancelled = true;
                    broadcastStatus({ tabId, phase: 'cancelled' });
                    sendResponse({ cancelled: true });
                    return;
                }
                if (settled === 'interrupted') {
                    const reason = await downloadErrorReason(buildRes.id);
                    if (reason && reason !== 'USER_CANCELED') {
                        throw new Error(`Export download failed (${reason})`);
                    }
                    bundleCancelled = true;
                    broadcastStatus({ tabId, phase: 'cancelled' });
                    sendResponse({ cancelled: true });
                    return;
                }
                if (settled === 'stalled') throw new Error('Export download did not finish (no progress)');
                if (settled === 'missing') throw new Error('Export download disappeared before completing');
            } else {
                const downloadStatus = buildRes.id != null
                    ? await waitForDownloadComplete(buildRes.id, 60_000)
                    : 'complete';
                // Same rule as the single-chat path: only a real interruption is a
                // cancel. A 'timeout' (completion unconfirmed in the window) still
                // continues rather than dropping the completion handling.
                if (downloadStatus === 'interrupted') {
                    bundleCancelled = true;
                    broadcastStatus({ tabId, phase: 'cancelled' });
                    sendResponse({ cancelled: true });
                    return;
                }
                if (downloadStatus === 'timeout') {
                    log('bundle export download completion unconfirmed; proceeding for id', buildRes.id);
                }
            }

            // Stream collected document attachments to disk now that the outer
            // zip name (= export folder base) is known. Registered with the tab
            // so STOP_EXPORT cancels the in-flight resolver fetches.
            let filesSummary: AttachmentDownloadSummary | undefined;
            if (buildOptions?.downloadFiles && bundleAttachItems.length && buildRes.filename && !bundleStops.has(tabId)) {
                const exportFolder = buildRes.folder ?? buildRes.filename.replace(/\.zip$/i, '');
                const filesAbort = new AbortController();
                trackFetch(tabId, filesAbort);
                try {
                    // Every chat's attachments are dispatched here; per-chat
                    // request-access pages are cleaned up post-download by the
                    // onChanged verifier (each candidate carries its chatFolder,
                    // so FAILURES.txt lands in the right per-chat folder).
                    filesSummary = await downloadAttachments(
                        bundleAttachItems,
                        exportFolder,
                        {
                            downloads,
                            log,
                            onProgress: (filesDone, filesTotal) =>
                                broadcastStatus({ tabId, phase: 'downloading-files', filesDone, filesTotal, bundleTotalChats: totalChats }),
                        },
                        filesAbort.signal,
                        // Bound concurrency on Firefox (no per-host cap there) so the
                        // bundle's many files don't burst SharePoint into throttling.
                        { maxConcurrent: isFirefox ? 6 : undefined },
                    );
                } finally {
                    untrackFetch(tabId, filesAbort);
                }
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
                filesSummary: filesSummary && toFilesSummaryWire(filesSummary),
            });

            const elapsedMs = Math.max(0, Date.now() - startedAt);
            // Bundle-level history kind. If any chat in the bundle came
            // back partial, the bundle as a whole is partial — even if
            // the rest succeeded. The most conservative reason wins
            // when partials carry mixed reasons (network beats
            // truncation, same as single-chat).
            const bundlePartialReason: 'network' | 'truncation' | undefined = partials.length
                ? (partials.some(p => p.reason === 'network') ? 'network' : 'truncation')
                : undefined;
            const partialSuffix = partials.length ? ` (${partials.length} partial)` : '';
            const failedSuffix = failures.length ? ` (${failures.length} failed)` : '';
            await persistHistoryEntry({
                id: makeEntryId(),
                tabId,
                kind: bundlePartialReason ? 'partial' : 'success',
                partialReason: bundlePartialReason,
                downloadId: buildRes.id,
                filename: buildRes.filename,
                // Title carries chat count + status markers. messageCount is
                // omitted because the per-chat counts don't aggregate into a
                // single meaningful number for a bundle row.
                title: `${entries.length} chats${failedSuffix}${partialSuffix}`,
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
        } finally {
            // Derive the outcome from the accumulators rather than threading a
            // flag through every return: cancelled wins, then truly-empty, then
            // all-failed, otherwise complete (covers the no-history-only save).
            const outcome = bundleCancelled
                ? 'cancelled'
                : (entries.length === 0 && failures.length === 0 && noHistory.length === 0)
                    ? 'empty'
                    : (entries.length === 0 && failures.length > 0)
                        ? 'failed'
                        : 'complete';
            getExportStats(tabId)?.log(Date.now(), outcome, { verbose: diagVerboseStatsEnabled });
            endExportStats(tabId);
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
                    // Floor the wait at the configured backoff. Teams
                    // sometimes returns Retry-After: 0; honouring 0
                    // literally would spin retries without giving the
                    // rate limit a chance to clear.
                    const backoffWait = backoffMs(attempt);
                    const retryAfter = resp.headers.get('Retry-After');
                    const retryAfterMs = retryAfter ? Math.min(parseInt(retryAfter, 10) * 1000, 30_000) : 0;
                    const waitMs = Math.max(backoffWait, Number.isFinite(retryAfterMs) ? retryAfterMs : 0);
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

// Image-fetch-fallback feature gate. Background owns the check
// because Firefox MV2 content scripts can't reliably access
// permissions.contains. Content asks once at scrape start. Returns
// the AND of: (a) user has flipped the toggle ON in Settings, (b)
// <all_urls> permission is currently granted. Both can drift apart
// if the user revokes the host permission outside our UI; the AND
// keeps the feature gate honest.
async function handleFallbackStatusMessage(sendResponse: (resp: unknown) => void) {
    try {
        const stored = await storage.local.get('teamsExporterOptions');
        const opts = (stored as { teamsExporterOptions?: { imageFetchFallback?: boolean } })
            .teamsExporterOptions;
        const optionOn = !!opts?.imageFetchFallback;
        if (!optionOn) {
            sendResponse({ enabled: false });
            return;
        }
        // @ts-ignore - browser global on Firefox; chrome polyfill on Chrome
        const permsApi = typeof browser !== 'undefined' ? browser.permissions : chrome.permissions;
        const granted = await permsApi.contains({ origins: ['<all_urls>'] });
        sendResponse({ enabled: !!granted });
    } catch {
        sendResponse({ enabled: false });
    }
}

// Direct upstream fetch — used by the image-fetch-fallback feature when
// Teams' proxy returns a permanent-shaped failure on a thumbnail. No
// auth headers, no retries, single attempt with a 10 s timeout. Caller
// is content.ts:fetchImageAsDataUrl, which only invokes this when the
// user has both opted in via the Settings toggle AND granted
// <all_urls>. We re-check the permission here as a safety guard —
// permission state can change between scrape-start and a later fetch.
// Reuses the tab-grouped abort tracking so STOP_EXPORT also kills
// these in-flight fetches.
async function handleFetchBlobDirectMessage(
    msg: { url: string; maxBytes?: number; minBytes?: number },
    sendResponse: (resp: unknown) => void,
    tabId?: number,
) {
    try {
        // @ts-ignore - browser global on Firefox; chrome polyfill on Chrome
        const permsApi = typeof browser !== 'undefined' ? browser.permissions : chrome.permissions;
        const granted = await permsApi.contains({ origins: ['<all_urls>'] });
        if (!granted) {
            sendResponse({ ok: false, error: 'permission-revoked' });
            return;
        }
    } catch {
        sendResponse({ ok: false, error: 'permission-check-failed' });
        return;
    }
    const maxBytes = msg.maxBytes ?? 5 * 1024 * 1024;
    const minBytes = msg.minBytes ?? 100;
    const TIMEOUT_MS = 10_000;
    const controller = new AbortController();
    trackFetch(tabId, controller);
    const timeoutId = setTimeout(() => controller.abort(), TIMEOUT_MS);
    try {
        const resp = await fetch(msg.url, { signal: controller.signal });
        if (!resp.ok) {
            sendResponse({ ok: false, status: resp.status, statusText: resp.statusText });
            return;
        }
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
        sendResponse({ ok: true, dataUrl, size: blob.size });
    } catch (e) {
        if (controller.signal.aborted) {
            sendResponse({ ok: false, cancelled: true });
            return;
        }
        sendResponse({ ok: false, error: String(e) });
    } finally {
        clearTimeout(timeoutId);
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

    if (msg.type === 'GET_DIAGNOSTICS_BG') {
        // Copy the buffer so a later mutation doesn't affect what the
        // popup receives via structured clone. Include the persistence
        // state and on-disk byte usage so the popup can render the
        // storage info row without an extra round-trip.
        (async () => {
            // Default to null so a getBytesInUse failure (Firefox's
            // older polyfill rejects with "Not supported" instead of
            // returning a number) shows up as "unknown size" in the
            // UI rather than misleading "0 B".
            let bytesUsed: number | null = null;
            if (diagLogPersistEnabled) {
                try {
                    bytesUsed = await chrome.storage.local.getBytesInUse(DIAG_LOG_STORAGE_KEY);
                } catch {
                    // Firefox only added storage.local.getBytesInUse in v132; on
                    // older Firefox it rejects with "not supported". Fall back to
                    // the in-memory byte counter (the same buffer we persist) so
                    // the size shows on Firefox too instead of "?".
                    bytesUsed = diagBufferBytes;
                }
            } else {
                bytesUsed = 0;
            }
            sendResponse({
                entries: diagLogBuffer.slice(),
                bytesUsed,
                persistEnabled: diagLogPersistEnabled,
                lastFlushError: diagLastFlushError,
            });
        })();
        return true;
    }

    if (msg.type === 'DIAG_LOG_FORWARD') {
        // Content scripts batch their captures into small arrays and
        // forward them here. Append each to the master buffer with
        // src: 'content'. Fire-and-forget from the content side, so
        // we acknowledge but don't need to ship a payload back.
        if (Array.isArray(msg.entries)) {
            for (const e of msg.entries) {
                if (typeof e?.line !== 'string' || typeof e?.ts !== 'number' || typeof e?.level !== 'string') continue;
                pushDiagEntry({ src: 'content', ts: e.ts, level: e.level, line: e.line });
            }
        }
        sendResponse({ ok: true });
        return;
    }

    if (msg.type === 'CLEAR_DIAGNOSTICS_LOGS') {
        // Wipe in-memory AND on-disk. Cancels any pending flush, then
        // waits for any in-flight flush to finish before remove so the
        // post-await set() can't land after remove and resurrect the
        // just-cleared key. We also temporarily flip persistence off
        // for the duration of the clear: otherwise any push during
        // the `await remove` window would schedule a new flush that
        // fires after remove and writes a partial buffer back.
        const wasEnabled = diagLogPersistEnabled;
        diagLogPersistEnabled = false;
        diagLogBuffer.length = 0;
        diagBufferBytes = 0;
        diagBufferDirty = false;
        if (diagLogFlushTimer) { clearTimeout(diagLogFlushTimer); diagLogFlushTimer = null; }
        diagLastFlushError = null;
        (async () => {
            if (diagInflightFlush) { try { await diagInflightFlush; } catch { /* noop */ } }
            try { await chrome.storage.local.remove(DIAG_LOG_STORAGE_KEY); }
            catch { /* best-effort */ }
            diagLogPersistEnabled = wasEnabled;
            sendResponse({ ok: true });
        })();
        return true;
    }

    if (msg.type === 'SET_DIAG_LOG_PERSIST') {
        // Caller (popup) has already saved the option. We update the
        // BG runtime flag here so the change takes effect without
        // waiting for an SW restart. Turning off also wipes the
        // existing on-disk buffer (no point keeping stale data).
        diagLogPersistEnabled = !!msg.enabled;
        (async () => {
            if (diagLogPersistEnabled) {
                // Snapshot what we have so a refresh on the diagnostics
                // page right after toggling shows the right size.
                try {
                    await chrome.storage.local.set({ [DIAG_LOG_STORAGE_KEY]: diagLogBuffer.slice() });
                    diagLastFlushError = null;
                } catch (e: any) {
                    diagLastFlushError = { ts: Date.now(), reason: e?.message || String(e) };
                }
            } else {
                if (diagLogFlushTimer) { clearTimeout(diagLogFlushTimer); diagLogFlushTimer = null; }
                // Wait out any in-flight write before remove, otherwise
                // its post-await set() can resurrect the key. Same
                // ordering as CLEAR_DIAGNOSTICS_LOGS above.
                if (diagInflightFlush) { try { await diagInflightFlush; } catch { /* noop */ } }
                try { await chrome.storage.local.remove(DIAG_LOG_STORAGE_KEY); }
                catch { /* best-effort */ }
                diagLastFlushError = null;
            }
            sendResponse({ ok: true });
        })();
        return true;
    }

    if (msg.type === 'SET_DIAG_VERBOSE_STATS') {
        // The popup already saved the option; mirror it into the BG runtime
        // flag so the next export's [export-stats] line reflects it without
        // waiting for an SW restart. No buffer side effects.
        diagVerboseStatsEnabled = !!msg.enabled;
        sendResponse({ ok: true });
        return true;
    }

    if (msg.type === 'RUN_PROBES_BG') {
        // BG-side probes are cheap and don't need to be parallel.
        // service_worker_alive is implicit: if we got the message,
        // the SW is alive enough to answer.
        const t0 = Date.now();
        (async () => {
            const results: { name: string; status: 'pass' | 'fail' | 'skipped'; detail?: string; ms: number }[] = [
                { name: 'service_worker_alive', status: 'pass', ms: 0 },
            ];
            // <all_urls> grant state. Promise-based on both Chrome MV3
            // and Firefox WebExt; explicit null on error so the popup
            // doesn't conflate failure with "not granted".
            const t1 = Date.now();
            try {
                const granted = await (typeof browser !== 'undefined' ? browser.permissions : chrome.permissions)
                    .contains({ origins: ['<all_urls>'] });
                // Strict boolean check. A truthy non-true value (e.g.
                // an unusual WebExt polyfill return) should not be
                // reported as a clean PASS. Only annotate the type when
                // it's unexpected — a plain `false` is just "not
                // granted" without noise.
                const isGranted = granted === true;
                let detail: string;
                if (isGranted) {
                    detail = 'granted';
                } else if (typeof granted === 'boolean') {
                    detail = 'not granted';
                } else {
                    detail = `not granted (returned ${typeof granted})`;
                }
                results.push({
                    name: 'all_urls_granted',
                    status: isGranted ? 'pass' : 'fail',
                    detail,
                    ms: Date.now() - t1,
                });
            } catch (e: any) {
                results.push({
                    name: 'all_urls_granted',
                    status: 'fail',
                    detail: e?.message || String(e),
                    ms: Date.now() - t1,
                });
            }
            sendResponse({ ok: true, results, totalMs: Date.now() - t0 });
        })();
        return true;
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
        // Stage timeline: the fetch sub-phases (api-fetch / images / avatars,
        // plus the DOM-fallback scroll / extract) only pass through here.
        getExportStats(senderTabId)?.markPhase(payload?.phase, Date.now());
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

    if (msg.type === 'FETCH_BLOB_DIRECT') {
        handleFetchBlobDirectMessage(msg, sendResponse, sender?.tab?.id);
        return true;
    }

    if (msg.type === 'FALLBACK_STATUS') {
        handleFallbackStatusMessage(sendResponse);
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

    // Files phase: the export artifact is ALREADY on disk; Stop should stop
    // the file downloads, not disown the export. Abort the tab's controllers
    // (breaks the settlement wait; the run cancels its in-flight downloads,
    // which land in FAILURES.txt's cancelled section) and let the export flow
    // emit the single coherent terminal — 'complete' with the cancelled-file
    // counts — instead of writing a 'cancelled' history row here that would
    // pair with the flow's success row (double-row) and claim nothing saved.
    // 'cancelling' is matched too so a second Stop press during the wind-down
    // stays idempotent instead of falling through to the generic path (which
    // would write a cancelled row alongside the flow's coming terminal).
    if (existing?.phase === 'downloading-files' || existing?.phase === 'cancelling') {
        log('STOP_EXPORT during Files phase — stopping file downloads, keeping the export');
        broadcastStatus({ tabId, phase: 'cancelling' });
        const filesAborted = abortFetchesForTab(tabId);
        if (filesAborted > 0) log(`aborted ${filesAborted} in-flight fetches for tab ${tabId}`);
        sendResponse({ ok: true });
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
