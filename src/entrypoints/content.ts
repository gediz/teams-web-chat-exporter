import { defineContentScript } from 'wxt/sandbox';
import { $ } from '../utils/dom';
import { uint8ToBase64 } from '../utils/base64';
import { makeDayDivider as buildDayDivider } from '../utils/messages';
import { cssEscape, isPlaceholderText, textFrom } from '../utils/text';
import { parseTimeStamp } from '../utils/time';
import { extractAttachments } from '../content/attachments';
import { resolveShareFile, resolveWithFallback } from '../content/share-resolver';
import { extractReactions } from '../content/reactions';
import { extractReplyContext } from '../content/replies';
import { cleanAltText, extractTables, extractTextWithEmojis, normalizeMentions } from '../content/text';
import { autoScrollAggregate as autoScrollAggregateHelper } from '../content/scroll';
import { extractChatTitle, extractChannelTitle } from '../content/title';
import { extractAvatarId } from '../utils/avatars';
import { TEAMS_MATCH_PATTERNS } from '../utils/teams-urls';
import { apiScrape, discover, ensureSkypeTokenCookies, extractConversationId, fetchSharePointFile, getGraphToken, getIc3Token, getLastApiScrapeFailure, getSkypeToken, listConversationsFromIdb, listConversationsFromIdbQuick, resolveAvatarPhotos } from '../content/api-client';
import { convertApiMessages } from '../content/api-converter';
import { runStandaloneProbes } from '../content/probes';
import type { ProbeResult } from '../utils/diagnostics';
import type { AggregatedItem, ExportMessage, OrderContext, ReplyContext, ScrapeOptions } from '../types/shared';

// Typed globals for Firefox builds
declare const browser: typeof chrome | undefined;

// Resolve after `ms`, or immediately if `signal` aborts first, so a Stop click
// during a backoff sleep isn't stuck waiting out the full timer.
function sleepOrAbort(ms: number, signal?: AbortSignal): Promise<void> {
    return new Promise(resolve => {
        if (signal?.aborted) { resolve(); return; }
        let timer: ReturnType<typeof setTimeout>;
        const onAbort = () => { clearTimeout(timer); resolve(); };
        timer = setTimeout(() => { signal?.removeEventListener('abort', onAbort); resolve(); }, ms);
        signal?.addEventListener('abort', onAbort, { once: true });
    });
}

// Map an HTTP status from a failed image/file fetch to one short reason word
// for the placeholder, the summary banner, and the failed-items manifest (see
// Attachment.failReason). 410 expired / 404 removed / 403 no-access / 401
// sign-in / 429 rate-limited / 5xx or 408 server-error / else unavailable.
function statusToReason(status: number): string {
  if (status === 410) return 'expired';
  if (status === 404) return 'removed';
  if (status === 403) return 'no-access';
  if (status === 401) return 'sign-in';
  if (status === 429) return 'rate-limited';
  if (status === 408 || status >= 500) return 'server-error';
  return 'unavailable';
}

// Bound concurrent shares-API resolves from the page (defence-in-depth on top
// of the background's resolve cap). A tiny FIFO semaphore.
const SHARE_RESOLVE_MAX = 4;
let shareResolveActive = 0;
const shareResolveQueue: Array<() => void> = [];
async function withShareResolveSlot<T>(fn: () => Promise<T>): Promise<T> {
  if (shareResolveActive >= SHARE_RESOLVE_MAX) {
    await new Promise<void>(res => shareResolveQueue.push(res));
  }
  shareResolveActive++;
  try { return await fn(); }
  finally {
    shareResolveActive--;
    shareResolveQueue.shift()?.();
  }
}

type ExtractedMessage = ExportMessage & {
    id: string;
    threadId?: string | null;
    author: string;
    timestamp: string;
    text: string;
    edited: boolean;
    avatar: string | null;
};

type ContentAggregated = AggregatedItem & { message?: ExtractedMessage };
type InternalScrapeOptions = ScrapeOptions & { __allowInlineThreadReplies?: boolean };

export default defineContentScript({
    matches: TEAMS_MATCH_PATTERNS,
    runAt: 'document_idle',
    allFrames: true,

    main() {
        const isTop = window.top === window;

        // Browser API compatibility for Firefox
        const runtime = typeof browser !== 'undefined' ? browser.runtime : chrome.runtime;

        // Diagnostic log tail.
        //
        // Content captures forward to BG (which owns the master buffer)
        // as small batched arrays. Each console wrap call enqueues a
        // line; a 100 ms debounced timer ships the accumulated batch.
        // Fire-and-forget so the hot path doesn't await IPC.
        //
        // No in-memory retention here. The previous design kept a
        // local ring buffer too; now the source of truth is BG, which
        // means the popup gets a consistent merged view from a single
        // place. The cost is one runtime.sendMessage per batch (a
        // handful per export).
        //
        // Guarded so a second module load (HMR, re-injection) doesn't
        // double-wrap and emit duplicates.
        const DIAG_PREFIX_RE = /^\[[A-Za-z][A-Za-z0-9 _\-:]{0,40}\]/;
        const DIAG_FORWARD_DEBOUNCE_MS = 100;
        const pendingDiagEntries: { ts: number; level: string; line: string }[] = [];
        let diagForwardTimer: ReturnType<typeof setTimeout> | null = null;
        // Counters so a silent forwarding failure (BG unreachable, SW
        // evicted, port dropped) is visible in the diagnostic report
        // instead of silently disappearing. The popup pulls these via
        // GET_DIAGNOSTICS_CONTENT and the JSON report carries them in
        // the logs section.
        let diagLostBatches = 0;
        let diagLostEntries = 0;
        let diagLastForwardError: string | null = null;
        function scheduleDiagForward() {
            if (diagForwardTimer) return;
            diagForwardTimer = setTimeout(() => {
                diagForwardTimer = null;
                if (pendingDiagEntries.length === 0) return;
                const batch = pendingDiagEntries.splice(0);
                const recordLoss = (err: unknown) => {
                    diagLostBatches++;
                    diagLostEntries += batch.length;
                    diagLastForwardError = err instanceof Error
                        ? err.message
                        : (typeof err === 'string' ? err : 'forwarding failed');
                };
                try {
                    const p = runtime.sendMessage({ type: 'DIAG_LOG_FORWARD', entries: batch });
                    if (p && typeof p.then === 'function') {
                        (p as Promise<unknown>).catch(recordLoss);
                    }
                } catch (err) { recordLoss(err); }
            }, DIAG_FORWARD_DEBOUNCE_MS);
        }
        function diagForwardingStats() {
            return { lostBatches: diagLostBatches, lostEntries: diagLostEntries, lastError: diagLastForwardError };
        }
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
                    if (level !== 'error') {
                        const first = args[0];
                        if (typeof first !== 'string' || !DIAG_PREFIX_RE.test(first)) return;
                    }
                    const line = args.map(formatArg).join(' ');
                    pendingDiagEntries.push({ ts: Date.now(), level, line });
                    scheduleDiagForward();
                } catch { /* never break logging */ }
            };
            console.log = (...args: unknown[]) => { capture('log', args); origLog(...args); };
            console.warn = (...args: unknown[]) => { capture('warn', args); origWarn(...args); };
            console.error = (...args: unknown[]) => { capture('error', args); origError(...args); };
            console.info = (...args: unknown[]) => { capture('info', args); origInfo(...args); };
            console.debug = (...args: unknown[]) => { capture('debug', args); origDebug(...args); };
        }

        // IDB shape probe used by the diagnostics page. Counts rows per
        // store without enumerating records, so no chat data is read.
        // Names embed tenant + user UUID; redaction is applied at the
        // popup's report-build step, not here.
        type DbOpenResult =
            | { status: 'opened'; db: IDBDatabase }
            | { status: 'blocked' }
            | { status: 'error'; reason: string };
        // Hard ceiling on a single DB open. An indexedDB.open that fires
        // none of success/error/blocked is a known browser-bug shape we
        // do not want hanging the entire diagnostics response. 5 s is
        // generous: a healthy open completes in milliseconds, a busy
        // one resolves via onblocked.
        const PROBE_OPEN_TIMEOUT_MS = 5000;
        async function openProbeDb(name: string): Promise<DbOpenResult> {
            return new Promise(resolve => {
                let settled = false;
                const settle = (r: DbOpenResult) => {
                    if (settled) return;
                    settled = true;
                    clearTimeout(timer);
                    resolve(r);
                };
                const timer = setTimeout(() => {
                    settle({ status: 'error', reason: `open timed out after ${PROBE_OPEN_TIMEOUT_MS}ms` });
                }, PROBE_OPEN_TIMEOUT_MS);
                try {
                    const req = indexedDB.open(name);
                    req.onsuccess = () => {
                        const db = req.result;
                        // If we already resolved (blocked or timeout path),
                        // close the late-arriving handle so we don't pin
                        // the DB open.
                        if (settled) { try { db.close(); } catch { /* noop */ } return; }
                        // Cooperate with Teams: if it requests an upgrade
                        // while we're holding the handle, release it.
                        db.onversionchange = () => { try { db.close(); } catch { /* noop */ } };
                        settle({ status: 'opened', db });
                    };
                    req.onerror = () => settle({ status: 'error', reason: req.error?.message || 'open failed' });
                    req.onblocked = () => settle({ status: 'blocked' });
                } catch (e) {
                    settle({ status: 'error', reason: e instanceof Error ? e.message : String(e) });
                }
            });
        }
        async function probeIdbShape(): Promise<
            | {
                available: true;
                databases: {
                    name: string;
                    version: number;
                    status: 'opened' | 'blocked' | 'error';
                    reason?: string;
                    stores: { name: string; count: number; error?: string }[];
                }[];
              }
            | { available: false; reason: string }
        > {
            try {
                if (typeof indexedDB === 'undefined' || typeof indexedDB.databases !== 'function') {
                    return { available: false, reason: 'indexedDB.databases unsupported' };
                }
                const all = await indexedDB.databases();
                const teamsDbs = all.filter(d => typeof d.name === 'string' && d.name.startsWith('Teams:'));
                const out: {
                    name: string;
                    version: number;
                    status: 'opened' | 'blocked' | 'error';
                    reason?: string;
                    stores: { name: string; count: number; error?: string }[];
                }[] = [];
                for (const meta of teamsDbs) {
                    const name = meta.name as string;
                    const version = typeof meta.version === 'number' ? meta.version : 0;
                    const opened = await openProbeDb(name);
                    if (opened.status !== 'opened') {
                        out.push({ name, version, status: opened.status, reason: 'reason' in opened ? opened.reason : undefined, stores: [] });
                        continue;
                    }
                    const db = opened.db;
                    const stores: { name: string; count: number; error?: string }[] = [];
                    for (const sn of Array.from(db.objectStoreNames)) {
                        const r = await new Promise<{ count: number; error?: string }>(resolve => {
                            try {
                                const tx = db.transaction(sn, 'readonly');
                                const req = tx.objectStore(sn).count();
                                req.onsuccess = () => resolve({ count: req.result || 0 });
                                req.onerror = () => resolve({ count: 0, error: req.error?.message || 'count failed' });
                            } catch (e) {
                                resolve({ count: 0, error: e instanceof Error ? e.message : String(e) });
                            }
                        });
                        stores.push({ name: sn, count: r.count, ...(r.error ? { error: r.error } : {}) });
                    }
                    try { db.close(); } catch { /* noop */ }
                    out.push({ name, version, status: 'opened', stores });
                }
                return { available: true, databases: out };
            } catch (e) {
                return { available: false, reason: e instanceof Error ? e.message : String(e) };
            }
        }

        let currentRunStartedAt: number | null = null;
        // One scrape can be active per content script. Holding the controller
        // here lets the STOP_SCRAPE handler abort whatever fetch loop is in
        // flight (IC3 pagination, image fetches, Graph photo fetches).
        let currentAbortController: AbortController | null = null;

        // Chat/channel detection intentionally relies ONLY on data-tid
        // attributes, never on aria-label text. Issues #10 and #19 reported
        // that English-hardcoded checks (e.g. aria-label^="Chat") broke the
        // extension for every non-English Teams UI — French ("Conversation"),
        // Japanese, German, etc. data-tid values are code identifiers baked
        // into Teams' build and don't change with UI language.
        //
        // Trade-off: we no longer surface a distinct "switch to the Chat
        // app in Teams" error for users who have the wrong nav pane
        // selected — we just report "open a chat before exporting", which
        // is true in every case where the surface isn't present. That
        // combined error is good enough; hyper-specific guidance isn't
        // worth re-introducing the locale bug.

        function hasChatMessageSurface() {
            return Boolean(
                document.querySelector('[data-tid="message-pane-list-viewport"], [data-tid="chat-message-list"], [data-tid="chat-pane"]')
            );
        }

        function hasChannelMessageSurface() {
            return Boolean(
                document.querySelector('[data-tid="channel-pane-runway"], [data-tid="channel-pane-message"], [data-tid="channel-pane"]')
            );
        }

        function checkChatContext(target: 'chat' | 'team' = 'chat') {
            if (target === 'team') {
                if (hasChannelMessageSurface()) return { ok: true };
                return { ok: false, reason: 'Open a team channel before exporting.' };
            }
            if (hasChatMessageSurface()) return { ok: true };
            return { ok: false, reason: 'Open a chat conversation before exporting.' };
        }

        // Core DOM hooks ------------------------------------------------
        function findScrollableAncestor(node: Element | null): Element | null {
            let current: Element | null = node;
            while (current) {
                const el = current as HTMLElement;
                const style = window.getComputedStyle(el);
                const overflowY = style.overflowY;
                if ((overflowY === 'auto' || overflowY === 'scroll' || overflowY === 'overlay') && el.scrollHeight > el.clientHeight) {
                    return el;
                }
                current = current.parentElement;
            }
            return null;
        }

        function isElementVisible(el: Element | null): el is HTMLElement {
            if (!el) return false;
            const style = window.getComputedStyle(el);
            if (style.display === 'none' || style.visibility === 'hidden') return false;
            const rect = (el as HTMLElement).getBoundingClientRect();
            return rect.width > 0 && rect.height > 0;
        }

        function getScroller(target: 'chat' | 'team' = 'chat') {
            if (target === 'team') {
                const viewport = document.querySelector<HTMLElement>('[data-tid="channel-pane-viewport"]');
                if (viewport && isElementVisible(viewport) && viewport.scrollHeight > viewport.clientHeight) {
                    return viewport;
                }
                const runway = findChannelRunway();
                if (runway) {
                    return findScrollableAncestor(runway) || document.scrollingElement;
                }
                const anchors = [
                    $('[data-tid="channel-pane-runway"]'),
                    $('[data-testid="virtual-list-loader"]'),
                    $('[data-testid="vl-placeholders"]'),
                    document.querySelector('[id^="channel-pane-"]'),
                ];
                const anchor = anchors.find(isElementVisible) || anchors.find(Boolean) || null;
                return findScrollableAncestor(anchor) || document.scrollingElement;
            }
            return $('[data-tid="message-pane-list-viewport"]') || $('[data-tid="chat-message-list"]') || document.scrollingElement;
        }


        function getAllDocs(): Document[] {
            const docs: Document[] = [document];
            // Try same-origin frames
            for (let i = 0; i < window.frames.length; i++) {
              try {
                const d = window.frames[i].document;
                if (d) docs.push(d);
              } catch {
                // cross-origin or inaccessible frame
              }
            }
            return docs;
          }
          
          function qAny<T extends Element = Element>(selector: string): T | null {
            for (const d of getAllDocs()) {
              const el = d.querySelector(selector) as T | null;
              if (el) return el;
            }
            return null;
          }
          

        function findChannelRunway(): Element | null {
            const explicit = Array.from(document.querySelectorAll<HTMLElement>('[data-tid="channel-pane-runway"]'));
            const visibleExplicit = explicit.find(isElementVisible);
            if (visibleExplicit) return visibleExplicit;
            if (explicit.length) return explicit[0];
            const candidates = Array.from(document.querySelectorAll<HTMLElement>('[id^="channel-pane-"]'));
            if (!candidates.length) return null;
            const filtered = candidates.filter(el => {
                if (el.getAttribute('data-tid') === 'channel-replies-runway') return false;
                if (el.id === 'channel-pane-l2') return false;
                return Boolean(el.querySelector('[data-tid="channel-pane-message"]'));
            });
            const visibleFiltered = filtered.filter(isElementVisible);
            if (visibleFiltered.length) return visibleFiltered[0];
            if (filtered.length) return filtered[0];
            const visibleCandidate = candidates.find(isElementVisible);
            return visibleCandidate || candidates[0] || null;
        }

        function getChannelItems(): Element[] {
            const runway = findChannelRunway();
            const listItems = runway ? Array.from(runway.querySelectorAll('li[role="none"]')) : [];
        
            let items: Element[] = [];
        
            if (runway) {
                const selectors = [
                    '[id^="message-body-"][aria-labelledby]',
                    '[data-tid="control-message-renderer"]',
                    '.fui-Divider__wrapper',
                ];
                const direct = Array.from(runway.querySelectorAll<HTMLElement>(selectors.join(', ')));
                if (direct.length) {
                    items = direct;
                }
            }
        
            if (!items.length) {
                const filtered = listItems.filter(item =>
                    item.querySelector('[data-tid="channel-pane-message"], [data-tid="control-message-renderer"], .fui-Divider__wrapper'),
                );
                if (filtered.length) {
                    items = filtered;
                } else {
                    items = Array.from(document.querySelectorAll('[data-tid="channel-pane-message"]'));
                }
            }
        
            return items;
        }
        

        function isVirtualListLoading(): boolean {
            const runway = findChannelRunway();
            const loader =
                runway?.parentElement?.querySelector<HTMLElement>('[data-testid="virtual-list-loader"]') ||
                runway?.querySelector<HTMLElement>('[data-testid="virtual-list-loader"]') ||
                document.querySelector<HTMLElement>('[data-testid="virtual-list-loader"]');
            if (loader && loader.offsetParent !== null) {
                const rect = loader.getBoundingClientRect();
                if (rect.height >= 1 || rect.width >= 1) return true;
            }
            return false;
        }

        // Author/timestamp/edited/avatar helpers ------------------------
        function resolveAuthor(body: Element, lastAuthor = ""): string {
            let author = textFrom($('[data-tid="message-author-name"]', body));
            if (!author) {
                const embedded = body.querySelector<HTMLElement>('[id^="author-"]');
                if (embedded) author = textFrom(embedded);
            }
            if (!author) {
                const aria = body.getAttribute('aria-labelledby') || '';
                const aId = aria.split(/\s+/).find(s => s.startsWith('author-'));
                if (aId) author = textFrom(document.getElementById(aId));
            }
            return author || lastAuthor || '';
        }
        function resolveTimestamp(item: Element): string {
            const t = $('time[datetime]', item) || $('time', item) || $('[data-tid="message-status"] time', item);
            return t?.getAttribute?.('datetime') || t?.getAttribute?.('title') || t?.getAttribute?.('aria-label') || textFrom(t) || '';
        }
        function resolveEdited(item: Element, body: Element): boolean {
            const aria = body?.getAttribute('aria-labelledby') || '';
            const editedId = aria.split(/\s+/).find(s => s.startsWith('edited-'));
            if (editedId) {
                const el = document.getElementById(editedId);
                if (el) {
                    const txt = (el.textContent || el.getAttribute('title') || '').trim();
                    if (/^edited\b/i.test(txt)) return true; // real badge only
                }
            }
            const badge = item.querySelector('[id^="edited-"]');
            if (badge) {
                const txt = (badge.textContent || badge.getAttribute('title') || '').trim();
                if (/^edited\b/i.test(txt)) return true;
            }
            return false;
        }
        function imgToDataURL(img: HTMLImageElement): string | null {
            try {
                if (!img.complete || !img.naturalWidth || !img.naturalHeight) return null;
                const canvas = document.createElement('canvas');
                canvas.width = img.naturalWidth;
                canvas.height = img.naturalHeight;
                const ctx = canvas.getContext('2d');
                if (!ctx) return null;
                ctx.drawImage(img, 0, 0);
                return canvas.toDataURL('image/png');
            } catch {
                return null;
            }
        }

        // The user's own messages don't include an avatar element in the DOM,
        // so capture it once from the top-bar MeControl and patch it in by UUID identity.
        function findSelfAvatar(): { dataUrl: string | null; name: string | null; uuid: string | null } {
            const trigger = document.querySelector('[data-tid="me-control-avatar"]') as HTMLElement | null;
            if (!trigger) return { dataUrl: null, name: null, uuid: null };
            const aria = trigger.getAttribute('aria-label') || '';
            const nameMatch = aria.match(/Profile picture of (.+?)\.?$/i);
            const name = nameMatch?.[1]?.trim() || null;
            const img = trigger.querySelector('img') as HTMLImageElement | null;
            const dataUrl = img ? imgToDataURL(img) : null;
            const uuid = extractUserUuidFromAvatarUrl(img?.src);
            return { dataUrl, name, uuid };
        }

        // Returns both the original profile-picture URL (carries the user's UUID for
        // identity) and a canvas-rendered data URL (the actual pixels we'll embed).
        type ResolvedAvatar = { url: string; dataUrl: string | null };

        function resolveAvatar(item: Element): ResolvedAvatar | null {
            // Try per-message avatar with various selectors
            const selectors = [
                '[data-tid="message-avatar"] img',
                '[data-tid="avatar"] img',
                '.fui-Avatar img',
                '[class*="avatar" i] img',
                'img[src*="profilepicture"]'
            ];

            const searchIn = (root: Element): ResolvedAvatar | null => {
                for (const selector of selectors) {
                    const img = $(selector, root) as HTMLImageElement | null;
                    if (img?.src && img.src.startsWith('http')) {
                        // Only accept individual user avatars (profilepicturev2), not group avatars
                        if (img.src.includes('/profilepicturev2/') || img.src.includes('/profilepicture/')) {
                            // Read pixels from the already-rendered <img>: Firefox content scripts
                            // can't send page cookies on fetch(), so the URL path 401s there.
                            return { url: img.src, dataUrl: imgToDataURL(img) };
                        }
                    }
                }
                return null;
            };

            const found = searchIn(item);
            if (found) return found;

            // In the new Teams DOM the avatar is a sibling of the message body, not inside it.
            // extractOne narrows itemScope to the body; climb to the outer wrapper to find it.
            const outer =
                item.closest('[data-tid="chat-pane-item"]') ||
                item.closest('li[role="none"]') ||
                null;
            if (outer && outer !== item) return searchIn(outer);

            // No per-message avatar found - return null (don't use group/header fallback)
            return null;
        }

        // Profile picture URLs embed the user's UUID at /users/<uuid>/profilepicturev2/.
        // Use this UUID as a stable identity to disambiguate users with identical names.
        function extractUserUuidFromAvatarUrl(url: string | null | undefined): string | null {
            if (!url) return null;
            const m = url.match(/\/users\/([a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12})\//i);
            return m?.[1] || null;
        }

        /**
        * Fetches avatar images and converts them to base64 data URLs.
        * This runs in the content script context which has access to Teams cookies.
        */
        async function fetchAvatarAsDataURL(url: string): Promise<string | null> {
            try {
                const res = await fetch(url, { credentials: 'include' });
                if (!res.ok) {
                    // log (not warn): per-avatar HTTP failures are
                    // external state (deleted user, expired blob URL,
                    // etc.) — not extension code errors. Avoids
                    // flooding Chrome's Errors panel for routine
                    // upstream unavailability.
                    console.log(`[Avatar Fetch] HTTP ${res.status} for ${url.substring(0, 100)}...`);
                    return null;
                }
                const blob = await res.blob();
                return new Promise<string | null>((resolve) => {
                    const reader = new FileReader();
                    reader.onloadend = () => resolve(reader.result as string);
                    reader.onerror = () => resolve(null);
                    reader.readAsDataURL(blob);
                });
            } catch (err) {
                console.error(`[Avatar Fetch] Failed for ${url.substring(0, 100)}...`, err);
                return null;
            }
        }

        /**
        * Fetches avatars and returns normalized messages with avatarId references
        * plus a deduplicated avatars map. This keeps the sendResponse payload small.
        */
        async function embedAvatarsInContent(messages: ExtractedMessage[]): Promise<{ messages: ExtractedMessage[]; avatars: Record<string, string> }> {
            const uniqueUrls = new Set<string>();
            // Collect data URL avatars from API mode (already fetched)
            const dataUrlAvatars = new Map<string, string>(); // author → dataUrl
            for (const m of messages) {
                if (m.avatar && m.avatar.startsWith('data:')) {
                    if (m.author && !dataUrlAvatars.has(m.author)) {
                        dataUrlAvatars.set(m.author, m.avatar);
                    }
                } else if (m.avatar) {
                    uniqueUrls.add(m.avatar);
                }
            }

            // Fetch unique HTTP avatar URLs (DOM mode)
            const avatars: Record<string, string> = {};
            const urlToId = new Map<string, string>();
            for (const url of uniqueUrls) {
                if (currentAbortController?.signal.aborted) break;
                const dataUrl = await fetchAvatarAsDataURL(url);
                if (dataUrl) {
                    const id = extractAvatarId(url);
                    avatars[id] = dataUrl;
                    urlToId.set(url, id);
                }
            }

            // Normalize API-mode data URL avatars into the avatars map
            let apiAvatarIdx = 0;
            const authorToId = new Map<string, string>();
            for (const [author, dataUrl] of dataUrlAvatars) {
                const id = `api-avatar-${apiAvatarIdx++}`;
                avatars[id] = dataUrl;
                authorToId.set(author, id);
            }

            // Replace avatar URLs/data with short avatarId references
            const normalized = messages.map(m => {
                if (!m.avatar) return m;
                if (m.avatar.startsWith('data:')) {
                    // API-mode: replace data URL with avatarId
                    const id = m.author ? authorToId.get(m.author) : undefined;
                    if (id) return { ...m, avatar: null, avatarId: id };
                    return { ...m, avatar: null };
                }
                // DOM-mode: replace HTTP URL with avatarId
                const id = urlToId.get(m.avatar);
                if (id) return { ...m, avatar: null, avatarId: id };
                return { ...m, avatar: null };
            });

            // Reverse map so author / self / reactor avatars share ids: a given
            // dataUrl is registered (or reused) once, no duplicate-byte entries.
            const dataUrlToId = new Map<string, string>();
            for (const id in avatars) dataUrlToId.set(avatars[id], id);
            const idForDataUrl = (dataUrl: string): string => {
                let id = dataUrlToId.get(dataUrl);
                if (!id) {
                    id = `api-avatar-${apiAvatarIdx++}`;
                    avatars[id] = dataUrl;
                    dataUrlToId.set(dataUrl, id);
                }
                return id;
            };

            // Self reactor avatars: the self user often reacts in chats where
            // they never sent a message. Reuse the captured self avatar (no
            // network), preferring an existing author entry.
            const self = findSelfAvatar();
            const selfAvatarId: string | undefined =
                (self.name ? authorToId.get(self.name) : undefined)
                || (self.dataUrl ? idForDataUrl(self.dataUrl) : undefined);

            type ReactorLike = { name: string; avatarId?: string; self?: boolean; uuid?: string };
            const eachReactor = (fn: (reactor: ReactorLike) => void) => {
                for (const m of normalized) {
                    const reactions = (m as { reactions?: unknown[] }).reactions;
                    if (!Array.isArray(reactions)) continue;
                    for (const r of reactions) {
                        const reactors = (r as { reactors?: ReactorLike[] }).reactors;
                        if (Array.isArray(reactors)) for (const reactor of reactors) fn(reactor);
                    }
                }
            };

            // Pass 1: resolve reactor avatars by self, then by author display
            // name. Collect the still-unresolved reactor UUIDs for a Graph fetch.
            const unresolved = new Set<string>();
            eachReactor(reactor => {
                if (reactor.avatarId) return;
                if (reactor.self && selfAvatarId) { reactor.avatarId = selfAvatarId; return; }
                const id = authorToId.get(reactor.name);
                if (id) { reactor.avatarId = id; return; }
                if (reactor.uuid) unresolved.add(reactor.uuid);
            });

            // Pass 2: Graph-fetch photos for the remaining (non-self, non-author)
            // reactors, deduped by UUID, capped, with 429 backoff.
            if (unresolved.size) {
                const photos = await fetchReactorPhotos([...unresolved]);
                if (photos.size) {
                    eachReactor(reactor => {
                        if (reactor.avatarId || !reactor.uuid) return;
                        const dataUrl = photos.get(reactor.uuid);
                        if (dataUrl) reactor.avatarId = idForDataUrl(dataUrl);
                    });
                }
            }

            return { messages: normalized, avatars };
        }

        // ── Inline Image Fetching (API mode) ──────────────────────
        const MAX_IMAGE_BYTES = 5 * 1024 * 1024; // 5 MB
        // Opt-in full-res images use a higher per-image cap. NOTE: this is a
        // post-download REJECTION gate (the body is fully fetched, then its size
        // is checked), not a download limiter — the real peak-heap bound is the
        // reduced concurrency below.
        const FULLRES_MAX_IMAGE_BYTES = 20 * 1024 * 1024; // 20 MB
        // Active per-image cap for the AMS image fetchers; set once per export in
        // fetchInlineImages (one export at a time, set before the concurrent
        // workers start, so reads are race-free). Avatars/SharePoint keep MAX_IMAGE_BYTES.
        let imgFetchMaxBytes: number = MAX_IMAGE_BYTES;
        const MAX_CONCURRENT_FETCHES = 12; // [limit-test] was 6; raise to probe the ceiling
        // Full-res images are up to 20MB each and the cap is post-download, so a
        // lower pool bounds peak heap when full-res is on.
        const FULLRES_CONCURRENT_FETCHES = 4;
        // When Teams' urlp proxy starts returning HTTP 429 we drop concurrency
        // hard for the rest of the export so the per-session rate limit can
        // clear and the per-image retries below (MAX_429_RETRIES) can recover
        // the tail instead of dropping it. Only engages after the first 429, so
        // exports that never rate-limit stay at MAX_CONCURRENT_FETCHES. 2 was
        // validated against a heavily-rate-limited Firefox multi-chat storm; 4
        // was not gentle enough to let the limit clear while retries fired.
        const MAX_CONCURRENT_FETCHES_THROTTLED = 2;
        // Per-image HTTP 429 retry budget. A single retry dropped the tail of
        // images under sustained rate-limiting at bundle scale; retry a few
        // times with escalating backoff (see below) so they recover.
        const MAX_429_RETRIES = 4;

        // Verbose image-fetch debugging. Two ways to enable:
        //   1. Build-time:  WXT_DEBUG_IMAGE_FETCH=1 pnpm build
        //   2. Runtime:     localStorage.setItem('__teams_exporter_debug_image_fetch', '1')
        //                   on the Teams tab, then re-export.
        // When on, we log:
        //   - JWT claims of the IC3 token (aud / iss / exp / tid / appid)
        //   - First 5 sample URLs (raw + transformed)
        //   - A _image-fetch-debug.txt sidecar in the export with all failed URLs
        // We never log the token itself, only its decoded claim payload.
        const DEBUG_IMAGE_FETCH: boolean = (() => {
            try {
                if ((import.meta as { env?: Record<string, unknown> }).env?.WXT_DEBUG_IMAGE_FETCH === '1') return true;
            } catch { /* import.meta in some contexts */ }
            try { return localStorage.getItem('__teams_exporter_debug_image_fetch') === '1'; }
            catch { return false; }
        })();
        if (DEBUG_IMAGE_FETCH) {
            console.log('[Teams Exporter DEBUG] image-fetch verbose mode active');
        }

        // userRegion comes from the chat-service discovery (e.g. "emea"); the proxy
        // host uses a different naming convention. Fall back to "eu" if unmapped.
        const USER_REGION_TO_PROXY_PREFIX: Record<string, string> = {
            emea: 'eu', amer: 'na', apac: 'as', uk: 'uk', au: 'au', in: 'in', jp: 'jp',
        };

        // Teams's authenticated image proxy. AMS URLs in message HTML can't be fetched
        // directly without per-domain Skype cookies; the proxy accepts our IC3 Bearer
        // token (same one used for chat service) and returns the same image bytes.
        function transformImageUrlToProxy(url: string, userId: string, userRegion: string): string | null {
            const targetHost = `${USER_REGION_TO_PROXY_PREFIX[userRegion.toLowerCase()] || 'eu'}-prod.asyncgw.teams.microsoft.com`;
            // AMS direct: https://[region]-api.asm.skype.com/v1/(objects/...)
            // Transforms to: https://{targetHost}/v1/{userId}/(rest)?v=1
            // [a-z0-9-]+ not just [a-z-]+: real region hostnames can contain digits
            // (eu-api-0, amer-api-1, euno1-api-0, etc.).
            const ams = url.match(/^https:\/\/[a-z0-9-]+\.asm\.skype\.com\/v1\/(.+)$/i);
            if (ams) {
                const sep = ams[1].includes('?') ? '&' : '?';
                return `https://${targetHost}/v1/${userId}/${ams[1]}${sep}v=1`;
            }
            // Asyncgw URL: insert {userId} immediately after the "/v1/" segment, keeping
            // any prefix path (like "urlp/") intact. Examples:
            //   /v1/objects/...         → /v1/{userId}/objects/...
            //   /urlp/v1/url/image/...  → /urlp/v1/{userId}/url/image/...
            const asyncgw = url.match(/^https:\/\/[a-z0-9-]+\.asyncgw\.teams\.microsoft\.com\/(.*?)v1\/(.*)$/i);
            if (asyncgw) {
                const [, prefix, rest] = asyncgw;
                // Skip if userId is already present immediately after /v1/.
                if (/^[a-f0-9-]{36}\//i.test(rest)) {
                    return `https://${targetHost}/${prefix}v1/${rest}`;
                }
                return `https://${targetHost}/${prefix}v1/${userId}/${rest}`;
            }
            // Any other http(s) URL — route through Teams' generic URL-image proxy.
            // Avoids needing host_permissions for every third-party image CDN, and
            // Teams's server validates the URL + shields the user's IP from the origin.
            if (/^https?:\/\//i.test(url)) {
                return `https://${targetHost}/urlp/v1/${userId}/url/image/Thumbnail?url=${encodeURIComponent(url)}`;
            }
            return null;
        }

        // Set by fetchInlineImages before any concurrent fetches start, so the
        // image fetcher can rewrite AMS/asyncgw URLs to the authenticated proxy form.
        let imgFetchAuth: { userId: string; userRegion: string; ic3Token: string } | null = null;

        // Toggled per-export by fetchInlineImages based on the user's
        // "Image fetch fallback" Setting + the current <all_urls>
        // permission state. When true, fetchImageAsDataUrl will retry
        // a failed proxy fetch directly against the original upstream
        // host via FETCH_BLOB_DIRECT in the background script. When
        // false, proxy failure stays a failure (legacy behaviour).
        let imgFetchFallbackEnabled = false;

        // ── /urlp/ URL-image-proxy fetch via page-context helper ──
        //
        // Cookies set by Teams' login flow on the asyncgw domain
        // (authtoken_asm_urlp, skypetoken_asm) are partitioned to the
        // teams.cloud.microsoft top-level origin in modern Firefox
        // (Total Cookie Protection) and Chrome (3rd-party cookie
        // phaseout). Three contexts can't see them:
        //
        //   1. Background fetch with credentials:'include' — uses the
        //      extension's partition key, gets nothing.
        //   2. Content-script direct fetch — Firefox content scripts
        //      use the extension's network privileges, also a separate
        //      partition. Verified empirically: returns 401 the same
        //      as the background path.
        //   3. The page itself — works. The page's MAIN world IS the
        //      partition that owns those cookies.
        //
        // So we inject a tiny helper script (loaded from
        // public/page-helpers/urlp-fetcher.js, declared as web-
        // accessible in wxt.config.ts) into the page and RPC into it
        // via window.postMessage. The helper does
        // fetch(...,{credentials:'include'}) from the page world,
        // where cookies attach normally, and ships the ArrayBuffer
        // back to us via postMessage with transfer.

        const URLP_REQ = 'tce-urlp-fetch';
        const URLP_RES = 'tce-urlp-result';
        const URLP_READY = 'tce-urlp-helper-ready';
        const URLP_CANCEL = 'tce-urlp-cancel';

        // Lazy, idempotent: only inject the helper on first urlp use,
        // then every subsequent fetch reuses it. The helper itself
        // also guards via a window-scoped flag in case anyone else
        // re-injects it.
        // Outcome of one helper load attempt. Image-fetch call sites
        // historically did not care which case occurred (they retry on
        // a per-call timeout anyway), but the diagnostics probe needs
        // to know real success from a silent timeout — otherwise it
        // reports PASS for a dead helper. Plain Promise<void> hid the
        // distinction.
        type UrlpHelperLoadStatus = 'ready' | 'script-error' | 'timeout';
        let urlpHelperPromise: Promise<UrlpHelperLoadStatus> | null = null;
        function ensureUrlpHelperLoaded(): Promise<UrlpHelperLoadStatus> {
            if (urlpHelperPromise) return urlpHelperPromise;
            urlpHelperPromise = new Promise<UrlpHelperLoadStatus>((resolve) => {
                let resolved = false;
                const finish = (status: UrlpHelperLoadStatus) => {
                    if (resolved) return;
                    resolved = true;
                    window.removeEventListener('message', onReady);
                    resolve(status);
                };
                function onReady(e: MessageEvent) {
                    if (e.source !== window) return;
                    const d = e.data as { type?: string } | null;
                    if (d?.type === URLP_READY) finish('ready');
                }
                window.addEventListener('message', onReady);
                const script = document.createElement('script');
                script.src = runtime.getURL('page-helpers/urlp-fetcher.js');
                script.onload = () => { script.remove(); };
                script.onerror = () => {
                    console.warn('[Teams Exporter] urlp helper script failed to load — urlp fetches will fail.');
                    finish('script-error');
                };
                (document.head || document.documentElement).appendChild(script);
                // Safety timeout: if the helper never signals ready
                // (CSP block, navigation, etc.), settle the promise so
                // call sites can fall back to per-fetch timeouts. The
                // 'timeout' status keeps the cause visible to callers
                // (the diagnostics probe distinguishes it from a real
                // ready event).
                setTimeout(() => finish('timeout'), 5000);
            });
            return urlpHelperPromise;
        }

        let urlpDirectCallCount = 0;
        // Tracked separately from urlpDirectCallCount so we can log the
        // first response of each path independently. AMS-direct uses
        // background bearer-auth (FETCH_BLOB → /v1/{userId}/objects/...);
        // urlp uses page-helper cookie-auth. If the issue #22 user's
        // pasted screenshots fail, knowing the AMS-direct first
        // response status tells us whether it's a tenant-specific
        // bearer-auth issue.
        let amsDirectCallCount = 0;
        // Per-export auth-mode tracking for the AMS-direct path
        // (/v1/{userId}/objects/...). Some tenants reject the IC3
        // Bearer with HTTP 401 even though the same auth works for
        // most users (see issue #22). When we see the first 401, flip
        // to 'failed' and route every subsequent AMS-direct fetch
        // in this export through the page-world cookie helper — same
        // path that already works for /urlp/ thumbnails on those
        // tenants. Reset on every fetchInlineImages call so each
        // export starts fresh; a tenant-state change between exports
        // (token refresh, re-login) is correctly re-detected.
        type AmsBearerStatus = 'unknown' | 'works' | 'failed';
        let amsBearerStatus: AmsBearerStatus = 'unknown';
        // Two-strikes counter for ambiguous network errors. A clear HTTP
        // 401 still flips the status to 'failed' immediately because it is
        // a definitive auth rejection from the server. Network errors
        // (TypeError: Failed to fetch / ERR_NAME_NOT_RESOLVED) can be
        // transient blips, so we require two in a row before flipping the
        // global flag and slowing the rest of the export to the cookie
        // path.
        let amsBearerFailureCount = 0;
        const AMS_BEARER_FAILURE_THRESHOLD = 2;
        // Set to true the first time any image fetch sees HTTP 429.
        // Stays true for the rest of the export so the batched loop
        // can permanently drop to throttled concurrency.
        let sawRateLimit = false;
        async function fetchUrlpDirect(url: string): Promise<{
            ok: boolean; dataUrl?: string; status?: number; statusText?: string; error?: string; sizeReason?: number;
        }> {
            urlpDirectCallCount++;
            const isFirstCall = urlpDirectCallCount === 1;
            await ensureUrlpHelperLoaded();
            if (isFirstCall) {
                console.log(`[Teams Exporter DEBUG] fetchUrlpDirect first call (page-helper RPC) — url: ${url.slice(0, 200)}`);
            }
            return new Promise((resolve) => {
                const id = `u-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 10)}`;
                const timer = setTimeout(() => {
                    window.removeEventListener('message', listener);
                    if (isFirstCall) console.warn('[Teams Exporter DEBUG] fetchUrlpDirect first call timed out at 30s');
                    resolve({ ok: false, error: 'urlp-helper-timeout-30s' });
                }, 30_000);
                function listener(e: MessageEvent) {
                    if (e.source !== window) return;
                    const d = e.data as {
                        type?: string; id?: string; ok?: boolean;
                        status?: number; statusText?: string; error?: string;
                        mime?: string; bytes?: ArrayBuffer;
                    } | null;
                    if (d?.type !== URLP_RES || d.id !== id) return;
                    window.removeEventListener('message', listener);
                    clearTimeout(timer);
                    if (isFirstCall) {
                        const dump = d.ok && d.bytes
                            ? `ok=true, bytes=${d.bytes.byteLength}, mime=${d.mime}`
                            : `ok=false, status=${d.status ?? '-'} ${d.statusText ?? ''}, error=${d.error ?? ''}`;
                        console.log(`[Teams Exporter DEBUG] fetchUrlpDirect first response — ${dump}`);
                    }
                    // The 30s timeout above is already cleared, so anything that
                    // throws here (e.g. a cross-compartment access on the
                    // page-world bytes) would leave this Promise unresolved and
                    // hang the whole image batch. Always resolve.
                    try {
                        if (d.ok && d.bytes) {
                            const buf = d.bytes;
                            if (buf.byteLength > imgFetchMaxBytes || buf.byteLength < 100) {
                                resolve({ ok: false, sizeReason: buf.byteLength });
                                return;
                            }
                            const bytes = new Uint8Array(buf);
                            resolve({ ok: true, dataUrl: `data:${d.mime};base64,${uint8ToBase64(bytes)}` });
                        } else {
                            resolve({ ok: false, status: d.status, statusText: d.statusText, error: d.error });
                        }
                    } catch (err) {
                        resolve({ ok: false, error: `urlp-decode failed: ${String((err as Error)?.message || err).slice(0, 80)}` });
                    }
                }
                window.addEventListener('message', listener);
                window.postMessage({ type: URLP_REQ, id, url }, '*');
            });
        }

        async function fetchImageAsDataUrl(url: string, outcome?: { reason?: string }): Promise<string | null> {
            // Image fetch routing, in order of attempts:
            //   1. Primary: Teams' authenticated proxy (asyncgw) via Bearer token.
            //   2. If the tenant rejects Bearer or the proxy host is unreachable,
            //      retry through the page-world cookie helper.
            //   3. For raw AMS object URLs, also retry the original AMS host with
            //      page-world cookies. Covers environments where asyncgw fails
            //      DNS but the raw AMS host still resolves.
            //   4. For public preview images, when the user has opted into the
            //      "Image fetch fallback" setting and granted <all_urls>, retry
            //      the upstream URL directly via FETCH_BLOB_DIRECT (background
            //      fetch with no credentials).
            const host = hostOf(url);
            const hostStats = getHostStats(host);
            // Record why this fetch failed (one short word) for the caller, which
            // stamps it onto the attachment for the placeholder + manifest. Never
            // overwrites a reason already set, and never fires on success/abort.
            const fail = (r?: string): null => { if (outcome && r && !outcome.reason) outcome.reason = r; return null; };
            const reasonOf = (rr?: { status?: number; error?: string; sizeReason?: number }): string | undefined =>
                rr?.status ? statusToReason(rr.status)
                    : rr?.error ? 'network'
                    : (typeof rr?.sizeReason === 'number' ? (rr.sizeReason > imgFetchMaxBytes ? 'too-large' : 'empty') : undefined);
            if (!imgFetchAuth?.userId || !imgFetchAuth?.userRegion || !imgFetchAuth?.ic3Token) {
                imgFetchStats.skippedDomain++;
                return fail('sign-in');
            }
            const fetchUrl = transformImageUrlToProxy(url, imgFetchAuth.userId, imgFetchAuth.userRegion);
            if (!fetchUrl) {
                imgFetchStats.skippedDomain++;
                return fail('unsupported');
            }
            // URL-image proxy paths (/urlp/.../url/image/Thumbnail) auth
            // via cookies set by Teams' login flow on the asyncgw domain
            // (authtoken_asm_urlp, skypetoken_asm). The IC3 Bearer we
            // use for /v1/{userId}/objects/... gets a 401 from this
            // endpoint — confirmed in a HAR of Teams' own UI requests,
            // which send no Authorization header at all.
            //
            // We have to fetch from the content script (not background)
            // because cookies on the asyncgw domain are partitioned to
            // the teams.cloud.microsoft top-level origin. A background
            // fetch with credentials:'include' uses a different
            // partition key and gets no cookies; a content-script fetch
            // shares the page's partition key and gets them. See
            // fetchUrlpDirect above.
            //
            // We've only verified this in one tenant (eu-prod). Other
            // tenants may differ; if 401s persist, the per-host
            // breakdown + verbose DEBUG logging will say which.
            // Routing decision:
            //   - /urlp/ paths ALWAYS go through the page-world cookie
            //     helper. Their Bearer-auth path 401s on every tenant.
            //   - /v1/{userId}/objects/... (AMS-direct) starts on the
            //     Bearer path. If we've already detected this tenant
            //     rejects Bearer (amsBearerStatus === 'failed'),
            //     skip the doomed Bearer call and route through the
            //     same cookie helper as urlp.
            const isUrlpPath = /\/urlp\//i.test(fetchUrl);
            const isRawAmsObject = /\.asm\.skype\.com\/v1\/objects\/[^/]+\/views\//i.test(url);
            const useCookies = isUrlpPath || amsBearerStatus === 'failed';
            const isFirstAmsDirectCall = !isUrlpPath && amsDirectCallCount === 0;
            if (isFirstAmsDirectCall) {
                amsDirectCallCount++;
                const mode = useCookies ? 'page-world cookies (tenant Bearer 401d earlier)' : 'background bearer-auth';
                console.log(`[Teams Exporter DEBUG] AMS-direct first call (${mode}) — url: ${fetchUrl.slice(0, 200)}`);
            }
            try {
                let resp: { ok: boolean; dataUrl?: string; status?: number; statusText?: string; error?: string; sizeReason?: number };
                // Once we have detected the tenant's Bearer auth is dead
                // (amsBearerStatus === 'failed') AND the URL is a raw AMS
                // object, the existing routing would call asyncgw via cookies
                // (which we already know fails for this tenant) and only
                // then fall into the rescue branch which calls the raw AMS
                // URL. Two helper calls per image when one suffices. Skip
                // straight to the raw AMS attempt and bail on its result so
                // the failure path stays at one helper call too (falling
                // through would cost three).
                if (amsBearerStatus === 'failed' && isRawAmsObject) {
                    const direct = await fetchUrlpDirect(url);
                    if (direct?.ok && direct.dataUrl) {
                        hostStats.ok++;
                        imgFetchStats.amsRawRecovered++;
                        return direct.dataUrl;
                    }
                    if (direct?.status) {
                        imgFetchStats.httpError++;
                        hostStats.httpError++;
                        if (hostStats.firstStatus === undefined) {
                            hostStats.firstStatus = direct.status;
                            hostStats.firstStatusText = direct.statusText || '';
                            hostStats.firstUrl = url.slice(0, 200);
                        }
                        imgFetchStats.failedUrls.push({ url, transformed: fetchUrl, status: direct.status });
                        if (!imgFetchStats.firstHttpError) {
                            imgFetchStats.firstHttpError = `HTTP ${direct.status} ${direct.statusText || ''} for ${url.slice(0, 200)}`;
                        }
                    } else if (direct?.error) {
                        imgFetchStats.threwError++;
                        hostStats.threwError++;
                        if (hostStats.firstError === undefined) {
                            hostStats.firstError = direct.error.slice(0, 120);
                            hostStats.firstUrl = url.slice(0, 200);
                        }
                        imgFetchStats.failedUrls.push({ url, transformed: fetchUrl, error: direct.error });
                        if (!imgFetchStats.firstThrow) {
                            imgFetchStats.firstThrow = `${direct.error.slice(0, 100)} for ${url.slice(0, 80)}`;
                        }
                    } else if (typeof direct?.sizeReason === 'number') {
                        if (direct.sizeReason > imgFetchMaxBytes) {
                            imgFetchStats.tooLarge++;
                            hostStats.tooLarge++;
                        } else {
                            imgFetchStats.tooSmall++;
                            hostStats.tooSmall++;
                        }
                    }
                    return fail(reasonOf(direct));
                }
                if (useCookies) {
                    resp = await fetchUrlpDirect(fetchUrl);
                } else {
                    resp = await runtime.sendMessage({
                        type: 'FETCH_BLOB',
                        url: fetchUrl,
                        bearerToken: imgFetchAuth.ic3Token,
                        maxBytes: imgFetchMaxBytes,
                        minBytes: 100,
                    }) as typeof resp;
                    // First AMS-direct response on a Bearer path tells
                    // us whether the tenant accepts Bearer at all. On a
                    // 401, flip the per-export switch and retry THIS
                    // image via the cookie helper. All subsequent
                    // AMS-direct fetches in this export will skip the
                    // doomed Bearer attempt and go straight to cookies.
                    if (!isUrlpPath && amsBearerStatus === 'unknown') {
                        if (resp?.ok) {
                            amsBearerStatus = 'works';
                            amsBearerFailureCount = 0;
                        } else if (resp?.status === 401) {
                            // Definitive auth rejection. Flip immediately.
                            amsBearerStatus = 'failed';
                            // log (not warn): the cookie fallback is a
                            // successful recovery path, not a failure.
                            // Chrome's chrome://extensions Errors panel
                            // surfaces both warn and error; keeping
                            // recovery-fired messages at log level
                            // avoids cosmetic "errors" for what is
                            // actually working as designed.
                            console.log('[Teams Exporter] AMS-direct Bearer auth returned 401 on first attempt; switching to page-world cookie auth for remaining AMS fetches in this export');
                            // Retry this specific image via cookies so
                            // we don't lose it just because it was the
                            // canary call that revealed the issue.
                            resp = await fetchUrlpDirect(fetchUrl);
                        } else if (resp?.error) {
                            // Ambiguous network error. Could be a one-off blip.
                            // Require two in a row before flipping the global
                            // flag, but still try this specific image via the
                            // cookie helper so we do not lose it.
                            amsBearerFailureCount++;
                            if (amsBearerFailureCount >= AMS_BEARER_FAILURE_THRESHOLD) {
                                amsBearerStatus = 'failed';
                                console.log(`[Teams Exporter] AMS-direct Bearer auth failed ${amsBearerFailureCount} time(s) (last: ${resp?.error || 'network error'}); switching to page-world cookie auth for remaining AMS fetches in this export`);
                            }
                            resp = await fetchUrlpDirect(fetchUrl);
                        }
                    }
                    // Concurrent sibling of the canary call: it was dispatched
                    // on the Bearer path while amsBearerStatus was still
                    // 'unknown', but its 401 arrived after another worker had
                    // already flipped the status to 'failed', so the block above
                    // no longer matched it. Retry it via the cookie helper too —
                    // otherwise the first batch of images is silently lost in a
                    // tenant that rejects Bearer (symptom: an image that loads in
                    // Teams but reads "(not included)" in the export).
                    if (!isUrlpPath && amsBearerStatus === 'failed' && resp?.status === 401) {
                        if (currentAbortController?.signal.aborted) return null;
                        resp = await fetchUrlpDirect(fetchUrl);
                    }
                }
                // 429 retry: Teams' urlp proxy rate-limits per session and
                // returns 429 under sustained load (notably a large multi-chat
                // export where every image goes through the cookie path). Retry
                // up to MAX_429_RETRIES times with escalating backoff + jitter so
                // images aren't dropped to a transient or sustained rate limit; a
                // single retry used to lose the tail. Each 429 flags sawRateLimit
                // so the batched loop in fetchInlineImages also drops concurrency
                // for the rest of the export. The backoff sleep is abortable, so
                // a Stop click during the wait bails immediately (not after up to
                // 8s).
                for (let r429 = 0; resp?.status === 429 && r429 < MAX_429_RETRIES; r429++) {
                    sawRateLimit = true;
                    const backoff = Math.min(1500 * 2 ** r429, 8000);
                    const delay = backoff + Math.floor(Math.random() * 600);
                    await sleepOrAbort(delay, currentAbortController?.signal);
                    if (currentAbortController?.signal.aborted) return null;
                    // Retry on the LIVE auth path: if the tenant already flipped
                    // to cookies, don't bounce a just-rescued sibling back onto
                    // the dead Bearer path (useCookies is a const captured before
                    // the flip, so a pre-flip sibling still has it false).
                    if (useCookies || amsBearerStatus === 'failed') {
                        resp = await fetchUrlpDirect(fetchUrl);
                    } else {
                        resp = await runtime.sendMessage({
                            type: 'FETCH_BLOB',
                            url: fetchUrl,
                            bearerToken: imgFetchAuth.ic3Token,
                            maxBytes: imgFetchMaxBytes,
                            minBytes: 100,
                        }) as typeof resp;
                    }
                }
                if (isFirstAmsDirectCall) {
                    const dump = resp?.ok
                        ? `ok=true, dataUrl=${resp.dataUrl ? `${resp.dataUrl.length}-char data: URL` : 'missing'}`
                        : `ok=false, status=${resp?.status ?? '-'} ${resp?.statusText ?? ''}, error=${resp?.error ?? ''}, sizeReason=${resp?.sizeReason ?? '-'}`;
                    console.log(`[Teams Exporter DEBUG] AMS-direct first response — ${dump}`);
                }
                if (resp?.ok && resp.dataUrl) {
                    hostStats.ok++;
                    return resp.dataUrl;
                }
                // Private Teams AMS fallback:
                // If the asyncgw/proxy path failed, but the original URL is a raw
                // *.asm.skype.com object URL, retry the ORIGINAL raw AMS URL through
                // the page-world cookie helper. This covers environments where
                // asyncgw DNS fails (ERR_NAME_NOT_RESOLVED) but the raw AMS host
                // still resolves and serves the image with page cookies.
                if (isRawAmsObject) {
                    let rawAmsResp: typeof resp | undefined;
                    try {
                        rawAmsResp = await fetchUrlpDirect(url);
                    } catch {
                        rawAmsResp = { ok: false, error: 'rescue threw' };
                    }
                    if (rawAmsResp?.ok && rawAmsResp.dataUrl) {
                        hostStats.ok++;
                        imgFetchStats.amsRawRecovered++;
                        return rawAmsResp.dataUrl;
                    }
                    // Record the rescue's own failure rather than falling through
                    // to record `resp` (the prior asyncgw failure). Otherwise the
                    // summary log mislabels rescue failures as asyncgw failures
                    // and the debug failedUrls dump never shows the raw AMS URL.
                    if (rawAmsResp?.status) {
                        imgFetchStats.httpError++;
                        hostStats.httpError++;
                        if (hostStats.firstStatus === undefined) {
                            hostStats.firstStatus = rawAmsResp.status;
                            hostStats.firstStatusText = rawAmsResp.statusText || '';
                            hostStats.firstUrl = url.slice(0, 200);
                        }
                        imgFetchStats.failedUrls.push({ url, transformed: fetchUrl, status: rawAmsResp.status });
                        if (!imgFetchStats.firstHttpError) {
                            imgFetchStats.firstHttpError = `HTTP ${rawAmsResp.status} ${rawAmsResp.statusText || ''} for ${url.slice(0, 200)}`;
                        }
                    } else if (rawAmsResp?.error) {
                        imgFetchStats.threwError++;
                        hostStats.threwError++;
                        if (hostStats.firstError === undefined) {
                            hostStats.firstError = rawAmsResp.error.slice(0, 120);
                            hostStats.firstUrl = url.slice(0, 200);
                        }
                        imgFetchStats.failedUrls.push({ url, transformed: fetchUrl, error: rawAmsResp.error });
                        if (!imgFetchStats.firstThrow) {
                            imgFetchStats.firstThrow = `${rawAmsResp.error.slice(0, 100)} for ${url.slice(0, 80)}`;
                        }
                    } else if (typeof rawAmsResp?.sizeReason === 'number') {
                        if (rawAmsResp.sizeReason > imgFetchMaxBytes) {
                            imgFetchStats.tooLarge++;
                            hostStats.tooLarge++;
                        } else {
                            imgFetchStats.tooSmall++;
                            hostStats.tooSmall++;
                        }
                    }
                    return fail(reasonOf(rawAmsResp));
                }
                // Image fetch fallback (opt-in feature). When Teams'
                // proxy/page-helper cannot return a public preview image,
                // retry the original upstream URL directly. v1.4.7 only
                // tried this for proxy HTTP 4xx *and* only when the raw URL
                // itself had a ?url= wrapper, so ordinary link previews like
                // https://cdn.overleaf.com/... never actually reached the
                // fallback. Some users also see network-layer failures
                // (TypeError: Failed to fetch / ERR_NAME_NOT_RESOLVED), so
                // treat resp.error as fallback-eligible too.
                const getDirectFallbackUrl = (): string | null => {
                    // A target URL is fallback-eligible only if it is a real
                    // public http(s) URL on a host the unauthenticated direct
                    // path can actually serve. Teams-own hosts (raw AMS, the
                    // asyncgw proxy) need auth, so a direct fetch will 401.
                    // Filtering them here also keeps the per-export privacy
                    // log honest: it lists third-party hosts the extension
                    // contacted, not Teams hosts the user already knows about.
                    const isFallbackEligible = (target: string): boolean => {
                        try {
                            const u = new URL(target);
                            if (!/^https?:$/i.test(u.protocol)) return false;
                            if (/\.asm\.skype\.com$/i.test(u.hostname)) return false;
                            if (/\.asyncgw\.teams\.microsoft\.com$/i.test(u.hostname)) return false;
                            return true;
                        } catch {
                            return false;
                        }
                    };

                    const wrapped = fetchUrl.match(/[?&]url=([^&]+)/) || url.match(/[?&]url=([^&]+)/);
                    if (wrapped) {
                        let target: string;
                        try { target = decodeURIComponent(wrapped[1]); } catch { return null; }
                        return target && isFallbackEligible(target) ? target : null;
                    }
                    // Raw public preview/GIF URLs are wrapped by transformImageUrlToProxy().
                    return isFallbackEligible(url) ? url : null;
                };

                const fallbackEligibleStatus = resp?.status !== undefined
                    && (resp.status === 410 || resp.status === 429 || resp.status === 403 || resp.status === 404);
                const fallbackEligibleNetworkError = !resp?.status && !!resp?.error;

                if (imgFetchFallbackEnabled && (fallbackEligibleStatus || fallbackEligibleNetworkError)) {
                    const upstreamUrl = getDirectFallbackUrl();
                    if (upstreamUrl) {
                        try { directFetchHosts.add(new URL(upstreamUrl).hostname); } catch { /* ignore */ }
                        try {
                            const directResp = await runtime.sendMessage({
                                type: 'FETCH_BLOB_DIRECT',
                                url: upstreamUrl,
                                maxBytes: imgFetchMaxBytes,
                                minBytes: 100,
                            }) as typeof resp;

                            if (directResp?.ok && directResp.dataUrl) {
                                hostStats.ok++;
                                hostStats.directRecovered = (hostStats.directRecovered || 0) + 1;
                                imgFetchStats.directRecovered++;
                                return directResp.dataUrl;
                            }

                            hostStats.directFailed = (hostStats.directFailed || 0) + 1;
                            imgFetchStats.directFailed++;
                        } catch { /* fall through to record proxy failure */ }
                    }
                }
                if (resp?.status) {
                    imgFetchStats.httpError++;
                    hostStats.httpError++;
                    if (hostStats.firstStatus === undefined) {
                        hostStats.firstStatus = resp.status;
                        hostStats.firstStatusText = resp.statusText || '';
                        hostStats.firstUrl = url.slice(0, 200);
                    }
                    imgFetchStats.failedUrls.push({ url, transformed: fetchUrl, status: resp.status });
                    if (!imgFetchStats.firstHttpError) {
                        // For the urlp thumbnail proxy, the useful identifier is the
                        // wrapped external URL in the `?url=` param, not our proxy URL.
                        const m = url.match(/[?&]url=([^&]+)/);
                        const label = m ? decodeURIComponent(m[1]).slice(0, 200) : url.slice(0, 200);
                        imgFetchStats.firstHttpError = `HTTP ${resp.status} ${resp.statusText || ''} for ${label}`;
                    }
                } else if (resp?.error) {
                    imgFetchStats.threwError++;
                    hostStats.threwError++;
                    if (hostStats.firstError === undefined) {
                        hostStats.firstError = resp.error.slice(0, 120);
                        hostStats.firstUrl = url.slice(0, 200);
                    }
                    imgFetchStats.failedUrls.push({ url, transformed: fetchUrl, error: resp.error });
                    if (!imgFetchStats.firstThrow) imgFetchStats.firstThrow = `${resp.error.slice(0, 100)} for ${url.slice(0, 80)}`;
                } else if (typeof resp?.sizeReason === 'number') {
                    if (resp.sizeReason > imgFetchMaxBytes) {
                        imgFetchStats.tooLarge++;
                        hostStats.tooLarge++;
                    } else {
                        imgFetchStats.tooSmall++;
                        hostStats.tooSmall++;
                    }
                }
                return fail(reasonOf(resp));
            } catch (e) {
                imgFetchStats.threwError++;
                hostStats.threwError++;
                if (!imgFetchStats.firstThrow) imgFetchStats.firstThrow = `bg-fetch ${String(e).slice(0, 80)}`;
                if (hostStats.firstError === undefined) hostStats.firstError = `bg-fetch ${String(e).slice(0, 100)}`;
                return fail('network');
            }
        }

        // Per-export image fetch counters; reset at the start of each fetchInlineImages call.
        type HostStats = {
            ok: number;
            httpError: number;
            threwError: number;
            tooLarge: number;
            tooSmall: number;
            // Image fetch fallback (opt-in feature): when the proxy
            // fetch fails with a 4xx that suggests permanent failure,
            // we retry directly against the upstream host. Tracked
            // separately so the diagnostic breakdown shows how often
            // the fallback rescued an image vs. how often it failed
            // alongside the proxy.
            directRecovered?: number;
            directFailed?: number;
            firstStatus?: number;
            firstStatusText?: string;
            firstError?: string;
            firstUrl?: string;
        };
        const imgFetchStats = {
            skippedDomain: 0, httpError: 0, threwError: 0, tooLarge: 0, tooSmall: 0,
            // Image fetch fallback counters (opt-in feature). Surface
            // in the post-export log so the user can see how many
            // images the fallback recovered vs. how many it couldn't.
            directRecovered: 0,
            directFailed: 0,
            // Raw AMS rescue counter. Incremented when the rescue branch
            // (or the Fix 1 short-circuit) saves a private image that the
            // asyncgw proxy could not deliver.
            amsRawRecovered: 0,
            firstHttpError: '' as string,
            firstThrow: '' as string,
            // Per-host breakdown. Key is the upstream host (the URL the user actually
            // sees on the chat: AMS / asyncgw / Google Docs preview / etc.) so the
            // diagnostic shows which class of image is failing rather than just
            // "the proxy" for everything routed through it.
            byHost: new Map<string, HostStats>(),
            // Failed URLs collected when verbose debug is enabled. Written into a
            // _image-fetch-debug.txt sidecar for the export so the user can share
            // the raw evidence with us instead of devtools logs.
            failedUrls: [] as Array<{ url: string; transformed: string; status?: number; error?: string }>,
        };
        // Hosts that received a direct (non-Teams-proxied) fetch during this
        // export. Populated only when the opt-in image-fetch fallback fires.
        // Surfaced in the post-export log so the user can see which third-
        // party hosts their browser contacted on their behalf.
        const directFetchHosts: Set<string> = new Set();
        const getHostStats = (host: string): HostStats => {
            let s = imgFetchStats.byHost.get(host);
            if (!s) {
                s = { ok: 0, httpError: 0, threwError: 0, tooLarge: 0, tooSmall: 0 };
                imgFetchStats.byHost.set(host, s);
            }
            return s;
        };
        const hostOf = (url: string): string => {
            // Prefer the wrapped upstream host for /urlp/.../Thumbnail?url=<X> URLs
            // (we want to see "lh7-us.googleusercontent.com" failing, not just
            // "asyncgw.teams.microsoft.com" because everything goes through it).
            const wrapped = url.match(/[?&]url=([^&]+)/);
            const target = wrapped ? decodeURIComponent(wrapped[1]) : url;
            try { return new URL(target).hostname; } catch { return 'unknown'; }
        };
        const resetImgFetchStats = () => {
            imgFetchStats.skippedDomain = 0;
            imgFetchStats.httpError = 0;
            imgFetchStats.threwError = 0;
            imgFetchStats.tooLarge = 0;
            imgFetchStats.tooSmall = 0;
            imgFetchStats.directRecovered = 0;
            imgFetchStats.directFailed = 0;
            imgFetchStats.amsRawRecovered = 0;
            imgFetchStats.firstHttpError = '';
            imgFetchStats.firstThrow = '';
            imgFetchStats.byHost.clear();
            imgFetchStats.failedUrls.length = 0;
            directFetchHosts.clear();
            urlpDirectCallCount = 0;
            amsDirectCallCount = 0;
            amsBearerStatus = 'unknown';
            amsBearerFailureCount = 0;
            sawRateLimit = false;
        };

        /**
         * Fetch inline images for API-mode messages.
         * Finds AMS image URLs in attachments and downloads them as data URLs.
         */
        // Direct GET against *.asm.skype.com for Teams Free inline images.
        // Auth = `authentication: skypetoken=<JWT>` header (same scheme the
        // chat-service proxy uses) + cookies established by ensureSkypeTokenCookies.
        // Returns a base64 data: URL ready to embed, or null on any failure.
        async function fetchTeamsFreeAmsImage(url: string, jwt: string, outcome?: { reason?: string }): Promise<string | null> {
            try {
                const resp = await fetch(url, {
                    headers: { 'authentication': `skypetoken=${jwt}` },
                    credentials: 'include',
                    signal: currentAbortController?.signal,
                });
                if (!resp.ok) {
                    imgFetchStats.httpError++;
                    if (!imgFetchStats.firstHttpError) {
                        imgFetchStats.firstHttpError = `HTTP ${resp.status} ${resp.statusText || ''} for ${url.slice(0, 200)}`;
                    }
                    if (outcome && !outcome.reason) outcome.reason = statusToReason(resp.status);
                    return null;
                }
                const buf = await resp.arrayBuffer();
                if (buf.byteLength > imgFetchMaxBytes) { imgFetchStats.tooLarge++; if (outcome && !outcome.reason) outcome.reason = 'too-large'; return null; }
                if (buf.byteLength < 100) { imgFetchStats.tooSmall++; if (outcome && !outcome.reason) outcome.reason = 'empty'; return null; }
                const mime = resp.headers.get('content-type')?.split(';')[0]?.trim() || 'image/jpeg';
                const bytes = new Uint8Array(buf);
                return `data:${mime};base64,${uint8ToBase64(bytes)}`;
            } catch (e: any) {
                imgFetchStats.threwError++;
                if (!imgFetchStats.firstThrow) imgFetchStats.firstThrow = `tfl-ams ${String(e?.message || e).slice(0, 80)}`;
                if (outcome && !outcome.reason) outcome.reason = 'network';
                return null;
            }
        }

        async function fetchInlineImages(
            messages: ExportMessage[],
            onProgress?: (done: number, total: number, rateLimited?: boolean) => void,
            auth?: { userId: string | null; userRegion: string; ic3Token: string },
            fullResImages = false,
        ): Promise<void> {
            resetImgFetchStats();
            // Full-res fetches the higher-cap original view; see FULLRES_MAX_IMAGE_BYTES.
            imgFetchMaxBytes = fullResImages ? FULLRES_MAX_IMAGE_BYTES : MAX_IMAGE_BYTES;
            // Per-export image dedup: the same hosted image can appear in
            // several messages; fetch each URL once and reuse the result. Stores
            // successes only (failures stay retryable). Cleared per export.
            const imageUrlCache = new Map<string, string>();
            // Resolve the image-fetch-fallback feature state for this
            // export by asking the background. Firefox MV2 content
            // scripts can't reliably call permissions.contains
            // themselves (the API is gated to the background), so we
            // delegate. Background ANDs the option flag with the
            // current <all_urls> permission state, so a revoked
            // permission disables the feature even if the option flag
            // still says on. Cached for the export to avoid one
            // message per image. Fail-closed: any error disables.
            imgFetchFallbackEnabled = false;
            try {
                const statusResp = await runtime.sendMessage({ type: 'FALLBACK_STATUS' }) as { enabled?: boolean } | undefined;
                imgFetchFallbackEnabled = !!statusResp?.enabled;
            } catch { /* fall closed */ }
            if (imgFetchFallbackEnabled) {
                console.log('[Teams Exporter] Image fetch fallback enabled (direct upstream retry on proxy/network failure).');
            }
            // Only the IC3 token is strictly required (Teams Free's direct
            // AMS fetch needs only the Skype JWT; Work/School's asyncgw
            // proxy ALSO needs userId+userRegion, but checking those at
            // the proxy call site means Teams Free still works when those
            // fields come back empty from the cached SKYPE-TOKEN entry).
            imgFetchAuth = auth?.ic3Token
                ? { userId: auth.userId || '', userRegion: auth.userRegion || '', ic3Token: auth.ic3Token }
                : null;
            if (!imgFetchAuth) {
                console.warn('[Teams Exporter] Image fetch auth missing — IC3/Skype token unavailable. ' +
                    `userId=${auth?.userId ? 'ok' : 'null'} region=${auth?.userRegion ? 'ok' : 'null'} ic3=${auth?.ic3Token ? 'ok' : 'null'}`);
            } else {
                // Auth-state log so the user's submitted log makes it obvious
                // whether the fields are set, without leaking the actual token.
                // Token length is a quick proxy for "looks like a real JWT" vs
                // "captured an empty string" (real Skype JWTs are ~1.5–3 kB).
                console.log(`[Teams Exporter] Image fetch auth state: userId=${imgFetchAuth.userId ? 'set' : 'MISSING'} (${imgFetchAuth.userId?.length || 0} chars), region=${imgFetchAuth.userRegion || 'MISSING'}, ic3Token=${imgFetchAuth.ic3Token ? 'set' : 'MISSING'} (${imgFetchAuth.ic3Token?.length || 0} chars)`);
                if (DEBUG_IMAGE_FETCH) {
                    // Decode the JWT payload (claims only — we never log the signature).
                    // Useful diagnostic when the proxy 401s: the audience claim tells
                    // us whether the token is scoped to the proxy at all.
                    try {
                        const parts = imgFetchAuth.ic3Token.split('.');
                        if (parts.length >= 2) {
                            const json = atob(parts[1].replace(/-/g, '+').replace(/_/g, '/'));
                            const claims = JSON.parse(json);
                            const now = Math.floor(Date.now() / 1000);
                            const exp = typeof claims?.exp === 'number' ? claims.exp : null;
                            const expIn = exp ? `${exp - now}s` : 'unknown';
                            console.log(`[Teams Exporter DEBUG] IC3 JWT claims: aud=${claims?.aud || '?'} iss=${claims?.iss || '?'} exp=${exp || '?'} (in ${expIn}) tid=${claims?.tid || '?'} appid=${claims?.appid || '?'}`);
                        }
                    } catch (e) {
                        console.log('[Teams Exporter DEBUG] Could not decode IC3 token claims:', String(e).slice(0, 80));
                    }
                }
            }
            // ── SharePoint image-file attachments ─────────────────────
            // Files attached via paperclip / drag-drop (as opposed to
            // pasted-as-image) get stored on the user's SharePoint and
            // referenced by URL. Teams renders them as a thumbnail in
            // the chat, so the user expects the image to be in the
            // export. We fetch the SharePoint URL directly with the
            // user's session cookies — `*.sharepoint.com` and
            // `*.sharepoint.us` are in host_permissions for this.
            // Work/school SharePoint hosts honour CORS for the Teams origin
            // and accept cookie-authenticated GETs from a content script.
            // Consumer-account paperclip uploads land on
            // my.microsoftpersonalcontent.com instead, but that host
            // returns a 302 to login.live.com which has no CORS headers
            // for non-interactive callers (verified via probe + console
            // capture in 2026-04). Excluded here so we don't waste a
            // 20-second timeout per legacy chat — the rendered HTML still
            // shows the file as a clickable link the user can open
            // manually.
            const SHAREPOINT_HOST_RE = /^https:\/\/[^/]+\.(sharepoint\.(com|us|cn)|sharepoint-mil\.us)\//i;
            const IMAGE_FILE_EXT_RE = /\.(bmp|png|jpe?g|gif|webp|svg|tiff?|heic|heif)$/i;
            const isImageFileAttachment = (att: { href?: string; type?: string | null; label?: string }) => {
                if (!att.href || !SHAREPOINT_HOST_RE.test(att.href)) return false;
                const t = (att.type || '').toLowerCase();
                if (['bmp', 'png', 'jpg', 'jpeg', 'gif', 'webp', 'svg', 'tif', 'tiff', 'heic', 'heif'].includes(t)) return true;
                return !!(att.label && IMAGE_FILE_EXT_RE.test(att.label));
            };
            const sharePointTasks: { att: { href?: string; dataUrl?: string; type?: string | null; failReason?: string }; url: string }[] = [];
            for (const m of messages) {
                if (!m.attachments) continue;
                for (const att of m.attachments) {
                    if (isImageFileAttachment(att)) sharePointTasks.push({ att, url: att.href! });
                }
            }
            if (sharePointTasks.length > 0) {
                console.log(`[Teams Exporter] Fetching ${sharePointTasks.length} SharePoint image-file attachment(s)…`);
                let spOk = 0;
                for (let i = 0; i < sharePointTasks.length; i += MAX_CONCURRENT_FETCHES) {
                    if (currentAbortController?.signal.aborted) break;
                    const batch = sharePointTasks.slice(i, i + MAX_CONCURRENT_FETCHES);
                    await Promise.all(batch.map(async ({ att, url }) => {
                        if (currentAbortController?.signal.aborted) return;
                        const result = await fetchSharePointFile(url, currentAbortController?.signal);
                        if (result.ok) {
                            if (result.bytes.byteLength > MAX_IMAGE_BYTES) {
                                // log (not warn): an oversized attachment
                                // is a soft skip — file too big to embed,
                                // not a code bug. Keeps Chrome's
                                // Errors panel scoped to extension issues.
                                console.log(`[Teams Exporter] SharePoint file too large (${result.bytes.byteLength} bytes), skipping: ${url.slice(0, 120)}`);
                                att.failReason = 'too-large';
                                return;
                            }
                            att.dataUrl = `data:${result.mime};base64,${uint8ToBase64(result.bytes)}`;
                            spOk++;
                        } else {
                            // log (not warn): SharePoint may decline to
                            // serve a file (auth scope, deleted folder,
                            // permission revoked, etc) — that's external
                            // state, not an extension code error.
                            console.log(`[Teams Exporter] SharePoint fetch failed (${result.status} ${result.statusText} — ${result.reason}) for ${url.slice(0, 120)}`);
                            att.failReason = result.status ? statusToReason(result.status) : 'network';
                        }
                    }));
                }
                console.log(`[Teams Exporter] SharePoint fetch: ${spOk}/${sharePointTasks.length} succeeded`);
            }

            // ── Inline AMS images via Teams URL-image proxy ───────────
            // Collect attachments whose href is an actual image (or routable through
            // the Teams URL-image proxy). Skip file attachments like SharePoint docs,
            // which have an http href but the proxy returns 415 for them.
            const tasks: { att: { href?: string; dataUrl?: string; kind?: 'preview'; type?: string | null; failReason?: string; failHref?: string }; url: string; fallbackUrl?: string }[] = [];
            for (const m of messages) {
                if (!m.attachments) continue;
                for (const att of m.attachments) {
                    if (!att.href || !/^https?:\/\//i.test(att.href)) continue;
                    if (att.dataUrl) continue; // already fetched (e.g. via Graph /shares above)
                    const isAmsObject = /\/v1\/objects\/[^/]+\/views\//i.test(att.href);
                    const isUrlp = /\/urlp\//i.test(att.href);
                    const isImageish = att.type === 'gif' || att.type === 'video' || att.kind === 'preview';
                    if (!isAmsObject && !isUrlp && !isImageish) continue;
                    // Full resolution: Teams' message HTML carries the downscaled
                    // display view (imgo, ~1280px cap, JPEG). The true original is
                    // the imgpsh_fullsize view (PNG) on the same object. Swap it in
                    // for the FETCH only (att.href stays the imgo link), and only
                    // for image objects — never the video/audio/thumbnail views.
                    let url = att.href;
                    let fallbackUrl: string | undefined;
                    if (fullResImages && /\/v1\/objects\/[^/]+\/views\/imgo\b/i.test(url)) {
                        url = url.replace(/(\/v1\/objects\/[^/]+\/views\/)imgo\b/i, '$1imgpsh_fullsize');
                        fallbackUrl = att.href; // original imgo; used if the full-res fetch fails or is over-cap
                    }
                    tasks.push({ att, url, fallbackUrl });
                }
            }

            if (!tasks.length) return;
            console.log(`[Teams Exporter] Fetching ${tasks.length} inline images…`);

            // Debug build: dump the first 5 raw + transformed URLs so we
            // can verify the URL transform produces what we expect for each
            // class of attachment (asyncgw, AMS, urlp wrap, etc.).
            if (DEBUG_IMAGE_FETCH && imgFetchAuth?.userId && imgFetchAuth?.userRegion) {
                const sample = tasks.slice(0, 5);
                console.log('[Teams Exporter DEBUG] Sample image URLs (first 5):');
                for (let i = 0; i < sample.length; i++) {
                    const t = sample[i];
                    const xform = transformImageUrlToProxy(t.url, imgFetchAuth.userId, imgFetchAuth.userRegion);
                    console.log(`  ${i + 1}. raw     : ${t.url.slice(0, 240)}`);
                    console.log(`     proxied : ${xform?.slice(0, 240) || '(transform returned null)'}`);
                }
            }

            // Teams Free has no authenticated image proxy host equivalent
            // to Work/School's *.asyncgw.teams.microsoft.com — direct fetches
            // against *.asm.skype.com using the Skype JWT + session cookies
            // are the only way to get the image bytes. The skypetokenauth
            // round-trip primes the *.asm.skype.com cookie jar; one call
            // covers every subsequent AMS GET in this export.
            const isTeamsFree = location.hostname.toLowerCase().includes('teams.live.com');
            if (isTeamsFree && imgFetchAuth?.ic3Token) {
                await ensureSkypeTokenCookies(imgFetchAuth.ic3Token, currentAbortController?.signal);
            }

            let done = 0;
            let succeeded = 0;
            // Forward export-level abort to the page-world helper so
            // pending fetch()s are torn down immediately when the user
            // clicks Stop. Without this, each in-flight helper call has
            // to wait out the content-side 30s timeout. One-shot listener
            // since we only need to fire the cancel once per export.
            currentAbortController?.signal.addEventListener('abort', () => {
                try { window.postMessage({ type: URLP_CANCEL }, '*'); } catch { /* ignore */ }
            }, { once: true });
            // Process in batches to limit concurrency. Concurrency starts at
            // MAX_CONCURRENT_FETCHES and drops to MAX_CONCURRENT_FETCHES_THROTTLED
            // for the rest of the export the first time any image fetch sees
            // a 429 from the urlp proxy. The transition is logged once.
            // Continuous fetch pool. Up to MAX_CONCURRENT_FETCHES requests stay
            // in flight, and a worker pulls the next task the moment its current
            // one finishes, instead of fixed batches that idle every finished
            // slot until the batch's slowest image returns (one 4 s image used
            // to stall five slots). Peak concurrency and the 429 throttle are
            // unchanged: workers never exceed MAX_CONCURRENT_FETCHES, and a
            // worker whose slot index is at or above the (reduced) target exits
            // the first time any fetch sees a 429, dropping to the throttled
            // count for the rest of the export. nextTask/succeeded/done are only
            // touched between awaits, so the single-threaded event loop makes the
            // shared-counter updates race-free.
            // Full-res images are up to 20MB each, fully buffered before the
            // post-download cap can reject them, so cut the pool to bound peak
            // heap. On a 429 it now drops to MAX_CONCURRENT_FETCHES_THROTTLED (2)
            // like the standard pool, so the rate limit can clear and the retries
            // recover the tail (this only lowers peak heap further).
            let target = fullResImages ? FULLRES_CONCURRENT_FETCHES : MAX_CONCURRENT_FETCHES;
            let nextTask = 0;
            let fullResFellBack = 0;
            const fetchWorker = async (slot: number) => {
                while (true) {
                    if (currentAbortController?.signal.aborted) return;
                    const throttledTarget = MAX_CONCURRENT_FETCHES_THROTTLED;
                    if (sawRateLimit && target !== throttledTarget) {
                        target = throttledTarget;
                        console.log(`[Teams Exporter] Rate limit detected (HTTP 429); reducing concurrency to ${target} for the rest of this export`);
                    }
                    if (slot >= target) return;
                    const idx = nextTask++;
                    if (idx >= tasks.length) return;
                    const { att, url, fallbackUrl } = tasks[idx];
                    // One dispatcher owns the routing (Teams Free vs proxy) so the
                    // full-res url and its imgo fallback take the identical path,
                    // cap (imgFetchMaxBytes), and dedup cache.
                    const fetchOne = (u: string, out?: { reason?: string }) =>
                        (isTeamsFree && /\.asm\.skype\.com\/v1\/objects\//i.test(u) && imgFetchAuth?.ic3Token)
                            ? fetchTeamsFreeAmsImage(u, imgFetchAuth.ic3Token, out)
                            : fetchImageAsDataUrl(u, out);
                    let dataUrl: string | null | undefined = imageUrlCache.get(url);
                    // Why the fetch failed (one short word), from the LAST attempt
                    // (primary, then the imgo fallback if it also failed). A cached
                    // success or a recovered fallback leaves it undefined. Stamped
                    // onto att below for the placeholder + failed-items manifest.
                    let failReason: string | undefined;
                    if (dataUrl === undefined) {
                        const out1: { reason?: string } = {};
                        dataUrl = await fetchOne(url, out1);
                        if (dataUrl) imageUrlCache.set(url, dataUrl);
                        else {
                            failReason = out1.reason;
                            if (fallbackUrl) {
                                // Full-res failed (over 20MB / 404 / error): fall back to
                                // the imgo view so the image is downscaled, never dropped.
                                if (currentAbortController?.signal.aborted) return;
                                dataUrl = imageUrlCache.get(fallbackUrl);
                                if (dataUrl === undefined) {
                                    const out2: { reason?: string } = {};
                                    dataUrl = await fetchOne(fallbackUrl, out2);
                                    if (dataUrl) imageUrlCache.set(fallbackUrl, dataUrl);
                                    else failReason = out2.reason ?? failReason;
                                }
                                if (dataUrl) {
                                    failReason = undefined;
                                    fullResFellBack++;
                                    // The over-cap primary already counted a too-large for
                                    // its host; it recovered via the fallback, so roll that
                                    // back on the aggregate and per-host tallies (over-cap is
                                    // the dominant fallback reason) to keep both honest.
                                    imgFetchStats.tooLarge = Math.max(0, imgFetchStats.tooLarge - 1);
                                    const hs = getHostStats(hostOf(url));
                                    hs.tooLarge = Math.max(0, hs.tooLarge - 1);
                                }
                            }
                        }
                    }
                    if (dataUrl) {
                        att.dataUrl = dataUrl;
                        // Clear any reason a prior path (e.g. the SharePoint fetch
                        // for a .gif) stamped: this attachment did embed, so it is
                        // not a failure for the count / banner / manifest.
                        delete att.failReason;
                        succeeded++;
                    } else {
                        if (failReason) att.failReason = failReason;
                        // Fetch failed. For previews/videos the href is a
                        // thumbnail the HTML renderer draws directly as <img>,
                        // so a failed auth-protected one must be dropped to
                        // avoid a broken icon. Inline images are different: the
                        // renderer shows a quiet "(not included)" placeholder
                        // for auth-protected image hrefs, so KEEP the href —
                        // dropping it would lose the only evidence the image
                        // existed (the card collapses to a bare "image" label)
                        // and strips the AMS URL the TXT/CSV summary needs to
                        // read it as [image] rather than [file: image].
                        // Public URLs (e.g. a giphy gif) are kept either way —
                        // they still load when online.
                        const authProtected = /\/v1\/objects\/[^/]+\/views\//i.test(url)
                            || /\/urlp\//i.test(url)
                            || /(asyncgw\.teams\.microsoft\.com|\.asm\.skype\.com)/i.test(url);
                        const isDirectThumbnail = att.kind === 'preview' || att.type === 'video' || att.type === 'audio';
                        // Clear the render href so the HTML doesn't draw a broken
                        // auth-protected <img>, but keep the URL in failHref so the
                        // failure count / manifest can still dedup by it and show
                        // the host (otherwise N copies of one failed thumbnail count
                        // as N distinct failures).
                        if (authProtected && isDirectThumbnail) { att.failHref = att.href; att.href = undefined; }
                    }
                    done++;
                    onProgress?.(done, tasks.length, sawRateLimit);
                }
            };
            await Promise.all(Array.from({ length: MAX_CONCURRENT_FETCHES }, (_, slot) => fetchWorker(slot)));
            console.log(`[Teams Exporter] Image fetch: ${succeeded} succeeded, ${tasks.length - succeeded} failed (of ${tasks.length} attempted)`);
            if (imgFetchStats.directRecovered > 0 || imgFetchStats.directFailed > 0) {
                console.log(`[Teams Exporter] Image fetch fallback: ${imgFetchStats.directRecovered} recovered, ${imgFetchStats.directFailed} still failed via direct upstream fetch.`);
            }
            if (imgFetchStats.amsRawRecovered > 0) {
                console.log(`[Teams Exporter] Recovered ${imgFetchStats.amsRawRecovered} private image(s) via raw AMS fallback`);
            }
            if (directFetchHosts.size > 0) {
                const hosts = [...directFetchHosts].sort().join(', ');
                console.log(`[Teams Exporter] Image fetch fallback contacted ${directFetchHosts.size} upstream host(s) directly: ${hosts}`);
            }
            // Full-res images that went over the cap but recovered via the imgo
            // fallback had their too-large rolled back (aggregate + per-host) as
            // each fallback succeeded, so the tallies below are already honest.
            if (fullResFellBack > 0) {
                console.log(`[Teams Exporter] ${fullResFellBack} full-res image(s) over the ${Math.round(FULLRES_MAX_IMAGE_BYTES / 1048576)}MB cap or unavailable; used the downscaled view instead`);
            }
            const failures = imgFetchStats.httpError + imgFetchStats.threwError + imgFetchStats.tooLarge + imgFetchStats.tooSmall + imgFetchStats.skippedDomain;
            if (failures > 0) {
                const detail = [
                    imgFetchStats.httpError && `${imgFetchStats.httpError} http-error`,
                    imgFetchStats.threwError && `${imgFetchStats.threwError} threw`,
                    imgFetchStats.tooLarge && `${imgFetchStats.tooLarge} too-large`,
                    imgFetchStats.tooSmall && `${imgFetchStats.tooSmall} too-small`,
                    imgFetchStats.skippedDomain && `${imgFetchStats.skippedDomain} domain-blocked`,
                ].filter(Boolean).join(', ');
                // log (not warn): these are reports of external state
                // (upstream returned 4xx, network blip, etc), not errors
                // in our own code. Chrome's chrome://extensions Errors
                // panel surfaces warn + error and treats anything there
                // as "the extension misbehaved" — surfacing external
                // failures via that channel misleads users into
                // thinking the extension is buggy. Stays accessible in
                // DevTools console for support / debugging.
                console.log(`[Teams Exporter] Image fetch failures — ${detail}`);
                if (imgFetchStats.firstHttpError) console.log(`[Teams Exporter] First http error: ${imgFetchStats.firstHttpError}`);
                if (imgFetchStats.firstThrow) console.log(`[Teams Exporter] First exception: ${imgFetchStats.firstThrow}`);
                // Per-host breakdown. The "first http error" line above shows
                // ONE error from ONE URL, which is misleading when 200+ images
                // fail across several different hosts (asyncgw, AMS, Google
                // Docs preview, etc.). The breakdown shows the failure pattern
                // by upstream host so the actual culprit is obvious.
                if (imgFetchStats.byHost.size > 0) {
                    const rows: Array<[string, HostStats]> = Array.from(imgFetchStats.byHost.entries())
                        .sort((a, b) => (b[1].httpError + b[1].threwError) - (a[1].httpError + a[1].threwError));
                    // log (not warn): the breakdown is informational —
                    // every host's success/failure tally, including all
                    // the fully-OK rows. Chrome's chrome://extensions
                    // Errors panel surfaces both warn and error, so an
                    // export with any failures used to flood it with
                    // dozens of "ok" rows. The summary line above
                    // (`Image fetch failures — N http-error`) stays at
                    // warn so real failures still surface there.
                    console.log('[Teams Exporter] Image fetch breakdown by upstream host:');
                    for (const [host, s] of rows) {
                        const total = s.ok + s.httpError + s.threwError + s.tooLarge + s.tooSmall;
                        const parts = [
                            s.httpError && `${s.httpError} http-error`,
                            s.threwError && `${s.threwError} threw`,
                            s.tooLarge && `${s.tooLarge} too-large`,
                            s.tooSmall && `${s.tooSmall} too-small`,
                        ].filter(Boolean).join(', ');
                        const firstFail = s.firstStatus !== undefined
                            ? `first: HTTP ${s.firstStatus} ${s.firstStatusText || ''}`.trim()
                            : s.firstError ? `first: ${s.firstError}` : '';
                        console.log(`  ${host}: ${s.ok}/${total} ok${parts ? ` (${parts})` : ''}${firstFail ? ` — ${firstFail}` : ''}`);
                    }
                }
                // Debug build: dump up to 30 failed URLs to console so the
                // user can copy/paste the log directly into a GitHub issue.
                // Earlier we considered a sidecar file in the export, but
                // that means threading state through the builder pipeline
                // for a one-shot debug release. Console is plenty.
                if (DEBUG_IMAGE_FETCH && imgFetchStats.failedUrls.length > 0) {
                    const sample = imgFetchStats.failedUrls.slice(0, 30);
                    console.warn(`[Teams Exporter DEBUG] First ${sample.length} failed URLs (raw → proxied → status/error):`);
                    for (let i = 0; i < sample.length; i++) {
                        const f = sample[i];
                        const tail = f.status ? `HTTP ${f.status}` : f.error ? f.error.slice(0, 120) : 'unknown';
                        console.warn(`  ${i + 1}. ${tail}`);
                        console.warn(`     raw     : ${f.url.slice(0, 240)}`);
                        console.warn(`     proxied : ${f.transformed.slice(0, 240)}`);
                    }
                    if (imgFetchStats.failedUrls.length > 30) {
                        console.warn(`  … (${imgFetchStats.failedUrls.length - 30} more failures truncated)`);
                    }
                }
            }
        }

        // ── Avatar Fetching (API mode) ────────────────────────────
        // Cross-chat profile-photo cache. The same person recurs across many
        // chats in a bundle, so each Graph photo is fetched once and reused
        // (measured ~75% of avatar fetches were repeats). Value is the data URL,
        // or '' for a known no-photo (404) so we don't re-query. Persists for
        // the content-script lifetime, i.e. the whole bundle. Transient failures
        // are NOT cached, so a later chat can retry them.
        const apiAvatarPhotoCache = new Map<string, string>();
        /**
         * Fetch profile photos via Graph API for API-mode messages.
         * Extracts unique author UUIDs, fetches photos, and sets avatar data URLs.
         */
        async function fetchApiAvatars(
            messages: ExportMessage[],
            rawMessages: Array<{ from?: string; imdisplayname?: string }>,
        ): Promise<void> {
            const graphToken = await getGraphToken();
            if (!graphToken) return;
            // context, so crossChatCacheHits = avatars a bundle-wide photo cache
            // would have skipped. Remove after perf run.

            // Build author → UUID map from raw API messages
            const authorToUuid = new Map<string, string>();
            for (const m of rawMessages) {
                if (!m.from || !m.imdisplayname) continue;
                const uuidMatch = m.from.match(/([a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12})/i);
                if (uuidMatch) authorToUuid.set(m.imdisplayname, uuidMatch[1]);
            }

            if (!authorToUuid.size) return;

            // Fetch profile photos via background (same path used for images) so both
            // browsers behave the same way; direct content-script fetch to Graph can
            // fail sporadically in Firefox even with host_permissions.
            // Track UNIQUE attempts and explicit failures separately. authorToUuid
            // can have multiple display names mapping to the same UUID, so comparing
            // "authors" to "successes" inflated the failure count by the duplicates.
            const uuidToDataUrl = new Map<string, string>();
            const attempted = new Set<string>();
            let firstPhotoError: string | null = null;
            let photo404 = 0;
            let nonPhotoFailures = 0;
            // Pre-compute the deduplicated total so we can emit "x/total" progress.
            // authorToUuid maps display names to UUIDs and the same UUID can repeat
            // across many display names; a unique count is the meaningful one.
            const totalAvatars = new Set(authorToUuid.values()).size;
            // Resolve which UUIDs still need a Graph fetch (after dedup + the
            // cross-chat cache), then fetch them through a small concurrency pool.
            // These hit Graph (graph.microsoft.com), not the 429-sensitive chat
            // message endpoint, and they were previously fetched strictly one at a
            // time (the serial loop was the second-biggest export phase). Counters
            // and caches are only touched between awaits, so the single-threaded
            // event loop keeps the shared updates race-free.
            const avatarsToFetch: string[] = [];
            let avProcessed = 0;
            for (const [, uuid] of authorToUuid) {
                if (attempted.has(uuid)) continue;
                attempted.add(uuid);
                // Cross-chat cache: reuse a photo already fetched for an earlier
                // chat in this bundle (or a known no-photo) instead of re-querying.
                const cachedPhoto = apiAvatarPhotoCache.get(uuid);
                if (cachedPhoto !== undefined) {
                    if (cachedPhoto) uuidToDataUrl.set(uuid, cachedPhoto);
                    avProcessed++;
                    continue;
                }
                avatarsToFetch.push(uuid);
            }
            // Resolve photos in Graph $batch first (20 per POST). Whatever it
            // returns is seeded here; only the UUIDs it could NOT process fall
            // through to the per-UUID FETCH_BLOB pool below. Self-probing: if
            // /$batch is unusable for photos, resolveAvatarPhotos returns an
            // empty map and everything falls through unchanged.
            if (avatarsToFetch.length && graphToken) {
                try {
                    const batched = await resolveAvatarPhotos(avatarsToFetch, graphToken, MAX_IMAGE_BYTES, currentAbortController?.signal);
                    if (batched.size) {
                        const remaining: string[] = [];
                        for (const uuid of avatarsToFetch) {
                            if (!batched.has(uuid)) { remaining.push(uuid); continue; }
                            const dataUrl = batched.get(uuid);
                            if (dataUrl) { uuidToDataUrl.set(uuid, dataUrl); apiAvatarPhotoCache.set(uuid, dataUrl); }
                            else { photo404++; apiAvatarPhotoCache.set(uuid, ''); } // no photo
                            avProcessed++;
                        }
                        avatarsToFetch.length = 0;
                        avatarsToFetch.push(...remaining);
                    }
                } catch (e) { console.log('[Teams Exporter] avatar $batch failed, using per-photo path:', e); }
            }
            // Emit one avatars-progress update now so the popup learns the total
            // even when the batch resolved everyone (then the worker pool below
            // does nothing and would never emit it). Common on the first chat.
            if (totalAvatars > 0) {
                try {
                    runtime.sendMessage({ type: 'SCRAPE_PROGRESS', payload: { phase: 'avatars', avatarsDone: avProcessed, avatarsTotal: totalAvatars } });
                } catch { /* ignore */ }
            }
            let nextAvatar = 0;
            const avatarWorker = async () => {
                while (true) {
                    if (currentAbortController?.signal.aborted) return;
                    const idx = nextAvatar++;
                    if (idx >= avatarsToFetch.length) return;
                    const uuid = avatarsToFetch[idx];
                    try {
                        const resp = await runtime.sendMessage({
                            type: 'FETCH_BLOB',
                            url: `https://graph.microsoft.com/v1.0/users/${uuid}/photo/$value`,
                            bearerToken: graphToken,
                            maxBytes: MAX_IMAGE_BYTES,
                            minBytes: 1,
                        }) as { ok: boolean; dataUrl?: string; status?: number; statusText?: string; error?: string; sizeReason?: number };
                        if (resp?.ok && resp.dataUrl) {
                            uuidToDataUrl.set(uuid, resp.dataUrl);
                            apiAvatarPhotoCache.set(uuid, resp.dataUrl);
                        } else if (resp?.status === 404) {
                            photo404++;
                            apiAvatarPhotoCache.set(uuid, '');  // known no-photo; don't re-query
                        } else {
                            nonPhotoFailures++;
                            if (!firstPhotoError) {
                                firstPhotoError = resp?.status
                                    ? `HTTP ${resp.status} ${resp.statusText || ''}`
                                    : (resp?.error || 'unknown error');
                            }
                        }
                    } catch (e) {
                        nonPhotoFailures++;
                        if (!firstPhotoError) firstPhotoError = String(e);
                    }
                    avProcessed++;
                    // Emit progress every few avatars (and on the last one) so the
                    // popup button can show "x/total" without spamming messages.
                    if (avProcessed % 5 === 0 || avProcessed === totalAvatars) {
                        try {
                            runtime.sendMessage({
                                type: 'SCRAPE_PROGRESS',
                                payload: { phase: 'avatars', avatarsDone: avProcessed, avatarsTotal: totalAvatars },
                            });
                        } catch { /* ignore */ }
                    }
                }
            };
            await Promise.all(Array.from({ length: MAX_CONCURRENT_FETCHES }, () => avatarWorker()));
            // how many of this chat's people were already fetched in an earlier
            // chat this run (= what a bundle-wide photo cache would save).
            const missing = attempted.size - uuidToDataUrl.size;
            if (missing > 0) {
                const parts: string[] = [];
                if (photo404 > 0) parts.push(`${photo404} have no photo (404)`);
                if (nonPhotoFailures > 0) parts.push(`${nonPhotoFailures} failed (${firstPhotoError || 'no error captured'})`);
                const level = nonPhotoFailures > 0 ? 'warn' : 'log';
                console[level](`[Teams Exporter] Graph photo: ${uuidToDataUrl.size}/${attempted.size} unique users fetched — ${parts.join(', ')}`);
            }

            if (!uuidToDataUrl.size) return;
            console.log(`[Teams Exporter] Fetched ${uuidToDataUrl.size} profile photos`);

            // Set avatar on converted messages
            for (const m of messages) {
                if (!m.author || m.avatar) continue;
                const uuid = authorToUuid.get(m.author);
                if (uuid) {
                    const dataUrl = uuidToDataUrl.get(uuid);
                    if (dataUrl) m.avatar = dataUrl;
                }
            }
        }

        // Graph-fetch profile photos for reactor UUIDs the self/name join could
        // not resolve (reactors who never sent a message in the chat). Reuses the
        // background FETCH_BLOB path used for author photos. Rate-limit safe: the
        // input is already deduped by UUID, capped, and backs off after repeated
        // 429s (see the documented API rate-limit caution).
        // Session cache so a reactor who appears across many chats in a bundle is
        // fetched at most once per Teams-tab session (null = tried, no photo).
        const reactorPhotoCache = new Map<string, string | null>();
        async function fetchReactorPhotos(uuids: string[]): Promise<Map<string, string>> {
            const out = new Map<string, string>();
            for (const u of uuids) { const c = reactorPhotoCache.get(u); if (c) out.set(u, c); }
            const toFetch = uuids.filter(u => !reactorPhotoCache.has(u));
            if (!toFetch.length) return out;
            const graphToken = await getGraphToken();
            if (!graphToken) return out;
            const CAP = 400;
            const targets = toFetch.slice(0, CAP);
            // Graph $batch first; only UUIDs it could not process fall through to
            // the per-UUID serial loop below (self-probing, returns empty if
            // /$batch is unusable for photos here).
            let serialTargets = targets;
            try {
                const batched = await resolveAvatarPhotos(targets, graphToken, MAX_IMAGE_BYTES, currentAbortController?.signal);
                if (batched.size) {
                    serialTargets = [];
                    for (const uuid of targets) {
                        if (!batched.has(uuid)) { serialTargets.push(uuid); continue; }
                        const dataUrl = batched.get(uuid);
                        if (dataUrl) { out.set(uuid, dataUrl); reactorPhotoCache.set(uuid, dataUrl); }
                        else reactorPhotoCache.set(uuid, null); // no photo / oversize
                    }
                }
            } catch { /* fall through to the serial FETCH_BLOB path */ }
            let consecutive429 = 0;
            for (const uuid of serialTargets) {
                if (currentAbortController?.signal.aborted) break;
                if (consecutive429 >= 3) {
                    console.warn('[Teams Exporter] Reactor photo fetch: backing off after repeated 429s');
                    break;
                }
                try {
                    const resp = await runtime.sendMessage({
                        type: 'FETCH_BLOB',
                        url: `https://graph.microsoft.com/v1.0/users/${uuid}/photo/$value`,
                        bearerToken: graphToken,
                        maxBytes: MAX_IMAGE_BYTES,
                        minBytes: 1,
                    }) as { ok: boolean; dataUrl?: string; status?: number };
                    if (resp?.ok && resp.dataUrl) { out.set(uuid, resp.dataUrl); reactorPhotoCache.set(uuid, resp.dataUrl); consecutive429 = 0; }
                    else if (resp?.status === 429) consecutive429++; // transient — don't cache
                    else { reactorPhotoCache.set(uuid, null); consecutive429 = 0; } // 404/other — cache the miss
                } catch { /* skip this reactor */ }
            }
            if (toFetch.length > CAP) {
                console.log(`[Teams Exporter] Reactor photos: capped at ${CAP} of ${toFetch.length} new reactors`);
            }
            if (out.size) console.log(`[Teams Exporter] Reactor profile photos: ${out.size} (session cache ${reactorPhotoCache.size})`);
            return out;
        }

        const extractCodeBlock = (el: Element) => {
            let code = '';
            const walkCode = (n: ChildNode) => {
                if (n.nodeType === Node.TEXT_NODE) { code += n.nodeValue; return; }
                if (n.nodeType !== Node.ELEMENT_NODE) return;
                const child = n as Element;
                const tagName = child.tagName;
                if (tagName === 'BR') { code += '\n'; return; }
                if (tagName === 'IMG') { code += cleanAltText(child.getAttribute('alt') || child.getAttribute('aria-label')); return; }
                for (const c of child.childNodes) walkCode(c);
            };
            walkCode(el);
            return code.replace(/\u00a0/g, ' ').replace(/\n+$/, '');
        };

        function extractCodeBlocks(root: Element | null): string[] {
            if (!root) return [];
            const skip = [
                '[data-tid="quoted-reply-card"]',
                '[data-tid="referencePreview"]',
                '[role="group"][aria-label^="Begin Reference"]',
            ];
            const out: string[] = [];
            const seen = new Set<string>();
            const pushBlock = (code: string) => {
                const cleaned = code.replace(/\u00a0/g, ' ').replace(/\n+$/, '');
                if (!cleaned.trim()) return;
                const key = cleaned.trim();
                if (seen.has(key)) return;
                seen.add(key);
                out.push(cleaned);
            };
            root.querySelectorAll('pre').forEach(pre => {
                if (skip.some(sel => pre.closest(sel))) return;
                pushBlock(extractCodeBlock(pre));
            });
            const containers = new Set<Element>();
            root.querySelectorAll<HTMLElement>('.cm-line').forEach(line => {
                const container = line.closest('pre, code') || line.parentElement;
                if (container) containers.add(container);
            });
            for (const container of containers) {
                if (container.tagName === 'PRE') continue;
                if (skip.some(sel => container.closest(sel))) continue;
                pushBlock(extractCodeBlock(container));
            }
            return out;
        }

        function extractRichTextAsMarkdown(root: Element | null): string {
            if (!root) return "";
          
            let out = "";
          
            const walk = (n: ChildNode) => {
              if (n.nodeType === Node.TEXT_NODE) {
                out += n.nodeValue ?? "";
                return;
              }
              if (n.nodeType !== Node.ELEMENT_NODE) return;
          
              const el = n as HTMLElement;
              const tag = el.tagName;
          
              // hard breaks
              if (tag === "BR") { out += "\n"; return; }
          
              // emojis / inline images
              if (tag === "IMG") {
                out += cleanAltText(el.getAttribute("alt") || el.getAttribute("aria-label"));
                return;
              }
          
              // inline code
              if (tag === "CODE") {
                out += "`";
                el.childNodes.forEach(walk);
                out += "`";
                return;
              }
          
              // code blocks
              if (tag === "PRE") {
                const code = extractCodeBlock(el);
                if (code) out += `\n\`\`\`\n${code}\n\`\`\`\n`;
                return;
              }
          
              // links
              if (tag === "A") {
                const href = el.getAttribute("href") || "";
                const before = out.length;
                el.childNodes.forEach(walk);
                const text = out.slice(before);
                out = out.slice(0, before);
                out += href ? `[${text}](${href})` : text;
                return;
              }
          
              // bold/italic/strike
              const wrap = (marker: string) => {
                out += marker;
                el.childNodes.forEach(walk);
                out += marker;
              };
          
              if (tag === "STRONG" || tag === "B") { wrap("**"); return; }
              if (tag === "EM" || tag === "I") { wrap("*"); return; }
              if (tag === "DEL" || tag === "S") { wrap("~~"); return; }
          
              // blockquotes
              if (tag === "BLOCKQUOTE") {
                const before = out.length;
                el.childNodes.forEach(walk);
                const chunk = out.slice(before).trim();
                out = out.slice(0, before);
                if (chunk) {
                  const lines = chunk.split(/\n/);
                  out += lines.map(l => (l ? `> ${l}` : `>`)).join("\n") + "\n";
                }
                return;
              }
          
              // default recursion
              const isBlock = /^(DIV|P|LI|BLOCKQUOTE|H[1-6])$/.test(tag);
              const start = out.length;
          
              el.childNodes.forEach(walk);
          
              // add paragraph-ish spacing
              if (isBlock && out.length > start) out += "\n";
            };
          
            root.childNodes.forEach(walk);
          
            return out.replace(/\n{3,}/g, "\n\n").trim();
          }
          

        // Helpers -------------------------------------------------------
        const sleep = (ms: number) => new Promise(resolve => setTimeout(resolve, ms));


        async function waitForPreviewImages(item: Element, timeoutMs = 350) {
            const imgs = Array.from(
                item.querySelectorAll<HTMLImageElement>(
                    '[data-tid="file-preview-root"][amspreviewurl] img[data-tid="rich-file-preview-image"],' +
                    'span[itemtype="http://schema.skype.com/AMSImage"] img[data-gallery-src],' +
                    'img[itemtype="http://schema.skype.com/AMSImage"][data-gallery-src]',
                ),
            );
            if (!imgs.length) return;

            const waits = imgs.map(img => {
                if (img.complete && img.naturalWidth > 0 && img.naturalHeight > 0) return Promise.resolve();
                if (typeof img.decode === 'function') return img.decode().catch(() => {});
                return new Promise<void>(resolve => {
                    const done = () => resolve();
                    img.addEventListener('load', done, { once: true });
                    img.addEventListener('error', done, { once: true });
                });
            });

            await Promise.race([Promise.all(waits), sleep(timeoutMs)]);
        }

        async function expandMessageContent(wrapper: Element | null) {
            if (!wrapper) return;
            const btn = wrapper.querySelector<HTMLButtonElement>(
                '[data-track-module-name="seeMoreButton"], [aria-controls^="see-more-content-"]',
            );
            if (!btn) return;
            if (btn.getAttribute('aria-expanded') === 'true') return;
            try { btn.click(); } catch { /* noop: best-effort expand; a failed click just exports the collapsed text */ }
            await sleep(160);
        }

        function findMainWrapper(item: Element) {
            const wrappers = Array.from(item.querySelectorAll<HTMLElement>('[data-testid="message-body-flex-wrapper"]'));
            const primary = wrappers.find(wrapper => {
                const mid = wrapper.getAttribute('data-mid');
                const chain = wrapper.getAttribute('data-reply-chain-id');
                if (mid && chain && mid === chain) return true;
                return Boolean(wrapper.querySelector('[data-tid="subject-line"]'));
            });
            return primary || wrappers[0] || null;
        }

        function findResponseSummaryButtonByParentId(parentId: string): HTMLButtonElement | null {
            // Find the post renderer by id if possible
            const post = qAny(`#post-message-renderer-${cssEscape(parentId)}`) ||
                         qAny(`#message-body-${cssEscape(parentId)}`) ||
                         qAny(`[data-mid="${cssEscape(parentId)}"]`)?.closest('[id^="post-message-renderer-"], [id^="message-body-"]');
          
            if (!post) return null;
          
            const surface = post.parentElement?.querySelector<HTMLElement>('[data-tid="response-surface"]') ||
                            post.querySelector<HTMLElement>('[data-tid="response-surface"]');
          
            return surface?.querySelector<HTMLButtonElement>('button[data-tid="response-summary-button"]') || null;
          }
          
          async function scrollPostIntoView(parentId: string) {
            const post =
              qAny(`#post-message-renderer-${cssEscape(parentId)}`) ||
              qAny(`#message-body-${cssEscape(parentId)}`) ||
              qAny(`[data-mid="${cssEscape(parentId)}"]`)?.closest('[id^="post-message-renderer-"], [id^="message-body-"]');
          
            if (!post) return false;
          
            // scroll the *correct scroller* (channel)
            const scroller = getScroller('team') as HTMLElement | null;
            if (!scroller) return false;
          
            post.scrollIntoView({ block: "center" });
            await sleep(120);
            return true;
          }
          

        function findReplyWrapper(item: Element) {
            const mid = $('[data-tid="reply-message-body"]', item)?.getAttribute('data-mid') || $('[data-tid="channel-pane-message"]', item)?.getAttribute('data-mid');
            if (!mid) return null;
            return item.querySelector<HTMLElement>(`[data-testid="message-body-flex-wrapper"][data-mid="${cssEscape(mid)}"]`) ||
                item.querySelector<HTMLElement>(`[data-testid="message-body-flex-wrapper"][data-reply-chain-id="${cssEscape(mid)}"]`);
        }

        async function expandSeeMore(item: Element) {
            const mainWrapper = findMainWrapper(item);
            await expandMessageContent(mainWrapper);
            const replyWrapper = findReplyWrapper(item);
            if (replyWrapper && replyWrapper !== mainWrapper) {
                await expandMessageContent(replyWrapper);
            }
        }

        async function waitForSelector(selector: string, timeoutMs = 2000) {
            const start = Date.now();
            while (Date.now() - start < timeoutMs) {
              const el = qAny(selector);
              if (el) return el;
              await sleep(100);
            }
            return null;
          }
          

        function deriveParentIdFromItem(item: Element): string | null {
            const itemRoot =
              item.closest<HTMLElement>('[data-tid="channel-pane-message"]') ||
              item.closest<HTMLElement>('li[role="none"]') ||
              (item as HTMLElement);
          
            // Prefer the main post wrapper mid
            const mid =
              itemRoot.querySelector<HTMLElement>('[data-testid="message-body-flex-wrapper"][data-mid]')?.getAttribute('data-mid') ||
              itemRoot.querySelector<HTMLElement>('[data-mid]')?.getAttribute('data-mid') ||
              itemRoot.getAttribute('data-mid');
          
            if (mid) return mid;
          
            // If no mid, sometimes the response-summary button id includes it: response-summary-<mid>
            const surface =
              itemRoot.parentElement?.querySelector<HTMLElement>('[data-tid="response-surface"]') ||
              itemRoot.querySelector<HTMLElement>('[data-tid="response-surface"]');
          
            const btn = surface?.querySelector<HTMLButtonElement>('[data-tid="response-summary-button"][id^="response-summary-"]');
            if (btn?.id) {
              const m = btn.id.match(/^response-summary-(.+)$/);
              if (m?.[1]) return m[1];
            }
          
            return null;
          }
          
          function getRepliesRunway(): Element | null {
            return (
              qAny('[data-tid="channel-replies-runway"]') ||
              qAny('#channel-pane-l2') ||
              null
            );
          }
          

        function getRepliesItems(): Element[] {
            const runway = getRepliesRunway();
            if (!runway) return [];
            const listItems = Array.from(runway.querySelectorAll('li'));
            const items: Element[] = [];
            for (const li of listItems) {
                const message = li.querySelector<HTMLElement>('[data-tid="channel-replies-pane-message"]');
                if (message) {
                    items.push(message);
                    continue;
                }
                const divider = li.querySelector<HTMLElement>('[data-testid="timestamp-divider"]');
                if (divider) {
                    items.push(divider);
                }
            }
            return items.length ? items : listItems;
        }

        function getReplyItemId(item: Element, index: number): string {
            const mid =
                item.querySelector('[data-testid="message-body-flex-wrapper"][data-mid]')?.getAttribute('data-mid') ||
                item.querySelector('[data-mid]')?.getAttribute('data-mid') ||
                item.getAttribute('data-mid');
            if (mid) return mid;
            return item.id || `reply-${index}`;
        }

        function getRepliesScroller(): Element | null {
            const runway = getRepliesRunway();
            if (!runway) return null;
            const primary = findScrollableAncestor(runway);
            if (primary) return primary;
            const items = getRepliesItems();
            for (const item of items) {
                const candidate = findScrollableAncestor(item);
                if (candidate) return candidate;
            }
            const replyPane =
                document.querySelector<HTMLElement>('[data-tid*="channel-replies"]') ||
                document.querySelector<HTMLElement>('[id^="channel-replies-"]');
            return findScrollableAncestor(replyPane) || document.scrollingElement;
        }

        function isRepliesLoading(): boolean {
            const runway = getRepliesRunway();
            if (!runway) return false;
            const loader = runway.parentElement?.querySelector<HTMLElement>('[data-testid="virtual-list-loader"]') ||
                runway.closest('[data-testid]')?.querySelector<HTMLElement>('[data-testid="virtual-list-loader"]') ||
                document.querySelector<HTMLElement>('[data-testid="virtual-list-loader"]');
            if (loader && loader.offsetParent !== null) {
                const rect = loader.getBoundingClientRect();
                if (rect.height >= 1 || rect.width >= 1) return true;
            }
            return false;
        }

        async function waitForRepliesPaneForParent(parentId: string, timeoutMs = 6000): Promise<boolean> {
            const start = Date.now();
            let sawAnyMessages = false;
            while (Date.now() - start < timeoutMs) {
              const runway = getRepliesRunway();
              if (!runway) { await sleep(120); continue; }
          
              // Strong signal: something in the replies pane references this chain id
              const match =
                runway.querySelector(`[data-reply-chain-id="${cssEscape(parentId)}"]`) ||
                runway.querySelector(`[data-tid="channel-replies-pane-message"] [data-reply-chain-id="${cssEscape(parentId)}"]`);
          
              // Backup signal: the pane actually has messages loaded (not empty / still transitioning)
              const items = getRepliesItems();
              const hasAnyMessages = items.some(el =>
                (el as HTMLElement).getAttribute?.('data-tid') === 'channel-replies-pane-message' ||
                el.querySelector?.('[data-tid="channel-replies-pane-message"]')
              );
              if (hasAnyMessages) sawAnyMessages = true;
          
              if (match || hasAnyMessages) {
                if (match) return true;
              }
          
              await sleep(150);
            }
            return sawAnyMessages;
          }

          
    
    
          async function openRepliesForItem(btn: HTMLButtonElement, parentId: string): Promise<OpenMode> {
            const maxTries = 3;
          
            for (let attempt = 1; attempt <= maxTries; attempt++) {
                await scrollPostIntoView(parentId);

                const liveBtn = findResponseSummaryButtonByParentId(parentId) || btn;
                await realClick(liveBtn);
          
              // Give layout a moment (Teams often needs this)
              await sleep(120);
          
              // 1) Pane path
              const runway = await waitForSelector(
                '[data-tid="channel-replies-runway"], #channel-pane-l2',
                3000
              );
          
              if (runway) {
                const ok = await waitForRepliesPaneForParent(parentId, 6500);
                if (ok) return "pane";
          
                await closeRepliesPane();
                await sleep(250);
                continue;
              }
          
          
              // Optional: quick second click inside the attempt
              await sleep(200);
              await realClick(btn);
              await sleep(200);
          
              const runway2 = await waitForSelector(
                '[data-tid="channel-replies-runway"], #channel-pane-l2',
                1500
              );
              if (runway2) {
                const ok2 = await waitForRepliesPaneForParent(parentId, 6500);
                if (ok2) return "pane";
                await closeRepliesPane();
                await sleep(250);
                continue;
              }
          
              await sleep(300);
            }

            return "fail";
          }
          
          
        async function closeRepliesPane() {
            const selectors = [
                '[data-tid="close-l2-view-button"]',
                '[data-tid="channel-replies-header"] button[aria-label*="Back"]',
                '[data-tid="channel-replies-header"] button[aria-label*="Close"]',
                'button[aria-label*="Back to channel"]',
                '[data-tid="close-replies-button"]',
            ];
            for (const selector of selectors) {
                const btn = document.querySelector<HTMLButtonElement>(selector);
                if (btn && btn.offsetParent !== null) {
                    try { btn.click(); } catch { /* noop: best-effort close; the Escape fallback and runway poll below recover */ }
                    await sleep(200);
                    break;
                }
            }
            try {
                document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', code: 'Escape', keyCode: 27, which: 27, bubbles: true }));
            } catch { /* noop: best-effort Escape; the runway poll below verifies whether the pane closed */ }
            const start = Date.now();
            while (getRepliesRunway() && Date.now() - start < 2000) {
                await sleep(100);
            }
            // NEW: let Teams finish layout so next "Open replies" click works reliably
            await sleep(250);
        }

        function buildReplyContext(msg: ExtractedMessage): ReplyContext {
            return {
                author: msg.author || '',
                timestamp: msg.timestamp || '',
                text: msg.text || '',
            };
        }

        const findPaneItemByMessageId = (id: string | null | undefined): Element | null => {
            if (!id) return null;
            const msgNode = qAny(`[data-mid="${cssEscape(id)}"]`);
            return (
                msgNode?.closest('[data-tid="chat-pane-item"]') ||
                msgNode?.closest('[data-tid="channel-pane-message"]') ||
                msgNode?.closest('[data-tid="channel-replies-pane-message"]') ||
                msgNode?.closest('li[role="none"]') ||
                null
            );
        };

        type OpenMode = "pane" | "inline" | "fail";

        function getResponseSurfaceForButton(btn: HTMLButtonElement): HTMLElement | null {
            return btn.closest<HTMLElement>('[data-tid="response-surface"]');
          }
          

        async function realClick(el: HTMLElement) {
        try { el.scrollIntoView({ block: "center" }); } catch { /* noop: positioning aid; the synthetic clicks below still dispatch */ }
        await sleep(80);

        const opts: MouseEventInit = { bubbles: true, cancelable: true, composed: true, view: window };
        el.dispatchEvent(new MouseEvent("pointerdown", opts));
        el.dispatchEvent(new MouseEvent("mousedown", opts));
        el.dispatchEvent(new MouseEvent("pointerup", opts));
        el.dispatchEvent(new MouseEvent("mouseup", opts));
        el.dispatchEvent(new MouseEvent("click", opts));
        }

        // Inline reply nodes tend to carry the parent chain id somewhere in the subtree.
        // We use the chain id to find “reply message-ish” containers.
        function findInlineReplyNodes(surface: Element, parentId: string): HTMLElement[] {
            const sel = `[data-testid="message-body-flex-wrapper"][data-reply-chain-id="${cssEscape(parentId)}"]`;
            const wrappers = Array.from(surface.querySelectorAll<HTMLElement>(sel));
          
            // Promote wrapper -> stable message-body group if possible
            const items = wrappers.map(w => w.closest<HTMLElement>('[id^="message-body-"][role="group"]') || w);
          
            // Dedup
            return Array.from(new Set(items));
          }
          
          

        async function hydrateSparseMessages(agg: Map<string, ContentAggregated>, opts: ScrapeOptions = {}) {
            if (!agg || agg.size === 0) return;

            const needsHydration = (message: ExtractedMessage, item: Element) => {
                const textNeeds = isPlaceholderText(message.text);
                let reactionsNeed = false;
                if (opts.includeReactions) {
                    const hadReactions = Array.isArray(message.reactions) && message.reactions.length > 0;
                    const missingEmoji =
                        hadReactions &&
                        (message.reactions || []).some(r => !r.emoji || !r.emoji.trim());
                    if ((!hadReactions || missingEmoji) && item?.querySelector('[data-tid="diverse-reaction-pill-button"]')) {
                        reactionsNeed = true;
                    }
                }
                let imagesNeed = false;
                if (item?.querySelector('[data-tid="file-preview-root"][amspreviewurl]')) {
                    const atts = Array.isArray(message.attachments) ? message.attachments : [];
                    const missingPreview = !atts.length || atts.some(att => {
                        const href = att.href || '';
                        if (!href) return false;
                        if (!/asm\.skype\.com|asyncgw\.teams\.microsoft\.com/i.test(href)) return false;
                        return !att.dataUrl;
                    });
                    if (missingPreview) imagesNeed = true;
                }
                return { textNeeds, reactionsNeed, imagesNeed, needs: textNeeds || reactionsNeed || imagesNeed };
            };

            let pending: { id: string; item: Element }[] = [];

            for (const [id, entry] of agg.entries()) {
                const msg = entry.message as ExtractedMessage | undefined;
                if (!msg || msg.system) continue;
                const item = findPaneItemByMessageId(id);
                if (!item) continue;
                const status = needsHydration(msg, item);
                if (status.needs) pending.push({ id, item });
            }

            if (!pending.length) return;

            let attempts = 0;
            while (pending.length && attempts < 3) {
                await sleep(attempts === 0 ? 450 : 650);
                const nextPending: { id: string; item: Element }[] = [];

                for (const task of pending) {
                    const { id } = task;
                    const existing = agg.get(id);
                    if (!existing || !existing.message) continue;

                    const item = findPaneItemByMessageId(id) || task.item;
                    if (!item) continue;

                    const statusBefore = needsHydration(existing.message, item);
                    if (statusBefore.imagesNeed) {
                        await waitForPreviewImages(item, attempts === 0 ? 350 : 700);
                    }

                    const lastAuthorRef = { value: existing.message.author || '' };
                    const ts = existing.message.timestamp ? Date.parse(existing.message.timestamp) : undefined;
                    const tempOrderCtx: OrderContext = {
                        lastTimeMs: Number.isNaN(ts) ? null : ts ?? null,
                        yearHint: Number.isNaN(ts) ? null : (ts ? new Date(ts).getFullYear() : null),
                        seqBase: Date.now(),
                        seq: 0,
                        lastAuthor: existing.message.author || '',
                        lastId: existing.message.id || null,
                        systemCursor: 0,
                    };

                    const reExtracted = await extractOne(
                        item,
                        {
                            includeSystem: opts.includeSystem,
                            includeReactions: opts.includeReactions,
                            includeReplies: opts.includeReplies,
                            startAtISO: null,
                            endAtISO: null,
                        },
                        lastAuthorRef,
                        tempOrderCtx
                    );

                    if (!reExtracted?.message) {
                        nextPending.push(task);
                        continue;
                    }

                    // --- HARD OVERRIDE STRATEGY ---
                    const merged: ExtractedMessage = {
                        // Keep a stable id if we already had one
                        id: existing.message.id || reExtracted.message.id || id,

                        // Prefer fresh extraction for content fields
                        author: reExtracted.message.author || existing.message.author || '',
                        timestamp: reExtracted.message.timestamp || existing.message.timestamp || '',
                        text: reExtracted.message.text || existing.message.text || '',

                        edited: Boolean(existing.message.edited || reExtracted.message.edited),
                        system: Boolean(existing.message.system || reExtracted.message.system),

                        // Avatar: new one wins, otherwise fall back
                        avatar: reExtracted.message.avatar ?? existing.message.avatar ?? null,

                        // Reactions / attachments / tables: trust the new extraction
                        reactions: reExtracted.message.reactions || existing.message.reactions || [],
                        attachments: reExtracted.message.attachments || existing.message.attachments || [],
                        tables: reExtracted.message.tables || existing.message.tables || [],

                        // Reply context: new replyTo wins if present
                        replyTo: reExtracted.message.replyTo ?? existing.message.replyTo ?? null,
                    };

                    const newTsMs =
                        reExtracted.tsMs ??
                        existing.tsMs ??
                        (merged.timestamp ? parseTimeStamp(merged.timestamp) : null);

                    const kind = existing.kind ?? reExtracted.kind;

                    agg.set(id, {
                        message: merged as ExtractedMessage,
                        orderKey: existing.orderKey,
                        tsMs: newTsMs,
                        kind,
                    });

                    const status = needsHydration(merged, item);
                    if (status.needs) {
                        nextPending.push({ id, item });
                    }
                }

                pending = nextPending;
                attempts++;
            }

            if (pending.length) {
                try {
                    console.debug(
                        '[Teams Exporter] hydration pending after retries',
                        pending.map(p => p.id)
                    );
                } catch (_) {
                    // ignore
                }
            }

        }

        function parseDateDividerText(txt: string, yearHint?: number | null) {
            if (!txt) return null;
            const monthMap = {
                january: 1, february: 2, march: 3, april: 4, may: 5, june: 6,
                july: 7, august: 8, september: 9, october: 10, november: 11, december: 12
            };

            const clean = txt.trim().replace(/\s+/g, ' ');
            const currentYear = typeof yearHint === 'number' ? yearHint : (yearHint ? Number(yearHint) : new Date().getFullYear());

            const tryBuild = (dayStr: string, monthStr: string, yearStr?: string | null) => {
                if (!dayStr || !monthStr) return null;
                const day = Number(dayStr);
                if (!Number.isFinite(day)) return null;
                const monthIdx = monthMap[monthStr.toLowerCase() as keyof typeof monthMap];
                if (!monthIdx) return null;
                const year = yearStr ? Number(yearStr) : currentYear;
                if (!Number.isFinite(year)) return null;
                const dt = new Date(year, monthIdx - 1, day);
                if (Number.isNaN(dt.getTime())) return null;
                return dt.getTime();
            };

            let m = clean.match(/^(?:[A-Za-z]+,\s*)?([A-Za-z]+)\s+(\d{1,2})(?:,\s*(\d{4}))?$/);
            if (m) {
                const ts = tryBuild(m[2], m[1], m[3]);
                if (ts != null) return ts;
            }

            m = clean.match(/^(\d{1,2})\s+([A-Za-z]+)(?:\s+(\d{4}))?$/);
            if (m) {
                const ts = tryBuild(m[1], m[2], m[3]);
                if (ts != null) return ts;
            }

            return null;
        }
        // Teams system messages emit timestamps as "MM/DD HH:MM AM/PM" or
        // "MM/DD/YY HH:MM AM/PM" (the year-bearing form appears on older
        // messages, notably "Meeting ended" controls). Without the year capture,
        // the parse failed and the message fell back to lastTimeMs at discovery,
        // landing it at the wrong place in the timeline (often at the very end).
        const controlTimeRe = /(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?\s+(\d{1,2}):(\d{2})\s*(AM|PM)/i;

        function parseControlTimestamp(text: string, yearHint?: number | null): number | null {
            if (!text) return null;
            const match = controlTimeRe.exec(text);
            if (!match) return null;
            const month = Number(match[1]);
            const day = Number(match[2]);
            const yearStr = match[3];
            let hour = Number(match[4]);
            const minute = Number(match[5]);
            const period = match[6];
            if (!Number.isFinite(month) || !Number.isFinite(day) || !Number.isFinite(hour) || !Number.isFinite(minute)) return null;
            if (period?.toUpperCase() === 'PM' && hour < 12) hour += 12;
            if (period?.toUpperCase() === 'AM' && hour === 12) hour = 0;
            let year: number;
            if (yearStr) {
                const n = Number(yearStr);
                year = n < 100 ? 2000 + n : n;
            } else {
                year = typeof yearHint === 'number' ? yearHint : new Date().getFullYear();
            }
            const date = new Date(year, month - 1, day, hour, minute, 0, 0);
            return Number.isNaN(date.getTime()) ? null : date.getTime();
        }

        const makeDayDivider = (dayKey: number, ts: number): ContentAggregated => {
            const base = buildDayDivider(dayKey, ts);
            return { ...base, message: base.message as ExtractedMessage };
        };

        function buildReplyContextFromChainId(chainId: string): ReplyContext | null {
            if (!chainId) return null;
            const anchor = qAny(`[data-mid="${cssEscape(chainId)}"]`);
            if (!anchor) return null;
            const parentItem =
                anchor.closest('[data-tid="channel-pane-message"]') ||
                anchor.closest('[data-tid="chat-pane-message"]') ||
                anchor.closest('[data-tid="channel-replies-pane-message"]') ||
                anchor.closest('li[role="none"]') ||
                anchor;
            const parentBody =
                parentItem.querySelector<HTMLElement>('[id^="message-body-"][aria-labelledby]') ||
                parentItem.querySelector<HTMLElement>('[data-tid="channel-pane-message"]') ||
                parentItem.querySelector<HTMLElement>('[data-tid="chat-pane-message"]') ||
                parentItem.querySelector<HTMLElement>('[data-tid="channel-replies-pane-message"]') ||
                (parentItem as HTMLElement);
            const author = resolveAuthor(parentBody, '');
            const timestamp = resolveTimestamp(parentItem);
            const contentEl = $('[id^="content-"]', parentBody) || $('[data-tid="message-body"]', parentBody) || parentBody;
            const text = extractTextWithEmojis(contentEl).trim();
            if (!author && !timestamp && !text) return null;
            return { author, timestamp, text, id: chainId };
        }

  
        // Extract one item into a message object + an orderKey
        async function extractOne(
            item: Element,
            opts: InternalScrapeOptions,
            lastAuthorRef: { value: string },
            orderCtx: OrderContext & { seq?: number }
        ): Promise<ContentAggregated | null> {
            // --- Special handling for thread replies --------------------------
            const isReplyItem =
                item instanceof HTMLElement &&
                item.getAttribute('data-tid') === 'channel-replies-pane-message';

            let wrapperWithMid: HTMLElement | null;
            let body: HTMLElement | null;
            let itemScope: Element;
            let hasMessage: boolean;

            if (isReplyItem) {
                // In the replies runway, `item` is already the message container.
                // Do NOT climb up with `closest`, or you risk grabbing the parent post.
                wrapperWithMid = item.querySelector<HTMLElement>('[data-mid]') || null;
                body = item;
                itemScope = item;
                hasMessage = true;
            } else {
                // --- Original non-reply initialization path -------------------
                wrapperWithMid =
                    item.querySelector<HTMLElement>('[data-testid="message-body-flex-wrapper"][data-mid]') ||
                    item.querySelector<HTMLElement>('[data-tid="channel-replies-pane-message"] [data-mid]') ||
                    item.querySelector<HTMLElement>('[data-mid]');

                const wrapperItem = wrapperWithMid
                    ? wrapperWithMid.closest<HTMLElement>(
                        '[data-tid="channel-replies-pane-message"], [data-tid="channel-pane-message"], [data-tid="chat-pane-message"], [id^="message-body-"][aria-labelledby]'
                    )
                    : null;

                body =
                    wrapperItem ||
                    item.querySelector<HTMLElement>('[data-tid="chat-pane-message"]') ||
                    item.querySelector<HTMLElement>('[data-tid="channel-pane-message"]') ||
                    (item instanceof HTMLElement && item.matches('[id^="message-body-"][aria-labelledby]') ? item : null) ||
                    item.querySelector<HTMLElement>('[id^="message-body-"][aria-labelledby]') ||
                    item.querySelector<HTMLElement>('[data-tid="channel-replies-pane-message"]') ||
                    (item as HTMLElement);

                itemScope =
                    wrapperItem ||
                    item.closest('[data-tid="channel-pane-message"], [data-tid="chat-pane-message"], [data-tid="channel-replies-pane-message"]') ||
                    item;

                hasMessage =
                    Boolean($('[data-tid="chat-pane-message"]', item)) ||
                    Boolean($('[data-tid="channel-pane-message"]', item)) ||
                    Boolean($('[data-tid="channel-replies-pane-message"]', item)) ||
                    (item instanceof HTMLElement && item.matches('[id^="message-body-"][aria-labelledby]')) ||
                    Boolean($('[id^="message-body-"][aria-labelledby]', item));
            }

            const isSystem = !hasMessage;

            // --- Skip inline thread preview replies when "Open X replies" exists ---
            //
            // In a Teams channel:
            //   - The main post is *outside* the response-surface.
            //   - Inline preview replies live *inside* a `data-tid="response-surface"` block
            //     which also hosts the "Open X replies" button.
            //
            // We don't want to export those preview replies, because we will scrape the
            // *full* thread from the right-hand replies pane instead. If there is NO
            // summary button, then we keep the inline replies (short threads).
            if (!isSystem && body) {
                const surface =
                  body.closest<HTMLElement>('[data-tid="response-surface"]') ||
                  (itemScope as HTMLElement).closest<HTMLElement>('[data-tid="response-surface"]');
              
                if (surface) {
                  const hasOpenButton = surface.querySelector<HTMLButtonElement>('[data-tid="response-summary-button"]');
              
                  // ✅ Only skip inline preview replies during the MAIN channel pass.
                  // Allow them when we're explicitly scraping inline thread replies.
                  if (hasOpenButton && !opts.__allowInlineThreadReplies) {
                    return null;
                  }
                }
              }
              
    
            // --- System / divider handling -----------------------------------
            if (isSystem) {
                if (!opts.includeSystem) return null;

                const dividerWrapper =
                    (item instanceof HTMLElement && item.matches('.fui-Divider__wrapper'))
                        ? item
                        : $('.fui-Divider__wrapper', item);

                const controlRenderer =
                    (item instanceof HTMLElement && item.matches('[data-tid="control-message-renderer"]'))
                        ? item
                        : $('[data-tid="control-message-renderer"]', item);

                // Pure date divider (no control renderer)
                if (dividerWrapper && !controlRenderer) {
                    const text = textFrom(dividerWrapper) || 'system';
                    const bodyMid =
                        dividerWrapper.getAttribute?.('data-mid') ||
                        $('[data-mid]', dividerWrapper)?.getAttribute('data-mid') ||
                        item.getAttribute('data-mid') ||
                        dividerWrapper.id;

                    const numericMid = bodyMid && Number(bodyMid);
                    const parsedTs = parseDateDividerText(text, orderCtx.yearHint);
                    const tsVal = Number.isFinite(parsedTs)
                        ? (parsedTs as number)
                        : Number.isFinite(numericMid)
                            ? Number(numericMid)
                            : Date.now();

                    return makeDayDivider(tsVal, tsVal);
                }

                const wrapper = controlRenderer || dividerWrapper || item;
                const text = textFrom(wrapper) || textFrom(item) || 'system';
                const bodyMid =
                    wrapper?.getAttribute?.('data-mid') ||
                    $('[data-mid]', wrapper || item)?.getAttribute('data-mid') ||
                    item.getAttribute('data-mid') ||
                    wrapper?.id;

                const dividerId = (bodyMid || text || 'system').toLowerCase();
                const numericMid = bodyMid && Number(bodyMid);

                let parsedTs = parseDateDividerText(text, orderCtx.yearHint);
                if (!Number.isFinite(parsedTs)) parsedTs = parseControlTimestamp(text, orderCtx.yearHint);

                const systemCursor = typeof orderCtx.systemCursor === 'number' ? orderCtx.systemCursor : -9e15;
                const approxMs: number = Number.isFinite(parsedTs)
                    ? (parsedTs as number)
                    : Number.isFinite(numericMid)
                        ? Number(numericMid)
                        : typeof orderCtx.lastTimeMs === 'number'
                            ? orderCtx.lastTimeMs - 1
                            : systemCursor;

                orderCtx.systemCursor = systemCursor + 1;

                if (Number.isFinite(parsedTs)) {
                    orderCtx.lastTimeMs = parsedTs as number;
                    orderCtx.yearHint = new Date(parsedTs as number).getFullYear();
                }

                return {
                    message: {
                        id: dividerId,
                        author: '[system]',
                        timestamp: '',
                        text,
                        reactions: [],
                        attachments: [],
                        edited: false,
                        avatar: null,
                        replyTo: null,
                        system: true,
                    },
                    orderKey: approxMs,
                    tsMs: approxMs,
                    kind: 'system-control',
                };
            }

            // --- Normal message (chat, channel, or reply) --------------------
            if (!body) body = itemScope as HTMLElement;

            // Compute mid once, for logging + ID + timestamp fallback.
            const mid =
                wrapperWithMid?.getAttribute('data-mid') ||
                body.getAttribute('data-mid') ||
                body.querySelector('[data-mid]')?.getAttribute('data-mid') ||
                item.getAttribute('data-mid') ||
                item.querySelector('[data-mid]')?.getAttribute('data-mid') ||
                item.id ||
                '';

            if (!mid) {
                try {
                    console.warn('[Teams Exporter] message with no data-mid:', (item as HTMLElement).outerHTML.slice(0, 200));
                } catch {
                    // ignore logging failure
                }
            }

            let ts = resolveTimestamp(item);
            let tms = ts ? Date.parse(ts) : NaN;

            const author = resolveAuthor(body, lastAuthorRef.value || orderCtx.lastAuthor || '');
            if (author) {
                lastAuthorRef.value = author;
                orderCtx.lastAuthor = author;
            }

            await expandSeeMore(item);

            // Prefer the content-block that corresponds to this mid when possible.
            const contentEl: Element =
                (mid
                    ? body.querySelector<HTMLElement>(`[data-tid="message-body"][data-mid="${cssEscape(mid)}"]`) ||
                    body
                        .querySelector<HTMLElement>(`[data-tid="message-body"] [data-mid="${cssEscape(mid)}"]`)
                        ?.closest<HTMLElement>('[data-tid="message-body"]') ||
                    null
                    : null) ||
                $('[id^="content-"]', body) ||
                $('[data-tid="message-content"]', body) ||
                body;

            // For replies, `body === itemScope` (the reply message) so this should
            // no longer “see” the parent post’s body.
            const tables = extractTables(contentEl);
            const codeBlocks = extractCodeBlocks(contentEl);

            const cleanRoot = stripQuotedPreview(contentEl) || contentEl;
            normalizeMentions(cleanRoot);

            let text = extractRichTextAsMarkdown(cleanRoot);


            // Subject line only really applies to top-level channel posts;
            // replies typically won't have it.
            const subjectEl = $('[data-tid="subject-line"]', item) || $('h2[data-tid="subject-line"]', item);
            const subject = textFrom(subjectEl).trim();
            if (subject) {
                const normalizedSubject = subject.replace(/\s+/g, ' ').trim();
                const normalizedText = (text || '').replace(/\s+/g, ' ').trim();
                if (!normalizedText.startsWith(normalizedSubject)) {
                    text = text ? `${subject}\n\n${text}` : subject;
                }
            }

            if (codeBlocks.length && !/```/.test(text)) {
                const fenced = codeBlocks.map(block => `\n\`\`\`\n${block}\n\`\`\`\n`).join('\n');
                text = text ? `${text}\n${fenced}` : fenced.replace(/^\n/, '');
            }

            const edited = resolveEdited(itemScope, body);
            const av = resolveAvatar(itemScope);
            const avatar = av?.dataUrl ?? av?.url ?? null;
            const avatarUrl = av?.url;
            const reactions = opts.includeReactions ? await extractReactions(itemScope) : [];

            await waitForPreviewImages(itemScope, 250);
            const attachments = await extractAttachments(itemScope, body);

            const chainId =
                body.getAttribute('data-reply-chain-id') ||
                body.querySelector('[data-reply-chain-id]')?.getAttribute('data-reply-chain-id') ||
                (itemScope as HTMLElement).getAttribute('data-reply-chain-id') ||
                itemScope.querySelector('[data-reply-chain-id]')?.getAttribute('data-reply-chain-id') ||
                item.getAttribute('data-reply-chain-id') ||
                item.querySelector('[data-reply-chain-id]')?.getAttribute('data-reply-chain-id');

            const threadId =
                chainId ||
                mid ||
                null;
              

            let replyTo = opts.includeReplies === false ? null : extractReplyContext(item, body);

            // Timestamp fallback from mid (some mids are ms since epoch).
            // 1e11 ms ≈ March 1973 — anything smaller can't be a real
            // Teams message timestamp, so it's a non-timestamp mid we
            // shouldn't accidentally interpret as a date.
            const MIN_PLAUSIBLE_MS_TIMESTAMP = 1e11;
            if ((!ts || Number.isNaN(tms)) && mid) {
                const midMs = Number(mid);
                if (Number.isFinite(midMs) && midMs > MIN_PLAUSIBLE_MS_TIMESTAMP) {
                    tms = midMs;
                    ts = new Date(midMs).toISOString();
                }
            }

            if (!Number.isNaN(tms)) {
                orderCtx.lastTimeMs = tms;
                orderCtx.yearHint = new Date(tms).getFullYear();
            }

            if (!replyTo && opts.includeReplies !== false && chainId && chainId !== mid) {
                replyTo = buildReplyContextFromChainId(chainId) || { author: '', timestamp: '', text: '', id: chainId };
            }

            const finalMid = mid || `${ts}#${author}`;
            const msg: ExtractedMessage = {
                id: finalMid,
                threadId,
                author,
                timestamp: ts,
                text,
                reactions,
                attachments,
                edited,
                avatar,
                avatarUrl,
                replyTo,
                tables,
                system: false,
            };

            const seqVal = orderCtx.seq ?? 0;
            orderCtx.seq = seqVal + 1;

            const orderKey = !Number.isNaN(tms) ? tms : orderCtx.seqBase + seqVal;
            const tsMs = !Number.isNaN(tms) ? tms : null;

            return { message: msg, orderKey, tsMs, kind: 'message' };
        }

          

          async function collectRepliesForThread(
            parentId: string,
            parentContext: ReplyContext,
            btn: HTMLButtonElement,
            includeReactions: boolean,
          ): Promise<ExtractedMessage[]> {
            const mode = await openRepliesForItem(btn, parentId);
            if (mode === "fail") return [];
          
            // INLINE MODE: scrape inline-expanded replies under the post
            if (mode === "inline") {
              const surface = getResponseSurfaceForButton(btn);
              if (!surface) return [];
          
              const nodes = findInlineReplyNodes(surface, parentId);
          
              const replies: ExtractedMessage[] = [];
              const seenIds = new Set<string>();
          
              const lastAuthorRef = { value: "" };
              const tempOrderCtx: OrderContext = {
                lastTimeMs: null,
                yearHint: null,
                seqBase: Date.now(),
                seq: 0,
                lastAuthor: "",
                lastId: null,
                systemCursor: -9e15,
              };
          
              for (let i = 0; i < nodes.length; i++) {
                const node = nodes[i];
          
                const extracted = await extractOne(
                    node,
                    {
                      includeSystem: false,
                      includeReactions,
                      includeReplies: false,
                      startAtISO: null,
                      endAtISO: null,
                      __allowInlineThreadReplies: true, // ✅
                    },
                    lastAuthorRef,
                    tempOrderCtx,
                  );
                  
          
                if (extracted?.message && extracted.kind === "message") {
                  const msg = extracted.message as ExtractedMessage;
          
                  const replyId = msg.id || "";
                  if (!replyId || replyId === parentId) continue;
                  if (seenIds.has(replyId)) continue;
          
                  if (!msg.replyTo) msg.replyTo = parentContext;
          
                  seenIds.add(replyId);
                  replies.push(msg);
                }
              }
          
              // In inline mode there is no replies pane to close.
              return replies;
            }
          
            // PANE MODE: scrape the right-hand replies pane
            let replies: ExtractedMessage[] = [];
            try {
              const scroller = getRepliesScroller();
              if (!scroller) {
                await closeRepliesPane();
                return [];
              }
          
              replies = await autoScrollAggregateHelper(
                {
                  runtime,
                  extractOne,
                  hydrateSparseMessages: async () => {},
                  getScroller: () => scroller,
                  getItems: getRepliesItems,
                  getItemId: getReplyItemId,
                  isLoading: isRepliesLoading,
                  makeDayDivider,
                  tuning: {
                    dwellMs: 350,
                    maxStagnant: 6,
                    maxStagnantAtTop: 3,
                    loadingStallPasses: 3,
                    loadingExtraDelayMs: 150,
                  },
                },
                {
                  includeSystem: false,
                  includeReactions,
                  includeReplies: false,
                  startAtISO: null,
                  endAtISO: null,
                },
                currentRunStartedAt,
              ) as ExtractedMessage[];
          
              // Defensive pass: grab any visible replies that the scroll loop missed.
              const seenIds = new Set<string>();
              for (const reply of replies) {
                if (reply?.id) seenIds.add(reply.id);
              }
          
              const visible = getRepliesItems();
              if (visible.length) {
                const lastAuthorRef = { value: "" };
                const tempOrderCtx: OrderContext = {
                  lastTimeMs: null,
                  yearHint: null,
                  seqBase: Date.now(),
                  seq: 0,
                  lastAuthor: "",
                  lastId: null,
                  systemCursor: -9e15,
                };
          
                for (let i = 0; i < visible.length; i++) {
                  const node = visible[i];
                  const idCandidate = getReplyItemId(node, i);
                  if (idCandidate && seenIds.has(idCandidate)) continue;
          
                  const extracted = await extractOne(
                    node,
                    {
                      includeSystem: false,
                      includeReactions,
                      includeReplies: false,
                      startAtISO: null,
                      endAtISO: null,
                      __allowInlineThreadReplies: true, // ✅
                    },
                    lastAuthorRef,
                    tempOrderCtx,
                  );
                  
          
                  if (extracted?.message && extracted.kind === "message") {
                    const msg = extracted.message as ExtractedMessage;
          
                    replies.push(msg);
                    if (msg.id) seenIds.add(msg.id);
                  }
                }
              }
            } catch (err) {
              console.warn("[Teams Exporter] failed to scrape replies", err);
            }
          
            const filtered: ExtractedMessage[] = [];
            for (const reply of replies) {
              const replyId = reply.id || "";
              if (!replyId || replyId === parentId) continue;
              if (!reply.replyTo) reply.replyTo = parentContext;
              filtered.push(reply);
            }
          
            await closeRepliesPane();
            return filtered;
          }
          

        function findOpenRepliesButton(itemRoot: Element): HTMLButtonElement | null {
            // Lock to the current message container / list item only
            const li =
              itemRoot.closest('li[role="none"]') ||
              itemRoot.closest('[data-tid="channel-pane-message"]') ||
              itemRoot;
          
            // Search ONLY inside this item
            return li.querySelector<HTMLButtonElement>(
              '[data-tid="response-surface"] button[data-tid="response-summary-button"]'
            );
          }
          

          function createReplyCollector() {
            const processed = new Set<string>();
            const repliesByParent = new Map<string, ExtractedMessage[]>();
          
            // NEW: serialize all reply scraping
            let queue: Promise<void> = Promise.resolve();
            const enqueue = (fn: () => Promise<void>) => {
              queue = queue.then(fn).catch(err => {
                console.warn("[teams-export] reply queue error", err);
              });
              return queue;
            };
          
            const maybeCollect = async (item: Element, message: ExtractedMessage | undefined, includeReactions: boolean) => {
              return enqueue(async () => {
                // EVERYTHING in here now runs one-at-a-time
          
                const itemRoot =
                item.closest('[data-tid="channel-pane-message"]') ||
                item.closest('li[role="none"]') ||
                item;
              
                const btn = findOpenRepliesButton(itemRoot);
                if (!btn) return;
          
                const chainId =
                  (itemRoot as HTMLElement).getAttribute('data-reply-chain-id') ||
                  itemRoot.querySelector('[data-reply-chain-id]')?.getAttribute('data-reply-chain-id') ||
                  '';
          
                const parentId =
                  chainId ||
                  deriveParentIdFromItem(itemRoot) ||
                  (message?.id || null);
          
                if (!parentId) {
                  console.warn("[teams-export] Found replies button but could not derive parentId");
                  return;
                }
                if (processed.has(parentId)) return;
                processed.add(parentId);
          
                const parentContext =
                  buildReplyContextFromChainId(parentId) ||
                  (message ? buildReplyContext(message) : { author: "", timestamp: "", text: "", id: parentId });
          
                const replies = await collectRepliesForThread(parentId, parentContext, btn, includeReactions);
          
                if (replies.length) repliesByParent.set(parentId, replies);
          
                // NEW: tiny settle delay so Teams finishes animations/layout
                await sleep(250);
              });
            };
          
            return { maybeCollect, repliesByParent };
          }
          
        function mergeRepliesIntoMessages(
            messages: ExtractedMessage[],
            repliesByParent: Map<string, ExtractedMessage[]>,
            ) {
            if (!repliesByParent.size) return messages;

            const baseIdToIndices = new Map<string, number[]>();
            messages.forEach((m, idx) => {
                if (!m?.id) return;
                const list = baseIdToIndices.get(m.id) || [];
                list.push(idx);
                baseIdToIndices.set(m.id, list);
            });

            // Any base (channel) message that also appears in the thread pane
            // should be hidden at top level and only shown inside the thread.
            const suppressedBaseIndices = new Set<number>();

            for (const replies of repliesByParent.values()) {
                for (const reply of replies) {
                if (!reply?.id) continue;
                const indices = baseIdToIndices.get(reply.id);
                if (!indices) continue;
                for (const idx of indices) {
                    suppressedBaseIndices.add(idx);
                }
                }
            }

            const out: ExtractedMessage[] = [];
            const existingIds = new Set<string>();
            const insertedParents = new Set<string>();

            // 1) Push channel messages in original order, but skip any that we know
            //    also appear in the thread (same author + timestamp + text).
            for (let i = 0; i < messages.length; i++) {
              if (suppressedBaseIndices.has(i)) continue; // drop inline preview duplicate
            
              const msg = messages[i];
              if (msg.id) existingIds.add(msg.id);
              out.push(msg);
            
              // Try to locate replies using multiple possible parent keys
              const keysToTry = [msg.threadId, msg.id].filter(Boolean) as string[];
            
              let replies: ExtractedMessage[] | undefined;
              let usedKey: string | null = null;
            
              for (const k of keysToTry) {
                const got = repliesByParent.get(k);
                if (got?.length) {
                  replies = got;
                  usedKey = k;
                  break;
                }
              }
            
              if (!replies || !replies.length || !usedKey) continue;
            
              // mark the actual parent key we matched so step (3) doesn't re-append later
              insertedParents.add(usedKey);

              // 2) Append replies for this parent, deduping by id.
              for (const reply of replies) {
                if (reply.id && existingIds.has(reply.id)) continue;
                if (reply.id) existingIds.add(reply.id);
                out.push(reply);
              }
            }
            
            // 3) Any replies whose parent wasn't in the main messages (rare) go at the end.
            for (const [parentId, replies] of repliesByParent.entries()) {
                if (insertedParents.has(parentId)) continue;
                for (const reply of replies) {
                if (reply.id && existingIds.has(reply.id)) continue;
                if (reply.id) existingIds.add(reply.id);
                out.push(reply);
                }
            }

            return out;
        }



        // Remove quoted/preview blocks from a cloned content node so root "text" doesn't include them
        function stripQuotedPreview(container: Element | null): Element | null {
            if (!container) return container;
            const clone = container.cloneNode(true) as Element;

            // Known containers for quoted/preview content
            const kill = [
                '[data-tid="quoted-reply-card"]',
                '[data-tid="referencePreview"]',
                '[role="group"][aria-label^="Begin Reference"]',
                'table[itemprop="copy-paste-table"]'
            ];
            for (const sel of kill) {
                clone.querySelectorAll(sel).forEach((n: Element) => n.remove());
            }
            const cardSelectors = ['[data-tid="adaptive-card"]', '.ac-adaptiveCard', '[aria-label*="card message"]'];
            clone.querySelectorAll(cardSelectors.join(',')).forEach((n: Element) => {
                if (n.querySelector('pre, code, .cm-line')) return;
                n.remove();
            });

            // Headings like "Begin Reference, …"
            clone.querySelectorAll('div[role="heading"]').forEach((h: Element) => {
                const txt = textFrom(h);
                if (/^Begin Reference,/i.test(txt)) h.remove();
            });

            return clone;
        }

        // Helper-dependent diagnostic probes. These live inside main()
        // because they close over the page-world urlp helper RPC
        // (constants and helper-loaded promise are all main-scope).
        // The standalone probes live in src/content/probes.ts.
        async function probePageWorldHelper(): Promise<ProbeResult> {
            const t0 = performance.now();
            // ensureUrlpHelperLoaded itself has a 5 s internal safety
            // timeout. We do NOT add a second race here because that
            // would create a deceptive "loaded vs timeout" ambiguity
            // (which timer wins is scheduler-dependent). Instead we
            // rely on the helper to report its own outcome.
            const status = await ensureUrlpHelperLoaded();
            const ms = Math.round(performance.now() - t0);
            if (status === 'ready') {
                return { name: 'page_world_helper', status: 'pass', ms };
            }
            if (status === 'script-error') {
                return { name: 'page_world_helper', status: 'fail', detail: 'script load error', ms };
            }
            return { name: 'page_world_helper', status: 'fail', detail: 'timeout: helper never signalled ready', ms };
        }

        async function probeCanaryImageFetch(): Promise<ProbeResult> {
            const t0 = performance.now();
            // Deliberately invalid AMS object id. The helper will issue a
            // real fetch via page cookies; we want to know the RPC path
            // works end to end. 4xx is a healthy "round-tripped, server
            // rejected as expected". 5xx means AMS itself is misbehaving
            // and is a real diagnostic signal — report it as fail.
            const url = 'https://us-api.asm.skype.com/v1/objects/probe-canary/views/imgo';
            const resp = await fetchUrlpDirect(url);
            const ms = Math.round(performance.now() - t0);
            if (resp.ok) {
                return { name: 'canary_image_fetch', status: 'pass', detail: 'returned bytes', ms };
            }
            if (typeof resp.status === 'number') {
                const is5xx = resp.status >= 500 && resp.status < 600;
                return {
                    name: 'canary_image_fetch',
                    status: is5xx ? 'fail' : 'pass',
                    detail: is5xx ? `HTTP ${resp.status} (server error)` : `HTTP ${resp.status}`,
                    ms,
                };
            }
            return { name: 'canary_image_fetch', status: 'fail', detail: resp.error || 'unknown failure', ms };
        }

        async function runAllContentProbes(): Promise<{ results: ProbeResult[]; totalMs: number }> {
            const t0 = performance.now();
            // Standalone probes (DOM, auth tokens, host reachability) and
            // the helper probes run in parallel: the slowest network call
            // sets the floor instead of stacking.
            const standalonePromise = runStandaloneProbes();
            const helperPromise = probePageWorldHelper();
            const standalone = await standalonePromise;
            const helper = await helperPromise;
            // Canary depends on the helper being loaded. We chain it after
            // the helper probe rather than racing in parallel; otherwise a
            // failed helper would also fail the canary with a generic
            // timeout, blurring which RPC stage actually broke.
            const canary = helper.status === 'pass'
                ? await probeCanaryImageFetch()
                : ({ name: 'canary_image_fetch', status: 'skipped' as const, detail: 'helper not loaded', ms: 0 });
            const results: ProbeResult[] = [...standalone, helper, canary];
            return { results, totalMs: Math.round(performance.now() - t0) };
        }

        if (!isTop) return;
        // Bridge --------------------------------------------------------
        runtime.onMessage.addListener((msg, _sender, sendResponse) => {
            (async () => {
                try {
                    if (msg.type === 'PING') { sendResponse({ ok: true }); return; }
                    if (msg.type === 'CHECK_CHAT_CONTEXT') { sendResponse(checkChatContext(msg.target)); return; }
                    if (msg.type === 'GET_DIAGNOSTICS_CONTENT') {
                        try {
                            const idbShape = await probeIdbShape();
                            sendResponse({ idbShape, forwardingStats: diagForwardingStats() });
                        } catch (e) {
                            sendResponse({
                                idbShape: { available: false, reason: e instanceof Error ? e.message : String(e) },
                                forwardingStats: diagForwardingStats(),
                            });
                        }
                        return;
                    }
                    if (msg.type === 'RUN_PROBES_CONTENT') {
                        try {
                            const { results, totalMs } = await runAllContentProbes();
                            sendResponse({ ok: true, results, totalMs });
                        } catch (e) {
                            sendResponse({ ok: false, reason: e instanceof Error ? e.message : String(e) });
                        }
                        return;
                    }
                    if (msg.type === 'DUMP_FILE_FIELDS') {
                        // Diagnostics: scrape the open chat's first message page
                        // and report the raw field NAMES (never values) of each
                        // file attachment record, so we can see whether Teams'
                        // message data carries a sharing-link field (shareUrl /
                        // defaultEncodingURL / ...) that the shares resolver
                        // needs. Read-only; nothing is saved.
                        try {
                            const convId = (await extractConversationId()) || undefined;
                            const res = await apiScrape(undefined, { conversationId: convId });
                            // Walk a file record collecting DOTTED key paths (names
                            // only, never values), 3 levels deep, so a sharing link
                            // nested in fileInfo/providerData/etc. is visible too.
                            const allPaths = new Set<string>();
                            const linkFieldNames = new Set<string>();
                            const walk = (o: unknown, prefix: string, depth: number) => {
                                if (!o || typeof o !== 'object' || depth > 3) return;
                                for (const [k, v] of Object.entries(o as Record<string, unknown>)) {
                                    const path = prefix ? `${prefix}.${k}` : k;
                                    allPaths.add(path);
                                    if (/url|link|share|encoding|relative|path/i.test(k)) linkFieldNames.add(path);
                                    if (v && typeof v === 'object') walk(v, path, depth + 1);
                                }
                            };
                            let fileRecords = 0;
                            for (const m of res?.messages || []) {
                                const rawFiles = (m.properties as Record<string, unknown> | undefined)?.files;
                                if (!rawFiles) continue;
                                let files: Array<Record<string, unknown>>;
                                try { files = typeof rawFiles === 'string' ? JSON.parse(rawFiles) : rawFiles as typeof files; } catch { continue; }
                                if (!Array.isArray(files)) continue;
                                for (const f of files) {
                                    if (!f || typeof f !== 'object') continue;
                                    walk(f, '', 0);
                                    fileRecords++;
                                    if (fileRecords >= 8) break;
                                }
                                if (fileRecords >= 8) break;
                            }
                            sendResponse({
                                ok: true,
                                messages: res?.messages?.length ?? 0,
                                fileRecords,
                                keys: Array.from(allPaths).sort(),
                                linkFields: Array.from(linkFieldNames).sort(),
                            });
                        } catch (e) {
                            sendResponse({ ok: false, error: e instanceof Error ? e.message : String(e) });
                        }
                        return;
                    }
                    if (msg.type === 'DOWNLOAD_RESOLVE_SHARE') {
                        // Files-phase resolver: resolve a file's sharing link to
                        // a pre-authenticated download URL (Bearer-token auth via
                        // the shares API). Called from the background per
                        // attachment during the Files phase. A local semaphore
                        // bounds concurrent shares calls as defence-in-depth on
                        // top of the background's own resolve concurrency cap.
                        try {
                            const shareUrl = String(msg.shareUrl || '');
                            const itemid = String(msg.itemid || '');
                            const href = String(msg.href || '');
                            // sharing link -> drive-item GUID (own uploads) -> raw, via
                            // the shared ladder (keeps size/date even on a URL-less resolve).
                            const { result: out } = await withShareResolveSlot(() =>
                                resolveWithFallback({ shareUrl, itemid, href }));
                            sendResponse({ ok: out.ok, downloadUrl: out.downloadUrl, blocksDownload: out.blocksDownload, size: out.size, lastModifiedDateTime: out.lastModifiedDateTime });
                        } catch (e) {
                            sendResponse({ ok: false, error: e instanceof Error ? e.message : String(e) });
                        }
                        return;
                    }
                    if (msg.type === 'BATCH_RESOLVE_HREFS') {
                        // Salvage tool (Diagnostics page): resolve a LIST of file
                        // links (e.g. from a FAILURES.txt) to pre-authenticated
                        // download URLs. Scrapes the open chat ONCE, builds a
                        // filename -> shareUrl map, then resolves each link. The
                        // popup downloads the resolved URLs serially. downloadUrl
                        // is short-lived — returned to the popup, never logged.
                        try {
                            const hrefs: string[] = Array.isArray(msg.hrefs) ? msg.hrefs.map((h: unknown) => String(h)) : [];
                            // Index the open chat's files by name. A file may carry a
                            // sharing link (shareUrl), a stable GUID (itemid), or both.
                            // Own-uploaded files have an itemid but NO share link, so we
                            // keep any file that has EITHER — the itemid feeds the
                            // drive-item-by-GUID resolve that recovers them.
                            const byName = new Map<string, { shareUrl?: string; itemid?: string; siteUrl?: string; fileName: string }>();
                            const convId = (await extractConversationId()) || undefined;
                            const res = await apiScrape(undefined, { conversationId: convId });
                            for (const m of res?.messages || []) {
                                const rawFiles = (m.properties as Record<string, unknown> | undefined)?.files;
                                if (!rawFiles) continue;
                                let files: Array<Record<string, unknown>>;
                                try { files = typeof rawFiles === 'string' ? JSON.parse(rawFiles) : rawFiles as typeof files; } catch { continue; }
                                if (!Array.isArray(files)) continue;
                                for (const f of files) {
                                    const info = (f?.fileInfo || {}) as Record<string, unknown>;
                                    const shareUrl = typeof info.shareUrl === 'string' ? info.shareUrl : undefined;
                                    const itemid = typeof f?.itemid === 'string' ? f.itemid : undefined;
                                    if (!shareUrl && !itemid) continue;
                                    const siteUrl = String(f?.objectUrl || f?.baseUrl || '');
                                    const rec = { shareUrl, itemid, siteUrl, fileName: String(f?.fileName || '') };
                                    const nm = String(f?.fileName || '').toLowerCase();
                                    const objName = decodeURIComponent(siteUrl.split('?')[0].split('/').pop() || '').toLowerCase();
                                    if (nm && !byName.has(nm)) byName.set(nm, rec);
                                    if (objName && !byName.has(objName)) byName.set(objName, rec);
                                }
                            }
                            const results = await Promise.all(hrefs.map(href => withShareResolveSlot(async () => {
                                const want = decodeURIComponent((href.split('?')[0].split('/').pop() || '')).toLowerCase();
                                const match = byName.get(want);
                                // Matched files resolve by their sharing link and/or GUID
                                // (GUID uses the file's own site URL); an unmatched href
                                // falls back to resolving the raw href as a share token.
                                const { result: out, via } = await resolveWithFallback(
                                    match ? { shareUrl: match.shareUrl, itemid: match.itemid, href: match.siteUrl || href } : { href });
                                return { href, name: match?.fileName || want, ok: out.ok, downloadUrl: out.downloadUrl, blocksDownload: out.blocksDownload, error: out.error, via };
                            })));
                            sendResponse({ ok: true, results });
                        } catch (e) {
                            sendResponse({ ok: false, error: e instanceof Error ? e.message : String(e) });
                        }
                        return;
                    }
                    if (msg.type === 'RESOLVE_SHARE_FILE') {
                        // File-access probe (Diagnostics page): resolve a file to
                        // its pre-authenticated downloadUrl the way Teams does —
                        // via the SHARING LINK (fileInfo.shareUrl), not the raw
                        // file path (which needs host cookies and 401s when the
                        // session is stale). Given a pasted file link, scrape the
                        // open chat, match the file record by name, and resolve
                        // its shareUrl. Falls back to resolving the pasted URL
                        // directly if no match is found. Runs here because the
                        // shares fetch needs the Teams page origin.
                        try {
                            const pasted = String(msg.href || '');
                            const wantName = decodeURIComponent((pasted.split('?')[0].split('/').pop() || '')).toLowerCase();
                            let shareUrl: string | undefined;
                            let matchedName: string | undefined;
                            try {
                                const convId = (await extractConversationId()) || undefined;
                                const res = await apiScrape(undefined, { conversationId: convId });
                                for (const m of res?.messages || []) {
                                    const rawFiles = (m.properties as Record<string, unknown> | undefined)?.files;
                                    if (!rawFiles) continue;
                                    let files: Array<Record<string, unknown>>;
                                    try { files = typeof rawFiles === 'string' ? JSON.parse(rawFiles) : rawFiles as typeof files; } catch { continue; }
                                    if (!Array.isArray(files)) continue;
                                    for (const f of files) {
                                        const info = (f?.fileInfo || {}) as Record<string, unknown>;
                                        const nm = String(f?.fileName || '').toLowerCase();
                                        const objName = decodeURIComponent(String(f?.objectUrl || '').split('?')[0].split('/').pop() || '').toLowerCase();
                                        if (typeof info.shareUrl === 'string' && (nm === wantName || objName === wantName || (wantName && nm && wantName.startsWith(nm.slice(0, 40))))) {
                                            shareUrl = info.shareUrl;
                                            matchedName = String(f?.fileName || '');
                                            break;
                                        }
                                    }
                                    if (shareUrl) break;
                                }
                            } catch { /* scrape failed — fall through to direct resolve */ }
                            const target = shareUrl || pasted;
                            const out = await resolveShareFile(target);
                            sendResponse({ ...out, via: shareUrl ? 'shareUrl' : 'rawUrl', matchedName });
                        } catch (e) {
                            sendResponse({ ok: false, status: 0, error: e instanceof Error ? e.message : String(e) });
                        }
                        return;
                    }
                    if (msg.type === 'GET_CONV_ID') {
                        // Delegates to the same extractor the scraper uses
                        // (IndexedDB → URL → DOM). Used by the popup on
                        // restore to confirm the persisted outcome still
                        // matches the active conversation, and by the
                        // background on cancel to tag the cancelled snapshot.
                        try {
                            const convId = await extractConversationId();
                            sendResponse({ convId: convId ?? null });
                        } catch (e) {
                            console.log('[Teams Exporter] GET_CONV_ID extraction failed:', e);
                            sendResponse({ convId: null });
                        }
                        return;
                    }
                    if (msg.type === 'LIST_CONVERSATIONS_QUICK') {
                        // Fast first-paint variant: pure IDB read, no
                        // Graph or roster network calls. Returns in
                        // ~50 ms. Picker can render immediately;
                        // popup follows up with the full LIST_CONVERSATIONS
                        // for resolved external/internal contact names.
                        try {
                            const { conversations, folders } = await listConversationsFromIdbQuick();
                            sendResponse({ ok: true, conversations, folders });
                        } catch (e) {
                            console.log('[Teams Exporter] LIST_CONVERSATIONS_QUICK failed:', e);
                            sendResponse({ ok: false, error: String((e as Error)?.message || e) });
                        }
                        return;
                    }
                    if (msg.type === 'LIST_CONVERSATIONS') {
                        // Reads the conversation list from Teams' own
                        // IndexedDB store (Teams:conversation-manager).
                        // IDB has the full local conversation set —
                        // including meeting-derived chats and other niche
                        // product types — and is consistent with the
                        // sidebar. The chat-service API path used to be
                        // a fallback; once we proved IDB is reliable
                        // across tenants/locales it became dead code,
                        // so it's gone. If IDB is genuinely empty
                        // (Teams hasn't booted, or the user is on a
                        // brand-new account) we return an empty list —
                        // the popup retry button covers re-fetching
                        // once Teams has populated.
                        try {
                            const skypeToken = await getSkypeToken();
                            if (!skypeToken) { sendResponse({ ok: false, error: 'no-skype-token' }); return; }
                            const ic3Token = await getIc3Token();
                            if (!ic3Token) { sendResponse({ ok: false, error: 'no-ic3-token' }); return; }
                            const { chatServiceUrl } = await discover(skypeToken);
                            const { conversations, folders } = await listConversationsFromIdb({ chatServiceUrl, ic3Token });
                            sendResponse({ ok: true, conversations, folders });
                        } catch (e) {
                            console.log('[Teams Exporter] LIST_CONVERSATIONS failed:', e);
                            sendResponse({ ok: false, error: String((e as Error)?.message || e) });
                        }
                        return;
                    }
                    if (msg.type === 'STOP_SCRAPE') {
                        if (currentAbortController) {
                            console.log('[Teams Exporter] STOP_SCRAPE received — aborting active scrape');
                            currentAbortController.abort();
                            sendResponse({ ok: true, cancelling: true });
                        } else {
                            sendResponse({ ok: false, reason: 'no-active-scrape' });
                        }
                        return;
                    }
                    if (msg.type === 'SCRAPE_TEAMS') {
                        console.log(`[Teams Exporter] export start build=${__BUILD_STAMP__}`);
                        const { startAt, endAt, includeReactions, includeSystem, includeReplies, exportTarget, formats, embedAvatars, downloadImages, fullResImages, conversationId, conversationTitle, noDomFallback } = msg.options || {};
                        const target = exportTarget === 'team' ? 'team' : 'chat';
                        const scrapeOpts = { startAtISO: startAt, endAtISO: endAt, includeSystem, includeReactions, includeReplies: includeReplies !== false };
                        // Skip expensive fetches none of the selected formats
                        // would render (see ScrapeOptions in types/shared.ts).
                        //   TXT / CSV          — never render images or avatars
                        //   JSON / HTML        — only need avatars if embedAvatars=true
                        //   HTML               — only needs image blobs if downloadImages=true
                        // The union across selected formats decides — fetch
                        // when ANY selected format wants the asset.
                        const fmtList: string[] = Array.isArray(formats) ? formats : [];
                        // HTML + PDF render inline images; TXT/CSV/JSON don't.
                        const needImages = (fmtList.includes('html') || fmtList.includes('pdf')) && downloadImages === true;
                        // JSON/HTML/PDF render avatars; TXT/CSV don't.
                        const needAvatars = (fmtList.includes('json') || fmtList.includes('html') || fmtList.includes('pdf')) && embedAvatars === true;
                        console.debug('[Teams Exporter] SCRAPE_TEAMS', location.href, msg.options, 'needImages=' + needImages, 'needAvatars=' + needAvatars);
                        currentRunStartedAt = Date.now();
                        // Fresh AbortController per scrape so a stop only affects the
                        // current run and aborts the IC3 pagination loop, image fetches,
                        // and Graph photo fetches via signal checks they share.
                        currentAbortController = new AbortController();
                        const abortSignal = currentAbortController.signal;

                        // ── API Mode (fast, no scrolling) ──────────────────
                        let messages: ExportMessage[] | null = null;
                        let title: string | null = null;
                        let avatars: Record<string, string> = {};
                        // Chat roster (API mode only): participant list + member count for meta.
                        let participants: import('../types/shared').Participant[] = [];
                        let memberCount: number | undefined;
                        // Partial-export tracking. Set when we detect a high-
                        // confidence "this scrape is incomplete" signal during
                        // the run. Wins are propagated into meta.partial so
                        // the background can rename the file (-PARTIAL),
                        // inject the in-file warning banner, and write a
                        // history entry with kind='partial'. 'network' beats
                        // 'truncation' if both fire (network is the higher-
                        // confidence root cause).
                        let partialReason: 'network' | 'truncation' | null = null;

                        try {
                            const apiResult = await apiScrape((p) => {
                                try {
                                    const mp = runtime.sendMessage({ type: 'SCRAPE_PROGRESS', payload: { phase: 'api-fetch', messagesVisible: p.messagesSoFar, passes: p.page } });
                                    if (mp && mp.catch) mp.catch(() => {});
                                } catch { /* ignore */ }
                            }, { startAtISO: scrapeOpts.startAtISO, signal: abortSignal, conversationId: typeof conversationId === 'string' ? conversationId : null });

                            if (apiResult) {
                                // A salvaged partial chat (a mid-pagination network error kept the
                                // pages fetched so far) flags the export partial-network, the same
                                // as the total-failure path below.
                                if (apiResult.partialNetwork) {
                                    partialReason = 'network';
                                    console.warn('[Teams Exporter] API returned a PARTIAL chat after a network error — flagging export as partial-network');
                                }
                                const converted = convertApiMessages(apiResult.messages, scrapeOpts, apiResult.userId);
                                // Two distinct empty cases:
                                //  - converted.length > 0     : normal — render messages.
                                //  - apiResult.messages.length === 0: API returned an
                                //    actual empty result. This is the genuine state of
                                //    the chat (most often a Teams Free legacy Skype-
                                //    imported 1:1 — Microsoft never migrated those
                                //    histories to the consumer chat backend, so the
                                //    server returns []). Produce an empty-but-valid
                                //    export rather than falling through to DOM scroll;
                                //    the empty export communicates "we asked, got
                                //    nothing" honestly.
                                //  - apiResult.messages.length > 0 && converted.length === 0:
                                //    raw API had messages but our filter dropped them
                                //    all (system-only chat etc.). DOM scroll might
                                //    surface them differently — keep the existing
                                //    fall-through behaviour for that case.
                                if (converted.length > 0 || apiResult.messages.length === 0) {
                                    // Fetch inline images (content script has auth cookies, URLs expire).
                                    // Skipped when the target format won't render them.
                                    if (needImages) {
                                        try { runtime.sendMessage({ type: 'SCRAPE_PROGRESS', payload: { phase: 'images', messagesVisible: converted.length } }); } catch { /* noop: progress ping is best-effort */ }
                                        await fetchInlineImages(converted, (done, total, rateLimited) => {
                                            // While rate-limited, emit on every completion (not every
                                            // 10th) so the bar keeps ticking as retries recover images
                                            // instead of looking frozen across the backoff windows.
                                            if (done % 10 === 0 || done === total || rateLimited) {
                                                try { runtime.sendMessage({ type: 'SCRAPE_PROGRESS', payload: { phase: 'images', messagesVisible: converted.length, imagesDone: done, imagesTotal: total } }); } catch { /* noop: progress ping is best-effort */ }
                                            }
                                        }, { userId: apiResult.userId, userRegion: apiResult.userRegion, ic3Token: apiResult.ic3Token }, fullResImages === true);
                                    } else {
                                        console.log(`[Teams Exporter] Skipping inline image fetch — formats=${fmtList.join(',')}, downloadImages=${downloadImages}`);
                                    }
                                    // Fetch profile photos via Graph API.
                                    // Skipped when the format/options don't render avatars.
                                    if (needAvatars) {
                                        try { runtime.sendMessage({ type: 'SCRAPE_PROGRESS', payload: { phase: 'avatars', messagesVisible: converted.length } }); } catch { /* noop: progress ping is best-effort */ }
                                        await fetchApiAvatars(converted, apiResult.messages);
                                    } else {
                                        console.log(`[Teams Exporter] Skipping avatar fetch — formats=${fmtList.join(',')}, embedAvatars=${embedAvatars}`);
                                    }
                                    messages = converted;
                                    // Prefer the picker-resolved title: when the user explicitly
                                    // picks chat X but their Teams tab is viewing chat Y, DOM
                                    // extraction would stamp Y's name on X's export. Fall through
                                    // to DOM extraction when the caller didn't supply a title.
                                    title = (typeof conversationTitle === 'string' && conversationTitle.trim())
                                      ? conversationTitle.trim()
                                      : (target === 'team' ? extractChannelTitle() : extractChatTitle());
                                    // Roster captured by the API scraper (current members + count).
                                    if (Array.isArray(apiResult.participants)) participants = apiResult.participants;
                                    if (typeof apiResult.memberCount === 'number') memberCount = apiResult.memberCount;
                                    console.log(`[Teams Exporter] API mode: ${converted.length} messages from ${apiResult.messages.length} raw; ${participants.length} participants`);
                                } else {
                                    console.log('[Teams Exporter] API returned messages but all filtered out, trying DOM scroll');
                                }
                            }
                        } catch (apiErr) {
                            // An abort surfaces here as a fetch AbortError; treat
                            // it as cancellation rather than as an API failure
                            // worth falling back to DOM scroll for.
                            if (abortSignal.aborted) throw apiErr;
                            console.log('[Teams Exporter] API mode failed, falling back to DOM scroll:', apiErr);
                        }

                        // ── DOM Scroll Fallback ────────────────────────────
                        if (!messages) {
                            if (abortSignal.aborted) {
                                currentRunStartedAt = null;
                                currentAbortController = null;
                                sendResponse({ ok: false, cancelled: true });
                                return;
                            }
                            // If the API failed because of a network error
                            // specifically (NetworkError / Failed to fetch),
                            // record it as a partial signal. We still try
                            // DOM scroll because Teams' UI may have already
                            // rendered some messages before the network
                            // dropped — those are recoverable from the DOM.
                            // The partial flag flows into meta and produces
                            // the -PARTIAL filename + warning banner.
                            const apiFailure = getLastApiScrapeFailure();
                            if (apiFailure?.isNetworkError) {
                                partialReason = 'network';
                                console.warn('[Teams Exporter] API failed with network error — flagging export as partial-network');
                            }
                            // In multi-chat bundle mode the DOM scroll would
                            // scrape whichever chat is currently visible in
                            // the user's tab — almost certainly NOT the one
                            // we're trying to export. Refuse to fall back
                            // and let the background loop record this as a
                            // per-chat failure. (The api-client already
                            // retries 403/5xx on a tight budget before
                            // surfacing the error here.)
                            if (noDomFallback) {
                                console.log('[Teams Exporter] API failed and noDomFallback is set (multi-chat) — reporting failure without DOM scroll');
                                currentRunStartedAt = null;
                                currentAbortController = null;
                                sendResponse({ ok: false, error: 'API scrape failed; DOM fallback disabled in multi-chat mode', isNetworkError: apiFailure?.isNetworkError === true });
                                return;
                            }
                            const replyCollector = createReplyCollector();
                            const includeRepliesEnabled = includeReplies !== false;

                            const extractWithReplies = async (
                                item: Element,
                                opts: ScrapeOptions,
                                lastAuthorRef: { value: string },
                                orderCtx: OrderContext & { seq?: number },
                            ) => {
                                const extracted = await extractOne(item, opts, lastAuthorRef, orderCtx);
                                if (target === 'team' && includeRepliesEnabled) {
                                  const msg = extracted?.message as ExtractedMessage | undefined;
                                  if (extracted?.kind === "message" && msg && !msg.system) {
                                    await replyCollector.maybeCollect(item, msg, Boolean(includeReactions));
                                  }
                                }
                                return extracted;
                            };

                            let scrollMessages = await autoScrollAggregateHelper(
                                {
                                    runtime,
                                    extractOne: target === 'team' && includeRepliesEnabled ? extractWithReplies : extractOne,
                                    hydrateSparseMessages,
                                    getScroller: () => getScroller(target),
                                    getItems: target === 'team' ? getChannelItems : undefined,
                                    isLoading: target === 'team' ? isVirtualListLoading : undefined,
                                    makeDayDivider,
                                    tuning: target === 'team' ? {
                                        dwellMs: 800,
                                        maxStagnant: 30,
                                        maxStagnantAtTop: 35,
                                        loadingStallPasses: 20,
                                        loadingExtraDelayMs: 700,
                                    } : undefined,
                                },
                                scrapeOpts,
                                currentRunStartedAt,
                                abortSignal,
                            );
                            if (target === 'team' && includeRepliesEnabled) {
                                scrollMessages = mergeRepliesIntoMessages(scrollMessages as ExtractedMessage[], replyCollector.repliesByParent);
                            }
                            messages = scrollMessages;
                            title = target === 'team' ? extractChannelTitle() : extractChatTitle();
                        }

                        try {
                            const msgPromise = runtime.sendMessage({ type: 'SCRAPE_PROGRESS', payload: { phase: 'extract', messagesExtracted: messages.length } });
                            if (msgPromise && msgPromise.catch) msgPromise.catch(() => { });
                        } catch (e) { /* ignore */ }

                        // Patch missing avatars using UUID identity. This covers two cases:
                        //   - Grouped message follow-ups (Teams omits the avatar <img> for consecutive
                        //     messages from the same author, so resolveAvatar returns null for them)
                        //   - The user's own messages (no avatar element in DOM; capture from MeControl)
                        // We use UUID (from the profile-picture URL) as the identity key so two users
                        // with the same display name don't get each other's avatars.
                        {
                            const uuidToAvatar = new Map<string, string>();      // identity → data URL
                            const nameToUuids = new Map<string, Set<string>>();  // name → identities seen

                            const observe = (name: string | undefined, uuid: string | null, dataUrl: string | null) => {
                                if (uuid && dataUrl && !uuidToAvatar.has(uuid)) uuidToAvatar.set(uuid, dataUrl);
                                if (name && uuid) {
                                    let s = nameToUuids.get(name);
                                    if (!s) { s = new Set(); nameToUuids.set(name, s); }
                                    s.add(uuid);
                                }
                            };

                            for (const m of messages) {
                                const url = (m as ExportMessage).avatarUrl;
                                const uuid = extractUserUuidFromAvatarUrl(url);
                                const dataUrl = m.avatar && m.avatar.startsWith('data:') ? m.avatar : null;
                                observe(m.author, uuid, dataUrl);
                            }

                            const self = findSelfAvatar();
                            if (self.dataUrl && self.uuid) {
                                observe(self.name || undefined, self.uuid, self.dataUrl);
                            }

                            let patched = 0, skippedAmbiguous = 0;
                            for (const m of messages) {
                                if (m.avatar) continue;
                                const url = (m as ExportMessage).avatarUrl;
                                const uuid = extractUserUuidFromAvatarUrl(url);
                                if (uuid && uuidToAvatar.has(uuid)) {
                                    m.avatar = uuidToAvatar.get(uuid)!;
                                    patched++;
                                    continue;
                                }
                                if (!m.author) continue;
                                const ids = nameToUuids.get(m.author);
                                if (ids?.size === 1) {
                                    const av = uuidToAvatar.get([...ids][0]);
                                    if (av) { m.avatar = av; patched++; }
                                } else if (ids && ids.size > 1) {
                                    skippedAmbiguous++;
                                }
                            }
                            if (patched > 0) console.log(`[Teams Exporter] Avatar patch: ${patched} messages filled from identity map`);
                            if (skippedAmbiguous > 0) console.log(`[Teams Exporter] Avatar patch: ${skippedAmbiguous} messages skipped (multiple users share the same display name)`);
                        }

                        // Fetch + normalize avatars in content script context.
                        // This also fetches DOM-mode HTTP avatar URLs, so we
                        // skip it entirely when the format/options won't use
                        // the resulting avatars map.
                        let normalizedMessages: ExtractedMessage[];
                        if (needAvatars) {
                            const messagesForAvatar = messages.map(m => ({
                                id: m.id || '',
                                threadId: m.threadId ?? null,
                                author: m.author || '',
                                timestamp: m.timestamp || '',
                                text: m.text || '',
                                edited: m.edited || false,
                                avatar: m.avatar ?? null,
                                ...m,
                            })) as ExtractedMessage[];
                            const result = await embedAvatarsInContent(messagesForAvatar);
                            normalizedMessages = result.messages;
                            avatars = result.avatars;
                        } else {
                            // Drop avatar bytes so nothing travels through
                            // the port unnecessarily. The builder's
                            // removeAvatars() would strip these downstream
                            // anyway when embedAvatars is false.
                            normalizedMessages = messages.map(m => ({ ...m, avatar: null })) as ExtractedMessage[];
                            avatars = {};
                        }

                        // reactor.uuid is a transient content-script aid for the
                        // avatar join; strip it here (covers BOTH branches above)
                        // so user UUIDs never reach the serialized export.
                        for (const m of normalizedMessages) {
                            const reactions = (m as { reactions?: Array<{ reactors?: Array<{ uuid?: string }> }> }).reactions;
                            if (Array.isArray(reactions)) for (const r of reactions) {
                                if (Array.isArray(r.reactors)) for (const rt of r.reactors) delete rt.uuid;
                            }
                        }

                        currentRunStartedAt = null;

                        // If a stop arrived during the avatar/image phase, drop
                        // everything we collected and report cancellation. We
                        // explicitly null the working arrays so the GC can reclaim
                        // potentially large data URLs and message buffers right
                        // away instead of waiting for closure teardown. Check
                        // BEFORE the streaming-prep work so we don't log a
                        // misleading "Streaming N messages" line on a cancelled run.
                        if (abortSignal.aborted) {
                            messages = null;
                            avatars = {};
                            currentRunStartedAt = null;
                            currentAbortController = null;
                            console.log('[Teams Exporter] Scrape cancelled — discarded collected data');
                            sendResponse({ ok: false, cancelled: true });
                            return;
                        }

                        // Include the resolved conversation id so the
                        // background can pin the persisted outcome snapshot
                        // to this specific chat (Teams is an SPA; tabId
                        // alone doesn't distinguish conversations).
                        let scrapeConvId: string | null = null;
                        try { scrapeConvId = await extractConversationId(); } catch { /* best-effort */ }
                        const meta: Record<string, unknown> = {
                            count: messages.length,
                            title,
                            startAt: startAt || null,
                            endAt: endAt || null,
                            avatars,
                        };
                        if (scrapeConvId) meta.conversationId = scrapeConvId;
                        // Roster (API mode only; DOM-scroll fallback can't fetch it).
                        if (participants.length) meta.participants = participants;
                        if (typeof memberCount === 'number') meta.memberCount = memberCount;
                        // Propagate the partial signal. Only flagged when we
                        // saw a high-confidence partial-condition AND we have
                        // SOME messages — empty exports aren't "partial",
                        // they're just empty and the bundle's NO_HISTORY.txt
                        // already covers that case.
                        if (partialReason && messages.length > 0) {
                            meta.partial = { reason: partialReason };
                            console.warn(`[Teams Exporter] Export flagged as partial (${partialReason}); meta.partial set on streamed result`);
                        }
                        const requestId = msg.requestId || `${Date.now()}`;

                        // Estimate payload for logging
                        let attDataUrlBytes = 0;
                        let attCount = 0;
                        for (const m of normalizedMessages) {
                            if (m.attachments) {
                                for (const att of m.attachments) {
                                    if (att.dataUrl) { attDataUrlBytes += att.dataUrl.length; attCount++; }
                                }
                            }
                        }
                        const avatarCount = Object.keys(avatars).length;
                        console.log(`[Teams Exporter] Streaming ${normalizedMessages.length} messages, ${avatarCount} avatars, ${attCount} attachment dataUrls (~${(attDataUrlBytes / 1024 / 1024).toFixed(1)}MB)`);

                        // Signal background that results will be streamed via port
                        // (avoids Chrome's 64MiB sendResponse limit)
                        sendResponse({ ok: true, streaming: true });

                        // Stream results via port in chunks
                        const port = runtime.connect({ name: `scrape-result:${requestId}` });
                        try {
                            port.postMessage({ type: 'meta', meta });
                            // Byte-aware batching. A single chrome.runtime port
                            // message is capped at 64 MiB (JSON-serialized), and
                            // full-res image dataUrls can push a fixed-count batch
                            // past it, which would drop the entire chat. Flush by
                            // accumulated dataUrl bytes (with a message-count
                            // ceiling) so no port message approaches the cap; the
                            // background receiver appends batches in order, so any
                            // batch size is fine.
                            const FLUSH_BYTES = 32 * 1024 * 1024;      // batch target, well under the 64 MiB cap
                            const PEEL_ABOVE_BYTES = 48 * 1024 * 1024; // a lone message over this peels its biggest images into standalone chunks
                            const SINGLE_CHUNK_MAX = 60 * 1024 * 1024; // a single image bigger than this can't fit even its own port message -> drop it
                            const MAX_BATCH_MSGS = 100;
                            // Estimate a message's serialized port-message weight: image
                            // dataUrls (the dominant term), any data: URL parked in href
                            // (the DOM-scrape link-preview path), plus the big text fields
                            // so the batch byte budget reflects real size, not just images.
                            const msgBytesOf = (m: { attachments?: Array<{ dataUrl?: string; href?: string }>; contentHtml?: string; text?: string }): number => {
                                let n = (m.contentHtml?.length || 0) + (m.text?.length || 0);
                                if (m.attachments) for (const a of m.attachments) {
                                    if (a.dataUrl) n += a.dataUrl.length;
                                    else if (a.href && a.href.startsWith('data:')) n += a.href.length;
                                }
                                return n;
                            };
                            let batch: typeof normalizedMessages = [];
                            let batchBytes = 0;
                            let batchNo = 0;
                            const flushBatch = () => {
                                if (!batch.length) return;
                                batchNo++;
                                port.postMessage({ type: 'messages', messages: batch });
                                console.debug(`[Teams Exporter] Sent batch ${batchNo}: ${batch.length} messages (~${(batchBytes / 1048576).toFixed(1)}MB)`);
                                batch = [];
                                batchBytes = 0;
                            };
                            // A single message whose own images exceed the cap can
                            // never fit in one port message. Peel its largest image
                            // dataUrls out and stream them as standalone chunks (one
                            // image is well under 64 MiB); the background reattaches
                            // each by message + attachment index, so nothing is lost.
                            const chunks: Array<{ mi: number; ai: number; dataUrl: string }> = [];
                            let mi = -1;
                            for (const m of normalizedMessages) {
                                mi++;
                                if (abortSignal.aborted) {
                                    console.log('[Teams Exporter] Streaming aborted mid-flight');
                                    break;
                                }
                                // A single image too large to fit even its own port
                                // message (>64 MiB) can't be chunked further, so drop
                                // it to a placeholder so it can never fail the stream.
                                // Belt for any uncapped path (e.g. the DOM-scrape
                                // canvas, whose data: URL can land in dataUrl OR href);
                                // the API/fetch image paths are byte-capped.
                                if (m.attachments) {
                                    for (const a of m.attachments) {
                                        if (a.dataUrl && a.dataUrl.length > SINGLE_CHUNK_MAX) {
                                            console.warn(`[Teams Exporter] One image is too large to stream even alone (~${Math.round(a.dataUrl.length / 1048576)}MB); dropped to placeholder`);
                                            delete a.dataUrl;
                                            a.failReason = a.failReason || 'too-large';
                                        } else if (a.href && a.href.startsWith('data:') && a.href.length > SINGLE_CHUNK_MAX) {
                                            console.warn(`[Teams Exporter] One image is too large to stream even alone (~${Math.round(a.href.length / 1048576)}MB); dropped to placeholder`);
                                            a.href = undefined;
                                            a.failReason = a.failReason || 'too-large';
                                        }
                                    }
                                }
                                let mBytes = msgBytesOf(m);
                                if (mBytes > PEEL_ABOVE_BYTES && m.attachments) {
                                    const heavy = m.attachments
                                        .map((a, ai) => ({ a, ai }))
                                        .filter(x => !!x.a.dataUrl)
                                        .sort((x, y) => y.a.dataUrl!.length - x.a.dataUrl!.length);
                                    let peeled = 0;
                                    for (const { a, ai } of heavy) {
                                        if (mBytes <= PEEL_ABOVE_BYTES) break;
                                        chunks.push({ mi, ai, dataUrl: a.dataUrl! });
                                        mBytes -= a.dataUrl!.length;
                                        delete a.dataUrl;
                                        peeled++;
                                    }
                                    console.debug(`[Teams Exporter] Peeled ${peeled} oversized image(s) from one message into standalone chunks`);
                                }
                                if (batch.length && (batchBytes + mBytes > FLUSH_BYTES || batch.length >= MAX_BATCH_MSGS)) {
                                    flushBatch();
                                }
                                batch.push(m);
                                batchBytes += mBytes;
                            }
                            if (!abortSignal.aborted) {
                                flushBatch();
                                // Send peeled chunks AFTER all message batches so the
                                // background already has every message to reattach to.
                                for (const c of chunks) {
                                    port.postMessage({ type: 'attachment-chunk', mi: c.mi, ai: c.ai, dataUrl: c.dataUrl });
                                }
                                if (chunks.length) console.debug(`[Teams Exporter] Sent ${chunks.length} oversized image chunk(s) separately`);
                            }
                            if (abortSignal.aborted) {
                                port.postMessage({ type: 'error', error: 'cancelled' });
                            } else {
                                port.postMessage({ type: 'done' });
                                console.debug('[Teams Exporter] Streaming complete');
                            }
                        } catch (streamErr: any) {
                            console.error('[Teams Exporter] Streaming error:', streamErr);
                            try { port.postMessage({ type: 'error', error: streamErr?.message || String(streamErr) }); } catch (_) { /* port may be dead */ }
                        } finally {
                            port.disconnect();
                        }
                        currentAbortController = null;
                    }
                } catch (e: any) {
                    // Abort is an expected error path — surface it as cancellation
                    // rather than a regular failure so the popup doesn't show a
                    // red banner.
                    if (currentAbortController?.signal.aborted) {
                        currentRunStartedAt = null;
                        currentAbortController = null;
                        console.log('[Teams Exporter] Scrape cancelled');
                        sendResponse({ ok: false, cancelled: true });
                        return;
                    }
                    currentAbortController = null;
                    console.error('[Teams Exporter] Error:', e);
                    currentRunStartedAt = null;
                    sendResponse({ error: e?.message || String(e) });
                }
            })();
            return true;
        });

    } // End of main()
}); // End of defineContentScript
