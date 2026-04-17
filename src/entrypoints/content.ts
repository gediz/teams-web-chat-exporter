/* eslint-disable no-console */
import { defineContentScript } from 'wxt/sandbox';
import { $, $$ } from '../utils/dom';
import { makeDayDivider as buildDayDivider } from '../utils/messages';
import { cssEscape, isPlaceholderText, textFrom } from '../utils/text';
import { formatElapsed, parseTimeStamp } from '../utils/time';
import { extractAttachments } from '../content/attachments';
import { extractReactions } from '../content/reactions';
import { extractReplyContext } from '../content/replies';
import { extractTables, extractTextWithEmojis, normalizeMentions } from '../content/text';
import { autoScrollAggregate as autoScrollAggregateHelper } from '../content/scroll';
import { extractChatTitle, extractChannelTitle } from '../content/title';
import { extractAvatarId } from '../utils/avatars';
import { TEAMS_MATCH_PATTERNS } from '../utils/teams-urls';
import { apiScrape, getGraphToken } from '../content/api-client';
import { convertApiMessages, buildApiMeta } from '../content/api-converter';
import type { AggregatedItem, Attachment, ExportMessage, OrderContext, Reaction, ReplyContext, ScrapeOptions } from '../types/shared';

// Typed globals for Firefox builds
declare const browser: typeof chrome | undefined;

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

        let hudEnabled = true;
        let currentRunStartedAt: number | null = null;

        function isChatNavSelected() {
            return Boolean(document.querySelector('[data-tid="app-bar-wrapper"] button[aria-pressed="true"][aria-label^="Chat" i]'));
        }

        function isTeamsNavSelected() {
            return Boolean(document.querySelector('[data-tid="app-bar-wrapper"] button[aria-pressed="true"][aria-label*="Teams" i]'));
        }

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
                const navSelected = isTeamsNavSelected();
                const hasSurface = hasChannelMessageSurface();

                if (!navSelected) {
                    return { ok: false, reason: 'Switch to the Teams app in Teams before exporting.' };
                }

                if (hasSurface) {
                    return { ok: true };
                }

                return { ok: false, reason: 'Open a team channel before exporting.' };
            }

            const navSelected = isChatNavSelected();
            const hasSurface = hasChatMessageSurface();

            if (navSelected && hasSurface) {
                return { ok: true };
            }

            if (!navSelected) {
                return { ok: false, reason: 'Switch to the Chat app in Teams before exporting.' };
            }

            return { ok: false, reason: 'Open a chat conversation before exporting.' };
        }

        function clearHUD() {
            const existing = document.getElementById("__teamsExporterHUD");
            if (existing) existing.remove();
        }

        // HUD -----------------------------------------------------------
        function ensureHUD() {
            if (!hudEnabled) return null;
            let hud = document.getElementById("__teamsExporterHUD");
            if (!hud) {
                hud = document.createElement("div");
                hud.id = "__teamsExporterHUD";
                hud.style.cssText = "position:fixed;right:12px;top:12px;z-index:999999;font:12px/1.3 system-ui,sans-serif;color:#111;background:rgba(255,255,255,.92);border:1px solid #ddd;border-radius:8px;padding:8px 10px;box-shadow:0 2px 8px rgba(0,0,0,.15);pointer-events:none;";
                hud.textContent = "Teams Exporter: idle";
                document.body.appendChild(hud);
            }
            return hud;
        }
        function hud(text: string, { includeElapsed = true }: { includeElapsed?: boolean } = {}) {
            if (!hudEnabled) return;
            const hudNode = ensureHUD();
            if (hudNode) {
                let final = `Teams Exporter: ${text}`;
                if (includeElapsed !== false && currentRunStartedAt) {
                    final += ` • elapsed ${formatElapsed(Date.now() - currentRunStartedAt)}`;
                }
                hudNode.textContent = final;
            }
            try {
                const msgPromise = runtime.sendMessage({ type: "SCRAPE_PROGRESS", payload: { phase: "hud", text } });
                if (msgPromise && msgPromise.catch) msgPromise.catch(() => { });
            } catch (e) { /* ignore */ }
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
                    console.warn(`[Avatar Fetch] HTTP ${res.status} for ${url.substring(0, 100)}...`);
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

            return { messages: normalized, avatars };
        }

        // ── Inline Image Fetching (API mode) ──────────────────────
        const MAX_IMAGE_BYTES = 5 * 1024 * 1024; // 5 MB
        const MAX_CONCURRENT_FETCHES = 6;

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

        async function fetchImageAsDataUrl(url: string): Promise<string | null> {
            // All image fetches go through Teams' authenticated URL-image proxy:
            //   - AMS object URLs and asyncgw proxy URLs get rewritten with {userId}
            //   - Any other http(s) URL gets wrapped in /urlp/v1/{userId}/url/image/Thumbnail
            // The Teams server validates + fetches the upstream, so we never touch
            // third-party CDNs directly (no extra host_permissions, no CORS pitfalls).
            if (!imgFetchAuth?.userId || !imgFetchAuth?.userRegion || !imgFetchAuth?.ic3Token) {
                imgFetchStats.skippedDomain++;
                return null;
            }
            const fetchUrl = transformImageUrlToProxy(url, imgFetchAuth.userId, imgFetchAuth.userRegion);
            if (!fetchUrl) {
                imgFetchStats.skippedDomain++;
                return null;
            }
            try {
                const resp = await runtime.sendMessage({
                    type: 'FETCH_BLOB',
                    url: fetchUrl,
                    bearerToken: imgFetchAuth.ic3Token,
                    maxBytes: MAX_IMAGE_BYTES,
                    minBytes: 100,
                }) as { ok: boolean; dataUrl?: string; status?: number; statusText?: string; error?: string; sizeReason?: number };
                if (resp?.ok && resp.dataUrl) return resp.dataUrl;
                if (resp?.status) {
                    imgFetchStats.httpError++;
                    if (!imgFetchStats.firstHttpError) imgFetchStats.firstHttpError = `HTTP ${resp.status} ${resp.statusText || ''} for ${url.slice(0, 80)}`;
                } else if (resp?.error) {
                    imgFetchStats.threwError++;
                    if (!imgFetchStats.firstThrow) imgFetchStats.firstThrow = `${resp.error.slice(0, 100)} for ${url.slice(0, 80)}`;
                } else if (typeof resp?.sizeReason === 'number') {
                    if (resp.sizeReason > MAX_IMAGE_BYTES) imgFetchStats.tooLarge++;
                    else imgFetchStats.tooSmall++;
                }
                return null;
            } catch (e) {
                imgFetchStats.threwError++;
                if (!imgFetchStats.firstThrow) imgFetchStats.firstThrow = `bg-fetch ${String(e).slice(0, 80)}`;
                return null;
            }
        }

        // Per-export image fetch counters; reset at the start of each fetchInlineImages call.
        const imgFetchStats = {
            skippedDomain: 0, httpError: 0, threwError: 0, tooLarge: 0, tooSmall: 0,
            firstHttpError: '' as string,
            firstThrow: '' as string,
        };
        const resetImgFetchStats = () => {
            imgFetchStats.skippedDomain = 0;
            imgFetchStats.httpError = 0;
            imgFetchStats.threwError = 0;
            imgFetchStats.tooLarge = 0;
            imgFetchStats.tooSmall = 0;
            imgFetchStats.firstHttpError = '';
            imgFetchStats.firstThrow = '';
        };

        /**
         * Fetch inline images for API-mode messages.
         * Finds AMS image URLs in attachments and downloads them as data URLs.
         */
        async function fetchInlineImages(
            messages: ExportMessage[],
            onProgress?: (done: number, total: number) => void,
            auth?: { userId: string | null; userRegion: string; ic3Token: string },
        ): Promise<void> {
            resetImgFetchStats();
            imgFetchAuth = (auth?.userId && auth.userRegion && auth.ic3Token)
                ? { userId: auth.userId, userRegion: auth.userRegion, ic3Token: auth.ic3Token }
                : null;
            if (!imgFetchAuth) {
                console.warn('[Teams Exporter] Image proxy auth missing — fetches will use raw URLs. ' +
                    `userId=${auth?.userId ? 'ok' : 'null'} region=${auth?.userRegion ? 'ok' : 'null'} ic3=${auth?.ic3Token ? 'ok' : 'null'}`);
            }
            // Collect attachments whose href is an actual image (or routable through
            // the Teams URL-image proxy). Skip file attachments like SharePoint docs,
            // which have an http href but the proxy returns 415 for them.
            const tasks: { att: { href?: string; dataUrl?: string }; url: string }[] = [];
            for (const m of messages) {
                if (!m.attachments) continue;
                for (const att of m.attachments) {
                    if (!att.href || !/^https?:\/\//i.test(att.href)) continue;
                    const isAmsObject = /\/v1\/objects\/[^/]+\/views\//i.test(att.href);
                    const isUrlp = /\/urlp\//i.test(att.href);
                    const isImageish = att.type === 'gif' || att.type === 'video' || att.kind === 'preview';
                    if (!isAmsObject && !isUrlp && !isImageish) continue;
                    tasks.push({ att, url: att.href });
                }
            }

            if (!tasks.length) return;
            console.log(`[Teams Exporter] Fetching ${tasks.length} inline images…`);

            let done = 0;
            let succeeded = 0;
            // Process in batches to limit concurrency
            for (let i = 0; i < tasks.length; i += MAX_CONCURRENT_FETCHES) {
                const batch = tasks.slice(i, i + MAX_CONCURRENT_FETCHES);
                await Promise.all(
                    batch.map(async ({ att, url }) => {
                        const dataUrl = await fetchImageAsDataUrl(url);
                        if (dataUrl) {
                            att.dataUrl = dataUrl;
                            succeeded++;
                        } else if (/\/urlp\/v1\/url\/image\//i.test(url)) {
                            // Thumbnail proxy failed (CORS/opaque) — clear href to avoid broken img
                            att.href = undefined;
                        }
                        done++;
                        onProgress?.(done, tasks.length);
                    }),
                );
            }
            console.log(`[Teams Exporter] Image fetch: ${succeeded} succeeded, ${tasks.length - succeeded} failed (of ${tasks.length} attempted)`);
            const failures = imgFetchStats.httpError + imgFetchStats.threwError + imgFetchStats.tooLarge + imgFetchStats.tooSmall + imgFetchStats.skippedDomain;
            if (failures > 0) {
                const detail = [
                    imgFetchStats.httpError && `${imgFetchStats.httpError} http-error`,
                    imgFetchStats.threwError && `${imgFetchStats.threwError} threw`,
                    imgFetchStats.tooLarge && `${imgFetchStats.tooLarge} too-large`,
                    imgFetchStats.tooSmall && `${imgFetchStats.tooSmall} too-small`,
                    imgFetchStats.skippedDomain && `${imgFetchStats.skippedDomain} domain-blocked`,
                ].filter(Boolean).join(', ');
                console.warn(`[Teams Exporter] Image fetch failures — ${detail}`);
                if (imgFetchStats.firstHttpError) console.warn(`[Teams Exporter] First http error: ${imgFetchStats.firstHttpError}`);
                if (imgFetchStats.firstThrow) console.warn(`[Teams Exporter] First exception: ${imgFetchStats.firstThrow}`);
            }
        }

        // ── Avatar Fetching (API mode) ────────────────────────────
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
            const uuidToDataUrl = new Map<string, string>();
            let firstPhotoError: string | null = null;
            let photo404 = 0;
            for (const [, uuid] of authorToUuid) {
                if (uuidToDataUrl.has(uuid)) continue;
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
                    } else if (resp?.status === 404) {
                        photo404++;
                    } else if (!firstPhotoError) {
                        firstPhotoError = resp?.status
                            ? `HTTP ${resp.status} ${resp.statusText || ''}`
                            : (resp?.error || 'unknown error');
                    }
                } catch (e) {
                    if (!firstPhotoError) firstPhotoError = String(e);
                }
            }
            const missing = authorToUuid.size - uuidToDataUrl.size;
            if (missing > 0) {
                const nonPhotoErrors = missing - photo404;
                const parts: string[] = [];
                if (photo404 > 0) parts.push(`${photo404} have no photo (404)`);
                if (nonPhotoErrors > 0) parts.push(`${nonPhotoErrors} failed (${firstPhotoError || 'no error captured'})`);
                const level = nonPhotoErrors > 0 ? 'warn' : 'log';
                console[level](`[Teams Exporter] Graph photo: ${uuidToDataUrl.size}/${authorToUuid.size} fetched — ${parts.join(', ')}`);
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

        const extractCodeBlock = (el: Element) => {
            let code = '';
            const walkCode = (n: ChildNode) => {
                if (n.nodeType === Node.TEXT_NODE) { code += n.nodeValue; return; }
                if (n.nodeType !== Node.ELEMENT_NODE) return;
                const child = n as Element;
                const tagName = child.tagName;
                if (tagName === 'BR') { code += '\n'; return; }
                if (tagName === 'IMG') { code += (child.getAttribute('alt') || child.getAttribute('aria-label') || ''); return; }
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
                out += (el.getAttribute("alt") || el.getAttribute("aria-label") || "");
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
            try { btn.click(); } catch {}
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
                    try { btn.click(); } catch {}
                    await sleep(200);
                    break;
                }
            }
            try {
                document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', code: 'Escape', keyCode: 27, which: 27, bubbles: true }));
            } catch {}
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
        try { el.scrollIntoView({ block: "center" }); } catch {}
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
        const controlTimeRe = /(\d{1,2})\/(\d{1,2})\s+(\d{1,2}):(\d{2})\s*(AM|PM)/i;

        function parseControlTimestamp(text: string, yearHint?: number | null): number | null {
            if (!text) return null;
            const match = controlTimeRe.exec(text);
            if (!match) return null;
            const month = Number(match[1]);
            const day = Number(match[2]);
            let hour = Number(match[3]);
            const minute = Number(match[4]);
            const period = match[5];
            if (!Number.isFinite(month) || !Number.isFinite(day) || !Number.isFinite(hour) || !Number.isFinite(minute)) return null;
            if (period?.toUpperCase() === 'PM' && hour < 12) hour += 12;
            if (period?.toUpperCase() === 'AM' && hour === 12) hour = 0;
            const baseYear = typeof yearHint === 'number' ? yearHint : new Date().getFullYear();
            const date = new Date(baseYear, month - 1, day, hour, minute, 0, 0);
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

            let wrapperWithMid: HTMLElement | null = null;
            let wrapperItem: HTMLElement | null = null;
            let body: HTMLElement | null = null;
            let itemScope: Element = item;
            let hasMessage = false;

            if (isReplyItem) {
                // In the replies runway, `item` is already the message container.
                // Do NOT climb up with `closest`, or you risk grabbing the parent post.
                wrapperWithMid = item.querySelector<HTMLElement>('[data-mid]') || null;
                wrapperItem = item;
                body = item;
                itemScope = item;
                hasMessage = true;
            } else {
                // --- Original non-reply initialization path -------------------
                wrapperWithMid =
                    item.querySelector<HTMLElement>('[data-testid="message-body-flex-wrapper"][data-mid]') ||
                    item.querySelector<HTMLElement>('[data-tid="channel-replies-pane-message"] [data-mid]') ||
                    item.querySelector<HTMLElement>('[data-mid]');

                wrapperItem = wrapperWithMid
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
            let contentEl: Element =
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

            // Timestamp fallback from mid (some mids are ms since epoch)
            if ((!ts || Number.isNaN(tms)) && mid) {
                const midMs = Number(mid);
                if (Number.isFinite(midMs) && midMs > 100000000000) {
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
                  hud,
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

        if (!isTop) return;
        // Bridge --------------------------------------------------------
        runtime.onMessage.addListener((msg, _sender, sendResponse) => {
            (async () => {
                try {
                    if (msg.type === 'PING') { sendResponse({ ok: true }); return; }
                    if (msg.type === 'CHECK_CHAT_CONTEXT') { sendResponse(checkChatContext(msg.target)); return; }
                    if (msg.type === 'SCRAPE_TEAMS') {
                        const { startAt, endAt, includeReactions, includeSystem, includeReplies, showHud, exportTarget } = msg.options || {};
                        const target = exportTarget === 'team' ? 'team' : 'chat';
                        hudEnabled = showHud !== false;
                        if (!hudEnabled) clearHUD();
                        const scrapeOpts = { startAtISO: startAt, endAtISO: endAt, includeSystem, includeReactions, includeReplies: includeReplies !== false };
                        console.debug('[Teams Exporter] SCRAPE_TEAMS', location.href, msg.options);
                        currentRunStartedAt = Date.now();
                        hud('starting…');

                        // ── API Mode (fast, no scrolling) ──────────────────
                        let messages: ExportMessage[] | null = null;
                        let title: string | null = null;
                        let avatars: Record<string, string> = {};

                        try {
                            hud('trying API mode…');
                            const apiResult = await apiScrape((p) => {
                                if (p.phase === 'discover') {
                                    hud('discovering API endpoint…');
                                } else {
                                    hud(`API: page ${p.page}, ${p.messagesSoFar} messages`);
                                }
                                try {
                                    const mp = runtime.sendMessage({ type: 'SCRAPE_PROGRESS', payload: { phase: 'api-fetch', messagesVisible: p.messagesSoFar, passes: p.page } });
                                    if (mp && mp.catch) mp.catch(() => {});
                                } catch { /* ignore */ }
                            }, { startAtISO: scrapeOpts.startAtISO });

                            if (apiResult) {
                                const converted = convertApiMessages(apiResult.messages, scrapeOpts);
                                if (converted.length > 0) {
                                    // Fetch inline images (content script has auth cookies, URLs expire)
                                    hud(`fetching inline images…`);
                                    try { runtime.sendMessage({ type: 'SCRAPE_PROGRESS', payload: { phase: 'images', messagesVisible: converted.length } }); } catch {}
                                    await fetchInlineImages(converted, (done, total) => {
                                        hud(`images: ${done}/${total}`);
                                        if (done % 10 === 0 || done === total) {
                                            try { runtime.sendMessage({ type: 'SCRAPE_PROGRESS', payload: { phase: 'images', messagesVisible: converted.length, imagesDone: done, imagesTotal: total } }); } catch {}
                                        }
                                    }, { userId: apiResult.userId, userRegion: apiResult.userRegion, ic3Token: apiResult.ic3Token });
                                    // Fetch profile photos via Graph API
                                    hud(`fetching avatars…`);
                                    try { runtime.sendMessage({ type: 'SCRAPE_PROGRESS', payload: { phase: 'avatars', messagesVisible: converted.length } }); } catch {}
                                    await fetchApiAvatars(converted, apiResult.messages);
                                    messages = converted;
                                    title = target === 'team' ? extractChannelTitle() : extractChatTitle();
                                    console.log(`[Teams Exporter] API mode: ${converted.length} messages from ${apiResult.messages.length} raw`);
                                    hud(`API: ${converted.length} messages`);
                                } else {
                                    console.log('[Teams Exporter] API returned messages but all filtered out, trying DOM scroll');
                                }
                            }
                        } catch (apiErr) {
                            console.log('[Teams Exporter] API mode failed, falling back to DOM scroll:', apiErr);
                        }

                        // ── DOM Scroll Fallback ────────────────────────────
                        if (!messages) {
                            hud('scrolling…');
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
                                    hud,
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
                                currentRunStartedAt
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
                        hud(`extracted ${messages.length} messages`);

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

                        // Fetch avatars in content script context (has access to Teams cookies)
                        hud('fetching avatars...');
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
                        const { messages: normalizedMessages, avatars: fetchedAvatars } = await embedAvatarsInContent(messagesForAvatar);
                        avatars = fetchedAvatars;

                        currentRunStartedAt = null;

                        const meta = { count: messages.length, title, startAt: startAt || null, endAt: endAt || null, avatars };
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
                            const BATCH = 100;
                            for (let i = 0; i < normalizedMessages.length; i += BATCH) {
                                const batch = normalizedMessages.slice(i, i + BATCH);
                                port.postMessage({ type: 'messages', messages: batch });
                                console.debug(`[Teams Exporter] Sent batch ${Math.floor(i / BATCH) + 1}: messages ${i + 1}-${i + batch.length}`);
                            }
                            port.postMessage({ type: 'done' });
                            console.debug('[Teams Exporter] Streaming complete');
                        } catch (streamErr: any) {
                            console.error('[Teams Exporter] Streaming error:', streamErr);
                            try { port.postMessage({ type: 'error', error: streamErr?.message || String(streamErr) }); } catch (_) { /* port may be dead */ }
                        } finally {
                            port.disconnect();
                        }
                    }
                } catch (e: any) {
                    console.error('[Teams Exporter] Error:', e);
                    hud(`error: ${e?.message || e}`);
                    currentRunStartedAt = null;
                    sendResponse({ error: e?.message || String(e) });
                }
            })();
            return true;
        });

    } // End of main()
}); // End of defineContentScript
