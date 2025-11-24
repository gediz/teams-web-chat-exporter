/* eslint-disable no-console */
import { defineContentScript } from 'wxt/sandbox';
import { $, $$ } from '../utils/dom';
import { makeDayDivider as buildDayDivider } from '../utils/messages';
import { cssEscape, isPlaceholderText, preferText, textFrom } from '../utils/text';
import { formatElapsed, parseTimeStamp, startOfLocalDay } from '../utils/time';
import { extractAttachments } from '../content/attachments';
import { extractReactions } from '../content/reactions';
import { extractReplyContext } from '../content/replies';
import { extractTextWithEmojis, normalizeMentions } from '../content/text';
import type { AggregatedItem, Attachment, ExportMessage, OrderContext, Reaction, ReplyContext, ScrapeOptions } from '../types/shared';

// Typed globals for Firefox builds
declare const browser: typeof chrome | undefined;

type ExtractedMessage = ExportMessage & {
    id: string;
    author: string;
    timestamp: string;
    text: string;
    edited: boolean;
    avatar: string | null;
};

type ContentAggregated = AggregatedItem & { message?: ExtractedMessage };
export default defineContentScript({
  matches: [
    'https://*.teams.microsoft.com/*',
    'https://teams.cloud.microsoft/*',
  ],
  runAt: 'document_idle',
  allFrames: true,

  main() {
// Browser API compatibility for Firefox
const runtime = typeof browser !== 'undefined' ? browser.runtime : chrome.runtime;

let hudEnabled = true;
let currentRunStartedAt: number | null = null;

function isChatNavSelected() {
    return Boolean(document.querySelector('[data-tid="app-bar-wrapper"] button[aria-pressed="true"][aria-label^="Chat" i]'));
}

function hasChatMessageSurface() {
    return Boolean(
        document.querySelector('[data-tid="message-pane-list-viewport"], [data-tid="chat-message-list"], [data-tid="chat-pane"]')
    );
}

function checkChatContext() {
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
function getScroller() {
    return $('[data-tid="message-pane-list-viewport"]') || $('[data-tid="chat-message-list"]') || document.scrollingElement;
}

// Author/timestamp/edited/avatar helpers ------------------------
function resolveAuthor(body: Element, lastAuthor = ""): string {
    let author = textFrom($('[data-tid="message-author-name"]', body));
    if (!author) {
        const aria = body.getAttribute('aria-labelledby') || '';
        const aId = aria.split(/\s+/).find(s => s.startsWith('author-'));
        if (aId) author = textFrom(document.getElementById(aId));
    }
    return author || lastAuthor || '';
}
function resolveTimestamp(item: Element): string {
    const t = $('time[datetime]', item) || $('time', item) || $('[data-tid="message-status"] time', item);
    return t?.getAttribute?.('datetime') || t?.getAttribute?.('title') || textFrom(t) || '';
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
function resolveAvatar(item: Element): string | null {
    const perMsg = $('[data-tid="message-avatar"] img', item) as HTMLImageElement | null; // per-message avatar
    if (perMsg?.src) return perMsg.src;
    const header = $('[data-tid="chat-title-avatar"] img') as HTMLImageElement | null;   // header fallback
    return header?.src || null;
}

// Text with emoji (IMG alt) + block breaks
function extractTextWithEmojis(root: Element | null): string {
    if (!root) return '';
    let out = '';
    const walk = (n: ChildNode) => {
        if (n.nodeType === Node.TEXT_NODE) { out += n.nodeValue; return; }
        if (n.nodeType !== Node.ELEMENT_NODE) return;
        const el = n as Element;
        const tag = el.tagName;
        if (tag === 'BR') { out += '\n'; return; }
        if (tag === 'IMG') { out += (el.getAttribute('alt') || el.getAttribute('aria-label') || ''); return; }
        const blockish = /^(DIV|P|LI|BLOCKQUOTE)$/;
        const start = out.length;
        for (const c of el.childNodes) walk(c);
        if (blockish.test(tag) && out.length > start) out += '\n';
    };
    walk(root);
    return out.replace(/\n{3,}/g, '\n\n').trim();
}


// Helpers -------------------------------------------------------
const sleep = (ms: number) => new Promise(resolve => setTimeout(resolve, ms));

const findPaneItemByMessageId = (id: string | null | undefined): Element | null => {
    if (!id) return null;
    const msgNode = document.querySelector(`[data-mid="${cssEscape(id)}"]`);
    return msgNode?.closest('[data-tid="chat-pane-item"]') || null;
};

async function hydrateSparseMessages(agg: Map<string, ContentAggregated>, opts: ScrapeOptions = {}) {
    if (!agg || agg.size === 0) return;

    const needsHydration = (message: ExtractedMessage, item: Element) => {
        const textNeeds = isPlaceholderText(message.text);
        let reactionsNeed = false;
        if (opts.includeReactions) {
            const hadReactions = Array.isArray(message.reactions) && message.reactions.length > 0;
            if (!hadReactions && item?.querySelector('[data-tid="diverse-reaction-pill-button"]')) {
                reactionsNeed = true;
            }
        }
        return { textNeeds, reactionsNeed, needs: textNeeds || reactionsNeed };
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
        const nextPending = [];

        for (const task of pending) {
            const { id } = task;
            const existing = agg.get(id);
            if (!existing) continue;

            if (!existing.message) continue;
            const item = findPaneItemByMessageId(id) || task.item;
            if (!item) continue;

            const lastAuthorRef = { value: existing.message.author || '' };
            const ts = existing.message.timestamp ? Date.parse(existing.message.timestamp) : undefined;
            const tempOrderCtx: OrderContext = {
                lastTimeMs: Number.isNaN(ts) ? null : ts ?? null,
                yearHint: Number.isNaN(ts) ? null : (ts ? new Date(ts).getFullYear() : null),
                seqBase: Date.now(),
                seq: 0,
                lastAuthor: existing.message.author || '',
                lastId: existing.message.id || null,
                systemCursor: 0
            };

            const reExtracted = await extractOne(item, { includeSystem: opts.includeSystem, includeReactions: opts.includeReactions }, lastAuthorRef, tempOrderCtx);
            if (!reExtracted?.message) {
                nextPending.push(task);
                continue;
            }

            const merged: ExtractedMessage = {
                id: existing.message.id || reExtracted.message.id || id,
                author: existing.message.author || reExtracted.message.author || '',
                timestamp: existing.message.timestamp || reExtracted.message.timestamp || '',
                text: preferText(existing.message.text || '', reExtracted.message.text || ''),
                edited: existing.message.edited || reExtracted.message.edited,
                system: existing.message.system || reExtracted.message.system,
                avatar: existing.message.avatar ?? reExtracted.message.avatar ?? null,
                reactions: existing.message.reactions,
                attachments: existing.message.attachments,
                replyTo: existing.message.replyTo ?? reExtracted.message.replyTo ?? null,
            };

            if (opts.includeReactions) {
                const newReacts = reExtracted.message.reactions || [];
                const prevCount = Array.isArray(merged.reactions) ? merged.reactions.length : 0;
                if (newReacts.length && newReacts.length >= prevCount) {
                    merged.reactions = newReacts;
                }
            }

            const newAttachments = reExtracted.message.attachments || [];
            const prevAttCount = Array.isArray(merged.attachments) ? merged.attachments.length : 0;
            if (newAttachments.length && newAttachments.length >= prevAttCount) {
                merged.attachments = newAttachments;
            }

            if (!merged.replyTo && reExtracted.message.replyTo) merged.replyTo = reExtracted.message.replyTo;
            if (!merged.avatar && reExtracted.message.avatar) merged.avatar = reExtracted.message.avatar;
            merged.edited = merged.edited || Boolean(reExtracted.message.edited);

            const newTsMs = reExtracted?.tsMs ?? existing.tsMs ?? (merged.timestamp ? parseTimeStamp(merged.timestamp) : null);
            const kind = existing.kind ?? reExtracted?.kind;
            agg.set(id, { message: merged as ExtractedMessage, orderKey: existing.orderKey, tsMs: newTsMs, kind });

            const status = needsHydration(merged, item);
        if (status.needs) nextPending.push({ id, item });
    }

        pending = nextPending;
        attempts++;
    }

    if (pending.length) {
        try {
            console.debug('[Teams Exporter] hydration pending after retries', pending.map(p => p.id));
        } catch (_) {
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

// Extract one item into a message object + an orderKey
async function extractOne(
    item: Element,
    opts: ScrapeOptions,
    lastAuthorRef: { value: string },
    orderCtx: OrderContext & { seq?: number }
): Promise<ContentAggregated | null> {
    const body = $('[data-tid="chat-pane-message"]', item) || item;
    const isSystem = !$('[data-tid="chat-pane-message"]', item);

    // Date/system divider
    if (isSystem) {
        if (!opts.includeSystem) return null;
        const dividerWrapper = $('.fui-Divider__wrapper', item);
        const controlRenderer = $('[data-tid="control-message-renderer"]', item);

        if (dividerWrapper && !controlRenderer) {
            const text = textFrom(dividerWrapper) || 'system';
            const bodyMid = dividerWrapper.getAttribute?.('data-mid') || $('[data-mid]', dividerWrapper)?.getAttribute('data-mid') || item.getAttribute('data-mid') || dividerWrapper.id;
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
        const bodyMid = wrapper?.getAttribute?.('data-mid') || $('[data-mid]', wrapper || item)?.getAttribute('data-mid') || item.getAttribute('data-mid') || wrapper?.id;
        const dividerId = (bodyMid || text || 'system').toLowerCase();
        const numericMid = bodyMid && Number(bodyMid);
        let parsedTs = parseDateDividerText(text, orderCtx.yearHint);
        if (!Number.isFinite(parsedTs)) parsedTs = parseControlTimestamp(text, orderCtx.yearHint);
        const systemCursor = typeof orderCtx.systemCursor === 'number' ? orderCtx.systemCursor : -9e15;
        const approxMs: number = Number.isFinite(parsedTs)
            ? parsedTs!
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
            message: { id: dividerId, author: '[system]', timestamp: '', text, reactions: [], attachments: [], edited: false, avatar: null, replyTo: null, system: true },
            orderKey: approxMs,
            tsMs: approxMs,
            kind: 'system-control'
        };
    }

    // Normal message
    const ts = resolveTimestamp(item);
    const tms = ts ? Date.parse(ts) : NaN;
    if (!Number.isNaN(tms)) {
        orderCtx.lastTimeMs = tms;
        orderCtx.yearHint = new Date(tms).getFullYear();
    }

    const author = resolveAuthor(body, lastAuthorRef.value || orderCtx.lastAuthor || '');
    if (author) {
        lastAuthorRef.value = author;
        orderCtx.lastAuthor = author;
    }

    const contentEl = $('[id^="content-"]', body) || $('[data-tid="message-content"]', body) || body;
    const cleanRoot = stripQuotedPreview(contentEl) || contentEl;
    normalizeMentions(cleanRoot);
    const text = extractTextWithEmojis(cleanRoot);
    const edited = resolveEdited(item, body);
    const avatar = resolveAvatar(item);
    const reactions = opts.includeReactions ? await extractReactions(item) : [];

    const attachments = extractAttachments(item, body);
    const replyTo = opts.includeReplies === false ? null : extractReplyContext(item, body);

    const mid = body.getAttribute('data-mid') || item.id || `${ts}#${author}`;
    const msg: ExtractedMessage = { id: mid, author, timestamp: ts, text, reactions, attachments, edited, avatar, replyTo, system: false };

    const seqVal = orderCtx.seq ?? 0;
    orderCtx.seq = seqVal + 1;
    const orderKey = !Number.isNaN(tms) ? tms : (orderCtx.seqBase + seqVal);
    const tsMs = !Number.isNaN(tms) ? tms : null;
    return { message: msg, orderKey, tsMs, kind: 'message' };
}

// Aggregate while scrolling so virtualization can’t drop items
async function collectCurrentVisible(agg: Map<string, ContentAggregated>, opts: ScrapeOptions, orderCtx: OrderContext) {
    const nodes = $$('[data-tid="chat-pane-item"]'); // preserve DOM order for system dividers, too
    const lastAuthorRef = { value: orderCtx.lastAuthor || '' };
    for (let i = 0; i < nodes.length; i++) {
        const item = nodes[i];
        const idCandidate = $('[data-tid="chat-pane-message"]', item)?.getAttribute('data-mid') || $('[data-tid="control-message-renderer"]', item)?.getAttribute('data-mid') || $('.fui-Divider__wrapper', item)?.id || item.id || `node-${i}`;
        if (agg.has(idCandidate)) continue;

        const extracted = await extractOne(item, opts, lastAuthorRef, orderCtx);
        if (!extracted) continue;
        if (extracted.kind === 'day-divider') {
            if (typeof extracted.tsMs === 'number' && Number.isFinite(extracted.tsMs)) {
                orderCtx.lastTimeMs = extracted.tsMs;
                orderCtx.yearHint = new Date(extracted.tsMs).getFullYear();
            }
            continue;
        }
        const { message, orderKey, tsMs, kind } = extracted;
        if (!message) continue;

        agg.set(message.id || `${orderKey}`, { message, orderKey, tsMs, kind });
        if (!message.system && message.timestamp) {
            const tms = Date.parse(message.timestamp);
            if (!Number.isNaN(tms)) { orderCtx.lastTimeMs = tms; orderCtx.yearHint = new Date(tms).getFullYear(); }
        }
        if (!message.system && message.author) {
            orderCtx.lastAuthor = message.author;
        }
    }
}

async function autoScrollAggregate({ startAtISO, endAtISO, includeSystem, includeReactions, includeReplies = true }: ScrapeOptions & { includeReplies?: boolean }) {
    const scroller = getScroller();
    if (!scroller) throw new Error('Scroller not found');

    const agg = new Map<string, ContentAggregated>();         // id -> {message, orderKey}
    const orderCtx: OrderContext = {
        lastTimeMs: null,
        yearHint: null,
        seqBase: Date.now(),
        seq: 0,
        lastAuthor: '',
        lastId: null,
        systemCursor: -9e15
    };

    // 0) Pre-capture bottom (newest window)
    scroller.scrollTop = scroller.scrollHeight;
    await new Promise(r => requestAnimationFrame(r));
    await sleep(300);
    await collectCurrentVisible(agg, { includeSystem, includeReactions, includeReplies }, orderCtx);

    // 1) Scroll to top repeatedly to load older history, collecting each pass
    let prevHeight = -1;
    let lastCount = -1;
    let passes = 0;
    let stagnantPasses = 0;
    let lastOldestId = null;
    const dwellMs = 700;

    const headerSentinel = document.querySelector('[data-tid="message-pane-header"]');
    let topReached = false;
    const observer = headerSentinel ? new IntersectionObserver((entries) => {
        const entry = entries[0];
        if (entry?.isIntersecting) topReached = true;
    }, { root: scroller, threshold: 0.01 }) : null;
    if (observer && headerSentinel) observer.observe(headerSentinel);

    const startLimit = typeof startAtISO === 'string' ? parseTimeStamp(startAtISO) : null;
    const endLimit = typeof endAtISO === 'string' ? parseTimeStamp(endAtISO) : null;

    try {
        while (true) {
            passes++;
            scroller.scrollTop = 0;
            await new Promise(r => requestAnimationFrame(r));
            await sleep(dwellMs);

            await collectCurrentVisible(agg, { includeSystem, includeReactions, includeReplies }, orderCtx);

            const nodes = $$('[data-tid="chat-pane-item"]');
            if (!nodes.length) break;
            const newCount = nodes.length;
            const newHeight = scroller.scrollHeight;
            const oldestNode = nodes[0];
            const oldestTimeAttr = $('time[datetime]', oldestNode)?.getAttribute('datetime') || null;
            const oldestTime = oldestTimeAttr;
            const oldestTs = parseTimeStamp(oldestTimeAttr);
            const oldestId = $('[data-tid="chat-pane-message"]', oldestNode)?.getAttribute('data-mid') || oldestNode?.id || null;

            // Expand any collapsed sections that block older history
            const hiddenButtons = Array.from(document.querySelectorAll<HTMLButtonElement>('[data-tid="show-hidden-chat-history-btn"]'))
                .filter(btn => btn && !btn.disabled && btn.offsetParent !== null);
            if (hiddenButtons.length) {
                console.debug('[Teams Exporter] expanding hidden history', { count: hiddenButtons.length });
                for (const btn of hiddenButtons) {
                    try { btn.click(); }
                    catch (err) { console.warn('[Teams Exporter] failed to click hidden-history button', err); }
                    await sleep(400);
                }
                // Reassert scroll position after Teams expands hidden blocks (it may jump to bottom)
                scroller.scrollTop = 0;
                await new Promise(r => requestAnimationFrame(r));
                scroller.scrollTop = 0;
                await sleep(300);
                // reset stagnation tracking and continue so new content can render next loop
                stagnantPasses = 0;
                prevHeight = -1;
                lastCount = -1;
                lastOldestId = null;
                await sleep(600);
                continue;
            }

            const elapsedMs = currentRunStartedAt ? Date.now() - currentRunStartedAt : null;
            const seen = agg.size;
            let filteredSeen = 0;
            for (const entry of agg.values()) {
                const candidate = entry?.tsMs ?? (entry?.message?.timestamp ? parseTimeStamp(entry.message.timestamp) : null);
                if (candidate == null) {
                    filteredSeen++;
                    continue;
                }
                if (startLimit != null && candidate < startLimit) continue;
                if (endLimit != null && candidate >= endLimit) continue;
                filteredSeen++;
            }

            try {
                const msgPromise = runtime.sendMessage({ type: 'SCRAPE_PROGRESS', payload: { phase: 'scroll', passes, newHeight, messagesVisible: newCount, aggregated: seen, seen: filteredSeen, filteredSeen, oldestTime, oldestId, elapsedMs } });
                if (msgPromise && msgPromise.catch) msgPromise.catch(() => { });
            } catch (e) { /* ignore */ }
            hud(`scroll pass ${passes} • seen ${filteredSeen}`);
            console.debug('[Teams Exporter] scroll pass', {
                passes,
                newHeight,
                newCount,
                aggregated: seen,
                oldestTime,
                oldestTs,
                oldestId,
                elapsedMs,
                reason: 'progress report'
            });

            if (startLimit != null && oldestTs != null && oldestTs <= startLimit) {
                console.debug('[Teams Exporter] breaking scroll: startAt reached', { oldestVisible: oldestTimeAttr, startAtISO });
                break;
            }

            const heightUnchanged = newHeight === prevHeight;
            const countUnchanged = newCount === lastCount;
            const oldestUnchanged = oldestId && lastOldestId === oldestId;

            if (heightUnchanged && countUnchanged) {
                stagnantPasses++;
                console.debug('[Teams Exporter] scroll metrics unchanged', { passes, stagnantPasses, newHeight, newCount });
            } else if (oldestUnchanged) {
                stagnantPasses++;
                console.debug('[Teams Exporter] oldest id unchanged', { passes, stagnantPasses, oldestId });
            } else {
                stagnantPasses = 0;
            }

            if (oldestId && lastOldestId !== oldestId) {
                lastOldestId = oldestId;
            }

            prevHeight = newHeight;
            lastCount = newCount;

            if (topReached && stagnantPasses >= 3) {
                console.debug('[Teams Exporter] breaking scroll: header sentinel reached & stagnant', { passes, oldestId, stagnantPasses });
                break;
            }

            if (!topReached && stagnantPasses >= 12) {
                console.debug('[Teams Exporter] breaking scroll: stagnation threshold', { passes, oldestId, stagnantPasses });
                break;
            }
        }
    } finally {
        if (observer && headerSentinel) observer.disconnect();
    }

    await hydrateSparseMessages(agg, { includeSystem, includeReactions });

    // 2) Build sorted list:
    const entries = Array.from(agg.values());
    entries.sort((a, b) => a.orderKey - b.orderKey); // timestamps first; dividers placed near their discovery/parsed time

    // Align system dividers just before the next message when their timestamp is ambiguous
    let nextMessageTs = null;
    for (let i = entries.length - 1; i >= 0; i--) {
        const entry = entries[i];
        if (entry.kind === 'message') {
            if (entry.tsMs != null) nextMessageTs = entry.tsMs;
            continue;
        }

        if (nextMessageTs != null) {
        if (entry.tsMs == null || entry.tsMs >= nextMessageTs) {
            entry.anchorTs = nextMessageTs;
            entry.tsMs = (entry.tsMs == null ? nextMessageTs : entry.tsMs) - 1;
            if (entry.tsMs != null) {
                entry.orderKey = entry.tsMs - 0.1;
            }
            }
        }
    }

    let filtered = entries.filter(entry => entry.kind !== 'day-divider');
    filtered.sort((a, b) => {
        const aTs = (a.tsMs ?? a.anchorTs ?? a.orderKey ?? 0);
        const bTs = (b.tsMs ?? b.anchorTs ?? b.orderKey ?? 0);
        if (aTs !== bTs) return aTs - bTs;
        return a.orderKey - b.orderKey;
    });

    filtered = filtered.filter(entry => {
        const ts = entry.anchorTs ?? entry.tsMs ?? (entry.message?.timestamp ? parseTimeStamp(entry.message.timestamp) : null);
        if (ts == null) return true;
        if (startLimit != null && ts < startLimit) return false;
        if (endLimit != null && ts >= endLimit) return false;
        return true;
    });

    const buckets = new Map<number, { ts: number; message: ExtractedMessage }[]>();
    const noDate: { ts: number; message: ExtractedMessage }[] = [];

    for (const entry of filtered) {
        const msg = entry.message;
        if (!msg) continue;
        if (msg.system && (!msg.text || msg.text.trim().toLowerCase() === 'system')) {
            continue;
        }
        const ts = entry.anchorTs ?? entry.tsMs ?? (msg.timestamp ? parseTimeStamp(msg.timestamp) : null);
        if (ts == null) {
            noDate.push({ ts: Number.MIN_SAFE_INTEGER, message: msg });
            continue;
        }
        const dayKey = startOfLocalDay(ts);
        if (!buckets.has(dayKey)) {
            buckets.set(dayKey, []);
        }
        const list = buckets.get(dayKey);
        if (list) list.push({ ts, message: msg });
    }

    const finalMessages: ExportMessage[] = [];
    const sortedDayKeys = Array.from(buckets.keys()).sort((a, b) => a - b);
    for (const dayKey of sortedDayKeys) {
        const items = buckets.get(dayKey);
        if (!items || !items.length) continue;
        const representativeTs = items[0].ts;
        const divider = makeDayDivider(dayKey, representativeTs);
        if (divider.message) finalMessages.push(divider.message);
        items.sort((a, b) => a.ts - b.ts);
        for (const item of items) {
            finalMessages.push(item.message);
        }
    }

    noDate.sort((a, b) => a.ts - b.ts);
    for (const entry of noDate) {
        finalMessages.push(entry.message);
    }

    return finalMessages;
}

// Remove quoted/preview blocks from a cloned content node so root "text" doesn't include them
function stripQuotedPreview(container: Element | null): Element | null {
  if (!container) return container;
  const clone = container.cloneNode(true) as Element;

  // Known containers for quoted/preview content
  const kill = [
    '[data-tid="quoted-reply-card"]',
    '[data-tid="referencePreview"]',
    '[role="group"][aria-label^="Begin Reference"]'
  ];
  for (const sel of kill) {
    clone.querySelectorAll(sel).forEach((n: Element) => n.remove());
  }

  // Headings like "Begin Reference, …"
  clone.querySelectorAll('div[role="heading"]').forEach((h: Element) => {
    const txt = textFrom(h);
    if (/^Begin Reference,/i.test(txt)) h.remove();
  });

  return clone;
}

// Bridge --------------------------------------------------------
runtime.onMessage.addListener((msg, _sender, sendResponse) => {
    (async () => {
        try {
            if (msg.type === 'PING') { sendResponse({ ok: true }); return; }
            if (msg.type === 'CHECK_CHAT_CONTEXT') { sendResponse(checkChatContext()); return; }
            if (msg.type === 'SCRAPE_TEAMS') {
                const { startAt, endAt, includeReactions, includeSystem, includeReplies, showHud } = msg.options || {};
                hudEnabled = showHud !== false;
                if (!hudEnabled) clearHUD();
                const scrapeOpts = { startAtISO: startAt, endAtISO: endAt, includeSystem, includeReactions, includeReplies: includeReplies !== false };
                console.debug('[Teams Exporter] SCRAPE_TEAMS', location.href, msg.options);
                currentRunStartedAt = Date.now();
                hud('starting…');
                const messages = await autoScrollAggregate(scrapeOpts);
                try {
                    const msgPromise = runtime.sendMessage({ type: 'SCRAPE_PROGRESS', payload: { phase: 'extract', messagesExtracted: messages.length } });
                    if (msgPromise && msgPromise.catch) msgPromise.catch(() => { });
                } catch (e) { /* ignore */ }
                hud(`extracted ${messages.length} messages`);
                currentRunStartedAt = null;
                // meta can keep title; add timeRange later if you want
                sendResponse({ messages, meta: { count: messages.length, title: document.title, startAt: startAt || null, endAt: endAt || null } });
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
