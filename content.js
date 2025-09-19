/* eslint-disable no-console */
const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));
const $ = (sel, root = document) => root.querySelector(sel);

let hudEnabled = true;
let currentRunStartedAt = null;

function isChatNavSelected() {
    return Boolean(document.querySelector('[data-tid="app-bar-wrapper"] button[aria-label="Chat"][aria-pressed="true"]'));
}

function hasChatMessageSurface() {
    return Boolean(
        document.querySelector('[data-tid="message-pane-list-viewport"], [data-tid="chat-message-list"], [data-tid="chat-pane"]')
    );
}

function parseTimeStamp(value) {
    if (!value) return null;
    const ts = Date.parse(value);
    if (!Number.isNaN(ts)) return ts;
    // Teams sometimes uses without timezone; treat as local.
    const normalized = value.replace(/ /g, 'T');
    const ts2 = Date.parse(normalized);
    return Number.isNaN(ts2) ? null : ts2;
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
function hud(text, { includeElapsed = true } = {}) {
    if (!hudEnabled) return;
    const hudNode = ensureHUD();
    if (hudNode) {
        let final = `Teams Exporter: ${text}`;
        if (includeElapsed !== false && currentRunStartedAt) {
            final += ` • elapsed ${formatElapsed(Date.now() - currentRunStartedAt)}`;
        }
        hudNode.textContent = final;
    }
    chrome.runtime.sendMessage({ type: "SCRAPE_PROGRESS", payload: { phase: "hud", text } }).catch(() => { });
}

function formatElapsed(ms) {
    const totalSeconds = Math.max(0, Math.floor(ms / 1000));
    const hours = Math.floor(totalSeconds / 3600);
    const minutes = Math.floor((totalSeconds % 3600) / 60);
    const seconds = totalSeconds % 60;
    if (hours > 0) {
        return `${hours}:${String(minutes).padStart(2, "0")}:${String(seconds).padStart(2, "0")}`;
    }
    return `${minutes}:${String(seconds).padStart(2, "0")}`;
}

// Core DOM hooks ------------------------------------------------
function getScroller() {
    return $('[data-tid="message-pane-list-viewport"]') || $('[data-tid="chat-message-list"]') || document.scrollingElement;
}

// Author/timestamp/edited/avatar helpers ------------------------
function resolveAuthor(body, lastAuthor = "") {
    let author = ($('[data-tid="message-author-name"]', body)?.innerText || '').trim();
    if (!author) {
        const aria = body.getAttribute('aria-labelledby') || '';
        const aId = aria.split(/\s+/).find(s => s.startsWith('author-'));
        if (aId) author = (document.getElementById(aId)?.innerText || '').trim();
    }
    return author || lastAuthor || '';
}
function resolveTimestamp(item) {
    const t = $('time[datetime]', item) || $('time', item) || $('[data-tid="message-status"] time', item);
    return t?.getAttribute?.('datetime') || t?.getAttribute?.('title') || t?.innerText || '';
}
function resolveEdited(item, body) {
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
function resolveAvatar(item) {
    const perMsg = $('[data-tid="message-avatar"] img', item); // per-message avatar
    if (perMsg?.src) return perMsg.src;                       // :contentReference[oaicite:4]{index=4}
    const header = $('[data-tid="chat-title-avatar"] img');   // header fallback
    return header?.src || null;                                // :contentReference[oaicite:5]{index=5}
}

// Text with emoji (IMG alt) + block breaks
function extractTextWithEmojis(root) {
    if (!root) return '';
    let out = ''; const walk = (n) => {
        if (n.nodeType === Node.TEXT_NODE) { out += n.nodeValue; return; }
        if (n.nodeType !== Node.ELEMENT_NODE) return;
        const el = n, tag = el.tagName;
        if (tag === 'BR') { out += '\n'; return; }
        if (tag === 'IMG') { out += (el.getAttribute('alt') || el.getAttribute('aria-label') || ''); return; }
        const blockish = /^(DIV|P|LI|BLOCKQUOTE)$/; const start = out.length;
        for (const c of el.childNodes) walk(c);
        if (blockish.test(tag) && out.length > start) out += '\n';
    }; walk(root);
    return out.replace(/\n{3,}/g, '\n\n').trim();
}

function extractReplyContext(item, body) {
  // 1) Structured quoted-reply card (preferred)
  const card = body?.querySelector('[data-tid="quoted-reply-card"]');
  if (card) {
    const tsEl = card.querySelector('[data-tid="quoted-reply-timestamp"]');
    const authorEl = tsEl?.previousElementSibling || null; // author sits right before timestamp
    const textEl = card.querySelector('[data-tid="quoted-reply-preview-content"]');

    const author = (authorEl?.innerText || '').trim();
    const timestamp = (tsEl?.innerText || '').trim(); // e.g. "12/09/2025, 11:12"
    const text = extractTextWithEmojis(textEl || card).trim(); // full preview text

    if (author || timestamp || text) return { author, timestamp, text };
  }

  // 2) Fallback: aria-label on the group with "Begin Reference, …, <author>, <timestamp>, End reference"
  const group = body?.querySelector('[role="group"][aria-label^="Begin Reference"]');
  if (group) {
    const aria = group.getAttribute('aria-label') || '';
    // Greedy capture for text; last two comma-separated tokens are author and timestamp
    const m = aria.match(/^Begin Reference,\s*(.*),\s*([^,]+),\s*([^,]+),\s*End reference$/s);
    if (m) {
      const [, text, author, timestamp] = m;
      return { author: (author||'').trim(), timestamp: (timestamp||'').trim(), text: (text||'').trim() };
    }
  }

  // 3) Legacy fallback: "Begin Reference, … by <author>"
  const heading = item.querySelector('div[role="heading"]');
  if (heading) {
    const raw = heading.innerText || '';
    const m = raw.match(/Begin Reference,\s*(.*?)\s*by\s*(.+)$/i);
    if (m) return { author: m[2].trim(), timestamp: '', text: m[1].trim() };
  }
  return null;
}


function extractReactions(item) {
  const pills = Array.from(item.querySelectorAll('[data-tid="diverse-reaction-pill-button"]'));
  const out = [];

  for (const btn of pills) {
    // Emoji icon (static)
    const emoji = btn.querySelector('[data-tid="emoticon-renderer"] img')?.getAttribute('alt') || '';

    // Pull the descriptive text WITHOUT interacting
    // Prefer aria-labelledby targets (Teams often renders the pill’s label there)
    let labelText = '';
    const labelledBy = btn.getAttribute('aria-labelledby');
    if (labelledBy) {
      labelText = labelledBy
        .split(/\s+/)
        .map(id => (document.getElementById(id)?.innerText || '').trim())
        .filter(Boolean)
        .join(' ')
        .trim();
    }
    if (!labelText) {
      labelText = (btn.getAttribute('aria-label') || '').trim();
    }

    // Parse count (first integer we find; Teams strings usually contain it)
    const count = parseInt((labelText.match(/\d+/) || [1])[0], 10) || 1;

    const entry = { emoji, count };

    // Inline-only names (no clicking): if label lists names before “react…”
    // e.g., "Alice, Bob and 2 others reacted with 👍"
    if (labelText) {
      const beforeReact = labelText.split(/react/i)[0]; // "Alice, Bob and 2 others "
      // Split out explicit names; ignore the "X others" tail
      let names = beforeReact
        .split(/,\s*|\s+and\s+/)
        .map(s => s.trim())
        .filter(Boolean)
        .filter(s => !/^\d+\s+others?$/i.test(s) && !/others$/i.test(s));
      if (names.length) entry.reactors = Array.from(new Set(names)).slice(0, 100);
    }

    out.push(entry);
  }
  return out;
}


// Attachments (robust)
// --- replace extractAttachments in content.js ---
function extractAttachments(item, body) {
    const map = new Map();
    const merge = (prev, next) => {
        const merged = { ...prev };
        if (!merged.href && next.href) merged.href = next.href;
        if (!merged.label && next.label) merged.label = next.label;
        for (const field of ['type', 'size', 'owner', 'metaText']) {
            if (!merged[field] && next[field]) merged[field] = next[field];
        }
        return merged;
    };
    const guessTypeFromLabel = (label = '') => {
        const m = label.trim().match(/\.([A-Za-z0-9]{1,6})$/);
        return m ? m[1].toUpperCase() : null;
    };
    const collectMetaText = (node, label) => {
        const parts = new Set();
        const add = (val) => {
            if (!val || typeof val !== 'string') return;
            const trimmed = val.trim();
            if (!trimmed) return;
            parts.add(trimmed);
        };
        if (node) {
            add(node.getAttribute?.('aria-label'));
            add(node.getAttribute?.('title'));
            const txt = node.textContent?.trim();
            if (txt && txt !== label?.trim()) add(txt);
        }
        if (!parts.size) return '';
        return Array.from(parts).join(' • ').replace(/\s+/g, ' ').trim();
    };
    const inferOwner = (text = '') => {
        const match = text.match(/(?:Shared by|Uploaded by|Sent by|From|Owner)\s*:?\s*([^•]+)/i);
        return match ? match[1].trim() : null;
    };
    const inferSize = (text = '') => {
        const match = text.match(/\b\d+(?:[.,]\d+)?\s*(?:bytes?|KB|MB|GB|TB)\b/i);
        return match ? match[0].replace(',', '.').trim() : null;
    };
    const inferType = (label, text) => {
        return guessTypeFromLabel(label) || (text ? (text.match(/\b(PDF|DOCX|XLSX|PPTX|TXT|PNG|JPE?G|GIF|ZIP|RAR|CSV|MP4|MP3)\b/i)?.[0]?.toUpperCase() || null) : null);
    };
    const push = (sourceNode, data = {}) => {
        const att = { ...data };
        if (!att.href && sourceNode?.href) att.href = sourceNode.href;
        if (!att.label) {
            const ariaLabel = sourceNode?.getAttribute?.('aria-label');
            if (ariaLabel) att.label = ariaLabel.split(/\n+/)[0].trim();
        }
        if (!att.label && sourceNode?.getAttribute?.('title')) {
            att.label = sourceNode.getAttribute('title').split(/\n+/)[0].trim();
        }
        if (!att.label && sourceNode?.textContent) {
            const text = sourceNode.textContent.trim();
            if (text) att.label = text.split(/\n+/)[0].trim();
        }
        if (!att.href && !att.label) return;

        const metaText = collectMetaText(sourceNode, att.label);
        if (metaText) att.metaText = metaText;
        const type = inferType(att.label || '', metaText);
        if (type) att.type = type;
        const size = inferSize(metaText);
        if (size) att.size = size;
        const owner = inferOwner(metaText);
        if (owner) att.owner = owner;

        const key = `${att.href || ''}@@${att.label || ''}`;
        const prev = map.get(key);
        map.set(key, prev ? merge(prev, att) : att);
    };
    const parseTitle = (t) => {
        if (!t) return null;
        const parts = t.split(/\n+/).map(s => s.trim()).filter(Boolean);
        if (parts.length >= 2 && /^https?:\/\//i.test(parts[1])) {
            return { label: parts[0], href: parts[1], metaText: parts.slice(2).join(' • ') };
        }
        const m = t.match(/https?:\/\/\S+/);
        if (m) {
            const url = m[0];
            const label = (parts[0] && parts[0] !== url) ? parts[0] : url;
            return { label, href: url, metaText: parts.slice(1).join(' • ') };
        }
        return null;
    };

    const roots = [];
    const aria = body?.getAttribute('aria-labelledby') || '';
    const attId = aria.split(/\s+/).find(s => s.startsWith('attachments-'));
    if (attId) {
        const el = document.getElementById(attId);
        if (el) roots.push(el);
    }
    ['[data-tid="file-attachment-grid"]', '[data-tid="file-preview-root"]', '[data-tid="attachments"]'].forEach(sel => {
        const el = body && body.querySelector(sel);
        if (el && !roots.includes(el)) roots.push(el);
    });

    for (const root of roots) {
        root.querySelectorAll('[data-testid="file-attachment"], [data-tid^="file-chiclet-"]').forEach(el => {
            const t = el.getAttribute('title') || el.getAttribute('aria-label') || '';
            const parsed = parseTitle(t);
            if (parsed) push(el, parsed);
            el.querySelectorAll('a[href^="http"]').forEach(a => {
                const label = a.innerText || a.getAttribute('aria-label') || a.title || a.href;
                push(a, { href: a.href, label });
            });
        });
        root.querySelectorAll('button[data-testid="rich-file-preview-button"][title]').forEach(btn => {
            const parsed = parseTitle(btn.getAttribute('title'));
            if (parsed) push(btn, parsed);
        });
        root.querySelectorAll('a[href^="http"]').forEach(a => {
            const label = a.innerText || a.getAttribute('aria-label') || a.title || a.href;
            push(a, { href: a.href, label });
        });
    }

    const contentRoot = body && (body.querySelector('[id^="content-"]') || body.querySelector('[data-tid="message-content"]'));
    if (contentRoot) {
        contentRoot.querySelectorAll('a[data-testid="atp-safelink"], a[href^="http"]').forEach(a => {
            push(a, { href: a.href, label: a.innerText || a.getAttribute('aria-label') || a.title || a.href });
        });
        contentRoot.querySelectorAll('[data-testid="lazy-image-wrapper"] img').forEach(img => {
            const src = img.getAttribute('src') || '';
            if (/^https?:\/\//i.test(src)) {
                push(img, { href: src, label: img.getAttribute('alt') || 'image' });
            }
        });
    }

    return Array.from(map.values());
}

// Helpers -------------------------------------------------------
const sleep = (ms) => new Promise(r => setTimeout(r, ms));

const cssEscape = (s) => {
    if (typeof CSS !== 'undefined' && CSS.escape) return CSS.escape(s);
    return (s || '').toString().replace(/([\0-\x1f\x7f-\x9f!"#$%&'()*+,./:;<=>?@[\\\]^`{|}~])/g, '\\$1');
};

const normalizeText = (s) => (s ?? '').replace(/\u00a0/g, ' ');
const isPlaceholderText = (s) => {
    const clean = normalizeText(s).trim();
    if (!clean) return true;
    return /^loading(?:\.\.\.?|…)?$/i.test(clean);
};
const textScore = (s) => normalizeText(s).trim().length;
const preferText = (prev, next) => {
    if (next == null) return prev;
    const nextPlaceholder = isPlaceholderText(next);
    const prevPlaceholder = isPlaceholderText(prev);
    if (nextPlaceholder && !prevPlaceholder) return prev;
    if (!nextPlaceholder && prevPlaceholder) return next;
    if (nextPlaceholder && prevPlaceholder) return prev;
    const prevLen = textScore(prev);
    const nextLen = textScore(next);
    return nextLen >= prevLen ? next : prev;
};

const findPaneItemByMessageId = (id) => {
    if (!id) return null;
    const msgNode = document.querySelector(`[data-mid="${cssEscape(id)}"]`);
    return msgNode?.closest('[data-tid="chat-pane-item"]') || null;
};

async function hydrateSparseMessages(agg, opts = {}) {
    if (!agg || agg.size === 0) return;

    const needsHydration = (message, item) => {
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

    let pending = [];

    for (const [id, entry] of agg.entries()) {
        const msg = entry.message;
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

            const item = findPaneItemByMessageId(id) || task.item;
            if (!item) continue;

            const lastAuthorRef = { value: existing.message.author || '' };
            const ts = existing.message.timestamp ? Date.parse(existing.message.timestamp) : undefined;
            const tempOrderCtx = {
                lastTimeMs: Number.isNaN(ts) ? undefined : ts,
                yearHint: Number.isNaN(ts) ? undefined : new Date(ts).getFullYear(),
                seqBase: Date.now(),
                seq: 0,
                lastAuthor: existing.message.author || ''
            };

            const reExtracted = await extractOne(item, { includeSystem: opts.includeSystem, includeReactions: opts.includeReactions }, lastAuthorRef, tempOrderCtx);
            if (!reExtracted?.message) {
                nextPending.push(task);
                continue;
            }

            const merged = { ...existing.message };
            merged.text = preferText(existing.message.text, reExtracted.message.text);

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
            merged.edited = merged.edited || reExtracted.message.edited;

            const newTsMs = reExtracted?.tsMs ?? existing.tsMs ?? (merged.timestamp ? parseTimeStamp(merged.timestamp) : null);
            const kind = existing.kind ?? reExtracted?.kind;
            agg.set(id, { message: merged, orderKey: existing.orderKey, tsMs: newTsMs, kind });

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

function parseDateDividerText(txt, yearHint) {
    if (!txt) return null;
    const monthMap = {
        january: 1, february: 2, march: 3, april: 4, may: 5, june: 6,
        july: 7, august: 8, september: 9, october: 10, november: 11, december: 12
    };

    const clean = txt.trim().replace(/\s+/g, ' ');
    const currentYear = typeof yearHint === 'number' ? yearHint : (yearHint ? Number(yearHint) : new Date().getFullYear());

    const tryBuild = (dayStr, monthStr, yearStr) => {
        if (!dayStr || !monthStr) return null;
        const day = Number(dayStr);
        if (!Number.isFinite(day)) return null;
        const monthIdx = monthMap[monthStr.toLowerCase()];
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

// Extract one item into a message object + an orderKey
async function extractOne(item, opts, lastAuthorRef, orderCtx) {
    const body = $('[data-tid="chat-pane-message"]', item) || item;
    const isSystem = !$('[data-tid="chat-pane-message"]', item);

    // Date/system divider
    if (isSystem) {
        if (!opts.includeSystem) return null;
        const wrapper = $('.fui-Divider__wrapper', item); // "7 September", "Monday", etc. :contentReference[oaicite:6]{index=6}
        const text = (wrapper?.innerText || item.innerText || '').trim() || 'system';
        const bodyMid = wrapper?.id || $('[data-mid]', item)?.getAttribute('data-mid') || item.getAttribute('data-mid');
        const dividerId = (bodyMid || text || 'system').toLowerCase();
        const numericMid = bodyMid && Number(bodyMid);

        let approxMs = parseDateDividerText(text, orderCtx.yearHint);
        if (!Number.isFinite(approxMs)) {
            if (Number.isFinite(numericMid)) {
                approxMs = numericMid;
            } else if (typeof orderCtx.lastTimeMs === 'number') {
                approxMs = orderCtx.lastTimeMs - 1;
            } else {
                approxMs = orderCtx.systemCursor++;
            }
        }

        orderCtx.lastTimeMs = approxMs;
        orderCtx.yearHint = new Date(approxMs).getFullYear();
        return {
            message: { id: dividerId, author: '[system]', timestamp: '', text, reactions: [], attachments: [], edited: false, avatar: null, replyTo: null, system: true },
            orderKey: approxMs,
            tsMs: approxMs,
            kind: 'system'
        };
    }

    // Normal message
    const ts = resolveTimestamp(item);
    const tms = ts ? Date.parse(ts) : NaN;
    if (!Number.isNaN(tms)) orderCtx.lastTimeMs = tms, orderCtx.yearHint = new Date(tms).getFullYear();

    const author = resolveAuthor(body, lastAuthorRef.value || orderCtx.lastAuthor || '');
    if (author) {
        lastAuthorRef.value = author;
        orderCtx.lastAuthor = author;
    }

    const contentEl = $('[id^="content-"]', body) || $('[data-tid="message-content"]', body) || body;
    const cleanRoot = stripQuotedPreview(contentEl);
    const text = extractTextWithEmojis(cleanRoot);
    const edited = resolveEdited(item, body);
    const avatar = resolveAvatar(item);
    const reactions = opts.includeReactions ? await extractReactions(item) : [];

    const attachments = extractAttachments(item, body);
    const replyTo = opts.includeReplies === false ? null : extractReplyContext(item, body);

    const mid = body.getAttribute('data-mid') || item.id || `${ts}#${author}`;
    const msg = { id: mid, author, timestamp: ts, text, reactions, attachments, edited, avatar, replyTo, system: false };

    const orderKey = !Number.isNaN(tms) ? tms : (orderCtx.seqBase + orderCtx.seq++);
    const tsMs = !Number.isNaN(tms) ? tms : null;
    return { message: msg, orderKey, tsMs, kind: 'message' };
}

// Aggregate while scrolling so virtualization can’t drop items
async function collectCurrentVisible(agg, opts, orderCtx) {
    const nodes = $$('[data-tid="chat-pane-item"]'); // preserve DOM order for system dividers, too
    const lastAuthorRef = { value: orderCtx.lastAuthor || '' };
    for (let i = 0; i < nodes.length; i++) {
        const item = nodes[i];
        const idCandidate = $('[data-tid="chat-pane-message"]', item)?.getAttribute('data-mid') || $('.fui-Divider__wrapper', item)?.id || item.id || `node-${i}`;
        if (agg.has(idCandidate)) continue;

        const extracted = await extractOne(item, opts, lastAuthorRef, orderCtx);
        if (!extracted) continue;
        const { message, orderKey, tsMs, kind } = extracted;

        agg.set(message.id, { message, orderKey, tsMs, kind });
        if (!message.system && message.timestamp) {
            const tms = Date.parse(message.timestamp);
            if (!Number.isNaN(tms)) { orderCtx.lastTimeMs = tms; orderCtx.yearHint = new Date(tms).getFullYear(); }
        }
        if (!message.system && message.author) {
            orderCtx.lastAuthor = message.author;
        }
    }
}

async function autoScrollAggregate({ stopAtISO, includeSystem, includeReactions, includeReplies = true }) {
    const scroller = getScroller();
    if (!scroller) throw new Error('Scroller not found');

    const agg = new Map();         // id -> {message, orderKey}
    const orderCtx = { lastTimeMs: undefined, yearHint: undefined, seqBase: Date.now(), seq: 0, lastAuthor: '', systemCursor: -9e15 };

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

    const stopLimit = typeof stopAtISO === 'string' ? parseTimeStamp(stopAtISO) : null;

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
            const hiddenButtons = Array.from(document.querySelectorAll('[data-tid="show-hidden-chat-history-btn"]'))
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
            let filteredSeen = seen;
            if (stopLimit != null) {
                filteredSeen = 0;
                for (const entry of agg.values()) {
                    const candidate = entry?.tsMs ?? (entry?.message?.timestamp ? parseTimeStamp(entry.message.timestamp) : null);
                    if (candidate == null || candidate >= stopLimit) filteredSeen++;
                }
            }

            chrome.runtime.sendMessage({ type: 'SCRAPE_PROGRESS', payload: { phase: 'scroll', passes, newHeight, messagesVisible: newCount, aggregated: seen, seen: filteredSeen, filteredSeen, oldestTime, oldestId, elapsedMs } }).catch(() => { });
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

            if (stopLimit != null && oldestTs != null && oldestTs <= stopLimit) {
                console.debug('[Teams Exporter] breaking scroll: stopAt reached', { oldestVisible: oldestTimeAttr, stopAtISO });
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

    let filtered = entries;
    if (stopLimit != null) {
        filtered = entries.filter(entry => {
            const ts = entry.tsMs ?? (entry.message?.timestamp ? parseTimeStamp(entry.message.timestamp) : null);
            if (entry.kind === 'system') {
                if (ts == null) return false;
                return ts >= stopLimit;
            }
            if (ts == null) return true;
            return ts >= stopLimit;
        });

        const firstMessageIdx = filtered.findIndex(entry => entry.kind === 'message');
        if (firstMessageIdx > 0) {
            const firstMessage = filtered[firstMessageIdx];
            const firstTs = firstMessage.tsMs ?? (firstMessage.message?.timestamp ? parseTimeStamp(firstMessage.message.timestamp) : null);
            const idxInEntries = entries.indexOf(firstMessage);
            for (let i = idxInEntries - 1; i >= 0; i--) {
                const candidate = entries[i];
                if (candidate.kind === 'system') {
                    const ts = candidate.tsMs ?? (candidate.message?.timestamp ? parseTimeStamp(candidate.message.timestamp) : null);
                    if (ts != null && ts >= stopLimit && (firstTs == null || ts <= firstTs)) {
                        if (!filtered.includes(candidate)) {
                            filtered.splice(firstMessageIdx, 0, candidate);
                        }
                    }
                    break;
                }
                if (candidate.kind === 'message') break;
            }
        }
    }

    return filtered.map(e => e.message);
}

// Remove quoted/preview blocks from a cloned content node so root "text" doesn't include them
function stripQuotedPreview(container) {
  if (!container) return container;
  const clone = container.cloneNode(true);

  // Known containers for quoted/preview content
  const kill = [
    '[data-tid="quoted-reply-card"]',
    '[data-tid="referencePreview"]',
    '[role="group"][aria-label^="Begin Reference"]'
  ];
  for (const sel of kill) {
    clone.querySelectorAll(sel).forEach(n => n.remove());
  }

  // Headings like "Begin Reference, …"
  clone.querySelectorAll('div[role="heading"]').forEach(h => {
    const txt = (h.innerText || '').trim();
    if (/^Begin Reference,/i.test(txt)) h.remove();
  });

  return clone;
}

// Bridge --------------------------------------------------------
chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
    (async () => {
        try {
            if (msg.type === 'PING') { sendResponse({ ok: true }); return; }
            if (msg.type === 'CHECK_CHAT_CONTEXT') { sendResponse(checkChatContext()); return; }
            if (msg.type === 'SCRAPE_TEAMS') {
                const { stopAt, includeReactions, includeSystem, includeReplies, showHud } = msg.options || {};
                hudEnabled = showHud !== false;
                if (!hudEnabled) clearHUD();
                const scrapeOpts = { stopAtISO: stopAt, includeSystem, includeReactions, includeReplies: includeReplies !== false };
                console.debug('[Teams Exporter] SCRAPE_TEAMS', location.href, msg.options);
                currentRunStartedAt = Date.now();
                hud('starting…');
                const messages = await autoScrollAggregate(scrapeOpts);
                chrome.runtime.sendMessage({ type: 'SCRAPE_PROGRESS', payload: { phase: 'extract', messagesExtracted: messages.length } }).catch(() => { });
                hud(`extracted ${messages.length} messages`);
                currentRunStartedAt = null;
                // meta can keep title; add timeRange later if you want
                sendResponse({ messages, meta: { count: messages.length, title: document.title } });
            }
        } catch (e) {
            console.error('[Teams Exporter] Error:', e);
            hud(`error: ${e.message}`);
            currentRunStartedAt = null;
            sendResponse({ error: e?.message || String(e) });
        }
    })();
    return true;
});
