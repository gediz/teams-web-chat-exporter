/* eslint-disable no-console */
const $$ = (sel, root = document) => Array.from(root.querySelectorAll(sel));
const $ = (sel, root = document) => root.querySelector(sel);

// HUD -----------------------------------------------------------
function ensureHUD() { let hud = document.getElementById("__teamsExporterHUD"); if (!hud) { hud = document.createElement("div"); hud.id = "__teamsExporterHUD"; hud.style.cssText = "position:fixed;right:12px;top:12px;z-index:999999;font:12px/1.3 system-ui,sans-serif;color:#111;background:rgba(255,255,255,.92);border:1px solid #ddd;border-radius:8px;padding:8px 10px;box-shadow:0 2px 8px rgba(0,0,0,.15);pointer-events:none;"; hud.textContent = "Teams Exporter: idle"; document.body.appendChild(hud); } return hud; }
function hud(text) { ensureHUD().textContent = `Teams Exporter: ${text}`; chrome.runtime.sendMessage({ type: "SCRAPE_PROGRESS", payload: { phase: "hud", text } }).catch(() => { }); }

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
    const out = [];
    const seen = new Set();
    const push = (href, label) => {
        if (!href && !label) return;
        const key = `${href || ''}@@${label || ''}`;
        if (seen.has(key)) return;
        seen.add(key);
        out.push({ href, label });
    };
    const parseTitle = (t) => {
        if (!t) return null;
        const parts = t.split(/\n+/).map(s => s.trim()).filter(Boolean);
        // Common case: "filename.ext\nhttps://sharepoint/..." (or OneDrive)
        if (parts.length >= 2 && /^https?:\/\//i.test(parts[1])) {
            return { label: parts[0], href: parts[1] };
        }
        // Fallback: extract first URL and use the remainder as label
        const m = t.match(/https?:\/\/\S+/);
        if (m) {
            const url = m[0];
            const label = (parts[0] && parts[0] !== url) ? parts[0] : url;
            return { label, href: url };
        }
        return null;
    };

    // Collect likely roots
    const roots = [];
    const aria = body?.getAttribute('aria-labelledby') || '';
    const attId = aria.split(/\s+/).find(s => s.startsWith('attachments-'));
    if (attId) {
        const el = document.getElementById(attId);
        if (el) roots.push(el); // explicit attachments container (covers attachment-only posts)
    }
    // Standard containers seen in your snippets
    ['[data-tid="file-attachment-grid"]', '[data-tid="file-preview-root"]', '[data-tid="attachments"]'].forEach(sel => {
        const el = body && body.querySelector(sel);
        if (el && !roots.includes(el)) roots.push(el);
    });

    // 1) File chiclets / tiles (title or aria-label contains "filename\nURL")
    for (const root of roots) {
        // Chiclet role
        root.querySelectorAll('[data-testid="file-attachment"], [data-tid^="file-chiclet-"]').forEach(el => {
            const t = el.getAttribute('title') || el.getAttribute('aria-label') || '';
            const parsed = parseTitle(t);
            if (parsed) push(parsed.href, parsed.label);
            // Nested anchor fallback
            el.querySelectorAll('a').forEach(a => push(a.href, a.innerText || a.getAttribute('aria-label') || a.title || a.href));
        });
        // Rich preview button (title has filename + URL)
        root.querySelectorAll('button[data-testid="rich-file-preview-button"][title]').forEach(btn => {
            const parsed = parseTitle(btn.getAttribute('title'));
            if (parsed) push(parsed.href, parsed.label);
        });
        // Generic anchors under the grid (belt-and-suspenders)
        root.querySelectorAll('a[href^="http"]').forEach(a => {
            push(a.href, a.innerText || a.getAttribute('aria-label') || a.title || a.href);
        });
    }

    // 2) Inline links inside the message content (e.g., OneDrive "safelinks")
    const contentRoot = body && (body.querySelector('[id^="content-"]') || body.querySelector('[data-tid="message-content"]'));
    if (contentRoot) {
        contentRoot.querySelectorAll('a[data-testid="atp-safelink"], a[href^="http"]').forEach(a => {
            push(a.href, a.innerText || a.getAttribute('aria-label') || a.title || a.href);
        });
        // 3) Posted images in the content
        contentRoot.querySelectorAll('[data-testid="lazy-image-wrapper"] img').forEach(img => {
            // Avoid tiny UI icons by requiring an http(s) src
            const src = img.getAttribute('src') || '';
            if (/^https?:\/\//i.test(src)) push(src, img.getAttribute('alt') || 'image');
        });
    }

    return out;
}

// Helpers -------------------------------------------------------
const sleep = (ms) => new Promise(r => setTimeout(r, ms));

function parseDateDividerText(txt, yearHint) {
    // "7 September" → try add current year; leave null if unparsable
    const m = txt.match(/^(\d{1,2})\s+([A-Za-z]+)(?:\s+(\d{4}))?$/);
    if (!m) return null;
    const y = m[3] || yearHint || String(new Date().getFullYear());
    const d = Date.parse(`${m[1]} ${m[2]} ${y}`);
    return Number.isNaN(d) ? null : d;
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
        const dividerId = (wrapper?.id || '').trim() || text.toLowerCase();
        let approxMs = parseDateDividerText(text, orderCtx.yearHint);
        if (approxMs == null) approxMs = (orderCtx.lastTimeMs ?? Date.now()) + 1; // place after last seen
        return {
            message: { id: dividerId, author: '[system]', timestamp: '', text, reactions: [], attachments: [], edited: false, avatar: null, replyTo: null, system: true },
            orderKey: approxMs
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
    const replyTo = extractReplyContext(item, body);

    const mid = body.getAttribute('data-mid') || item.id || `${ts}#${author}`;
    const msg = { id: mid, author, timestamp: ts, text, reactions, attachments, edited, avatar, replyTo, system: false };

    const orderKey = !Number.isNaN(tms) ? tms : (orderCtx.seqBase + orderCtx.seq++);
    return { message: msg, orderKey };
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
        const { message, orderKey } = extracted;

        agg.set(message.id, { message, orderKey });
        if (!message.system && message.timestamp) {
            const tms = Date.parse(message.timestamp);
            if (!Number.isNaN(tms)) { orderCtx.lastTimeMs = tms; orderCtx.yearHint = new Date(tms).getFullYear(); }
        }
        if (!message.system && message.author) {
            orderCtx.lastAuthor = message.author;
        }
    }
}

async function autoScrollAggregate({ stopAtISO, includeSystem, includeReactions }) {
    const scroller = getScroller();
    if (!scroller) throw new Error('Scroller not found');

    const agg = new Map();         // id -> {message, orderKey}
    const orderCtx = { lastTimeMs: undefined, yearHint: undefined, seqBase: Date.now(), seq: 0, lastAuthor: '' };

    // 0) Pre-capture bottom (newest window)
    scroller.scrollTop = scroller.scrollHeight;
    await new Promise(r => requestAnimationFrame(r));
    await sleep(300);
    await collectCurrentVisible(agg, { includeSystem, includeReactions }, orderCtx);

    // 1) Scroll to top repeatedly to load older history, collecting each pass
    let prevHeight = -1, lastCount = -1, passes = 0;
    const maxPasses = 150, dwellMs = 700;

    while (passes < maxPasses) {
        passes++;
        scroller.scrollTop = 0;
        await new Promise(r => requestAnimationFrame(r));
        await sleep(dwellMs);

        await collectCurrentVisible(agg, { includeSystem, includeReactions }, orderCtx);

        const nodes = $$('[data-tid="chat-pane-item"]');
        const newCount = nodes.length;
        const newHeight = scroller.scrollHeight;

        if (passes % 4 === 0) {
            chrome.runtime.sendMessage({ type: 'SCRAPE_PROGRESS', payload: { phase: 'scroll', passes, newHeight, messagesVisible: newCount, aggregated: agg.size } }).catch(() => { });
            hud(`scroll pass ${passes} • height ${newHeight} • seen ${agg.size}`);
        }

        if (stopAtISO) {
            const oldestVisible = $('time[datetime]', nodes[0])?.getAttribute('datetime');
            if (oldestVisible && oldestVisible <= stopAtISO) break;
        }
        if (newHeight === prevHeight && newCount === lastCount) break;
        prevHeight = newHeight; lastCount = newCount;
    }

    // 2) Build sorted list:
    const entries = Array.from(agg.values());
    entries.sort((a, b) => a.orderKey - b.orderKey); // timestamps first; dividers placed near their discovery/parsed time

    return entries.map(e => e.message);
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
            if (msg.type === 'SCRAPE_TEAMS') {
                const { stopAt, includeReactions, includeSystem } = msg.options || {};
                console.debug('[Teams Exporter] SCRAPE_TEAMS', location.href, msg.options);
                hud('starting…');
                const messages = await autoScrollAggregate({ stopAtISO: stopAt, includeSystem, includeReactions });
                chrome.runtime.sendMessage({ type: 'SCRAPE_PROGRESS', payload: { phase: 'extract', messagesExtracted: messages.length } }).catch(() => { });
                hud(`extracted ${messages.length} messages`);
                // meta can keep title; add timeRange later if you want
                sendResponse({ messages, meta: { count: messages.length, title: document.title } });
            }
        } catch (e) {
            console.error('[Teams Exporter] Error:', e);
            hud(`error: ${e.message}`);
            sendResponse({ error: e?.message || String(e) });
        }
    })();
    return true;
});
