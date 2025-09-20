// ===== service-worker.js (replace) =====
function log(...a) { try { console.log("[Teams Exporter SW]", ...a) } catch { } }
log("boot");

chrome.runtime.onInstalled.addListener(() => log("onInstalled"));
chrome.runtime.onStartup?.addListener(() => log("onStartup"));

const activeExports = new Map(); // tabId -> { startedAt, lastStatus }
// TERMINAL_PHASES: 'complete' = success, 'error' = failure, 'empty' = no data found (not a failure)
const TERMINAL_PHASES = new Set(['complete', 'error', 'empty']);

function updateActiveExport(tabId, patch = {}) {
    if (tabId == null) return;
    const prev = activeExports.get(tabId) || {};
    const next = { ...prev, ...patch };
    activeExports.set(tabId, next);
    return next;
}

function sanitizeBase(name) {
    const raw = (name || "teams-chat").toString();
    const cleaned = raw.replace(/[<>:"/\\|?*\x00-\x1F]/g, "-").replace(/\s+/g, " ").trim().replace(/[. ]+$/g, "");
    return (cleaned || "teams-chat").slice(0, 80);
}

function formatRangeLabel(startISO, endISO) {
    if (!startISO && !endISO) return null;
    const fmt = new Intl.DateTimeFormat(undefined, { dateStyle: 'medium', timeStyle: 'short' });
    const start = startISO ? fmt.format(new Date(startISO)) : null;
    const end = endISO ? fmt.format(new Date(endISO)) : null;
    if (start && end) return `${start} → ${end}`;
    if (start) return `Since ${start}`;
    if (end) return `Until ${end}`;
    return null;
}

// --- Builders (text only; good enough for chat exports)
const esc = s => (s ?? "").toString().replaceAll('"', '""');

async function fetchAsDataURL(url) {
    const res = await fetch(url, { credentials: "include" });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const buf = await res.arrayBuffer();
    const bytes = new Uint8Array(buf);
    // Convert to base64 (small images; OK for 64x64 avatars)
    let bin = "";
    for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
    const b64 = btoa(bin);
    const ct = res.headers.get("content-type") || "image/png";
    return `data:${ct};base64,${b64}`;
}

async function embedAvatarsInRows(rows) {
    // Deduplicate identical avatar URLs
    const map = new Map(); // url -> dataURL|null (if failed)
    for (const m of rows) {
        const u = m.avatar;
        if (!u || u.startsWith("data:")) continue;
        if (!map.has(u)) {
            try { map.set(u, await fetchAsDataURL(u)); }
            catch { map.set(u, null); }
        }
    }
    // Return a shallow-cloned array with avatars replaced
    return rows.map(m => {
        const u = m.avatar;
        if (!u || u.startsWith("data:")) return m;
        const inlined = map.get(u);
        return { ...m, avatar: inlined || null };
    });
}


function toCSV(messages) {
    const header = [
        'id',
        'author',
        'timestamp',
        'text',
        'edited',
        'system',
        'reactions_json',
        'attachments_json'
    ];

    const rows = (messages || []).map(m => {
        const row = [];
        const text = (m.text || '').replaceAll('\n', '\\n');
        row.push(
            m.id ?? '',
            m.author ?? '',
            m.timestamp ?? '',
            text,
            m.edited ? 'true' : 'false',
            m.system ? 'true' : 'false'
        );

        const reactions = Array.isArray(m.reactions) ? m.reactions : [];
        row.push(reactions.length ? JSON.stringify(reactions) : '');

        const attachments = Array.isArray(m.attachments) ? m.attachments : [];
        row.push(attachments.length ? JSON.stringify(attachments) : '');

        return row.map(v => `"${(v ?? '').toString().replaceAll('"', '""')}"`).join(',');
    });

    return [header.join(','), ...rows].join('\n');
}

function toHTML(rows, meta = {}) {
  const isImg = (url = "") => /\.(png|jpe?g|gif|webp)(\?|#|$)/i.test(url);
  const fmtTs = (s) => {
    if (!s) return "";
    const d = new Date(s);
    if (isNaN(d)) return s;
    return new Intl.DateTimeFormat(undefined, { dateStyle: "medium", timeStyle: "short", hour12: false }).format(d);
  };
  const relFmt = typeof Intl !== "undefined" && Intl.RelativeTimeFormat ? new Intl.RelativeTimeFormat(undefined, { numeric: "auto" }) : null;
  const relLabel = (s) => {
    if (!s || !relFmt) return "";
    const d = new Date(s);
    if (isNaN(d)) return "";
    const diffMs = Date.now() - d.getTime();
    const tense = diffMs >= 0 ? -1 : 1; // negative -> past
    const absMs = Math.abs(diffMs);
    const minute = 60 * 1000;
    const hour = 60 * minute;
    const day = 24 * hour;
    const month = 30 * day;
    const year = 365 * day;
    const choose = (value, unit) => relFmt.format(value * tense, unit);
    if (absMs < minute) return choose(Math.round(absMs / 1000) || 0, "second");
    if (absMs < hour) return choose(Math.round(absMs / minute), "minute");
    if (absMs < day) return choose(Math.round(absMs / hour), "hour");
    if (absMs < month) return choose(Math.round(absMs / day), "day");
    if (absMs < year) return choose(Math.round(absMs / month), "month");
    return choose(Math.round(absMs / year), "year");
  };
  const initials = (name="") => (name.trim().split(/\s+/).map(p=>p[0]).join("").slice(0,2) || "•");

  // --- NEW: URL helpers
  const urlRe = /https?:\/\/[^\s<>"']+/g;
  const urlsIn = (plain) => {
    const set = new Set();
    if (!plain) return set;
    let m;
    while ((m = urlRe.exec(plain)) !== null) set.add(m[0]);
    return set;
  };
  const escapeHtml = (str = "") =>
    str.replace(/&/g, "&amp;")
       .replace(/</g, "&lt;")
       .replace(/>/g, "&gt;")
       .replace(/"/g, "&quot;")
       .replace(/'/g, "&#39;");
  const autolink = (plain) => {
    const safe = escapeHtml(plain || "");
    return safe.replace(urlRe, (u) => {
      const safeUrl = escapeHtml(u);
      return `<a href="${safeUrl}" target="_blank" rel="noopener">${safeUrl}</a>`;
    });
  };

  const style = `<style>
    :root { --muted:#6b7280; --border:#e5e7eb; --bg:#ffffff; --chip:#f3f4f6; }
    body{font:14px system-ui, -apple-system, Segoe UI, Roboto; background:#fff; color:#111; padding:20px}
    h1{margin:0 0 10px 0}
    .meta{color:var(--muted); margin:0 0 12px 0}
    .toolbar{margin-bottom:12px; display:flex; gap:8px; align-items:center}
    .toolbar button{border:1px solid var(--border); background:#f9fafb; color:#111; padding:6px 10px; border-radius:6px; cursor:pointer; font:13px system-ui}
    .toolbar button:hover{background:#eef2f7}
    .msg{display:flex; gap:10px; margin:12px 0; padding:12px; border:1px solid var(--border); border-radius:12px; background:var(--bg)}
    .avt{flex:0 0 36px; width:36px; height:36px; border-radius:50%; background:#eef2f7; overflow:hidden; display:flex; align-items:center; justify-content:center; font-weight:600; color:#334155}
    .avt img{width:36px; height:36px; border-radius:50%; display:block}
    .main{flex:1}
    .hdr{color:var(--muted); font-size:12px; margin-bottom:6px}
    .hdr .rel{margin-left:6px; font-style:italic}
    .hdr .edited{font-style:italic}
    .reply{background:#f8fafc; border-left:3px solid #d1d5db; padding:8px 10px; border-radius:8px; margin:8px 0; font-size:13px; color:#374151}
    .reply .reply-meta{display:flex; flex-wrap:wrap; gap:6px; font-size:12px; color:#6b7280; margin-bottom:4px}
    .reply blockquote{margin:0; padding:0; border:none; color:#1f2937; word-wrap:break-word; overflow-wrap:anywhere}
    .atts{display:grid; grid-template-columns:repeat(auto-fill,minmax(220px,1fr)); gap:8px; margin-top:8px}
    .att, .att-img{border:1px solid var(--border); border-radius:10px; padding:8px; background:#fff; transition:max-height .2s ease, opacity .2s ease}
    .att a{word-break:break-word; overflow-wrap:anywhere; text-decoration:none}
    .att-meta{margin-top:6px; font-size:12px; color:#6b7280}
    .att-img{padding:0; overflow:hidden}
    .att-img img{display:block; width:100%; height:auto; max-height:340px; object-fit:contain; background:#fff}
    .att-img .att-meta{padding:8px}
    .reactions{margin-top:6px; font-size:12px; color:#374151}
    .divider{position:relative; text-align:center; margin:18px 0}
    .divider:before, .divider:after{content:""; position:absolute; top:50%; width:42%; height:1px; background:var(--border)}
    .divider:before{left:0} .divider:after{right:0}
    .divider span{display:inline-block; padding:0 10px; color:var(--muted); background:#fff; font-weight:600}
    .compact .msg{padding:10px}
    .compact .reactions{display:none}
    .compact .reply{display:none}
    .compact .att{max-height:0; opacity:0; pointer-events:none; padding:0; border:none; margin:0}
    .compact .att-img{max-height:0; opacity:0; pointer-events:none; padding:0; border:none; margin:0}
    .compact .atts{gap:0; margin-top:0}
    .compact .avt{display:none}
    .compact .msg{margin:8px 0; border-color:rgba(0,0,0,0.08)}
    .main > div{word-break:break-word; overflow-wrap:anywhere}
  </style>`;

  const head = `<h1>${meta.title || "Teams Chat Export"}</h1>
    <p class="meta"><b>Messages:</b> ${rows.length}${meta.timeRange ? ` &nbsp; <b>Range:</b> ${meta.timeRange}` : ""}</p>
    <div class="toolbar">
      <button type="button" data-toggle-compact>Toggle compact view</button>
    </div><hr/>`;

  const body = (rows || []).map(m => {
    // System/date rows
    if (m.system) {
      const label = (m.text || "[system]").replace(/</g,"&lt;").replace(/>/g,"&gt;");
      return `<div class="divider"><span>${label}</span></div>`;
    }

    // --- NEW: autolink text & collect URLs present in text
    const textPlain = m.text || "";
    const textHtml  = autolink(textPlain).replace(/\n/g, "<br/>");
    const urlSet    = urlsIn(textPlain);

    const reactions = (m.reactions || [])
      .map(rx => rx.reactors?.length
        ? `${rx.emoji} ${rx.count} (${rx.reactors.join(", ")})`
        : `${rx.emoji} ${rx.count}`)
      .join(" ");

    // Filter out attachment entries that are just the same naked URL already in text
    const filteredAtts = (m.attachments || []).filter(a => {
      if (!a || !a.href) return true;
      if (isImg(a.href)) return true; // keep images as thumbnails
      const label = a.label || "";
      const nakedDup = urlSet.has(a.href) && (!label || label === a.href);
      return !nakedDup;
    });

    const attHtml = filteredAtts.map(a => {
      const metaBits = [];
      if (a.type) metaBits.push(escapeHtml(a.type));
      if (a.size) metaBits.push(escapeHtml(a.size));
      if (a.owner) metaBits.push(escapeHtml(a.owner));
      if (!metaBits.length && a.metaText) metaBits.push(escapeHtml(a.metaText));
      const metaHtml = metaBits.length ? `<div class="att-meta">${metaBits.join(' • ')}</div>` : "";

      if (isImg(a.href)) {
        return `<div class="att-img"><a href="${a.href}" target="_blank" rel="noopener">
          <img src="${a.href}" alt="${escapeHtml(a.label || "image")}"/>
        </a>${metaHtml}</div>`;
      }
      const label = escapeHtml(a.label || a.href || "attachment");
      return `<div class="att"><a href="${a.href || "#"}" target="_blank" rel="noopener">${label}</a>${metaHtml}</div>`;
    }).join("");

    const reply = m.replyTo
      ? `<div class="reply">`
        + `<div class="reply-meta">↩︎ <strong>${escapeHtml(m.replyTo.author || "Unknown")}</strong>${m.replyTo.timestamp ? `<span>• ${escapeHtml(m.replyTo.timestamp)}</span>` : ""}</div>`
        + `<blockquote>${autolink(m.replyTo.text || "").replace(/\n/g, "<br/>") || "<i>(no preview)</i>"}</blockquote>`
        + `</div>`
      : "";

    const hasImg = m.avatar && m.avatar.startsWith("data:");
    const avatarEl = hasImg ? `<img src="${m.avatar}" alt="avatar"/>` : `${initials(m.author || "")}`;

    const hdrTs = fmtTs(m.timestamp);
    const relTs = relLabel(m.timestamp);
    const timeHtml = hdrTs
      ? `<span title="${m.timestamp}">${hdrTs}</span>${relTs ? `<span class="rel">(${relTs})</span>` : ""}`
      : (m.timestamp || "");
    const hdr = `${m.author || "Unknown"} — ${timeHtml}${
      m.edited ? ` <span class="edited">• edited</span>` : ""
    }${reactions ? ` — ${reactions}` : ""}`;

    return `<div class="msg">
      <div class="avt">${avatarEl}</div>
      <div class="main">
        <div class="hdr">${hdr}</div>
        ${reply}
        <div>${textHtml}</div>
        ${attHtml ? `<div class="atts">${attHtml}</div>` : ""}
      </div>
    </div>`;
  }).join("");

  const script = `<script>(function(){
    const toggleBtn = document.querySelector('[data-toggle-compact]');
    if (!toggleBtn) return;
    const root = document.body;
    const key = 'teamsExporterCompact';
    const apply = (state) => {
      if (state) {
        root.classList.add('compact');
        toggleBtn.textContent = 'Switch to expanded view';
      } else {
        root.classList.remove('compact');
        toggleBtn.textContent = 'Switch to compact view';
      }
    };
    const stored = localStorage.getItem(key);
    let compact = stored === '1';
    apply(compact);
    toggleBtn.addEventListener('click', () => {
      compact = !compact;
      apply(compact);
      try { localStorage.setItem(key, compact ? '1' : '0'); } catch (_) {}
    });
  })();</script>`;

  return `<!doctype html><meta charset="utf-8">${style}${head}${body}${script}`;
}


// Encode text to a data URL to download from SW (works reliably in MV3)
function textToDataUrl(text, mime) {
    // Avoid huge base64 for massive files; but for a few MB we’re fine.
    const b64 = btoa(unescape(encodeURIComponent(text)));
    return `data:${mime};base64,${b64}`;
}

const sendMessageToTab = (tabId, msg) => new Promise((resolve, reject) => {
    chrome.tabs.sendMessage(tabId, msg, (resp) => {
        const err = chrome.runtime.lastError;
        if (err) {
            reject(new Error(err.message || 'Failed to reach tab context'));
            return;
        }
        resolve(resp);
    });
});

async function ensureContentScript(tabId) {
    try {
        const pong = await sendMessageToTab(tabId, { type: 'PING' });
        if (pong?.ok) return;
    } catch (_) {
        // fallback to injection
    }
    await chrome.scripting.executeScript({ target: { tabId, allFrames: true }, files: ['content.js'] });
    const pong2 = await sendMessageToTab(tabId, { type: 'PING' });
    if (!pong2?.ok) throw new Error('Content script did not respond after injection');
}

async function requestScrape(tabId, options) {
    const res = await sendMessageToTab(tabId, { type: 'SCRAPE_TEAMS', options });
    if (!res) throw new Error('No response from content script');
    if (res.error) throw new Error(res.error);
    return res;
}

async function buildAndDownload({ messages = [], meta = {}, format = 'json', saveAs = true, embedAvatars = false }) {
    let rows = messages;
    if (format === 'html' && embedAvatars) {
        try {
            rows = await embedAvatarsInRows(messages);
        } catch (e) {
            log('embed avatars failed', e?.message || e);
            rows = messages;
        }
    }

    const rangeLabel = formatRangeLabel(meta.startAt, meta.endAt);
    const enrichedMeta = { ...meta };
    if (rangeLabel) enrichedMeta.timeRange = rangeLabel;

    const baseTitle = sanitizeBase(enrichedMeta.title || 'teams-chat');
    const stamp = new Date().toISOString().replace(/[:.]/g, '-');
    const base = `${baseTitle}-${stamp}`;
    let filename, mime, content;

    if (format === 'json') {
        filename = `${base}.json`;
        mime = 'application/json';
        const payload = { meta: { ...enrichedMeta, count: messages.length }, messages };
        content = JSON.stringify(payload, null, 2);
    } else if (format === 'csv') {
        filename = `${base}.csv`;
        mime = 'text/csv';
        content = toCSV(messages);
    } else if (format === 'html') {
        filename = `${base}.html`;
        mime = 'text/html';
        content = toHTML(rows, { ...enrichedMeta, count: messages.length });
    } else {
        throw new Error('Unknown format: ' + format);
    }

    const url = textToDataUrl(content, mime);
    log('download build', { format, bytes: content.length, filename });

    try {
        const id = await chrome.downloads.download({ url, filename, saveAs });
        log('download started', { id, filename });
        return { ok: true, filename, id };
    } catch (e) {
        log('download primary failed', e?.message || e);
        const safe = `${sanitizeBase('teams-chat')}-${Date.now()}.${format === 'html' ? 'html' : format === 'csv' ? 'csv' : 'json'}`;
        try {
            const id2 = await chrome.downloads.download({ url, filename: safe, saveAs });
            return { ok: true, filename: safe, id: id2 };
        } catch (e2) {
            log('download fallback failed', e2?.message || e2);
            throw new Error(e2?.message || String(e2));
        }
    }
}

function broadcastStatus(payload) {
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
    chrome.runtime.sendMessage({ type: 'EXPORT_STATUS', ...enriched }).catch(() => { });
    updateBadgeForStatus(payload);
}

function handleBuildAndDownloadMessage(msg, sendResponse) {
    (async () => {
        try {
            const result = await buildAndDownload(msg.data || {});
            sendResponse(result);
        } catch (err) {
            sendResponse({ error: err?.message || String(err) });
        }
    })();
}

function handleStartExportMessage(msg, sendResponse) {
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
            const ctx = await sendMessageToTab(tabId, { type: 'CHECK_CHAT_CONTEXT' });
            if (!ctx?.ok) {
                const message = ctx?.reason || 'Open a chat conversation before exporting.';
                sendResponse({ error: message });
                return;
            }

            startedAt = Date.now();
            updateActiveExport(tabId, { startedAt, phase: 'starting', lastStatus: null });
            broadcastStatus({ tabId, phase: 'starting', startedAt });
            setBadge('0', '#2563eb');

            broadcastStatus({ tabId, phase: 'scrape:start' });
            const scrapeRes = await requestScrape(tabId, scrapeOptions);
            const totalMessages = Array.isArray(scrapeRes.messages) ? scrapeRes.messages.length : 0;
            broadcastStatus({ tabId, phase: 'scrape:complete', messages: totalMessages });
            if (totalMessages === 0) {
                const message = 'No messages found for the selected range.';
                broadcastStatus({ tabId, phase: 'empty', message });
                setBadge('0', '#6b7280');
                clearBadgeSoon(2000);
                sendResponse({ empty: true, message, code: 'EMPTY_RESULTS' });
                return;
            }
            }

            const buildRes = await buildAndDownload({
                messages: scrapeRes.messages || [],
                meta: scrapeRes.meta || {},
                format: buildOptions.format || 'json',
                saveAs: buildOptions.saveAs !== false,
                embedAvatars: Boolean(buildOptions.embedAvatars)
            });

            broadcastStatus({ tabId, phase: 'complete', filename: buildRes.filename });
            sendResponse({ ok: true, filename: buildRes.filename, downloadId: buildRes.id });
        } catch (err) {
            const message = err?.message || String(err);
            broadcastStatus({ tabId, phase: 'error', error: message });
            sendResponse({ error: message });
        } finally {
            if (startedAt) {
                activeExports.delete(tabId);
            }
        }
    })();
}

const BADGE_COLOR_EMPTY = '#6b7280';

function updateBadgeForStatus(payload) {
    try {
        const phase = payload?.phase;
        if (phase === 'scrape:complete') {
            const total = payload?.messages ?? payload?.messagesExtracted;
            if (typeof total === 'number') setBadge(String(total));
        } else if (phase === 'empty') {
            setBadge('0', BADGE_COLOR_EMPTY);
            clearBadgeSoon(2000);
        } else if (phase === 'complete') {
            setBadge('✔', '#16a34a');
            clearBadgeSoon(2000);
        } else if (phase === 'error') {
            setBadge('!', '#dc2626');
            clearBadgeSoon(3000);
        } else if (phase === 'scrape:start') {
            setBadge('…', '#2563eb');
        }
    } catch (_) {
        // ignore badge errors
    }
}

function updateBadgeForProgress(progress) {
    if (!progress) return;
    const seen = progress.filteredSeen ?? progress.seen ?? progress.aggregated ?? progress.messagesVisible;
    if (typeof seen === 'number' && seen >= 0) {
        setBadge(String(seen));
    }
}

function setBadge(text, color = '#1d4ed8') {
    try {
        chrome.action.setBadgeBackgroundColor({ color });
        chrome.action.setBadgeText({ text: text || '' });
    } catch (_) {
        // ignore
    }
}

function clearBadge() {
    try {
        chrome.action.setBadgeText({ text: '' });
    } catch (_) {
        // ignore
    }
}

let clearBadgeTimer = null;
function clearBadgeSoon(delay = 0) {
    if (clearBadgeTimer) clearTimeout(clearBadgeTimer);
    clearBadgeTimer = setTimeout(() => {
        clearBadge();
        clearBadgeTimer = null;
    }, delay);
}

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
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

    if (msg.type === 'SCRAPE_PROGRESS') {
        updateBadgeForProgress(msg.payload || msg);
        return;
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
