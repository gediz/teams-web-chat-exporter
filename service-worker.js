// ===== service-worker.js (replace) =====
function log(...a) { try { console.log("[Teams Exporter SW]", ...a) } catch { } }
log("boot");

chrome.runtime.onInstalled.addListener(() => log("onInstalled"));
chrome.runtime.onStartup?.addListener(() => log("onStartup"));

function sanitizeBase(name) {
    const raw = (name || "teams-chat").toString();
    const cleaned = raw.replace(/[<>:"/\\|?*\x00-\x1F]/g, "-").replace(/\s+/g, " ").trim().replace(/[. ]+$/g, "");
    return (cleaned || "teams-chat").slice(0, 80);
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
    const header = ['id', 'author', 'timestamp', 'text', 'reactions', 'attachments'].join(',');
    const rows = (messages || []).map(m => {
        const reactions = (m.reactions || []).map(r => r.reactors?.length ? `${r.emoji} ${r.count} [${r.reactors.join('; ')}]` : `${r.emoji} ${r.count}`).join('|');
        const attachments = (m.attachments || []).map(a => (a.label ? `${a.label}=>${a.href || ''}` : (a.href || ''))).join('|');
        const flat = [
            m.id,
            m.author,
            m.timestamp,
            (m.text || '').replaceAll('\n', '\\n'),
            reactions,
            attachments
        ];
        return flat.map(v => `"${(v ?? '').toString().replaceAll('"', '""')}"`).join(',');
    });
    return [header, ...rows].join('\n');
}

function toMD(rows, meta = {}) {
    const head = `# ${meta.title || "Teams Chat Export"}\n\n_Messages_: ${rows.length}\n${meta.timeRange ? `_Time range_: ${meta.timeRange}\n` : ``}\n---\n`;
    const body = (rows || []).map(m => {
        const r = (m.reactions || []).join(' ');
        const at = (m.attachments || []).map(a => `\n  - [${a.label || a.href}](${a.href || "#"})`).join('');
        return `- **${m.author || 'Unknown'}** — _${m.timestamp || ''}_ ${r}\n\n  ${m.text || ''}${at}\n`;
    }).join('\n');
    return head + body;
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
  const autolink = (plain) =>
    (plain || "").replace(urlRe, (u) => `<a href="${u}" target="_blank" rel="noopener">${u}</a>`);

  const style = `<style>
    :root { --muted:#6b7280; --border:#e5e7eb; --bg:#ffffff; --chip:#f3f4f6; }
    body{font:14px system-ui, -apple-system, Segoe UI, Roboto; background:#fff; color:#111; padding:20px}
    h1{margin:0 0 10px 0}
    .meta{color:var(--muted); margin:0 0 12px 0}
    .msg{display:flex; gap:10px; margin:12px 0; padding:12px; border:1px solid var(--border); border-radius:12px; background:var(--bg)}
    .avt{flex:0 0 36px; width:36px; height:36px; border-radius:50%; background:#eef2f7; overflow:hidden; display:flex; align-items:center; justify-content:center; font-weight:600; color:#334155}
    .avt img{width:36px; height:36px; border-radius:50%; display:block}
    .main{flex:1}
    .hdr{color:var(--muted); font-size:12px; margin-bottom:6px}
    .hdr .rel{margin-left:6px; font-style:italic}
    .hdr .edited{font-style:italic}
    .reply{background:#f8fafc; border-left:3px solid #e5e7eb; padding:6px 8px; border-radius:8px; margin:6px 0; font-size:13px; color:#374151}
    .atts{display:grid; grid-template-columns:repeat(auto-fill,minmax(220px,1fr)); gap:8px; margin-top:8px}
    .att, .att-img{border:1px solid var(--border); border-radius:10px; padding:8px; background:#fff}
    .att a{word-break:break-all; text-decoration:none}
    .att-img{padding:0; overflow:hidden}
    .att-img img{display:block; width:100%; height:auto; max-height:340px; object-fit:contain; background:#fff}
    .reactions{margin-top:6px; font-size:12px; color:#374151}
    .divider{position:relative; text-align:center; margin:18px 0}
    .divider:before, .divider:after{content:""; position:absolute; top:50%; width:42%; height:1px; background:var(--border)}
    .divider:before{left:0} .divider:after{right:0}
    .divider span{display:inline-block; padding:0 10px; color:var(--muted); background:#fff; font-weight:600}
  </style>`;

  const head = `<h1>${meta.title || "Teams Chat Export"}</h1>
    <p class="meta"><b>Messages:</b> ${rows.length}${meta.timeRange ? ` &nbsp; <b>Range:</b> ${meta.timeRange}` : ""}</p><hr/>`;

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
      if (isImg(a.href)) {
        return `<div class="att-img"><a href="${a.href}" target="_blank" rel="noopener">
          <img src="${a.href}" alt="${(a.label || "image").replace(/"/g, "&quot;")}"/>
        </a></div>`;
      }
      const label = (a.label || a.href || "attachment").replace(/</g,"&lt;").replace(/>/g,"&gt;");
      return `<div class="att"><a href="${a.href || "#"}" target="_blank" rel="noopener">${label}</a></div>`;
    }).join("");

    const reply = m.replyTo
      ? `<div class="reply">↩︎ <b>${m.replyTo.author || "Unknown"}</b>${m.replyTo.timestamp ? ` — <i>${m.replyTo.timestamp}</i>` : ""}<br>${(m.replyTo.text || "").replace(/\n/g, "<br/>")}</div>`
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

  return `<!doctype html><meta charset="utf-8">${style}${head}${body}`;
}


// Encode text to a data URL to download from SW (works reliably in MV3)
function textToDataUrl(text, mime) {
    // Avoid huge base64 for massive files; but for a few MB we’re fine.
    const b64 = btoa(unescape(encodeURIComponent(text)));
    return `data:${mime};base64,${b64}`;
}

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
    (async () => {
        if (!msg || !msg.type) return;

        if (msg.type === "PING_SW") { sendResponse({ ok: true, now: Date.now() }); return; }

        if (msg.type === "BUILD_AND_DOWNLOAD") {
            const { messages = [], meta = {}, format = "json", saveAs = true, embedAvatars = false } = msg.data || {};

            let rows = messages;
            if (format === "html" && embedAvatars) {
                try {
                    rows = await embedAvatarsInRows(messages);
                } catch (_) {
                    rows = messages; // fall back silently
                }
            }
            const base = `${sanitizeBase(meta.title)}-${new Date().toISOString().replace(/[:.]/g, "-")}`;
            let filename, mime, content;

            if (format === "json") { filename = `${base}.json`; mime = "application/json"; content = JSON.stringify({ meta, messages }, null, 2); }
            else if (format === "csv") { filename = `${base}.csv`; mime = "text/csv"; content = toCSV(messages); }
            else if (format === "md") { filename = `${base}.md`; mime = "text/markdown"; content = toMD(messages, meta); }
            else if (format === "html") { filename = `${base}.html`; mime="text/html"; content = toHTML(rows, meta); }
            else if (format === "ndjson") { filename = `${base}.ndjson`; mime = "application/x-ndjson"; content = (messages || []).map(m => JSON.stringify(m)).join("\n"); }
            else { sendResponse({ error: "Unknown format: " + format }); return; }

            const url = textToDataUrl(content, mime);
            log("BUILD_AND_DOWNLOAD", { format, bytes: content.length, filename });

            try {
                const id = await chrome.downloads.download({ url, filename, saveAs });
                log("download started", { id, filename });
                sendResponse({ ok: true, filename, id });
            } catch (e) {
                log("download error", e?.message || String(e));
                // Try with an ultra-safe fallback name
                const safe = `${sanitizeBase("teams-chat")}-${Date.now()}.${format === "ndjson" ? "ndjson" : format === "md" ? "md" : format === "html" ? "html" : format === "csv" ? "csv" : "json"}`;
                try {
                    const id2 = await chrome.downloads.download({ url, filename: safe, saveAs });
                    sendResponse({ ok: true, filename: safe, id: id2 });
                } catch (e2) {
                    sendResponse({ error: e2?.message || String(e2) });
                }
            }
            return;
        }
    })();
    return true; // keep SW alive for async sendResponse
});
