import type { ExportMessage, ExportMeta } from '../types/shared';

export async function embedAvatarsInRows(rows: ExportMessage[]) {
  const map = new Map(); // url -> dataURL|null (if failed)
  for (const m of rows) {
    const u = m.avatar;
    if (!u || u.startsWith('data:')) continue;
    if (!map.has(u)) {
      try {
        map.set(u, await fetchAsDataURL(u));
      } catch {
        map.set(u, null);
      }
    }
  }
  return rows.map(m => {
    const u = m.avatar;
    if (!u || u.startsWith('data:')) return m;
    const inlined = map.get(u);
    return { ...m, avatar: inlined || null };
  });
}

async function fetchAsDataURL(url: string) {
  const res = await fetch(url, { credentials: 'include' });
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  const buf = await res.arrayBuffer();
  const bytes = new Uint8Array(buf);
  let bin = '';
  for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
  const b64 = btoa(bin);
  const ct = res.headers.get('content-type') || 'image/png';
  return `data:${ct};base64,${b64}`;
}

export function toCSV(messages: ExportMessage[]) {
  const header = ['id', 'author', 'timestamp', 'text', 'edited', 'system', 'reactions_json', 'attachments_json'];

  const rows = (messages || []).map(m => {
    const row = [];
    const text = (m.text || '').replace(/\n/g, '\\n');
    row.push(m.id ?? '', m.author ?? '', m.timestamp ?? '', text, m.edited ? 'true' : 'false', m.system ? 'true' : 'false');

    const reactions = Array.isArray(m.reactions) ? m.reactions : [];
    row.push(reactions.length ? JSON.stringify(reactions) : '');

    const attachments = Array.isArray(m.attachments) ? m.attachments : [];
    row.push(attachments.length ? JSON.stringify(attachments) : '');

    return row.map(v => `"${(v ?? '').toString().split('"').join('""')}"`).join(',');
  });

  return [header.join(','), ...rows].join('\n');
}

export function toHTML(rows: ExportMessage[], meta: ExportMeta = {}): string {
  // Restore the richer HTML layout (avatars, replies, attachment grid, divider, compact mode)
  const fmtTs = (s: string | number) => {
    if (!s) return '';
    const d = new Date(s);
    if (Number.isNaN(d.getTime())) return s as string;
    return new Intl.DateTimeFormat(undefined, { dateStyle: 'medium', timeStyle: 'short', hour12: false }).format(d);
  };
  const relFmt = typeof Intl !== 'undefined' && Intl.RelativeTimeFormat ? new Intl.RelativeTimeFormat(undefined, { numeric: 'auto' }) : null;
  const relLabel = (s: string | number) => {
    if (!s || !relFmt) return '';
    const d = new Date(s);
    if (Number.isNaN(d.getTime())) return '';
    const diffMs = Date.now() - d.getTime();
    const tense = diffMs >= 0 ? -1 : 1;
    const absMs = Math.abs(diffMs);
    const minute = 60 * 1000;
    const hour = 60 * minute;
    const day = 24 * hour;
    const month = 30 * day;
    const year = 365 * day;
    const choose = (value: number, unit: Intl.RelativeTimeFormatUnit) => relFmt.format(value * tense, unit);
    if (absMs < minute) return choose(Math.round(absMs / 1000) || 0, 'second');
    if (absMs < hour) return choose(Math.round(absMs / minute), 'minute');
    if (absMs < day) return choose(Math.round(absMs / hour), 'hour');
    if (absMs < month) return choose(Math.round(absMs / day), 'day');
    if (absMs < year) return choose(Math.round(absMs / month), 'month');
    return choose(Math.round(absMs / year), 'year');
  };
  const escapeHtml = (str = '') =>
    str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  const autolink = (plain: string) => {
    const safe = escapeHtml(plain || '');
    return safe.replace(/https?:\/\/[^\s<>"']+/g, u => `<a href="${escapeHtml(u)}" target="_blank" rel="noopener">${escapeHtml(u)}</a>`);
  };
  const initials = (name = '') => (name.trim().split(/\s+/).map(p => p[0]).join('').slice(0, 2) || '•');

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

  const metaParts = [];
  if (meta.messages != null || meta.count != null) metaParts.push(`<b>Messages:</b> ${escapeHtml(String(meta.messages ?? meta.count ?? ''))}`);
  if (meta.timeRange) metaParts.push(`<b>Range:</b> ${escapeHtml(meta.timeRange)}`);
  const metaLine = metaParts.length ? `<p class="meta">${metaParts.join(' &nbsp; ')}</p>` : '';

  const head = `<h1>${escapeHtml(meta.title || 'Teams Chat Export')}</h1>
    ${metaLine}
    <div class="toolbar"><button type="button" data-toggle-compact>Toggle compact view</button></div><hr/>`;

  const body = rows
    .map((m, idx) => {
      if (m.system) {
        const label = escapeHtml(m.text || m.author || '[system]');
        return `<div class="divider"><span>${label}</span></div>`;
      }
      const ts = m.timestamp || '';
      const rel = relLabel(ts);
      const tsLabel = fmtTs(ts);
      const reactions = Array.isArray(m.reactions) ? m.reactions : [];
      const atts = Array.isArray(m.attachments) ? m.attachments : [];
      const replyTo = m.replyTo;
      const text = autolink(m.text || '');
      const avatar = m.avatar
        ? `<img src="${escapeHtml(m.avatar)}" alt="avatar" />`
        : escapeHtml((m.author || '').split(' ').map(p => p[0]).join('').slice(0, 2) || '•');

      const reactHtml = reactions
        .map(r => `<span class="chip">${escapeHtml(r.emoji || '')} ${r.count}${r.reactors ? ` · ${escapeHtml(r.reactors.join(', '))}` : ''}</span>`)
        .join(' ');

      const attsHtml = atts
        .map(att => {
          const label = escapeHtml(att.label || att.href || 'attachment');
          const href = att.href ? escapeHtml(att.href) : '';
          const metaText = att.metaText ? `<div class="att-meta">${escapeHtml(att.metaText)}</div>` : '';
          const type = att.type ? ` [${escapeHtml(att.type)}]` : '';
          const size = att.size ? ` (${escapeHtml(att.size)})` : '';
          const owner = att.owner ? ` — ${escapeHtml(att.owner)}` : '';
          const isImage = att.href && /\.(png|jpe?g|gif|webp)(\?|#|$)/i.test(att.href);
          if (isImage) {
            return `<div class="att-img"><img src="${href}" alt="${label}" />${metaText}</div>`;
          }
          const link = href ? `<a href="${href}" target="_blank" rel="noopener">${label}</a>` : label;
          return `<div class="att">${link}${type}${size}${owner}${metaText}</div>`;
        })
        .join('');

      const replyHtml = replyTo
        ? `<div class="reply"><div class="reply-meta">↩︎ <strong>${escapeHtml(replyTo.author || '')}</strong>${replyTo.timestamp ? `<span>• ${escapeHtml(replyTo.timestamp)}</span>` : ''}</div><blockquote>${escapeHtml(replyTo.text || '')}</blockquote></div>`
        : '';

      return `<div class="msg" id="msg-${idx}">
      <div class="avt">${avatar}</div>
      <div class="main">
        <div class="hdr">${escapeHtml(m.author || '')} — <span title="${escapeHtml(ts)}">${tsLabel}</span>${rel ? `<span class="rel">(${rel})</span>` : ''}${m.edited ? ' <span class="edited">• edited</span>' : ''}</div>
        ${replyHtml}
        <div>${text}</div>
        ${reactHtml ? `<div class="reactions">${reactHtml}</div>` : ''}
        ${attsHtml ? `<div class="atts">${attsHtml}</div>` : ''}
      </div>
    </div>`;
    })
    .join('');

  const script = `<script>(()=>{const btn=document.querySelector('[data-toggle-compact]');if(!btn)return;const key='teamsExporterCompact';const apply=(c)=>{document.body.classList.toggle('compact',c);btn.textContent=c?'Switch to expanded view':'Switch to compact view';};const stored=localStorage.getItem(key);let compact=stored==='1';apply(compact);btn.addEventListener('click',()=>{compact=!compact;apply(compact);try{localStorage.setItem(key,compact?'1':'0');}catch(_){}});})();</script>`;

  return `<!doctype html><meta charset="utf-8">${style}${head}${body}${script}`;
}

// Encode text to a data URL to download from SW (works reliably in MV3)
export function textToDataUrl(text: string, mime: string) {
  const b64 = btoa(unescape(encodeURIComponent(text)));
  return `data:${mime};base64,${b64}`;
}

// Firefox-compatible: Create blob URL (Firefox blocks data URLs in downloads)
export function textToBlobUrl(text: string, mime: string) {
  const blob = new Blob([text], { type: mime });
  return URL.createObjectURL(blob);
}
