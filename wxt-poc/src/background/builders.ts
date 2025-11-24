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
  const isImg = (url = '') => /\.(png|jpe?g|gif|webp)(\?|#|$)/i.test(url);
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
  const initials = (name = '') => (name.trim().split(/\s+/).map(p => p[0]).join('').slice(0, 2) || '•');

  const urlRe = /https?:\/\/[^\s<>"']+/g;
  const escapeHtml = (str = '') =>
    str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  const autolink = (plain: string) => {
    const safe = escapeHtml(plain || '');
    return safe.replace(urlRe, u => {
      const safeUrl = escapeHtml(u);
      return `<a href="${safeUrl}" target="_blank" rel="noopener">${safeUrl}</a>`;
    });
  };

  const style = `<style>
    :root { --muted:#6b7280; --border:#e5e7eb; --bg:#ffffff; --chip:#f3f4f6; }
    body{font:14px system-ui, -apple-system, Segoe UI, Roboto; background:#fff; color:#111; padding:24px; margin:0}
    h1{margin:0 0 10px 0}
    .meta{color:var(--muted); margin:0 0 12px 0}
    .toolbar{margin:0 0 12px 0; display:flex; gap:8px; align-items:center}
    .toolbar button{border:1px solid var(--border); background:#f9fafb; color:#111; padding:6px 10px; border-radius:6px; cursor:pointer; font:13px system-ui}
    .toolbar button:hover{background:#eef2f7}
    .msg{display:flex; gap:10px; margin:12px 0; padding:12px; border:1px solid var(--border); border-radius:12px; background:var(--bg)}
    .msg.system{background:#f8fafc; border-style:dashed; color:#374151}
    body.compact .msg{padding:10px; margin:8px 0}
    body.compact .header{flex-wrap:wrap}
    .avatar{width:36px; height:36px; border-radius:50%; background:var(--chip); display:flex; align-items:center; justify-content:center; font-weight:600; color:#111; overflow:hidden}
    .avatar img{width:36px; height:36px; border-radius:50%; object-fit:cover}
    .avatar.system{background:#e5e7eb; color:#4b5563}
    .content{flex:1; min-width:0}
    .header{display:flex; align-items:center; gap:8px; font-weight:600}
    .author{font-weight:600}
    .timestamp{color:var(--muted); font-size:12px}
    .meta-line{color:var(--muted); font-size:12px; margin-top:4px}
    .text{white-space:pre-wrap; margin:8px 0}
    .chip{display:inline-flex; align-items:center; gap:4px; background:var(--chip); color:#111; border-radius:999px; padding:4px 8px; font-size:12px}
    .row{display:flex; flex-wrap:wrap; gap:6px; margin-top:6px}
    .attachments{margin-top:8px; display:flex; flex-direction:column; gap:4px; font-size:13px}
    .attachment a{color:#1d4ed8; text-decoration:none}
    .attachment a:hover{text-decoration:underline}
    .reactions{margin-top:6px; display:flex; gap:6px; flex-wrap:wrap}
    .reaction{background:var(--chip); padding:4px 8px; border-radius:12px; font-size:12px}
  </style>`;

  const metaLines = [];
  if (meta.title) metaLines.push(`<div><strong>Title:</strong> ${escapeHtml(meta.title || '')}</div>`);
  if (meta.timeRange) metaLines.push(`<div><strong>Range:</strong> ${escapeHtml(meta.timeRange || '')}</div>`);
  if (meta.count != null) metaLines.push(`<div><strong>Messages:</strong> ${escapeHtml(String(meta.count))}</div>`);
  if (meta.startAt) metaLines.push(`<div><strong>Start:</strong> ${fmtTs(meta.startAt)}</div>`);
  if (meta.endAt) metaLines.push(`<div><strong>End:</strong> ${fmtTs(meta.endAt)}</div>`);

  const head = `
    <h1>Teams Chat Export</h1>
    <div class="meta">${metaLines.join(' ')}<div>Generated ${fmtTs(Date.now())}</div></div>
    <div class="toolbar">
      <button type="button" data-toggle-compact>Toggle compact view</button>
    </div>
  `;

  const body = rows
    .map((m, idx) => {
      const ts = m.timestamp || '';
      const rel = relLabel(ts);
      const tsLabel = fmtTs(ts);
      const text = autolink(m.text || '');
      const reactions = Array.isArray(m.reactions) ? m.reactions : [];
      const atts = Array.isArray(m.attachments) ? m.attachments : [];
      const replyTo = m.replyTo;
      const avatar = m.avatar && isImg(m.avatar) ? `<img src="${m.avatar}" alt="avatar" />` : `<span>${initials(m.author || '')}</span>`;

      const reactHtml = reactions
        .map(r => `<div class="reaction">${escapeHtml(r.emoji || '')} ${r.count}${r.reactors ? ` · ${escapeHtml(r.reactors.join(', '))}` : ''}</div>`)
        .join('');

      const attHtml = atts
        .map(att => {
          const label = escapeHtml(att.label || att.href || 'attachment');
          const href = att.href ? escapeHtml(att.href) : '';
          const metaText = att.metaText ? ` — ${escapeHtml(att.metaText)}` : '';
          const type = att.type ? ` [${escapeHtml(att.type)}]` : '';
          const size = att.size ? ` (${escapeHtml(att.size)})` : '';
          const owner = att.owner ? ` — ${escapeHtml(att.owner)}` : '';
          const link = href ? `<a href="${href}" target="_blank" rel="noopener">${label}</a>` : label;
          return `<div class="attachment">• ${link}${type}${size}${owner}${metaText}</div>`;
        })
        .join('');

      const replyHtml = replyTo
        ? `<div class="meta-line">Replying to ${escapeHtml(replyTo.author || '')}${replyTo.timestamp ? ` • ${escapeHtml(replyTo.timestamp)}` : ''}</div>
           <div class="meta-line">${escapeHtml((replyTo.text || '').slice(0, 300))}</div>`
        : '';

      const systemClass = m.system ? ' system' : '';
      return `<div class="msg${systemClass}" id="msg-${idx}">
        <div class="avatar${systemClass}">${avatar}</div>
        <div class="content">
          <div class="header">
            <span class="author">${escapeHtml(m.author || '')}${m.system ? ' [system]' : ''}</span>
            <span class="timestamp">${tsLabel}${rel ? ` • ${rel}` : ''}</span>
            ${m.edited ? `<span class="meta-line">(edited)</span>` : ''}
            ${m.system ? `<span class="meta-line">[system]</span>` : ''}
          </div>
          ${replyHtml}
          <div class="text">${text || '<span class="meta-line">(no text)</span>'}</div>
          ${reactHtml ? `<div class="reactions">${reactHtml}</div>` : ''}
          ${attHtml ? `<div class="attachments">${attHtml}</div>` : ''}
        </div>
      </div>`;
    })
    .join('\n');

  const script = `<script>(()=>{const btn=document.querySelector('[data-toggle-compact]');if(!btn)return;const key='teamsExporterCompact';const apply=(c)=>{document.body.classList.toggle('compact',c);btn.textContent=c?'Switch to expanded view':'Switch to compact view';};const stored=localStorage.getItem(key);let compact=stored==='1';apply(compact);btn.addEventListener('click',()=>{compact=!compact;apply(compact);try{localStorage.setItem(key,compact?'1':'0');}catch(_){}});})();</script>`;

  return `<!doctype html><html><head><meta charset="utf-8">${style}</head><body>${head}${body}${script}</body></html>`;
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
