import type { ExportMessage, ExportMeta } from '../types/shared';

/**
 * Converts avatar HTTP URLs to base64 data URLs for offline viewing.
 * Returns both the converted messages and a map of unique avatars.
 */
export async function embedAvatarsInRows(rows: ExportMessage[]) {
  const map = new Map<string, string | null>(); // url -> dataURL|null (if failed)
  for (const m of rows) {
    const u = m.avatar;
    if (!u || u.startsWith('data:')) continue;
    if (!map.has(u)) {
      try {
        const dataUrl = await fetchAsDataURL(u);
        map.set(u, dataUrl);
      } catch (err) {
        console.error(`[Avatar Fetch] FAILED for ${u.substring(0, 100)}...`, err);
        map.set(u, null);
      }
    }
  }
  const converted = rows.map(m => {
    const u = m.avatar;
    if (!u || u.startsWith('data:')) return m;
    const inlined = map.get(u);
    return { ...m, avatar: inlined || null };
  });
  return { messages: converted, avatarMap: map };
}

/**
 * Extracts a stable user ID from a Teams avatar URL.
 * E.g., "https://.../8:orgid:cf7134d2-b5df-4b93-bbeb-e68d4545bb89/..." -> "8orgid-cf7134d2"
 */
export function extractAvatarId(url: string): string {
  const match = url.match(/\/([^/]+)\/profilepicturev2/);
  if (match) {
    const fullId = match[1];
    // Extract the first part of the UUID for a shorter ID
    const parts = fullId.split(':');
    if (parts.length >= 3) {
      const prefix = parts[1]; // "orgid"
      const uuid = parts[2].split('-')[0]; // First part of UUID
      return `${parts[0]}${prefix}-${uuid}`;
    }
  }
  // Fallback: use a hash of the URL
  let hash = 0;
  for (let i = 0; i < url.length; i++) {
    const char = url.charCodeAt(i);
    hash = ((hash << 5) - hash) + char;
    hash = hash & hash; // Convert to 32-bit integer
  }
  return `avatar-${Math.abs(hash).toString(36)}`;
}

/**
 * Normalizes avatars by moving them to a meta.avatars map with IDs.
 * Returns messages with avatarId instead of avatar URL.
 */
export function normalizeAvatars(messages: ExportMessage[], avatarMap: Map<string, string | null>) {
  const avatars: Record<string, string> = {};
  const urlToId = new Map<string, string>();

  // Build avatars map with stable IDs
  avatarMap.forEach((dataUrl, url) => {
    if (dataUrl) {
      const id = extractAvatarId(url);
      avatars[id] = dataUrl;
      urlToId.set(url, id);
    }
  });

  // Replace avatar URLs with avatarId references
  const normalized = messages.map(m => {
    if (!m.avatar) return m;
    const id = urlToId.get(m.avatar);
    if (id) {
      const { avatar, ...rest } = m;
      return { ...rest, avatarId: id };
    }
    // If avatar failed to convert, remove it
    const { avatar, ...rest } = m;
    return rest;
  });

  return { messages: normalized, avatars };
}

/**
 * Removes avatar data entirely from messages.
 */
export function removeAvatars(messages: ExportMessage[]) {
  return messages.map(m => {
    if (!m.avatar) return m;
    const { avatar, ...rest } = m;
    return rest;
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
  const urlRe = /https?:\/\/[^\s<>"']+/g;
  const autolinkEscaped = (escaped: string) =>
    escaped.replace(urlRe, u => `<a href="${escapeHtml(u)}" target="_blank" rel="noopener">${escapeHtml(u)}</a>`);
  const formatInline = (segment: string) => {
    const parts = segment.split('`');
    let html = '';
    if (parts.length >= 3 && parts.length % 2 === 1) {
      html = parts
        .map((part, idx) => {
          const escaped = escapeHtml(part);
          if (idx % 2 === 1) return `<code>${escaped}</code>`;
          return autolinkEscaped(escaped);
        })
        .join('');
    } else {
      const escaped = escapeHtml(segment);
      html = autolinkEscaped(escaped);
    }
    return html
      .replace(/\r\n/g, '\n')
      .replace(/\n{2,}/g, '<br>&nbsp;<br>')
      .replace(/\n/g, '<br>');
  };
  const formatWithQuotes = (segment: string) => {
    const lines = segment.split(/\r?\n/);
    let out = '';
    let mode: 'normal' | 'quote' = 'normal';
    let buf: string[] = [];
    const flush = () => {
      if (!buf.length) return;
      const text = buf.join('\n');
      if (mode === 'quote') {
        out += `<blockquote>${formatInline(text)}</blockquote>`;
      } else {
        out += formatInline(text);
      }
      buf = [];
    };
    for (const line of lines) {
      const isQuote = /^>\s?/.test(line);
      const cleaned = isQuote ? line.replace(/^>\s?/, '') : line;
      if (isQuote && mode !== 'quote') {
        flush();
        mode = 'quote';
      } else if (!isQuote && mode === 'quote') {
        flush();
        mode = 'normal';
      }
      buf.push(cleaned);
    }
    flush();
    return out;
  };
  const formatText = (plain: string) => {
    const raw = plain || '';
    const fenceParts = raw.split('```');
    if (fenceParts.length >= 3 && fenceParts.length % 2 === 1) {
      return fenceParts
        .map((part, idx) => {
          if (idx % 2 === 1) {
            const code = part.replace(/^\n/, '').replace(/\n$/, '');
            return `<pre class="code-block"><code>${escapeHtml(code)}</code></pre>`;
          }
          return formatWithQuotes(part);
        })
        .join('');
    }
    return formatWithQuotes(raw);
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
    code{font-family:ui-monospace,SFMono-Regular,Menlo,Consolas,"Liberation Mono",monospace;background:#f3f4f6;border:1px solid #e5e7eb;padding:1px 4px;border-radius:4px}
    pre.code-block{background:#0b1020;color:#e5e7eb;border-radius:10px;padding:10px 12px;overflow:auto;margin:8px 0;border:1px solid #111827}
    pre.code-block code{background:none;border:none;padding:0;color:inherit;white-space:pre}
    .hdr{color:var(--muted); font-size:12px; margin-bottom:6px}
    .hdr .rel{margin-left:6px; font-style:italic}
    .hdr .edited{font-style:italic}
    .reply{background:#f8fafc; border-left:3px solid #d1d5db; padding:8px 10px; border-radius:8px; margin:8px 0; font-size:13px; color:#374151}
    .reply .reply-meta{display:flex; flex-wrap:wrap; gap:6px; font-size:12px; color:#6b7280; margin-bottom:4px}
    .reply blockquote{margin:0; padding:0; border:none; color:#1f2937; word-wrap:break-word; overflow-wrap:anywhere}
    blockquote{margin:8px 0; padding:8px 10px; border-left:3px solid #d1d5db; background:#f8fafc; color:#374151}
    .atts{display:grid; grid-template-columns:repeat(auto-fill,minmax(220px,1fr)); gap:8px; margin-top:8px}
    .att, .att-img{border:1px solid var(--border); border-radius:10px; padding:8px; background:#fff; transition:max-height .2s ease, opacity .2s ease}
    .att a{word-break:break-word; overflow-wrap:anywhere; text-decoration:none}
    .att-meta{margin-top:6px; font-size:12px; color:#6b7280}
    .att-img{padding:0; overflow:hidden}
    .att-img img{display:block; width:100%; height:auto; max-height:340px; object-fit:contain; background:#fff; cursor:zoom-in}
    .att-img .att-meta{padding:8px}
    .att-preview{border:1px solid var(--border); border-radius:12px; overflow:hidden; background:#fff; display:flex; flex-direction:column}
    .att-preview img{display:block; width:100%; height:auto; background:#111}
    .att-preview-body{padding:8px 10px}
    .att-preview-source{font-size:12px; color:var(--muted); margin-bottom:4px}
    .att-preview-title{font-weight:600; margin-bottom:4px}
    .att-preview-lines{font-size:13px; color:#374151}
    .tbl-wrap{margin:8px 0; border:1px solid var(--border); border-radius:10px; overflow-x:auto; background:#fff}
    .tbl{width:100%; border-collapse:collapse; font-size:13px}
    .tbl td,.tbl th{padding:8px 10px; border-bottom:1px solid var(--border); vertical-align:top}
    .tbl tr:last-child td{border-bottom:none}
    .tbl tr:nth-child(even) td{background:#f8fafc}
    .img-modal{position:fixed; inset:0; background:rgba(0,0,0,0.8); display:flex; align-items:center; justify-content:center; z-index:9999}
    .img-modal[hidden]{display:none}
    .img-modal img{max-width:96vw; max-height:92vh; object-fit:contain; box-shadow:0 12px 40px rgba(0,0,0,0.4); background:#111}
    .img-modal .close{position:fixed; top:16px; right:16px; width:36px; height:36px; border-radius:18px; border:0; background:#111; color:#fff; font-size:20px; line-height:36px; cursor:pointer}
    .reactions{margin-top:6px; font-size:12px; color:#374151}
    .chip{display:inline-flex; gap:6px; align-items:center; padding:2px 8px; border-radius:999px; background:#f3f4f6; border:1px solid transparent}
    .chip.self{border-color:#2563eb; box-shadow:0 0 0 1px rgba(37,99,235,0.2) inset}
    .divider{position:relative; text-align:center; margin:18px 0}
    .divider:before, .divider:after{content:""; position:absolute; top:50%; width:42%; height:1px; background:var(--border)}
    .divider:before{left:0} .divider:after{right:0}
    .divider span{display:inline-block; padding:0 10px; color:var(--muted); background:#fff; font-weight:600}
    .compact .msg{padding:10px}
    .compact .reactions{display:none}
    .compact .reply{display:none}
    .compact .att{max-height:0; opacity:0; pointer-events:none; padding:0; border:none; margin:0}
    .compact .att-img{max-height:0; opacity:0; pointer-events:none; padding:0; border:none; margin:0}
    .compact .att-preview{max-height:0; opacity:0; pointer-events:none; padding:0; border:none; margin:0}
    .compact .tbl-wrap{max-height:0; opacity:0; pointer-events:none; padding:0; border:none; margin:0}
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
      const tables = Array.isArray(m.tables) ? m.tables : [];
      const replyTo = m.replyTo;
      const text = formatText(m.text || '');
      const avatar = m.avatar
        ? `<img src="${escapeHtml(m.avatar)}" alt="avatar" />`
        : escapeHtml((m.author || '').split(' ').map(p => p[0]).join('').slice(0, 2) || '•');

      const reactHtml = reactions
        .map(r => `<span class="chip${r.self ? ' self' : ''}">${escapeHtml(r.emoji || '')} ${r.count}${r.reactors ? ` · ${escapeHtml(r.reactors.join(', '))}` : ''}</span>`)
        .join(' ');

      const attsHtml = atts
        .map(att => {
          const label = escapeHtml(att.label || att.href || 'attachment');
          const href = att.href ? escapeHtml(att.href) : '';
          const metaText = att.metaText ? `<div class="att-meta">${escapeHtml(att.metaText)}</div>` : '';
          const type = att.type ? ` [${escapeHtml(att.type)}]` : '';
          const size = att.size ? ` (${escapeHtml(att.size)})` : '';
          const owner = att.owner ? ` — ${escapeHtml(att.owner)}` : '';
          if (att.kind === 'preview') {
            const lines = (att.metaText || '')
              .split(/\n+/)
              .map(s => s.trim())
              .filter(Boolean);
            const title = escapeHtml(lines[0] || att.label || 'Preview');
            const rest = lines.slice(1);
            const restHtml = rest.length
              ? `<div class="att-preview-lines">${rest.map(l => `<div>${escapeHtml(l)}</div>`).join('')}</div>`
              : '';
            const img = href ? `<img src="${href}" alt="${label}" />` : '';
            const source = att.label ? `<div class="att-preview-source">${label}</div>` : '';
            return `<div class="att-preview">${img}<div class="att-preview-body">${source}<div class="att-preview-title">${title}</div>${restHtml}</div></div>`;
          }
          const isImage =
            !!att.href &&
            (
              /\.(png|jpe?g|gif|webp)(\?|#|$)/i.test(att.href) ||
              /asyncgw\.teams\.microsoft\.com/i.test(att.href) ||
              /asm\.skype\.com/i.test(att.href) ||
              /\.(png|jpe?g|gif|webp)(\?|#|$)/i.test(att.label || '') ||
              /^(png|jpe?g|gif|webp)$/i.test(att.type || '')
            );

          if (isImage && href) {
            return `<div class="att-img"><img src="${href}" alt="${label}" data-full="${href}" />${metaText}</div>`;
          }
          const link = href ? `<a href="${href}" target="_blank" rel="noopener">${label}</a>` : label;
          return `<div class="att">${link}${type}${size}${owner}${metaText}</div>`;
        })
        .join('');

      const tablesHtml = tables
        .map(table => {
          if (!Array.isArray(table) || !table.length) return '';
          const rowsHtml = table
            .map(row => {
              if (!Array.isArray(row) || !row.length) return '';
              const cells = row.map(cell => `<td>${formatText(cell || '')}</td>`).join('');
              return `<tr>${cells}</tr>`;
            })
            .join('');
          if (!rowsHtml) return '';
          return `<div class="tbl-wrap"><table class="tbl"><tbody>${rowsHtml}</tbody></table></div>`;
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
        <div>${text || '<span class="meta">(no text)</span>'}</div>
        ${tablesHtml || ''}
        ${reactHtml ? `<div class="reactions">${reactHtml}</div>` : ''}
        ${attsHtml ? `<div class="atts">${attsHtml}</div>` : ''}
      </div>
    </div>`;
    })
    .join('');

  const modal = `<div class="img-modal" id="img-modal" hidden>
    <button class="close" type="button" aria-label="Close">X</button>
    <img alt="full size image" />
  </div>`;

  const script = `<script>(()=>{const btn=document.querySelector('[data-toggle-compact]');const key='teamsExporterCompact';const apply=(c)=>{document.body.classList.toggle('compact',c);if(btn)btn.textContent=c?'Switch to expanded view':'Switch to compact view';};const stored=localStorage.getItem(key);let compact=stored==='1';apply(compact);if(btn){btn.addEventListener('click',()=>{compact=!compact;apply(compact);try{localStorage.setItem(key,compact?'1':'0');}catch(_){}});}const modal=document.getElementById('img-modal');const modalImg=modal?modal.querySelector('img'):null;const closeBtn=modal?modal.querySelector('.close'):null;const close=()=>{if(modal){modal.hidden=true;}};const open=(src,alt)=>{if(!modal||!modalImg)return;modalImg.src=src;modalImg.alt=alt||'image';modal.hidden=false;};if(closeBtn){closeBtn.addEventListener('click',close);}if(modal){modal.addEventListener('click',(e)=>{if(e.target===modal)close();});}document.addEventListener('keydown',(e)=>{if(e.key==='Escape')close();});document.body.addEventListener('click',(e)=>{const t=e.target;if(!(t instanceof Element))return;const img=t.closest('.att-img img');if(!img)return;const src=img.getAttribute('data-full')||img.getAttribute('src');if(!src)return;open(src,img.getAttribute('alt')||'image');});})();</script>`;

  return `<!doctype html><meta charset="utf-8">${style}${head}${body}${modal}${script}`;
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
