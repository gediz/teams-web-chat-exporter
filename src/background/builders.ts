import type { ExportMessage, ExportMeta, Reaction, TableData, MergeRegion } from '../types/shared';
import { mintBlobUrlViaOffscreen, revokeBlobUrlViaOffscreen } from '../utils/offscreen-client';

/**
 * Removes avatar data entirely from messages.
 */
export function removeAvatars(messages: ExportMessage[]) {
  return messages.map(m => {
    if (!m.avatar && !m.avatarId) return m;
    const { avatar, avatarId, ...rest } = m;
    return rest;
  });
}

// Compact one-line summary of a message's attachments — used by TXT/CSV
// when the body is empty so attachment-only messages (image pastes,
// file drops) don't render as silent blank rows. Keeps to a few entries
// to stay readable; trailing "+N more" when truncated.
//
// Discriminator: Attachment.kind is only set for link previews; for
// everything else we look at `type` (a file extension like "png", "mp4",
// or null for inline AMS images) to pick a reasonable label.
const IMAGE_EXT_RE = /^(png|jpe?g|gif|webp|bmp|svg|ico|tif|tiff|heic)$/i;
const VIDEO_EXT_RE = /^(mp4|webm|mov|avi|mkv|m4v)$/i;
const AUDIO_EXT_RE = /^(mp3|wav|ogg|m4a|flac|aac)$/i;

export function summarizeAttachments(message: ExportMessage): string {
  const atts = Array.isArray(message.attachments) ? message.attachments : [];
  if (!atts.length) return '';
  const labelFor = (a: typeof atts[number]): string => {
    const name = (a.label || '').trim();
    if (a.kind === 'preview') {
      // metaText may carry "Title\nURL" — first line is the human label.
      const title = (a.metaText || '').split('\n')[0].trim();
      return title ? `[link: ${title}]` : '[link]';
    }
    const ext = (a.type || '').trim();
    if (IMAGE_EXT_RE.test(ext) || (!ext && /\.(png|jpe?g|gif|webp|bmp)$/i.test(name))) {
      return name && !/^image$/i.test(name) ? `[image: ${name}]` : '[image]';
    }
    if (VIDEO_EXT_RE.test(ext)) return name ? `[video: ${name}]` : '[video]';
    if (AUDIO_EXT_RE.test(ext)) return name ? `[audio: ${name}]` : '[audio]';
    // Default: generic file.
    return name ? `[file: ${name}]` : '[file]';
  };
  const MAX = 3;
  const labels = atts.slice(0, MAX).map(labelFor);
  if (atts.length > MAX) labels.push(`[+${atts.length - MAX} more]`);
  return labels.join(' ');
}

export function toCSV(messages: ExportMessage[], meta: ExportMeta = {}) {
  // Partial-export warning rendered as a leading comment block. CSV
  // doesn't have a comment syntax in the standard, but most readers
  // (Excel, LibreOffice, pandas) tolerate '#' lines at the top by
  // either showing them as a single column or skipping them with a
  // skip-rows option. Better to show the warning loudly even in CSV
  // than to hide it because the format is rigid.
  const partial = meta.partial as { reason?: string } | undefined;
  const partialBanner = partial
    ? [
        `# WARNING: This export may be incomplete (${partial.reason || 'partial'}).`,
        `# Captured ${messages?.length ?? 0} messages. The chat may contain more.`,
        '',
      ].join('\n')
    : '';
  const header = ['id', 'author', 'timestamp', 'text', 'edited', 'system', 'subject', 'importance', 'mentions', 'reactions_json', 'attachments_json', 'forwarded'];

  const rows = (messages || []).map(m => {
    const row = [];
    // RFC 4180 (https://www.rfc-editor.org/rfc/rfc4180): a field containing a
    // line break is wrapped in double quotes and the break is kept LITERAL
    // inside the quotes (not escaped to "\n"). Every cell is quoted below, so
    // we keep real newlines here; just normalize CRLF -> LF for consistency.
    let text = (m.text || '').replace(/\r\n?/g, '\n');
    // Empty body + attachments → fill the readable text column with a
    // brief summary. The structured attachments_json column still has
    // the full data; this is purely so spreadsheet readers see
    // something meaningful in the "text" column instead of "".
    if (!text.trim()) {
      const summary = summarizeAttachments(m);
      if (summary) text = summary;
    }
    row.push(m.id ?? '', m.author ?? '', m.timestamp ?? '', text, m.edited ? 'true' : 'false', m.system ? 'true' : 'false');

    row.push(m.subject ?? '');
    row.push(m.importance ?? '');
    const mentions = Array.isArray(m.mentions) ? m.mentions.map(n => n.name).join(', ') : '';
    row.push(mentions);

    const reactions = Array.isArray(m.reactions) ? m.reactions : [];
    row.push(reactions.length ? JSON.stringify(reactions) : '');

    const attachments = Array.isArray(m.attachments) ? m.attachments : [];
    row.push(attachments.length ? JSON.stringify(attachments) : '');

    // Forwarded body. CSV had no column for it, so a forwarded message (whose
    // own m.text is usually empty) lost its content here while HTML/TXT/JSON
    // kept it. Full text, newlines kept literal in-quote per RFC 4180.
    const fwd = m.forwarded;
    const forwarded = fwd && (fwd.originalAuthor || fwd.originalText)
      ? `${fwd.originalAuthor ? `${fwd.originalAuthor}: ` : ''}${(fwd.originalText || '').replace(/\r\n?/g, '\n')}`
      : '';
    row.push(forwarded);

    // RFC 4180 escaping: wrap every field in double quotes and double any
    // internal quote. Quoting every cell means embedded commas and newlines
    // need no special handling.
    return row.map(v => `"${(v ?? '').toString().split('"').join('""')}"`).join(',');
  });

  return partialBanner + [header.join(','), ...rows].join('\n');
}

export function toHTML(rows: ExportMessage[], meta: ExportMeta = {}): string[] {
  // Restore the richer HTML layout (avatars, replies, attachment grid, divider, compact mode)
  const formatRecDuration = (seconds: number) => {
    if (!Number.isFinite(seconds) || seconds <= 0) return '';
    const h = Math.floor(seconds / 3600);
    const mi = Math.floor((seconds % 3600) / 60);
    const s = seconds % 60;
    const out: string[] = [];
    if (h > 0) out.push(`${h}h`);
    if (mi > 0) out.push(`${mi}m`);
    if (s > 0 || out.length === 0) out.push(`${s}s`);
    return out.join(' ');
  };
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
  // `escaped` is already HTML-escaped, so the matched URL `u` already carries
  // entity-encoded ampersands (a `&` query separator is `&amp;` here). Do NOT
  // re-escape it, or `&amp;` becomes `&amp;amp;` and the href/link text break.
  const autolinkEscaped = (escaped: string) =>
    escaped.replace(urlRe, u => `<a href="${u}" target="_blank" rel="noopener">${u}</a>`);
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
  const avatarMap = meta.avatars || {};
  const safeCssId = (id: string) => id.replace(/[^a-zA-Z0-9_-]/g, '');

  // Render a parsed table as a real <table> with native rowspan/colspan.
  // Covered positions (the cells a span sits on top of) are skipped; the
  // span is emitted once on its anchor. Header rows become <th>.
  const renderHtmlTable = (t: TableData): string => {
    if (!t.columns || !t.rows.length) return '';
    const covered = new Set<string>();
    const anchor = new Map<string, MergeRegion>();
    for (const mg of t.merges) {
      anchor.set(`${mg.row},${mg.col}`, mg);
      for (let dr = 0; dr < mg.rowspan; dr++) {
        for (let dc = 0; dc < mg.colspan; dc++) {
          if (dr || dc) covered.add(`${mg.row + dr},${mg.col + dc}`);
        }
      }
    }
    let rowsHtml = '';
    t.rows.forEach((row, r) => {
      let cells = '';
      for (let c = 0; c < t.columns; c++) {
        if (covered.has(`${r},${c}`)) continue;
        const mg = anchor.get(`${r},${c}`);
        const tag = r < t.headerRowCount ? 'th' : 'td';
        const span = mg
          ? `${mg.rowspan > 1 ? ` rowspan="${mg.rowspan}"` : ''}${mg.colspan > 1 ? ` colspan="${mg.colspan}"` : ''}`
          : '';
        cells += `<${tag}${span}>${formatText(row[c] || '')}</${tag}>`;
      }
      rowsHtml += `<tr>${cells}</tr>`;
    });
    return `<div class="tbl-wrap"><table class="tbl">${rowsHtml}</table></div>`;
  };

  const style = `<style>
    :root { --muted:#6b7280; --border:#e5e7eb; --bg:#ffffff; --chip:#f3f4f6; --thread-bg:#f8fafc; --thread-border:#dbeafe; --thread-accent:#3b82f6; }
    body{font:14px system-ui, -apple-system, Segoe UI, Roboto; background:#fff; color:#111; padding:20px}
    h1{margin:0 0 10px 0}
    .meta{color:var(--muted); margin:0 0 12px 0}
    .partial-warning{background:#fef3c7; border:1px solid #f59e0b; border-radius:6px; padding:10px 14px; margin:0 0 14px 0; color:#78350f; font-size:13px; line-height:1.5}
    .partial-warning strong{color:#7c2d12}
    .partial-warning .partial-tag{font-family:ui-monospace,Menlo,monospace; font-size:11px; color:#92400e; opacity:0.85}
    .toolbar{margin-bottom:12px; display:flex; gap:8px; align-items:center}
    .toolbar button{border:1px solid var(--border); background:#f9fafb; color:#111; padding:6px 10px; border-radius:6px; cursor:pointer; font:13px system-ui}
    .toolbar button:hover{background:#eef2f7}
    .msg{position:relative; display:flex; gap:10px; margin:12px 0; padding:12px; border:1px solid var(--border); border-radius:12px; background:var(--bg)}
    /* Issue #20: highlight messages authored by the current Teams user.
       Uses a tinted background + accent-colored left rail to match the
       "it's me" affordance from the Teams UI without overpowering the
       document. Inherits radius + borders from .msg above. */
    .own-msg{background:rgba(37,99,235,0.06); border-color:rgba(37,99,235,0.25); border-left:4px solid #2563eb; padding-left:9px}
    .avt{flex:0 0 36px; width:36px; height:36px; border-radius:50%; background:#eef2f7; overflow:hidden; display:flex; align-items:center; justify-content:center; font-weight:600; color:#334155}
    .avt img{width:36px; height:36px; border-radius:50%; display:block}
    .avt-img{width:36px; height:36px; border-radius:50%; display:block; background-size:cover; background-position:center}
    .main{flex:1}
    code{font-family:ui-monospace,SFMono-Regular,Menlo,Consolas,"Liberation Mono",monospace;background:#f3f4f6;border:1px solid #e5e7eb;padding:1px 4px;border-radius:4px}
    pre.code-block{background:#0b1020;color:#e5e7eb;border-radius:10px;padding:10px 12px;overflow:auto;margin:8px 0;border:1px solid #111827;max-width:100%;box-sizing:border-box;white-space:pre-wrap;word-break:break-word;overflow-wrap:anywhere}
    pre.code-block code{background:none;border:none;padding:0;color:inherit;white-space:inherit}
    .hdr{color:var(--muted); font-size:12px; margin-bottom:6px}
    .hdr .rel{margin-left:6px; font-style:italic}
    .hdr .edited{font-style:italic}
    .reply{background:#f8fafc; border-left:3px solid #d1d5db; padding:8px 10px; border-radius:8px; margin:8px 0; font-size:13px; color:#374151}
    .reply .reply-meta{display:flex; flex-wrap:wrap; gap:6px; font-size:12px; color:#6b7280; margin-bottom:4px}
    .reply blockquote{margin:0; padding:0; border:none; color:#1f2937; word-wrap:break-word; overflow-wrap:anywhere}
    .forward{background:#f0f4ff; border:1px solid #dbeafe; border-left:3px solid #60a5fa; border-radius:8px; padding:8px 12px; margin:6px 0; font-size:13px}
    .forward-header{color:#6b7280; font-size:12px; display:flex; align-items:center; gap:6px; flex-wrap:wrap; margin-bottom:4px}
    .forward-icon{font-size:14px}
    .forward-ts{color:#9ca3af}
    .forward-body{color:#374151; margin-top:4px; word-wrap:break-word; overflow-wrap:anywhere; white-space:pre-wrap}
    blockquote{margin:8px 0; padding:8px 10px; border-left:3px solid #d1d5db; background:#f8fafc; color:#374151}
    .thread{border:1px solid var(--thread-border); background:var(--thread-bg); border-radius:14px; padding:10px 12px; margin:14px 0}
    .thread-parent .msg{margin:0; border-left:4px solid var(--thread-accent)}
    .thread-meta{display:flex; align-items:center; gap:8px; font-size:12px; color:var(--muted); margin:8px 2px 2px 2px}
    .thread-toggle{border:1px solid var(--border); background:#f9fafb; color:#111; padding:2px 8px; border-radius:999px; cursor:pointer; font:12px system-ui}
    .thread-toggle:hover{background:#eef2f7}
    .thread.collapsed .thread-replies{display:none}
    .thread.collapsed .thread-meta{opacity:.85}
    .thread-replies{margin-top:6px; padding-left:18px; position:relative}
    .thread-replies:before{content:""; position:absolute; left:7px; top:0; bottom:0; width:2px; background:var(--thread-border)}
    .msg.reply-msg{margin:10px 0 0 0; background:#fff}
    .msg.reply-msg:before{content:""; position:absolute; left:-16px; top:25px; width:10px; height:10px; border-radius:50%; background:var(--thread-accent)}
    .atts{display:grid; grid-template-columns:repeat(auto-fill,minmax(220px,1fr)); gap:8px; margin-top:8px}
    .att, .att-img{border:1px solid var(--border); border-radius:10px; padding:8px; background:#fff; transition:max-height .2s ease, opacity .2s ease}
    .att a{word-break:break-word; overflow-wrap:anywhere; text-decoration:none}
    .att-audio{max-width:320px}
    .att-audio audio{width:100%; margin-top:6px; border-radius:8px}
    .att-video{position:relative; max-width:240px}
    .att-video img{border-radius:10px; cursor:pointer !important}
    .video-badge{position:absolute; bottom:8px; left:8px; background:rgba(0,0,0,0.7); color:#fff; font-size:12px; font-weight:600; padding:2px 8px; border-radius:4px}
    .att-meta{margin-top:6px; font-size:12px; color:#6b7280}
    .att-missing{color:#6b7280; font-size:13px; background:#fafafa; border-style:dashed}
    .att-missing .att-meta-hint{color:#9ca3af; font-size:11px; margin-left:4px}
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
    .tbl{border-collapse:collapse; font-size:13px}
    /* Full cell borders so rowspan/colspan merges read as one cell. */
    .tbl td,.tbl th{padding:7px 10px; border:1px solid var(--border); vertical-align:top; text-align:left}
    .tbl th{background:#f1f5f9; font-weight:600}
    .img-modal{position:fixed; inset:0; background:rgba(0,0,0,0.8); display:flex; align-items:center; justify-content:center; z-index:9999}
    .img-modal[hidden]{display:none}
    .img-modal img{max-width:96vw; max-height:92vh; object-fit:contain; box-shadow:0 12px 40px rgba(0,0,0,0.4); background:#111}
    .img-modal .close{position:fixed; top:16px; right:16px; width:36px; height:36px; border-radius:18px; border:0; background:#111; color:#fff; font-size:20px; line-height:36px; cursor:pointer}
    .subject{font-weight:600; font-size:15px; margin-bottom:4px; color:#0f172a}
    .badge-urgent{display:inline-block; background:#dc2626; color:#fff; font-size:11px; font-weight:600; padding:1px 6px; border-radius:4px; margin-left:6px; vertical-align:middle}
    .badge-important{display:inline-block; background:#d97706; color:#fff; font-size:11px; font-weight:600; padding:1px 6px; border-radius:4px; margin-left:6px; vertical-align:middle}
    .mention{background:#dbeafe; padding:0 3px; border-radius:3px; color:#1d4ed8; font-weight:500}
    .reactions{margin-top:6px; font-size:12px; color:#374151}
    .chip{display:inline-flex; gap:6px; align-items:center; padding:2px 8px; border-radius:999px; background:#f3f4f6; border:1px solid transparent; position:relative; cursor:default}
    .chip.self{border-color:#2563eb; box-shadow:0 0 0 1px rgba(37,99,235,0.2) inset}
    /* Reactor chip v9 (issue #17/#28).
       - avatar dot stack: up to 3 overlapping circles (+N overflow badge)
       - inline names: adaptive (1 / 2-3 / 4+)
       - popover on hover / :focus-within (keyboard + mobile tap) */
    .chip-emoji{font-size:14px; line-height:1}
    .chip-avatars{display:inline-flex; align-items:center}
    .chip-avatars .avt-dot{width:16px; height:16px; border-radius:50%; background-size:cover; background-position:center; background-color:#e5e7eb; color:#374151; font-size:8px; font-weight:600; display:inline-flex; align-items:center; justify-content:center; box-shadow:0 0 0 2px var(--bg, #ffffff)}
    .chip-avatars .avt-dot + .avt-dot, .chip-avatars .avt-dot + .avt-more{margin-left:-5px}
    .chip-avatars .avt-more{min-width:16px; height:16px; padding:0 3px; border-radius:10px; background:#9ca3af; color:#fff; font-size:9px; font-weight:700; display:inline-flex; align-items:center; justify-content:center; box-shadow:0 0 0 2px var(--bg, #ffffff)}
    .chip-names{font-size:11px; color:#4b5563}
    .chip-popover{position:absolute; bottom:calc(100% + 6px); left:0; background:var(--bg, #fff); color:var(--text, #111827); border:1px solid var(--border, #e5e7eb); border-radius:8px; padding:8px 10px; box-shadow:0 8px 24px -8px rgba(0,0,0,0.25); min-width:180px; max-width:280px; z-index:10; display:none; pointer-events:none}
    .chip:hover .chip-popover, .chip:focus-within .chip-popover{display:block}
    .chip-pop-row{display:flex; align-items:center; gap:8px; padding:3px 0; font-size:12px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis}
    .chip-pop-row .avt-dot{width:18px; height:18px; flex-shrink:0; border-radius:50%; background-size:cover; background-position:center; background-color:#e5e7eb; box-shadow:none}
    .chip-pop-self{font-weight:600; color:#2563eb}
    .divider{display:flex; align-items:center; gap:10px; margin:18px 0}
    .divider:before, .divider:after{content:""; flex:1; height:1px; background:var(--border)}
    .divider > span{color:var(--muted); font-weight:600; white-space:nowrap; font-size:13px; display:inline-flex; gap:8px; align-items:baseline}
    .divider-time{color:var(--muted); font-weight:400; font-size:12px; opacity:0.75}
    /* Call/meeting system events: 3-column header (duration | label | timestamp)
       between horizontal lines, with an optional attendees row below using
       lighter hairlines. Same shape for start (no duration) and end. */
    .divider-block{margin:18px 0}
    .divider-row{display:flex; align-items:center; gap:10px}
    .divider-row:before, .divider-row:after{content:""; flex:1; height:1px; background:var(--border)}
    .divider-header{display:inline-flex; gap:16px; align-items:baseline; padding:0 8px}
    .divider-h-left{color:var(--muted); font-weight:500; font-size:12px; opacity:0.85; min-width:80px; text-align:right}
    .divider-h-center{color:var(--muted); font-weight:600; font-size:13px; white-space:nowrap}
    .divider-h-right{color:var(--muted); font-weight:400; font-size:12px; opacity:0.85; min-width:160px}
    .divider-att-row{display:flex; align-items:center; gap:10px; margin-top:4px}
    .divider-att-row:before, .divider-att-row:after{content:""; flex:1; height:1px; background:var(--border); opacity:0.4}
    .divider-att-row > span{color:var(--muted); font-weight:400; font-size:12px; opacity:0.85; max-width:65%; text-align:center; line-height:1.4}
    .divider-icon{display:inline-block; opacity:0.65; font-size:12px; vertical-align:-1px}
    .compact .msg{padding:10px}
    .compact .reactions{display:none}
    .compact .reply{display:none}
    .compact .att{max-height:0; opacity:0; pointer-events:none; padding:0; border:none; margin:0}
    .compact .att-img{max-height:0; opacity:0; pointer-events:none; padding:0; border:none; margin:0}
    .compact .att-preview{max-height:0; opacity:0; pointer-events:none; padding:0; border:none; margin:0}
    .att-summary{display:none; font-size:12px; color:var(--muted); margin-top:4px}
    .compact .att-summary{display:block}
    .compact .tbl-wrap{max-height:0; opacity:0; pointer-events:none; padding:0; border:none; margin:0}
    .compact .atts{gap:0; margin-top:0}
    .compact .avt{display:none}
    .compact .msg{margin:8px 0; border-color:rgba(0,0,0,0.08)}
    .compact .thread{padding:6px 8px}
    .compact .thread-replies{padding-left:12px}
    .compact .msg.reply-msg:before{left:-15px; top:24px; width:8px; height:8px}
    .main > div{word-break:break-word; overflow-wrap:anywhere}
    ${(() => {
      // Avatar CSS source selection.
      //   Default (inline): generate `url("data:...")` with the base64
      //   avatar inline — works for standalone HTML.
      //   Zip-mode (`_avatarsAsFiles`): reference `avatars/<file>` files
      //   that the zip builder will drop in alongside the HTML. Keeps
      //   the HTML payload small when many avatars are present.
      const asFiles = (meta as { _avatarsAsFiles?: boolean })._avatarsAsFiles === true;
      const fileMap = (meta as { _avatarFileMap?: Record<string, string> })._avatarFileMap || {};
      return Object.entries(avatarMap).map(([id, dataUrl]) => {
        if (asFiles) {
          const file = fileMap[id];
          // Fall back to inline if, for some reason, the file map
          // missed this id — keeps the export from rendering blank.
          if (!file) return `.avt-${safeCssId(id)}{background-image:url("${dataUrl.replace(/["\\]/g, '')}")}`;
          return `.avt-${safeCssId(id)}{background-image:url("avatars/${file}")}`;
        }
        return `.avt-${safeCssId(id)}{background-image:url("${dataUrl.replace(/["\\]/g, '')}")}`;
      }).join('\n    ');
    })()}
  </style>`;

  const metaParts = [];
  if (meta.messages != null || meta.count != null) metaParts.push(`<b>Messages:</b> ${escapeHtml(String(meta.messages ?? meta.count ?? ''))}`);
  if (meta.timeRange) metaParts.push(`<b>Range:</b> ${escapeHtml(meta.timeRange)}`);
  const metaLine = metaParts.length ? `<p class="meta">${metaParts.join(' &nbsp; ')}</p>` : '';

  // Partial-export warning banner. Shown right at the top of the
  // body when the scrape signalled an incomplete-data condition. The
  // banner stays in the rendered file forever, so a user opening the
  // export weeks later still sees the caveat. The reason tag (network
  // / truncation) is included verbatim to help bug-report triage.
  const partial = meta.partial as { reason?: string } | undefined;
  const partialBanner = partial
    ? `<div class="partial-warning" role="alert">
         <strong>⚠ Export may be incomplete.</strong>
         ${partial.reason === 'network'
           ? 'A network interruption was detected during scraping; some messages may be missing.'
           : 'Some messages may not have fully loaded before the export finished.'}
         Captured ${escapeHtml(String(meta.count ?? meta.messages ?? '?'))} messages.
         <span class="partial-tag">[${escapeHtml(String(partial.reason || 'partial'))}]</span>
       </div>`
    : '';

  const head = `<h1>${escapeHtml(meta.title || 'Teams Chat Export')}</h1>
    ${metaLine}
    ${partialBanner}
    <div class="toolbar"><button type="button" data-toggle-compact>Toggle compact view</button></div><hr/>`;

  // Reactor chip (v9 spec, issue #17/#28). Renders emoji + avatar dot
  // stack (up to 3 + "+N" overflow) + adaptive inline names + popover.
  // Back-compat: if `reactors` is an older string[] shape, treat each
  // entry as a name-only reactor.
  function renderReactorChip(r: Reaction, avatars: Record<string, string>): string {
    const emoji = escapeHtml(r.emoji || '');
    const self = r.self ? ' self' : '';
    const reactorsList = Array.isArray(r.reactors) ? r.reactors : [];
    // Normalize to ReactorInfo shape (handles legacy string[] gracefully).
    const normalized = reactorsList.map(x => {
      if (typeof x === 'string') return { name: x } as { name: string; avatarId?: string; self?: boolean };
      return x as { name: string; avatarId?: string; self?: boolean };
    });
    // No reactors resolved (e.g. self-chat reaction): fall back to the
    // compact emoji + count format — keeps old exports' look.
    if (!normalized.length) {
      return `<span class="chip${self}"><span class="chip-emoji">${emoji}</span> ${r.count}</span>`;
    }

    // Avatar stack: up to 3 circles, "+N" badge if more. Avatars use the
    // same CSS classes as message avatars (same keys in avatarMap); for
    // reactors without a resolved avatarId we render initials.
    const MAX_DOTS = 3;
    const dots = normalized.slice(0, MAX_DOTS).map(reactor => {
      if (reactor.avatarId && avatars[reactor.avatarId]) {
        return `<span class="avt-dot avt-${safeCssId(reactor.avatarId)}"></span>`;
      }
      const initials = (reactor.name || '').split(/\s+/).map(p => p[0] || '').join('').slice(0, 2).toUpperCase() || '?';
      return `<span class="avt-dot">${escapeHtml(initials)}</span>`;
    }).join('');
    const overflow = normalized.length > MAX_DOTS
      ? `<span class="avt-more">+${normalized.length - MAX_DOTS}</span>`
      : '';

    // Adaptive inline names:
    //   1 reactor      → "Name" (or "You" for self)
    //   2-3 reactors   → "Name1, Name2, Name3"
    //   4+ reactors    → "First & N" (where N = total - 1)
    const nameFor = (reactor: { name: string; self?: boolean }) => reactor.self ? 'You' : reactor.name;
    let inlineNames: string;
    if (normalized.length === 1) {
      inlineNames = nameFor(normalized[0]);
    } else if (normalized.length <= 3) {
      inlineNames = normalized.map(nameFor).join(', ');
    } else {
      inlineNames = `${nameFor(normalized[0])} & ${normalized.length - 1}`;
    }

    // Popover: full list, one row per reactor. Avatar dot + name. Shown
    // on :hover and :focus-within via CSS.
    const popoverRows = normalized.map(reactor => {
      let dot: string;
      if (reactor.avatarId && avatars[reactor.avatarId]) {
        dot = `<span class="avt-dot avt-${safeCssId(reactor.avatarId)}"></span>`;
      } else {
        const init = (reactor.name || '').split(/\s+/).map(p => p[0] || '').join('').slice(0, 2).toUpperCase() || '?';
        dot = `<span class="avt-dot">${escapeHtml(init)}</span>`;
      }
      const nameHtml = escapeHtml(reactor.self ? 'You' : reactor.name);
      return `<span class="chip-pop-row${reactor.self ? ' chip-pop-self' : ''}">${dot}${nameHtml}</span>`;
    }).join('');

    return `<span class="chip${self}" tabindex="0">
      <span class="chip-emoji">${emoji}</span>
      <span class="chip-avatars">${dots}${overflow}</span>
      <span class="chip-names">${escapeHtml(inlineNames)}</span>
      <span class="chip-popover" role="tooltip">${popoverRows}</span>
    </span>`;
  }

  let msgIndex = 0;
  const renderMessage = (m: ExportMessage, opts: { isReply?: boolean } = {}) => {
    const idx = msgIndex++;
    const ts = m.timestamp || '';
    const rel = relLabel(ts);
    const tsLabel = escapeHtml(fmtTs(ts));
    const reactions = Array.isArray(m.reactions) ? m.reactions : [];
    const atts = Array.isArray(m.attachments) ? m.attachments : [];
    const tables = Array.isArray(m.tables) ? m.tables : [];
    const replyTo = m.replyTo;
    const text = formatText(m.text || '');
    // Use CSS class for avatars (avoids duplicating large data URLs in every message)
    const avatarId = m.avatarId || '';
    const hasAvatar = avatarId && avatarMap[avatarId];
    const avatar = hasAvatar
      ? `<span class="avt-img avt-${safeCssId(avatarId)}"></span>`
      : m.avatar
        ? `<img src="${escapeHtml(m.avatar)}" alt="avatar" />`
        : escapeHtml((m.author || '').split(' ').map(p => p[0]).join('').slice(0, 2) || '?');

    const reactHtml = reactions
      .map(r => renderReactorChip(r, avatarMap))
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
        // Video attachments: show thumbnail with play overlay + link to video
        if (att.type === 'video') {
          const thumbSrc = att.dataUrl || href;
          const videoUrl = att.owner || ''; // video URL stored in owner field
          const videoLink = videoUrl ? `<a href="${escapeHtml(videoUrl)}" target="_blank" rel="noopener">` : '';
          const videoLinkEnd = videoUrl ? '</a>' : '';
          if (thumbSrc) {
            return `<div class="att-img att-video">${videoLink}<img src="${escapeHtml(thumbSrc)}" alt="${escapeHtml(att.label || 'Video')}" /><div class="video-badge">${escapeHtml(att.size || '▶')}</div>${videoLinkEnd}</div>`;
          }
          return `<div class="att">${videoLink}🎬 ${escapeHtml(att.label || 'Video')}${videoLinkEnd}</div>`;
        }
        // Audio attachments (voice messages)
        if (att.type === 'audio') {
          const audioSrc = att.dataUrl || href;
          if (audioSrc) {
            return `<div class="att att-audio">🎤 ${escapeHtml(att.label || 'Voice message')}<br><audio controls preload="none" src="${escapeHtml(audioSrc)}">Your browser does not support audio.</audio></div>`;
          }
          return `<div class="att">🎤 ${label}</div>`;
        }
        const isEmbeddedImage = !!att.dataUrl || /^data:image\//i.test(att.href || '');
        const isAmsImage = /asyncgw\.teams\.microsoft\.com|asm\.skype\.com/i.test(att.href || '');
        const isLocalImage = /^images\//i.test(att.href || '');
        // Hosts that 401/403 (or redirect to a sign-in page) when an
        // `<img src=...>` is loaded from a saved HTML file with no
        // session cookies. Renders as a placeholder card instead of
        // a broken-icon img tag.
        const isAuthProtected = isAmsImage
          || /sharepoint\.com|sharepoint\.us|sharepoint-mil\.us|sharepoint\.cn|microsoftpersonalcontent\.com/i.test(att.href || '');
        const looksLikeImage =
          /\.(png|jpe?g|gif|webp|bmp|svg|tiff?|heic)(\?|#|$)/i.test(att.href || '') ||
          /\.(png|jpe?g|gif|webp|bmp|svg|tiff?|heic)(\?|#|$)/i.test(att.label || '') ||
          /^(png|jpe?g|gif|webp|bmp|svg|tiff?|heic)$/i.test(att.type || '');
        // Only render as <img> if we have the data locally — either as a
        // data: URL (embedded) or a relative path inside the zip. Any URL
        // that needs auth (Teams AMS, SharePoint, OneDrive) would 401 as
        // an img src when the saved file is opened later, so we render
        // those as a clean placeholder card instead of a broken icon.
        const isImage = !!att.href && (isEmbeddedImage || isLocalImage || (looksLikeImage && !isAuthProtected));

        if (isImage && href) {
          return `<div class="att-img"><img src="${href}" alt="${label}" data-full="${href}" />${metaText}</div>`;
        }
        // Auth-protected image that wasn't downloaded into this export.
        // Show a quiet placeholder so users see "there was an image here"
        // without the visual noise of a broken-icon `<img>` tag.
        // An auth-protected image, or one whose unfetchable remote href was
        // cleared during scraping, renders as a quiet placeholder rather than
        // a broken-icon <img>.
        if (looksLikeImage && (isAuthProtected || !att.href)) {
          // We couldn't embed the image (it needs auth). If we still have its
          // original URL, link it so the user can open it in the browser where
          // they may be signed in, instead of losing it; otherwise show a quiet
          // "not included" card. Generic across hosts (SharePoint, OneDrive,
          // AMS, ...) — the link points at wherever the image actually lives.
          const canLink = !!href && /^https?:\/\//i.test(att.href || '');
          if (canLink) {
            return `<div class="att att-missing">🖼️ <a href="${href}" target="_blank" rel="noopener">${label}</a> <span class="att-meta-hint">(open original)</span></div>`;
          }
          return `<div class="att att-missing">🖼️ ${label} <span class="att-meta-hint">(not included)</span></div>`;
        }
        const link = href ? `<a href="${href}" target="_blank" rel="noopener">${label}</a>` : label;
        return `<div class="att">${link}${type}${size}${owner}${metaText}</div>`;
      })
      .join('');

    // Legacy DOM-scrape tables: flat string[][][], appended after the text.
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

    // API mode: ordered text/table blocks. Renders text and real tables in
    // their original positions. Falls back to the flat text + legacy tables
    // when no bodyBlocks were parsed (non-table messages, DOM-scrape mode).
    const blocks = Array.isArray(m.bodyBlocks) ? m.bodyBlocks : null;
    const bodyHtml = blocks && blocks.length
      ? blocks
          .map(b => (b.type === 'table' ? renderHtmlTable(b.table) : `<div>${formatText(b.text)}</div>`))
          .join('')
      : `<div>${text || (m.forwarded || atts.length ? '' : '<span class="meta">(no text)</span>')}</div>${tablesHtml || ''}`;

    // Forward card (API mode provides ForwardContext)
    let forwardHtml = '';
    if (m.forwarded && !opts.isReply) {
      const fwd = m.forwarded;
      const origAuthor = fwd.originalAuthor || 'unknown';
      const origTs = fwd.originalTimestamp ? new Date(fwd.originalTimestamp).toLocaleString() : '';
      const origText = fwd.originalText ? escapeHtml(fwd.originalText) : '';
      forwardHtml = `<div class="forward"><div class="forward-header"><span class="forward-icon">↪</span> Forwarded from <strong>${escapeHtml(origAuthor)}</strong>${origTs ? ` <span class="forward-ts">${escapeHtml(origTs)}</span>` : ''}</div>${origText ? `<div class="forward-body">${origText}</div>` : ''}</div>`;
    }

    // Reply preview
    const hasReplyPreview = replyTo && (replyTo.author || replyTo.timestamp || replyTo.text);
    const replyHtml = hasReplyPreview && !opts.isReply && !m.forwarded
      ? `<div class="reply"><div class="reply-meta">↩︎ <strong>${escapeHtml(replyTo!.author || '')}</strong>${replyTo!.timestamp ? `<span>• ${escapeHtml(replyTo!.timestamp)}</span>` : ''}</div><blockquote>${escapeHtml(replyTo!.text || '')}</blockquote></div>`
      : '';

    // Own-message CSS hook (issue #20). Lets a single rule paint the
     // viewer's own messages with a distinct accent without needing to
     // rebuild the whole message markup per-author.
     const msgClass = `msg${opts.isReply ? ' reply-msg' : ''}${m.isOwn ? ' own-msg' : ''}`;
    return `<div class="${msgClass}" id="msg-${idx}">
      <div class="avt">${avatar}</div>
      <div class="main">
        <div class="hdr">${escapeHtml(m.author || '')} — <span title="${escapeHtml(ts)}">${tsLabel}</span>${rel ? `<span class="rel">(${rel})</span>` : ''}${m.edited ? ' <span class="edited">• edited</span>' : ''}${m.importance === 'urgent' ? '<span class="badge-urgent">URGENT</span>' : m.importance === 'high' ? '<span class="badge-important">IMPORTANT</span>' : ''}</div>
        ${m.subject ? `<div class="subject">${escapeHtml(m.subject)}</div>` : ''}
        ${forwardHtml}${replyHtml}
        ${bodyHtml}
        ${reactHtml ? `<div class="reactions">${reactHtml}</div>` : ''}
        ${attsHtml ? `<div class="atts">${attsHtml}</div>` : ''}
        ${atts.length ? `<div class="att-summary">📎 ${atts.length} attachment${atts.length > 1 ? 's' : ''}</div>` : ''}
      </div>
    </div>`;
  };

  const normalizeKey = (value?: string | null) => (value || '').trim().toLowerCase().replace(/\s+/g, ' ');
  const parseMs = (value?: string | null) => {
    if (!value) return null;
    const ms = Date.parse(value);
    if (!Number.isNaN(ms)) return ms;
    const normalized = value.replace(/ /g, 'T');
    const ms2 = Date.parse(normalized);
    return Number.isNaN(ms2) ? null : ms2;
  };
  const minuteKey = (ms: number) => Math.floor(ms / 60000);
  const textMatches = (a: string, b: string) => {
    if (!a || !b) return false;
    return a.includes(b) || b.includes(a);
  };

  type ParentEntry = {
    idx: number;
    authorKey: string;
    textKey: string;
    tsKey: string;
    tsMs: number | null;
    minute: number | null;
  };

  const parentByIndex = new Map<number, ParentEntry>();
  const parentsByAuthorTimestamp = new Map<string, number[]>();
  const parentsByAuthorMinute = new Map<string, number[]>();
  const parentsByMinute = new Map<number, number[]>();

  for (let i = 0; i < rows.length; i++) {
    const msg = rows[i];
    if (!msg || msg.system) continue;
    const authorKey = normalizeKey(msg.author);
    const textKey = normalizeKey((msg.text || '').slice(0, 280));
    const tsKey = normalizeKey(msg.timestamp);
    const tsMs = parseMs(msg.timestamp);
    const minute = typeof tsMs === 'number' ? minuteKey(tsMs) : null;
    const entry: ParentEntry = { idx: i, authorKey, textKey, tsKey, tsMs, minute };
    parentByIndex.set(i, entry);
    if (authorKey && tsKey) {
      const key = `${authorKey}|${tsKey}`;
      const list = parentsByAuthorTimestamp.get(key) || [];
      list.push(i);
      parentsByAuthorTimestamp.set(key, list);
    }
    if (authorKey && minute != null) {
      const key = `${authorKey}|${minute}`;
      const list = parentsByAuthorMinute.get(key) || [];
      list.push(i);
      parentsByAuthorMinute.set(key, list);
    }
    if (minute != null) {
      const list = parentsByMinute.get(minute) || [];
      list.push(i);
      parentsByMinute.set(minute, list);
    }
  }

  const repliesByParent = new Map<number, { index: number; msg: ExportMessage }[]>();
  const replyIndices = new Set<number>();

  const findParentIndex = (reply: ExportMessage, replyIndex: number): number | null => {
    const replyTo = reply.replyTo;
    if (!replyTo) return null;
    if (replyTo.id) {
      const byId = rows.findIndex((m, idx) => idx <= replyIndex && m?.id === replyTo.id);
      if (byId >= 0) return byId;
    }
    const authorKey = normalizeKey(replyTo.author);
    const textKey = normalizeKey((replyTo.text || '').slice(0, 280));
    const tsKey = normalizeKey(replyTo.timestamp);
    const tsMs = parseMs(replyTo.timestamp);
    const minute = typeof tsMs === 'number' ? minuteKey(tsMs) : null;

    let candidates: number[] = [];
    if (authorKey && tsKey) {
      candidates = parentsByAuthorTimestamp.get(`${authorKey}|${tsKey}`) || [];
    }
    if (authorKey && minute != null) {
      candidates = candidates.length ? candidates : (parentsByAuthorMinute.get(`${authorKey}|${minute}`) || []);
    }
    if (!candidates.length && minute != null) {
      candidates = parentsByMinute.get(minute) || [];
    }
    if (!candidates.length && authorKey && textKey) {
      candidates = Array.from(parentByIndex.values())
        .filter(p => p.authorKey === authorKey && textMatches(p.textKey, textKey))
        .map(p => p.idx);
    }
    if (!candidates.length && textKey) {
      candidates = Array.from(parentByIndex.values())
        .filter(p => textMatches(p.textKey, textKey))
        .map(p => p.idx);
    }
    if (!candidates.length) return null;

    const earlier = candidates.filter(idx => idx <= replyIndex);
    if (!earlier.length) return null;
    candidates = earlier;

    if (candidates.length === 1) return candidates[0];

    let bestIdx = candidates[0];
    let bestScore = Number.POSITIVE_INFINITY;
    for (const idx of candidates) {
      const parent = parentByIndex.get(idx);
      if (!parent) continue;
      let score = 0;
      if (textKey && parent.textKey && textMatches(parent.textKey, textKey)) {
        score -= 1000000;
      }
      if (tsMs != null && parent.tsMs != null) {
        score += Math.abs(parent.tsMs - tsMs);
      } else {
        score += Math.abs(idx - replyIndex) * 60000;
      }
      if (score < bestScore) {
        bestScore = score;
        bestIdx = idx;
      }
    }
    return bestIdx;
  };

  // First pass: wire each reply to its immediate parent.
  const directParentOf = new Map<number, number>();
  for (let i = 0; i < rows.length; i++) {
    const m = rows[i];
    if (!m || !m.replyTo || m.system) continue;
    const parentIdx = findParentIndex(m, i);
    if (parentIdx == null || parentIdx === i) continue;
    directParentOf.set(i, parentIdx);
    replyIndices.add(i);
  }
  // Second pass: fold reply chains onto the top-most non-reply ancestor.
  // Teams UIs let users reply to a reply, which produces chains like
  // G → P → C. The renderer below walks each top-level parent's
  // repliesByParent list once; if we left P→C un-flattened, C would
  // land in repliesByParent[P] and never be visited (because P is in
  // replyIndices and skipped in the main render loop). Walk up until
  // we hit a node with no direct parent — that's the thread root.
  for (const replyIdx of directParentOf.keys()) {
    let root = directParentOf.get(replyIdx)!;
    const seen = new Set<number>([replyIdx, root]);
    while (directParentOf.has(root)) {
      const next = directParentOf.get(root)!;
      if (seen.has(next)) break; // defensive: malformed data, no infinite loop
      seen.add(next);
      root = next;
    }
    const list = repliesByParent.get(root) || [];
    list.push({ index: replyIdx, msg: rows[replyIdx] });
    repliesByParent.set(root, list);
  }

  const parts: string[] = [];
  for (let i = 0; i < rows.length; i++) {
    if (replyIndices.has(i)) continue;
    const m = rows[i];
    if (m.system) {
      const rawText = m.text || m.author || '[system]';
      const ts = m.timestamp ? fmtTs(m.timestamp) : '';
      const timeIso = m.timestamp || '';
      const attendees = m.systemAttendees;

      // CallRecording messages re-use the same 3-column divider layout as the
      // Meeting started/ended events: ⏱ duration on the left, "Recording —
      // {title}" in the center, timestamp 🕒 on the right; organizer + attendees
      // appear as below-divider rows. The recording's playable artifacts on
      // asyncgw / SharePoint require auth that a static export can't supply,
      // so we don't render any links — just the metadata we have.
      const rec = m.recordingDetails;
      if (rec) {
        const title = rec.title || rec.meetingSubject || 'Meeting recording';
        const dur = rec.durationSec ? formatRecDuration(rec.durationSec) : '';
        // Right-side timestamp: prefer the meeting's actual start (when the
        // call began) over the recording-message composetime.
        const rightTs = rec.meetingStart || rec.meetingEnd || timeIso;
        const rightDisplay = rightTs ? fmtTs(rightTs) : '';
        const leftHtml = `<span class="divider-icon">⏱</span>${dur ? ' ' + escapeHtml(dur) : ''}`;
        const rightHtml = rightDisplay
          ? `${escapeHtml(rightDisplay)} <span class="divider-icon">🕒</span>`
          : '';
        const detailRows: string[] = [];
        if (rec.organizerUpn) {
          detailRows.push(
            `<div class="divider-att-row"><span>organized by ${escapeHtml(rec.organizerUpn)}</span></div>`,
          );
        }
        if (rec.attendees && rec.attendees.length) {
          detailRows.push(
            `<div class="divider-att-row"><span>${escapeHtml(rec.attendees.join(', '))}</span></div>`,
          );
        }
        parts.push(
          `<div class="divider-block">` +
          `<div class="divider-row">` +
          `<div class="divider-header">` +
          `<div class="divider-h-left">${leftHtml}</div>` +
          `<div class="divider-h-center">Recording — ${escapeHtml(title)}</div>` +
          `<div class="divider-h-right" title="${escapeHtml(rightTs || '')}">${rightHtml}</div>` +
          `</div>` +
          `</div>` +
          detailRows.join('') +
          `</div>`,
        );
        continue;
      }

      // Call/meeting events use the 3-column header: duration | label | timestamp.
      // Other system events (member joined/left, recording, transcript, date dividers)
      // keep the simple horizontal divider.
      const callMatch = rawText.match(/^(Meeting|Call) (started|ended)(?: — (.+))?$/);
      if (callMatch) {
        const labelText = `${callMatch[1]} ${callMatch[2]}`;
        const duration = callMatch[3] || '';
        const isStarted = callMatch[2] === 'started';
        const leftIcon = isStarted ? '▶' : '⏱';
        const leftHtml = `<span class="divider-icon">${leftIcon}</span>${duration ? ' ' + escapeHtml(duration) : ''}`;
        const rightHtml = ts
          ? `${escapeHtml(ts)} <span class="divider-icon">🕒</span>`
          : '';
        const attRowHtml = attendees && attendees.length
          ? `<div class="divider-att-row"><span>${escapeHtml(attendees.join(', '))}</span></div>`
          : '';
        parts.push(
          `<div class="divider-block">` +
          `<div class="divider-row">` +
          `<div class="divider-header">` +
          `<div class="divider-h-left">${leftHtml}</div>` +
          `<div class="divider-h-center">${escapeHtml(labelText)}</div>` +
          `<div class="divider-h-right" title="${escapeHtml(timeIso)}">${rightHtml}</div>` +
          `</div>` +
          `</div>` +
          `${attRowHtml}` +
          `</div>`,
        );
      } else {
        const label = escapeHtml(rawText);
        const timeHtml = ts
          ? ` <span class="divider-time" title="${escapeHtml(timeIso)}">${escapeHtml(ts)}</span>`
          : '';
        parts.push(`<div class="divider"><span>${label}${timeHtml}</span></div>`);
      }
      continue;
    }

    const grouped = repliesByParent.get(i);
    if (grouped && grouped.length) {
      const sortedReplies = grouped
        .slice()
        .sort((a, b) => a.index - b.index)
        .map(r => r.msg);
      const replyHtml = sortedReplies.map(r => renderMessage(r, { isReply: true })).join('');
      const countLabel = sortedReplies.length == 1 ? '1 reply' : `${sortedReplies.length} replies`;
      parts.push(
        `<div class="thread">` +
        `<div class="thread-parent">${renderMessage(m)}</div>` +
        `<div class="thread-meta"><span>${countLabel}</span><button type="button" class="thread-toggle" data-thread-toggle>Collapse</button></div>` +
        `<div class="thread-replies">${replyHtml}</div>` +
        `</div>`,
      );
      continue;
    }

    parts.push(renderMessage(m));
  }

  const modal = `<div class="img-modal" id="img-modal" hidden>
    <button class="close" type="button" aria-label="Close">X</button>
    <img alt="full size image" />
  </div>`;

  const script = `<script>(()=>{const btn=document.querySelector('[data-toggle-compact]');const key='teamsExporterCompact';const apply=(c)=>{document.body.classList.toggle('compact',c);if(btn)btn.textContent=c?'Switch to expanded view':'Switch to compact view';};const stored=localStorage.getItem(key);let compact=stored==='1';apply(compact);if(btn){btn.addEventListener('click',()=>{compact=!compact;apply(compact);try{localStorage.setItem(key,compact?'1':'0');}catch(_){}});}document.querySelectorAll('.thread').forEach((thread)=>{const toggle=thread.querySelector('[data-thread-toggle]');if(!toggle)return;toggle.addEventListener('click',()=>{const collapsed=thread.classList.toggle('collapsed');toggle.textContent=collapsed?'Expand':'Collapse';});});const modal=document.getElementById('img-modal');const modalImg=modal?modal.querySelector('img'):null;const closeBtn=modal?modal.querySelector('.close'):null;const close=()=>{if(modal){modal.hidden=true;}};const open=(src,alt)=>{if(!modal||!modalImg)return;modalImg.src=src;modalImg.alt=alt||'image';modal.hidden=false;};if(closeBtn){closeBtn.addEventListener('click',close);}if(modal){modal.addEventListener('click',(e)=>{if(e.target===modal)close();});}document.addEventListener('keydown',(e)=>{if(e.key==='Escape')close();});document.body.addEventListener('click',(e)=>{const t=e.target;if(!(t instanceof Element))return;const img=t.closest('.att-img img');if(!img)return;if(img.closest('.att-video'))return;const src=img.getAttribute('data-full')||img.getAttribute('src');if(!src)return;open(src,img.getAttribute('alt')||'image');});})();</script>`;

  // Return as array of chunks to avoid "Invalid string length" on large exports.
  // Callers use Blob(chunks) directly instead of concatenating into one string.
  return [`<!doctype html><meta charset="utf-8">${style}${head}`, ...parts, modal, script];
}

/** Lazy check for Blob URL support (service workers may not have it). */
function canCreateBlobUrls(): boolean {
  try {
    return typeof URL.createObjectURL === 'function';
  } catch {
    return false;
  }
}

/** True when chrome.offscreen exists (Chromium MV3 only). */
function hasOffscreen(): boolean {
  try {
    return typeof chrome !== 'undefined' && !!chrome.offscreen;
  } catch {
    return false;
  }
}

// Above this payload size, an MV3 service worker mints the download URL via
// the offscreen document (a real blob: URL) instead of a base64 data: URL.
// A data: URL is a single string bounded by V8's ~512 MB limit, so large
// exports throw "Invalid string length" when base64-encoded (issue #27).
// Below the threshold the cheap inline data: URL is kept (no offscreen
// round-trip). 50 MB is well under the ~384 MB failure point and also
// sidesteps the memory cost of holding a large base64 string.
const OFFSCREEN_MINT_THRESHOLD = 50 * 1024 * 1024;

/** Create a downloadable URL from text (string or chunked string[]). Uses Blob URL when available. */
export function textToDownloadUrl(text: string | string[], mime: string): string {
  const parts = Array.isArray(text) ? text : [text];
  if (canCreateBlobUrls()) {
    try {
      return URL.createObjectURL(new Blob(parts, { type: mime }));
    } catch { /* fall through */ }
  }
  // Service worker fallback: encode chunks individually to avoid single huge string
  const encoder = new TextEncoder();
  const encoded = parts.map(p => encoder.encode(p));
  const totalLen = encoded.reduce((s, a) => s + a.length, 0);
  const bytes = new Uint8Array(totalLen);
  let offset = 0;
  for (const chunk of encoded) { bytes.set(chunk, offset); offset += chunk.length; }
  return binaryToDownloadUrl(bytes, mime);
}

/**
 * Create a downloadable URL from a Blob. Uses Blob URL when the API
 * is available (everywhere except some Chrome MV3 SW configurations);
 * falls back to a base64 data: URL by reading the blob bytes when it
 * isn't. The data: URL path adds ~33% memory overhead and won't scale
 * past a few hundred MB of payload, but it works wherever blob URLs
 * fail. Used by the zip-output download paths (per-chat-zip,
 * html-zip, bundle-outer) which all receive a Blob directly from
 * buildZipAsync.
 */
export async function blobToDownloadUrl(blob: Blob, mime: string): Promise<string> {
  if (canCreateBlobUrls()) {
    try {
      return URL.createObjectURL(blob);
    } catch { /* fall through */ }
  }
  // MV3 service worker (no createObjectURL). For large payloads a base64
  // data: URL would exceed V8's max string length and throw "Invalid
  // string length" (issue #27), so mint a real blob: URL in the offscreen
  // document. On any offscreen failure we fall through to the data: URL,
  // which is no worse than the previous behavior.
  if (hasOffscreen() && blob.size >= OFFSCREEN_MINT_THRESHOLD) {
    const url = await mintBlobUrlViaOffscreen(blob);
    if (url) return url;
  }
  const buf = await blob.arrayBuffer();
  return binaryToDownloadUrl(new Uint8Array(buf), mime);
}

/** Create a downloadable URL from binary data. Uses Blob URL when available, data URL in service workers. */
export function binaryToDownloadUrl(data: Uint8Array, mime: string): string {
  if (canCreateBlobUrls()) {
    try {
      return URL.createObjectURL(new Blob([data as BlobPart], { type: mime }));
    } catch { /* fall through */ }
  }
  // Service worker fallback: chunked base64 data URL
  // Process in small chunks to avoid string length limits
  const CHUNK = 32768;
  const binaryChunks: string[] = [];
  for (let i = 0; i < data.length; i += CHUNK) {
    const chunk = data.subarray(i, i + CHUNK);
    binaryChunks.push(String.fromCharCode.apply(null, chunk as unknown as number[]));
  }
  return `data:${mime};base64,${btoa(binaryChunks.join(''))}`;
}

/** Revoke a URL if it's a Blob URL (no-op for data URLs). */
export function revokeDownloadUrl(url: string) {
  if (!url.startsWith('blob:')) return;
  if (canCreateBlobUrls()) {
    try { URL.revokeObjectURL(url); } catch { /* noop */ }
    return;
  }
  // SW path: the blob: URL was minted in the offscreen document, so it
  // must be revoked there. Fire-and-forget.
  void revokeBlobUrlViaOffscreen(url);
}
