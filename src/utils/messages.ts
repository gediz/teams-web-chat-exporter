import type { AggregatedItem, Attachment, ExportMessage } from '../types/shared';
import { formatDayLabelForExport, parseTimeStamp } from './time';

export const makeDayDivider = (dayKey: number, ts: number): AggregatedItem => {
  const label = formatDayLabelForExport(ts);
  return {
    message: {
      id: `day-${dayKey}`,
      author: '[system]',
      timestamp: '',
      text: label,
      reactions: [],
      attachments: [],
      edited: false,
      avatar: null,
      replyTo: null,
      system: true,
    },
    orderKey: ts,
    tsMs: ts,
    kind: 'day-divider',
    timeLabel: label,
  };
};

// Windows reserved names (case-insensitive, with or without extension).
// Using one of these as a base name causes chrome.downloads to fail with
// "filename must not contain illegal characters" on Windows.
const WINDOWS_RESERVED = /^(con|prn|aux|nul|com[1-9]|lpt[1-9])(\..*)?$/i;

export const sanitizeBase = (name: string | null | undefined): string => {
  const raw = (name || 'teams-chat').toString();

  let cleaned = raw.replace(/^\(\d+\)\s+/, '');
  cleaned = cleaned.replace(/\s*\|\s*Microsoft Teams.*$/i, '');
  cleaned = cleaned.replace(/\s*\|\s*Teams.*$/i, '');
  cleaned = cleaned.replace(/\s*-\s*Microsoft Teams.*$/i, '');

  const parts = cleaned.split('|').map(p => p.trim());
  if (parts.length > 1) {
    cleaned = parts.find(p => p.length > 0) || parts[0];
  }

  // Strip control chars + characters illegal on common filesystems.
  // Includes DEL (\x7F) alongside the C0 range.
  cleaned = cleaned.replace(/[<>:"/\\|?*\x00-\x1F\x7F]/g, '-');
  // Collapse whitespace.
  cleaned = cleaned.replace(/\s+/g, ' ').trim();
  // Strip leading dots (prevents Unix "hidden file" names like ".config")
  // and leading dashes (avoids names that look like CLI flags).
  cleaned = cleaned.replace(/^[.\-\s]+/, '');
  // Trailing dots/spaces are illegal on Windows NTFS.
  cleaned = cleaned.replace(/[. ]+$/g, '');
  // Collapse runs of "..", which otherwise reintroduce path traversal in
  // contexts that glue a subpath (`bundle.zip/<name>/`).
  cleaned = cleaned.replace(/\.{2,}/g, '.');

  // Reject Windows reserved device names by suffixing. Matching is done
  // on the whole cleaned string; `CON.json` would also be reserved once
  // an extension is appended, but the extension is added *after* we
  // return, so we guard the base alone.
  if (WINDOWS_RESERVED.test(cleaned)) {
    cleaned = `${cleaned}_`;
  }

  // Truncate before the final empty check so a pathological input of
  // all-whitespace or all-dots still falls back to the default.
  cleaned = cleaned.slice(0, 80);

  return cleaned || 'teams-chat';
};

export const formatRangeLabel = (startISO?: string | null, endISO?: string | null): string | null => {
  if (!startISO && !endISO) return null;
  const fmt = new Intl.DateTimeFormat(undefined, { dateStyle: 'medium', timeStyle: 'short' });
  const start = startISO ? fmt.format(new Date(startISO)) : null;
  const end = endISO ? fmt.format(new Date(endISO)) : null;
  if (start && end) return `${start} → ${end}`;
  if (start) return `Since ${start}`;
  if (end) return `Until ${end}`;
  return null;
};

export const mergeAttachments = (existing: Attachment[] | undefined, next: Attachment[] | undefined): Attachment[] => {
  const map = new Map<string, Attachment>();
  const mergeOne = (att: Attachment | undefined) => {
    if (!att) return;
    const key = `${att.href || ''}@@${att.label || ''}`;
    const prev = map.get(key) || {};
    map.set(key, { ...prev, ...att });
  };
  (existing || []).forEach(mergeOne);
  (next || []).forEach(mergeOne);
  return Array.from(map.values());
};

export const ensureMessageTs = (entry: AggregatedItem) => {
  const ts = entry.anchorTs ?? entry.tsMs ?? (entry.message?.timestamp ? parseTimeStamp(entry.message.timestamp) : null);
  return ts;
};
