import type { AggregatedItem } from '../types/shared';
import { formatDayLabelForExport } from './time';

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
  // Strip zero-width and bidirectional formatting characters that
  // sometimes appear invisibly in chat names (often pasted from rich
  // text). Chrome's chrome.downloads.download rejects filenames
  // containing these as "filename must not contain illegal
  // characters" (issue #21). Removed entirely (not replaced with a
  // dash) since they have no visible representation — substituting
  // would produce names like "Jane---Doe" for what looks like
  // "Jane Doe". Covers:
  //   U+200B–U+200F  zero-width space, ZWNJ, ZWJ, LRM, RLM
  //   U+202A–U+202E  directional formatting overrides (LRE/RLE/PDF/LRO/RLO)
  //   U+2060–U+2064  word joiner + invisible math operators
  //   U+FEFF         BOM / zero-width no-break space
  cleaned = cleaned.replace(/[​-‏‪-‮⁠-⁤﻿]/g, '');
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

