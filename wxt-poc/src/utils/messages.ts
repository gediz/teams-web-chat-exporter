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

export const sanitizeBase = (name: string | null | undefined): string => {
  const raw = (name || 'teams-chat').toString();
  const cleaned = raw.replace(/[<>:"/\\|?*\x00-\x1F]/g, '-').replace(/\s+/g, ' ').trim().replace(/[. ]+$/g, '');
  return (cleaned || 'teams-chat').slice(0, 80);
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
