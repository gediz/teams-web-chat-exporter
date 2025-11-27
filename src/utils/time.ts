export const MS_PER_DAY = 24 * 60 * 60 * 1000;

export function parseTimeStamp(value: string | null | undefined): number | null {
  if (!value) return null;
  const ts = Date.parse(value);
  if (!Number.isNaN(ts)) return ts;
  const normalized = value.replace(/ /g, 'T');
  const ts2 = Date.parse(normalized);
  return Number.isNaN(ts2) ? null : ts2;
}

export function startOfLocalDay(ts: number) {
  const d = new Date(ts);
  d.setHours(0, 0, 0, 0);
  return d.getTime();
}

export function formatDayLabelForExport(ts: number): string {
  const dayStart = startOfLocalDay(ts);
  const todayStart = startOfLocalDay(Date.now());
  const diffDays = Math.round((todayStart - dayStart) / MS_PER_DAY);

  if (diffDays === 0) return 'Today';
  if (diffDays === 1) return 'Yesterday';
  if (diffDays === -1) return 'Tomorrow';
  if (diffDays >= -6 && diffDays <= 6) {
    return new Intl.DateTimeFormat(undefined, { weekday: 'long' }).format(new Date(ts));
  }
  return new Intl.DateTimeFormat(undefined, { dateStyle: 'long' }).format(new Date(ts));
}

export function formatElapsed(ms: number) {
  const totalSeconds = Math.max(0, Math.floor(ms / 1000));
  const hours = Math.floor(totalSeconds / 3600);
  const minutes = Math.floor((totalSeconds % 3600) / 60);
  const seconds = totalSeconds % 60;
  if (hours > 0) {
    return `${hours}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;
  }
  return `${minutes}:${String(seconds).padStart(2, '0')}`;
}

export const formatElapsedSuffix = (ms: number) => ` — Elapsed: ${formatElapsed(ms)}`;

export function localInputToISO(localValue: string) {
  if (!localValue) return '';
  let normalized = localValue.trim();
  if (!normalized) return '';
  normalized = normalized.replace(/\//g, '-').replace(/\s+/g, ' ');
  if (/^\d{4}-\d{2}-\d{2}$/.test(normalized)) {
    normalized += ' 00:00';
  }
  if (normalized.includes(' ')) {
    normalized = normalized.replace(' ', 'T');
  }
  const date = new Date(normalized);
  if (Number.isNaN(date.getTime())) return '';
  return date.toISOString();
}

export function isoToLocalInput(isoValue: string) {
  if (!isoValue) return '';
  const date = new Date(isoValue);
  if (Number.isNaN(date.getTime())) return '';
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');
  return `${year}-${month}-${day} ${hours}:${minutes}`;
}
