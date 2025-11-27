import type { ExportStatusPayload } from '../types/shared';

type BadgeAction = Pick<typeof chrome.action, 'setBadgeText' | 'setBadgeBackgroundColor'>;

export const BADGE_COLORS = {
  default: '#1d4ed8',
  empty: '#6b7280',
  success: '#16a34a',
  error: '#dc2626',
  progress: '#2563eb',
} as const;

export type BadgeProgress = { filteredSeen?: number; seen?: number; aggregated?: number; messagesVisible?: number };

const ONE_THOUSAND = 1000;
const ONE_MILLION = 1_000_000;
const THOUSAND_ROUND_THRESHOLD = 10_000;
const MILLION_ROUND_THRESHOLD = 10_000_000;

const formatBadgeCount = (value: string | number | null | undefined): string => {
  if (value === null || value === undefined) return '';
  const num = typeof value === 'number' ? value : Number(value);
  if (!Number.isFinite(num)) return String(value);
  const abs = Math.abs(num);
  const sign = num < 0 ? '-' : '';
  if (abs < ONE_THOUSAND) return `${num}`;
  if (abs < ONE_MILLION) {
    const scaled = (abs / ONE_THOUSAND).toFixed(abs >= THOUSAND_ROUND_THRESHOLD ? 0 : 1);
    return `${sign}${scaled.replace(/\.0$/, '')}k`;
  }
  const scaled = (abs / ONE_MILLION).toFixed(abs >= MILLION_ROUND_THRESHOLD ? 0 : 1);
  return `${sign}${scaled.replace(/\.0$/, '')}m`;
};

export const createBadgeManager = (action: BadgeAction) => {
  let clearTimer: ReturnType<typeof setTimeout> | null = null;

  const set = (text: string | number, color: string = BADGE_COLORS.default) => {
    try {
      let finalText: string | number | null | undefined = text;
      if (typeof finalText === 'number') {
        finalText = formatBadgeCount(finalText);
      } else if (typeof finalText === 'string') {
        const trimmed = finalText.trim();
        if (trimmed) {
          const numeric = Number(trimmed);
          if (Number.isFinite(numeric)) {
            finalText = formatBadgeCount(numeric);
          }
        }
      }
      action.setBadgeBackgroundColor({ color });
      const textValue = finalText == null ? '' : `${finalText}`;
      action.setBadgeText({ text: textValue });
    } catch {
      // ignore badge errors
    }
  };

  const reset = () => set('', BADGE_COLORS.default);

  const clearSoon = (delay = 0) => {
    if (clearTimer) clearTimeout(clearTimer);
    clearTimer = setTimeout(() => {
      reset();
      clearTimer = null;
    }, delay);
  };

  const updateForStatus = (payload: ExportStatusPayload) => {
    const phase = payload?.phase;
    if (!phase) return;
    try {
      if (phase === 'starting' || phase === 'scrape:start') {
        set('…', BADGE_COLORS.progress);
        return;
      }
      if (phase === 'scrape:complete') {
        const total = payload?.messages ?? payload?.messagesExtracted;
        if (typeof total === 'number') set(total);
        return;
      }
      if (phase === 'empty') {
        set('0', BADGE_COLORS.empty);
        clearSoon(2000);
        return;
      }
      if (phase === 'complete') {
        set('✔', BADGE_COLORS.success);
        clearSoon(2000);
        return;
      }
      if (phase === 'error') {
        set('!', BADGE_COLORS.error);
        clearSoon(3000);
        return;
      }
    } catch {
      // ignore badge errors
    }
  };

  const updateForProgress = (progress: BadgeProgress) => {
    if (!progress) return;
    const seen = progress.filteredSeen ?? progress.seen ?? progress.aggregated ?? progress.messagesVisible;
    if (typeof seen === 'number' && seen >= 0) {
      set(seen);
    }
  };

  return { set, reset, clearSoon, updateForStatus, updateForProgress };
};
