import { isoToLocalInput, localInputToISO } from './time';
import type { HistoryEntry } from '../types/shared';

export type OptionFormat = 'json' | 'csv' | 'html' | 'txt';
export type Theme = 'light' | 'dark';
export type ExportTarget = 'chat' | 'team';
// What the extension does automatically after a successful export:
// 'manual' -> the post-export tile shows Open/Show buttons; nothing else
//             happens until the user clicks. (default; labelled "Let me
//             decide" in the Settings UI).
// 'show'   -> auto-trigger chrome.downloads.show(id) on success. Works
//             whether the popup is open or not — show() doesn't need a
//             user gesture. Action buttons still appear on the tile.
//
// History: 'open' and 'none' used to be valid values. 'open' was dropped
// because chrome.downloads.open() requires user activation and is not
// reliably callable from background or a closed popup in MV3. 'none' was
// dropped because it behaved identically to the default ("do nothing
// automatic"). 'ask' was renamed to 'manual' because nothing actually
// prompts — we just render the two action buttons.
export type AfterExport = 'manual' | 'show';

export type Options = {
  lang?: string;
  startAt: string;
  startAtISO: string;
  endAt: string;
  endAtISO: string;
  exportTarget: ExportTarget;
  format: OptionFormat;
  includeReplies: boolean;
  includeReactions: boolean;
  includeSystem: boolean;
  embedAvatars: boolean;
  downloadImages: boolean;
  showHud: boolean;
  theme: Theme;
  afterExport: AfterExport;
};

export type StoredError = { message: string; timestamp?: number };

export const OPTIONS_STORAGE_KEY = 'teamsExporterOptions';
export const ERROR_STORAGE_KEY = 'teamsExporterLastError';
export const HISTORY_STORAGE_KEY = 'teamsExporterHistory';
// Last time the user opened the History page. Entries with savedAt > this
// value are considered "new" and trigger the breathing dot on the history
// icon. Stored under its own key so it survives history clears.
export const HISTORY_VIEWED_KEY = 'teamsExporterHistoryViewedAt';

export const DEFAULT_OPTIONS: Options = {
  lang: 'en',
  startAt: '',
  startAtISO: '',
  endAt: '',
  endAtISO: '',
  exportTarget: 'chat',
  format: 'json',
  includeReplies: true,
  includeReactions: true,
  includeSystem: false,
  embedAvatars: false,
  downloadImages: false,
  showHud: false,
  theme: 'light',
  afterExport: 'manual',
};

type StorageArea = Pick<chrome.storage.StorageArea, 'get' | 'set' | 'remove'>;
type ExtensionStorage = { local: StorageArea };

const normalizeOptions = (raw: Partial<Options>, defaults: Options = DEFAULT_OPTIONS): Options => {
  const merged: Options = { ...defaults, ...raw };
  merged.startAt = merged.startAt || isoToLocalInput(merged.startAtISO);
  merged.endAt = merged.endAt || isoToLocalInput(merged.endAtISO);
  // Clamp afterExport to the current enum. A previously-stored value of
  // 'ask' / 'open' / 'none' falls back to the default so the Settings
  // dropdown doesn't render blank. This isn't a formal migration — more
  // of a defensive floor for an out-of-range value.
  if (merged.afterExport !== 'manual' && merged.afterExport !== 'show') {
    merged.afterExport = defaults.afterExport;
  }
  return merged;
};

export async function loadOptions(storage: ExtensionStorage, defaults: Options = DEFAULT_OPTIONS): Promise<Options> {
  try {
    const stored = await storage.local.get(OPTIONS_STORAGE_KEY);
    const loaded = (stored?.[OPTIONS_STORAGE_KEY] || {}) as Partial<Options>;
    return normalizeOptions(loaded, defaults);
  } catch {
    return { ...defaults };
  }
}

export async function saveOptions(
  storage: ExtensionStorage,
  options: Options,
  defaults: Options = DEFAULT_OPTIONS,
): Promise<Options> {
  const startISO = localInputToISO(options.startAt);
  const endISO = localInputToISO(options.endAt);
  const payload: Options = {
    ...normalizeOptions(options, defaults),
    startAtISO: startISO || '',
    endAtISO: endISO || '',
  };
  try {
    await storage.local.set({ [OPTIONS_STORAGE_KEY]: payload });
  } catch {
    // ignore
  }
  return payload;
}

export function validateRange(options: Pick<Options, 'startAt' | 'endAt'>): { startISO: string | null; endISO: string | null } {
  const rawStart = (options.startAt || '').trim();
  const rawEnd = (options.endAt || '').trim();
  const startISO = rawStart ? localInputToISO(rawStart) : null;
  if (rawStart && !startISO) {
    throw new Error('Enter a valid start date/time.');
  }
  const endISO = rawEnd ? localInputToISO(rawEnd) : null;
  if (rawEnd && !endISO) {
    throw new Error('Enter a valid end date/time.');
  }
  if (startISO && endISO) {
    const startMs = Date.parse(startISO);
    const endMs = Date.parse(endISO);
    if (!Number.isNaN(startMs) && !Number.isNaN(endMs) && startMs > endMs) {
      throw new Error('Start date must be before end date.');
    }
  }
  return { startISO, endISO };
}

export async function loadLastError(storage: ExtensionStorage): Promise<StoredError | null> {
  try {
    const res = await storage.local.get(ERROR_STORAGE_KEY);
    return (res?.[ERROR_STORAGE_KEY] as StoredError) || null;
  } catch {
    return null;
  }
}

export async function persistErrorMessage(storage: ExtensionStorage, message: string) {
  try {
    await storage.local.set({ [ERROR_STORAGE_KEY]: { message, timestamp: Date.now() } });
  } catch {
    // ignore
  }
}

export async function clearLastError(storage: ExtensionStorage) {
  try {
    await storage.local.remove(ERROR_STORAGE_KEY);
  } catch {
    // ignore
  }
}

// =====================================================================
// Export history — array stored under HISTORY_STORAGE_KEY, newest first.
// Each entry is metadata-only (~250 bytes). No formal cap; the user can
// "Clear all" or remove rows individually from the History page.
// =====================================================================

const isHistoryEntry = (raw: unknown): raw is HistoryEntry => {
  if (!raw || typeof raw !== 'object') return false;
  const e = raw as HistoryEntry;
  return typeof e.id === 'string'
      && typeof e.tabId === 'number'
      && typeof e.savedAt === 'number'
      && (e.kind === 'success' || e.kind === 'cancelled');
};

export async function loadHistory(storage: ExtensionStorage): Promise<HistoryEntry[]> {
  try {
    const res = await storage.local.get(HISTORY_STORAGE_KEY);
    const raw = res?.[HISTORY_STORAGE_KEY];
    if (!Array.isArray(raw)) return [];
    // Drop any malformed entries silently; better to lose one than crash the page.
    return raw.filter(isHistoryEntry);
  } catch {
    return [];
  }
}

export async function appendHistoryEntry(storage: ExtensionStorage, entry: HistoryEntry): Promise<void> {
  try {
    const current = await loadHistory(storage);
    // Newest first.
    const next = [entry, ...current];
    await storage.local.set({ [HISTORY_STORAGE_KEY]: next });
  } catch {
    // ignore — best-effort; losing one entry on storage error is acceptable
  }
}

export async function removeHistoryEntry(storage: ExtensionStorage, id: string): Promise<void> {
  try {
    const current = await loadHistory(storage);
    const next = current.filter(e => e.id !== id);
    await storage.local.set({ [HISTORY_STORAGE_KEY]: next });
  } catch {
    // ignore
  }
}

// Patch a single entry in place. Used to mark fileExists once we observe
// the file is missing — that observation is durable and shouldn't be lost
// when the popup closes.
export async function updateHistoryEntry(
  storage: ExtensionStorage,
  id: string,
  patch: Partial<HistoryEntry>,
): Promise<void> {
  try {
    const current = await loadHistory(storage);
    let changed = false;
    const next = current.map(e => {
      if (e.id !== id) return e;
      changed = true;
      return { ...e, ...patch };
    });
    if (changed) {
      await storage.local.set({ [HISTORY_STORAGE_KEY]: next });
    }
  } catch {
    // ignore
  }
}

export async function clearHistory(storage: ExtensionStorage): Promise<void> {
  try {
    await storage.local.remove(HISTORY_STORAGE_KEY);
  } catch {
    // ignore
  }
}

// Mark all current history as "seen" (clears the new-entry dot on the
// history icon). Stored timestamp is compared against each entry's savedAt
// to compute "any new entries?" without mutating the entries themselves.
export async function markHistorySeen(storage: ExtensionStorage): Promise<void> {
  try {
    await storage.local.set({ [HISTORY_VIEWED_KEY]: Date.now() });
  } catch {
    // ignore
  }
}

export async function loadHistoryViewedAt(storage: ExtensionStorage): Promise<number> {
  try {
    const res = await storage.local.get(HISTORY_VIEWED_KEY);
    const raw = res?.[HISTORY_VIEWED_KEY];
    return typeof raw === 'number' ? raw : 0;
  } catch {
    return 0;
  }
}
