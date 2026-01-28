import { isoToLocalInput, localInputToISO } from './time';

export type OptionFormat = 'json' | 'csv' | 'html' | 'txt';
export type Theme = 'light' | 'dark';
export type ExportTarget = 'chat' | 'team';

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
};

export type StoredError = { message: string; timestamp?: number };

export const OPTIONS_STORAGE_KEY = 'teamsExporterOptions';
export const ERROR_STORAGE_KEY = 'teamsExporterLastError';

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
};

type StorageArea = Pick<chrome.storage.StorageArea, 'get' | 'set' | 'remove'>;
type ExtensionStorage = { local: StorageArea };

const normalizeOptions = (raw: Partial<Options>, defaults: Options = DEFAULT_OPTIONS): Options => {
  const merged: Options = { ...defaults, ...raw };
  merged.startAt = merged.startAt || isoToLocalInput(merged.startAtISO);
  merged.endAt = merged.endAt || isoToLocalInput(merged.endAtISO);
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
