import { embedAvatarsInRows, textToBlobUrl, textToDataUrl, toCSV, toHTML } from './builders';
import { formatRangeLabel, sanitizeBase } from '../utils/messages';
import type { BuildOptions, ExportMessage, ExportMeta } from '../types/shared';

type DownloadsApi = Pick<typeof chrome.downloads, 'download'>;

export type BuildDownloadDeps = {
  downloads: DownloadsApi;
  isFirefox: boolean;
};

export async function buildAndDownload(
  deps: BuildDownloadDeps,
  { messages = [], meta = {}, format = 'json', saveAs = true, embedAvatars = false }: { messages?: ExportMessage[]; meta?: ExportMeta; format?: 'json' | 'csv' | 'html'; saveAs?: boolean; embedAvatars?: boolean },
) {
  const { downloads, isFirefox } = deps;
  let rows = messages;
  if (format === 'html' && embedAvatars) {
    try {
      rows = await embedAvatarsInRows(messages);
    } catch {
      rows = messages;
    }
  }

  const rangeLabel = formatRangeLabel(meta.startAt, meta.endAt);
  const enrichedMeta = { ...meta };
  if (rangeLabel) enrichedMeta.timeRange = rangeLabel;

  const baseTitle = sanitizeBase(enrichedMeta.title || 'teams-chat');
  const stamp = new Date().toISOString().replace(/[:.]/g, '-');
  const base = `${baseTitle}-${stamp}`;
  let filename: string;
  let mime: string;
  let content: string;

  if (format === 'json') {
    filename = `${base}.json`;
    mime = 'application/json';
    const payload = { meta: { ...enrichedMeta, count: messages.length }, messages };
    content = JSON.stringify(payload, null, 2);
  } else if (format === 'csv') {
    filename = `${base}.csv`;
    mime = 'text/csv';
    content = toCSV(messages);
  } else if (format === 'html') {
    filename = `${base}.html`;
    mime = 'text/html';
    content = toHTML(rows, { ...enrichedMeta, count: messages.length });
  } else {
    throw new Error('Unknown format: ' + format);
  }

  const url = isFirefox ? textToBlobUrl(content, mime) : textToDataUrl(content, mime);
  try {
    const id = await downloads.download({ url, filename, saveAs });
    if (isFirefox) setTimeout(() => URL.revokeObjectURL(url), 60000);
    return { ok: true, filename, id };
  } catch (e: any) {
    const safe = `${sanitizeBase('teams-chat')}-${Date.now()}.${format === 'html' ? 'html' : format === 'csv' ? 'csv' : 'json'}`;
    if (isFirefox) URL.revokeObjectURL(url);
    try {
      const url2 = isFirefox ? textToBlobUrl(content, mime) : textToDataUrl(content, mime);
      const id2 = await downloads.download({ url: url2, filename: safe, saveAs });
      if (isFirefox) setTimeout(() => URL.revokeObjectURL(url2), 60000);
      return { ok: true, filename: safe, id: id2 };
    } catch (e2: any) {
      throw new Error(e2?.message || String(e2));
    }
  }
}
