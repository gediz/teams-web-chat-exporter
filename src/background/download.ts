import { normalizeAvatars, removeAvatars, textToBlobUrl, textToDataUrl, toCSV, toHTML } from './builders';
import { formatRangeLabel, sanitizeBase } from '../utils/messages';
import type { ExportMessage, ExportMeta } from '../types/shared';

type DownloadsApi = Pick<typeof chrome.downloads, 'download'>;

export type BuildDownloadDeps = {
  downloads: DownloadsApi;
  isFirefox: boolean;
};

export async function buildAndDownload(
  deps: BuildDownloadDeps,
  { messages = [], meta = {}, format = 'json', saveAs = true, embedAvatars = false }: { messages?: ExportMessage[]; meta?: ExportMeta; format?: 'json' | 'csv' | 'html' | 'txt'; saveAs?: boolean; embedAvatars?: boolean },
) {
  const { downloads, isFirefox } = deps;

  // Process avatars based on format and embedAvatars option
  let processedMessages = messages;
  let enrichedMeta = { ...meta };

  if (embedAvatars && (format === 'json' || format === 'html')) {
    // Avatars are already base64 data URLs from content script
    // Just need to normalize them for JSON format
    if (format === 'json') {
      // Build avatar map using original URLs for ID extraction
      const avatarMap = new Map<string, string | null>();
      for (const m of messages) {
        if (m.avatar && m.avatar.startsWith('data:') && m.avatarUrl) {
          // Use original HTTP URL as key for proper ID extraction
          avatarMap.set(m.avatarUrl, m.avatar);
        }
      }

      if (avatarMap.size > 0) {
        // Temporarily restore avatarUrl to avatar field for normalizeAvatars
        const msgsWithUrls = messages.map(m =>
          m.avatarUrl ? { ...m, avatar: m.avatarUrl } : m
        );
        const { messages: normalized, avatars } = normalizeAvatars(msgsWithUrls, avatarMap);
        // Remove avatarUrl field from final output
        processedMessages = normalized.map(m => {
          const { avatarUrl, ...rest } = m as any;
          return rest;
        });
        enrichedMeta.avatars = avatars;
      } else {
        // No avatars found, remove avatarUrl field
        processedMessages = messages.map(m => {
          const { avatarUrl, ...rest } = m as any;
          return rest;
        });
      }
    } else {
      // For HTML: messages already have base64, just remove avatarUrl field
      processedMessages = messages.map(m => {
        const { avatarUrl, ...rest } = m as any;
        return rest;
      });
    }
  } else if (!embedAvatars) {
    // Remove avatars entirely when option is disabled
    processedMessages = removeAvatars(messages).map(m => {
      const { avatarUrl, ...rest } = m as any;
      return rest;
    });
  }

  const rangeLabel = formatRangeLabel(meta.startAt, meta.endAt);
  if (rangeLabel) enrichedMeta.timeRange = rangeLabel;

  const baseTitle = sanitizeBase(enrichedMeta.title || 'UnknownChat');
  const stamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19).replace('T', '_');
  const base = `TeamsExport_${baseTitle}_${stamp}`;
  let filename: string;
  let mime: string;
  let content: string;

  if (format === 'json') {
    filename = `${base}.json`;
    mime = 'application/json';
    const payload = { meta: { ...enrichedMeta, count: messages.length }, messages: processedMessages };
    content = JSON.stringify(payload, null, 2);
  } else if (format === 'csv') {
    filename = `${base}.csv`;
    mime = 'text/csv';
    content = toCSV(messages);
  } else if (format === 'html') {
    filename = `${base}.html`;
    mime = 'text/html';
    content = toHTML(processedMessages, { ...enrichedMeta, count: messages.length });
  } else if (format === 'txt') {
    filename = `${base}.txt`;
    mime = 'text/plain';
    content = toPlainText(messages);
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

function toPlainText(messages: ExportMessage[]) {
  const lines: string[] = [];
  for (const m of messages) {
    const ts = m.timestamp || '';
    const author = m.author || '[unknown]';
    const text = (m.text || '').replace(/\r\n/g, '\n').replace(/\n{2,}/g, '\n\n');
    lines.push(`[${ts}] ${author}: ${text}`);
  }
  return lines.join('\n');
}
