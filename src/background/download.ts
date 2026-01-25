import { normalizeAvatars, removeAvatars, textToBlobUrl, textToDataUrl, toCSV, toHTML } from './builders';
import { formatRangeLabel, sanitizeBase } from '../utils/messages';
import { buildZip } from './zip';
import type { Attachment, ExportMessage, ExportMeta } from '../types/shared';

type DownloadsApi = Pick<typeof chrome.downloads, 'download'>;

type ImageMode = 'none' | 'data-url' | 'files';

type InlineImage = { filename: string; dataUrl: string };

export type BuiltExport = {
  baseFolder: string;
  filename: string;
  content: string;
  mime: string;
  inlineImages: InlineImage[];
};

export type BuildDownloadDeps = {
  downloads: DownloadsApi;
  isFirefox: boolean;
  onStatus?: (payload: { phase?: string; message?: string; filename?: string; messages?: number }) => void;
};

type BuildExportOptions = {
  messages?: ExportMessage[];
  meta?: ExportMeta;
  format?: 'json' | 'csv' | 'html' | 'txt';
  saveAs?: boolean;
  embedAvatars?: boolean;
  downloadImages?: boolean;
};

const DATA_URL_RE = /^data:([^;]+);base64,(.*)$/i;

function parseDataUrl(dataUrl: string) {
  const match = dataUrl.match(DATA_URL_RE);
  if (!match) return null;
  return { mime: match[1].toLowerCase(), data: match[2] };
}

function dataUrlToBytes(dataUrl: string): Uint8Array | null {
  const parsed = parseDataUrl(dataUrl);
  if (!parsed) return null;
  const bin = atob(parsed.data);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

function bytesToDataUrl(bytes: Uint8Array, mime: string): string {
  const chunkSize = 0x8000;
  let bin = '';
  for (let i = 0; i < bytes.length; i += chunkSize) {
    bin += String.fromCharCode(...bytes.subarray(i, i + chunkSize));
  }
  return `data:${mime};base64,${btoa(bin)}`;
}

function guessExtension(label?: string | null) {
  const match = (label || '').trim().match(/\.([A-Za-z0-9]{1,6})$/);
  return match ? match[1].toLowerCase() : null;
}

function mimeToExtension(mime: string) {
  if (mime === 'image/jpeg' || mime === 'image/jpg') return 'jpg';
  if (mime === 'image/png') return 'png';
  if (mime === 'image/gif') return 'gif';
  if (mime === 'image/webp') return 'webp';
  if (mime === 'image/bmp') return 'bmp';
  if (mime === 'image/svg+xml') return 'svg';
  return 'png';
}

function ensureUniqueName(name: string, used: Set<string>) {
  const canonical = name.toLowerCase();
  if (!used.has(canonical)) {
    used.add(canonical);
    return name;
  }
  const match = name.match(/^(.*?)(\.[^.]+)?$/);
  const base = match?.[1] || name;
  const ext = match?.[2] || '';
  let idx = 2;
  let candidate = `${base}-${idx}${ext}`;
  while (used.has(candidate.toLowerCase())) {
    idx += 1;
    candidate = `${base}-${idx}${ext}`;
  }
  used.add(candidate.toLowerCase());
  return candidate;
}

function stripAttachmentDataUrl(att: Attachment) {
  const { dataUrl, ...rest } = att;
  return rest;
}

function stripMessageDataUrls(messages: ExportMessage[]) {
  return messages.map(message => {
    const atts = Array.isArray(message.attachments) ? message.attachments : null;
    if (!atts) return message;
    const nextAtts = atts.map(att => stripAttachmentDataUrl(att));
    return { ...message, attachments: nextAtts };
  });
}

function applyInlineImages(messages: ExportMessage[], mode: ImageMode) {
  const inlineImages: InlineImage[] = [];
  if (mode === 'none') {
    return { messages: stripMessageDataUrls(messages), inlineImages };
  }

  const dataToName = new Map<string, string>();
  const usedNames = new Set<string>();
  let counter = 0;

  const rewriteAttachment = (att: Attachment): Attachment => {
    const dataUrl = att.dataUrl || (att.href && att.href.startsWith('data:') ? att.href : null);
    const base = stripAttachmentDataUrl(att);
    if (!dataUrl) return base;
    if (mode === 'data-url') {
      return { ...base, href: dataUrl };
    }

    let filename = dataToName.get(dataUrl);
    if (!filename) {
      counter += 1;
      const parsed = parseDataUrl(dataUrl);
      const ext = guessExtension(att.label) || (parsed ? mimeToExtension(parsed.mime) : 'png');
      const labelBase = (att.label || '').replace(/\.[^.]+$/, '').trim();
      const fallback = `image-${counter}`;
      const baseName = sanitizeBase(labelBase || fallback) || fallback;
      filename = ensureUniqueName(`${baseName}.${ext}`, usedNames);
      dataToName.set(dataUrl, filename);
      inlineImages.push({ filename: `images/${filename}`, dataUrl });
    }
    return { ...base, href: `images/${filename}` };
  };

  const nextMessages = messages.map(message => {
    const atts = Array.isArray(message.attachments) ? message.attachments : null;
    if (!atts || !atts.length) return message;
    const nextAtts = atts.map(att => rewriteAttachment(att));
    return { ...message, attachments: nextAtts };
  });

  return { messages: nextMessages, inlineImages };
}

function prepareMessages(messages: ExportMessage[], meta: ExportMeta, format: BuildExportOptions['format'], embedAvatars: boolean) {
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
        const msgsWithUrls = messages.map(m => (m.avatarUrl ? { ...m, avatar: m.avatarUrl } : m));
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

  return { processedMessages, enrichedMeta };
}

function buildExportInternal(options: BuildExportOptions, imageMode: ImageMode): BuiltExport {
  const { messages = [], meta = {}, format = 'json', embedAvatars = false } = options;
  const { processedMessages, enrichedMeta } = prepareMessages(messages, meta, format, embedAvatars);

  const rangeLabel = formatRangeLabel(meta.startAt, meta.endAt);
  if (rangeLabel) enrichedMeta.timeRange = rangeLabel;

  const baseTitle = sanitizeBase(enrichedMeta.title || 'UnknownChat');
  const stamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19).replace('T', '_');
  const base = `TeamsExport_${baseTitle}_${stamp}`;
  const baseFolder = base;

  let filename = base;
  let mime = 'application/json';
  let content = '';
  let inlineImages: InlineImage[] = [];
  let finalMessages = processedMessages;

  if (format === 'html') {
    const mode: ImageMode = imageMode;
    const applied = applyInlineImages(processedMessages, mode);
    finalMessages = applied.messages;
    inlineImages = applied.inlineImages;
    filename = mode === 'files' ? 'index.html' : `${base}.html`;
    mime = 'text/html';
    content = toHTML(finalMessages, { ...enrichedMeta, count: messages.length });
  } else if (format === 'csv') {
    filename = `${base}.csv`;
    mime = 'text/csv';
    finalMessages = stripMessageDataUrls(processedMessages);
    content = toCSV(finalMessages);
  } else if (format === 'txt') {
    filename = `${base}.txt`;
    mime = 'text/plain';
    finalMessages = stripMessageDataUrls(processedMessages);
    content = toPlainText(finalMessages);
  } else {
    filename = `${base}.json`;
    mime = 'application/json';
    finalMessages = stripMessageDataUrls(processedMessages);
    const payload = { meta: { ...enrichedMeta, count: messages.length }, messages: finalMessages };
    content = JSON.stringify(payload, null, 2);
  }

  return { baseFolder, filename, content, mime, inlineImages };
}

export function buildExport(options: BuildExportOptions): BuiltExport {
  const format = options.format || 'json';
  const downloadImages = Boolean(options.downloadImages);
  const mode: ImageMode = format === 'html' && downloadImages ? 'files' : 'none';
  return buildExportInternal(options, mode);
}

export async function buildAndDownload(
  deps: BuildDownloadDeps,
  { messages = [], meta = {}, format = 'json', saveAs = true, embedAvatars = false, downloadImages = false }: BuildExportOptions,
) {
  const { downloads, isFirefox, onStatus } = deps;
  const mode: ImageMode = format === 'html' && downloadImages ? 'data-url' : 'none';
  const built = buildExportInternal({ messages, meta, format, embedAvatars, downloadImages }, mode);

  onStatus?.({ phase: 'build', filename: built.filename, messages: messages.length });

  const url = isFirefox ? textToBlobUrl(built.content, built.mime) : textToDataUrl(built.content, built.mime);
  try {
    const id = await downloads.download({ url, filename: built.filename, saveAs });
    if (isFirefox) setTimeout(() => URL.revokeObjectURL(url), 60000);
    return { ok: true, filename: built.filename, id };
  } catch (e: any) {
    const safe = `${sanitizeBase('teams-chat')}-${Date.now()}.${format === 'html' ? 'html' : format === 'csv' ? 'csv' : format === 'txt' ? 'txt' : 'json'}`;
    if (isFirefox) URL.revokeObjectURL(url);
    try {
      const url2 = isFirefox ? textToBlobUrl(built.content, built.mime) : textToDataUrl(built.content, built.mime);
      const id2 = await downloads.download({ url: url2, filename: safe, saveAs });
      if (isFirefox) setTimeout(() => URL.revokeObjectURL(url2), 60000);
      return { ok: true, filename: safe, id: id2 };
    } catch (e2: any) {
      throw new Error(e2?.message || String(e2));
    }
  }
}

export async function buildAndDownloadZip(
  deps: BuildDownloadDeps,
  { messages = [], meta = {}, embedAvatars = false, downloadImages = false }: BuildExportOptions,
) {
  const { downloads, isFirefox, onStatus } = deps;
  const built = buildExportInternal(
    { messages, meta, format: 'html', embedAvatars, downloadImages },
    downloadImages ? 'files' : 'none',
  );

  onStatus?.({ phase: 'build', filename: `${built.baseFolder}.zip`, messages: messages.length });

  const encoder = new TextEncoder();
  const files: { path: string; data: Uint8Array }[] = [];
  files.push({
    path: `${built.baseFolder}/index.html`,
    data: encoder.encode(built.content),
  });

  for (const img of built.inlineImages) {
    const bytes = dataUrlToBytes(img.dataUrl);
    if (!bytes) continue;
    files.push({
      path: `${built.baseFolder}/${img.filename}`,
      data: bytes,
    });
  }

  const zipBytes = buildZip(files);
  const mime = 'application/zip';
  const zipName = `${built.baseFolder}.zip`;
  const url = isFirefox ? URL.createObjectURL(new Blob([zipBytes], { type: mime })) : bytesToDataUrl(zipBytes, mime);
  try {
    const id = await downloads.download({ url, filename: zipName, saveAs: true });
    if (isFirefox) setTimeout(() => URL.revokeObjectURL(url), 60000);
    return { ok: true, filename: zipName, id };
  } catch (e: any) {
    if (isFirefox) URL.revokeObjectURL(url);
    throw new Error(e?.message || String(e));
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
