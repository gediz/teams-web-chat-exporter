import { binaryToDownloadUrl, removeAvatars, revokeDownloadUrl, textToDownloadUrl, toCSV, toHTML } from './builders';
import { buildPdf } from './pdf';
import { formatRangeLabel, sanitizeBase } from '../utils/messages';
import { buildZip } from './zip';
import type { Attachment, ExportMessage, ExportMeta } from '../types/shared';

type DownloadsApi = Pick<typeof chrome.downloads, 'download'>;

type ImageMode = 'none' | 'data-url' | 'files';

type InlineImage = { filename: string; dataUrl: string };

type BuiltExport = {
  baseFolder: string;
  filename: string;
  content: string | string[];  // string[] for large HTML (chunked to avoid "Invalid string length")
  mime: string;
  inlineImages: InlineImage[];
};

export type BuildDownloadDeps = {
  downloads: DownloadsApi;
  isFirefox: boolean;
  onStatus?: (payload: { phase?: string; message?: string; filename?: string; messages?: number; messagesBuilt?: number; messagesTotal?: number }) => void;
};

type SingleFormat = 'json' | 'csv' | 'html' | 'txt' | 'pdf';
type AvatarMode = 'inline' | 'files';

// Subset of PdfOptions we pass through from popup to pdf.ts. Mirrors
// PdfOptions in src/background/pdf.ts — kept structural (not imported)
// because download.ts doesn't otherwise depend on pdf.ts's types.
type PdfKnobs = {
  pdfPageSize?: 'a4' | 'letter';
  pdfBodyFontSize?: number;
  pdfShowPageNumbers?: boolean;
  pdfIncludeAvatars?: boolean;
};

type BuildExportOptions = {
  messages?: ExportMessage[];
  meta?: ExportMeta;
  format?: SingleFormat;
  saveAs?: boolean;
  embedAvatars?: boolean;
  downloadImages?: boolean;
  // Only read by the zip paths; single-file HTML always inlines avatars
  // because one loose .html file can't reference a sibling folder.
  avatarMode?: AvatarMode;
} & PdfKnobs;

type BuildBundleOptions = {
  messages?: ExportMessage[];
  meta?: ExportMeta;
  formats: SingleFormat[];
  embedAvatars?: boolean;
  downloadImages?: boolean;
  avatarMode?: AvatarMode;
} & PdfKnobs;

// Translate PdfKnobs -> buildPdf's PdfOptions shape. The property rename
// keeps buildPdf's public surface narrow (it doesn't know about the
// outer "pdf*" prefix we use to distinguish these in BuildOptions).
function toPdfOptions(knobs: PdfKnobs) {
  return {
    pageSize: knobs.pdfPageSize,
    bodyFontSize: knobs.pdfBodyFontSize,
    showPageNumbers: knobs.pdfShowPageNumbers,
    includeAvatars: knobs.pdfIncludeAvatars,
  };
}

// Compute the canonical `TeamsExport_<chat>_<stamp>` base name. Shared
// between buildExportInternal (text formats) and the PDF path so a
// single-format PDF export and a bundle containing PDF use the same
// filename scheme. The stamp collapses to second resolution; within a
// single export call all builders see the same value as long as they
// share this helper's result.
function computeBaseName(meta: ExportMeta): string {
  const baseTitle = sanitizeBase(meta.title || 'UnknownChat');
  const stamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19).replace('T', '_');
  return `TeamsExport_${baseTitle}_${stamp}`;
}

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

// Split the embedded-avatar map out of the HTML export meta into a set
// of `avatars/<id>.<ext>` files. Used only on zip paths (HTML.zip,
// bundle.zip) so the HTML stays small and extractors see a
// human-browsable avatars folder.
//
// Returns:
//   - files: list of { path, data } entries ready to drop into the zip
//   - meta:  a shallow clone with `_avatarsAsFiles` + `_avatarFileMap`
//            set so toHTML knows to emit `url("avatars/<file>")`
//            instead of inlining base64
// If no avatars are present we return an empty list and a pass-through
// meta (toHTML still renders just fine with no avatar map at all).
function collectAvatarFiles(meta: ExportMeta): { files: InlineImage[]; meta: ExportMeta } {
  const avatars = meta.avatars;
  if (!avatars || !Object.keys(avatars).length) return { files: [], meta };
  const files: InlineImage[] = [];
  const fileMap: Record<string, string> = {};
  for (const [id, dataUrl] of Object.entries(avatars)) {
    const parsed = parseDataUrl(dataUrl);
    if (!parsed) continue;
    const ext = mimeToExtension(parsed.mime);
    const filename = `${id}.${ext}`;
    fileMap[id] = filename;
    files.push({ filename: `avatars/${filename}`, dataUrl });
  }
  // Mutate-free: fresh meta with the two routing flags set. Keeping the
  // original avatars map on the side (untouched) means JSON co-outputs
  // in a bundle still see the inline form if they want it.
  const nextMeta: ExportMeta = {
    ...meta,
    _avatarsAsFiles: true,
    _avatarFileMap: fileMap,
  } as ExportMeta;
  return { files, meta: nextMeta };
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
    // Keep audio data URLs inline (not as image files)
    if (att.type === 'audio') {
      return { ...base, dataUrl };
    }
    if (mode === 'data-url') {
      return { ...base, href: dataUrl };
    }

    let filename = dataToName.get(dataUrl);
    if (!filename) {
      counter += 1;
      const parsed = parseDataUrl(dataUrl);
      const isPreview = att.kind === 'preview';
      const ext = (isPreview ? null : guessExtension(att.label)) || (parsed ? mimeToExtension(parsed.mime) : 'png');
      const labelBase = isPreview ? '' : (att.label || '').replace(/\.[^.]+$/, '').trim();
      const fallback = isPreview ? `preview-${counter}` : `image-${counter}`;
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
  const avatarMap = meta.avatars || {};

  if (embedAvatars && (format === 'json' || format === 'html')) {
    if (format === 'json') {
      // Content script already normalized: messages have avatarId, meta has avatars map
      // Just pass through — avatarId references and meta.avatars are already set
      processedMessages = messages;
    } else {
      // HTML: keep avatarId on messages — the avatar data URLs will be embedded
      // via CSS background-image classes in the HTML head (much more efficient than
      // duplicating a ~50KB data URL in every message's <img> tag).
      // The toHTML function reads avatarId and applies the CSS class.
      processedMessages = messages;
      // Keep avatars map in meta for toHTML to generate CSS
    }
  } else if (!embedAvatars) {
    // Remove avatars entirely when option is disabled
    processedMessages = removeAvatars(messages);
    delete enrichedMeta.avatars;
  }

  return { processedMessages, enrichedMeta };
}

function buildExportInternal(options: BuildExportOptions, imageMode: ImageMode): BuiltExport {
  const { messages = [], meta = {}, format = 'json', embedAvatars = false } = options;
  const { processedMessages, enrichedMeta } = prepareMessages(messages, meta, format, embedAvatars);

  const rangeLabel = formatRangeLabel(meta.startAt, meta.endAt);
  if (rangeLabel) enrichedMeta.timeRange = rangeLabel;

  const base = computeBaseName(enrichedMeta);
  const baseFolder = base;

  let filename = base;
  let mime = 'application/json';
  let content: string | string[] = '';
  let inlineImages: InlineImage[] = [];
  let finalMessages = processedMessages;

  if (format === 'html') {
    const mode: ImageMode = imageMode;
    const applied = applyInlineImages(processedMessages, mode);
    finalMessages = applied.messages;
    inlineImages = applied.inlineImages;
    filename = mode === 'files' ? 'index.html' : `${base}.html`;
    mime = 'text/html';
    content = toHTML(finalMessages, { ...enrichedMeta, count: messages.length }); // returns string[]
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

export async function buildAndDownload(
  deps: BuildDownloadDeps,
  options: BuildExportOptions,
) {
  const { messages = [], meta = {}, format = 'json', saveAs = true, embedAvatars = false, downloadImages = false } = options;
  const { downloads, onStatus } = deps;

  // PDF is fundamentally different from the text formats — it produces
  // binary bytes via pdf-lib, not a string, and its builder is async
  // because it fetches the embedded font. Route it entirely outside of
  // buildExportInternal.
  if (format === 'pdf') {
    onStatus?.({ phase: 'build', messages: messages.length });
    await Promise.resolve();
    const base = computeBaseName(meta);
    const filename = `${base}.pdf`;
    const bytes = await buildPdf(messages, meta, (done, total) => {
      onStatus?.({ phase: 'build', filename, messages: messages.length, messagesBuilt: done, messagesTotal: total });
    }, toPdfOptions(options));
    onStatus?.({ phase: 'build', filename, messages: messages.length, messagesBuilt: messages.length, messagesTotal: messages.length });
    const url = binaryToDownloadUrl(bytes, 'application/pdf');
    try {
      const id = await downloads.download({ url, filename, saveAs });
      setTimeout(() => revokeDownloadUrl(url), 60_000);
      return { ok: true, filename, id };
    } catch (e: any) {
      revokeDownloadUrl(url);
      throw new Error(e?.message || String(e));
    }
  }

  // Auto-upgrade to zip for large HTML exports to avoid "Invalid string length".
  // Embedding many base64 images/avatars inline creates strings too large for V8.
  if (format === 'html') {
    let totalDataBytes = 0;
    for (const m of messages) {
      if (m.attachments) for (const a of m.attachments) { if (a.dataUrl) totalDataBytes += a.dataUrl.length; }
    }
    if (totalDataBytes > 5_000_000 || (downloadImages && messages.length > 500)) {
      return buildAndDownloadZip(deps, { messages, meta, embedAvatars, downloadImages });
    }
  }

  const mode: ImageMode = format === 'html' && downloadImages ? 'data-url' : 'none';
  // Surface the build phase to the popup BEFORE the synchronous serialization
  // runs, then yield once so the popup's render tick fires and the segment-4
  // stripe becomes visible. For JSON/CSV/TXT this flashes briefly; for HTML
  // with large embedded data, the stripe stays visible for the full build.
  onStatus?.({ phase: 'build', messages: messages.length });
  await Promise.resolve();
  const built = buildExportInternal({ messages, meta, format, embedAvatars, downloadImages }, mode);
  onStatus?.({ phase: 'build', filename: built.filename, messages: messages.length, messagesBuilt: messages.length, messagesTotal: messages.length });

  const url = textToDownloadUrl(built.content, built.mime);
  try {
    const id = await downloads.download({ url, filename: built.filename, saveAs });
    setTimeout(() => revokeDownloadUrl(url), 60000);
    return { ok: true, filename: built.filename, id };
  } catch (e: any) {
    const safe = `${sanitizeBase('teams-chat')}-${Date.now()}.${format === 'html' ? 'html' : format === 'csv' ? 'csv' : format === 'txt' ? 'txt' : 'json'}`;
    revokeDownloadUrl(url);
    try {
      const url2 = textToDownloadUrl(built.content, built.mime);
      const id2 = await downloads.download({ url: url2, filename: safe, saveAs });
      setTimeout(() => URL.revokeObjectURL(url2), 60000);
      return { ok: true, filename: safe, id: id2 };
    } catch (e2: any) {
      throw new Error(e2?.message || String(e2));
    }
  }
}

export async function buildAndDownloadZip(
  deps: BuildDownloadDeps,
  { messages = [], meta = {}, embedAvatars = false, downloadImages = false, avatarMode = 'inline' }: BuildExportOptions,
) {
  const { downloads, onStatus } = deps;
  // Show segment 4 active before the synchronous build/zip work begins so the
  // popup has a render tick to paint the indeterminate stripe.
  onStatus?.({ phase: 'build', messages: messages.length });
  await Promise.resolve();

  // Extract avatars into files when embedAvatars is on AND the user
  // opted into files mode. If either is off, `collectAvatarFiles`
  // effectively no-ops (empty list + pass-through meta) and avatars
  // stay inline in the HTML's style block.
  const avatarCollect = embedAvatars && avatarMode === 'files' ? collectAvatarFiles(meta) : { files: [], meta };

  const built = buildExportInternal(
    { messages, meta: avatarCollect.meta, format: 'html', embedAvatars, downloadImages },
    downloadImages ? 'files' : 'none',
  );

  onStatus?.({ phase: 'build', filename: `${built.baseFolder}.zip`, messages: messages.length, messagesBuilt: messages.length, messagesTotal: messages.length });

  const encoder = new TextEncoder();
  const files: { path: string; data: Uint8Array }[] = [];

  // Place avatar files first so they group together in the zip listing.
  for (const a of avatarCollect.files) {
    const bytes = dataUrlToBytes(a.dataUrl);
    if (!bytes) continue;
    files.push({ path: `${built.baseFolder}/${a.filename}`, data: bytes });
  }
  // Handle chunked HTML content (string[]) by encoding each chunk and concatenating
  const contentParts = Array.isArray(built.content) ? built.content : [built.content];
  const encodedParts = contentParts.map(part => encoder.encode(part));
  const totalLen = encodedParts.reduce((sum, p) => sum + p.length, 0);
  const contentBytes = new Uint8Array(totalLen);
  let offset = 0;
  for (const part of encodedParts) {
    contentBytes.set(part, offset);
    offset += part.length;
  }
  files.push({
    path: `${built.baseFolder}/index.html`,
    data: contentBytes,
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
  const url = binaryToDownloadUrl(zipBytes, mime);
  try {
    const id = await downloads.download({ url, filename: zipName, saveAs: true });
    setTimeout(() => revokeDownloadUrl(url), 60_000);
    return { ok: true, filename: zipName, id };
  } catch (e: any) {
    revokeDownloadUrl(url);
    throw new Error(e?.message || String(e));
  }
}

// Build every requested format and pack them into a single bundle.zip.
// Used when the user has 2+ formats selected. The HTML build (if present
// and downloadImages=true) contributes its inline images to a shared
// images/ folder; other formats are dropped in alongside as standalone
// files. The bundle uses the same `TeamsExport_<chat>_<stamp>/` folder
// inside the zip as buildAndDownloadZip, so extraction behaviour is
// uniform regardless of which path produced the .zip.
export async function buildAndDownloadBundle(
  deps: BuildDownloadDeps,
  options: BuildBundleOptions,
) {
  const { messages = [], meta = {}, formats, embedAvatars = false, downloadImages = false, avatarMode = 'inline' } = options;
  const { downloads, onStatus } = deps;
  if (!formats.length) throw new Error('buildAndDownloadBundle: formats is empty');

  onStatus?.({ phase: 'build', messages: messages.length });
  await Promise.resolve();

  // Pre-compute the bundle's canonical folder name so PDF (async) and the
  // text builders (sync) share the same stamp. buildExportInternal also
  // calls computeBaseName, but its result for the FIRST text format is
  // only used as a tiebreaker — for PDF we must have the name up front
  // because the PDF builder never touches buildExportInternal.
  let baseFolder = computeBaseName(meta);

  // HTML-in-bundle gets the avatars-as-files treatment when avatars
  // are enabled, HTML is selected, AND the user chose 'files' mode.
  // Otherwise it's a no-op that returns the input meta.
  const needAvatarsAsFiles = embedAvatars && formats.includes('html') && avatarMode === 'files';
  const avatarCollect = needAvatarsAsFiles ? collectAvatarFiles(meta) : { files: [], meta };
  const htmlMeta = avatarCollect.meta;

  // Each text format runs through buildExportInternal independently. HTML
  // uses 'files' image mode when downloadImages is on (so its <img src>
  // point at images/foo.png and the data URLs come back as inlineImages
  // we can place into the zip); other formats use 'none' to strip data
  // URLs from the JSON payload (otherwise base64 blobs would inflate it).
  // PDF takes a separate async path below.
  const htmlImageMode: ImageMode = downloadImages ? 'files' : 'none';
  const files: { path: string; data: Uint8Array }[] = [];
  const imagePaths = new Set<string>();
  const encoder = new TextEncoder();

  // Avatar files — one per unique avatarId, PNG/JPEG bytes. Written to
  // <baseFolder>/avatars/<id>.<ext>. HTML references them via CSS urls.
  for (const a of avatarCollect.files) {
    const bytes = dataUrlToBytes(a.dataUrl);
    if (!bytes) continue;
    files.push({ path: `${baseFolder}/${a.filename}`, data: bytes });
  }

  for (const format of formats) {
    if (format === 'pdf') {
      // PDF: async bytes straight from pdf-lib. Named <baseFolder>.pdf so
      // it matches the other file naming inside the bundle. PDF gets
      // the original meta (with inline avatars) because pdf-lib embeds
      // its own copies — the avatars/ folder is HTML-only. Progress
      // callbacks flow through to the popup so the "building" segment
      // actually ticks during the (potentially long) PDF phase.
      const bytes = await buildPdf(
        messages,
        meta,
        (done, total) => {
          onStatus?.({ phase: 'build', filename: `${baseFolder}.zip`, messages: messages.length, messagesBuilt: done, messagesTotal: total });
        },
        toPdfOptions(options),
      );
      files.push({ path: `${baseFolder}/${baseFolder}.pdf`, data: bytes });
      continue;
    }

    const mode: ImageMode = format === 'html' ? htmlImageMode : 'none';
    // HTML uses the meta that routes avatars to files; other formats
    // get the original inline-avatars meta (they don't read the flag,
    // but the JSON path emits the avatars map either way).
    const formatMeta = format === 'html' ? htmlMeta : meta;
    const built = buildExportInternal({ messages, meta: formatMeta, format, embedAvatars, downloadImages }, mode);

    // Always rename to <baseFolder>.<ext> rather than reusing built.filename:
    //   - HTML in 'files' mode produces 'index.html' (only meaningful when
    //     HTML is the sole occupant of the zip)
    //   - Other formats use their own (possibly drift-affected) timestamp
    // Forcing the extension keeps every file name in the bundle consistent.
    const perFormatName = `${baseFolder}.${format}`;
    const contentParts = Array.isArray(built.content) ? built.content : [built.content];
    const encodedParts = contentParts.map(part => encoder.encode(part));
    const totalLen = encodedParts.reduce((sum, p) => sum + p.length, 0);
    const contentBytes = new Uint8Array(totalLen);
    let offset = 0;
    for (const part of encodedParts) {
      contentBytes.set(part, offset);
      offset += part.length;
    }
    files.push({ path: `${baseFolder}/${perFormatName}`, data: contentBytes });

    // Inline images only come from the HTML+files build. Dedupe in case
    // some hypothetical future format also produces them.
    for (const img of built.inlineImages) {
      const path = `${baseFolder}/${img.filename}`;
      if (imagePaths.has(path)) continue;
      const bytes = dataUrlToBytes(img.dataUrl);
      if (!bytes) continue;
      imagePaths.add(path);
      files.push({ path, data: bytes });
    }
  }

  onStatus?.({ phase: 'build', filename: `${baseFolder}.zip`, messages: messages.length, messagesBuilt: messages.length, messagesTotal: messages.length });

  const zipBytes = buildZip(files);
  const zipName = `${baseFolder}.zip`;
  const url = binaryToDownloadUrl(zipBytes, 'application/zip');
  try {
    const id = await downloads.download({ url, filename: zipName, saveAs: true });
    setTimeout(() => revokeDownloadUrl(url), 60_000);
    return { ok: true, filename: zipName, id };
  } catch (e: any) {
    revokeDownloadUrl(url);
    throw new Error(e?.message || String(e));
  }
}

function toPlainText(messages: ExportMessage[]) {
  const lines: string[] = [];
  for (const m of messages) {
    const ts = m.timestamp || '';
    const author = m.author || '[unknown]';
    const text = (m.text || '').replace(/\r\n/g, '\n').replace(/\n{2,}/g, '\n\n');
    const urgencyTag = m.importance === 'urgent' ? ' [!URGENT]' : m.importance === 'high' ? ' [!IMPORTANT]' : '';
    const subjectLine = m.subject ? `[Subject: ${m.subject}] ` : '';
    let line = `[${ts}] ${author}${urgencyTag}: ${subjectLine}${text}`;

    // Include forward/reply context
    if (m.forwarded?.originalAuthor) {
      const fwdFrom = `[forwarded from ${m.forwarded.originalAuthor}]`;
      const fwdText = m.forwarded.originalText ? m.forwarded.originalText.replace(/\n/g, ' ').slice(0, 300) : '';
      line = text
        ? `[${ts}] ${author}${urgencyTag}: ${subjectLine}${text}\n  ${fwdFrom}: ${fwdText}`
        : `[${ts}] ${author}${urgencyTag} ${fwdFrom}: ${subjectLine}${fwdText}`;
    }
    if (m.replyTo?.text) {
      const quotedText = m.replyTo.text.replace(/\n/g, ' ').slice(0, 200);
      const attribution = m.replyTo.author ? `${m.replyTo.author}: ` : '';
      line += `\n  > ${attribution}${quotedText}`;
    }

    lines.push(line);
  }
  return lines.join('\n');
}
