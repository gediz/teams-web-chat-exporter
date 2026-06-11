import { blobToDownloadUrl, removeAvatars, revokeDownloadUrl, summarizeAttachments, toCSV, toHTML } from './builders';
// Every download URL is minted via blobToDownloadUrl: it returns a blob:
// URL where one is available, and on an MV3 service worker (no
// createObjectURL) it routes large payloads through the offscreen document
// instead of a base64 data: URL, which would throw "Invalid string length"
// past ~384 MB (issue #27).
import { buildPdfResilient } from './pdf';
import { formatRangeLabel, sanitizeBase } from '../utils/messages';
import { buildZipAsync } from './zip';
import type { Attachment, ExportMessage, ExportMeta, TableData } from '../types/shared';

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
  // -PARTIAL suffix when the scrape detected an incomplete-export
  // condition (NetworkError caught mid-scrape, etc). Visible at the
  // OS file-browser level so users can spot a partial export without
  // opening the file. Pairs with the in-file warning banner the
  // builders inject.
  const partial = meta.partial ? '-PARTIAL' : '';
  return `TeamsExport_${baseTitle}_${stamp}${partial}`;
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
    content = toCSV(finalMessages, enrichedMeta);
  } else if (format === 'txt') {
    filename = `${base}.txt`;
    mime = 'text/plain';
    finalMessages = stripMessageDataUrls(processedMessages);
    content = toPlainText(finalMessages, enrichedMeta);
  } else {
    filename = `${base}.json`;
    mime = 'application/json';
    finalMessages = stripMessageDataUrls(processedMessages);
    // Surface parsed tables as a clean, additive `tables: TableData[]` field and
    // drop the internal `bodyBlocks` render aid. `text` and contentHtml are left
    // untouched. API messages derive `tables` from bodyBlocks; legacy DOM-scrape
    // messages have their flat `tables` normalized to the same shape so the JSON
    // schema is uniform. Non-table messages pass through unchanged.
    const jsonMessages: unknown[] = finalMessages.map(m => {
      const hasBlocks = Array.isArray(m.bodyBlocks) && m.bodyBlocks.length > 0;
      const hasLegacy = Array.isArray(m.tables) && m.tables.length > 0;
      if (!hasBlocks && !hasLegacy) return m;
      const { bodyBlocks, ...rest } = m;
      const tables: TableData[] = hasBlocks
        ? m.bodyBlocks!.flatMap(b => (b.type === 'table' ? [b.table] : []))
        : (m.tables as string[][][]).map(legacyToTableData);
      return tables.length ? { ...rest, tables } : rest;
    });
    const payload = { meta: { ...enrichedMeta, count: messages.length }, messages: jsonMessages };
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
    const bytes = await buildPdfResilient(messages, meta, (done, total) => {
      onStatus?.({ phase: 'build', filename, messages: messages.length, messagesBuilt: done, messagesTotal: total });
    }, toPdfOptions(options));
    onStatus?.({ phase: 'build', filename, messages: messages.length, messagesBuilt: messages.length, messagesTotal: messages.length });
    const url = await blobToDownloadUrl(new Blob([bytes as BlobPart], { type: 'application/pdf' }), 'application/pdf');
    try {
      console.log(`[pdf-single] downloads.download calling: filename=${filename} bytes=${bytes.byteLength}`);
      const id = await downloads.download({ url, filename, saveAs });
      console.log(`[pdf-single] downloads.download resolved: id=${id}`);
      setTimeout(() => revokeDownloadUrl(url), 60_000);
      return { ok: true, filename, id };
    } catch (e: any) {
      console.log(`[pdf-single] downloads.download rejected: ${e?.message || String(e)}`);
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

  const url = await blobToDownloadUrl(
    new Blob(Array.isArray(built.content) ? built.content : [built.content], { type: built.mime }),
    built.mime,
  );
  try {
    console.log(`[text-single] downloads.download calling: filename=${built.filename} format=${format}`);
    const id = await downloads.download({ url, filename: built.filename, saveAs });
    console.log(`[text-single] downloads.download resolved: id=${id}`);
    setTimeout(() => revokeDownloadUrl(url), 60_000);
    return { ok: true, filename: built.filename, id };
  } catch (e: any) {
    console.log(`[text-single] downloads.download rejected: ${e?.message || String(e)} — retrying with sanitized filename`);
    const safe = `${sanitizeBase('teams-chat')}-${Date.now()}.${format === 'html' ? 'html' : format === 'csv' ? 'csv' : format === 'txt' ? 'txt' : 'json'}`;
    revokeDownloadUrl(url);
    // Hoisted so the catch can release the retry URL too. For >50MB exports
    // this blob: URL is minted in the offscreen document, which has no
    // automatic cleanup; if the retry download also fails we must revoke it
    // explicitly or the export blob stays resident until the doc unloads.
    let url2: string | undefined;
    try {
      url2 = await blobToDownloadUrl(
        new Blob(Array.isArray(built.content) ? built.content : [built.content], { type: built.mime }),
        built.mime,
      );
      const id2 = await downloads.download({ url: url2, filename: safe, saveAs });
      console.log(`[text-single] downloads.download retry resolved: id=${id2}`);
      setTimeout(() => revokeDownloadUrl(url2!), 60_000);
      return { ok: true, filename: safe, id: id2 };
    } catch (e2: any) {
      console.log(`[text-single] downloads.download retry rejected: ${e2?.message || String(e2)}`);
      if (url2) revokeDownloadUrl(url2);
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

  const zipBlob = await buildZipAsync(files, 'zip-html-images');
  const zipName = `${built.baseFolder}.zip`;
  const url = await blobToDownloadUrl(zipBlob, 'application/zip');
  try {
    console.log(`[html-zip] downloads.download calling: filename=${zipName} zipBytes=${zipBlob.size}`);
    const id = await downloads.download({ url, filename: zipName, saveAs: true });
    console.log(`[html-zip] downloads.download resolved: id=${id}`);
    setTimeout(() => revokeDownloadUrl(url), 60_000);
    return { ok: true, filename: zipName, id };
  } catch (e: any) {
    console.log(`[html-zip] downloads.download rejected: ${e?.message || String(e)}`);
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
      const bytes = await buildPdfResilient(
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

  const zipBlob = await buildZipAsync(files, 'zip-per-chat-bundle');
  const zipName = `${baseFolder}.zip`;
  const tUrl = performance.now();
  const url = await blobToDownloadUrl(zipBlob, 'application/zip');
  console.log(`[per-chat-zip] download-url created in ${Math.round(performance.now() - tUrl)}ms (zipBytes=${zipBlob.size}, scheme=${url.slice(0, 5)})`);
  try {
    const tDl = performance.now();
    console.log(`[per-chat-zip] downloads.download calling: filename=${zipName}`);
    const id = await downloads.download({ url, filename: zipName, saveAs: true });
    console.log(`[per-chat-zip] downloads.download resolved: id=${id} ms=${Math.round(performance.now() - tDl)}`);
    setTimeout(() => revokeDownloadUrl(url), 60_000);
    return { ok: true, filename: zipName, id };
  } catch (e: any) {
    console.log(`[per-chat-zip] downloads.download rejected: ${e?.message || String(e)}`);
    revokeDownloadUrl(url);
    throw new Error(e?.message || String(e));
  }
}

// Build every requested format for a single conversation as in-memory
// bytes, ready to be packed into an outer multi-chat bundle.zip. No
// download is triggered; the caller chooses the chat's folder name and
// emits the final outer zip (see buildAndDownloadBundlesZip).
//
// File layout returned (no chat-folder prefix — caller adds it):
//   messages.json
//   messages.html        when 'html' is selected
//   messages.csv / .txt / .pdf
//   avatars/<id>.<ext>   when html + embedAvatars + avatarMode='files'
//   images/<name>.<ext>  when html + downloadImages
//
// Generic filenames (`messages.*`) instead of the single-chat-export
// `<baseFolder>.<ext>` pattern keep paths short inside the bundle and
// match the v1.4 plan's example layout.
type PerChatFile = { relativePath: string; data: Uint8Array };

export type BuildOneChatBundleOptions = {
  messages: ExportMessage[];
  meta: ExportMeta;
  formats: SingleFormat[];
  embedAvatars: boolean;
  downloadImages: boolean;
  avatarMode: AvatarMode;
  // Reported as the per-chat build advances; the caller wraps these to
  // bubble bundle context (chat N of M) up to the popup.
  onPdfProgress?: (done: number, total: number) => void;
} & PdfKnobs;

export type PerChatFormatFailure = { format: SingleFormat; error: string };
export type OneChatBundleResult = { files: PerChatFile[]; formatFailures: PerChatFormatFailure[] };

export async function buildOneChatForBundle(
  options: BuildOneChatBundleOptions,
): Promise<OneChatBundleResult> {
  const { messages, meta, formats, embedAvatars, downloadImages, avatarMode, onPdfProgress } = options;
  if (!formats.length) throw new Error('buildOneChatForBundle: formats is empty');

  const needAvatarsAsFiles = embedAvatars && formats.includes('html') && avatarMode === 'files';
  const avatarCollect = needAvatarsAsFiles ? collectAvatarFiles(meta) : { files: [], meta };
  const htmlMeta = avatarCollect.meta;

  const htmlImageMode: ImageMode = downloadImages ? 'files' : 'none';
  const out: PerChatFile[] = [];
  // Per-format failures are collected, not thrown, so one bad format (e.g. an
  // unbuildable PDF) doesn't drop the whole chat from the bundle.
  const formatFailures: PerChatFormatFailure[] = [];
  const imagePaths = new Set<string>();
  const encoder = new TextEncoder();

  for (const a of avatarCollect.files) {
    const bytes = dataUrlToBytes(a.dataUrl);
    if (!bytes) continue;
    out.push({ relativePath: a.filename, data: bytes });
  }

  // Title used in diagnostic logs so the SW console shows which chat
  // is being built when something dies. The chat is identified by
  // meta.title in normal use; falls back to "?" if upstream didn't
  // stamp one (shouldn't happen in the bundle path but defensive).
  const titleForLog = (meta?.title || '?').slice(0, 60);

  for (const format of formats) {
    // Yield to the event loop before each format. Without this, building
    // 5 formats × 12k+ messages ties the extension process up for several
    // seconds straight — observed effect: Firefox popup mount sees its
    // chrome.storage.local read balloon from ~25ms to ~400ms+ and renders
    // as a white pixel until the build finishes. setTimeout(0) gives any
    // pending popup mount or storage IO a chance to interleave.
    await new Promise<void>(r => setTimeout(r, 0));

    // Per-format diagnostic — the bundle-handler's catch only sees the
    // bare error message ("The operation was aborted." was the recent
    // mystery), so it can't tell which format died. Log when we start
    // each format and re-throw with format context if anything inside
    // explodes. No behavior change: the bundle handler still pushes
    // the failure into FAILURES.txt and continues to the next chat.
    try {
      console.log(`[bundle] building ${format} for "${titleForLog}"`);
      if (format === 'pdf') {
        const bytes = await buildPdfResilient(messages, meta, onPdfProgress, toPdfOptions(options));
        out.push({ relativePath: 'messages.pdf', data: bytes });
        continue;
      }

      const mode: ImageMode = format === 'html' ? htmlImageMode : 'none';
      const formatMeta = format === 'html' ? htmlMeta : meta;
      const built = buildExportInternal({ messages, meta: formatMeta, format, embedAvatars, downloadImages }, mode);

      const contentParts = Array.isArray(built.content) ? built.content : [built.content];
      const encodedParts = contentParts.map(part => encoder.encode(part));
      const totalLen = encodedParts.reduce((sum, p) => sum + p.length, 0);
      const contentBytes = new Uint8Array(totalLen);
      let offset = 0;
      for (const part of encodedParts) {
        contentBytes.set(part, offset);
        offset += part.length;
      }
      const ext = format;
      out.push({ relativePath: `messages.${ext}`, data: contentBytes });

      for (const img of built.inlineImages) {
        if (imagePaths.has(img.filename)) continue;
        const bytes = dataUrlToBytes(img.dataUrl);
        if (!bytes) continue;
        imagePaths.add(img.filename);
        out.push({ relativePath: img.filename, data: bytes });
      }
    } catch (e: unknown) {
      // Per-format resilience: record this format's failure and keep building
      // the rest, so the chat still gets a folder with whatever formats
      // succeed. The bundle handler lists the failed format in FAILURES.txt.
      // (Previously this re-threw and dropped the ENTIRE chat — all formats —
      // when a single format, almost always the PDF, failed.)
      const inner = e instanceof Error ? e.message : String(e);
      console.log(`[bundle] build:${format} failed for "${titleForLog}": ${inner}`);
      formatFailures.push({ format, error: `build:${format} — ${inner}` });
    }
  }

  return { files: out, formatFailures };
}

// Sanitise + dedupe a chat title for use as a folder name inside a
// multi-chat bundle. Collisions get `(2)`, `(3)`, ... appended.
//
// Mutates `used` to reserve the chosen name so subsequent calls in
// the same bundle keep picking unique values.
export function pickBundleFolderName(rawTitle: string, used: Set<string>): string {
  const base = sanitizeBase(rawTitle || 'UnnamedChat') || 'UnnamedChat';
  if (!used.has(base.toLowerCase())) {
    used.add(base.toLowerCase());
    return base;
  }
  let idx = 2;
  while (used.has(`${base} (${idx})`.toLowerCase())) idx += 1;
  const chosen = `${base} (${idx})`;
  used.add(chosen.toLowerCase());
  return chosen;
}

// Per-chat success entry passed to buildAndDownloadBundlesZip. The
// folder name has already been picked (and deduped) by the caller so
// progress messages can show the exact label used in the zip.
export type BundleEntry = {
  folderName: string;
  files: PerChatFile[];
};

export type BundleFailure = {
  // The folder name we WOULD have used so the FAILURES.txt line is
  // human-recognisable even if the chat never produced output.
  folderName: string;
  conversationId: string;
  reason: string;
};

// "Empty" chats — the API succeeded but the server has no message
// history for that conversation (Teams Free legacy Skype-imported 1:1s
// are the common case: Microsoft never migrated those messages to the
// consumer chat backend). Listed at the bundle root in NO_HISTORY.txt
// rather than producing per-chat folders full of empty files — saves
// ~7 MB of empty-PDF noise per chat, and the user can see at a glance
// which chats had nothing to fetch.
export type BundleEmpty = {
  folderName: string;
  conversationId: string;
};

// "Partial" chats — the API succeeded and produced output, but the
// scrape detected an incomplete-data condition (NetworkError mid-scrape,
// DOM-scroll truncation). Listed at the bundle root in PARTIAL.txt so
// the user can see at a glance which chats may be missing content.
// Distinct from BundleFailure (failures = no output) and BundleEmpty
// (empty = API said the chat is empty server-side).
export type BundlePartial = {
  folderName: string;
  conversationId: string;
  reason: 'network' | 'truncation';
};

// Pack every per-chat result into a single outer zip and download it.
// Failures (if any) are written as one line each into FAILURES.txt at
// the zip's root so the user gets a concrete punch-list of what didn't
// export, without aborting the rest of the run.
export async function buildAndDownloadBundlesZip(
  deps: BuildDownloadDeps,
  entries: BundleEntry[],
  failures: BundleFailure[],
  noHistory: BundleEmpty[] = [],
  partials: BundlePartial[] = [],
) {
  const { downloads } = deps;

  const outerFiles: { path: string; data: Uint8Array }[] = [];
  for (const entry of entries) {
    for (const f of entry.files) {
      outerFiles.push({ path: `${entry.folderName}/${f.relativePath}`, data: f.data });
    }
  }

  if (failures.length) {
    const lines = failures.map(f => `${f.folderName}\t${f.conversationId}\t${f.reason}`);
    // A reason of "build:<format> — ..." is a per-format failure: that chat's
    // folder still exists with the formats that did build. Reasons without a
    // "build:" prefix (e.g. scrape failures) mean the whole chat is absent.
    const header = '# Export failures. Columns: folder\tconversationId\treason\n'
      + '# "build:<format> — ..." = only that format is missing; the chat folder has the rest.';
    const body = `${header}\n${lines.join('\n')}\n`;
    outerFiles.push({ path: 'FAILURES.txt', data: new TextEncoder().encode(body) });
  }

  if (noHistory.length) {
    const lines = noHistory.map(e => `${e.folderName}\t${e.conversationId}`);
    const header = '# Chats with no retrievable message history. The API succeeded\n'
      + '# but returned 0 messages — most often legacy Skype-imported 1:1\n'
      + '# chats where Microsoft did not migrate the message history into\n'
      + '# the Teams Free chat backend. Not a failure; nothing to export.\n'
      + '# Columns: folder\tconversationId';
    const body = `${header}\n${lines.join('\n')}\n`;
    outerFiles.push({ path: 'NO_HISTORY.txt', data: new TextEncoder().encode(body) });
  }

  if (partials.length) {
    // Bundle-level partial signal. Lists which chats came back with a
    // partial-data warning and the cause tag. Each chat's individual
    // files inside its folder also carry the in-file banner +
    // -PARTIAL filename suffix from the single-chat code path; this
    // file is the bundle-root summary so users see "the bundle has
    // problems" without opening every folder.
    const lines = partials.map(p => `${p.folderName}\t${p.conversationId}\t${p.reason}`);
    const header = '# Chats whose export may be incomplete. The API ran and produced\n'
      + '# output, but a partial-data condition was detected during scraping\n'
      + '# (e.g. a network interruption). The chat\'s own files carry their\n'
      + '# warning banner and a -PARTIAL filename suffix; this list is the\n'
      + '# bundle-root summary. Columns: folder\tconversationId\treason';
    const body = `${header}\n${lines.join('\n')}\n`;
    outerFiles.push({ path: 'PARTIAL.txt', data: new TextEncoder().encode(body) });
  }

  // Async zip — yields to the event loop between file additions.
  // Critical for multi-chat bundles: a 100-chat ~900 MB sync zip
  // blocked the SW thread for ~34 s and froze any open popup.
  const totalOuterBytes = outerFiles.reduce((a, f) => a + f.data.byteLength, 0);
  console.log(`[bundle-outer] entering buildZipAsync: entries=${entries.length} failures=${failures.length} partials=${partials.length} files=${outerFiles.length} totalBytes=${totalOuterBytes}`);
  const zipBlob = await buildZipAsync(outerFiles, 'zip-outer-bundle');
  console.log(`[bundle-outer] zip built: outBytes=${zipBlob.size}`);
  const stamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19).replace('T', '_');
  // Bundle outer-zip filename gets -PARTIAL when any chat in the
  // bundle is partial. Visible in the OS file browser without
  // opening the zip.
  const partialSuffix = partials.length ? '-PARTIAL' : '';
  const zipName = `TeamsExport_bundle_${stamp}${partialSuffix}.zip`;
  const tUrl = performance.now();
  const url = await blobToDownloadUrl(zipBlob, 'application/zip');
  console.log(`[bundle-outer] download-url created in ${Math.round(performance.now() - tUrl)}ms (scheme=${url.slice(0, 5)})`);
  try {
    const tDl = performance.now();
    console.log(`[bundle-outer] downloads.download calling: filename=${zipName}`);
    const id = await downloads.download({ url, filename: zipName, saveAs: true });
    console.log(`[bundle-outer] downloads.download resolved: id=${id} ms=${Math.round(performance.now() - tDl)}`);
    setTimeout(() => revokeDownloadUrl(url), 60_000);
    return { ok: true, filename: zipName, id };
  } catch (e: any) {
    console.log(`[bundle-outer] downloads.download rejected: ${e?.message || String(e)}`);
    revokeDownloadUrl(url);
    throw new Error(e?.message || String(e));
  }
}

// Normalize a legacy DOM-scrape table (flat string[][], no spans) into the
// same TableData shape the API path produces, so the JSON `tables` field is
// one uniform schema regardless of capture mode.
function legacyToTableData(rows: string[][]): TableData {
  const columns = rows.reduce((m, r) => Math.max(m, r.length), 0);
  return {
    columns,
    headerRowCount: 0,
    rows: rows.map(r => {
      const out = r.slice();
      for (let c = 0; c < columns; c++) if (out[c] === undefined) out[c] = '';
      return out;
    }),
    merges: [],
  };
}

// Render a parsed table as a monospace grid with TRUE merges: a spanned cell's
// value is printed once and the internal border is omitted, so it reads as one
// larger cell. This is the reStructuredText / Pandoc grid-table convention, the
// only plain-text form that expresses merged cells. Cell text is flattened to a
// single line (newlines -> spaces) to keep the box art aligned.
function renderTxtTable(t: TableData): string {
  const W = t.columns;
  if (!W || !t.rows.length) return '';

  const covered = new Set<string>(); // positions a span sits on top of (not the anchor)
  for (const mg of t.merges) {
    for (let dr = 0; dr < mg.rowspan; dr++) {
      for (let dc = 0; dc < mg.colspan; dc++) {
        if (dr || dc) covered.add(`${mg.row + dr},${mg.col + dc}`);
      }
    }
  }
  // Value shown at a position: blank when covered, so a span prints only once.
  const shown = (r: number, c: number) =>
    covered.has(`${r},${c}`) ? '' : (t.rows[r]?.[c] ?? '').replace(/\s*\n\s*/g, ' ').trim();

  const colW = new Array(W).fill(1);
  for (let r = 0; r < t.rows.length; r++) {
    for (let c = 0; c < W; c++) colW[c] = Math.max(colW[c], shown(r, c).length);
  }
  const boxW = (c: number) => colW[c] + 2; // one space of padding each side

  // Is there a horizontal border below row r in column c? No when a rowspan
  // (or row+colspan) merge straddles the r / r+1 boundary in that column.
  const borderBelow = (r: number, c: number) => {
    if (r >= t.rows.length - 1) return true; // bottom edge
    for (const mg of t.merges) {
      const inCols = mg.col <= c && c < mg.col + mg.colspan;
      const straddles = mg.row <= r && mg.row + mg.rowspan - 1 >= r + 1;
      if (inCols && straddles) return false;
    }
    return true;
  };
  // Is there a vertical border to the right of column c in row r? No when a
  // colspan merge straddles the c / c+1 boundary in that row, so a colspan
  // cell reads as one wide cell instead of split boxes.
  const borderRight = (r: number, c: number) => {
    if (c >= W - 1) return true; // outer right edge
    for (const mg of t.merges) {
      const inRows = mg.row <= r && r < mg.row + mg.rowspan;
      const straddles = mg.col <= c && mg.col + mg.colspan - 1 >= c + 1;
      if (inRows && straddles) return false;
    }
    return true;
  };

  // Is a vertical wall present at column c's right edge across the separator
  // below row rb? Absent only inside a cell that spans both rows and columns.
  const wallRight = (rb: number, c: number) =>
    borderRight(rb, c) || (rb + 1 < t.rows.length ? borderRight(rb + 1, c) : true);

  // A horizontal rule. `rb` is the row whose bottom this rule draws (null = a
  // full rule, used for the top and bottom edges).
  const rule = (rb: number | null) => {
    const dashAt = (c: number) => (rb === null ? true : borderBelow(rb, c));
    let s = dashAt(0) ? '+' : '|';
    for (let c = 0; c < W; c++) {
      const dash = dashAt(c);
      s += (dash ? '-' : ' ').repeat(boxW(c));
      const dashRight = c + 1 < W ? dashAt(c + 1) : false;
      // '+' where a horizontal rule meets the boundary; else a continuing wall
      // '|', or a blank inside a cell merged across both rows and columns.
      if (dash || dashRight) s += '+';
      else if (rb === null || c >= W - 1 || wallRight(rb, c)) s += '|';
      else s += ' ';
    }
    return s;
  };

  const out: string[] = [rule(null)];
  for (let r = 0; r < t.rows.length; r++) {
    let line = '|';
    for (let c = 0; c < W; c++) {
      line += ` ${shown(r, c).padEnd(colW[c])} ` + (c >= W - 1 || borderRight(r, c) ? '|' : ' ');
    }
    out.push(line);
    out.push(r < t.rows.length - 1 ? rule(r) : rule(null));
  }
  return out.join('\n');
}

function toPlainText(messages: ExportMessage[], meta: ExportMeta = {}) {
  const lines: string[] = [];
  // Partial-export warning: prepend a clearly visible banner so a
  // user opening the .txt file in any editor sees it immediately.
  // Plain-text format means we can't style it; ASCII border + caps
  // is the universally-readable equivalent of an alert box.
  const partial = meta.partial as { reason?: string } | undefined;
  if (partial) {
    const reason = partial.reason === 'network'
      ? 'A network interruption was detected during scraping; some messages may be missing.'
      : 'Some messages may not have fully loaded before the export finished.';
    lines.push(
      '======================================================================',
      `  WARNING: This export may be incomplete. [${partial.reason || 'partial'}]`,
      `  ${reason}`,
      `  Captured ${messages?.length ?? 0} messages.`,
      '======================================================================',
      '',
    );
  }
  // Flatten a quote/forward body to one line, truncated with an ellipsis so
  // it is clear the text was shortened (the TXT format keeps quotes brief).
  const clip = (s: string, max: number) => {
    const flat = s.replace(/\n/g, ' ');
    // Slice by code point so an emoji / surrogate pair at the cut boundary
    // isn't split into a broken half before the ellipsis.
    const cps = Array.from(flat);
    return cps.length > max ? cps.slice(0, max).join('') + '…' : flat;
  };
  // Compact single-line reaction summary, e.g. "[👍 Alice, Bob · ❤️ Carol]".
  // Reactor names when present (the self reactor shows as "You"), else "×count".
  const reactionSummary = (m: ExportMessage): string => {
    const rs = Array.isArray(m.reactions) ? m.reactions : [];
    const parts = rs.map(r => {
      const names = (r.reactors || []).map(x => (x.self ? 'You' : x.name)).filter(Boolean);
      const who = names.length ? ` ${names.join(', ')}` : (r.count ? ` ×${r.count}` : '');
      return `${r.emoji}${who}`;
    });
    return parts.length ? `  [${parts.join(' · ')}]` : '';
  };
  for (const m of messages) {
    const ts = m.timestamp || '';
    const author = m.author || '[unknown]';
    let text = (m.text || '').replace(/\r\n/g, '\n').replace(/\n{2,}/g, '\n\n');
    // API mode: rebuild the body from ordered text/table blocks so pasted
    // tables render as aligned grids in their original positions instead of
    // the flat "cell | cell" form left in `text`.
    if (m.bodyBlocks?.length) {
      text = m.bodyBlocks
        .map(b => (b.type === 'table' ? `\n${renderTxtTable(b.table)}\n` : b.text))
        .join('\n')
        .replace(/\n{3,}/g, '\n\n')
        .trim();
    }
    // When the body is empty but the message carried attachments
    // (image paste, file drop, link preview), surface a short
    // attachment summary so the row isn't silently blank.
    if (!text.trim()) {
      const summary = summarizeAttachments(m);
      if (summary) text = summary;
    }
    // Indent continuation lines of a multi-line body so they aren't mistaken
    // for new "[timestamp] author:" message lines (e.g. pasted logs). Single
    // -line bodies are unchanged.
    const body = text.includes('\n') ? text.replace(/\n/g, '\n    ') : text;
    const urgencyTag = m.importance === 'urgent' ? ' [!URGENT]' : m.importance === 'high' ? ' [!IMPORTANT]' : '';
    const subjectLine = m.subject ? `[Subject: ${m.subject}] ` : '';
    let line = `[${ts}] ${author}${urgencyTag}: ${subjectLine}${body}`;

    // Include forward/reply context
    if (m.forwarded?.originalAuthor) {
      const fwdFrom = `[forwarded from ${m.forwarded.originalAuthor}]`;
      const fwdText = m.forwarded.originalText ? clip(m.forwarded.originalText, 300) : '';
      line = text
        ? `[${ts}] ${author}${urgencyTag}: ${subjectLine}${body}\n  ${fwdFrom}: ${fwdText}`
        : `[${ts}] ${author}${urgencyTag} ${fwdFrom}: ${subjectLine}${fwdText}`;
    }
    // Edited marker + reactions on the same line as the message (HTML/CSV/JSON
    // already carry both; this keeps the leanest TXT format consistent).
    if (m.edited) line += ' (edited)';
    const reactions = reactionSummary(m);
    if (reactions) line += reactions;
    if (m.replyTo?.text) {
      const quotedText = clip(m.replyTo.text, 200);
      const attribution = m.replyTo.author ? `${m.replyTo.author}: ` : '';
      line += `\n  > ${attribution}${quotedText}`;
    }

    lines.push(line);
  }
  return lines.join('\n');
}
