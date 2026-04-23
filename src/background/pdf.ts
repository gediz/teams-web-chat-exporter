// Chat -> PDF builder.
//
// Uses pdf-lib + @pdf-lib/fontkit to embed a real Unicode font (Roboto
// Regular + Bold, bundled under public/fonts/). pdf-lib's built-in
// fonts are WinAnsi-only and would mangle Turkish / Cyrillic / Greek
// characters, which is unacceptable for the primary user's data.
//
// Note on bundle size: these deps add ~1.2 MB to background.js. Tried
// dynamic import() to split them off, but MV3 service workers can't
// load ES chunks at runtime, so Vite inlines the code regardless. Accept
// the SW cold-start parse cost — it's a one-time ~30ms on modern
// hardware, and the alternative (ship without PDF) is worse.
//
// Design choices
// - Single-column layout, A4 portrait, 48pt margins.
// - Each message block: bold author + small timestamp header, then the
//   text body word-wrapped to the content width.
// - Reply context rendered as a short quote above the body.
// - Reactions rendered as a compact "👍 3 · ❤️ 1" line under the body.
//   Emoji render in color via Twemoji SVGs rasterized through
//   OffscreenCanvas at build time (see prewarmTwemoji).
// - Attachments: one line per attachment, "[attachment] label". No
//   embedded raster images in v1 (keeps file size predictable).
// - Page breaks: measured per block; if the current block won't fit we
//   add a new page before rendering it.

import { PDFDocument, rgb, type PDFFont, type PDFImage, type PDFPage } from 'pdf-lib';
import fontkit from '@pdf-lib/fontkit';
import type { Attachment, ExportMessage, ExportMeta } from '../types/shared';

// Page dimensions (points, 1/72 inch). A4 is the default; US Letter is
// selectable via PdfOptions.pageSize.
const PAGE_DIMS = {
  a4: { w: 595.28, h: 841.89 },
  letter: { w: 612, h: 792 },
} as const;
const MARGIN = 48;

// Image rendering: attachments that carry a dataUrl are embedded and
// drawn scaled to the text column width. We cap the displayed height
// at IMAGE_MAX_H to avoid a single image taking a full page; pdf-lib
// keeps the underlying image data either way (no re-compression).
const IMAGE_MAX_H = 320;

// Size ratios relative to the body font size. Picking 10pt as the
// reference makes the math read naturally: at bodySize=10 every
// computed value matches the earlier hard-coded defaults.
const RATIO_TITLE = 1.6;     // 16 at body 10
const RATIO_HEADER = 0.9;    // 9 at body 10  — timestamp
const RATIO_AUTHOR = 1.0;    // 10 at body 10 — bold author
const RATIO_META = 0.8;      // 8 at body 10  — reactions, reply quote
const RATIO_LEAD_BODY = 1.3; // 13 at body 10 — body line height
const RATIO_LEAD_META = 1.1; // 11 at body 10
const RATIO_BLOCK_GAP = 1.0; // 10 at body 10
// Avatar square is ~1.6× body size (16 at default). Scales with font
// so a larger font doesn't leave a tiny avatar floating next to it.
const RATIO_AVATAR = 1.6;
const AVATAR_PAD = 6;        // constant — not size-dependent

// Noto family — chosen for international coverage. Noto Sans covers
// Latin + Cyrillic + Greek + Hebrew + Arabic basic; Noto Sans SC adds
// Chinese / Japanese / Korean ideographs (10 MB, the heaviest font in
// the bundle). Emoji are handled separately via Twemoji SVG (see below)
// so we no longer bundle Noto Emoji.
const FONT_REGULAR_PATH = 'fonts/NotoSans-Regular.ttf';
const FONT_BOLD_PATH = 'fonts/NotoSans-Bold.ttf';
const FONT_CJK_PATH = 'fonts/NotoSansSC-Regular.ttf';

// Twemoji vendored SVG set (public/twemoji/<key>.svg). The manifest
// enumerates the ~4000 available emoji keys so we can tell "this is
// an emoji I can render" from "this is a regular character" without
// doing a 404 per unknown codepoint. Loaded once per export via
// loadTwemojiManifest.
const TWEMOJI_MANIFEST_PATH = 'twemoji/manifest.json';

// Emoji rendered inline with text as a small colored image. We
// rasterize each unique emoji sequence once via OffscreenCanvas, embed
// the resulting PNG via pdf-lib, and draw it at font-size dimensions.
// EMOJI_RASTER_PX is the canvas render size — ~2x the PDF font size so
// the image stays crisp at 100% zoom on retina.
const EMOJI_RASTER_PX = 72;

type Fonts = {
  regular: PDFFont;
  bold: PDFFont;
  cjk: PDFFont;
};

// Layout derived from user-picked PDF options (page size, body font
// size). All pixel dimensions are computed once at the start of a
// build. Scaling everything off bodyFontSize keeps proportions intact
// when the user goes from 8pt to 14pt.
type Layout = {
  pageW: number;
  pageH: number;
  contentWidth: number;
  textColX: number;
  textWidth: number;
  avatarSize: number;
  avatarGutter: number;
  sizeTitle: number;
  sizeHeader: number;
  sizeAuthor: number;
  sizeBody: number;
  sizeMeta: number;
  leadBody: number;
  leadMeta: number;
  blockGap: number;
  // Toggles forwarded from PdfOptions.
  showPageNumbers: boolean;
  includeAvatars: boolean;
};

// Bundle of everything text measurement + drawing needs: the three
// fonts, the emoji manifest (what keys exist on disk), the pre-warmed
// cache of embedded emoji PDFImage refs, and the computed layout.
type TextCtx = {
  fonts: Fonts;
  emojiManifest: Set<string>;
  twemoji: Map<string, PDFImage | null>;
  layout: Layout;
};

// Embedded-asset caches, scoped to one buildPdf call. Key = source data
// URL for images, avatarId for avatars. Dedupe-by-key cuts both embed
// time and file size when the same asset appears on many messages.
type AssetCache = {
  avatars: Map<string, PDFImage>;    // avatarId -> embedded image
  images: Map<string, PDFImage>;     // dataUrl -> embedded image
  // Twemoji cache: key is the emoji filename stem (e.g. "1f600"), value
  // is the embedded image or `null` when the key is known but could not
  // be rasterized/embedded (skip without retrying).
  twemoji: Map<string, PDFImage | null>;
  // Raw bytes for avatars we already failed to embed — retry is
  // pointless, so we remember failures and skip silently.
  failedAvatars: Set<string>;
  failedImages: Set<string>;
};

// Public options fed to buildPdf. Only pdf-specific knobs live here;
// global embedAvatars is handled separately in the layout (see the
// `includeAvatars` flag).
export type PdfOptions = {
  pageSize?: 'a4' | 'letter';
  bodyFontSize?: number;     // clamped to [8, 16]
  showPageNumbers?: boolean;
  includeAvatars?: boolean;
};

function buildLayout(opts: PdfOptions): Layout {
  const pageSize = opts.pageSize === 'letter' ? 'letter' : 'a4';
  const dims = PAGE_DIMS[pageSize];
  const fs = Math.max(8, Math.min(16, Math.round(opts.bodyFontSize ?? 10)));
  const includeAvatars = opts.includeAvatars !== false;
  const avatarSize = fs * RATIO_AVATAR;
  // Collapse the gutter when avatars are disabled: no point reserving
  // ~22pt of whitespace the user will never use. Text reclaims the
  // full content width, which both looks cleaner and fits more per line.
  const avatarGutter = includeAvatars ? avatarSize + AVATAR_PAD : 0;
  const contentWidth = dims.w - MARGIN * 2;
  return {
    pageW: dims.w,
    pageH: dims.h,
    contentWidth,
    textColX: MARGIN + avatarGutter,
    textWidth: contentWidth - avatarGutter,
    avatarSize,
    avatarGutter,
    sizeTitle: fs * RATIO_TITLE,
    sizeHeader: fs * RATIO_HEADER,
    sizeAuthor: fs * RATIO_AUTHOR,
    sizeBody: fs,
    sizeMeta: fs * RATIO_META,
    leadBody: fs * RATIO_LEAD_BODY,
    leadMeta: fs * RATIO_LEAD_META,
    blockGap: fs * RATIO_BLOCK_GAP,
    showPageNumbers: opts.showPageNumbers !== false,
    includeAvatars,
  };
}

// Running cursor state across pages. Holds the current page + y offset
// plus the Layout so page-break/ensureSpace can add correctly-sized
// pages without threading ctx through every primitive.
type Cursor = {
  doc: PDFDocument;
  page: PDFPage;
  pageIndex: number;
  y: number;         // current baseline y, measured from bottom per pdf-lib convention
  layout: Layout;
};

// ----- helpers that don't need pdf-lib --------------------------------

function formatTimestamp(ts?: string): string {
  if (!ts) return '';
  const d = new Date(ts);
  if (Number.isNaN(d.getTime())) return ts;
  const pad = (n: number) => n.toString().padStart(2, '0');
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

// ----- public entry point ---------------------------------------------

// Single set of rgb() colors, declared once. Safe at module scope now
// that pdf-lib is a static import.
const COLOR_BODY = rgb(0.12, 0.12, 0.15);
const COLOR_META = rgb(0.45, 0.45, 0.50);
const COLOR_AUTHOR = rgb(0.08, 0.14, 0.32);
const COLOR_RULE = rgb(0.86, 0.86, 0.90);
// Own-message accent. Matches the HTML `.own-msg` left rail color
// (#2563eb) so the visual cue is consistent across formats.
const COLOR_OWN_ACCENT = rgb(0.145, 0.388, 0.922);

export async function buildPdf(
  messages: ExportMessage[],
  meta: ExportMeta,
  onProgress?: (done: number, total: number) => void,
  pdfOptions: PdfOptions = {},
): Promise<Uint8Array> {
  // Layout first — page size + font size drive every subsequent
  // dimension. Stays constant throughout the build (we don't mix
  // settings mid-document).
  const layout = buildLayout(pdfOptions);

  const doc = await PDFDocument.create();
  doc.registerFontkit(fontkit as unknown as Parameters<typeof doc.registerFontkit>[0]);

  const fonts = await loadFonts(doc);
  doc.setTitle(meta.title || 'Teams Chat Export');
  doc.setCreator('Teams Chat Exporter');
  doc.setProducer('Teams Chat Exporter');
  doc.setCreationDate(new Date());

  const firstPage = doc.addPage([layout.pageW, layout.pageH]);
  const cursor: Cursor = {
    doc,
    page: firstPage,
    pageIndex: 0,
    y: layout.pageH - MARGIN,
    layout,
  };

  const cache: AssetCache = {
    avatars: new Map(),
    images: new Map(),
    twemoji: new Map(),
    failedAvatars: new Set(),
    failedImages: new Set(),
  };
  const avatars = meta.avatars || {};

  // Prewarm emoji images so rendering can stay synchronous. This is a
  // single pass over all message text that fetches + rasterizes each
  // unique emoji sequence in parallel.
  const emojiManifest = await loadTwemojiManifest();
  await prewarmTwemoji(doc, cache, messages);

  const ctx: TextCtx = { fonts, emojiManifest, twemoji: cache.twemoji, layout };

  const metaWithCount: ExportMeta = { ...meta, count: messages.length } as ExportMeta;
  renderHeader(cursor, metaWithCount, ctx);

  const total = messages.length;
  const PROGRESS_INTERVAL = 50;
  for (let i = 0; i < total; i++) {
    await renderMessage(cursor, messages[i], ctx, { doc, cache, avatars });
    if (onProgress && (i % PROGRESS_INTERVAL === 0 || i === total - 1)) {
      onProgress(i + 1, total);
    }
  }

  // Page numbers — second pass. Must run after all content is laid out
  // so we know the final page count. pdf-lib lets us revisit any page
  // via getPages(); the footer is a single centered line at the
  // bottom margin using the regular font.
  if (layout.showPageNumbers) {
    const pages = doc.getPages();
    const totalPages = pages.length;
    const fsize = Math.max(8, layout.sizeMeta);
    for (let p = 0; p < totalPages; p++) {
      const page = pages[p];
      const label = `${p + 1} / ${totalPages}`;
      // Measure via the regular font so we can center.
      let textW = 0;
      try { textW = fonts.regular.widthOfTextAtSize(label, fsize); } catch { textW = label.length * fsize * 0.5; }
      const x = (layout.pageW - textW) / 2;
      const y = MARGIN / 2;
      try {
        page.drawText(label, { x, y, size: fsize, font: fonts.regular, color: COLOR_META });
      } catch { /* ignore — extremely unlikely */ }
    }
  }

  return doc.save();
}

// ----- font loading ---------------------------------------------------

async function fetchFont(path: string): Promise<ArrayBuffer> {
  const url = chrome.runtime.getURL(path);
  const resp = await fetch(url);
  if (!resp.ok) throw new Error(`Failed to load font at ${path}: HTTP ${resp.status}`);
  return resp.arrayBuffer();
}

async function loadFonts(doc: PDFDocument): Promise<Fonts> {
  const [regularBytes, boldBytes, cjkBytes] = await Promise.all([
    fetchFont(FONT_REGULAR_PATH),
    fetchFont(FONT_BOLD_PATH),
    fetchFont(FONT_CJK_PATH),
  ]);
  // subset: true is critical for CJK — the full NotoSansSC has 40k+
  // glyphs and 10 MB of data; subsetting drops the PDF down to just
  // the codepoints actually used (often a handful per export).
  const regular = await doc.embedFont(regularBytes, { subset: true });
  const bold = await doc.embedFont(boldBytes, { subset: true });
  const cjk = await doc.embedFont(cjkBytes, { subset: true });
  return { regular, bold, cjk };
}

// Per-character font selection for NON-emoji characters. Emoji are
// handled separately via Twemoji SVG (see segmentText). Order:
//   1. Primary (Noto Sans Regular/Bold) — Latin, Cyrillic, Greek, Hebrew,
//      Arabic basic, Western-European extras
//   2. CJK (Noto Sans SC) — Chinese/Japanese/Korean ideographs + Hangul
type PreferredWeight = 'regular' | 'bold';
function pickFontForChar(ch: string, weight: PreferredWeight, fonts: Fonts): PDFFont {
  const cp = ch.codePointAt(0) ?? 0;

  // CJK Unified Ideographs + extensions + Hiragana / Katakana / Hangul.
  if (
    (cp >= 0x3040 && cp <= 0x30FF) ||
    (cp >= 0x3400 && cp <= 0x4DBF) ||
    (cp >= 0x4E00 && cp <= 0x9FFF) ||
    (cp >= 0xAC00 && cp <= 0xD7AF) ||
    (cp >= 0x1100 && cp <= 0x11FF) ||
    (cp >= 0x20000 && cp <= 0x2FFFF)
  ) {
    return fonts.cjk;
  }

  return weight === 'bold' ? fonts.bold : fonts.regular;
}

// ----- Twemoji (color emoji via inline SVG → PNG) --------------------

// Module-scoped cache of available emoji keys. Loaded on first use per
// SW lifetime. Cleared on SW restart; manifest is tiny (~75 KB) so the
// reload cost is negligible.
let _twemojiIndex: Set<string> | null = null;
async function loadTwemojiManifest(): Promise<Set<string>> {
  if (_twemojiIndex) return _twemojiIndex;
  try {
    const url = chrome.runtime.getURL(TWEMOJI_MANIFEST_PATH);
    const resp = await fetch(url);
    if (!resp.ok) { _twemojiIndex = new Set(); return _twemojiIndex; }
    const arr = (await resp.json()) as string[];
    _twemojiIndex = new Set(arr);
    return _twemojiIndex;
  } catch {
    _twemojiIndex = new Set();
    return _twemojiIndex;
  }
}

// Match-the-longest emoji sequence that starts at text[i]. Returns the
// filename key (codepoints joined with "-") and the number of UTF-16
// code units consumed. Returns null when no known emoji starts at i.
// Tries progressively shorter prefixes so e.g. "👨‍💻" matches the full
// ZWJ sequence when available, and falls back to the base emoji "👨"
// when the specific sequence isn't in the Twemoji set.
function readEmojiSequence(text: string, i: number, index: Set<string>): { key: string; len: number } | null {
  const cp0 = text.codePointAt(i);
  if (cp0 === undefined) return null;

  // Collect up to N codepoints worth of a plausible emoji sequence:
  // starter + any joiners (ZWJ, VS16), modifiers (skin tone), or
  // subsequent emoji codepoints that follow a ZWJ.
  const cps: number[] = [cp0];
  let j = i + (cp0 > 0xFFFF ? 2 : 1);
  // Cap at 10 codepoints so we don't walk the entire string on
  // pathological input; real Twemoji sequences top out around 7.
  for (let k = 0; k < 10 && j < text.length; k++) {
    const cp = text.codePointAt(j);
    if (cp === undefined) break;
    const isZwj = cp === 0x200D;
    const isVS = cp === 0xFE0F;
    const isSkin = cp >= 0x1F3FB && cp <= 0x1F3FF;
    if (isZwj || isVS || isSkin) {
      cps.push(cp);
      j += cp > 0xFFFF ? 2 : 1;
      if (isZwj && j < text.length) {
        // The codepoint after a ZWJ is part of the sequence regardless
        // of what it is (woman + technologist, etc).
        const next = text.codePointAt(j);
        if (next !== undefined) {
          cps.push(next);
          j += next > 0xFFFF ? 2 : 1;
        }
      }
      continue;
    }
    break;
  }

  // Match longest→shortest so "family" finds its full ZWJ sequence
  // before degrading to the base emoji. Also try stripping VS16 since
  // Twemoji filenames sometimes omit FE0F.
  for (let n = cps.length; n >= 1; n--) {
    const sub = cps.slice(0, n);
    const key = sub.map(c => c.toString(16)).join('-');
    if (index.has(key)) {
      const len = sub.reduce((s, c) => s + (c > 0xFFFF ? 2 : 1), 0);
      return { key, len };
    }
    const keyNoVs = sub.filter(c => c !== 0xFE0F).map(c => c.toString(16)).join('-');
    if (keyNoVs !== key && index.has(keyNoVs)) {
      const len = sub.reduce((s, c) => s + (c > 0xFFFF ? 2 : 1), 0);
      return { key: keyNoVs, len };
    }
  }
  return null;
}

// Walk every text field on every message collecting unique emoji keys
// so we can pre-embed them in parallel, up front. Rendering then stays
// synchronous.
function collectEmojiKeys(messages: ExportMessage[], index: Set<string>): Set<string> {
  const keys = new Set<string>();
  const visit = (t: string | undefined | null) => {
    if (!t) return;
    let i = 0;
    while (i < t.length) {
      const seq = readEmojiSequence(t, i, index);
      if (seq) {
        keys.add(seq.key);
        i += seq.len;
      } else {
        const cp = t.codePointAt(i);
        i += (cp !== undefined && cp > 0xFFFF) ? 2 : 1;
      }
    }
  };
  for (const m of messages) {
    visit(m.text);
    visit(m.author);
    visit(m.subject);
    visit(m.replyTo?.text);
    visit(m.replyTo?.author);
    visit(m.forwarded?.originalAuthor);
    visit(m.forwarded?.originalText);
    if (Array.isArray(m.reactions)) for (const r of m.reactions) visit(r.emoji);
    if (Array.isArray(m.attachments)) for (const a of m.attachments) visit(a.label || '');
  }
  return keys;
}

// Rasterize a Twemoji SVG to PNG bytes via OffscreenCanvas + createImageBitmap.
// Both APIs are available in MV3 service workers (Chromium 2020+, Firefox 113+).
async function rasterizeTwemoji(key: string): Promise<Uint8Array | null> {
  try {
    const url = chrome.runtime.getURL(`twemoji/${key}.svg`);
    const resp = await fetch(url);
    if (!resp.ok) return null;
    const svgBlob = await resp.blob();
    const bitmap = await createImageBitmap(svgBlob, {
      resizeWidth: EMOJI_RASTER_PX,
      resizeHeight: EMOJI_RASTER_PX,
      resizeQuality: 'high',
    });
    const canvas = new OffscreenCanvas(EMOJI_RASTER_PX, EMOJI_RASTER_PX);
    const ctx = canvas.getContext('2d');
    if (!ctx) return null;
    ctx.drawImage(bitmap, 0, 0);
    const pngBlob = await canvas.convertToBlob({ type: 'image/png' });
    const buf = await pngBlob.arrayBuffer();
    return new Uint8Array(buf);
  } catch {
    return null;
  }
}

// Populate the emoji image cache for every key actually used in this
// export, in parallel. Failed/missing keys are cached as null so the
// rendering pass skips them without retrying.
async function prewarmTwemoji(
  doc: PDFDocument,
  cache: AssetCache,
  messages: ExportMessage[],
): Promise<void> {
  const index = await loadTwemojiManifest();
  if (!index.size) return;
  const keys = collectEmojiKeys(messages, index);
  await Promise.all(Array.from(keys).map(async key => {
    if (cache.twemoji.has(key)) return;
    const bytes = await rasterizeTwemoji(key);
    if (!bytes) { cache.twemoji.set(key, null); return; }
    try {
      const img = await doc.embedPng(bytes);
      cache.twemoji.set(key, img);
    } catch {
      cache.twemoji.set(key, null);
    }
  }));
}

// ----- text segmentation (text + emoji runs) -------------------------

type Run =
  | { type: 'text'; font: PDFFont; text: string }
  | { type: 'emoji'; image: PDFImage };

// Split a string into mixed runs of font-backed text and inline emoji
// images. The emoji lookup uses only the pre-warmed cache — if a key
// isn't in it, we fall back to rendering the raw codepoints through
// the font stack (which will usually produce a box since Noto Sans
// doesn't cover emoji). That's the acceptable long tail.
function segmentText(
  text: string,
  weight: PreferredWeight,
  fonts: Fonts,
  twemoji: Map<string, PDFImage | null>,
  manifest: Set<string>,
): Run[] {
  const out: Run[] = [];
  const pushText = (ch: string) => {
    const font = pickFontForChar(ch, weight, fonts);
    const last = out[out.length - 1];
    if (last && last.type === 'text' && last.font === font) last.text += ch;
    else out.push({ type: 'text', font, text: ch });
  };

  let i = 0;
  while (i < text.length) {
    const seq = readEmojiSequence(text, i, manifest);
    if (seq) {
      const img = twemoji.get(seq.key);
      if (img) {
        out.push({ type: 'emoji', image: img });
        i += seq.len;
        continue;
      }
      // Known key but rasterization failed, OR not prewarmed. Skip the
      // whole sequence as raw text runs.
      for (let k = 0; k < seq.len;) {
        const ch = text[i + k];
        // Preserve surrogate pairs.
        const cp = text.codePointAt(i + k);
        const chLen = cp !== undefined && cp > 0xFFFF ? 2 : 1;
        pushText(text.slice(i + k, i + k + chLen));
        k += chLen;
      }
      i += seq.len;
      continue;
    }
    const cp = text.codePointAt(i);
    const chLen = cp !== undefined && cp > 0xFFFF ? 2 : 1;
    pushText(text.slice(i, i + chLen));
    i += chLen;
  }
  return out;
}

// ----- asset embedding ------------------------------------------------

const DATA_URL_RE = /^data:(image\/[a-z0-9.+-]+);base64,(.+)$/i;

function dataUrlToBytes(dataUrl: string): { mime: string; bytes: Uint8Array } | null {
  const m = dataUrl.match(DATA_URL_RE);
  if (!m) return null;
  try {
    const bin = atob(m[2]);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return { mime: m[1].toLowerCase(), bytes };
  } catch {
    return null;
  }
}

// Embed an image data URL using the right pdf-lib codec. PDF natively
// supports PNG and JPEG; any other MIME we return null for and the
// caller falls back to a text placeholder. WebP / GIF / SVG can't be
// embedded without re-encoding, which we deliberately don't attempt.
async function embedImage(doc: PDFDocument, dataUrl: string): Promise<PDFImage | null> {
  const parsed = dataUrlToBytes(dataUrl);
  if (!parsed) return null;
  try {
    if (parsed.mime === 'image/png') return await doc.embedPng(parsed.bytes);
    if (parsed.mime === 'image/jpeg' || parsed.mime === 'image/jpg') return await doc.embedJpg(parsed.bytes);
    return null;
  } catch {
    return null;
  }
}

async function getAvatar(doc: PDFDocument, cache: AssetCache, avatarId: string, avatars: Record<string, string>): Promise<PDFImage | null> {
  const cached = cache.avatars.get(avatarId);
  if (cached) return cached;
  if (cache.failedAvatars.has(avatarId)) return null;
  const dataUrl = avatars[avatarId];
  if (!dataUrl) { cache.failedAvatars.add(avatarId); return null; }
  const img = await embedImage(doc, dataUrl);
  if (!img) { cache.failedAvatars.add(avatarId); return null; }
  cache.avatars.set(avatarId, img);
  return img;
}

async function getAttachmentImage(doc: PDFDocument, cache: AssetCache, dataUrl: string): Promise<PDFImage | null> {
  const cached = cache.images.get(dataUrl);
  if (cached) return cached;
  if (cache.failedImages.has(dataUrl)) return null;
  const img = await embedImage(doc, dataUrl);
  if (!img) { cache.failedImages.add(dataUrl); return null; }
  cache.images.set(dataUrl, img);
  return img;
}

// ----- word wrap ------------------------------------------------------

// Measure a run's width. Emoji runs render at the font em-size (1em
// square) so we return `size` as the width. Text runs sum each
// character's measured width in its font.
function runWidth(run: Run, size: number): number {
  if (run.type === 'emoji') return size;
  let w = 0;
  for (const ch of run.text) {
    try { w += run.font.widthOfTextAtSize(ch, size); }
    catch { /* skip unmeasurable */ }
  }
  return w;
}

function wrapText(text: string, weight: PreferredWeight, ctx: TextCtx, size: number, maxWidth: number): string[] {
  if (!text) return [''];
  const paragraphs = text.split(/\r?\n/);
  const out: string[] = [];
  for (const para of paragraphs) {
    if (!para) { out.push(''); continue; }
    const words = para.split(/(\s+)/);
    let line = '';
    for (const w of words) {
      const candidate = line + w;
      if (safeWidthOf(candidate, weight, ctx, size) <= maxWidth) {
        line = candidate;
        continue;
      }
      if (line.trim()) out.push(line.replace(/\s+$/, ''));
      if (safeWidthOf(w, weight, ctx, size) > maxWidth) {
        const broken = hardBreak(w, weight, ctx, size, maxWidth);
        for (let i = 0; i < broken.length - 1; i++) out.push(broken[i]);
        line = broken[broken.length - 1];
      } else {
        line = w.replace(/^\s+/, '');
      }
    }
    if (line) out.push(line.replace(/\s+$/, ''));
  }
  return out.length ? out : [''];
}

function hardBreak(word: string, weight: PreferredWeight, ctx: TextCtx, size: number, maxWidth: number): string[] {
  const lines: string[] = [];
  let current = '';
  for (const ch of word) {
    const w = safeWidthOf(current + ch, weight, ctx, size);
    if (w > maxWidth && current) {
      lines.push(current);
      current = ch;
    } else {
      current += ch;
    }
  }
  if (current) lines.push(current);
  return lines.length ? lines : [word];
}

// Measure a mixed text-and-emoji string. Text characters are measured
// via their matching font; emoji sequences are treated as a square of
// size `size` (1em). Characters no font can render stay 0-width so
// wrap calculations don't explode.
function safeWidthOf(text: string, weight: PreferredWeight, ctx: TextCtx, size: number): number {
  let w = 0;
  for (const run of segmentText(text, weight, ctx.fonts, ctx.twemoji, ctx.emojiManifest)) {
    w += runWidth(run, size);
  }
  return w;
}

// Draw text+emoji at (x,y). Text runs go through drawText; emoji runs
// draw an inline PDFImage scaled to `size`. The image's baseline is
// set so the visual center aligns roughly with the x-height of the
// surrounding text — a small downward nudge from the baseline.
function drawMixed(
  page: PDFPage,
  text: string,
  x: number,
  y: number,
  size: number,
  weight: PreferredWeight,
  ctx: TextCtx,
  color: Color,
) {
  let cx = x;
  for (const run of segmentText(text, weight, ctx.fonts, ctx.twemoji, ctx.emojiManifest)) {
    if (run.type === 'text') {
      try {
        page.drawText(run.text, { x: cx, y, size, font: run.font, color });
      } catch { /* skip */ }
      cx += runWidth(run, size);
    } else {
      // Emoji square — size x size, shifted down ~20% so it visually
      // aligns with lowercase letter bodies instead of baseline.
      const dy = -size * 0.2;
      try {
        page.drawImage(run.image, { x: cx, y: y + dy, width: size, height: size });
      } catch { /* skip */ }
      cx += size;
    }
  }
}

// ----- drawing primitives --------------------------------------------

function ensureSpace(cursor: Cursor, needed: number) {
  if (cursor.y - needed < MARGIN) {
    cursor.page = cursor.doc.addPage([cursor.layout.pageW, cursor.layout.pageH]);
    cursor.pageIndex += 1;
    cursor.y = cursor.layout.pageH - MARGIN;
  }
}

type Color = ReturnType<typeof rgb>;

// Text lines are drawn in the text column (right of the avatar gutter).
// `indent` is an additional offset inside the column (for reply/forward
// quotes).
function drawLines(
  cursor: Cursor,
  lines: string[],
  weight: PreferredWeight,
  ctx: TextCtx,
  size: number,
  color: Color,
  leading: number,
  indent = 0,
) {
  for (const line of lines) {
    ensureSpace(cursor, leading);
    cursor.y -= leading;
    drawMixed(cursor.page, line, ctx.layout.textColX + indent, cursor.y, size, weight, ctx, color);
  }
}

function drawRule(cursor: Cursor) {
  ensureSpace(cursor, 4);
  cursor.y -= 4;
  // Rule spans the text column only — the avatar gutter is visually a
  // separate column and looks cleaner without a line cutting through it.
  cursor.page.drawLine({
    start: { x: cursor.layout.textColX, y: cursor.y },
    end: { x: MARGIN + cursor.layout.contentWidth, y: cursor.y },
    thickness: 0.5,
    color: COLOR_RULE,
  });
}

// Draw an embedded image in the text column, scaled to fit cursor.layout.textWidth
// and IMAGE_MAX_H while preserving aspect ratio. Returns consumed
// vertical space (before the block gap) so the caller can track it.
function drawImage(cursor: Cursor, img: PDFImage): number {
  const maxW = cursor.layout.textWidth;
  const maxH = IMAGE_MAX_H;
  const scale = Math.min(maxW / img.width, maxH / img.height, 1);
  const w = img.width * scale;
  const h = img.height * scale;
  ensureSpace(cursor, h + 4);
  cursor.y -= h + 2;
  cursor.page.drawImage(img, { x: cursor.layout.textColX, y: cursor.y, width: w, height: h });
  return h + 4;
}

// ----- rendering ------------------------------------------------------

function renderHeader(cursor: Cursor, meta: ExportMeta, ctx: TextCtx) {
  const title = (meta.title || 'Teams Chat Export').trim();
  const lines = wrapText(title, 'bold', ctx, ctx.layout.sizeTitle, ctx.layout.contentWidth);
  drawLines(cursor, lines, 'bold', ctx, ctx.layout.sizeTitle, COLOR_AUTHOR, ctx.layout.sizeTitle + 4);
  const subParts: string[] = [];
  if (meta.timeRange) subParts.push(String(meta.timeRange));
  const count = typeof (meta as { count?: number }).count === 'number' ? (meta as { count: number }).count : undefined;
  if (typeof count === 'number') subParts.push(`${count.toLocaleString()} messages`);
  const sub = subParts.join(' · ');
  if (sub) {
    drawLines(cursor, wrapText(sub, 'regular', ctx, ctx.layout.sizeMeta, ctx.layout.contentWidth), 'regular', ctx, ctx.layout.sizeMeta, COLOR_META, ctx.layout.leadMeta);
  }
  cursor.y -= ctx.layout.blockGap;
}

async function renderMessage(
  cursor: Cursor,
  m: ExportMessage,
  ctx: TextCtx,
  ac: { doc: PDFDocument; cache: AssetCache; avatars: Record<string, string> },
) {
  drawRule(cursor);
  cursor.y -= 4;

  // Author + timestamp headline on one baseline.
  const author = (m.author || '[unknown]').trim();
  const ts = formatTimestamp(m.timestamp);
  const urgency = m.importance === 'urgent' ? ' [URGENT]' : m.importance === 'high' ? ' [IMPORTANT]' : '';
  const editedTag = m.edited ? ' (edited)' : '';
  const authorText = `${author}${urgency}${editedTag}`;
  ensureSpace(cursor, ctx.layout.sizeAuthor + 4);
  cursor.y -= ctx.layout.sizeAuthor + 2;
  const authorBaselineY = cursor.y;

  // Avatar goes in the gutter, vertically aligned so its bottom sits a
  // little below the author baseline (reads as "belongs to this row").
  if (m.avatarId && ctx.layout.includeAvatars) {
    const img = await getAvatar(ac.doc, ac.cache, m.avatarId, ac.avatars);
    if (img) {
      // 2pt below baseline looks right optically.
      const avatarY = authorBaselineY - (ctx.layout.avatarSize - ctx.layout.sizeAuthor) / 2 - 2;
      cursor.page.drawImage(img, {
        x: MARGIN,
        y: avatarY,
        width: ctx.layout.avatarSize,
        height: ctx.layout.avatarSize,
      });
    }
  }

  // Own-message accent: a short blue rectangle at the author baseline,
  // positioned just left of the text column. Mirrors the HTML left
  // rail — simple enough to always fit on one line without needing
  // to track message extent across page breaks.
  if (m.isOwn) {
    const railX = Math.max(MARGIN, ctx.layout.textColX - 3);
    const railH = ctx.layout.sizeAuthor + 2;
    cursor.page.drawRectangle({
      x: railX,
      y: authorBaselineY - 2,
      width: 2,
      height: railH,
      color: COLOR_OWN_ACCENT,
    });
  }

  // Author line: mixed-font + emoji so author names with CJK/Latin/
  // emoji all render cleanly. drawMixed handles per-run font + emoji
  // image dispatch. Own-messages get a tinted color to match the HTML
  // .own-msg styling.
  const authorColor = m.isOwn ? COLOR_OWN_ACCENT : COLOR_AUTHOR;
  drawMixed(cursor.page, authorText, ctx.layout.textColX, cursor.y, ctx.layout.sizeAuthor, 'bold', ctx, authorColor);
  if (ts) {
    const authorWidth = safeWidthOf(authorText, 'bold', ctx, ctx.layout.sizeAuthor);
    drawMixed(cursor.page, ts, ctx.layout.textColX + authorWidth + 8, cursor.y, ctx.layout.sizeHeader, 'regular', ctx, COLOR_META);
  }

  if (m.replyTo?.text) {
    cursor.y -= 4;
    const who = m.replyTo.author ? `${m.replyTo.author}: ` : '';
    const truncated = m.replyTo.text.length > 200 ? m.replyTo.text.slice(0, 200) + '…' : m.replyTo.text;
    const quote = `> ${who}${truncated}`;
    drawLines(cursor, wrapText(quote, 'regular', ctx, ctx.layout.sizeMeta, ctx.layout.textWidth - 12), 'regular', ctx, ctx.layout.sizeMeta, COLOR_META, ctx.layout.leadMeta, 12);
  }

  if (m.forwarded?.originalAuthor) {
    cursor.y -= 4;
    const fwd = `[forwarded from ${m.forwarded.originalAuthor}]`;
    drawLines(cursor, wrapText(fwd, 'regular', ctx, ctx.layout.sizeMeta, ctx.layout.textWidth), 'regular', ctx, ctx.layout.sizeMeta, COLOR_META, ctx.layout.leadMeta);
    if (m.forwarded.originalText) {
      const truncated = m.forwarded.originalText.length > 300 ? m.forwarded.originalText.slice(0, 300) + '…' : m.forwarded.originalText;
      drawLines(cursor, wrapText(truncated, 'regular', ctx, ctx.layout.sizeBody, ctx.layout.textWidth - 12), 'regular', ctx, ctx.layout.sizeBody, COLOR_BODY, ctx.layout.leadBody, 12);
    }
  }

  if (m.subject) {
    cursor.y -= 2;
    drawLines(cursor, wrapText(`Subject: ${m.subject}`, 'bold', ctx, ctx.layout.sizeBody, ctx.layout.textWidth), 'bold', ctx, ctx.layout.sizeBody, COLOR_BODY, ctx.layout.leadBody);
  }

  const text = (m.text || '').replace(/\r\n/g, '\n');
  if (text) {
    cursor.y -= 2;
    drawLines(cursor, wrapText(text, 'regular', ctx, ctx.layout.sizeBody, ctx.layout.textWidth), 'regular', ctx, ctx.layout.sizeBody, COLOR_BODY, ctx.layout.leadBody);
  }

  if (Array.isArray(m.reactions) && m.reactions.length) {
    const parts = m.reactions.map(r => `${r.emoji} ${r.count}`).join('  ');
    drawLines(cursor, wrapText(parts, 'regular', ctx, ctx.layout.sizeMeta, ctx.layout.textWidth), 'regular', ctx, ctx.layout.sizeMeta, COLOR_META, ctx.layout.leadMeta);
  }

  // Attachments. If a dataUrl is present and embeds as PNG/JPEG we draw
  // the image inline. Everything else (non-image types, failed embeds,
  // or images with no dataUrl because downloadImages was off) falls
  // back to a single text line so the user still sees the filename.
  if (Array.isArray(m.attachments) && m.attachments.length) {
    for (const att of m.attachments) {
      const drewImage = await tryDrawAttachmentImage(cursor, att, ac);
      if (drewImage) continue;
      const label = att.label || att.href || '[attachment]';
      drawLines(cursor, wrapText(`[attachment] ${label}`, 'regular', ctx, ctx.layout.sizeMeta, ctx.layout.textWidth), 'regular', ctx, ctx.layout.sizeMeta, COLOR_META, ctx.layout.leadMeta);
    }
  }

  cursor.y -= ctx.layout.blockGap;
}

async function tryDrawAttachmentImage(
  cursor: Cursor,
  att: Attachment,
  ac: { doc: PDFDocument; cache: AssetCache },
): Promise<boolean> {
  const dataUrl = att.dataUrl || (att.href && att.href.startsWith('data:') ? att.href : null);
  if (!dataUrl) return false;
  const img = await getAttachmentImage(ac.doc, ac.cache, dataUrl);
  if (!img) return false;
  drawImage(cursor, img);
  return true;
}
