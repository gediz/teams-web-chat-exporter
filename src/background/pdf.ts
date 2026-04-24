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

import { PDFDocument, PDFName, PDFString, PDFArray, rgb, type PDFFont, type PDFImage, type PDFPage } from 'pdf-lib';
import fontkit from '@pdf-lib/fontkit';
import type { Attachment, ExportMessage, ExportMeta } from '../types/shared';
import { subsetFont } from './font-subset';

// Match-anywhere regex for URL detection inside rendered text. Kept
// deliberately conservative — URLs end at whitespace or any character
// that's unusual in real URLs (quote, paren, bracket, curly). Trailing
// punctuation (.,;!?) is trimmed off post-match to avoid "visit..."
// swallowing the sentence's trailing period.
const URL_RE = /https?:\/\/[^\s<>"')\]}]+/g;

function trimTrailingPunct(url: string): string {
  return url.replace(/[.,;:!?)\]}]+$/, '');
}

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

  // Scan every rendered text field once: collects the used codepoints
  // (→ font subset targets) and emoji sequences (→ prewarm list). Doing
  // it up front keeps the downstream rendering pass synchronous and
  // guarantees the subset contains every glyph the renderer will draw,
  // because both steps walk the same fields with the same code.
  const emojiManifest = await loadTwemojiManifest();
  const scan = scanDocument(messages, meta, emojiManifest);

  const fonts = await loadFonts(doc, scan.codepoints);
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

  // Fetch + rasterize the Twemoji images we know we'll need in
  // parallel. Rendering then stays synchronous.
  await prewarmTwemojiKeys(doc, cache, scan.emojiKeys);

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

// Page-number codepoints that always get rendered even if no message
// mentions them. Keeps "1 / 42" style footers from going tofu on a
// chat that happens not to use any digit, slash, or space character.
const PAGE_NUMBER_CODEPOINTS = '0123456789/ ';

async function loadFonts(doc: PDFDocument, codepoints: Set<number>): Promise<Fonts> {
  const [regularBytes, boldBytes, cjkBytes] = await Promise.all([
    fetchFont(FONT_REGULAR_PATH),
    fetchFont(FONT_BOLD_PATH),
    fetchFont(FONT_CJK_PATH),
  ]);

  // Subset each font to exactly the codepoints the document uses via
  // HarfBuzz. This is both correct (HarfBuzz is the reference subsetter
  // used by browsers) and tiny (a 10 MB CJK font drops to <1 MB for a
  // typical conversation; a 620 KB Latin font drops to a few KB for a
  // short English chat). The result is a valid TTF that pdf-lib embeds
  // verbatim with `subset: false` — bypassing fontkit's buggy subsetter.
  //
  // If subsetting fails for any reason (corrupt input, WASM error, an
  // empty codepoint set), we fall back to embedding the original full
  // font. That costs ~10 MB extra in the PDF but keeps every glyph
  // available — fail-safe, not fail-broken.
  const regular = await embedSubsetOrFull(doc, regularBytes, codepoints);
  const bold = await embedSubsetOrFull(doc, boldBytes, codepoints);
  const cjk = await embedSubsetOrFull(doc, cjkBytes, codepoints);
  return { regular, bold, cjk };
}

async function embedSubsetOrFull(
  doc: PDFDocument,
  fontBytes: ArrayBuffer,
  codepoints: Set<number>,
): Promise<PDFFont> {
  try {
    const subset = await subsetFont(new Uint8Array(fontBytes), codepoints);
    if (subset && subset.byteLength > 0) {
      return await doc.embedFont(subset, { subset: false });
    }
  } catch { /* fall through */ }
  return doc.embedFont(fontBytes, { subset: false });
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

// Walk every text field on every message, collecting (a) unique emoji
// keys so we can pre-embed the emoji images and (b) every non-emoji
// codepoint so we can build a minimal font subset for the document.
// Single pass over all rendered text — the two outputs stay in lockstep
// with the renderer by construction.
type DocScan = { emojiKeys: Set<string>; codepoints: Set<number> };

function scanDocument(
  messages: ExportMessage[],
  meta: ExportMeta,
  index: Set<string>,
): DocScan {
  const emojiKeys = new Set<string>();
  const codepoints = new Set<number>();
  const visit = (t: string | undefined | null) => {
    if (!t) return;
    let i = 0;
    while (i < t.length) {
      const seq = readEmojiSequence(t, i, index);
      if (seq) {
        emojiKeys.add(seq.key);
        // Also add each emoji codepoint to the font subset as a
        // last-resort fallback: if Twemoji rasterization fails at
        // runtime, drawMixed routes the codepoints through the font.
        // Without the subset containing them, the glyphs get stripped
        // entirely (invisible reactions) rather than rendering as
        // visible tofu boxes. Tofu is ugly; invisible is a bug.
        let j = i;
        const end = i + seq.len;
        while (j < end) {
          const cp = t.codePointAt(j);
          if (cp !== undefined) codepoints.add(cp);
          j += cp !== undefined && cp > 0xFFFF ? 2 : 1;
        }
        i += seq.len;
        continue;
      }
      const cp = t.codePointAt(i);
      if (cp !== undefined) codepoints.add(cp);
      i += (cp !== undefined && cp > 0xFFFF) ? 2 : 1;
    }
  };
  // Header fields: title + "N messages" suffix.
  visit(meta.title);
  visit(meta.timeRange as string | undefined);
  // Per-message text fields. Kept in sync with renderMessage's reads.
  for (const m of messages) {
    visit(m.text);
    visit(m.author);
    visit(m.subject);
    visit(m.replyTo?.text);
    visit(m.replyTo?.author);
    visit(m.forwarded?.originalAuthor);
    visit(m.forwarded?.originalText);
    if (Array.isArray(m.reactions)) {
      for (const r of m.reactions) {
        visit(r.emoji);
        if (r.reactors) for (const reactor of r.reactors) visit(reactor.name);
      }
    }
    if (Array.isArray(m.attachments)) {
      for (const a of m.attachments) visit(a.label || '');
    }
  }
  // Literal strings renderMessage stamps on the page that users don't
  // type. Covering them here ensures the subset always has glyphs for
  // the scaffolding even if the conversation is, say, emoji-only.
  visit('[unknown]');
  visit('[URGENT]');
  visit('[IMPORTANT]');
  visit('(edited)');
  visit('[forwarded from ]');
  visit('[attachment] ');
  visit('Subject: ');
  visit('Teams Chat Export');
  visit('messages');
  visit('> ');
  visit('…');
  // Digits + space + slash for the page-number footer.
  for (const ch of PAGE_NUMBER_CODEPOINTS) codepoints.add(ch.codePointAt(0)!);
  return { emojiKeys, codepoints };
}

// Rasterize a Twemoji SVG to PNG bytes via OffscreenCanvas +
// createImageBitmap. Works in MV3 service workers on Chromium 2020+
// and Firefox 113+.
//
// Twemoji SVGs carry only `viewBox="0 0 36 36"` with no width/height
// attributes. Firefox's createImageBitmap silently fails to rasterize
// such "viewport-only" SVGs even when resizeWidth/Height are passed —
// the returned bitmap is a 0×0 phantom that draws as transparent,
// producing tofu boxes in the final PDF. Fix: fetch as text, inject
// explicit width + height attributes, then hand the fixed string to
// createImageBitmap. Chromium works either way; Firefox only works
// with the sized version.
async function rasterizeTwemoji(key: string): Promise<Uint8Array | null> {
  try {
    const url = chrome.runtime.getURL(`twemoji/${key}.svg`);
    const resp = await fetch(url);
    if (!resp.ok) {
      console.warn(`[pdf] Twemoji ${key}: fetch failed HTTP ${resp.status}`);
      return null;
    }
    let svgText = await resp.text();
    // Stamp width/height onto the root <svg> tag. If they're already
    // present (later Twemoji revisions may add them), this replacement
    // is idempotent — the regex matches any first <svg attributes.
    svgText = svgText.replace(/<svg\b([^>]*)>/i, (_m, attrs: string) => {
      const stripped = attrs
        .replace(/\s+width="[^"]*"/i, '')
        .replace(/\s+height="[^"]*"/i, '');
      return `<svg${stripped} width="${EMOJI_RASTER_PX}" height="${EMOJI_RASTER_PX}">`;
    });

    // Firefox MV2 background has a DOM (window.Image, HTMLImageElement),
    // and its SVG rendering through <img> is more reliable than
    // createImageBitmap on svg blobs — which has a long tail of
    // viewport-inference bugs. Prefer the DOM path when available.
    const png = (typeof Image !== 'undefined' && typeof document !== 'undefined')
      ? await rasterizeViaImage(svgText)
      : await rasterizeViaImageBitmap(svgText);
    if (!png) {
      console.warn(`[pdf] Twemoji ${key}: rasterize returned null`);
      return null;
    }
    return png;
  } catch (err) {
    console.warn(`[pdf] Twemoji ${key}: rasterize threw`, err);
    return null;
  }
}

// DOM-backed rasterization path. Only valid in contexts with an HTML
// document (Firefox MV2 background, or future popup contexts). Loads
// the SVG into an Image via data URL, draws into an OffscreenCanvas,
// and exports PNG bytes. Canvas is detached from DOM; only the Image
// needs to live long enough to reach the 'load' event.
async function rasterizeViaImage(svgText: string): Promise<Uint8Array | null> {
  const dataUrl = 'data:image/svg+xml;base64,' + btoa(unescape(encodeURIComponent(svgText)));
  const img = new Image();
  img.width = EMOJI_RASTER_PX;
  img.height = EMOJI_RASTER_PX;
  const loaded = new Promise<boolean>(resolve => {
    img.onload = () => resolve(true);
    img.onerror = () => resolve(false);
  });
  img.src = dataUrl;
  const ok = await loaded;
  if (!ok) return null;
  const canvas = new OffscreenCanvas(EMOJI_RASTER_PX, EMOJI_RASTER_PX);
  const ctx = canvas.getContext('2d');
  if (!ctx) return null;
  try {
    ctx.drawImage(img, 0, 0, EMOJI_RASTER_PX, EMOJI_RASTER_PX);
  } catch {
    return null;
  }
  const pngBlob = await canvas.convertToBlob({ type: 'image/png' });
  return new Uint8Array(await pngBlob.arrayBuffer());
}

// Service-worker (Chrome MV3) rasterization path. createImageBitmap
// is the only async image loader available in that context. Firefox
// bug note: without resizeWidth/resizeHeight, createImageBitmap on an
// SVG blob can return a 0×0 phantom that draws as transparent. The
// resize options pre-allocate a fixed target so the decoder can't
// silently skip sizing.
async function rasterizeViaImageBitmap(svgText: string): Promise<Uint8Array | null> {
  const svgBlob = new Blob([svgText], { type: 'image/svg+xml' });
  const bitmap = await createImageBitmap(svgBlob, {
    resizeWidth: EMOJI_RASTER_PX,
    resizeHeight: EMOJI_RASTER_PX,
    resizeQuality: 'high',
  });
  if (bitmap.width === 0 || bitmap.height === 0) {
    bitmap.close?.();
    return null;
  }
  const canvas = new OffscreenCanvas(EMOJI_RASTER_PX, EMOJI_RASTER_PX);
  const ctx = canvas.getContext('2d');
  if (!ctx) return null;
  ctx.drawImage(bitmap, 0, 0, EMOJI_RASTER_PX, EMOJI_RASTER_PX);
  bitmap.close?.();
  const pngBlob = await canvas.convertToBlob({ type: 'image/png' });
  return new Uint8Array(await pngBlob.arrayBuffer());
}

// Populate the emoji image cache for every key actually used in this
// export, in parallel. Failed/missing keys are cached as null so the
// rendering pass skips them without retrying. Keys come from scanDocument.
async function prewarmTwemojiKeys(
  doc: PDFDocument,
  cache: AssetCache,
  keys: Set<string>,
): Promise<void> {
  if (!keys.size) return;
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

// Attach a clickable link annotation to the given rectangle on a
// page. pdf-lib doesn't expose a high-level helper for this, so we
// build the annotation dict directly against the document's low-level
// object graph. Rect coords are in PDF user space (origin bottom-left).
function addLinkAnnotation(
  page: PDFPage,
  x1: number,
  y1: number,
  x2: number,
  y2: number,
  url: string,
) {
  const ctx = page.doc.context;
  const actionDict = ctx.obj({
    Type: 'Action',
    S: 'URI',
    URI: PDFString.of(url),
  });
  const annotDict = ctx.obj({
    Type: 'Annot',
    Subtype: 'Link',
    Rect: [x1, y1, x2, y2],
    Border: [0, 0, 0],
    A: actionDict,
  });
  const annotRef = ctx.register(annotDict);
  const annotsKey = PDFName.of('Annots');
  const existing = page.node.get(annotsKey);
  if (existing instanceof PDFArray) {
    existing.push(annotRef);
  } else {
    page.node.set(annotsKey, ctx.obj([annotRef]));
  }
}

// Scan a rendered text line for URLs and register link annotations
// for each one. Uses String.matchAll under the hood so we can iterate
// without mutating regex state. Skips URL substrings that don't begin
// with http(s):// (e.g. the tail half of a URL that got wrapped to a
// second line — annotating partial URLs would misdirect the click).
function addLinkAnnotationsForLine(
  page: PDFPage,
  line: string,
  x: number,
  y: number,
  size: number,
  weight: PreferredWeight,
  ctx: TextCtx,
) {
  for (const m of line.matchAll(URL_RE)) {
    const raw = m[0];
    const url = trimTrailingPunct(raw);
    if (!url) continue;
    const prefix = line.slice(0, m.index ?? 0);
    const prefixW = safeWidthOf(prefix, weight, ctx, size);
    const urlW = safeWidthOf(url, weight, ctx, size);
    const x1 = x + prefixW;
    const x2 = x1 + urlW;
    // 1pt of space below baseline catches descenders; top goes to y+size
    // so the hit area roughly matches the glyph cap height.
    addLinkAnnotation(page, x1, y - 1, x2, y + size, url);
  }
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
  opts?: { links?: boolean },
) {
  const linkify = opts?.links === true;
  for (const line of lines) {
    ensureSpace(cursor, leading);
    cursor.y -= leading;
    const lineX = ctx.layout.textColX + indent;
    drawMixed(cursor.page, line, lineX, cursor.y, size, weight, ctx, color);
    if (linkify && line.includes('http')) {
      addLinkAnnotationsForLine(cursor.page, line, lineX, cursor.y, size, weight, ctx);
    }
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
    drawLines(cursor, wrapText(quote, 'regular', ctx, ctx.layout.sizeMeta, ctx.layout.textWidth - 12), 'regular', ctx, ctx.layout.sizeMeta, COLOR_META, ctx.layout.leadMeta, 12, { links: true });
  }

  if (m.forwarded?.originalAuthor) {
    cursor.y -= 4;
    const fwd = `[forwarded from ${m.forwarded.originalAuthor}]`;
    drawLines(cursor, wrapText(fwd, 'regular', ctx, ctx.layout.sizeMeta, ctx.layout.textWidth), 'regular', ctx, ctx.layout.sizeMeta, COLOR_META, ctx.layout.leadMeta);
    if (m.forwarded.originalText) {
      const truncated = m.forwarded.originalText.length > 300 ? m.forwarded.originalText.slice(0, 300) + '…' : m.forwarded.originalText;
      drawLines(cursor, wrapText(truncated, 'regular', ctx, ctx.layout.sizeBody, ctx.layout.textWidth - 12), 'regular', ctx, ctx.layout.sizeBody, COLOR_BODY, ctx.layout.leadBody, 12, { links: true });
    }
  }

  if (m.subject) {
    cursor.y -= 2;
    drawLines(cursor, wrapText(`Subject: ${m.subject}`, 'bold', ctx, ctx.layout.sizeBody, ctx.layout.textWidth), 'bold', ctx, ctx.layout.sizeBody, COLOR_BODY, ctx.layout.leadBody);
  }

  const text = (m.text || '').replace(/\r\n/g, '\n');
  if (text) {
    cursor.y -= 2;
    drawLines(cursor, wrapText(text, 'regular', ctx, ctx.layout.sizeBody, ctx.layout.textWidth), 'regular', ctx, ctx.layout.sizeBody, COLOR_BODY, ctx.layout.leadBody, 0, { links: true });
  }

  if (Array.isArray(m.reactions) && m.reactions.length) {
    const parts = m.reactions.map(r => `${r.emoji} ${r.count}`).join('  ');
    drawLines(cursor, wrapText(parts, 'regular', ctx, ctx.layout.sizeMeta, ctx.layout.textWidth), 'regular', ctx, ctx.layout.sizeMeta, COLOR_META, ctx.layout.leadMeta);
  }

  // Attachments. If a dataUrl is present and embeds as PNG/JPEG we draw
  // the image inline. Everything else (non-image types, failed embeds,
  // or images with no dataUrl because downloadImages was off) falls
  // back to a single text line so the user still sees the filename.
  // When att.href is an http(s) URL, the entire fallback line becomes
  // a clickable link so the user can jump to the original file.
  if (Array.isArray(m.attachments) && m.attachments.length) {
    for (const att of m.attachments) {
      const drewImage = await tryDrawAttachmentImage(cursor, att, ac);
      if (drewImage) continue;
      const label = att.label || att.href || '[attachment]';
      const attLine = `[attachment] ${label}`;
      const wrapped = wrapText(attLine, 'regular', ctx, ctx.layout.sizeMeta, ctx.layout.textWidth);
      const isHttp = !!(att.href && /^https?:\/\//i.test(att.href));
      drawLines(cursor, wrapped, 'regular', ctx, ctx.layout.sizeMeta, COLOR_META, ctx.layout.leadMeta);
      // Attach a link annotation over the single-line case. Multi-line
      // attachment labels would need per-line baseline tracking through
      // page breaks, and attachments rarely wrap, so the common case
      // is enough. The URL is on cursor.page at cursor.y.
      if (isHttp && att.href && wrapped.length === 1) {
        const w = safeWidthOf(wrapped[0], 'regular', ctx, ctx.layout.sizeMeta);
        addLinkAnnotation(
          cursor.page,
          ctx.layout.textColX,
          cursor.y - 1,
          ctx.layout.textColX + w,
          cursor.y + ctx.layout.sizeMeta,
          att.href,
        );
      }
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
