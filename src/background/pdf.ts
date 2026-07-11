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
// - Reactions rendered under the body as "👍 3  Name, Name & N" (reactor
//   names follow the HTML chip rule), or a compact "👍 3  ❤️ 1" line when
//   no reactor names were resolved.
//   Emoji render in color via a per-export Type 3 font whose glyphs
//   are PNG-rasterized Twemoji SVGs. ToUnicode CMap maps glyph codes
//   back to the source codepoints (including multi-CP ZWJ sequences),
//   so emoji are selectable, searchable, and copy/paste-able as text.
//   See pdf-type3-emoji.ts.
// - Attachments: one line per attachment, "[attachment] label". No
//   embedded raster images in v1 (keeps file size predictable).
// - Page breaks: measured per block; if the current block won't fit we
//   add a new page before rendering it.

import {
  PDFDocument,
  PDFName,
  PDFString,
  PDFArray,
  PDFNumber,
  PDFOperator,
  PDFOperatorNames,
  PDFHexString,
  type PDFDict,
  rgb,
  type PDFFont,
  type PDFImage,
  type PDFPage,
} from 'pdf-lib';
import fontkit from '@pdf-lib/fontkit';
import type { Attachment, ExportMessage, ExportMeta, Reaction, ReactorInfo, TableData, MergeRegion } from '../types/shared';
import { subsetFont } from './font-subset';
import { imageSummaryLine } from './builders';
import { rasterizeSvgInDom } from '../utils/svg-rasterize';
import { rasterizeViaOffscreen } from '../utils/offscreen-client';
import {
  buildType3EmojiFont,
  glyphCodeToHex,
  type EmojiEntry,
  type EmojiFontResult,
} from './pdf-type3-emoji';

// Resource names under which our fonts are registered in each page's
// /Resources /Font dict, so the raw operators emitted in drawMixed for
// mixed text+emoji lines can reference them by name. pdf-lib's drawText
// also auto-registers fonts under names like /F1, /F2; both sets of
// names coexist and point to the same underlying font refs.
const TEXT_FONT_REGULAR_NAME = 'FT';
const TEXT_FONT_BOLD_NAME = 'FB';
const TEXT_FONT_CJK_NAME = 'FC';
const TEXT_FONT_KR_NAME = 'FK';
const EMOJI_FONT_RESOURCE_NAME = 'FE';

// Match-anywhere regex for URL detection inside rendered text. Kept
// deliberately conservative — URLs end at whitespace, any character that's
// unusual in real URLs (quote, paren, bracket, curly), or any non-ASCII
// character (so a link stops at CJK text / full-width punctuation rather than
// running through a Chinese sentence). Trailing punctuation (.,;!?) is trimmed
// off post-match to avoid "visit..." swallowing the sentence's trailing period.
const URL_RE = /https?:\/\/[^\s<>"')\]}\u0080-\uFFFF]+/g;

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

// Noto family, chosen for international coverage. Noto Sans covers
// Latin, Cyrillic, Greek, Hebrew, and basic Arabic. Noto Sans SC covers
// Han ideographs plus Hiragana and Katakana (10 MB, the heaviest font in
// the bundle). Noto Sans SC does NOT contain Korean Hangul, so Hangul is
// served by a separate Noto Sans KR (issue #28). Emoji are handled
// separately via a per-export Type 3 font assembled from Twemoji rasters,
// so we don't bundle Noto Emoji.
// Fonts ship as TTF and are subset at runtime by HarfBuzz. WOFF2 would be
// smaller, but the only viable service-worker decoder (wawoff2) relies on
// dynamic code evaluation, which the MV3 extension CSP forbids; see git
// history for that attempt.
const FONT_REGULAR_PATH = 'fonts/NotoSans-Regular.ttf';
const FONT_BOLD_PATH = 'fonts/NotoSans-Bold.ttf';
const FONT_CJK_PATH = 'fonts/NotoSansSC-Regular.ttf';
const FONT_KR_PATH = 'fonts/NotoSansKR-Regular.ttf';

// Twemoji vendored set. index.json enumerates the ~3,700 available emoji
// keys so we can tell "this is an emoji I can render" from a regular
// character in O(1) without fetching anything. The SVG bodies live in a
// single pack.json (see scripts/vendor-twemoji.mjs and TWEMOJI_PACK_PATH
// below), fetched lazily only when an export actually has emoji. Named
// index.json / pack.json, not manifest.json, so the Chrome Web Store
// package validator does not flag the build for multiple manifests (its
// check rejects any file literally named manifest.json anywhere in the
// tree, even though only the one at the package root is the real manifest).
const TWEMOJI_MANIFEST_PATH = 'twemoji/index.json';
// SVG bodies for all emoji, packed into one file (see scripts/vendor-twemoji.mjs).
// Fetched lazily on first rasterization so emoji-free exports never load it.
const TWEMOJI_PACK_PATH = 'twemoji/pack.json';

// Each unique emoji sequence becomes one glyph in a per-export Type 3
// font. EMOJI_RASTER_PX is the PNG render size used as the glyph image
// — ~2x the PDF font size so the image stays crisp at 100% zoom on
// retina. Rasterization runs in the offscreen document on Chromium MV3
// (the SW lacks DOM), or directly in the Firefox MV2 background page.
const EMOJI_RASTER_PX = 72;

type Fonts = {
  regular: PDFFont;
  bold: PDFFont;
  // null when the document has no CJK / Hangul codepoints (see loadFonts):
  // the 10 MB CJK and 6 MB KR faces are then never fetched, subset or embedded.
  cjk: PDFFont | null;
  kr: PDFFont | null;
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
// fonts, the emoji manifest (what keys exist on disk), the Type 3
// emoji font assembled per export, and the computed layout.
type TextCtx = {
  fonts: Fonts;
  emojiManifest: Set<string>;
  emojiFont: EmojiFontResult | null;
  layout: Layout;
  // Per-document memo for safeWidthOf: `${weight}|${size}|${text}` -> width.
  // wrapText measures every word-and-whitespace token of every paragraph, and
  // those tokens repeat heavily across a chat (spaces, common words, names).
  // _runWidthCache only caches AFTER segmentText splits a string into runs, so
  // it never saved the segmentation pass itself; this caches at the whole-token
  // level, skipping segmentText + readEmojiSequence for a repeated token. Fresh
  // per buildPdf, so it is collected with the document and never accumulates.
  swCache: Map<string, number>;
};

// Embedded-asset caches, scoped to one buildPdf call. Key = source data
// URL for images, avatarId for avatars. Dedupe-by-key cuts both embed
// time and file size when the same asset appears on many messages.
type AssetCache = {
  avatars: Map<string, PDFImage>;    // avatarId -> embedded image
  images: Map<string, PDFImage>;     // dataUrl -> embedded image
  // Type 3 emoji font assembled once per export. null when no emoji
  // were used, or when all rasterizations failed. drawMixed falls back
  // to font-subset rendering (typically tofu) when null.
  emojiFont: EmojiFontResult | null;
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
  // Skip HarfBuzz subsetting and embed the full fonts. Only set by the
  // resilient retry below when a subset build throws (rare per-chat font
  // corruption). Normal builds leave this off so PDFs stay small.
  disableSubset?: boolean;
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
const COLOR_DOT_BG = rgb(0.898, 0.906, 0.922);   // #e5e7eb — reactor initial placeholder fill
const COLOR_DOT_FG = rgb(0.216, 0.255, 0.318);   // #374151 — initials drawn on the placeholder
// Own-message accent. Matches the HTML `.own-msg` left rail color
// (#2563eb) so the visual cue is consistent across formats.
const COLOR_OWN_ACCENT = rgb(0.145, 0.388, 0.922);
// Amber for the partial-export warning banner — matches the HTML
// styling and the History page's amber badge for partial entries.
const COLOR_WARN_BG = rgb(0.996, 0.953, 0.780);   // #fef3c7
const COLOR_WARN_BORDER = rgb(0.961, 0.620, 0.043); // #f59e0b
const COLOR_WARN_TEXT = rgb(0.471, 0.208, 0.067);   // #78350f

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

  const fonts = await loadFonts(doc, scan.codepoints, pdfOptions.disableSubset);
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
    emojiFont: null,
    failedAvatars: new Set(),
    failedImages: new Set(),
  };
  const avatars = meta.avatars || {};

  // Rasterize every emoji used in this export and assemble them into
  // one Type 3 font. Rendering pass then stays synchronous and references
  // glyphs by index in the font's content stream operators.
  await prewarmType3EmojiFont(doc, cache, scan.emojiKeys);

  const ctx: TextCtx = { fonts, emojiManifest, emojiFont: cache.emojiFont, layout, swCache: new Map() };

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
      let textW: number;
      try { textW = fonts.regular.widthOfTextAtSize(label, fsize); } catch { textW = label.length * fsize * 0.5; }
      const x = (layout.pageW - textW) / 2;
      const y = MARGIN / 2;
      // Draw via the shared manual-operator path (stable /FT name) rather than
      // page.drawText, which would re-register the font into each page's
      // resource dict under a fresh uniqueKey (1001 redundant entries on a big
      // chat). drawMixed swallows its own draw errors.
      drawMixed(page, label, x, y, fsize, 'regular', ctx, COLOR_META);
    }
  }

  // Add the per-export fonts to every page's /Resources /Font dict so the
  // manual BT/ET operators in content streams can resolve /FT, /FB, /FC, /FE.
  // Runs after all content is rendered — PDF readers parse resources before
  // content, so as long as the entries are in place at save time, they'll
  // be found.
  registerFontsOnPages(doc, fonts, cache.emojiFont);

  const out = await doc.save();
  return out;
}

// Build a PDF, retrying once with subsetting disabled if the subset build
// throws. Certain glyph sets make HarfBuzz emit a subset that pdf-lib can't
// read back ("Trying to access beyond buffer length"), and that error only
// surfaces at save time — too late for the per-face fallback in
// embedSubsetOrFull. The retry embeds the original full fonts, which are
// always readable. Only the affected chat pays the larger size; every chat
// whose subset build succeeds (the overwhelming majority) stays small.
export async function buildPdfResilient(
  messages: ExportMessage[],
  meta: ExportMeta,
  onProgress?: (done: number, total: number) => void,
  pdfOptions: PdfOptions = {},
): Promise<Uint8Array> {
  try {
    return await buildPdf(messages, meta, onProgress, pdfOptions);
  } catch (e) {
    const msg = (e as Error)?.message || String(e);
    console.log(`[pdf] subset build failed (${msg}); retrying with full fonts`);
    return buildPdf(messages, meta, onProgress, { ...pdfOptions, disableSubset: true });
  }
}

// ----- font loading ---------------------------------------------------

// Font bytes are static extension assets, identical for every chat in a
// bundle. Cache them at module level so a 321-chat export reads each face
// once instead of re-fetching + re-allocating ~17 MB per chat. Read-only:
// subsetFont copies into the WASM heap and the full-font fallback slices
// before embedding, so the cached buffers are never mutated/detached.
const _fontByteCache = new Map<string, ArrayBuffer>();
async function fetchFont(path: string): Promise<ArrayBuffer> {
  const cached = _fontByteCache.get(path);
  if (cached) return cached;
  const url = chrome.runtime.getURL(path);
  const resp = await fetch(url);
  if (!resp.ok) throw new Error(`Failed to load font at ${path}: HTTP ${resp.status}`);
  const buf = await resp.arrayBuffer();
  _fontByteCache.set(path, buf);
  return buf;
}

// Page-number codepoints that always get rendered even if no message
// mentions them. Keeps "1 / 42" style footers from going tofu on a
// chat that happens not to use any digit, slash, or space character.
const PAGE_NUMBER_CODEPOINTS = '0123456789/ ';

async function loadFonts(doc: PDFDocument, codepoints: Set<number>, disableSubset = false): Promise<Fonts> {
  // Only fetch/subset/embed the CJK (10 MB) and KR (6 MB) faces when the
  // document actually routes a codepoint to them. The predicates are the SAME
  // ones pickFontForChar uses, so a face is skipped only when no character
  // would ever select it — no tofu risk. A Latin/Cyrillic chat skips ~16 MB
  // of font fetch + HarfBuzz subsetting; this is the dominant per-chat font
  // cost for non-CJK conversations.
  let needCjk = false, needKr = false;
  for (const cp of codepoints) {
    if (!needKr && cpRoutesToKr(cp)) needKr = true;
    if (!needCjk && cpRoutesToCjk(cp)) needCjk = true;
    if (needCjk && needKr) break;
  }
  const [regularBytes, boldBytes, cjkBytes, krBytes] = await Promise.all([
    fetchFont(FONT_REGULAR_PATH),
    fetchFont(FONT_BOLD_PATH),
    needCjk ? fetchFont(FONT_CJK_PATH) : Promise.resolve(null),
    needKr ? fetchFont(FONT_KR_PATH) : Promise.resolve(null),
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
  // font. That costs more bytes in the PDF but keeps every glyph
  // available — fail-safe, not fail-broken.
  const regular = await embedSubsetOrFull(doc, regularBytes, codepoints, 'regular', disableSubset);
  const bold = await embedSubsetOrFull(doc, boldBytes, codepoints, 'bold', disableSubset);
  const cjk = cjkBytes ? await embedSubsetOrFull(doc, cjkBytes, codepoints, 'cjk', disableSubset) : null;
  const kr = krBytes ? await embedSubsetOrFull(doc, krBytes, codepoints, 'kr', disableSubset) : null;
  return {
    regular: memoizeEncodeText(regular),
    bold: memoizeEncodeText(bold),
    cjk: cjk ? memoizeEncodeText(cjk) : null,
    kr: kr ? memoizeEncodeText(kr) : null,
  };
}

// Memoize a font's encodeText so the DRAW path stops re-shaping. pdf-lib's
// drawText -> encodeText runs a full fontkit layout() pass on every call, even
// for identical strings, and the subset embedder's encodeText also re-includes
// each glyph (includeGlyph is idempotent: same glyph -> same subset id). Chat
// text repeats words, sender names and timestamps heavily, so the same run
// string is drawn many times. Caching the returned PDFHexString collapses those
// repeats to one layout pass. This is the draw-side twin of _runWidthCache,
// which only covered measurement. Output is byte-identical: encodeText is
// deterministic, glyph-registration order is unchanged (a first occurrence is a
// cache miss = real call in document order; a repeat would only ever have
// re-returned an already-assigned subset id), and we hand back the same hex
// object. The cache lives on the per-document PDFFont, so it is collected with
// the document and never accumulates across chats.
function memoizeEncodeText(font: PDFFont): PDFFont {
  const cache = new Map<string, PDFHexString>();
  const orig = font.encodeText.bind(font);
  font.encodeText = (text: string): PDFHexString => {
    let hex = cache.get(text);
    if (hex === undefined) { hex = orig(text); cache.set(text, hex); }
    return hex;
  };
  return font;
}

async function embedSubsetOrFull(
  doc: PDFDocument,
  fontBytes: ArrayBuffer,
  codepoints: Set<number>,
  label = 'font',
  disableSubset = false,
): Promise<PDFFont> {
  // Resilient retry path: skip subsetting entirely and embed the original
  // (known-good) font. Used when a prior subset build threw a font read error
  // ("Trying to access beyond buffer length") that surfaced only at save time.
  if (disableSubset) {
    console.log(`[hb-subset] ${label}: subsetting disabled, embedding FULL font (${fontBytes.byteLength} bytes)`);
    return doc.embedFont(fontBytes.slice(0), { subset: false });
  }
  try {
    const subset = await subsetFont(new Uint8Array(fontBytes), codepoints);
    if (subset && subset.byteLength > 0) {
      // Per-face full->subset sizes land in the diagnostics console log so a
      // real export proves subsetting ran end-to-end. Emitted via console.log
      // (not warn/error) so they never surface in the extension's error badge;
      // the diagnostics buffer captures them by the [hb-subset] prefix
      // regardless of level. (subsetFont pads the bytes so @pdf-lib/fontkit's
      // empty-glyph cbox over-read can't crash save(); see font-subset.ts.)
      console.log(`[hb-subset] ${label}: ${fontBytes.byteLength} -> ${subset.byteLength} bytes (subset OK)`);
      return await doc.embedFont(subset, { subset: false });
    }
    console.log(`[hb-subset] ${label}: subset returned empty; embedding FULL font (${fontBytes.byteLength} bytes)`);
  } catch (e) {
    // A silent catch here is exactly what hid the MV3-CSP failure (WASM
    // blocked -> full-font fallback) across many releases. Surface it as a
    // log line (not an error) so it lands in diagnostics without tripping the
    // extension's error badge.
    console.log(`[hb-subset] ${label}: subset FAILED, embedding FULL font (${fontBytes.byteLength} bytes):`, (e as Error)?.message || e);
  }
  return doc.embedFont(fontBytes.slice(0), { subset: false });
}

// Per-character font selection for NON-emoji characters. Emoji are
// handled separately via Twemoji SVG (see segmentText). Order:
//   1. Korean (Noto Sans KR): Hangul syllables + all Jamo blocks. Must be
//      checked before CJK because Noto Sans SC has no Hangul (issue #28).
//   2. CJK (Noto Sans SC): Han ideographs + Hiragana + Katakana.
//   3. Primary (Noto Sans Regular/Bold): Latin, Cyrillic, Greek, Hebrew,
//      basic Arabic, Western-European extras.
type PreferredWeight = 'regular' | 'bold';
// Korean Hangul: conjoining Jamo (1100-11FF), Compatibility Jamo (3130-318F),
// Jamo Extended-A (A960-A97F), syllables (AC00-D7A3), Jamo Extended-B
// (D7B0-D7FF). NotoSansSC contains none of these, so they must route to
// NotoSansKR or render as tofu (issue #28). U+FFA0 (halfwidth Hangul filler)
// is the one codepoint in the fullwidth block below that SC lacks but KR has.
function cpRoutesToKr(cp: number): boolean {
  return (
    (cp >= 0x1100 && cp <= 0x11FF) ||
    (cp >= 0x3130 && cp <= 0x318F) ||
    (cp >= 0xA960 && cp <= 0xA97F) ||
    (cp >= 0xAC00 && cp <= 0xD7A3) ||
    (cp >= 0xD7B0 && cp <= 0xD7FF) ||
    cp === 0xFFA0
  );
}

// CJK Unified Ideographs + extensions + Hiragana/Katakana + the CJK
// punctuation / fullwidth blocks (absent from Latin Noto Sans). NotoSansSC
// covers every assigned codepoint in these blocks (the lone exception, U+FFA0,
// is handled by cpRoutesToKr above). Callers must check cpRoutesToKr FIRST.
function cpRoutesToCjk(cp: number): boolean {
  return (
    (cp >= 0x3000 && cp <= 0x303F) ||
    (cp >= 0x3040 && cp <= 0x30FF) ||
    (cp >= 0x3400 && cp <= 0x4DBF) ||
    (cp >= 0x4E00 && cp <= 0x9FFF) ||
    (cp >= 0xFE10 && cp <= 0xFE1F) ||
    (cp >= 0xFE30 && cp <= 0xFE4F) ||
    (cp >= 0xFE50 && cp <= 0xFE6F) ||
    (cp >= 0xFF00 && cp <= 0xFFEF) ||
    (cp >= 0x20000 && cp <= 0x2FFFF)
  );
}

function pickFontForChar(ch: string, weight: PreferredWeight, fonts: Fonts): PDFFont {
  const cp = ch.codePointAt(0) ?? 0;
  // The `?? fonts.regular` fallbacks can never be hit: loadFonts only leaves
  // cjk/kr null when no codepoint routes to them (same predicates). They keep
  // the return type non-null and fail safe rather than crashing.
  if (cpRoutesToKr(cp)) return fonts.kr ?? fonts.regular;
  if (cpRoutesToCjk(cp)) return fonts.cjk ?? fonts.regular;
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
// Cheap reject for the ~99% of characters that begin no Twemoji sequence.
// Without it, readEmojiSequence allocates a codepoint array and builds hex keys
// for every character of every message; a CPU profile of a large export put it
// at 16% of all PDF time. The set holds the FIRST codepoint of every manifest
// key, so keycap starters ('#', '*', '0'-'9') stay in while ordinary letters
// drop out. parseInt(key, 16) stops at the '-' separator, giving that first
// codepoint for free. Memoised against the manifest's identity.
let _emojiStarters: Set<number> | null = null;
let _emojiStartersFor: Set<string> | null = null;
function emojiStarterSet(index: Set<string>): Set<number> {
  if (_emojiStartersFor === index && _emojiStarters) return _emojiStarters;
  const starters = new Set<number>();
  for (const key of index) {
    const cp = parseInt(key, 16);
    if (!Number.isNaN(cp)) starters.add(cp);
  }
  _emojiStarters = starters;
  _emojiStartersFor = index;
  return starters;
}

function readEmojiSequence(text: string, i: number, index: Set<string>): { key: string; len: number } | null {
  const cp0 = text.codePointAt(i);
  if (cp0 === undefined) return null;
  // A character that starts no known emoji key cannot match below; the longest
  // candidate the matcher would build still begins with cp0, so this is output
  // identical, it just skips the per-character allocation and Set probes.
  if (!emojiStarterSet(index).has(cp0)) return null;

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
    // Table cells render via bodyBlocks, and a cell's text can carry glyphs
    // (emoji from <img alt>, mention @names) that the flat m.text misses, so
    // scan them too or those glyphs would be tofu / absent in the subset.
    if (Array.isArray(m.bodyBlocks)) {
      for (const b of m.bodyBlocks) {
        if (b.type === 'text') visit(b.text);
        else for (const row of b.table.rows) for (const cell of row) visit(cell);
      }
    }
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

// Rasterize a Twemoji SVG to PNG bytes. Fetches the SVG asset from the
// extension package, then dispatches to a context-appropriate rasterizer:
// the shared DOM-backed pipeline for contexts that have HTMLImageElement
// (Firefox MV2 background, popup), or the offscreen-document client for
// Chromium MV3 service workers (which lack DOM).
// The SVG bodies live in one packed map (twemoji/pack.json: { key: svg }).
// Loaded lazily on first rasterization and cached for the SW lifetime, so a
// PDF export with no emoji never fetches the ~8 MB pack. Single-flight so
// parallel rasterizations don't each fetch it.
let _twemojiPack: Record<string, string> | null = null;
let _twemojiPackLoading: Promise<Record<string, string>> | null = null;
async function loadTwemojiPack(): Promise<Record<string, string>> {
  if (_twemojiPack) return _twemojiPack;
  if (!_twemojiPackLoading) {
    _twemojiPackLoading = (async () => {
      try {
        const resp = await fetch(chrome.runtime.getURL(TWEMOJI_PACK_PATH));
        if (!resp.ok) return {};
        return (await resp.json()) as Record<string, string>;
      } catch {
        return {};
      }
    })().then(p => { _twemojiPack = p; return p; });
  }
  return _twemojiPackLoading;
}

// Rasterized emoji PNG bytes are identical for a given key across every chat,
// but the rasterize step (SVG decode via DOM Image/OffscreenCanvas on Firefox,
// or an offscreen-document IPC round-trip on Chrome) is the expensive part and
// was re-run per chat. Cache successes module-wide so each unique emoji is
// rasterized once per SW lifetime. Failures are NOT cached (may be transient).
const _emojiPngCache = new Map<string, Uint8Array>();
async function rasterizeTwemoji(key: string): Promise<Uint8Array | null> {
  const cached = _emojiPngCache.get(key);
  if (cached) return cached;
  try {
    const pack = await loadTwemojiPack();
    const svgText = pack[key];
    if (!svgText) {
      console.warn(`[pdf] Twemoji ${key}: not in pack`);
      return null;
    }
    // rasterizeSvgInDom and the offscreen client both apply the width/height
    // injection internally; the SW just forwards raw SVG text.
    const png = (typeof Image !== 'undefined' && typeof document !== 'undefined')
      ? await rasterizeSvgInDom(svgText, EMOJI_RASTER_PX)
      : await rasterizeViaOffscreen(svgText, EMOJI_RASTER_PX);
    if (!png) {
      console.warn(`[pdf] Twemoji ${key}: rasterize returned null`);
      return null;
    }
    _emojiPngCache.set(key, png);
    return png;
  } catch (err) {
    console.warn(`[pdf] Twemoji ${key}: rasterize threw`, err);
    return null;
  }
}

// Rasterize every emoji key the document needs and assemble them into a
// single Type 3 font with one glyph per emoji. Result is stored on the
// AssetCache; null when there are no emoji keys, or when all rasterizations
// fail. Rasterization runs in parallel; font construction is synchronous.
async function prewarmType3EmojiFont(
  doc: PDFDocument,
  cache: AssetCache,
  keys: Set<string>,
): Promise<void> {
  if (keys.size === 0) {
    cache.emojiFont = null;
    return;
  }
  const rasterized = await Promise.all(
    Array.from(keys).map(async key => {
      const bytes = await rasterizeTwemoji(key);
      if (!bytes) return null;
      try {
        const image = await doc.embedPng(bytes);
        return { key, image };
      } catch {
        return null;
      }
    }),
  );
  const entries: EmojiEntry[] = [];
  for (const r of rasterized) {
    if (!r) continue;
    entries.push({
      key: r.key,
      codepoints: r.key.split('-').map(c => parseInt(c, 16)),
      image: r.image,
    });
  }
  if (entries.length === 0) {
    cache.emojiFont = null;
    return;
  }
  try {
    cache.emojiFont = buildType3EmojiFont(doc, entries);
  } catch (err) {
    // buildType3EmojiFont throws on >256 entries (single-byte encoding
    // limit). Falls back to font-subset rendering for emoji codepoints,
    // which produces tofu for codepoints Noto Sans doesn't cover.
    console.warn('[pdf] Type 3 emoji font build failed:', err);
    cache.emojiFont = null;
  }
}

// Register text and emoji fonts on every page under stable resource
// names. Manual operators in content streams (the single-BT/ET path in
// drawMixed for mixed text+emoji lines) reference /FT, /FB, /FC, /FK, /FE;
// pdf-lib's auto-assigned /F1, /F2 etc. coexist for drawText-only lines.
// Called once at the end of buildPdf, after all pages are added: PDF
// readers parse resources before content, so writing them at finalize
// time is fine.
function registerFontsOnPages(
  doc: PDFDocument,
  fonts: Fonts,
  emojiFont: EmojiFontResult | null,
): void {
  for (const page of doc.getPages()) {
    const resources = page.node.Resources();
    if (!resources) continue;
    let fontMap = resources.lookup(PDFName.of('Font')) as PDFDict | undefined;
    if (!fontMap) {
      fontMap = doc.context.obj({}) as PDFDict;
      resources.set(PDFName.of('Font'), fontMap);
    }
    fontMap.set(PDFName.of(TEXT_FONT_REGULAR_NAME), fonts.regular.ref);
    fontMap.set(PDFName.of(TEXT_FONT_BOLD_NAME), fonts.bold.ref);
    // cjk/kr are null when skipped (no codepoint routes to them), so no page
    // references their resource name — nothing to register.
    if (fonts.cjk) fontMap.set(PDFName.of(TEXT_FONT_CJK_NAME), fonts.cjk.ref);
    if (fonts.kr) fontMap.set(PDFName.of(TEXT_FONT_KR_NAME), fonts.kr.ref);
    if (emojiFont) {
      fontMap.set(PDFName.of(EMOJI_FONT_RESOURCE_NAME), emojiFont.ref);
    }
  }
}

// Map a text font to its stable resource name. Used by drawMixed's
// single-BT/ET path to emit Tf operators against pdf.ts's known names
// rather than pdf-lib's auto-assigned /F1, /F2 etc.
function getTextFontResourceName(font: PDFFont, fonts: Fonts): string {
  if (font === fonts.bold) return TEXT_FONT_BOLD_NAME;
  if (font === fonts.cjk) return TEXT_FONT_CJK_NAME;
  if (font === fonts.kr) return TEXT_FONT_KR_NAME;
  return TEXT_FONT_REGULAR_NAME;
}

// ----- text segmentation (text + emoji runs) -------------------------

type Run =
  | { type: 'text'; font: PDFFont; text: string }
  | { type: 'emoji'; glyphCode: number };

// Split a string into mixed runs of font-backed text and Type 3 emoji
// glyphs. Emoji lookup uses the per-export Type 3 font's glyph-code map.
// When a key isn't in the font (rasterization failed, or no emoji at
// all in this export), the codepoints fall through to font-subset
// rendering — typically tofu in Noto Sans, which is the acceptable
// long tail.
function segmentText(
  text: string,
  weight: PreferredWeight,
  fonts: Fonts,
  emojiFont: EmojiFontResult | null,
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
      const glyphCode = emojiFont?.glyphCodeByKey.get(seq.key);
      if (glyphCode !== undefined) {
        out.push({ type: 'emoji', glyphCode });
        i += seq.len;
        continue;
      }
      // Known emoji key but no Type 3 glyph (rasterization failed, or
      // emojiFont is null). Render codepoints as raw text runs; Noto Sans
      // subset will typically show tofu, but the codepoint is preserved.
      for (let k = 0; k < seq.len;) {
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

// Per-font width cache: (font, size, text) -> width. The same run text is
// measured up to three times for one line (wrapText's per-word measure, then
// drawMixed's cursor-advance measure, then pdf-lib's own encode when drawing),
// and chat text repeats words heavily. Each measure is a fontkit layout pass,
// so caching collapses the repeats to one. Keyed by the PDFFont object via a
// WeakMap, so it clears automatically when a document's fonts are collected
// (bounds memory to one document's unique runs). The width for a given (font,
// size, text) is deterministic, so the cache is output-identical.
const _runWidthCache = new WeakMap<PDFFont, Map<string, number>>();

// Measure a run's width. Emoji runs render at the font em-size (1em
// square) so we return `size` as the width. Text runs sum each
// character's measured width in its font.
function runWidth(run: Run, size: number): number {
  if (run.type === 'emoji') return size;
  let perFont = _runWidthCache.get(run.font);
  if (!perFont) { perFont = new Map(); _runWidthCache.set(run.font, perFont); }
  const key = `${size}|${run.text}`;
  const cached = perFont.get(key);
  if (cached !== undefined) return cached;
  // Measure the whole run in one fontkit layout pass. The old code summed
  // widthOfTextAtSize per character, which made one layout call per character
  // and showed up as ~40% of PDF-build CPU in a profile. For our subset fonts
  // the two are exactly equal (verified: 0 of 576k real runs differed) and the
  // whole-string width also matches how the run is actually drawn. Fall back to
  // the per-character sum only when the whole-string measure throws on an
  // unmappable character, preserving the old graceful-degradation behaviour.
  let w: number;
  try {
    w = run.font.widthOfTextAtSize(run.text, size);
  } catch {
    w = 0;
    for (const ch of run.text) {
      try { w += run.font.widthOfTextAtSize(ch, size); }
      catch { /* skip unmeasurable */ }
    }
  }
  perFont.set(key, w);
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
    // Track the line's width incrementally. Run widths carry no kerning here
    // (verified: 0 of 576k real runs differed from the per-char sum), so width
    // is additive: width(line + w) === width(line) + width(w). Measuring each
    // word once instead of re-measuring the growing `line + w` every word turns
    // line wrapping from O(words^2) shaping into O(words). Output is identical;
    // the wrap comparison value is the same.
    let lineW = 0;
    for (const w of words) {
      const wW = safeWidthOf(w, weight, ctx, size);
      if (lineW + wW <= maxWidth) {
        line += w;
        lineW += wW;
        continue;
      }
      if (line.trim()) out.push(line.replace(/\s+$/, ''));
      if (wW > maxWidth) {
        const broken = hardBreak(w, weight, ctx, size, maxWidth);
        for (let i = 0; i < broken.length - 1; i++) out.push(broken[i]);
        line = broken[broken.length - 1];
      } else {
        line = w.replace(/^\s+/, '');
      }
      lineW = safeWidthOf(line, weight, ctx, size);
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
  const key = `${weight}|${size}|${text}`;
  const cached = ctx.swCache.get(key);
  if (cached !== undefined) return cached;
  let w = 0;
  for (const run of segmentText(text, weight, ctx.fonts, ctx.emojiFont, ctx.emojiManifest)) {
    w += runWidth(run, size);
  }
  ctx.swCache.set(key, w);
  return w;
}

// Encode a URL into the body of a PDF literal string for a /URI action.
// Two distinct problems, both of which corrupt the PDF if unhandled:
//   1. pdf-lib's PDFString.of writes each JS character as a single low
//      byte, so any non-ASCII character is truncated. A Chinese folder
//      name in a SharePoint URL is the common case: 天 (U+5929) truncates
//      to byte 0x29, i.e. ")", which closes the literal early and makes the
//      rest of the annotation parse as garbage. Percent-encode everything
//      outside printable ASCII so the value is pure, valid-URI ASCII (this
//      also makes those links actually resolve when clicked).
//   2. Inside a PDF literal, "\", "(" and ")" are special. Escape them so a
//      genuine ASCII paren in a URL (e.g. SharePoint "/Doc_(v2)") does not
//      break the literal either.
function encodeUriForPdf(url: string): string {
  let out = '';
  for (const ch of url) {
    const cp = ch.codePointAt(0) ?? 0;
    if (cp <= 0x20 || cp > 0x7E) out += encodeURIComponent(ch);
    else if (ch === '\\' || ch === '(' || ch === ')') out += '\\' + ch;
    else out += ch;
  }
  return out;
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
    URI: PDFString.of(encodeUriForPdf(url)),
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

// A clickable URL fragment on one wrapped line: [start,end) are character
// offsets within that line's text; url is the FULL target, identical for
// every fragment of a URL that wrapped across lines or pages.
interface LinkSpan {
  start: number;
  end: number;
  url: string;
}

// Map the URLs in a paragraph's source text onto the wrapped lines that
// render it, so one hyperlink that word-wraps across several lines (or
// flows across a page break) gets a separate annotation rect per visible
// fragment, all pointing at the full URL. URLs are detected on the source
// text, where surrounding whitespace still bounds them correctly, then
// located in the concatenation of the wrapped lines. hardBreak splits a
// long URL with no inserted separators, so its characters stay contiguous
// in that concatenation; soft (whitespace) breaks only drop whitespace,
// which a URL never contains, so they can't merge two links.
function computeLinkSpans(sourceText: string, lines: string[]): LinkSpan[][] {
  const spans: LinkSpan[][] = lines.map(() => []);
  if (!sourceText || !/https?:\/\//i.test(sourceText)) return spans;
  const urls: string[] = [];
  for (const m of sourceText.matchAll(URL_RE)) {
    const url = trimTrailingPunct(m[0]);
    if (url) urls.push(url);
  }
  if (!urls.length) return spans;
  const concat = lines.join('');
  const lineStart: number[] = [];
  let acc = 0;
  for (const l of lines) { lineStart.push(acc); acc += l.length; }
  let from = 0;
  for (const url of urls) {
    const at = concat.indexOf(url, from);
    if (at < 0) continue;
    from = at + url.length;
    const us = at, ue = at + url.length;
    for (let i = 0; i < lines.length; i++) {
      const ls = lineStart[i], le = ls + lines[i].length;
      const os = Math.max(us, ls), oe = Math.min(ue, le);
      if (os < oe) spans[i].push({ start: os - ls, end: oe - ls, url });
    }
  }
  return spans;
}

// Draw text+emoji at (x,y).
//
// Two paths:
//   * Text-only lines (no emoji runs): drawText for each run. pdf-lib
//     handles BT/ET, font selection, encoding, color. Triple-click in
//     PDF readers selects naturally.
//   * Mixed lines (one or more emoji runs): a single BT/ET text run
//     with Tf font switches between text and the Type 3 emoji font.
//     One text run keeps the line as one selection unit, so triple-
//     click selects the whole line including emoji instead of skipping
//     over them as separate text runs.
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
  const runs = segmentText(text, weight, ctx.fonts, ctx.emojiFont, ctx.emojiManifest);
  // Nothing to draw (empty or whitespace-that-segments-to-nothing line). The
  // old per-run fast path iterated zero times and emitted nothing here; return
  // early so a blank line does not push an empty q/BT/ET/Q text object.
  if (!runs.length) return;

  // One BT/ET block with all runs sharing the same text-matrix origin. Tj
  // operators auto-advance the text position within the block, so no Td is
  // needed per run. This single path serves both text-only and emoji lines.
  // The previous text-only fast path called page.drawText per run, and EVERY
  // such call ran setOrEmbedFont -> newFontDictionary -> uniqueKey, registering
  // the same font into the page's resource dict under a fresh random key on
  // each line (hundreds of thousands of redundant /F.. entries, the build's #1
  // allocation source and a major GC driver) plus pdf-lib's ~14 per-call type
  // assertions. Emitting our own operators against the stable /FT../FE names
  // (registered once per page by registerFontsOnPages) skips both. Glyph
  // positions are unchanged: Tj advances by the same per-glyph widths the old
  // `cx += runWidth(run)` summed (run widths are additive, no kerning here),
  // so the rendered output is visually identical.
  const ops: PDFOperator[] = [];
  // Save state + set fill color, then restore on exit. Scoping the
  // color change avoids leaking it into subsequent operators emitted
  // by other drawText calls on the page.
  ops.push(PDFOperator.of(PDFOperatorNames.PushGraphicsState));
  ops.push(PDFOperator.of(PDFOperatorNames.NonStrokingColorRgb, [
    PDFNumber.of(color.red),
    PDFNumber.of(color.green),
    PDFNumber.of(color.blue),
  ]));
  ops.push(PDFOperator.of(PDFOperatorNames.BeginText));
  ops.push(PDFOperator.of(PDFOperatorNames.MoveText, [PDFNumber.of(x), PDFNumber.of(y)]));

  for (const run of runs) {
    if (run.type === 'text') {
      const fontName = getTextFontResourceName(run.font, ctx.fonts);
      ops.push(PDFOperator.of(PDFOperatorNames.SetFontAndSize, [
        PDFName.of(fontName),
        PDFNumber.of(size),
      ]));
      try {
        ops.push(PDFOperator.of(PDFOperatorNames.ShowText, [run.font.encodeText(run.text)]));
      } catch {
        // Whole-run encode failed (a glyph fontkit can't lay out). Fall back to
        // per-character encoding so every character that DOES encode is still
        // drawn and advances the pen, mirroring runWidth's per-char fallback.
        // A character that still fails contributes no glyph and no advance,
        // exactly as the old per-run drawText path did (its runWidth sum also
        // skipped unmeasurable chars), so later runs on the line never shift.
        // In practice this is unreachable: fonts embed subset:false, so pdf-lib
        // uses the pure CustomFontEmbedder, whose encodeText does not throw for
        // the in-coverage runs segmentText produces. It is defence-in-depth.
        for (const ch of run.text) {
          try { ops.push(PDFOperator.of(PDFOperatorNames.ShowText, [run.font.encodeText(ch)])); }
          catch { /* unmappable character: skip, as the old path did */ }
        }
      }
    } else {
      ops.push(PDFOperator.of(PDFOperatorNames.SetFontAndSize, [
        PDFName.of(EMOJI_FONT_RESOURCE_NAME),
        PDFNumber.of(size),
      ]));
      ops.push(PDFOperator.of(PDFOperatorNames.ShowText, [
        PDFHexString.of(glyphCodeToHex(run.glyphCode)),
      ]));
    }
  }

  ops.push(PDFOperator.of(PDFOperatorNames.EndText));
  ops.push(PDFOperator.of(PDFOperatorNames.PopGraphicsState));

  try {
    page.pushOperators(...ops);
  } catch { /* skip on error */ }
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
  opts?: { linkSpans?: LinkSpan[][]; lineLink?: string },
) {
  const { linkSpans, lineLink } = opts ?? {};
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    ensureSpace(cursor, leading);
    cursor.y -= leading;
    const lineX = ctx.layout.textColX + indent;
    drawMixed(cursor.page, line, lineX, cursor.y, size, weight, ctx, color);
    // Per-URL fragments (message body, quotes). Annotated on the current
    // page after ensureSpace may have advanced it, so a link that crosses
    // a page break stays clickable on every page it spans. 1pt below the
    // baseline catches descenders; the top is y+size for cap height.
    const fragments = linkSpans?.[i];
    if (fragments) {
      for (const span of fragments) {
        const x1 = lineX + safeWidthOf(line.slice(0, span.start), weight, ctx, size);
        const x2 = x1 + safeWidthOf(line.slice(span.start, span.end), weight, ctx, size);
        addLinkAnnotation(cursor.page, x1, cursor.y - 1, x2, cursor.y + size, span.url);
      }
    }
    // Whole-line link (attachment fallback): the entire label is clickable,
    // every wrapped line, across page breaks.
    if (lineLink && line) {
      const w = safeWidthOf(line, weight, ctx, size);
      addLinkAnnotation(cursor.page, lineX, cursor.y - 1, lineX + w, cursor.y + size, lineLink);
    }
  }
}

// Wrap `text` and draw it with per-URL link annotations that survive line
// wrapping and page breaks. Used for message bodies and reply/forward
// quotes, where a long URL would otherwise lose clickability past the
// first line.
function drawLinkedText(
  cursor: Cursor,
  text: string,
  weight: PreferredWeight,
  ctx: TextCtx,
  size: number,
  color: Color,
  leading: number,
  maxWidth: number,
  indent = 0,
) {
  const lines = wrapText(text, weight, ctx, size, maxWidth);
  const spans = computeLinkSpans(text, lines);
  drawLines(cursor, lines, weight, ctx, size, color, leading, indent, { linkSpans: spans });
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

const COLOR_TABLE_HEAD_BG = rgb(0.945, 0.957, 0.97);

// Draw a parsed table as a real bordered grid in the text column. rowspan/
// colspan merges are drawn as one box on the anchor cell (covered positions
// skipped), matching the HTML/TXT output. Cells wrap to their column width;
// row height is the tallest cell. The grid page-breaks at row boundaries and
// keeps each merged block on a single page (ensures its full height up front).
function drawTable(cursor: Cursor, table: TableData, ctx: TextCtx) {
  const W = table.columns;
  const R = table.rows.length;
  if (!W || !R) return;

  const layout = ctx.layout;
  const PAD = 4;
  const PAD2 = 2 * PAD;

  const covered = new Set<string>();
  const anchorAt = new Map<string, MergeRegion>();
  for (const mg of table.merges) {
    anchorAt.set(`${mg.row},${mg.col}`, mg);
    for (let dr = 0; dr < mg.rowspan; dr++) {
      for (let dc = 0; dc < mg.colspan; dc++) {
        if (dr || dc) covered.add(`${mg.row + dr},${mg.col + dc}`);
      }
    }
  }
  const isHeader = (r: number) => r < table.headerRowCount;
  const cellText = (r: number, c: number) => (table.rows[r]?.[c] ?? '').replace(/\r\n/g, '\n');
  const isCovered = (r: number, c: number) => covered.has(`${r},${c}`);

  const avail = Math.max(1, layout.textWidth); // guard against a degenerate (<=0) layout

  // Column widths via min/max content distribution (like CSS auto table
  // layout), measured at a given font size. minW = widest unbreakable token, so
  // a column is never squeezed below the width of a single word (which would
  // wrap "Priority" -> "Priori ty"). maxW = widest full line, the column's ideal
  // width. Sizing each column to at least minW and handing leftover space to the
  // columns that want it most (maxW - minW) stops a long "Summary" column from
  // starving the short "Priority" column when the table is scaled to fit.
  const measureCols = (size: number) => {
    const lineW = (txt: string, weight: PreferredWeight) =>
      txt.split('\n').reduce((m, l) => Math.max(m, safeWidthOf(l, weight, ctx, size)), 0);
    const wordW = (txt: string, weight: PreferredWeight) =>
      txt.split(/\s+/).reduce((m, w) => Math.max(m, safeWidthOf(w, weight, ctx, size)), 0);
    const minW = new Array<number>(W).fill(8 + PAD2);
    const maxW = new Array<number>(W).fill(16);
    for (let r = 0; r < R; r++) {
      for (let c = 0; c < W; c++) {
        if (isCovered(r, c)) continue;
        const a = anchorAt.get(`${r},${c}`);
        if (a && a.colspan > 1) continue;
        const weight: PreferredWeight = isHeader(r) ? 'bold' : 'regular';
        maxW[c] = Math.max(maxW[c], lineW(cellText(r, c), weight) + PAD2);
        minW[c] = Math.max(minW[c], wordW(cellText(r, c), weight) + PAD2);
      }
    }
    // colspan anchors: spread any width deficit across the columns they cover.
    for (let r = 0; r < R; r++) {
      for (let c = 0; c < W; c++) {
        if (isCovered(r, c)) continue;
        const a = anchorAt.get(`${r},${c}`);
        if (!a || a.colspan <= 1) continue;
        const weight: PreferredWeight = isHeader(r) ? 'bold' : 'regular';
        const needMax = lineW(cellText(r, c), weight) + PAD2;
        const needMin = wordW(cellText(r, c), weight) + PAD2;
        let haveMax = 0, haveMin = 0;
        for (let k = 0; k < a.colspan; k++) { haveMax += maxW[c + k]; haveMin += minW[c + k]; }
        if (needMax > haveMax) { const add = (needMax - haveMax) / a.colspan; for (let k = 0; k < a.colspan; k++) maxW[c + k] += add; }
        if (needMin > haveMin) { const add = (needMin - haveMin) / a.colspan; for (let k = 0; k < a.colspan; k++) minW[c + k] += add; }
      }
    }
    // Cap each column's min so one no-space token (long URL / CJK run) can't
    // claim more than half the width and starve the others.
    for (let c = 0; c < W; c++) minW[c] = Math.min(minW[c], maxW[c], avail * 0.5);
    return { minW, maxW, totalMin: minW.reduce((a, b) => a + b, 0) };
  };

  // Auto-shrink the cell font so a very wide table stays on one line instead of
  // wrapping every cell to fragments. Start at the normal body-derived size; if
  // the column minimums overflow the text width, step the font down (0.5pt at a
  // time, to a readable 6pt floor) and take the largest size whose minimums fit.
  // A table that already fits keeps the normal size, so only over-wide tables
  // change. Past the floor we fall back to the old proportional squeeze + wrap.
  const CELL_BASE = Math.max(8, layout.sizeBody - 1);
  const CELL_FLOOR = 6;
  let cellSize = CELL_BASE;
  let cols = measureCols(CELL_BASE);
  if (cols.totalMin >= avail) {
    for (let s = CELL_BASE - 0.5; s >= CELL_FLOOR; s -= 0.5) {
      cellSize = s;
      cols = measureCols(s);
      if (cols.totalMin < avail) break;
    }
  }
  const cellLead = cellSize * 1.32;
  const { minW, maxW, totalMin } = cols;
  let colW: number[];
  if (totalMin >= avail) {
    // Even the minimums overflow (very wide / many-column table): scale them
    // down proportionally; some mid-word wrapping is then unavoidable.
    colW = minW.map(w => (w / totalMin) * avail);
  } else {
    // Each column gets its min; slack goes to the columns that want to grow.
    const slack = avail - totalMin;
    const want = maxW.map((w, c) => Math.max(0, w - minW[c]));
    const totalWant = want.reduce((a, b) => a + b, 0) || 1;
    colW = minW.map((w, c) => w + slack * (want[c] / totalWant));
  }
  const colLeft = (c: number) => { let x = 0; for (let i = 0; i < c; i++) x += colW[i]; return x; };

  // Wrap each anchor cell to its (possibly spanned) width; size the rows.
  const wrapped: Record<string, string[]> = {};
  const rowH = new Array<number>(R).fill(cellLead + 2 * PAD);
  const multiRow: Array<{ r: number; a: MergeRegion; lines: number }> = [];
  for (let r = 0; r < R; r++) {
    for (let c = 0; c < W; c++) {
      if (isCovered(r, c)) continue;
      const a = anchorAt.get(`${r},${c}`);
      const colspan = a ? a.colspan : 1;
      const rowspan = a ? a.rowspan : 1;
      let w = 0; for (let k = 0; k < colspan; k++) w += colW[c + k];
      const weight: PreferredWeight = isHeader(r) ? 'bold' : 'regular';
      const lines = wrapText(cellText(r, c), weight, ctx, cellSize, Math.max(8, w - 2 * PAD));
      wrapped[`${r},${c}`] = lines;
      const contentH = lines.length * cellLead + 2 * PAD;
      if (rowspan === 1) rowH[r] = Math.max(rowH[r], contentH);
      else if (a) multiRow.push({ r, a, lines: lines.length });
    }
  }
  // Grow the last spanned row so a tall rowspan cell's text fits its box.
  for (const { r, a, lines } of multiRow) {
    const contentH = lines * cellLead + 2 * PAD;
    let spanH = 0; for (let k = 0; k < a.rowspan; k++) spanH += rowH[r + k];
    if (contentH > spanH) rowH[r + a.rowspan - 1] += contentH - spanH;
  }

  const usable = layout.pageH - 2 * MARGIN;
  const tableX = layout.textColX;
  cursor.y -= 4; // gap above the table

  for (let r = 0; r < R; r++) {
    const ownCells: number[] = [];
    for (let c = 0; c < W; c++) if (!isCovered(r, c)) ownCells.push(c);
    if (!ownCells.length) { cursor.y -= rowH[r]; continue; } // fully covered row

    // Keep a merged block on one page: reserve the tallest span starting here.
    // Capped at `usable`: a single block taller than a full page can't be kept
    // intact and will overflow the bottom margin (rare; a cell with dozens of
    // wrapped lines). Normal tables, including wide ones, page-break cleanly.
    let keepH = rowH[r];
    for (const c of ownCells) {
      const a = anchorAt.get(`${r},${c}`);
      if (a && a.rowspan > 1) { let h = 0; for (let k = 0; k < a.rowspan; k++) h += rowH[r + k]; keepH = Math.max(keepH, h); }
    }
    ensureSpace(cursor, Math.min(keepH, usable));
    const rowTopY = cursor.y;

    for (const c of ownCells) {
      const a = anchorAt.get(`${r},${c}`);
      const rowspan = a ? a.rowspan : 1;
      const colspan = a ? a.colspan : 1;
      let h = 0; for (let k = 0; k < rowspan; k++) h += rowH[r + k];
      let w = 0; for (let k = 0; k < colspan; k++) w += colW[c + k];
      const x = tableX + colLeft(c);
      const yBottom = rowTopY - h;
      if (isHeader(r)) cursor.page.drawRectangle({ x, y: yBottom, width: w, height: h, color: COLOR_TABLE_HEAD_BG });
      cursor.page.drawRectangle({ x, y: yBottom, width: w, height: h, borderColor: COLOR_RULE, borderWidth: 0.7 });
      const weight: PreferredWeight = isHeader(r) ? 'bold' : 'regular';
      let ty = rowTopY - PAD - cellSize;
      for (const ln of wrapped[`${r},${c}`] || []) {
        if (ty < yBottom + 1) break; // clip text that overflows the cell box
        drawMixed(cursor.page, ln, x + PAD, ty, cellSize, weight, ctx, COLOR_BODY);
        ty -= cellLead;
      }
    }
    cursor.y = rowTopY - rowH[r];
  }
  cursor.y -= 2;
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
  // Participant roster — "who's in the chat", wrapped under the subtitle. The
  // PDF has no collapsible UI, so it's a plain comma-joined line.
  const roster = Array.isArray(meta.participants) ? meta.participants : [];
  if (roster.length) {
    const total = typeof meta.memberCount === 'number' && meta.memberCount > roster.length
      ? `${roster.length} of ${meta.memberCount}` : `${roster.length}`;
    const label = `Participants (${total}): ${roster.map(p => p.name).join(', ')}`;
    cursor.y -= ctx.layout.leadMeta / 2;
    drawLines(cursor, wrapText(label, 'regular', ctx, ctx.layout.sizeMeta, ctx.layout.contentWidth), 'regular', ctx, ctx.layout.sizeMeta, COLOR_META, ctx.layout.leadMeta);
  }
  // Partial-export banner. Amber block immediately under the
  // header so it's the first thing on page 1. Drawn as a filled
  // rectangle with text laid on top via drawMixed (NOT drawLines —
  // drawLines positions text at textColX, the message-body column
  // past the avatar gutter, which is the wrong column for a banner
  // that spans the full content width).
  const partial = (meta as { partial?: { reason?: string } }).partial;
  if (partial) {
    cursor.y -= ctx.layout.blockGap / 2;
    const bodyText = partial.reason === 'network'
      ? 'A network interruption was detected during scraping; some messages may be missing.'
      : 'Some messages may not have fully loaded before the export finished.';
    const titleText = `WARNING: This export may be incomplete. [${partial.reason || 'partial'}]`;
    const bodyLines = wrapText(bodyText, 'regular', ctx, ctx.layout.sizeMeta, ctx.layout.contentWidth - 16);
    const lineHeight = ctx.layout.leadMeta;
    const padY = 8;
    const padX = 8;
    const fontSize = ctx.layout.sizeMeta;
    const totalLines = 1 /* title */ + bodyLines.length;
    const blockH = padY * 2 + totalLines * lineHeight;
    // Rectangle anchored at the TOP of the banner block:
    //   top    = cursor.y
    //   bottom = cursor.y - blockH
    // pdf-lib's drawRectangle takes y=bottom, so y = cursor.y - blockH.
    const top = cursor.y;
    cursor.page.drawRectangle({
      x: MARGIN,
      y: top - blockH,
      width: ctx.layout.contentWidth,
      height: blockH,
      color: COLOR_WARN_BG,
      borderColor: COLOR_WARN_BORDER,
      borderWidth: 1,
    });
    // Lay the lines inside the box. drawMixed uses (x, y) where y is
    // the baseline. First baseline sits one full lineHeight below the
    // top inset, so the ascenders fit within the padY breathing room.
    let baselineY = top - padY - fontSize;
    drawMixed(cursor.page, titleText, MARGIN + padX, baselineY, fontSize, 'bold', ctx, COLOR_WARN_TEXT);
    for (const line of bodyLines) {
      baselineY -= lineHeight;
      drawMixed(cursor.page, line, MARGIN + padX, baselineY, fontSize, 'regular', ctx, COLOR_WARN_TEXT);
    }
    cursor.y = top - blockH;
  }
  // Failed-image summary: a one-line banner below the partial banner, using the
  // same warning-block visual language, shown only when some images failed.
  const imgSummary = imageSummaryLine((meta as { attachmentStats?: ExportMeta['attachmentStats'] }).attachmentStats);
  if (imgSummary) {
    cursor.y -= ctx.layout.blockGap / 2;
    const fontSize = ctx.layout.sizeMeta;
    const lineHeight = ctx.layout.leadMeta;
    const padY = 8;
    const padX = 8;
    const lines = wrapText(`Note: ${imgSummary}`, 'regular', ctx, fontSize, ctx.layout.contentWidth - 16);
    const blockH = padY * 2 + lines.length * lineHeight;
    const top = cursor.y;
    cursor.page.drawRectangle({
      x: MARGIN,
      y: top - blockH,
      width: ctx.layout.contentWidth,
      height: blockH,
      color: COLOR_WARN_BG,
      borderColor: COLOR_WARN_BORDER,
      borderWidth: 1,
    });
    let baselineY = top - padY - fontSize;
    for (const line of lines) {
      drawMixed(cursor.page, line, MARGIN + padX, baselineY, fontSize, 'regular', ctx, COLOR_WARN_TEXT);
      baselineY -= lineHeight;
    }
    cursor.y = top - blockH;
  }
  cursor.y -= ctx.layout.blockGap;
}

// Inline reactor names for one reaction, matching the HTML chip rule
// (renderReactorChip in builders.ts): 1 reactor -> the name, 2-3 -> a comma
// list, 4+ -> "First & N". A self reactor shows as "You". Returns '' when no
// reactors were resolved (self-chat / unresolved counts), so the caller keeps
// the bare "emoji count".
// Up to two initials from a reactor's display name, mirroring the HTML chip
// (builders.ts). For CJK names the first character stands in for the initial.
function reactorInitials(name: string): string {
  return (name || '').split(/\s+/).map(p => p[0] || '').join('').slice(0, 2).toUpperCase() || '?';
}

// Every reactor's name, comma-joined. The PDF has no hover popover, so the
// full list is the only way to see who reacted; drawReactionLine wraps it
// onto continuation lines, and the compact fallback wraps it via wrapText.
function reactorNames(r: Reaction): string {
  const list = r.reactors;
  if (!list || !list.length) return '';
  const nameOf = (x: ReactorInfo) => (x.self ? 'You' : x.name);
  return list.map(nameOf).join(', ');
}

// Draw one reaction starting on its own line: "emoji count", up to 3 reactor
// dots (avatar when resolved, else a placeholder circle with initials) + a "+N"
// marker, then every reactor's name wrapped onto continuation lines. Mirrors
// the HTML chip, with the wrapped name list standing in for its hover popover.
async function drawReactionLine(
  cursor: Cursor,
  r: Reaction,
  ctx: TextCtx,
  ac: { doc: PDFDocument; cache: AssetCache; avatars: Record<string, string> },
) {
  const size = ctx.layout.sizeMeta;
  ensureSpace(cursor, ctx.layout.leadMeta);
  cursor.y -= ctx.layout.leadMeta;
  const baseline = cursor.y;
  const rightEdge = ctx.layout.textColX + ctx.layout.textWidth;
  let x = ctx.layout.textColX;
  const head = `${r.emoji} ${r.count} `;
  drawMixed(cursor.page, head, x, baseline, size, 'regular', ctx, COLOR_META);
  x += safeWidthOf(head, 'regular', ctx, size);

  // Up to 3 reactor dots (avatar when resolved, else a placeholder circle),
  // then "+N" when there are more — matching the HTML chip's dot stack.
  const reactors = Array.isArray(r.reactors) ? r.reactors : [];
  if (ctx.layout.includeAvatars && reactors.length) {
    const dot = size + 1;
    for (const rt of reactors.slice(0, 3)) {
      const img = rt.avatarId ? await getAvatar(ac.doc, ac.cache, rt.avatarId, ac.avatars) : null;
      if (img) {
        cursor.page.drawImage(img, { x, y: baseline - 1, width: dot, height: dot });
      } else {
        // No avatar: gray circle with the reactor's initials centered inside,
        // mirroring the HTML chip's initial badge.
        const cx = x + dot / 2;
        const cy = baseline - 1 + dot / 2;
        cursor.page.drawCircle({ x: cx, y: cy, size: dot / 2, color: COLOR_DOT_BG });
        const initials = reactorInitials(rt.name);
        const initSize = dot * 0.5;
        const iw = safeWidthOf(initials, 'regular', ctx, initSize);
        drawMixed(cursor.page, initials, cx - iw / 2, cy - initSize * 0.34, initSize, 'regular', ctx, COLOR_DOT_FG);
      }
      x += dot + 1;
    }
    if (reactors.length > 3) {
      const more = `+${reactors.length - 3} `;
      drawMixed(cursor.page, more, x + 1, baseline, size, 'regular', ctx, COLOR_META);
      x += safeWidthOf(more, 'regular', ctx, size) + 1;
    } else {
      x += 2;
    }
  }

  // Names: list every reactor, wrapping onto continuation lines aligned under
  // the first name. The first line shares the emoji/dots baseline; each later
  // line advances the cursor. Each reactor's name stays intact (wrap happens at
  // ", " boundaries); a single over-long name is ellipsized to the column width.
  const names = reactorNames(r);
  if (names) {
    const nameStartX = x;
    const avail = rightEdge - nameStartX;
    const parts = names.split(', ');
    let line = '';
    const flush = () => {
      drawMixed(cursor.page, line, nameStartX, cursor.y, size, 'regular', ctx, COLOR_META);
    };
    for (let part of parts) {
      const candidate = line ? `${line}, ${part}` : part;
      if (line && safeWidthOf(candidate, 'regular', ctx, size) > avail) {
        flush();
        ensureSpace(cursor, ctx.layout.leadMeta);
        cursor.y -= ctx.layout.leadMeta;
        line = '';
      }
      if (!line && safeWidthOf(part, 'regular', ctx, size) > avail) {
        while (part.length > 1 && safeWidthOf(`${part}…`, 'regular', ctx, size) > avail) part = part.slice(0, -1);
        part += '…';
      }
      line = line ? `${line}, ${part}` : part;
    }
    if (line) flush();
  }
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
    drawLinkedText(cursor, quote, 'regular', ctx, ctx.layout.sizeMeta, COLOR_META, ctx.layout.leadMeta, ctx.layout.textWidth - 12, 12);
  }

  if (m.forwarded?.originalAuthor) {
    cursor.y -= 4;
    const fwd = `[forwarded from ${m.forwarded.originalAuthor}]`;
    drawLines(cursor, wrapText(fwd, 'regular', ctx, ctx.layout.sizeMeta, ctx.layout.textWidth), 'regular', ctx, ctx.layout.sizeMeta, COLOR_META, ctx.layout.leadMeta);
    if (m.forwarded.originalText) {
      const truncated = m.forwarded.originalText.length > 300 ? m.forwarded.originalText.slice(0, 300) + '…' : m.forwarded.originalText;
      drawLinkedText(cursor, truncated, 'regular', ctx, ctx.layout.sizeBody, COLOR_BODY, ctx.layout.leadBody, ctx.layout.textWidth - 12, 12);
    }
  }

  if (m.subject) {
    cursor.y -= 2;
    drawLines(cursor, wrapText(`Subject: ${m.subject}`, 'bold', ctx, ctx.layout.sizeBody, ctx.layout.textWidth), 'bold', ctx, ctx.layout.sizeBody, COLOR_BODY, ctx.layout.leadBody);
  }

  // API mode: ordered text/table blocks, so pasted tables draw as real grids
  // in their original positions. Falls back to the flat `text` otherwise.
  const blocks = Array.isArray(m.bodyBlocks) ? m.bodyBlocks : null;
  if (blocks && blocks.length) {
    for (const b of blocks) {
      if (b.type === 'table') {
        drawTable(cursor, b.table, ctx);
      } else {
        const bt = b.text.replace(/\r\n/g, '\n');
        if (bt) {
          cursor.y -= 2;
          drawLinkedText(cursor, bt, 'regular', ctx, ctx.layout.sizeBody, COLOR_BODY, ctx.layout.leadBody, ctx.layout.textWidth, 0);
        }
      }
    }
  } else {
    const text = (m.text || '').replace(/\r\n/g, '\n');
    if (text) {
      cursor.y -= 2;
      // Deleted-message tombstone body draws muted (grey) to match the HTML
      // .deleted-body styling; the "[message deleted]" placeholder rides in
      // m.text, so its glyphs are already in the subset (no pre-stamp needed).
      const bodyColor = m.deleted ? COLOR_META : COLOR_BODY;
      drawLinkedText(cursor, text, 'regular', ctx, ctx.layout.sizeBody, bodyColor, ctx.layout.leadBody, ctx.layout.textWidth, 0);
    }
  }

  if (Array.isArray(m.reactions) && m.reactions.length) {
    // When any reactor has a resolved avatar, render each reaction on its own
    // line with the avatar dots (drawReactionLine). Otherwise keep the compact
    // text layout: "emoji count  Name, Name & N" (one per line if any names
    // resolved, else a single "emoji count  emoji count" line).
    const anyAvatars = ctx.layout.includeAvatars
      && m.reactions.some(r => Array.isArray(r.reactors) && r.reactors.some(rt => rt.avatarId));
    if (anyAvatars) {
      for (const r of m.reactions) await drawReactionLine(cursor, r, ctx, ac);
    } else {
      const anyNames = m.reactions.some(r => reactorNames(r));
      const parts = m.reactions.map(r => {
        const names = reactorNames(r);
        return names ? `${r.emoji} ${r.count}  ${names}` : `${r.emoji} ${r.count}`;
      });
      const text = parts.join(anyNames ? '\n' : '  ');
      drawLines(cursor, wrapText(text, 'regular', ctx, ctx.layout.sizeMeta, ctx.layout.textWidth), 'regular', ctx, ctx.layout.sizeMeta, COLOR_META, ctx.layout.leadMeta);
    }
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
      // Append the one-word failure reason (expired / sign-in / ...) when the
      // fetch failed, so a missing item reads "[attachment] label (expired)".
      const why = att.failReason ? ` (${att.failReason})` : '';
      const attLine = `[attachment] ${label}${why}`;
      const wrapped = wrapText(attLine, 'regular', ctx, ctx.layout.sizeMeta, ctx.layout.textWidth);
      const isHttp = !!(att.href && /^https?:\/\//i.test(att.href));
      // When the attachment has an http(s) target, every wrapped line of
      // the label is made clickable. drawLines applies the link per line
      // on the current page, so a label that wraps or crosses a page break
      // stays fully clickable; otherwise it is plain text.
      drawLines(
        cursor, wrapped, 'regular', ctx, ctx.layout.sizeMeta, COLOR_META, ctx.layout.leadMeta, 0,
        isHttp && att.href ? { lineLink: att.href } : undefined,
      );
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
