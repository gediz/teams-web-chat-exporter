// Type 3 font builder for PDF emoji.
//
// Builds a single PDF Type 3 font where each glyph is a PNG image of an
// emoji. Each glyph's content stream is just "draw the embedded image"
// (`/Im<N> Do`); the font's ToUnicode CMap maps glyph codes back to the
// emoji's Unicode codepoints, including multi-codepoint sequences (ZWJ
// joiners, variation selectors, skin-tone modifiers).
//
// Why Type 3: it's the PDF spec's mechanism for arbitrary-content glyphs
// in selectable text streams. Modern PDF readers (Chrome PDFium, Firefox
// pdf.js, Acrobat, Apple Preview, pdftotext) all extract glyph codes
// through the ToUnicode CMap on text selection / copy / search, so
// emoji become first-class selectable text instead of inline images.
//
// Constraints:
// - Single-byte Encoding with Differences caps usage at 256 glyphs per
//   font. Real-world chat exports peak at ~100 unique emoji; we throw
//   on overflow rather than silently splitting (multi-font split is a
//   future enhancement).
// - Each glyph occupies a 1000x1000 box in font units (FontMatrix
//   scales to 1 user unit per glyph). The image inside each glyph is
//   shifted down by 200 font units so it sits on the text baseline
//   visually, matching other emoji-in-PDF implementations.

import {
  PDFDocument,
  PDFName,
  PDFNumber,
  PDFRawStream,
  type PDFRef,
  type PDFImage,
} from 'pdf-lib';

/** One emoji to include in the font. */
export type EmojiEntry = {
  /** Twemoji filename stem, e.g. "1f604" or "1f926-200d-2642-fe0f". */
  key: string;
  /** The actual codepoints to map back via ToUnicode. */
  codepoints: number[];
  /** Pre-embedded PNG (via `doc.embedPng(bytes)`). */
  image: PDFImage;
};

/** Result of buildType3EmojiFont. */
export type EmojiFontResult = {
  /** Indirect reference to the Type 3 font; register under a page's
   *  /Resources /Font dict to use it. */
  ref: PDFRef;
  /** Maps emoji key -> 0-based glyph code. Use with `glyphCodeToHex`
   *  to emit Tj operators. */
  glyphCodeByKey: Map<string, number>;
};

/** Maximum glyphs per Type 3 font given single-byte encoding. */
const MAX_GLYPHS_PER_FONT = 256;

/** Visual Y shift inside each CharProc, in font units. Drops the image
 *  ~0.2 em below the baseline so it visually aligns with surrounding
 *  text (which sits ascent-above-descent around the baseline). */
const Y_SHIFT_FONT_UNITS = -200;

/** Build a Type 3 font from the given emoji entries. Idempotent w.r.t.
 *  the input array (no mutation); each call produces a new font dict.
 *  Throws on empty input or overflow past MAX_GLYPHS_PER_FONT. */
export function buildType3EmojiFont(
  doc: PDFDocument,
  entries: EmojiEntry[],
): EmojiFontResult {
  if (entries.length === 0) {
    throw new Error('buildType3EmojiFont: at least one entry required');
  }
  if (entries.length > MAX_GLYPHS_PER_FONT) {
    throw new Error(
      `buildType3EmojiFont: ${entries.length} entries exceeds single-byte ` +
      `font limit of ${MAX_GLYPHS_PER_FONT}. Future enhancement: split into ` +
      `multiple Type 3 fonts (FE1, FE2, ...).`,
    );
  }

  const glyphCodeByKey = new Map<string, number>();
  for (let i = 0; i < entries.length; i++) {
    glyphCodeByKey.set(entries[i].key, i);
  }

  // CharProcs: one content stream per glyph that draws its embedded image.
  // The d0 operator declares glyph width (1000 font units = 1.0 em).
  const charProcRefs: Record<string, PDFRef> = {};
  for (let i = 0; i < entries.length; i++) {
    const streamText = `1000 0 d0
q
1000 0 0 1000 0 ${Y_SHIFT_FONT_UNITS} cm
/Im${i} Do
Q`;
    const bytes = new TextEncoder().encode(streamText);
    const stream = PDFRawStream.of(
      doc.context.obj({ Length: bytes.length }),
      bytes,
    );
    charProcRefs[`em${i}`] = doc.context.register(stream);
  }

  // The Type 3 font's own /Resources /XObject must include each image
  // referenced by the CharProcs. PDFImage shares the same indirect ref
  // across all referencing dicts, so we don't pay for image duplication.
  const xobjectDict: Record<string, PDFRef> = {};
  for (let i = 0; i < entries.length; i++) {
    xobjectDict[`Im${i}`] = entries[i].image.ref;
  }

  // Differences: [0 /em0 /em1 /em2 ...] assigns char codes 0..N-1 to
  // the named CharProcs in order.
  const differencesArr: (PDFNumber | PDFName)[] = [PDFNumber.of(0)];
  for (let i = 0; i < entries.length; i++) {
    differencesArr.push(PDFName.of(`em${i}`));
  }

  // All glyphs are full-em wide. Per-glyph customization isn't useful
  // since each emoji image is square and rendered at FontSize.
  const widthsArr = entries.map(() => PDFNumber.of(1000));

  // ToUnicode CMap: lets viewers reconstruct the original Unicode
  // codepoints on text extraction. Single-byte input range <00..FF>;
  // values are UTF-16BE hex strings.
  const toUnicodeBody = buildToUnicodeCMap(entries);
  const toUnicodeBytes = new TextEncoder().encode(toUnicodeBody);
  const toUnicodeStream = PDFRawStream.of(
    doc.context.obj({ Length: toUnicodeBytes.length }),
    toUnicodeBytes,
  );
  const toUnicodeRef = doc.context.register(toUnicodeStream);

  const fontDict = doc.context.obj({
    Type: 'Font',
    Subtype: 'Type3',
    FontBBox: [0, 0, 1000, 1000],
    FontMatrix: [0.001, 0, 0, 0.001, 0, 0],
    CharProcs: doc.context.obj(charProcRefs),
    Encoding: doc.context.obj({
      Type: 'Encoding',
      Differences: doc.context.obj(differencesArr),
    }),
    FirstChar: 0,
    LastChar: entries.length - 1,
    Widths: doc.context.obj(widthsArr),
    Resources: doc.context.obj({
      XObject: doc.context.obj(xobjectDict),
      ProcSet: doc.context.obj(['PDF', 'ImageC']),
    }),
    ToUnicode: toUnicodeRef,
  });
  const ref = doc.context.register(fontDict);

  return { ref, glyphCodeByKey };
}

/** Convert a glyph code (0..255) to the 2-char uppercase hex string
 *  used as the operand of a Tj operator. Throws on out-of-range input. */
export function glyphCodeToHex(code: number): string {
  if (!Number.isInteger(code) || code < 0 || code > 255) {
    throw new Error(`glyphCodeToHex: code ${code} out of range [0, 255]`);
  }
  return code.toString(16).padStart(2, '0').toUpperCase();
}

function buildToUnicodeCMap(entries: EmojiEntry[]): string {
  const bfchars = entries.map((entry, i) => {
    const code = i.toString(16).padStart(2, '0').toUpperCase();
    const utf16 = codepointsAsUtf16BEHex(entry.codepoints);
    return `<${code}> <${utf16}>`;
  }).join('\n');

  return `/CIDInit /ProcSet findresource begin
12 dict begin
begincmap
/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def
/CMapName /Adobe-Identity-UCS def
/CMapType 2 def
1 begincodespacerange
<00> <FF>
endcodespacerange
${entries.length} beginbfchar
${bfchars}
endbfchar
endcmap
CMapName currentdict /CMap defineresource pop
end
end`;
}

function codepointsAsUtf16BEHex(codepoints: number[]): string {
  const hexBytes: string[] = [];
  for (const cp of codepoints) {
    if (cp > 0xFFFF) {
      // Supplementary plane: encode as UTF-16 surrogate pair.
      const v = cp - 0x10000;
      const hi = 0xD800 | (v >> 10);
      const lo = 0xDC00 | (v & 0x3FF);
      hexBytes.push(((hi >> 8) & 0xFF).toString(16).padStart(2, '0'));
      hexBytes.push((hi & 0xFF).toString(16).padStart(2, '0'));
      hexBytes.push(((lo >> 8) & 0xFF).toString(16).padStart(2, '0'));
      hexBytes.push((lo & 0xFF).toString(16).padStart(2, '0'));
    } else {
      hexBytes.push(((cp >> 8) & 0xFF).toString(16).padStart(2, '0'));
      hexBytes.push((cp & 0xFF).toString(16).padStart(2, '0'));
    }
  }
  return hexBytes.join('').toUpperCase();
}
