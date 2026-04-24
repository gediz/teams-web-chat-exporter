// Runtime font subsetting using HarfBuzz (via the hb-subset WASM binary
// vendored to public/wasm/hb-subset.wasm at install time).
//
// Why: pdf-lib's built-in subsetting (via @pdf-lib/fontkit) drops glyphs
// unreliably, leaving PDFs with scattered-missing-character corruption.
// Embedding the full ~12 MB of Noto fonts works but makes every 3-line
// PDF at least 13 MB. HarfBuzz's subsetter is the gold standard (Chrome
// and Firefox use it internally) — given a font and the set of
// codepoints a document actually uses, it produces a correct, compact
// subset containing only those glyphs plus their dependencies
// (composites, ligature components, etc.). We then hand the subsetted
// bytes to pdf-lib with `subset: false`, bypassing fontkit entirely.
//
// Typical results: a font that's 620 KB in full becomes 3-8 KB after
// subsetting to ~50 English characters. A CJK font normally weighing
// 10 MB drops to 200-500 KB when subsetting to a few hundred actual
// ideographs used in a conversation.
//
// Implementation notes:
// - The WASM is loaded once per service worker lifetime and cached.
// - The `exports` object talks directly to HarfBuzz's C API via WASM —
//   all memory allocation is manual (malloc/free).
// - Every subsetFont call creates + destroys its own HarfBuzz objects.
//   Running multiple subset calls in parallel would race on the shared
//   WASM memory, so the callers serialize (Promise.all with per-font
//   awaits — not truly parallel).

const WASM_PATH = 'wasm/hb-subset.wasm';

// Subset of the harfbuzz C API we call. Exact types are opaque (i32
// handles managed by the WASM); we narrow to number for TS's benefit.
type HbExports = {
  memory: WebAssembly.Memory;
  malloc: (size: number) => number;
  free: (ptr: number) => void;
  hb_blob_create: (data: number, length: number, mode: number, userData: number, destroy: number) => number;
  hb_blob_destroy: (blob: number) => void;
  hb_blob_get_data: (blob: number, lengthPtr: number) => number;
  hb_blob_get_length: (blob: number) => number;
  hb_face_create: (blob: number, index: number) => number;
  hb_face_destroy: (face: number) => void;
  hb_face_reference_blob: (face: number) => number;
  hb_set_add: (set: number, codepoint: number) => void;
  hb_subset_input_create_or_fail: () => number;
  hb_subset_input_destroy: (input: number) => void;
  hb_subset_input_unicode_set: (input: number) => number;
  hb_subset_input_set: (input: number, setType: number) => number;
  hb_subset_or_fail: (face: number, input: number) => number;
};

// hb_subset_sets_t — see harfbuzz/src/hb-subset.h. We only need one.
const HB_SUBSET_SETS_DROP_TABLE_TAG = 3;

// OpenType table tags are 4-char ASCII packed big-endian into a uint32.
// HB_TAG('G','S','U','B') = 0x47535542, etc.
function tag(a: string, b: string, c: string, d: string): number {
  return ((a.charCodeAt(0) & 0xff) << 24) |
         ((b.charCodeAt(0) & 0xff) << 16) |
         ((c.charCodeAt(0) & 0xff) << 8) |
          (d.charCodeAt(0) & 0xff);
}
// Tables we want HarfBuzz to strip from the subset. See the callsite
// for the reasoning — short version: pdf-lib draws glyphs per character
// and measures per character without applying GSUB substitutions or
// GPOS kerning consistently, so leaving those tables in causes visible
// positioning bugs ("confl uence" with a gap after "fl"). Dropping them
// forces simple 1:1 codepoint→glyph output with predictable widths.
const DROP_TABLES = [
  tag('G', 'S', 'U', 'B'),  // ligatures, script shaping
  tag('G', 'P', 'O', 'S'),  // kerning, mark positioning
  tag('G', 'D', 'E', 'F'),  // glyph class defs (pointless without GSUB/GPOS)
  tag('k', 'e', 'r', 'n'),  // legacy kerning table
];

let _loading: Promise<HbExports> | null = null;

function loadHarfbuzz(): Promise<HbExports> {
  if (!_loading) {
    _loading = (async () => {
      const url = chrome.runtime.getURL(WASM_PATH);
      const resp = await fetch(url);
      if (!resp.ok) throw new Error(`hb-subset.wasm fetch failed: HTTP ${resp.status}`);
      const bytes = await resp.arrayBuffer();
      const { instance } = await WebAssembly.instantiate(bytes);
      return instance.exports as unknown as HbExports;
    })();
    // Don't cache a failed load — next call should retry.
    _loading.catch(() => { _loading = null; });
  }
  return _loading;
}

/** Subset a font to exactly the given codepoints. Returns the subsetted
 *  TTF bytes, or null if subsetting failed (corrupt input, empty
 *  codepoint set, WASM error). The original font bytes are unchanged.
 */
export async function subsetFont(fontBytes: Uint8Array, codepoints: Iterable<number>): Promise<Uint8Array | null> {
  const hb = await loadHarfbuzz();
  const heap = new Uint8Array(hb.memory.buffer);

  // Copy font bytes into WASM memory.
  const fontPtr = hb.malloc(fontBytes.byteLength);
  heap.set(fontBytes, fontPtr);

  // HB_MEMORY_MODE_WRITABLE = 2. The blob takes ownership of the buffer;
  // hb_blob_destroy will free it later.
  const blob = hb.hb_blob_create(fontPtr, fontBytes.byteLength, 2, 0, 0);
  const face = hb.hb_face_create(blob, 0);
  hb.hb_blob_destroy(blob);

  const input = hb.hb_subset_input_create_or_fail();
  if (input === 0) {
    hb.hb_face_destroy(face);
    hb.free(fontPtr);
    return null;
  }

  const unicodeSet = hb.hb_subset_input_unicode_set(input);
  let added = 0;
  for (const cp of codepoints) {
    if (cp > 0 && cp <= 0x10FFFF) {
      hb.hb_set_add(unicodeSet, cp);
      added++;
    }
  }
  if (added === 0) {
    // Empty input would subset to a font with no glyphs, which pdf-lib
    // can't embed. Bail and let the caller use the original.
    hb.hb_subset_input_destroy(input);
    hb.hb_face_destroy(face);
    return null;
  }

  // Tell HarfBuzz to strip GSUB/GPOS/GDEF/kern. Measurement in pdf-lib
  // is per-character — without these tables, the glyphs the font
  // reports match the glyphs fontkit emits at draw time, eliminating
  // the ligature-advance-width gap bugs.
  const dropSet = hb.hb_subset_input_set(input, HB_SUBSET_SETS_DROP_TABLE_TAG);
  if (dropSet !== 0) {
    for (const t of DROP_TABLES) hb.hb_set_add(dropSet, t);
  }

  const subsetFace = hb.hb_subset_or_fail(face, input);
  hb.hb_subset_input_destroy(input);

  if (subsetFace === 0) {
    hb.hb_face_destroy(face);
    return null;
  }

  const resultBlob = hb.hb_face_reference_blob(subsetFace);
  const resultLen = hb.hb_blob_get_length(resultBlob);
  if (resultLen === 0) {
    hb.hb_blob_destroy(resultBlob);
    hb.hb_face_destroy(subsetFace);
    hb.hb_face_destroy(face);
    return null;
  }
  // Note: `memory.buffer` view can be invalidated across WASM calls that
  // grow the heap, so we re-view AFTER all WASM activity. Copy out into
  // a detached Uint8Array before anything else happens.
  const dataPtr = hb.hb_blob_get_data(resultBlob, 0);
  const freshHeap = new Uint8Array(hb.memory.buffer);
  const out = new Uint8Array(resultLen);
  out.set(freshHeap.subarray(dataPtr, dataPtr + resultLen));

  hb.hb_blob_destroy(resultBlob);
  hb.hb_face_destroy(subsetFace);
  hb.hb_face_destroy(face);
  return out;
}
