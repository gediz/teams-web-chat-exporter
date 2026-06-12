// Fast Uint8Array -> base64.
//
// Two constraints shape this:
//
// 1. SPEED. Building a binary string one byte at a time
//    (`bin += String.fromCharCode(b)`) is O(n) string concatenations that
//    SpiderMonkey/V8 turn into a string-interning / rope-flattening storm. A
//    Firefox CPU profile of an image-heavy export measured that pattern (in the
//    image-fetch -> data URL paths) as the largest single share of CPU. So we
//    batch into chunks and append n/CHUNK times instead.
//
// 2. CROSS-COMPARTMENT SAFETY. In a Firefox content script, image bytes can
//    arrive from the PAGE world (e.g. via window.postMessage from the urlp
//    page-helper). Such a Uint8Array is an Xray-wrapped, cross-compartment
//    object. `bytes.subarray()` / `String.fromCharCode.apply(null, bytes)` do a
//    species (`.constructor`) lookup, which the wrapper DENIES ("Permission
//    denied to access property constructor") — and the throw, deep inside a
//    fetch-response handler, left the image Promise unresolved and HUNG the
//    export. Plain numeric index reads (`bytes[k]`) ARE permitted on wrapped
//    typed arrays, so we copy each chunk into a local plain array by index and
//    only `apply` over that local array.
export function uint8ToBase64(bytes: Uint8Array): string {
  const n = bytes.length;
  const CHUNK = 8192; // stays well under the fromCharCode.apply argument cap
  const parts: string[] = [];
  for (let i = 0; i < n; i += CHUNK) {
    const end = Math.min(i + CHUNK, n);
    const chunk = new Array<number>(end - i);
    for (let k = i; k < end; k++) chunk[k - i] = bytes[k]; // index reads only — Xray-safe
    parts.push(String.fromCharCode.apply(null, chunk));
  }
  return btoa(parts.join(''));
}
