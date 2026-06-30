import { Zip, ZipPassThrough, deflateSync } from 'fflate';

// `mtime` (optional) sets the entry's archived modified-date. fflate packs it
// into the DOS date field using LOCAL getters, so a Date is written as the user's
// local wall-clock. NOTE: whether the extracted file actually keeps this date is
// EXTRACTOR-DEPENDENT — Info-ZIP unzip / 7-Zip / WinRAR / macOS / Linux honor it,
// but Windows File Explorer "Extract All" re-stamps Mark-of-the-Web downloads to
// extraction time, and cloud/Android extractors drop it. The Date must be in
// fflate's DOS range (year 1980-2099) and valid, or fflate throws / writes junk;
// callers pass only validated dates.
type ZipFile = { path: string; data: Uint8Array; mtime?: Date };

// CRC-32 (fflate doesn't export one). Used to stamp pre-deflated entries.
const CRC_T = (() => {
  const t = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
    t[i] = c >>> 0;
  }
  return t;
})();
function crc32(d: Uint8Array): number {
  let c = 0xFFFFFFFF;
  for (let i = 0; i < d.length; i++) c = CRC_T[(c ^ d[i]) & 0xFF] ^ (c >>> 8);
  return (c ^ 0xFFFFFFFF) >>> 0;
}

// A pre-deflated ZIP entry: the pushed bytes are ALREADY raw DEFLATE; crc/size
// are of the ORIGINAL data. This is byte-identical to what fflate's ZipDeflate
// emits (verified), but holds NO deflate state — so it does not pin the V8
// allocation generation (see the comment on makeZipItem below). It implements
// just the minimal fflate Zip-entry surface (filename/compression/crc/size +
// ondata/push) that Zip.add reads for the local header + data descriptor.
class PreDeflatedEntry {
  filename: string;
  compression = 8;
  crc: number;
  size: number;
  // fflate reads `mtime` off the entry object generically (at zip.add time), so
  // declaring it here is enough for the DOS date to be written — no other plumbing.
  mtime?: Date;
  ondata!: (err: unknown, dat: Uint8Array, final: boolean) => void;
  constructor(filename: string, crc: number, size: number, mtime?: Date) {
    this.filename = filename;
    this.crc = crc;
    this.size = size;
    this.mtime = mtime;
  }
  push(data: Uint8Array, final: boolean) { this.ondata(null, data, final); }
}

// Files whose bytes are already compressed (PDF streams, JPEG/PNG/WebP, nested
// zips, MP3/MP4/WebM audio-video, gzipped archives) are STORED via ZipPassThrough
// (compression 0): bytes pass straight to output, no retained buffer. Everything
// else is DEFLATED — but one-shot, not via fflate's streaming ZipDeflate.
//
// Why one-shot: fflate's streaming Zip keeps every ZipDeflate's Deflate object
// alive in Zip.u until end(). Each Deflate retains its input buffer (this.d.b)
// AND a hash table (state .h/.p, Uint16Array), and — worse — keeping ~1000+
// long-lived Deflate objects pins a V8 allocation generation so the per-push
// deflate CHURN never gets swept until finish(). Measured: a 294-chat-shaped
// workload held ~1.5 GB of unreclaimable churn this way (heap-retainers on a peak
// snapshot showed 925 MB Uint16Array + 782 MB Uint8Array of fflate state).
// Deflating each file with deflateSync (working memory freed on return) and
// adding the result as a PreDeflatedEntry drops that to ~the live set: the same
// workload fell 1777 MB -> 290 MB. Output is byte-identical to ZipDeflate.
//
// (The download-size trade for STORING media still stands: deflating media would
// only shrink ~7% but would re-introduce per-entry deflate buffers, so media
// stays stored.)
const ALREADY_COMPRESSED_RE = /\.(zip|pdf|jpe?g|png|gif|webp|mp[34]|webm|gz|7z|rar|bz2|xz)$/i;

function addEntry(zip: Zip, path: string, data: Uint8Array, mtime?: Date): void {
  // mtime must be set BEFORE zip.add() — fflate writes the local header there.
  if (ALREADY_COMPRESSED_RE.test(path)) {
    const item = new ZipPassThrough(path);
    if (mtime) item.mtime = mtime;
    zip.add(item);
    item.push(data, true);
  } else {
    const compressed = deflateSync(data, { level: 6 });   // one-shot; working mem freed on return
    const item = new PreDeflatedEntry(path, crc32(data), data.length, mtime) as unknown as ZipPassThrough;
    zip.add(item);
    item.push(compressed, true);
  }
}

/**
 * Async, streaming zip builder. Yields to the event loop between every
 * file so the SW thread isn't pinned for many seconds straight, which
 * would freeze any popup that's open.
 *
 * Returns a Blob, NOT a Uint8Array. This is load-bearing for large
 * bundles: a contiguous Uint8Array of the final zip bytes (followed by
 * `new Blob([uint8])` to feed `URL.createObjectURL`) doubles peak
 * memory at the worst possible moment. A 289-chat bundle producing a
 * 2.5 GB zip blew past Firefox's allocation cap with "allocation size
 * overflow" at exactly that step. Constructing the Blob from the
 * chunks[] array directly lets the browser keep chunk references (or
 * page to disk for very large Blobs) without ever materialising a
 * contiguous in-memory copy.
 *
 * IMPORTANT: This routes through fflate's *synchronous* streaming `Zip`
 * + `ZipDeflate` classes, NOT the higher-level async `zip()` function.
 * fflate's `zip()` spawns a Web Worker via `new Worker(blob:...)` for
 * parallel compression, which Firefox extension CSP forbids by default
 * — `worker-src` inherits from `script-src 'self' 'wasm-unsafe-eval'`
 * and `blob:` is not allowed. The `Worker` constructor succeeds, but
 * the worker script never runs, so fflate hangs forever waiting for
 * messages. Going through the streaming `Zip` class keeps everything
 * on the main extension thread (no workers, no CSP), and we get the
 * yield benefit by awaiting `setTimeout(0)` between file additions.
 *
 * Compression: level 6 for raw text (.json/.csv/.txt/.html), level 0
 * (store) for already-compressed payloads. Same final size as
 * level-6-everywhere because already-compressed bytes don't shrink
 * further; saves ~4× CPU on the outer-zip step (bench-confirmed).
 */
export function buildZipAsync(files: ZipFile[], label = 'zip'): Promise<Blob> {
  const totalInputBytes = files.reduce((a, f) => a + f.data.byteLength, 0);
  console.log(`[${label}] start: files=${files.length} inputBytes=${totalInputBytes}`);
  const stream = createZipStream(label);
  return (async () => {
    for (const file of files) await stream.add(file.path, file.data, file.mtime);
    return stream.finish();
  })();
}

// An incrementally-fed version of buildZipAsync. Same streaming Zip + per-file
// yield + multi-source Blob (no contiguous copy), but the caller pushes files
// one at a time and finishes when done, instead of handing over the whole
// files[] array up front.
//
// Why it exists: a multi-chat bundle used to accumulate EVERY chat's built bytes
// in an entries[] array AND then build all the fflate buffers at the end, so peak
// was ~2x the uncompressed bundle (~2.8 GB off-heap for a ~300-chat run, invisible
// to JSHeapUsedSize, near Firefox's single-allocation ceiling). Streaming drops
// the separate entries[] copy. It does NOT make peak "output plus one chat":
// fflate's streaming Zip has no API to release a finished entry, so every
// ZipDeflate keeps its internal input buffer (~1x that file) alive until finish().
// The mitigation is makeZipItem — already-compressed files (images/PDF, usually the
// bulk) use ZipPassThrough and retain NOTHING, so only the compressible text
// formats (json/csv/html/txt) hold a ~1x Deflate buffer to finish(). Net at
// finish(): ~the compressed output + ~1x the uncompressed TEXT, roughly half the
// old peak and far less for media-heavy exports — but not zero.
export interface ZipStream {
  /** Compress + append one file. Yields to the event loop first (popup tick).
   *  Optional `mtime` sets the entry's archived modified-date (see ZipFile). */
  add(path: string, data: Uint8Array, mtime?: Date): Promise<void>;
  /** Close the archive and resolve the final multi-source Blob. */
  finish(): Promise<Blob>;
}

export function createZipStream(label = 'zip'): ZipStream {
  const chunks: Uint8Array[] = [];
  let zipError: unknown = null;
  let fileCount = 0;
  let inputBytes = 0;
  const tStart = performance.now();
  let resolveFinal!: (b: Blob) => void;
  let rejectFinal!: (e: unknown) => void;
  const finalPromise = new Promise<Blob>((res, rej) => { resolveFinal = res; rejectFinal = rej; });
  // Guard: if the archive errors before finish() is awaited, this keeps the
  // rejection from surfacing as an unhandled promise rejection in the worker.
  finalPromise.catch(() => { /* observed via finish() */ });

  const zip = new Zip((err, chunk, final) => {
    if (err) { zipError = err; rejectFinal(err); return; }
    if (chunk && chunk.length) chunks.push(chunk);
    if (final) {
      if (zipError) { rejectFinal(zipError); return; }
      let total = 0;
      for (const c of chunks) total += c.length;
      console.log(`[${label}] streaming-done: files=${fileCount} chunks=${chunks.length} chunkSumBytes=${total} inputBytes=${inputBytes} streamMs=${Math.round(performance.now() - tStart)}`);
      // Multi-source Blob constructor: takes the chunks array as-is, no
      // contiguous copy. The browser may keep chunk references or page to disk;
      // either way, peak memory does NOT double.
      resolveFinal(new Blob(chunks as BlobPart[], { type: 'application/zip' }));
    }
  });

  return {
    async add(path: string, data: Uint8Array, mtime?: Date): Promise<void> {
      if (zipError) throw zipError;
      // Yield BEFORE compressing so the popup mount / message handlers get a
      // tick. Per-file compression (deflateSync) is still synchronous.
      await new Promise<void>(r => setTimeout(r, 0));
      if (zipError) throw zipError;
      addEntry(zip, path.replace(/\\/g, '/'), data, mtime);
      fileCount += 1;
      inputBytes += data.byteLength;
      if (fileCount % 50 === 0) {
        console.log(`[${label}] progress: ${fileCount} files, elapsedMs=${Math.round(performance.now() - tStart)}`);
      }
    },
    finish(): Promise<Blob> {
      try { zip.end(); } catch (e) { rejectFinal(e); }
      return finalPromise;
    },
  };
}
