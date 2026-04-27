import { Zip, ZipDeflate } from 'fflate';

type ZipFile = { path: string; data: Uint8Array };

// Files whose bytes are already compressed (PDF streams, JPEG/PNG/WebP,
// nested zips, MP3/MP4/WebM audio-video, gzipped archives). Re-deflating
// them at level 6 burns CPU + holds 32 KB sliding-window buffers per
// stream while producing output the same size or marginally larger.
// Storing them verbatim (level 0) is faster, lower-memory, and yields
// the same final zip size to within rounding.
const ALREADY_COMPRESSED_RE = /\.(zip|pdf|jpe?g|png|gif|webp|mp[34]|webm|gz|7z|rar|bz2|xz)$/i;

function pickLevel(path: string): 0 | 6 {
  return ALREADY_COMPRESSED_RE.test(path) ? 0 : 6;
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
  return new Promise((resolve, reject) => {
    const chunks: Uint8Array[] = [];
    let zipError: unknown = null;

    const totalInputBytes = files.reduce((a, f) => a + f.data.byteLength, 0);
    const tStart = performance.now();
    console.log(`[${label}] start: files=${files.length} inputBytes=${totalInputBytes}`);

    const zip = new Zip((err, chunk, final) => {
      if (err) { zipError = err; return; }
      if (chunk && chunk.length) chunks.push(chunk);
      if (final) {
        if (zipError) { reject(zipError as Error); return; }
        let total = 0;
        for (const c of chunks) total += c.length;
        const tDone = performance.now();
        console.log(`[${label}] streaming-done: chunks=${chunks.length} chunkSumBytes=${total} streamMs=${Math.round(tDone - tStart)}`);
        // Multi-source Blob constructor: takes the chunks array as-is,
        // no contiguous copy. The browser may keep chunk references or
        // page to disk; either way, peak memory does NOT double.
        resolve(new Blob(chunks as BlobPart[], { type: 'application/zip' }));
      }
    });

    (async () => {
      try {
        for (let i = 0; i < files.length; i++) {
          // Yield BEFORE each file so the popup mount / message handlers
          // get a tick. Per-file compression is still synchronous (one
          // call to ZipDeflate.push), but per-file is much shorter than
          // the whole bundle — for a 50 MB JSON file it's about 1 s,
          // and most files in a per-chat folder are much smaller.
          await new Promise<void>(r => setTimeout(r, 0));
          if (zipError) throw zipError;
          const file = files[i];
          const path = file.path.replace(/\\/g, '/');
          const item = new ZipDeflate(path, { level: pickLevel(path) });
          zip.add(item);
          item.push(file.data, true);
          if ((i + 1) % 50 === 0 || i + 1 === files.length) {
            console.log(`[${label}] progress: ${i + 1}/${files.length} elapsedMs=${Math.round(performance.now() - tStart)}`);
          }
        }
        zip.end();
      } catch (e) {
        console.log(`[${label}] error during add-loop: ${e instanceof Error ? e.message : String(e)}`);
        reject(e instanceof Error ? e : new Error(String(e)));
      }
    })();
  });
}
