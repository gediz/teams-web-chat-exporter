// Bench: validate two claims about the outer-zip step in download.ts:
//   1. Resolving to a Blob (instead of a contiguous Uint8Array) halves the
//      peak memory at finalisation.
//   2. level: 0 (store) is materially faster than level: 6 when the
//      input bytes are already compressed (PDFs, per-chat zips, JPEGs),
//      with negligible final-size impact.
//
// Run:  node --expose-gc scripts/bench-outer-zip.mjs

import { Zip, ZipDeflate } from 'fflate';
import { performance } from 'node:perf_hooks';
import { randomBytes } from 'node:crypto';

const MB = 1_000_000;

// Realistic file mix matching what the outer zip actually contains in a
// 289-chat bundle:
//   - per-chat bundle.zip (already DEFLATE-compressed; incompressible)
//   - per-chat .pdf (mostly already DEFLATE-compressed)
//   - per-chat .json / .csv / .txt / .html (raw text; compressible)
//
// We model "already compressed" with crypto random bytes (DEFLATE can't
// shrink them) and "compressible text" with a repeating pattern that
// DEFLATE does shrink.
function makeFiles({ chatCount, perChatBytes, compressibleFraction }) {
  const files = [];
  const compressibleBytes = Math.floor(perChatBytes * compressibleFraction);
  const incompressibleBytes = perChatBytes - compressibleBytes;

  // Repeating text — high-entropy at first glance but compresses to ~5%.
  const textChunk = Buffer.from(
    '{"timestamp":"2024-01-01T00:00:00Z","author":"someone","text":"a typical chat message body, repeated to fill bytes"}\n'.repeat(64),
  );

  function buildText(size) {
    const buf = Buffer.allocUnsafe(size);
    let off = 0;
    while (off < size) {
      const n = Math.min(textChunk.length, size - off);
      textChunk.copy(buf, off, 0, n);
      off += n;
    }
    return new Uint8Array(buf.buffer, buf.byteOffset, buf.byteLength);
  }

  for (let i = 0; i < chatCount; i++) {
    const folder = `chat-${String(i).padStart(4, '0')}`;
    if (compressibleBytes > 0) {
      files.push({ path: `${folder}/messages.json`, data: buildText(Math.floor(compressibleBytes * 0.5)) });
      files.push({ path: `${folder}/messages.csv`,  data: buildText(Math.floor(compressibleBytes * 0.3)) });
      files.push({ path: `${folder}/messages.txt`,  data: buildText(Math.floor(compressibleBytes * 0.2)) });
    }
    if (incompressibleBytes > 0) {
      // Random bytes mimic a per-chat bundle.zip + a PDF.
      files.push({ path: `${folder}/bundle.zip`, data: new Uint8Array(randomBytes(Math.floor(incompressibleBytes * 0.6))) });
      files.push({ path: `${folder}/messages.pdf`, data: new Uint8Array(randomBytes(incompressibleBytes - Math.floor(incompressibleBytes * 0.6))) });
    }
  }
  return files;
}

// Variant A: current production code — accumulates chunks then COPIES into
// a contiguous Uint8Array before resolving.
function buildZipUint8Array(files, level) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let zipError = null;
    const zip = new Zip((err, chunk, final) => {
      if (err) { zipError = err; return; }
      if (chunk && chunk.length) chunks.push(chunk);
      if (final) {
        if (zipError) { reject(zipError); return; }
        let total = 0;
        for (const c of chunks) total += c.length;
        const out = new Uint8Array(total);
        let offset = 0;
        for (const c of chunks) { out.set(c, offset); offset += c.length; }
        resolve(out);
      }
    });
    (async () => {
      try {
        for (const f of files) {
          await new Promise(r => setImmediate(r));
          if (zipError) throw zipError;
          const item = new ZipDeflate(f.path, { level });
          zip.add(item);
          item.push(f.data, true);
        }
        zip.end();
      } catch (e) { reject(e); }
    })();
  });
}

// Variant B: proposed refactor — resolves to a Blob constructed directly
// from the chunks array. Browsers (and Node 18+) keep chunk references
// without materialising a single contiguous buffer.
//
// `levelFn(path)` lets us pick a deflate level per file. Pass a constant
// for variants that compress everything the same way; pass a smart fn
// to skip compression on already-compressed payloads (.zip, .pdf, .jpg,
// etc.) while still compressing raw text (.json, .csv, .txt, .html).
function buildZipBlob(files, levelFn) {
  return new Promise((resolve, reject) => {
    const chunks = [];
    let zipError = null;
    const zip = new Zip((err, chunk, final) => {
      if (err) { zipError = err; return; }
      if (chunk && chunk.length) chunks.push(chunk);
      if (final) {
        if (zipError) { reject(zipError); return; }
        resolve(new Blob(chunks, { type: 'application/zip' }));
      }
    });
    (async () => {
      try {
        for (const f of files) {
          await new Promise(r => setImmediate(r));
          if (zipError) throw zipError;
          const lvl = typeof levelFn === 'function' ? levelFn(f.path) : levelFn;
          const item = new ZipDeflate(f.path, { level: lvl });
          zip.add(item);
          item.push(f.data, true);
        }
        zip.end();
      } catch (e) { reject(e); }
    })();
  });
}

// Already-compressed file extensions: re-deflating them is pure waste.
const ALREADY_COMPRESSED = /\.(zip|pdf|jpe?g|png|gif|webp|mp[34]|webm|gz|7z|rar|bz2|xz)$/i;
const smartLevel = (path) => ALREADY_COMPRESSED.test(path) ? 0 : 6;

function snapshotRss() {
  if (global.gc) global.gc();
  return process.memoryUsage().rss;
}

async function runVariant(name, fn, files) {
  // Sampling: poll RSS every 50 ms during the run so we capture the peak,
  // not just before/after (the concat-into-Uint8Array spike is brief).
  let peakRss = snapshotRss();
  const baselineRss = peakRss;
  const sampler = setInterval(() => {
    const r = process.memoryUsage().rss;
    if (r > peakRss) peakRss = r;
  }, 50);

  const t0 = performance.now();
  const out = await fn(files);
  const t1 = performance.now();

  clearInterval(sampler);
  const outBytes = out instanceof Blob ? out.size : out.byteLength;

  return {
    name,
    timeMs: Math.round(t1 - t0),
    outBytes,
    peakDeltaMB: Math.round((peakRss - baselineRss) / MB),
    peakRssMB: Math.round(peakRss / MB),
    outputType: out instanceof Blob ? 'Blob' : 'Uint8Array',
  };
}

function fmtMB(bytes) { return (bytes / MB).toFixed(0) + 'MB'; }

async function main() {
  // Three workloads:
  //   small  — sanity check, fast
  //   medium — single-chat-class bundle
  //   large  — approximates the 289-chat bundle that hung in the field
  const workloads = [
    { label: 'small',  chatCount: 50,   perChatBytes: 1 * MB },     // ~50 MB
    { label: 'medium', chatCount: 200,  perChatBytes: 2 * MB },     // ~400 MB
    { label: 'large',  chatCount: 300,  perChatBytes: 5 * MB },     // ~1.5 GB
  ];
  const compressibleFraction = 0.30;

  for (const w of workloads) {
    console.log(`\n=== workload: ${w.label} (${w.chatCount} chats × ${fmtMB(w.perChatBytes)} = ${fmtMB(w.chatCount * w.perChatBytes)} input) ===`);
    const files = makeFiles({ ...w, compressibleFraction });
    const totalInput = files.reduce((a, f) => a + f.data.byteLength, 0);
    console.log(`  generated ${files.length} files, ${fmtMB(totalInput)} total input`);

    const variants = [
      { name: 'A: Uint8Array, level=6 (current)',     fn: f => buildZipUint8Array(f, 6) },
      { name: 'B: Blob,       level=6',               fn: f => buildZipBlob(f, 6) },
      { name: 'C: Uint8Array, level=0',               fn: f => buildZipUint8Array(f, 0) },
      { name: 'D: Blob,       level=0',               fn: f => buildZipBlob(f, 0) },
      { name: 'E: Blob,       smart per-extension',   fn: f => buildZipBlob(f, smartLevel) },
    ];

    const results = [];
    for (const v of variants) {
      // Drop any cached state from prior variant
      if (global.gc) { global.gc(); global.gc(); }
      const r = await runVariant(v.name, v.fn, files);
      results.push(r);
      console.log(`  ${v.name.padEnd(38)}  time=${String(r.timeMs).padStart(6)}ms  out=${fmtMB(r.outBytes).padStart(7)}  peakΔ=${String(r.peakDeltaMB).padStart(5)}MB  rss=${String(r.peakRssMB).padStart(5)}MB`);
    }

    // Highlight the two head-to-head comparisons that matter:
    const a = results.find(r => r.name.startsWith('A'));
    const b = results.find(r => r.name.startsWith('B'));
    const d = results.find(r => r.name.startsWith('D'));
    if (a && b) {
      const memSaved = a.peakDeltaMB - b.peakDeltaMB;
      const memSavedPct = a.peakDeltaMB > 0 ? Math.round(100 * memSaved / a.peakDeltaMB) : 0;
      console.log(`  → Blob vs Uint8Array (level 6): peak Δ saved ${memSaved}MB (${memSavedPct}% of variant A's spike)`);
    }
    if (a && d) {
      const speedup = a.timeMs / Math.max(d.timeMs, 1);
      const sizeDiffPct = Math.round(100 * (d.outBytes - a.outBytes) / a.outBytes);
      console.log(`  → Blob+level0 vs Uint8Array+level6: ${speedup.toFixed(2)}× faster, output size diff ${sizeDiffPct >= 0 ? '+' : ''}${sizeDiffPct}%`);
    }
    const e = results.find(r => r.name.startsWith('E'));
    if (a && e) {
      const speedup = a.timeMs / Math.max(e.timeMs, 1);
      const sizeDiffPct = Math.round(100 * (e.outBytes - a.outBytes) / a.outBytes);
      const memSaved = a.peakDeltaMB - e.peakDeltaMB;
      console.log(`  → Blob+smart vs Uint8Array+level6: ${speedup.toFixed(2)}× faster, output size diff ${sizeDiffPct >= 0 ? '+' : ''}${sizeDiffPct}%, peak Δ saved ${memSaved}MB`);
    }
    // Free workload memory before next round
    files.length = 0;
  }
}

main().catch(e => { console.error(e); process.exit(1); });
