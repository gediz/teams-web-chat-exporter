// Copy Twemoji SVG assets from node_modules into src/public/twemoji so
// WXT's static-asset pipeline picks them up at build time. Runs from the
// package's `postinstall` script.
//
// Why vendor-via-copy rather than ship in the repo: the emoji set is
// ~3,700 small files. Committing them would bloat every `git clone` with
// 19 MB of assets that are better modeled as a dependency. Copying at
// install time keeps the repo clean while preserving offline builds.
//
// Also generates an index.json listing all available emoji keys so the
// PDF builder can answer "is this codepoint-sequence renderable?" in O(1)
// without one HTTP 404 per unknown character. Named index.json (not
// manifest.json) so the Chrome Web Store package validator does not
// flag the build as having multiple manifests — its check rejects any
// file literally named manifest.json anywhere in the tree, even though
// only the one at the package root is the real extension manifest.

import { copyFileSync, mkdirSync, readdirSync, readFileSync, rmSync, writeFileSync, existsSync, statSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
import { optimize as optimizeSvg } from 'svgo';

const require = createRequire(import.meta.url);
const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');

// --- Twemoji SVGs ---------------------------------------------------------

const twemojiSrc = join(repoRoot, 'node_modules', '@twemoji', 'svg');
const twemojiDest = join(repoRoot, 'src', 'public', 'twemoji');

if (existsSync(twemojiSrc)) {
  // Pack the SVGs into ONE file instead of copying ~3,700 separate files.
  // The store packages the extension as a zip, which DEFLATEs each file
  // independently; thousands of tiny, near-identical SVGs defeat that
  // (the compression dictionary resets at every file boundary), so the
  // set ships at ~4 MB. Concatenating them into a single JSON map lets the
  // zip dedupe across all of them, dropping the download to ~1.4 MB, with
  // no runtime decoder needed (the SW just parses one file).
  //
  // index.json (keys only, tiny) is kept so the per-export "is this emoji
  // renderable?" check stays cheap and the larger pack.json (SVG bodies)
  // is fetched lazily, only when an export actually contains emoji.
  // pack.json/index.json are NOT named manifest.json so the Chrome Web
  // Store validator does not flag the build for multiple manifests.
  // Minify each SVG with SVGO before packing. This is vector-only (no
  // rasterization): it strips metadata/whitespace/redundant attributes and
  // rounds coordinates to 0.1 unit in the 36-unit viewBox (0.2 px at our
  // 72 px raster, invisible). SVGO's default preset keeps viewBox (it no
  // longer removes it), so the runtime width/height injection still scales
  // correctly. Roughly halves the packed size with no visible change.
  const SVGO_CONFIG = { multipass: true, floatPrecision: 1 };
  rmSync(twemojiDest, { recursive: true, force: true });
  mkdirSync(twemojiDest, { recursive: true });
  const files = readdirSync(twemojiSrc).filter(f => f.endsWith('.svg'));
  const pack = {};
  for (const f of files) {
    const raw = readFileSync(join(twemojiSrc, f), 'utf8');
    pack[f.replace(/\.svg$/, '')] = optimizeSvg(raw, SVGO_CONFIG).data;
  }
  writeFileSync(join(twemojiDest, 'pack.json'), JSON.stringify(pack));
  writeFileSync(join(twemojiDest, 'index.json'), JSON.stringify(Object.keys(pack)));
  console.log(`[vendor] packed ${files.length} Twemoji SVGs into pack.json + index.json`);
} else {
  console.warn(`[vendor] ${twemojiSrc} not found — skipping Twemoji pack`);
}

// --- HarfBuzz subsetter WASM ---------------------------------------------
// Copied to the extension so the PDF builder can load it via
// chrome.runtime.getURL in the service worker. harfbuzzjs itself is
// just JS bindings — we load the WASM binary at runtime.

try {
  const hbWasmSrc = require.resolve('harfbuzzjs/hb-subset.wasm');
  const hbDest = join(repoRoot, 'src', 'public', 'wasm');
  mkdirSync(hbDest, { recursive: true });
  copyFileSync(hbWasmSrc, join(hbDest, 'hb-subset.wasm'));
  const size = statSync(join(hbDest, 'hb-subset.wasm')).size;
  console.log(`[vendor] copied hb-subset.wasm (${Math.round(size / 1024)} KB)`);
} catch (err) {
  console.warn(`[vendor] harfbuzzjs not found — skipping hb-subset.wasm copy: ${err?.message || err}`);
}
