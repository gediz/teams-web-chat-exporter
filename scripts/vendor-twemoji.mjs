// Copy Twemoji SVG assets from node_modules into src/public/twemoji so
// WXT's static-asset pipeline picks them up at build time. Runs from the
// package's `postinstall` script.
//
// Why vendor-via-copy rather than ship in the repo: the emoji set is
// ~3,700 small files. Committing them would bloat every `git clone` with
// 19 MB of assets that are better modeled as a dependency. Copying at
// install time keeps the repo clean while preserving offline builds.
//
// Also generates a manifest.json listing all available emoji keys so the
// PDF builder can answer "is this codepoint-sequence renderable?" in O(1)
// without one HTTP 404 per unknown character.

import { copyFileSync, mkdirSync, readdirSync, writeFileSync, existsSync, statSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';

const require = createRequire(import.meta.url);
const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');

// --- Twemoji SVGs ---------------------------------------------------------

const twemojiSrc = join(repoRoot, 'node_modules', '@twemoji', 'svg');
const twemojiDest = join(repoRoot, 'src', 'public', 'twemoji');

if (existsSync(twemojiSrc)) {
  mkdirSync(twemojiDest, { recursive: true });
  const files = readdirSync(twemojiSrc).filter(f => f.endsWith('.svg'));
  for (const f of files) copyFileSync(join(twemojiSrc, f), join(twemojiDest, f));
  const keys = files.map(f => f.replace(/\.svg$/, ''));
  writeFileSync(join(twemojiDest, 'manifest.json'), JSON.stringify(keys));
  console.log(`[vendor] copied ${files.length} Twemoji SVGs + manifest.json`);
} else {
  console.warn(`[vendor] ${twemojiSrc} not found — skipping Twemoji copy`);
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
