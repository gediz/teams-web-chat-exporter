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

import { copyFileSync, mkdirSync, readdirSync, writeFileSync, existsSync } from 'node:fs';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const repoRoot = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const srcDir = join(repoRoot, 'node_modules', '@twemoji', 'svg');
const destDir = join(repoRoot, 'src', 'public', 'twemoji');

if (!existsSync(srcDir)) {
  console.warn(`[vendor-twemoji] ${srcDir} not found — skipping. (Did @twemoji/svg install?)`);
  process.exit(0);
}

mkdirSync(destDir, { recursive: true });
const files = readdirSync(srcDir).filter(f => f.endsWith('.svg'));
for (const f of files) {
  copyFileSync(join(srcDir, f), join(destDir, f));
}
const keys = files.map(f => f.replace(/\.svg$/, ''));
writeFileSync(join(destDir, 'manifest.json'), JSON.stringify(keys));
console.log(`[vendor-twemoji] copied ${files.length} SVGs and wrote manifest.json`);
