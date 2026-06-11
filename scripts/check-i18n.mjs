#!/usr/bin/env node
// Locale parity guard. Every file under src/i18n/locales must carry exactly
// en.json's key set (no missing, no extra), no empty string values, and the
// same {placeholder} tokens per key. Catches a silently-desynced locale before
// it ships. Run with `pnpm check:i18n` (also folded into `pnpm check`); exits
// non-zero on any divergence.
import { readdirSync, readFileSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

const dir = join(dirname(fileURLToPath(import.meta.url)), '..', 'src', 'i18n', 'locales');

// Flatten nested objects to dotted keys: { a: { b: 1 } } -> { 'a.b': 1 }.
function flatten(obj, prefix = '', out = {}) {
  for (const [k, v] of Object.entries(obj)) {
    const key = prefix ? `${prefix}.${k}` : k;
    if (v && typeof v === 'object' && !Array.isArray(v)) flatten(v, key, out);
    else out[key] = v;
  }
  return out;
}

const tokensOf = (s) => new Set(typeof s === 'string' ? s.match(/\{[^}]+\}/g) || [] : []);

const files = readdirSync(dir).filter((f) => f.endsWith('.json')).sort();
if (!files.includes('en.json')) {
  console.error('[i18n] reference en.json is missing');
  process.exit(1);
}

const en = flatten(JSON.parse(readFileSync(join(dir, 'en.json'), 'utf8')));
const enKeys = Object.keys(en);
const enKeySet = new Set(enKeys);
let problems = 0;

for (const file of files) {
  if (file === 'en.json') continue;
  const loc = file.slice(0, -5);
  let data;
  try {
    data = flatten(JSON.parse(readFileSync(join(dir, file), 'utf8')));
  } catch (e) {
    console.error(`[i18n] ${loc}: invalid JSON — ${e.message}`);
    problems++;
    continue;
  }
  const keys = new Set(Object.keys(data));
  const missing = enKeys.filter((k) => !keys.has(k));
  const extra = [...keys].filter((k) => !enKeySet.has(k));
  const empty = Object.entries(data)
    .filter(([, v]) => typeof v === 'string' && v.trim() === '')
    .map(([k]) => k);
  const badTokens = enKeys
    .filter((k) => keys.has(k))
    .filter((k) => {
      const a = tokensOf(en[k]);
      const b = tokensOf(data[k]);
      return a.size !== b.size || [...a].some((t) => !b.has(t));
    });

  if (missing.length || extra.length || empty.length || badTokens.length) {
    problems++;
    const cap = (arr) => arr.slice(0, 10).join(', ') + (arr.length > 10 ? ` … (+${arr.length - 10})` : '');
    console.error(`[i18n] ${loc}:`);
    if (missing.length) console.error(`  missing ${missing.length}: ${cap(missing)}`);
    if (extra.length) console.error(`  extra ${extra.length}: ${cap(extra)}`);
    if (empty.length) console.error(`  empty ${empty.length}: ${cap(empty)}`);
    if (badTokens.length) console.error(`  placeholder mismatch ${badTokens.length}: ${cap(badTokens)}`);
  }
}

if (problems) {
  console.error(`\n[i18n] FAIL — ${problems} locale(s) diverge from en.json (${enKeys.length} keys).`);
  process.exit(1);
}
console.log(`[i18n] OK — ${files.length - 1} locales all match en.json (${enKeys.length} keys); no empties, placeholders consistent.`);
