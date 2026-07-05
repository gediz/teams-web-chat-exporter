// ESLint flat config. Recommended rule sets with the noisy-but-benign
// rules demoted to warn while the baseline is triaged; promote a rule
// to error once its findings are fixed or individually suppressed.
// Run locally via `pnpm lint`; not wired into CI until rules are at
// error severity (warn-only CI output is noise nobody is forced to read).
import eslint from '@eslint/js';
import tseslint from 'typescript-eslint';
import svelte from 'eslint-plugin-svelte';
import globals from 'globals';

export default tseslint.config(
  {
    ignores: [
      '.output/**',
      '.wxt/**',
      // Generated / vendored assets (postinstall writes twemoji + wasm)
      'src/public/twemoji/**',
      'src/public/wasm/**',
      // Session-memory scratch files, not project code
      '.remember/**',
    ],
  },
  eslint.configs.recommended,
  tseslint.configs.recommended,
  svelte.configs.recommended,
  {
    languageOptions: {
      globals: {
        ...globals.browser,
        ...globals.webextensions,
        // Compile-time constant injected by wxt.config.ts (vite define);
        // declared for TS in src/build-stamp.d.ts.
        __BUILD_STAMP__: 'readonly',
      },
    },
  },
  {
    // TypeScript inside <script lang="ts"> blocks of Svelte components
    files: ['**/*.svelte'],
    languageOptions: {
      parserOptions: { parser: tseslint.parser },
    },
  },
  {
    // Node contexts: build/config files and the maintenance scripts
    files: ['scripts/**/*.mjs', 'wxt.config.ts', 'svelte.config.js', 'eslint.config.js'],
    languageOptions: { globals: { ...globals.node } },
  },
  {
    rules: {
      // The 2026-07 baseline was triaged (fixed, or justified per site
      // with an inline disable) and every clean rule promoted back to its
      // recommended error severity. Two stay demoted on purpose:
      // - no-explicit-any: the per-site typing batch is still in flight.
      // - no-async-promise-executor: the share-resolver rewrite is
      //   deferred until a test harness can lock it (SR-EXEC in the
      //   maintainer notes); the warning IS the tracking signal, so do
      //   not disable it inline.
      '@typescript-eslint/no-explicit-any': 'warn',
      'no-async-promise-executor': 'warn',
      // argsIgnorePattern/varsIgnorePattern keep the `_`-prefix
      // convention; caughtErrors:none matches the deliberate
      // catch { /* noop */ } style; ignoreRestSiblings covers
      // destructure-to-omit.
      '@typescript-eslint/no-unused-vars': ['error', { argsIgnorePattern: '^_', varsIgnorePattern: '^_', caughtErrors: 'none', ignoreRestSiblings: true }],
    },
  },
);
