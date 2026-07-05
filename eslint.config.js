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
      // Baseline demotions: every rule that fired on the 2026-07 baseline
      // run starts at warn (see the header comment). Promote a rule to
      // error once its findings are fixed or individually suppressed.
      // argsIgnorePattern keeps the established `_`-prefix convention for
      // intentionally unused params; caughtErrors:none matches the many
      // deliberate `catch { /* noop */ }` blocks.
      '@typescript-eslint/no-explicit-any': 'warn',
      '@typescript-eslint/no-unused-vars': ['warn', { argsIgnorePattern: '^_', caughtErrors: 'none' }],
      '@typescript-eslint/ban-ts-comment': 'warn',
      '@typescript-eslint/no-unused-expressions': 'warn',
      'prefer-const': 'warn',
      'no-useless-assignment': 'warn',
      'no-empty': 'warn',
      'preserve-caught-error': 'warn',
      'no-irregular-whitespace': 'warn',
      'no-control-regex': 'warn',
      'no-useless-escape': 'warn',
      'no-async-promise-executor': 'warn',
      'svelte/require-each-key': 'warn',
      'svelte/prefer-svelte-reactivity': 'warn',
      'svelte/infinite-reactive-loop': 'warn',
      'svelte/no-unused-svelte-ignore': 'warn',
    },
  },
);
