import { defineConfig } from 'vitest/config';
import { WxtVitest } from 'wxt/testing';

// WxtVitest wires WXT's Vite config (Svelte compile, path aliases, auto-imports)
// and stubs the `chrome`/`browser` globals with fakeBrowser. happy-dom is chosen
// over jsdom to dodge wxt#1575 (TextEncoder invariant under jsdom + WxtVitest).
export default defineConfig({
  plugins: [WxtVitest()],
  test: {
    environment: 'happy-dom',
    setupFiles: ['./test/setup.ts'],
    include: ['test/**/*.test.ts', 'src/**/*.test.ts'],
    coverage: {
      provider: 'v8',
      // Report the WHOLE source tree (all: true counts files no test touches as
      // 0%), so coverage reflects the real codebase, not just imported modules.
      all: true,
      include: ['src/**/*.{ts,svelte}'],
      exclude: ['src/**/*.d.ts', 'src/i18n/locales/**', '**/*.test.ts'],
      reporter: ['text-summary', 'json-summary', 'html', 'lcov'],
      reportsDirectory: './coverage',
      // Per-file floors, set a few points BELOW current coverage: they protect the
      // now-locked critical files from regression and ratchet up as coverage grows.
      // Deliberately NO global threshold — the global % is a denominator artifact of
      // the three near-0% mega-files (pdf/content/background entrypoints). Raise a
      // floor when a file's coverage rises; add a file here once it earns real specs.
      thresholds: {
        'src/background/zip.ts': { lines: 90, functions: 90 },
        'src/background/download-wait.ts': { lines: 90, functions: 90 },
        'src/background/builders.ts': { lines: 45, functions: 55 },
        'src/content/api-converter.ts': { lines: 35, functions: 60 },
        'src/utils/teams-urls.ts': { lines: 85 },
        'src/utils/messages.ts': { lines: 50 },
      },
    },
  },
});
