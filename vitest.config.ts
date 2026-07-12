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
  },
});
