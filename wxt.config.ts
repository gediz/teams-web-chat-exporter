import { defineConfig } from 'wxt';
import { svelte } from '@sveltejs/vite-plugin-svelte';
import { TEAMS_MATCH_PATTERNS, API_FETCH_PATTERNS } from './src/utils/teams-urls';

export default defineConfig({
  srcDir: 'src',
  extensionApi: 'chrome',
  runner: {
    disabled: true, // Don't auto-open browser (you'll load manually)
  },
  dev: {
    reloadOnChange: false, // Disable auto-reload to prevent Chrome throttling
  } as any,
  vite: () => ({
    plugins: [svelte()],
  }),
  // AMO sources zip: WXT's defaults exclude hidden files, node_modules,
  // and tests, but everything else in the repo root ships. `debug/` is
  // a local-only scratch directory (gitignored) — it can hold hundreds
  // of MB of test captures and personal data that must NEVER reach the
  // public sources zip submitted to addons.mozilla.org. Same idea for
  // screenshots and the Claude Code instructions file.
  zip: {
    excludeSources: [
      'debug/**',
      'screenshots/**',
      'CLAUDE.md',
      // Build artefacts that wouldn't normally land here, defensive:
      '.output/**',
      '.wxt/**',
      '*.zip',
    ],
  },
  manifest: ({ manifestVersion }) => ({
    name: 'Teams Chat Exporter',
    version: '1.4.10',
    description: 'Export Microsoft Teams web chat conversations to JSON, CSV, HTML, TXT, or PDF with full message history.',
    homepage_url: 'https://github.com/gediz/teams-web-chat-exporter',
    browser_specific_settings: {
      gecko: {
        id: 'n.gedizaydindogmus@gmail.com',
        strict_min_version: '109.0',
        // @ts-ignore - data_collection_permissions is required by Firefox but not yet in WXT types
        data_collection_permissions: {
          required: ["none"], // This extension does not collect or transmit any data
        },
      } as any,
    },
    // Firefox Add-ons data collection disclosure
    // This extension does not collect or transmit any user data
    // All exports are saved locally to the user's device
    //
    // Optional <all_urls> for the "Image fetch fallback" feature in
    // Settings. Off by default; users opt in deliberately, and the
    // browser shows the permission prompt only when they flip the
    // toggle. Manifest V3 (Chrome) and MV2 (Firefox) split where
    // host wildcards live:
    //   - Chrome MV3: dedicated `optional_host_permissions` key.
    //     Putting URL patterns in `optional_permissions` is REJECTED
    //     by Chrome MV3 — the manifest function gates the wrong key
    //     out per platform.
    //   - Firefox MV2: combined into `optional_permissions`.
    // Listing in optional_* never appears in the install dialog —
    // it's invisible until permissions.request() is called from the
    // Settings toggle.
    ...(manifestVersion === 3
      ? { optional_host_permissions: ['<all_urls>'] as any }
      : { optional_permissions: ['<all_urls>'] as any }
    ),
    permissions: [
      'scripting',
      'activeTab',
      'downloads',
      // downloads.open is needed so the popup can "Open" a saved export
      // via chrome.downloads.open(id); downloads.show is already part of
      // the base 'downloads' permission.
      'downloads.open',
      'storage',
    ],
    host_permissions: [...TEAMS_MATCH_PATTERNS, ...API_FETCH_PATTERNS],
    // page-helpers/urlp-fetcher.js runs in the Teams page's MAIN world
    // so it can use the page's cookie partition for fetches against
    // the asyncgw URL-image proxy. Content script injects it via a
    // <script src=runtime.getURL(...)> element; that requires the file
    // to be web-accessible. Restricted to Teams hosts so the helper
    // can't be loaded from unrelated pages.
    web_accessible_resources: [
      {
        resources: ['page-helpers/urlp-fetcher.js'],
        matches: TEAMS_MATCH_PATTERNS,
      },
    ],
    icons: {
      16: 'icons/action-16.png',
      32: 'icons/action-32.png',
      48: 'icons/action-48.png',
      128: 'icons/action-128.png',
    },
    action: {
      default_title: 'Teams Chat Exporter',
      default_icon: {
        16: 'icons/action-16.png',
        32: 'icons/action-32.png',
        48: 'icons/action-48.png',
        128: 'icons/action-128.png',
      },
    },
  }),
});
