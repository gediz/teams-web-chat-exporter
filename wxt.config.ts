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
  manifest: {
    name: 'Teams Chat Exporter',
    version: '1.3.1',
    description: 'Export Microsoft Teams web chat conversations to JSON, CSV, HTML, or Text with full message history.',
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
    optional_permissions: [],
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
  },
});
