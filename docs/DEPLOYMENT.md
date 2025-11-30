# Deployment Guide

Guide for building and deploying the extension to browser stores.

## Prerequisites

See [DEVELOPMENT.md](DEVELOPMENT.md) for:
- Build commands (`npm run build`, `npm run build:firefox`)
- Creating store ZIPs (`npm run zip`, `npm run zip:firefox`)
- Testing checklist

## Local Testing

See [MANUAL_INSTALL.md](MANUAL_INSTALL.md) for loading unpacked extensions in Chrome and Firefox.

## Store Deployment

### Chrome Web Store
1. Build and create ZIP (see [DEVELOPMENT.md](DEVELOPMENT.md))
2. Upload to [Chrome Web Store Developer Dashboard](https://chrome.google.com/webstore/devconsole)
3. Create new item or update existing package

### Firefox Add-ons (AMO)
1. Ensure unique ID is set in `wxt.config.ts`
2. Build and create ZIP (see [DEVELOPMENT.md](DEVELOPMENT.md))
3. Upload to [Firefox Add-ons Developer Hub](https://addons.mozilla.org/developers/)
4. Provide build instructions in review notes (see below)

#### Build Instructions for Reviewers
Copy this into "Notes to Reviewer":
```
# Teams Chat Exporter - Build Instructions

## Overview
Built with WXT (https://wxt.dev/). Source: https://github.com/gediz/teams-web-chat-exporter

## Requirements
- Node.js 18+ (LTS recommended)
- npm 9+

## Build Steps
1. Clone repo: git clone https://github.com/gediz/teams-web-chat-exporter.git
2. Install: npm install
3. Build: npm run build:firefox
4. Output: .output/firefox-mv2/

## Architecture
- src/entrypoints/popup/: Svelte UI
- src/entrypoints/background.ts: Service worker
- src/entrypoints/content.ts: Content script
- wxt.config.ts: Configuration

## Data Collection
No data collection. All exports are local.
```

### Edge Add-ons
1. Build and create ZIP (see [DEVELOPMENT.md](DEVELOPMENT.md))
2. Upload to [Edge Partner Center](https://partner.microsoft.com/dashboard/microsoftedge)
3. Use the Chrome ZIP (same as Chrome Web Store)

## Version Management

1. Update version in `wxt.config.ts`
2. Test in both browsers (see [DEVELOPMENT.md](DEVELOPMENT.md) testing checklist)
3. Create git tag
4. Build release ZIPs (see [DEVELOPMENT.md](DEVELOPMENT.md))
5. Upload to stores

## Build Output

- `.output/chrome-mv3/`: Chrome build (Manifest V3).
- `.output/firefox-mv2/`: Firefox build (Manifest V2).
- `.output/*.zip`: Store-ready files.

## Troubleshooting

- **Module not found**: Run `rm -rf node_modules .output && npm install`.
- **Chrome Obfuscation**: Enable sourcemaps in `wxt.config.ts` if rejected.
- **Firefox Source Review**: Link to GitHub repo in submission notes.
