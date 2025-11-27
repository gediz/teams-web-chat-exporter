# Deployment Guide

Guide for building, testing, and deploying the extension.

## Quick Reference

```bash
# Development
npm run dev              # Chrome (manual reload)
npm run dev:firefox      # Firefox (auto-reload)

# Production
npm run build            # Chrome build
npm run build:firefox    # Firefox build

# Store ZIPs
npm run zip              # Chrome ZIP
npm run zip:firefox      # Firefox ZIP
```

## Build Commands

### Development
- **Chrome**: `npm run dev`. Output: `.output/chrome-mv3/`. No hot reload.
- **Firefox**: `npm run dev:firefox`. Output: `.output/firefox-mv2/`. Hot reload enabled.

### Production
- **Chrome**: `npm run build`. Minified.
- **Firefox**: `npm run build:firefox`. Minified, Manifest V2.
- **All**: `npm run build && npm run build:firefox`.

### Store ZIPs
- **Chrome**: `npm run zip`. Creates `.output/teams-chat-exporter-*-chrome.zip`.
- **Firefox**: `npm run zip:firefox`. Creates `.output/teams-chat-exporter-*-firefox.zip`.

## Testing

### Chrome
See [Manual Installation Guide](MANUAL_INSTALL.md).
- **Update**: Run `npm run build`, then click refresh on the extension card.
- **Troubleshooting**: If "reloaded too frequently", use `npm run build`.

### Firefox
See [Manual Installation Guide](MANUAL_INSTALL.md).
- **Update**: Auto-reloads with `npm run dev:firefox`.
- **Limitation**: Temporary installs are removed when Firefox closes.

#### Permanent Install (Development)
Use `web-ext` to sign the extension:
```bash
npm install -g web-ext
npm run build:firefox
cd .output/firefox-mv2
web-ext sign --api-key=KEY --api-secret=SECRET
```

## Store Deployment

### Chrome Web Store
1. **Build**: `npm run zip`.
2. **Upload**: Go to [Chrome Web Store Developer Dashboard](https://chrome.google.com/webstore/devconsole).
3. **Submit**: Create new item or update existing package.

### Firefox Add-ons (AMO)
1. **Config**: Set unique ID in `wxt.config.ts`.
2. **Build**: `npm run zip:firefox`.
3. **Upload**: Go to [Firefox Add-ons Developer Hub](https://addons.mozilla.org/developers/).
4. **Review Notes**: Provide build instructions (see below).

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
1. **Build**: `npm run zip`.
2. **Upload**: Go to [Edge Partner Center](https://partner.microsoft.com/dashboard/microsoftedge).
3. **Submit**: Upload the Chrome ZIP.

## Version Management

1. Update version in `wxt.config.ts`.
2. Build and test both browsers.
3. Create git tag.
4. Build release ZIPs.
5. Upload to stores.

## Build Output

- `.output/chrome-mv3/`: Chrome build (Manifest V3).
- `.output/firefox-mv2/`: Firefox build (Manifest V2).
- `.output/*.zip`: Store-ready files.

## Troubleshooting

- **Module not found**: Run `rm -rf node_modules .output && npm install`.
- **Chrome Obfuscation**: Enable sourcemaps in `wxt.config.ts` if rejected.
- **Firefox Source Review**: Link to GitHub repo in submission notes.
