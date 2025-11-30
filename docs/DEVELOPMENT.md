# Developer Guide

This document contains technical details for the Teams Chat Exporter extension, built with the WXT Framework.

## Project Structure

```
.
├── src/
│   ├── entrypoints/
│   │   ├── popup/          # Svelte UI for the popup
│   │   ├── background.ts   # Service worker
│   │   └── content.ts      # Content script (scraper)
│   ├── background/         # Background script modules
│   ├── content/            # Content script modules
│   ├── utils/              # Shared utilities
│   ├── types/              # TypeScript types
│   └── i18n/               # Internationalization
├── public/                 # Static assets (icons)
├── docs/                   # Documentation
├── wxt.config.ts           # WXT configuration
└── package.json            # Dependencies & Scripts
```

## Quick Start

### 1. Install Dependencies

```bash
npm install
```

### 2. Development

```bash
# Chrome (Manual reload required)
npm run dev

# Firefox (Auto-reload enabled)
npm run dev:firefox
```

### 3. Build for Production

```bash
# Build for Chrome/Edge
npm run build

# Build for Firefox
npm run build:firefox

# Build for all browsers
npm run build && npm run build:firefox

# Create store-ready ZIPs
npm run zip              # Chrome
npm run zip:firefox      # Firefox
```

See [DEPLOYMENT.md](DEPLOYMENT.md) for store publishing instructions.

## Testing Checklist

Run these tests on both Chrome and Firefox before release:

### Core Functionality
- [ ] Extension loads without errors.
- [ ] Popup opens and displays correctly.
- [ ] Theme toggle works.
- [ ] Date range inputs work.
- [ ] Export button triggers scraping.
- [ ] Messages collected correctly.
- [ ] Badge updates during scraping.
- [ ] Empty chat shows banner.
- [ ] Date filtering works.

### Export Formats
- [ ] JSON export downloads and contains correct data.
- [ ] CSV export downloads and is formatted correctly.
- [ ] HTML export downloads and renders correctly.
- [ ] Avatar embedding works (HTML).

### Browser-Specifics
- [ ] **Firefox**: Downloads work (uses blob URL fallback).
- [ ] **Firefox**: Storage persistence works across restarts.
- [ ] **Performance**: Memory usage is stable during large exports.

## Troubleshooting

- **Build Fails**: Try `rm -rf .output .wxt node_modules && npm install`.
- **Hot Reload**: Works best in Firefox (`npm run dev:firefox`). Chrome requires manual reload for content scripts.
