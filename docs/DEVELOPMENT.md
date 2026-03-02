# Development

This document uses commands exactly as defined in `package.json`.

## Requirements

- Node.js LTS
- npm

## Install

```bash
npm install
```

## Run in development

```bash
# Chrome/Edge target
npm run dev

# Firefox target
npm run dev:firefox
```

`wxt.config.ts` has `runner.disabled: true` (browser won't auto-open) and `dev.reloadOnChange: false` (no auto-reload). Load the extension manually — see [MANUAL_INSTALL.md](MANUAL_INSTALL.md).

## Build

```bash
# Chrome/Edge
npm run build

# Firefox
npm run build:firefox
```

Build output folders:

- `.output/chrome-mv3/`
- `.output/firefox-mv2/`

## Create zip packages

```bash
# Chrome/Edge zip
npm run zip

# Firefox zip
npm run zip:firefox
```

## Type check

```bash
npm run check
```

## Recommended local checks before PR

```bash
npm run check
npm run build
```

If your change affects Firefox-specific behavior, also run:

```bash
npm run build:firefox
```

## Quick test checklist

### Core
- [ ] Extension loads without errors.
- [ ] Popup opens and displays correctly.
- [ ] Theme toggle works.
- [ ] Date range inputs work.
- [ ] Export button triggers scraping.
- [ ] Badge updates during scraping.
- [ ] Empty chat shows banner.
- [ ] Options persist across popup close/reopen.

### Export formats
- [ ] JSON export downloads and contains correct data.
- [ ] CSV export downloads and is formatted correctly.
- [ ] HTML export downloads and renders correctly.
- [ ] Text export downloads and reads correctly.
- [ ] Avatar embedding works (HTML).

### Include toggles
- [ ] Replies toggle works.
- [ ] Reactions toggle works.
- [ ] System messages toggle works.
- [ ] Date range filter works.

### Targets
- [ ] Chat export works.
- [ ] Team channel export works.

### Browser-specific
- [ ] **Firefox**: Downloads work (uses blob URL fallback).
- [ ] **Firefox**: Storage persistence works across restarts.

### Large exports (pre-release)
- [ ] Export a chat with 5,000+ messages (all options enabled).
- [ ] Export completes without 64MiB message errors.
- [ ] JSON export file is valid and contains all messages.
- [ ] HTML+images zip export renders correctly in browser.
- [ ] Avatars appear correctly in HTML and JSON exports.
- [ ] Memory usage is stable during large exports.

## Troubleshooting

- If extension changes are not visible, reload the unpacked extension manually.
- If build output looks stale, remove `.output` and `.wxt`, then rebuild.
- If dependencies get out of sync, run `rm -rf node_modules .output .wxt && npm install`.
