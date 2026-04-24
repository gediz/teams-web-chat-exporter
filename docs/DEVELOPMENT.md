# Development

This document uses scripts exactly as defined in `package.json`. Invoke them with `pnpm` — this repo uses pnpm as its package manager (lockfile is `pnpm-lock.yaml`).

## Requirements

- Node.js 24+
- pnpm 10+

## Install

```bash
pnpm install
```

The `postinstall` script runs `wxt prepare` and `node scripts/vendor-twemoji.mjs`, which copies the Twemoji SVG set from `@twemoji/svg` and the `hb-subset.wasm` from `harfbuzzjs` into `src/public/`. Both subtrees are gitignored.

## Run in development

```bash
# Chrome/Edge target
pnpm dev

# Firefox target
pnpm dev:firefox
```

`wxt.config.ts` has `runner.disabled: true` (browser won't auto-open) and `dev.reloadOnChange: false` (no auto-reload). Load the extension manually — see [MANUAL_INSTALL.md](MANUAL_INSTALL.md).

## Build

```bash
# Chrome/Edge
pnpm build

# Firefox
pnpm build:firefox
```

Build output folders:

- `.output/chrome-mv3/`
- `.output/firefox-mv2/`

## Create zip packages

```bash
# Chrome/Edge zip
pnpm zip

# Firefox zip
pnpm zip:firefox
```

## Type check

```bash
pnpm check
```

## Recommended local checks before PR

```bash
pnpm check
pnpm build
pnpm build:firefox
```

Always run both builds. Chrome MV3 (service worker) and Firefox MV2 (background script) run the same source on different runtimes and some APIs (`createImageBitmap` on SVG blobs, `browser.action` vs `browser.browserAction`) behave differently between them.

## Quick test checklist

### Core
- [ ] Extension loads without errors.
- [ ] Popup opens and displays correctly.
- [ ] Theme toggle works (light/dark).
- [ ] Date range inputs + quick-range pills work.
- [ ] Export button triggers export.
- [ ] Stop button aborts an in-progress export.
- [ ] Badge updates during export.
- [ ] Empty chat shows banner.
- [ ] Options persist across popup close/reopen.
- [ ] Last open page (main / settings / history) is restored on reopen.

### Popup pages
- [ ] Settings page opens, About card shows name, version, source/issue/author links.
- [ ] History page lists recent exports; Open + Show in folder work.
- [ ] Onboarding overlay shows on first launch, disappears after Got it / Skip, stays dismissed.

### Scraping modes
- [ ] API mode fetches messages (look for `[API]` log lines).
- [ ] DOM scroll fallback works when API is unavailable.

### Export formats
- [ ] JSON export downloads and contains correct data.
- [ ] CSV export downloads and is formatted correctly.
- [ ] HTML export downloads and renders correctly.
- [ ] Text (TXT) export downloads and reads correctly.
- [ ] PDF export downloads, text is selectable, hyperlinks clickable.
- [ ] Multi-format selection (2+) produces a single `bundle.zip`.
- [ ] Avatar embedding works (HTML, JSON, PDF).
- [ ] HTML + `Avatars in HTML → Save as separate files` produces `HTML.zip` with an `avatars/` folder.

### Include toggles
- [ ] Replies toggle works.
- [ ] Reactions toggle works (reactor names show in HTML / PDF hover).
- [ ] System messages toggle works.
- [ ] Date range filter works.
- [ ] Inline images toggle works (HTML, PDF).

### Targets
- [ ] Chat export works.
- [ ] Team channel export works.

### Browser-specific
- [ ] **Firefox**: Downloads work (uses blob URLs).
- [ ] **Firefox**: PDF emoji rasterization works (reactions show colour emoji, not tofu).
- [ ] **Firefox**: Storage persistence works across restarts.
- [ ] **Chrome**: Same as above under MV3 service worker.

### Large exports (pre-release)
- [ ] Export a chat with 5,000+ messages (all options enabled, all formats selected).
- [ ] Export completes without 64 MiB message errors.
- [ ] JSON export file is valid and contains all messages.
- [ ] HTML+images zip export renders correctly in browser.
- [ ] PDF renders all messages without missing characters (ligature + subset check).
- [ ] Avatars appear correctly in HTML, JSON, PDF exports.
- [ ] Memory usage is stable during large exports.

## Troubleshooting

- If extension changes are not visible, reload the unpacked extension manually.
- If build output looks stale, remove `.output` and `.wxt`, then rebuild.
- If dependencies get out of sync, run `rm -rf node_modules .output .wxt && pnpm install`.
- If PDF emoji or fonts look broken after an upgrade, also clear `src/public/twemoji/` and `src/public/wasm/` so the `postinstall` step re-vendors them.