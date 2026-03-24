# CLAUDE.md - Assistant Guide

This file gives quick, accurate context for AI assistants working in this repository.

## Project Snapshot

- Name: Teams Chat Exporter
- Version: see `package.json` and `wxt.config.ts` (keep in sync)
- Stack: WXT + TypeScript + Svelte 5
- Browsers: Chrome, Edge, Firefox
- Export formats: JSON, CSV, HTML, TXT
- UI locales: 24 files in `src/i18n/locales/`
- Scraping: API-based fetch (primary), DOM scroll (fallback)

## Source of Truth Files

- Runtime/build config: [wxt.config.ts](wxt.config.ts)
- NPM scripts and dependencies: [package.json](package.json)
- Popup UI entry: [src/entrypoints/popup/App.svelte](src/entrypoints/popup/App.svelte)
- Background entry: [src/entrypoints/background.ts](src/entrypoints/background.ts)
- Content entry: [src/entrypoints/content.ts](src/entrypoints/content.ts)
- API client: [src/content/api-client.ts](src/content/api-client.ts)
- API message converter: [src/content/api-converter.ts](src/content/api-converter.ts)
- Teams URL patterns: [src/utils/teams-urls.ts](src/utils/teams-urls.ts)

## Docs To Use

- [README.md](README.md)
- [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md)
- [docs/CONTRIBUTING.md](docs/CONTRIBUTING.md)
- [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md)
- [docs/DEPLOYMENT.md](docs/DEPLOYMENT.md)
- [docs/MANUAL_INSTALL.md](docs/MANUAL_INSTALL.md)
- [docs/TODO.md](docs/TODO.md)

## Assistant Rules

1. Verify claims against code before stating them.
2. Keep wording short and literal.
3. Do not describe behavior that is not in the code.
4. For commands, use scripts exactly as defined in `package.json`.
