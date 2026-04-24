# Contributing

## Before you start

- Read [DEVELOPMENT.md](DEVELOPMENT.md)
- Run install and checks locally

## Code style used in this repo

- TypeScript + Svelte 5
- Package manager: **pnpm** (the lockfile is `pnpm-lock.yaml` and CI installs via `pnpm install --frozen-lockfile`).
- Components: PascalCase filenames
- Utility/content modules: kebab-case filenames
- Shared types: `src/types/`
- UI text: locale keys in `src/i18n/locales/*.json` (24 locales, all at 100 % parity with `en.json`)
- Svelte components all use the legacy `export let` + `createEventDispatcher` style. No file currently uses `$props()` runes — match the existing style when editing.

### TypeScript conventions

- Prefer `type` over `interface` for simple structures.
- Use `as const` for literal types.
- Avoid `any` — use `unknown` if the type is truly dynamic.

## Common changes

### Add a translation key

1. Add the key in `src/i18n/locales/en.json`.
2. Add the same key to every other locale file (24 total). `t()` falls back to English, but the doc-level rule is parity — don't ship the key to just one locale.
3. Use `t('key', params, lang)` in UI code.

### Add an export format

1. Update `OptionFormat` in `src/utils/options.ts` and the `formats` tuples in `src/types/shared.ts` (`ScrapeOptions`, `BuildOptions`).
2. Add the builder:
   - Simple text formats (like JSON/TXT) serialize inline in `buildExportInternal` in `src/background/download.ts`.
   - Formats with shared renderer code (like CSV/HTML) live in `src/background/builders.ts`.
   - Large async formats (like PDF) get their own module — see `src/background/pdf.ts` for the pattern.
3. Wire the new format into `buildAndDownload`, `buildAndDownloadZip`, and `buildAndDownloadBundle` in `src/background/download.ts`.
4. Add the UI entry in `src/entrypoints/popup/components/FormatSection.svelte` (`allFormats` array + the `format.<id>` i18n key).
5. Add i18n labels for the format and any new settings it introduces.

### Change scraping logic

API-based scraping (primary path):
- API client: `src/content/api-client.ts`
- API response conversion: `src/content/api-converter.ts`

DOM scroll fallback:
- Entry: `src/entrypoints/content.ts`
- Helpers: `src/content/scroll.ts`, `src/content/text.ts`, `src/content/attachments.ts`, etc.

Prefer stable selectors (for example `data-tid` when present).

Teams DOM structure changes without notice. Test scraping changes thoroughly.

## Pull request checklist

1. Run `pnpm check`.
2. Run `pnpm build` AND `pnpm build:firefox` — the two targets (Chrome MV3, Firefox MV2) are separate build outputs and a change can pass one while silently breaking the other.
3. Test the feature in both browsers when practical.
4. Open PR with clear summary and test notes.

## CI in this repo

Workflow file: `.github/workflows/ci.yml`

CI runs on Ubuntu with Node 24 and pnpm 10:

- `pnpm install --frozen-lockfile`
- `pnpm check`
- `pnpm build`