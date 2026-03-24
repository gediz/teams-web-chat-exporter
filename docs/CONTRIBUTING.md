# Contributing

## Before you start

- Read [DEVELOPMENT.md](DEVELOPMENT.md)
- Run install and checks locally

## Code style used in this repo

- TypeScript + Svelte 5
- Components: PascalCase filenames
- Utility/content modules: kebab-case filenames
- Shared types: `src/types/`
- UI text: locale keys in `src/i18n/locales/*.json` (24 locales)
- Svelte components use mixed syntax:
  - Most components use legacy `export let` props and `createEventDispatcher`.
  - Only `StatusBar.svelte` uses `$props()` runes.
  - Match the style of the file you're editing.

### TypeScript conventions

- Prefer `type` over `interface` for simple structures.
- Use `as const` for literal types.
- Avoid `any` — use `unknown` if the type is truly dynamic.

## Common changes

### Add a translation key

1. Add the key in `src/i18n/locales/en.json`.
2. Add the same key to every other locale file (24 total).
3. Use `t('key', params, lang)` in UI code.

### Add an export format

1. Update output format types in `src/utils/options.ts` and `src/types/shared.ts`.
2. Add builder logic in `src/background/builders.ts`.
3. Update download flow in `src/background/download.ts`.
4. Add UI option in `src/entrypoints/popup/components/FormatSection.svelte`.
5. Add i18n labels.

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

1. Run `npm run check`.
2. Run `npm run build`.
3. If Firefox-related changes were made, run `npm run build:firefox`.
4. Test the feature in browser.
5. Open PR with clear summary and test notes.

## CI in this repo

Workflow file: `.github/workflows/ci.yml`

CI runs on Node 22 with npm 11:

- `npm run check`
- `npm run build`