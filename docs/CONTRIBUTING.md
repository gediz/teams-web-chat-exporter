# Contributing

Thank you for your interest in contributing. This project uses the WXT Framework to support Chrome, Edge, and Firefox.

## Getting Started

See [DEVELOPMENT.md](DEVELOPMENT.md) for setup, build commands, and testing instructions.

## Code Conventions

### Naming
- **Components**: PascalCase (e.g., `DateRangeSection.svelte`)
- **Utilities**: camelCase (e.g., `formatElapsed`)
- **Types**: PascalCase (e.g., `Options`, `Message`)
- **Constants**: UPPER_SNAKE_CASE (e.g., `DEFAULT_OPTIONS`)

### TypeScript
- Use strict mode
- Prefer `type` over `interface` for simple structures
- Use `as const` for literal types
- Avoid `any` - use `unknown` if truly dynamic

### Svelte
- **Version**: Svelte 5 (using legacy syntax, not runes)
- **Props**: Export variables with defaults: `export let theme: Theme = 'light'`
- **Events**: Use `createEventDispatcher<T>()` with typed events
- **Reactivity**: Use `$:` reactive statements
- **Styling**: Scoped CSS in `<style>` blocks
- **I18N**: Use `t(key, params, lang)` function

### Messaging
Messages between components use typed events. See [src/types/messaging.ts](../src/types/messaging.ts) for type definitions.

### Linting
Run `npm run check` to verify TypeScript types before committing.

## Common Development Tasks

### Adding a Translation

1. Add key to [src/i18n/locales/en.json](../src/i18n/locales/en.json)
2. Translate to all 14 languages
3. Use in code: `t('new.key', {}, currentLang())`

**Batch update example**:
```bash
# Add new key to all locales
for file in src/i18n/locales/*.json; do
  jq '. + {"new.key": "Translation"}' "$file" > tmp && mv tmp "$file"
done
```

### Adding a New Export Format

1. Update `OptionFormat` type in [src/types/shared.ts](../src/types/shared.ts)
2. Add builder function in [src/background/builders.ts](../src/background/builders.ts)
3. Update download handler in [src/background/download.ts](../src/background/download.ts)
4. Add UI option in [src/entrypoints/popup/components/FormatSection.svelte](../src/entrypoints/popup/components/FormatSection.svelte)
5. Add translation keys for new format

### Adding a New UI Component

1. Create `src/entrypoints/popup/components/NewSection.svelte`
2. Import in [App.svelte](../src/entrypoints/popup/App.svelte)
3. Export props with defaults: `export let data: Item[] = []`
4. Use typed event dispatcher if needed: `createEventDispatcher<{ change: ItemType }>()`
5. Add scoped styles in `<style>` block

### Modifying Scraping Logic

Content script: [src/entrypoints/content.ts](../src/entrypoints/content.ts)
- **DOM selectors**: Use `data-tid` attributes when available
- **Scroll logic**: [src/content/scroll.ts](../src/content/scroll.ts)
- **Reactions**: [src/content/reactions.ts](../src/content/reactions.ts)
- **Replies**: [src/content/replies.ts](../src/content/replies.ts)
- **Attachments**: [src/content/attachments.ts](../src/content/attachments.ts)

**Note**: Teams DOM structure changes frequently. Test thoroughly after updates.

## Submitting Changes

1. Fork the repository.
2. Create a feature branch: `git checkout -b feature/my-feature`
3. Make your changes.
4. Test in both Chrome and Firefox (see [DEVELOPMENT.md](DEVELOPMENT.md) for testing checklist).
5. Run `npm run check` to verify types.
6. Ensure builds work for both browsers (see [DEVELOPMENT.md](DEVELOPMENT.md) for build commands).
7. Commit with descriptive message: `feat: add export cancellation`
8. Submit a Pull Request with:
   - Description of the change
   - Testing performed
   - Screenshots (if UI changes)

## CI (Continuous Integration)

Every pull request is automatically checked by GitHub Actions before it can be merged. The workflow runs two checks:

1. **`npm run check`** — Runs `svelte-check` to catch TypeScript type errors.
2. **`npm run build`** — Builds the Chrome extension to catch build failures.

If either check fails, the PR will show a red X and cannot be merged until the issue is fixed.

**What this means for contributors:**
- Run `npm run check` and `npm run build` locally before pushing. If they pass on your machine, they will pass in CI.
- If CI fails on your PR, click the red X on GitHub to see the error log, fix the issue, and push again.

The workflow file lives at [.github/workflows/ci.yml](../.github/workflows/ci.yml).
