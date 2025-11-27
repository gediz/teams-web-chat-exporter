# CLAUDE.md - AI Assistant Guide

This file provides AI assistants with context about the Teams Chat Exporter codebase.

## Project Overview

**Teams Chat Exporter** is a cross-browser extension (Chrome, Edge, Firefox) that exports Microsoft Teams web chat conversations to JSON, CSV, HTML, or text formats.

- **Version**: 2.0.0
- **Framework**: WXT (Vite-based)
- **Languages**: TypeScript + Svelte 5
- **Supported Browsers**: Chrome, Edge, Firefox
- **UI Languages**: 14 (en, zh-CN, pt-BR, nl, fr, de, it, ja, ko, ru, es, tr, ar, he)

## Quick Reference

### Key Files
- **Popup**: [src/entrypoints/popup/App.svelte](src/entrypoints/popup/App.svelte)
- **Background**: [src/entrypoints/background.ts](src/entrypoints/background.ts)
- **Content**: [src/entrypoints/content.ts](src/entrypoints/content.ts)
- **Config**: [wxt.config.ts](wxt.config.ts)

### Build Commands
```bash
npm run dev              # Chrome development
npm run dev:firefox      # Firefox development
npm run build            # Chrome production
npm run build:firefox    # Firefox production
```

### Testing
See [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md) for testing checklist.

## Architecture

See [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) for:
- Technology stack
- Component communication diagram
- Message flow
- Data structures
- Browser compatibility details

## Code Conventions

### File Organization
- **Entrypoints**: `src/entrypoints/` (popup, background, content)
- **Components**: `src/entrypoints/popup/components/`
- **Utils**: `src/utils/` (shared code)
- **I18N**: `src/i18n/locales/` (translation JSON files)
- **Types**: `src/types/` (TypeScript interfaces)

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

### Svelte Conventions
- **Svelte 5 Runes**: Use `$state`, `$derived`, `$effect` (modern reactive syntax)
- **Props**: Destructure with defaults: `let { theme = 'light' } = $props()`
- **Events**: Use `createEventDispatcher<T>()` with typed events
- **Styling**: Scoped CSS in `<style>` blocks
- **I18N**: Use `t(key, params, lang)` function

### Messaging
Messages between components use typed events:

```typescript
// Popup → Background
chrome.runtime.sendMessage({ type: 'START_EXPORT', options })

// Background → Content
chrome.tabs.sendMessage(tabId, { type: 'SCRAPE_TEAMS', options })

// Content → Background (via response)
sendResponse({ messages, meta })

// Background → Popup
chrome.runtime.sendMessage({ type: 'EXPORT_STATUS', status, phase, count })
```

See [src/types/messaging.ts](src/types/messaging.ts) for type definitions.

## Common Tasks

### Adding a Translation

1. Add key to [src/i18n/locales/en.json](src/i18n/locales/en.json)
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

1. Update `OptionFormat` type in [src/types/shared.ts](src/types/shared.ts)
2. Add builder function in [src/background/builders.ts](src/background/builders.ts)
3. Update download handler in [src/background/download.ts](src/background/download.ts)
4. Add UI option in [src/entrypoints/popup/components/FormatSection.svelte](src/entrypoints/popup/components/FormatSection.svelte)
5. Add translation keys for new format

### Adding a New UI Component

1. Create `src/entrypoints/popup/components/NewSection.svelte`
2. Import in `App.svelte`
3. Add props with defaults: `let { data = [] } = $props()`
4. Use typed event dispatcher if needed
5. Add scoped styles in `<style>` block

### Modifying Scraping Logic

Content script: [src/entrypoints/content.ts](src/entrypoints/content.ts)
- **DOM selectors**: Use `data-tid` attributes when available
- **Scroll logic**: [src/content/scroll.ts](src/content/scroll.ts)
- **Reactions**: [src/content/reactions.ts](src/content/reactions.ts)
- **Replies**: [src/content/replies.ts](src/content/replies.ts)
- **Attachments**: [src/content/attachments.ts](src/content/attachments.ts)

**Note**: Teams DOM structure changes frequently. Test thoroughly after updates.

## Browser Compatibility

### Chrome/Edge
- Manifest V3
- `chrome.*` namespace
- Service worker background
- Data URL downloads

### Firefox
- Manifest V2 (auto-converted by WXT)
- `browser.*` namespace (polyfilled)
- Background page (not service worker)
- Blob URL downloads (data URLs blocked)

WXT handles these differences automatically via [wxt.config.ts](wxt.config.ts).

## Known Issues

1. **Content Script Injection**: Manual fallback relies on WXT filename (`content.js`). May break if build config changes.
2. **Data URL Limits**: Large HTML exports (>50 MB) may fail in Chrome. Firefox uses Blob URLs (safer).
3. **HUD Timer Stalls**: Occasional UI freezes during very long exports (investigating).

See [TODO.md](TODO.md) for roadmap.

## Testing

Run through [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md) testing checklist before releases.

**Critical tests**:
- Export all formats (JSON, CSV, HTML, Text)
- Date filtering (presets and custom ranges)
- Cross-browser (Chrome, Firefox)
- Large chats (1000+ messages)
- Empty chats (should show banner, no download)
- Badge updates during export
- Theme toggle (light/dark)
- Language switching (14 languages)

## Documentation

- **README.md**: User guide
- **docs/ARCHITECTURE.md**: Technical design
- **docs/DEVELOPMENT.md**: Build/test instructions
- **docs/CONTRIBUTING.md**: Contribution guidelines
- **docs/DEPLOYMENT.md**: Store publishing
- **docs/MANUAL_INSTALL.md**: Installation steps
- **docs/MIGRATION_NOTES.md**: WXT migration history

## When Helping Users

### For Code Questions
1. Check [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) for high-level design
2. Reference specific files with line numbers (e.g., `App.svelte:42`)
3. Use TypeScript types from [src/types/](src/types/)
4. Follow Svelte 5 runes syntax (not legacy stores)

### For Build Issues
1. Check [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md)
2. Common fix: `rm -rf .output .wxt node_modules && npm install`
3. Firefox needs separate build: `npm run build:firefox`

### For Feature Requests
1. Check [TODO.md](TODO.md) for planned features
2. Consider cross-browser compatibility
3. Verify Teams DOM selectors won't break
4. Add i18n keys for new UI strings

## Migration from Vanilla JS

This project was migrated from vanilla JavaScript to WXT in November 2025. The old codebase is preserved in git history (commit `df796ad`). Key improvements:

- **Cross-browser**: Chrome → Chrome + Edge + Firefox
- **Type safety**: JavaScript → TypeScript
- **Modern UI**: Vanilla JS → Svelte 5
- **Build system**: None → Vite (WXT)
- **I18N**: English only → 14 languages

See [docs/MIGRATION_NOTES.md](docs/MIGRATION_NOTES.md) for details.

---

**Last Updated**: 2025-11-27
**Version**: 2.0.0
