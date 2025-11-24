# Svelte/TypeScript Migration Status

## Completed
- Popup migrated to Svelte + TypeScript (`App.svelte`, `main.ts`, `popup.css`).
- Shared type definitions added in `src/types/shared.ts` (messages, meta, reactions, attachments, scrape/build options, status payloads, order context).
- Background entrypoint typed using shared types.
- Content script fully typed (removed `@ts-nocheck`), added typed helpers for DOM text extraction, attachments, reactions, reply parsing, and scroll aggregation.
- TS config set to strict with `allowJs: false`; builds and `npm run check` pass.
- Docs updated to reflect Svelte/TS popup and TS entrypoints.

## Remaining / Follow-ups
- Refactor popup into smaller Svelte components for readability (header, range, options, advanced done; footer/status/run button still inline).
- Consider extracting shared UI/logic utilities (date/range helpers, messaging payloads) if duplication grows.
- Optional tooling: linting/formatting setup, add tests (unit for helpers, E2E for export flow).
- Optional: further module extraction in content script (attachments/reactions/parse helpers) if desired.
- HUD/elapsed timer occasional stall is noted for later investigation.
