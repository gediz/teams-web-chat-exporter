# Internationalization TODO

Goal: add easy-to-contribute i18n for the popup UI (status/error/labels/quick ranges). HUD and export file contents stay English.

- [ ] Scaffolding
  - [ ] Add `src/i18n/i18n.ts` with `t(key, params?)`, `setLanguage(lang)`, fallback to `en`, and `dir/lang` update for RTL.
  - [ ] Add locale JSON files under `src/i18n/locales/` for: `en`, `zh-CN`, `pt-BR`, `nl`, `fr`, `de`, `it`, `ja`, `ko`, `ru`, `es`, `tr`, `ar`, `he`. Machine translations are fine initially.
- [ ] Popup wiring
  - [ ] Add language selector to options; store selected `lang` in options storage.
  - [ ] Apply selected `lang` on load (set `dir` for `ar`/`he`) and expose `t` for components.
  - [ ] Replace hardcoded strings in popup components (`App.svelte`, Header, QuickRange, Options, Advanced, Action sections) with `t(...)`.
- [ ] Keys/content
  - [ ] Define shared keys: labels (sections, fields, toggles), quick ranges, statuses (`preparing`, `running`, `building`, `complete`, `error`, `empty`), errors (`invalidRange`, `startAfterEnd`, generic), buttons (`export`, `cancel`, `retry` if any), placeholders.
  - [ ] Keep HUD strings English; export payload text unchanged.
- [ ] RTL check
  - [ ] When `lang` is `ar`/`he`, set `dir="rtl"` on `<html>`/`body`; visually spot-check layout (chips/buttons) and adjust minimal CSS if needed (e.g., alignments).
- [ ] Docs
  - [ ] Add short CONTRIBUTING note on how to add a new locale JSON (copy `en`, translate values).

Open question: Keep a single `en` (no US/UK split) and a single `zh-CN` (add `zh-TW` later if contributed).
