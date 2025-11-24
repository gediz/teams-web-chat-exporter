# Svelte Migration TODO

Incremental plan to port to Svelte with full TypeScript. Start with the popup, then migrate background/content scripts.

- [x] Add deps: `npm i -D svelte @sveltejs/vite-plugin-svelte svelte-check @tsconfig/svelte` (keep `@types/chrome`).
- [x] Config wiring (`wxt.config.ts`):
  - [x] Import `svelte` plugin and set `vite: { plugins: [svelte()] }` (keep `srcDir`, runner/dev settings).
- [x] TS/Svelte scaffolding:
  - [x] Add `tsconfig.json` extending `@tsconfig/svelte/tsconfig.json` (include `src` + `wxt.config.ts`; enable strict mode).
  - [x] Add `svelte.config.js` (or `.ts`) with default export for the plugin (minimal is fine for now).
  - [x] Add `svelte.d.ts` for tooling if needed.
  - [x] Update `package.json` scripts: add `check` (`svelte-check`), optional `lint` placeholder.
- [x] Prepare popup entrypoint shell:
  - [x] Update `src/entrypoints/popup/index.html` to host `<div id="app"></div>` and a module script that imports `./main.ts`.
  - [x] Add `src/entrypoints/popup/main.ts` that mounts the Svelte app (TS).
- [x] Build the Svelte UI:
  - [x] Create `src/entrypoints/popup/App.svelte` with existing layout.
  - [x] Move option state (dates, toggles, format) into Svelte component state/stores (TS types for options/messages).
  - [x] Port runtime/tabs/storage messaging; handle listeners with `onMount`/cleanup.
  - [x] Reuse validation/range helpers from current JS (or translate to TS as needed).
  - [x] Wire banners/status/busy states and quick-range buttons.
- [x] Styles (no inline styles):
  - [x] Create a popup stylesheet (e.g., `src/entrypoints/popup/popup.css`) and import it in `App.svelte` or `main.ts`.
  - [x] Keep component-specific styling scoped via Svelte `<style>` blocks only when necessary; prefer shared CSS modules for reusable classes.
- [ ] Testing passes:
  - [ ] `npm run dev` and `npm run dev:firefox` for manual reload/hot reload.
  - [x] `npm run build` and `npm run build:firefox`; load unpacked in Chrome/Firefox, verify persistence/export.
- [x] Docs:
  - [x] Update `README.md`/`QUICKSTART.md`/`DEPLOYMENT_GUIDE.md`/`FIXES.md` to reflect Svelte popup and new scripts.
  - [x] Note mixed stack (popup in Svelte, background/content still JS) in docs/notes.
- [ ] Optional polish:
  - [x] Add `npm run check` using `svelte-check` and `@tsconfig/svelte`.
  - [ ] Migrate background/content scripts to TypeScript and share utilities as a follow-up.
  - [ ] Factor popup into smaller Svelte components (header, range, options, advanced, footer) for readability.
  - [ ] Investigate HUD/elapsed timer occasional stalls; ensure timers are cleaned up and status updates stay in sync.
