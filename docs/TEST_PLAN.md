# Test Plan

Greenfield. There are no automated tests today (`pnpm check` is the only thing). This doc lays out a layered plan, the tooling, what to build first, and what stays manual.

## Constraints

What we can test deterministically (no Teams auth, no network):
- Pure functions: text extraction, filename sanitiser, time/format helpers, URL matchers, options normalisation.
- Builders (`toJSON` / `toCSV` / `toHTML` / `toTXT` / `toPDF`) given fixture inputs.
- The API converter (`src/content/api-converter.ts`) given fixture chat-service blobs.
- Svelte components in isolation against mocked stores/props.
- The popup as a static page with fixture data injected via a `?demo=1` flag.

What we cannot automate without giving up reliability:
- Live Teams scraping (auth + page-state moving target).
- IDB reads against real Teams data.
- Network-dependent flows (Graph user resolution, image fetching, AMS proxy).

In the middle:
- Content-script DOM scraping. Could be tested with recorded Teams DOM fixtures, but Teams' DOM ages fast. Worth doing for the small set of stable extractor functions, not for end-to-end DOM scrape flows.

## Layered architecture

```
Layer 6  Visual regression       (Playwright + screenshot diff)
Layer 5  Extension smoke         (Playwright with --load-extension)
Layer 4  Demo-mode + screenshots (Popup as static page, Playwright)
Layer 3  Component tests         (Svelte + @testing-library/svelte + Vitest browser)
Layer 2  Unit tests              (Vitest, pure functions + builders + converter)
Layer 1  Static hygiene          (i18n parity, manifest sanity, bundle budget)
Layer 0  Type check + builds     (already in place: pnpm check, pnpm build*)
```

Each layer adds confidence without making lower layers redundant. Layers 1 and 2 are the highest leverage and should ship first.

## Layer 1 — Static hygiene

Three checks that prevent the bugs we kept hitting manually:

1. **i18n parity.** Fail if any of the 24 locale files is missing a key from `en.json`, or carries an orphan key no source file references. We've been re-running a Python audit for this; a real test eliminates the next slip. Implementation: scan `src/` for every `t('key.path', ...)` call, build the master key set from `en.json`, diff against each locale.

2. **Manifest sanity.** Fail if `host_permissions` doesn't include all `TEAMS_MATCH_PATTERNS`, or if `permissions: ['scripting']` is missing, or if `content_scripts[0].matches` drifts from `TEAMS_MATCH_PATTERNS`. Catches refactors that drop a URL.

3. **Bundle size budget.** Fail if `.output/chrome-mv3` total goes over a threshold (current baseline 22.76 MB). Forces a conscious "yes, this dependency is worth it" moment when the bundle grows.

Run via `pnpm test:hygiene`. Fast (under 1 s).

## Layer 2 — Unit tests

**Tool:** [Vitest](https://vitest.dev/). WXT-native, ESM-first, fast, official Svelte plugin.

**Targets, in priority order:**

| Module | Why |
|---|---|
| `src/builders/*` | Hot path. Every export goes through these. Snapshot-test outputs given fixture `ExportMessage[]`. PDF gets page-count + metadata + first-page rasterisation hash. |
| `src/content/api-converter.ts` | Hot path on v1.4. Catches subtle regressions in message classification (system events, replies, mentions) when Teams ships a new field. |
| `src/utils/options.ts` | `normalizeOptions` is migration code. Test every legacy shape we still try to handle. |
| `src/utils/text.ts` | Mention normalisation, alt-text cleaning, table extraction. Lots of small invariants that quietly break. |
| `src/utils/messages.ts`, `src/utils/avatars.ts`, `src/utils/teams-urls.ts` | Pure helpers, easy wins. |
| `src/content/text.ts`, `src/content/attachments.ts`, `src/content/reactions.ts`, `src/content/replies.ts` | Each takes a small DOM/JSON snippet and returns a structured value. Ideal for fixture-driven tests. |

**Fixtures live in** `src/__fixtures__/`:
- `conversation.minimal.json` — one user, three messages, one reaction, one reply.
- `conversation.system_events.json` — name change, added/removed members, joining-enabled-update.
- `conversation.attachments.json` — file, image, link preview, audio.
- `conversation.threading.json` — parent + N replies, replies-out-of-order.
- `api_blob.thread_v2.json` — raw chat-service response for a `@thread.v2`.
- `api_blob.oneToOne.json` — raw response for a 1:1.
- `api_blob.channel.json` — raw response for a team channel.

Coverage target: 70 % on `src/builders/`, `src/utils/`, `src/content/api-converter.ts`. Other modules opportunistic.

Run via `pnpm test:unit`. Should stay under 3 s.

## Layer 3 — Component tests

**Tool:** `@testing-library/svelte` + Vitest browser mode (real Chromium, not jsdom — Svelte 5 plus `vanilla-calendar-pro` need a real DOM).

**Targets:**

1. **`ConversationPicker`** — every state transition (idle → loading → ok → error → empty → stillLoading), multi-select toggle, folder rail filtering, kind tab switching, bulk-select pill dropdown, chevron collapse + auto-scroll-into-view of active filter, "Loaded N" header.
2. **`OnboardingOverlay`** — prompt → tour transition, autoStart bypass, picker collapse coordination, `data-tour` target resolution, scrim hole positioning, scroll/resize rAF throttling.
3. **`FormatSection`** — toggle logic, "always at least one format selected" invariant, `bundle.zip` / `HTML.zip` pill conditions including `embedAvatars + avatarMode='files'`.
4. **`HistoryPage`** — entry kinds (success / failed / cancelled / no-history), reopen / show-folder events, the "new entry" indicator clearing.
5. **`SettingsPage`** — option dispatch shapes, replay tour wire-up, language change side-effect.
6. **`DateRangeSection`** — quick-range presets, calendar interaction, manual input parsing.

Run via `pnpm test:components`. Order of 5–10 s.

## Layer 4 — Demo-mode popup + screenshots

**Tool:** Playwright.

Add a `?demo=1` URL parameter to the popup that bypasses normal data loading and seeds fixture conversations / history / settings into a fake `chrome.storage.local`. A Playwright script then opens the built popup HTML in headless Chromium, drives it through states, and saves PNGs.

**Two outputs from one investment:**
- **Store-listing screenshots** — same fixture data, art-directed states, deterministic.
- **Visual regression coverage** — baseline PNGs committed under `screenshots/baseline/`, the test fails when a captured screenshot diffs by more than N pixels from baseline.

The hero shot (browser + popup over Teams) and the bundle-result screenshot (rendered HTML in a tab) stay manual once per release. Everything else regenerates with `pnpm screenshots`.

**Implementation outline:**
- `src/demo/fixtures.ts` — fixture conversations / history / settings, exported from a single object.
- `src/entrypoints/popup/App.svelte` — early in init, check `new URLSearchParams(location.search).get('demo')`. If set, hand-load fixtures into the in-memory state and skip storage / message-passing entirely.
- `playwright/screenshots.spec.ts` — opens the popup at `?demo=1`, drives through scripted states, calls `page.screenshot({ path: 'screenshots/<name>.png' })`.
- `playwright/visual.spec.ts` — same script, but uses Playwright's `await expect(page).toHaveScreenshot()` against baselines.

Run via `pnpm screenshots` (regenerate) or `pnpm test:visual` (assert against baselines).

## Layer 5 — Extension smoke

Playwright launches Chromium with `--load-extension=.output/chrome-mv3` and `--disable-extensions-except=.output/chrome-mv3`. Tests:

1. Service worker boots and `runtime.onInstalled` fires.
2. Popup opens at `chrome-extension://<id>/popup/index.html` without console errors.
3. On a non-Teams URL, picker shows the "Open the Teams web app tab first" empty state.
4. Settings page navigates and back-buttons correctly.
5. Replay tour from Settings fires the overlay.

Doesn't touch Teams. Catches manifest issues, popup-init crashes (the white-screen class of bug), service-worker registration failures.

Run via `pnpm test:e2e`. Order of 15–30 s (most of which is extension load).

## Layer 6 — Visual regression

Built into Layer 4 above. `playwright/visual.spec.ts` uses `await expect(page).toHaveScreenshot()` to diff against committed baselines. PR diffs make UI regressions obvious.

Baseline storage choice: commit them under `screenshots/baseline/`. Pros: PR-reviewable. Cons: ~5–10 MB of PNGs in the repo. Acceptable given the repo isn't huge already.

## CI plumbing

**Workflow:** `.github/workflows/ci.yml`, runs on PR + push to main:

```yaml
- pnpm install --frozen-lockfile
- pnpm check               # type check
- pnpm test:hygiene        # i18n + manifest + bundle
- pnpm test:unit           # vitest
- pnpm build               # chrome
- pnpm build:firefox       # firefox
- pnpm test:components     # svelte components in chromium
- pnpm test:e2e            # extension load + smoke
- pnpm test:visual         # screenshot diff
```

Cache: `~/.local/share/pnpm`, `node_modules`, Playwright browsers (`~/.cache/ms-playwright`).

## What to build first

If three afternoons of work were budgeted:

**Day 1.** Layer 1 (i18n parity + manifest sanity + bundle budget) + Vitest scaffolding + 4 builder snapshot tests. ~250 lines of test code, immediately catches regressions in our most-touched paths.

**Day 2.** Demo-mode plumbing in the popup (`?demo=1` + fixture loader) + Playwright screenshot script. Solves both store-screenshot-automation and visual-regression in one stroke.

**Day 3.** Component tests for `ConversationPicker` + `OnboardingOverlay` (the two most complex / most-touched components).

The rest (api-converter, attachments, reactions, settings, history, smoke) can come incrementally: every PR adds tests for the surface area it touches.

## Decisions to lock down before starting

1. **Vitest** as the runner (recommendation: yes).
2. **Fixture strategy** — synthetic only (handcrafted JSON), or also recorded-then-anonymised real Teams API responses? Recorded ones catch realistic edge cases; synthetic ones are smaller and carry no PII risk. Recommendation: synthetic for happy paths, anonymised recordings for known-quirky shapes (system events, federated 1:1s, threadProperties with hidden flag).
3. **Visual regression baselines** — commit to git, or store as build artifacts? Recommendation: commit. ~5 MB acceptable, makes PR review trivial.
4. **CI** — GitHub Actions now, or stay local-only? Recommendation: local first (Day 1), wire CI when there's something worth gating PRs on.
5. **Coverage goal** — 70 % on `src/builders/`, `src/utils/`, `src/content/api-converter.ts`; opportunistic elsewhere.

## What stays manual

- Hero screenshot (browser + popup over Teams) — once per release, art-directed.
- Rendered HTML / PDF export quality review — eye check on a real chat once per release.
- New-Teams-version sanity scrape — periodic, when Microsoft ships a UI change.

These are the cases where automation would either be fragile or miss the point of the check.
