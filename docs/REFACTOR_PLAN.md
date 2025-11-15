# Refactor Plan

## Context
The codebase is starting to accrete logic inside a few large files (`content.js`, `service-worker.js`, `popup.js`). We want to refactor while avoiding regressions.

## Suggested Direction
- **Module boundaries**
  - Extract date/timestamp helpers (e.g., `parseDateDividerText`, day divider generation) into a shared helper module.
  - Separate DOM/HUD/scroll loop helpers in `content.js` from the data aggregation logic.
  - Split export builders (JSON/CSV/HTML) and chrome badge logic into dedicated modules in the service worker.
  - For `popup.js`, isolate option persistence, elapsed timer, and banner helpers into smaller utility functions.

- **Typed structures / documentation**
  - Define explicit structures (e.g., `AggregatedEntry`) via JSDoc or TypeScript so fields like `kind`, `tsMs`, and `anchorTs` are documented.
  - Track run state (counts, start time, filters) in a small object to avoid passing loose parameter bags.

- **Chrome API wrappers**
  - Wrap `chrome.runtime.sendMessage`, `chrome.action.setBadgeText`, etc., in utility functions to decouple business logic from the API surface (easier testing/mocking).

- **Testing strategy**
  - Start with pure helpers (date parsing, day divider formatting, CSV builder) and add unit tests so we can refactor with confidence.
  - After helpers are tested, migrate larger flows in small steps, reusing the helper modules.

- **Future improvements**
  - Consider gradual TypeScript adoption once modules are split (or enhance JSDoc for IDE support).
  - Update lint/build tooling to enforce consistent imports after modules are created.

## Next Steps
1. Extract the date/timestamp/day-divider helpers into a `date-helpers.js` (or similar) and add unit tests.
2. Refactor `content.js` to consume these helpers and simplify the aggregation flow.
3. Split export builders and badge managers out of `service-worker.js`.
4. Clean up `popup.js` by moving persistence and timer logic into small utilities.
5. Evaluate TypeScript / lint upgrades once modules and tests are in place.
