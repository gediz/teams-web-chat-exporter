# Development Improvements & Testing Guide

This note collects suggested upgrades for the project’s tooling and a primer on how to test the extension effectively.

## Roadmap Ideas

- **Linting & Formatting**
  - Add ESLint (with the Chrome Extensions plugin) plus Prettier to keep popup, service worker, and content script consistent.
  - Wire a `npm run lint` script and optionally a pre-commit hook (Husky) so checks run automatically.

- **TypeScript Adoption**
  - Introduce TypeScript incrementally by creating `tsconfig.json`, renaming one surface at a time (e.g., service worker → `service-worker.ts`).
  - Leverage `@types/chrome` for accurate API signatures; this catches mis-typed messages and storage usage early.

- **Bundling & Dev Server**
  - Use Vite (with the MV3 plugin) or another bundler to compile TS, support modern syntax, and minify production builds.
  - Add a lightweight dev server that watches files, builds to `dist/`, and triggers a Chrome extension reload script for faster iteration.

- **Automated Testing**
  - Unit-test pure helpers (timestamp parsing, HTML export rendering) with Vitest or Jest.
  - Add integration tests with Playwright/Puppeteer: spin up a mocked Teams page (reuse `etc/snippet-*.html`), inject the content script, and assert collected messages, badge updates, etc.
  - Run these checks in CI (GitHub Actions) so PRs get automatic feedback.

- **Release & Documentation**
  - Create a CHANGELOG, plus scripts to bump the manifest version and zip the `dist/` folder.
  - Document manual QA scenarios (date ranges, empty export, badge resets) to keep release testing consistent.

- **Developer Ergonomics**
  - Expand the snippets into a local HTML gallery served via Vite/Express so you can load pages that mimic Teams quickly.
  - Provide a CLI script that uses `chrome-cli`/`chromedriver` to open the extension popup for smoke testing after each build.

## Testing the Extension

### Manual Testing Workflow

1. **Build & Load**
   - Bundle (or use the raw source) and open `chrome://extensions` → toggle *Developer mode* → *Load unpacked*, pointing to the project root (or `dist/`).
2. **Pick Scenarios**
   - Real account: open a Teams web chat with varied content (mentions, attachments, replies).
   - Mock data: open one of the HTML snippets in `etc/` by dragging it into a Chrome tab; they reproduce specific edge cases.
3. **Run Exports**
   - Use the popup to run through different formats (JSON/CSV/HTML) and date ranges (start-only, end-only, empty range) and confirm downloads.
   - Observe the badge: ensure it updates during runs, clears after refresh, and shows the formatted counts.
4. **Validate Output**
   - Check JSON/CSV payloads in a viewer; in HTML exports verify mentions, replies, and reactions render cleanly.
5. **Monitor Logs**
   - Open *Service Worker* and *Content Script* DevTools (chrome://extensions → *Inspect views*) to catch errors or warnings.

### Automated Testing Options

- **Content Script Unit Tests**
  - Extract utilities (timestamp parsing, mention normalization) into modules that can be imported by Vitest/Jest.
  - Use DOM testing libraries (JSDOM) to simulate minimal DOM fragments and assert transformed output.

- **Integration via Playwright or Puppeteer**
  - Launch a headless Chrome instance with the extension loaded (`--load-extension` flag).
  - Serve mock Teams pages (from `etc/` or generated fixtures) via a simple HTTP server.
  - Drive the popup UI: fill date ranges, start exports, wait for badge/notifications, and read the generated downloads (e.g., intercept via `chrome.downloads` API in the background page).

- **CI Setup**
  - GitHub Actions job: install dependencies, run lint/tests, and optionally execute headless integration tests using Playwright’s container image.
  - Upload artifacts (e.g., exported HTML snapshots) when tests fail for easier debugging.

### Getting Started Without Prior Testing Experience

- Start small: write one unit test for `formatBadgeCount` or mention normalization using Vitest. Running `npx vitest` locally gives fast feedback.
- Record a Playwright script in headed mode to interact with the popup—you can convert it into a regression test later.
- Keep a checklist of manual scenarios (stored in the repo) and run through it before each release until automated coverage is in place.

These additions can be adopted gradually—pick the areas that address your biggest pain points first (e.g., TypeScript + linting, then scripted integration tests).

