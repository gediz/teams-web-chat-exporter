# CLAUDE.md - Teams Chat Exporter Codebase Guide

This document provides AI assistants with a comprehensive understanding of the Teams Chat Exporter codebase, its architecture, conventions, and development workflows.

## Table of Contents
- [Project Overview](#project-overview)
- [Architecture](#architecture)
- [File Structure](#file-structure)
- [Key Concepts](#key-concepts)
- [Data Flow](#data-flow)
- [Code Conventions](#code-conventions)
- [Development Workflow](#development-workflow)
- [Testing Strategy](#testing-strategy)
- [Common Tasks](#common-tasks)
- [Known Issues & Future Work](#known-issues--future-work)

---

## Project Overview

**Teams Chat Exporter** is a Chrome Manifest V3 extension that exports Microsoft Teams web chat conversations to JSON, CSV, or HTML formats. The extension:

- Scrapes chat messages from the Teams web UI by scrolling and parsing DOM elements
- Supports date range filtering, threaded replies, reactions, and system messages
- Embeds avatars as base64 in HTML exports
- Uses a service worker for background processing and file generation
- Provides a popup UI for configuration and progress monitoring

**Current Version**: 1.0.1
**Codebase Size**: ~2,275 lines of JavaScript
**Browser Support**: Chrome (Firefox port planned but blocked on refactors)

---

## Architecture

The extension follows Chrome Manifest V3 architecture with three main components:

```
┌─────────────────┐
│   popup.html    │  ← User Interface
│   popup.js      │  ← Configuration & Status Display
└────────┬────────┘
         │ chrome.runtime.sendMessage
         ↓
┌─────────────────┐
│ service-worker  │  ← Background Orchestration
│   .js           │  ← Export Building & Downloads
└────────┬────────┘
         │ chrome.tabs.sendMessage
         ↓
┌─────────────────┐
│  content.js     │  ← DOM Scraping & Scrolling
│                 │  ← Injected into Teams tabs
└─────────────────┘
```

### Component Responsibilities

#### 1. **popup.js** (~588 lines)
- Renders extension UI with date pickers, format selectors, and toggles
- Persists user preferences to `chrome.storage.local`
- Initiates export by sending `START_EXPORT` message to service worker
- Displays real-time progress updates and elapsed time
- Implements theme toggle (light/dark mode)

**Key Functions**:
- `gatherOptions()`: Collects and validates user input
- `setStatus()`: Updates status text with elapsed time
- `handleExportStatus()`: Processes status messages from service worker

#### 2. **service-worker.js** (~651 lines)
- Orchestrates export workflow by messaging content script
- Builds export files (JSON/CSV/HTML) from collected message data
- Manages Chrome downloads API
- Updates extension badge with message counts
- Handles avatar embedding for HTML exports

**Key Functions**:
- `runExport()`: Main export orchestration (service-worker.js:145)
- `buildJSON()`, `buildCSV()`, `buildHTML()`: Format-specific builders
- `embedAvatarsInRows()`: Fetches and converts avatars to base64 (service-worker.js:93)
- `formatBadgeCount()`: Humanizes badge numbers (e.g., "1.2k") (service-worker.js:37)

#### 3. **content.js** (~1,036 lines)
- Injected into all Teams tabs via `content_scripts` in manifest
- Scrapes DOM to extract messages, authors, timestamps, reactions, replies
- Auto-scrolls chat pane to load older messages
- Uses `IntersectionObserver` to detect when top of chat is reached
- Provides optional in-page HUD overlay for progress feedback

**Key Functions**:
- `checkChatContext()`: Validates Teams chat is open (content.js:28)
- `startScrape()`: Main scraping orchestration (content.js:~700)
- `scrollToLoadMore()`: Auto-scroll implementation with sentinel detection
- `resolveAuthor()`, `resolveTimestamp()`: DOM parsing helpers (content.js:91-100)
- `parseDateDividerText()`: Extracts dates from Teams "day divider" elements

---

## File Structure

```
teams-web-chat-exporter/
├── manifest.json          # Extension configuration (MV3)
├── popup.html             # UI markup with embedded CSS (~648 lines)
├── popup.js               # UI logic & state management
├── service-worker.js      # Background processing & export building
├── content.js             # DOM scraping & scroll automation
├── icons/                 # Extension icons (16-256px)
│   ├── action-16.png
│   ├── action-32.png
│   ├── action-48.png
│   ├── action-128.png
│   └── action-256.png
├── README.md              # User-facing documentation
├── TODO.md                # Tracked tasks & feature roadmap
├── REFACTOR_PLAN.md       # Proposed code organization improvements
├── DEV_IMPROVEMENTS.md    # Tooling wishlist (TypeScript, linting, tests)
└── CLAUDE.md              # This file
```

**No build step**: Source files are loaded directly. Future plans include Vite/bundler integration.

---

## Key Concepts

### Message Aggregation

Messages are collected into `AggregatedEntry` objects with this structure:

```javascript
{
  kind: "message" | "dayDivider",
  tsMs: number,              // Message timestamp in milliseconds
  anchorTs: number,          // Teams-internal anchor timestamp
  author: string,
  avatar: string,            // URL or base64 data URL
  body: string,              // Plain text or HTML
  edited: boolean,
  replyCount: number,
  replyToId: string | null,  // Parent message ID for threaded replies
  replyToAuthor: string | null,
  reactions: Array<{emoji, participants}>,
  attachments: Array<{name, type, size, preview}>,
  mentions: Array<{name, id}>
}
```

**Day Dividers**: Special entries with `kind: "dayDivider"` inserted to mark date boundaries in exports.

### Scroll Loop Strategy

The content script uses a sophisticated scroll loop to load chat history:

1. **Sentinel Observer**: Uses `IntersectionObserver` on the topmost message element to detect when true top is reached
2. **Height Stabilization**: Tracks `scrollHeight` changes to detect when Teams stops loading new messages
3. **Author Carry-over**: Preserves `lastAuthor` across scroll passes to handle Teams' UI quirk of omitting author names on consecutive messages
4. **Hard Timeout**: Exits after max passes or when oldest message ID stops changing

See `scrollToLoadMore()` in content.js for implementation details.

### Chrome Messaging

#### Message Types

**Popup → Service Worker**:
- `START_EXPORT`: Initiates export with options payload

**Service Worker → Content Script**:
- `DO_SCRAPE`: Starts scraping with date range filters

**Content Script → Service Worker** (via popup):
- `SCRAPE_PROGRESS`: Updates during scroll/extraction (`{phase, passes, seen}`)
- `SCRAPE_COMPLETE`: Returns collected messages
- `SCRAPE_ERROR`: Reports failures

**Service Worker → Popup**:
- `EXPORT_STATUS`: Progress updates (`{status, phase, count, error}`)

### Badge Management

The extension uses `chrome.action.setBadgeText` to show message counts:

- Updates in real-time during scraping
- Humanizes large numbers: `1234 → "1.2k"`, `1234567 → "1.2m"`
- Resets when tabs reload or service worker restarts
- Blue badge during export, gray for empty results

---

## Data Flow

### Export Workflow (Happy Path)

```
1. User opens popup, configures options, clicks "Export current chat"
   ↓
2. popup.js validates inputs, sends START_EXPORT to service worker
   ↓
3. service-worker.js sends DO_SCRAPE to active Teams tab
   ↓
4. content.js checks chat context (isChatNavSelected, hasChatMessageSurface)
   ↓
5. content.js scrolls to top, observes sentinel, collects messages
   ├─ Updates badge via SCRAPE_PROGRESS messages
   └─ Shows in-page HUD (if enabled)
   ↓
6. content.js returns SCRAPE_COMPLETE with aggregated messages
   ↓
7. service-worker.js builds export (JSON/CSV/HTML)
   ├─ Embeds avatars if requested (HTML only)
   └─ Generates filename with chat name + timestamp
   ↓
8. service-worker.js triggers download via chrome.downloads.download
   ↓
9. service-worker.js sends EXPORT_STATUS "complete" to popup
   ↓
10. popup.js shows success status, re-enables "Export" button
```

### Error Handling

- **Invalid context**: Content script returns error if not on chat view
- **No messages**: Service worker skips download, shows banner instead of file
- **Scroll timeout**: Content script exits gracefully with partial results
- **Avatar fetch failures**: Service worker logs error, uses placeholder or original URL

---

## Code Conventions

### Naming Patterns

- **Selectors**: Use descriptive names: `$` for single element, `$$` for arrays
- **Functions**: Verb-first imperative (e.g., `resolveAuthor`, `buildHTML`, `formatBadgeCount`)
- **Constants**: UPPER_SNAKE_CASE (e.g., `DAY_MS`, `TERMINAL_PHASES`)
- **Message types**: UPPER_SNAKE_CASE with namespace (e.g., `START_EXPORT`, `SCRAPE_COMPLETE`)

### DOM Querying

Teams uses `data-tid` attributes extensively. Selectors rely on these:

```javascript
// Good: Target specific elements by data-tid
const viewport = document.querySelector('[data-tid="message-pane-list-viewport"]');
const authorEl = document.querySelector('[data-tid="message-author-name"]');

// Fallback: Use aria-labelledby or classList when data-tid unavailable
```

### Timestamp Handling

Always work with **milliseconds since epoch** internally:

```javascript
function parseTimeStamp(value) {
    // Normalize Teams' inconsistent formats
    const ts = Date.parse(value);
    if (!Number.isNaN(ts)) return ts;

    // Fallback: treat as local time
    const normalized = value.replace(/ /g, 'T');
    return Date.parse(normalized);
}
```

### Async/Await

All Chrome APIs use promises; use `async/await` consistently:

```javascript
// Good
async function getActiveTeamsTab() {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    return tab;
}

// Avoid callback hell
```

### Error Logging

Wrap console calls to avoid exceptions in production:

```javascript
function log(...args) {
    try { console.log("[Teams Exporter]", ...args); } catch {}
}
```

---

## Development Workflow

### Local Development

1. **Load Extension**:
   ```bash
   # Open chrome://extensions/
   # Enable "Developer mode"
   # Click "Load unpacked" → select project root
   ```

2. **Make Changes**: Edit JS/HTML/CSS directly (no build step)

3. **Reload Extension**:
   - Click refresh icon on `chrome://extensions` card
   - Or use keyboard shortcut (configure via Extensions Reloader)

4. **Debug**:
   - **Popup**: Right-click extension icon → "Inspect popup"
   - **Service Worker**: Click "Inspect views: service worker" on extensions page
   - **Content Script**: Open Teams tab, press F12, check console for `[content.js]` logs

### Testing on Real Teams Chats

1. Open `teams.microsoft.com` in Chrome
2. Navigate to Chat app
3. Open a conversation with varied content (mentions, replies, reactions, attachments)
4. Click extension icon, configure options, export
5. Validate downloaded file format and content

### Making Commits

Follow these commit message conventions:

- **feat**: New feature (e.g., "feat: add PDF export format")
- **fix**: Bug fix (e.g., "fix: preserve author across scroll passes")
- **refactor**: Code restructuring without behavior change
- **docs**: Documentation updates
- **chore**: Tooling/config changes

**Example**:
```bash
git add content.js
git commit -m "fix: handle empty author name in Teams v2 UI"
```

### Updating Version

1. Edit `manifest.json` → bump `version` field
2. Update README.md if user-facing changes
3. Consider adding entry to CHANGELOG (if created)

---

## Testing Strategy

### Current State

**No automated tests exist yet.** Testing is manual via:

- Loading extension in Chrome dev mode
- Exporting real Teams chats
- Inspecting downloaded files

### Recommended Test Coverage (from DEV_IMPROVEMENTS.md)

#### Unit Tests (with Vitest/Jest)

Test pure utility functions:

```javascript
// Example: test/formatBadgeCount.test.js
import { formatBadgeCount } from '../service-worker.js';

test('formats thousands', () => {
    expect(formatBadgeCount(1234)).toBe('1.2k');
    expect(formatBadgeCount(10500)).toBe('11k');
});

test('formats millions', () => {
    expect(formatBadgeCount(1234567)).toBe('1.2m');
});
```

**Testable modules** (extract first):
- Date/timestamp parsing (`parseTimeStamp`, `parseDateDividerText`)
- Badge formatting (`formatBadgeCount`)
- Export builders (`buildJSON`, `buildCSV`, `buildHTML`)
- Filename sanitization (`sanitizeBase`)

#### Integration Tests (with Playwright/Puppeteer)

1. Launch Chrome with extension loaded (`--load-extension` flag)
2. Serve mock Teams HTML page (e.g., from `etc/` fixtures)
3. Open popup, configure options, start export
4. Assert on:
   - Badge updates
   - Downloaded file content
   - Error states

**Example scenario**:
```javascript
test('exports empty chat gracefully', async ({ page }) => {
    await page.goto('http://localhost:3000/empty-chat.html');
    await page.click('text=Export current chat');
    // Assert: banner shows "No messages found", no download triggered
});
```

#### Manual Test Checklist

Before releases, manually verify:

- [ ] Export JSON/CSV/HTML formats
- [ ] Date range filtering (start only, end only, both, neither)
- [ ] Empty chat → shows banner, no download
- [ ] Badge updates during export
- [ ] Badge resets on tab reload
- [ ] Light/dark theme toggle
- [ ] Avatar embedding (HTML only)
- [ ] Long chat (>1000 messages) completes successfully
- [ ] Teams v2 UI compatibility

---

## Common Tasks

### Adding a New Export Format

**Example**: Add Markdown export

1. **Update UI** (popup.html):
   ```html
   <select id="format">
       <option value="json">JSON</option>
       <option value="csv">CSV</option>
       <option value="html">HTML</option>
       <option value="markdown">Markdown</option> <!-- NEW -->
   </select>
   ```

2. **Implement Builder** (service-worker.js):
   ```javascript
   function buildMarkdown(rows, opts) {
       let md = `# ${opts.chatName || 'Teams Chat'}\n\n`;
       for (const m of rows) {
           if (m.kind === 'dayDivider') {
               md += `\n---\n**${m.label}**\n---\n\n`;
               continue;
           }
           const ts = new Date(m.tsMs).toLocaleString();
           md += `### ${m.author} (${ts})\n${m.body}\n\n`;
       }
       return md;
   }
   ```

3. **Wire Builder** (service-worker.js, in `runExport`):
   ```javascript
   let content, mime, ext;
   if (format === 'json') { /* ... */ }
   else if (format === 'csv') { /* ... */ }
   else if (format === 'html') { /* ... */ }
   else if (format === 'markdown') {
       content = buildMarkdown(rows, opts);
       mime = 'text/markdown; charset=utf-8';
       ext = 'md';
   }
   ```

4. **Test**: Export a chat, verify `.md` file renders correctly

### Adding a New Message Filter

**Example**: Filter by author name

1. **Add UI Control** (popup.html):
   ```html
   <input id="filterAuthor" type="text" placeholder="Author name" />
   ```

2. **Capture in Options** (popup.js, `gatherOptions`):
   ```javascript
   const opts = {
       // ... existing options
       filterAuthor: controls.filterAuthor.value.trim()
   };
   ```

3. **Apply Filter** (content.js, in aggregation loop):
   ```javascript
   if (opts.filterAuthor && !author.includes(opts.filterAuthor)) {
       continue; // Skip messages not matching author
   }
   ```

4. **Test**: Export with author filter, verify only matching messages included

### Changing DOM Selectors (Teams UI Updates)

When Teams updates their DOM structure:

1. **Identify Breaking Change**: Check console for selector errors
2. **Inspect Teams DOM**: Use DevTools to find new `data-tid` or aria attributes
3. **Update Selectors** (content.js):
   ```javascript
   // Old
   const viewport = $('[data-tid="old-viewport-id"]');

   // New
   const viewport = $('[data-tid="new-viewport-id"]') ||
                    $('[data-tid="old-viewport-id"]'); // Fallback for compatibility
   ```
4. **Test on Both Versions**: Verify extension works on old and new Teams UI

### Debugging Message Collection Issues

**Symptom**: Export is missing messages or has duplicates

**Steps**:

1. **Enable HUD**: Check "Show in-page progress HUD" in Advanced section
2. **Monitor Console**: Open DevTools on Teams tab, watch content script logs
3. **Check Scroll Loop**:
   ```javascript
   // In content.js, add debug logging
   log(`Scroll pass ${passCount}: height=${scroller.scrollHeight}, seen=${aggregated.length}`);
   ```
4. **Inspect Aggregation**:
   ```javascript
   // Before sending SCRAPE_COMPLETE, log duplicates
   const ids = aggregated.map(m => m.anchorTs);
   const dupes = ids.filter((id, i) => ids.indexOf(id) !== i);
   if (dupes.length) log('WARNING: Duplicate message IDs:', dupes);
   ```
5. **Adjust Scroll Strategy**: If messages are missed, increase `maxScrollPasses` or adjust `heightStableCount` threshold

---

## Known Issues & Future Work

### Active TODOs (from TODO.md)

**Critical**:
- [ ] Trim `host_permissions` to minimal required scope (manifest.json:28-32)
- [ ] Add `homepage_url` and `support_url` to manifest

**UX Enhancements**:
- [ ] Move export button to top of popup for faster access
- [ ] Add progress bar (currently text-only status updates)

**Future Features** (see TODO.md):
- [ ] Incremental exports (persist last timestamp, diff collection)
- [ ] Participant filtering (export only messages from specific users)
- [ ] Export summary/analytics report
- [ ] PDF export option
- [ ] Batch export (all chats)

**Browser Support**:
- [ ] Firefox port (blockers: MV3 differences, service worker APIs)

### Refactoring Opportunities (from REFACTOR_PLAN.md)

**Module Extraction**:
- Extract date/timestamp helpers into `date-helpers.js`
- Split DOM/scroll logic from aggregation in content.js
- Create separate modules for export builders (json-builder.js, csv-builder.js, etc.)
- Isolate popup UI helpers (timer, banner, storage persistence)

**Type Safety**:
- Add JSDoc annotations for `AggregatedEntry` and message types
- Consider gradual TypeScript migration (`.ts` + `tsconfig.json` + bundler)

**Testing Tooling** (from DEV_IMPROVEMENTS.md):
- Set up ESLint + Prettier for code consistency
- Add Vitest for unit tests on pure functions
- Configure Playwright for E2E tests with mock Teams pages
- Set up GitHub Actions CI pipeline

**Build Pipeline**:
- Introduce Vite/Rollup for bundling and minification
- Add dev server with hot reload
- Create `dist/` folder for production builds
- Automate version bumping and release zipping

### Performance Considerations

**Current Bottlenecks**:
- **Scroll Loop**: Can take several minutes on very long chats (>5000 messages)
  - Mitigation: Show progress HUD, update badge frequently
- **Avatar Embedding**: Fetches each unique avatar sequentially
  - Future: Parallelize with `Promise.all` (be mindful of rate limits)
- **HTML Export Size**: Large chats with embedded avatars can exceed 50MB
  - Mitigation: Warn user when `embedAvatars` is enabled

**Optimization Ideas**:
- Use `requestIdleCallback` for non-urgent DOM parsing
- Implement virtual scrolling for huge exports (reduce memory footprint)
- Add option to export in chunks (e.g., 1000 messages per file)

---

## Working with This Codebase

### For AI Assistants

When helping users with this codebase:

1. **Understand Context**: Always check which component (popup/service-worker/content) needs changes
2. **Preserve Patterns**: Follow existing naming conventions and message types
3. **Test Holistically**: Changes to content.js often require updates to service-worker.js message handlers
4. **Maintain Backward Compat**: Teams UI varies across regions/versions; preserve fallback selectors
5. **Document Decisions**: Add comments for non-obvious DOM selectors or timing hacks

### Quick Reference

**Find where...**:
- Export is triggered: `popup.js` → `gatherOptions()` + `START_EXPORT` message
- Messages are collected: `content.js` → `startScrape()` → scroll loop → aggregation
- Files are built: `service-worker.js` → `runExport()` → format-specific builders
- Badge is updated: `service-worker.js` → `updateActiveExport()` + `chrome.action.setBadgeText`
- Timestamps are parsed: `content.js` → `parseTimeStamp()`, `parseDateDividerText()`
- Options are persisted: `popup.js` → `chrome.storage.local.set(STORAGE_KEY, opts)`

**Common Gotchas**:
- Content script runs in isolated world; can't directly access page JavaScript
- Service worker may restart; use `chrome.storage` for persistent state
- Teams loads messages lazily; must scroll to trigger fetch
- Author names sometimes missing (Teams quirk); carry over from previous message
- Timestamps in Teams DOM are inconsistent; normalize with fallbacks

---

## Additional Resources

- **Chrome Extension Docs**: https://developer.chrome.com/docs/extensions/mv3/
- **Teams Web Reverse Engineering**: Inspect DOM via DevTools (no official API)
- **Project Files**:
  - README.md: User installation & usage guide
  - TODO.md: Active tasks & feature roadmap
  - REFACTOR_PLAN.md: Code organization proposals
  - DEV_IMPROVEMENTS.md: Tooling wishlist & testing guide

---

**Last Updated**: 2025-01-15
**Maintainer**: See README for contact info
