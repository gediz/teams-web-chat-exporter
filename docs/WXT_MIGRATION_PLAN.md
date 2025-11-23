# WXT Migration Plan

**Version**: 1.0
**Date**: 2025-01-23
**Objective**: Migrate Teams Chat Exporter from vanilla Chrome extension to WXT framework for cross-browser support

---

## Table of Contents
- [Executive Summary](#executive-summary)
- [Why WXT?](#why-wxt)
- [Migration Phases](#migration-phases)
- [Detailed Steps](#detailed-steps)
- [File Structure Mapping](#file-structure-mapping)
- [Code Changes Required](#code-changes-required)
- [Testing Strategy](#testing-strategy)
- [Rollback Plan](#rollback-plan)
- [Timeline Estimate](#timeline-estimate)

---

## Executive Summary

**Current State**: Chrome-only Manifest V3 extension with ~2,278 lines of vanilla JavaScript
**Target State**: Cross-browser extension (Chrome, Firefox, Edge, Safari) built with WXT framework
**Risk Level**: Low-Medium (straightforward migration, well-supported patterns)
**Estimated Effort**: 8-12 hours for initial migration + 4-6 hours testing

### Key Benefits
1. ✅ **Instant Firefox support** - WXT auto-generates Firefox builds with browser API polyfills
2. ✅ **Better developer experience** - Hot module reload, TypeScript support, modern tooling
3. ✅ **Maintainability** - Modular code structure, type safety, tree-shaking
4. ✅ **Future-proof** - Abstracts browser API differences automatically
5. ✅ **Testing-ready** - Easier integration with Vitest, Playwright

### Key Challenges
1. ⚠️ Manual content script injection fallback needs adjustment
2. ⚠️ Build system adds slight complexity vs. direct file loading
3. ⚠️ Initial learning curve for WXT conventions

---

## Why WXT?

### Current Blockers for Firefox
Your codebase uses standard Web Extension APIs, but Firefox has subtle differences:
- `chrome.*` namespace → requires `browser.*` polyfill
- Service worker limitations → needs fallback handling
- Download API quirks → different blob handling

**WXT solves all of these automatically** via:
- Built-in `webextension-polyfill` integration
- Cross-browser manifest generation
- Browser-specific build outputs

### Alignment with Existing Plans
From `docs/DEV_IMPROVEMENTS.md`, WXT provides:
- ✅ TypeScript support (optional, gradual adoption)
- ✅ ESLint/Prettier integration
- ✅ Vite-based build system
- ✅ Test framework support (Vitest)
- ✅ Hot reload development

---

## Migration Phases

### Phase 1: Setup & Proof of Concept (2-3 hours)
- Install WXT and dependencies
- Create basic project structure
- Configure `wxt.config.ts`
- Verify build output

### Phase 2: Code Migration (4-6 hours)
- Convert popup (HTML + JS)
- Migrate service worker
- Port content script
- Update manifest configuration

### Phase 3: Testing & Validation (4-6 hours)
- Chrome functional testing
- Firefox compatibility testing
- Edge testing (optional)
- Performance comparison

### Phase 4: Documentation & Deployment (1-2 hours)
- Update README.md
- Update CLAUDE.md
- Create release builds
- Test store submission process

---

## Detailed Steps

### Step 1: Initialize WXT Project

```bash
# Create new WXT project in a separate directory (proof of concept)
npm create wxt@latest teams-chat-exporter-wxt

# Options to select:
# - Package Manager: npm
# - TypeScript: No (start with JS, migrate later)
# - Framework: None (vanilla)
```

### Step 2: Install Dependencies

```bash
cd teams-chat-exporter-wxt
npm install
```

No additional dependencies needed - WXT includes everything required.

### Step 3: Configure `wxt.config.ts`

```typescript
import { defineConfig } from 'wxt';

export default defineConfig({
  extensionApi: 'chrome',
  modules: ['@wxt-dev/module-react'],
  manifest: {
    name: 'Teams Chat Exporter',
    version: '1.0.1',
    description: 'Export Microsoft Teams web chat conversations to JSON, CSV, or HTML with full message history.',
    homepage_url: 'https://github.com/gediz/teams-web-chat-exporter',
    permissions: [
      'scripting',
      'activeTab',
      'downloads',
      'storage',
    ],
    host_permissions: [
      'https://*.teams.microsoft.com/*',
      'https://teams.cloud.microsoft/*',
    ],
    action: {
      default_title: 'Teams Chat Exporter',
    },
  },
  runner: {
    disabled: false,
    chromiumArgs: [],
  },
});
```

### Step 4: File Structure Migration

Create the following structure:

```
teams-chat-exporter-wxt/
├── entrypoints/
│   ├── background.ts          # service-worker.js → background.ts
│   ├── content.ts              # content.js → content.ts
│   └── popup/
│       ├── index.html          # popup.html → index.html
│       └── main.ts             # popup.js → main.ts
├── public/
│   └── icons/                  # Copy icons/ directory here
│       ├── action-16.png
│       ├── action-32.png
│       ├── action-48.png
│       ├── action-128.png
│       └── action-256.png
├── wxt.config.ts               # New: WXT configuration
├── package.json                # Auto-generated
└── tsconfig.json               # Auto-generated (even for JS projects)
```

### Step 5: Migrate Popup

#### `entrypoints/popup/index.html`
Copy `popup.html` **without** the `<script>` tag:

```html
<!doctype html>
<html>
<head>
    <meta charset="utf-8" />
    <style>
        /* Copy all styles from popup.html */
    </style>
</head>
<body data-theme="light">
    <!-- Copy all body content from popup.html -->
</body>
</html>
```

#### `entrypoints/popup/main.ts` (or `main.js`)
Copy `popup.js` content with these changes:

```javascript
// WXT auto-imports browser API - no changes needed for most code
// Just rename file to main.js or main.ts

// Original code works as-is, but you can optionally use browser.* instead of chrome.*
// Example:
// chrome.runtime.sendMessage(...) → browser.runtime.sendMessage(...)
```

**Recommendation**: Start with `.js` extension and copy code as-is. WXT supports vanilla JS.

### Step 6: Migrate Service Worker

#### `entrypoints/background.ts` (or `background.js`)

```javascript
// Copy service-worker.js content here
// WXT automatically registers this as the background service worker

// Optional: Replace chrome.* with browser.* for cross-browser compatibility
// WXT's polyfill makes this automatic, but explicit is better:

export default defineBackground(() => {
  // Original service-worker.js code goes here
  // Remove any top-level execution; wrap in this function if needed

  log("boot");

  chrome.runtime.onInstalled.addListener(() => {
    log("onInstalled");
    resetBadge();
  });

  // ... rest of service-worker.js code
});
```

**Recommendation**: Copy code as-is initially. WXT handles registration automatically.

### Step 7: Migrate Content Script

#### `entrypoints/content.ts` (or `content.js`)

```javascript
// Copy content.js here
// WXT auto-registers content scripts based on manifest config

export default defineContentScript({
  matches: [
    'https://*.teams.microsoft.com/*',
    'https://teams.cloud.microsoft/*',
  ],
  runAt: 'document_idle',
  allFrames: true,

  main() {
    // Original content.js code goes here
    // WXT automatically executes main() in the content script context

    chrome.runtime.onMessage.addListener((msg, _sender, sendResponse) => {
      // ... existing message handlers
    });
  },
});
```

**Recommendation**: Wrap existing code in `main()` function. Rest stays the same.

### Step 8: Update Manifest Icons

In `wxt.config.ts`, add icon configuration:

```typescript
export default defineConfig({
  // ... existing config
  manifest: {
    // ... existing manifest fields
    icons: {
      16: '/icons/action-16.png',
      32: '/icons/action-32.png',
      48: '/icons/action-48.png',
      128: '/icons/action-128.png',
    },
    action: {
      default_title: 'Teams Chat Exporter',
      default_icon: {
        16: '/icons/action-16.png',
        32: '/icons/action-32.png',
        48: '/icons/action-48.png',
        128: '/icons/action-128.png',
      },
    },
  },
});
```

---

## Code Changes Required

### Minimal Changes (Chrome build)

If targeting Chrome only initially:
- ✅ **No API changes needed** - `chrome.*` works as-is
- ✅ **File reorganization only** - Move files to WXT structure
- ✅ **Manifest extracted** - Move to `wxt.config.ts`

### Recommended Changes (Cross-browser)

For Firefox/Edge support, replace `chrome.*` with `browser.*`:

```javascript
// Before
chrome.runtime.sendMessage({ ... })
chrome.storage.local.get(...)
chrome.tabs.query(...)

// After (WXT polyfills this automatically)
browser.runtime.sendMessage({ ... })
browser.storage.local.get(...)
browser.tabs.query(...)
```

**Good news**: WXT's polyfill makes `chrome.*` work on Firefox automatically, so this is optional!

### Content Script Injection Adjustment

Your current code manually injects the content script as a fallback:

```javascript
// service-worker.js:378
await chrome.scripting.executeScript({
  target: { tabId, allFrames: true },
  files: ['content.js']
});
```

**WXT equivalent**:
```javascript
// No longer needed - WXT auto-registers content scripts
// Remove ensureContentScript() fallback or adjust for bundled path
```

**Recommendation**: Keep the fallback but update the file path to WXT's output structure.

---

## File Structure Mapping

| Current File | WXT Location | Notes |
|-------------|--------------|-------|
| `manifest.json` | `wxt.config.ts` | Converted to config |
| `popup.html` | `entrypoints/popup/index.html` | Remove `<script>` tag |
| `popup.js` | `entrypoints/popup/main.js` | Rename, auto-imported |
| `service-worker.js` | `entrypoints/background.js` | Auto-registered |
| `content.js` | `entrypoints/content.js` | Auto-registered |
| `icons/` | `public/icons/` | Copied to public dir |
| `docs/` | `docs/` | No change |
| `README.md` | `README.md` | Update build instructions |
| - | `wxt.config.ts` | **New**: Main config |
| - | `package.json` | **New**: Auto-generated |
| - | `.output/` | **New**: Build output |

---

## Testing Strategy

### Phase 1: Build Verification
```bash
npm run dev      # Start dev server with hot reload
npm run build    # Production build
npm run zip      # Create store-ready ZIP
```

**Checklist**:
- [ ] Build completes without errors
- [ ] Output in `.output/chrome-mv3/` contains all files
- [ ] Manifest.json generated correctly
- [ ] Icons copied to output

### Phase 2: Chrome Testing
```bash
npm run dev      # Load .output/chrome-mv3 in chrome://extensions
```

**Test Cases**:
- [ ] Extension loads without errors
- [ ] Popup opens and displays correctly
- [ ] Popup theme toggle works
- [ ] Date range inputs work
- [ ] Export button triggers scraping
- [ ] Messages collected correctly
- [ ] Badge updates during scraping
- [ ] JSON export downloads successfully
- [ ] CSV export downloads successfully
- [ ] HTML export downloads successfully
- [ ] Avatar embedding works (HTML)
- [ ] Empty chat shows banner (no download)
- [ ] Date filtering works correctly
- [ ] Quick range buttons work

### Phase 3: Firefox Testing
```bash
npm run dev:firefox    # If WXT supports it, otherwise:
npm run build
# Load .output/firefox-mv2/ in about:debugging
```

**Firefox-Specific Tests**:
- [ ] Extension loads in Firefox
- [ ] All Chrome test cases pass
- [ ] Download API works with blob URLs
- [ ] Storage persistence works
- [ ] Badge updates work

### Phase 4: Performance Testing

Compare against current version:
- [ ] Extension size (before vs. after)
- [ ] Load time
- [ ] Scraping performance (1000+ messages)
- [ ] Memory usage during large exports

---

## Rollback Plan

### If Migration Fails

1. **Keep current version separate**
   ```bash
   # Don't delete original directory
   # WXT is in separate folder: teams-chat-exporter-wxt/
   ```

2. **Branch strategy**
   ```bash
   git checkout -b wxt-migration
   # All WXT work on this branch
   # main branch stays untouched
   ```

3. **Gradual adoption**
   - Start with proof-of-concept in separate directory
   - Only merge to main after full testing
   - Keep v1.0.1 tagged for quick revert

### Red Flags to Stop Migration

- Build output > 2x current size (> 160KB)
- Scraping performance degrades > 20%
- Critical Chrome APIs unsupported
- Firefox build has showstopper bugs

---

## Timeline Estimate

| Phase | Tasks | Duration |
|-------|-------|----------|
| **Phase 1: Setup** | Install WXT, create structure, first build | 2-3 hours |
| **Phase 2: Migration** | Copy files, adjust code, configure manifest | 4-6 hours |
| **Phase 3: Testing** | Chrome + Firefox functional testing | 4-6 hours |
| **Phase 4: Docs** | Update README, CLAUDE.md, deployment guide | 1-2 hours |
| **Total** | End-to-end migration | **11-17 hours** |

### Breakdown by Role

**Developer Time**:
- Setup: 2 hours
- Code migration: 4 hours
- Testing: 6 hours
- **Total: 12 hours**

**Documentation Time**:
- Update guides: 2 hours

---

## Success Criteria

### Must-Have
- ✅ Chrome build works identically to current version
- ✅ Firefox build loads and exports chats successfully
- ✅ All existing features work (date filter, reactions, replies, avatars)
- ✅ Build size < 200KB (minified)
- ✅ No performance degradation

### Nice-to-Have
- ✅ TypeScript migration started (at least types for messages)
- ✅ ESLint configured
- ✅ Vitest test suite setup (even if minimal)
- ✅ Edge/Safari builds tested

---

## Post-Migration Tasks

### Immediate (After Merge)
1. Update Chrome Web Store listing with new build
2. Submit to Firefox Add-ons (mozilla.org)
3. Update README.md with "Available on Firefox" badge
4. Tag release: `v1.1.0-wxt`

### Short-Term (1-2 weeks)
1. Add TypeScript types for message objects
2. Set up ESLint + Prettier
3. Write unit tests for date/timestamp helpers
4. Add Playwright E2E test for basic export

### Long-Term (1-3 months)
1. Full TypeScript migration
2. Module extraction (per REFACTOR_PLAN.md)
3. CI/CD pipeline (GitHub Actions)
4. Automated release builds

---

## Resources

### Official Docs
- **WXT**: https://wxt.dev/
- **WXT Guide**: https://wxt.dev/guide/
- **webextension-polyfill**: https://github.com/mozilla/webextension-polyfill

### Example Projects
- WXT Examples: https://github.com/wxt-dev/wxt-examples
- Vanilla JS example: https://github.com/wxt-dev/wxt-examples/tree/main/examples/vanilla-js

### Community
- WXT Discord: https://discord.gg/ZFsZqGery9
- GitHub Discussions: https://github.com/wxt-dev/wxt/discussions

---

## Appendix: Common WXT Patterns

### A. Browser API Usage

```javascript
// WXT auto-imports browser from webextension-polyfill
import { browser } from 'wxt/browser';

// Use browser.* for cross-browser compatibility
browser.runtime.sendMessage({ type: 'PING' });
browser.storage.local.get('key');
browser.tabs.query({ active: true });
```

### B. Content Script Communication

```javascript
// entrypoints/content.ts
export default defineContentScript({
  matches: ['https://*.teams.microsoft.com/*'],

  main() {
    browser.runtime.onMessage.addListener((msg, sender, sendResponse) => {
      // Handle messages
      if (msg.type === 'SCRAPE_TEAMS') {
        // ... scraping logic
        sendResponse({ messages: [...] });
      }
      return true; // Keep channel open for async response
    });
  },
});
```

### C. Storage Helpers

```javascript
// WXT provides typed storage helpers
import { storage } from 'wxt/storage';

// Simpler API than chrome.storage
await storage.setItem('local:options', { theme: 'dark' });
const opts = await storage.getItem('local:options');
```

### D. Development Hot Reload

```javascript
// WXT automatically reloads extension on file changes in dev mode
// No special code needed - it just works!

// To manually reload:
npm run dev
// Edit any file → extension reloads instantly
```

---

**Document Version**: 1.0
**Last Updated**: 2025-01-23
**Next Review**: After Phase 1 completion
