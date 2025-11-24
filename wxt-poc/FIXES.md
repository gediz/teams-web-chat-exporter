# WXT POC Fixes Applied

**Note:** The popup is now implemented in Svelte + TypeScript (`src/entrypoints/popup/App.svelte` + `main.ts`). Background and content scripts have been moved to TypeScript entrypoints (`background.ts`, `content.ts`) with compatibility shims.

## Issue: Default export not found

**Error**: `Default export not found, did you forget to call "export default defineBackground(...)"?`

**Cause**: WXT requires entrypoints to use specific wrapper functions, but we copied the raw JavaScript files.

**Fixes Applied**:

### 1. `src/entrypoints/background.js`
Wrapped entire service worker code in `defineBackground()`:

```javascript
export default defineBackground(() => {
  // All original service-worker.js code here
});
```

### 2. `src/entrypoints/content.js`
Wrapped content script in `defineContentScript()` with configuration:

```javascript
export default defineContentScript({
  matches: [
    'https://*.teams.microsoft.com/*',
    'https://teams.cloud.microsoft/*',
  ],
  runAt: 'document_idle',
  allFrames: true,

  main() {
    // All original content.js code here
  }
});
```

### 3. `src/entrypoints/popup/main.js`
No changes needed - popup scripts don't require special wrappers.

---

### 4. `wxt.config.ts` - Disable Auto-Browser & Hot Reload
Added `runner.disabled: true` to prevent WXT from auto-opening a blank Chrome window.
Added `dev.reloadOnChange: false` to prevent Chrome's "reloaded too frequently" throttling error.

### 5. `src/entrypoints/popup/index.html` - Script Tag
Added `<script type="module" src="./main.js"></script>` so WXT injects the popup script.

### 6. `src/entrypoints/popup/main.js` - Firefox Compatibility
Added browser API compatibility layer at the top:

```javascript
const runtime = typeof browser !== 'undefined' ? browser.runtime : chrome.runtime;
const tabs = typeof browser !== 'undefined' ? browser.tabs : chrome.tabs;
const storage = typeof browser !== 'undefined' ? browser.storage : chrome.storage;
```

Replaced all `chrome.tabs`, `chrome.runtime`, and `chrome.storage` calls with the compatibility wrappers.

**Why**: Firefox uses `browser.*` API instead of `chrome.*`. This ensures cross-browser compatibility.

### 7. `src/entrypoints/background.js` - Firefox Compatibility
Added browser API compatibility layer:

```javascript
const runtime = typeof browser !== 'undefined' ? browser.runtime : chrome.runtime;
const tabs = typeof browser !== 'undefined' ? browser.tabs : chrome.tabs;
// Firefox MV2 uses browserAction, Chrome MV3 uses action
const action = typeof browser !== 'undefined'
    ? (browser.action || browser.browserAction)
    : chrome.action;
const downloads = typeof browser !== 'undefined' ? browser.downloads : chrome.downloads;
const scripting = typeof browser !== 'undefined' ? browser.scripting : chrome.scripting;
```

**Note**: Firefox uses Manifest V2 (`browserAction`) while Chrome uses Manifest V3 (`action`).

Replaced all `chrome.*` API calls with the compatibility wrappers.

**Special fix for `runtime.sendMessage`**: Firefox doesn't always return a Promise, so wrapped in try-catch:
```javascript
try {
    const msgPromise = runtime.sendMessage({ type: 'EXPORT_STATUS', ...enriched });
    if (msgPromise && msgPromise.catch) msgPromise.catch(() => { });
} catch (e) {
    // Ignore errors when popup is closed
}
```

### 8. `src/entrypoints/content.js` - Firefox Compatibility
Added browser API compatibility:

```javascript
const runtime = typeof browser !== 'undefined' ? browser.runtime : chrome.runtime;
```

Wrapped all `runtime.sendMessage` calls to handle Firefox's non-Promise behavior:
```javascript
try {
    const msgPromise = runtime.sendMessage({ type: "SCRAPE_PROGRESS", payload: { ... } });
    if (msgPromise && msgPromise.catch) msgPromise.catch(() => { });
} catch (e) { /* ignore */ }
```

### 9. `src/entrypoints/background.js` - Firefox Download Fix
Firefox blocks data URLs in the downloads API. Added blob URL support:

```javascript
// Firefox-compatible: Create blob URL (Firefox blocks data URLs in downloads)
function textToBlobUrl(text, mime) {
    const blob = new Blob([text], { type: mime });
    return URL.createObjectURL(blob);
}

const isFirefox = typeof browser !== 'undefined' && navigator.userAgent.includes('Firefox');

// Use blob URLs for Firefox, data URLs for Chrome
const url = isFirefox ? textToBlobUrl(content, mime) : textToDataUrl(content, mime);
```

**Why**: Firefox security policy blocks data URLs in `browser.downloads.download()`. Blob URLs work in both browsers but require cleanup.

**Cleanup**: Blob URLs are revoked 60 seconds after download starts to prevent memory leaks.

### 10. `src/entrypoints/background.js` - Firefox Badge Fix
Firefox uses different badge API. Updated action compatibility:

```javascript
// Firefox MV2 uses browserAction, Chrome MV3 uses action
const action = typeof browser !== 'undefined'
    ? (browser.action || browser.browserAction)
    : chrome.action;
```

**Why**: WXT builds Firefox as Manifest V2 which uses `browserAction` instead of `action`. This fallback ensures badges work on both browsers.

---

## Ready to Install

Now you can run:

```bash
npm install
npm run dev
```

**What happens**:
- ✅ No blank browser window opens
- ✅ Build completes successfully
- ✅ Extension ready in `.output/chrome-mv3/`

**Load manually in Chrome**:
1. Open `chrome://extensions/`
2. Enable "Developer mode"
3. Click "Load unpacked"
4. Select `.output/chrome-mv3/`

**Load manually in Firefox**:
```bash
npm run dev:firefox
```
1. Open `about:debugging#/runtime/this-firefox`
2. Click "Load Temporary Add-on..."
3. Navigate to `.output/firefox-mv2/`
4. Select the `manifest.json` file
5. Click "Open"

The popup buttons will now work on both browsers! 🚀
