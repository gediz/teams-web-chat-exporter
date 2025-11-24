# Teams Chat Exporter - WXT Proof of Concept

This is a proof-of-concept migration of Teams Chat Exporter to the WXT framework.

## 📁 Structure

```
wxt-poc/
├── src/
│   ├── entrypoints/
│   │   ├── popup/
│   │   │   ├── index.html          # Popup root (mounts Svelte)
│   │   │   ├── main.ts             # Popup bootstrap (Svelte mount)
│   │   │   ├── App.svelte          # Popup UI/logic (TypeScript + Svelte)
│   │   │   └── popup.css           # Popup styles
│   │   ├── background.ts           # Service worker (from service-worker.js)
│   │   └── content.ts              # Content script (from content.js)
│   └── public/
│       └── icons/                  # Extension icons
├── wxt.config.ts                   # WXT configuration
├── package.json                    # Dependencies
└── .gitignore                      # Git ignore rules
```

## 🚀 Quick Start

### 1. Install Dependencies

```bash
cd wxt-poc
npm install
```

### 2. Development Mode (with Hot Reload)

```bash
# Chrome
npm run dev

# Firefox
npm run dev:firefox
```

This will:
- Start WXT in development mode
- Watch for file changes and auto-reload
- Output to `.output/chrome-mv3/` or `.output/firefox-mv2/`

### 3. Load in Browser

#### Chrome
1. Open `chrome://extensions/`
2. Enable "Developer mode" (top-right toggle)
3. Click "Load unpacked"
4. Select `.output/chrome-mv3/` directory
5. Pin the extension to toolbar

#### Firefox
1. Open `about:debugging#/runtime/this-firefox`
2. Click "Load Temporary Add-on"
3. Navigate to `.output/firefox-mv2/` and select `manifest.json`

### 4. Production Build

```bash
# Build for all browsers
npm run build

# Build for specific browser
npm run build:firefox

# Create store-ready ZIP
npm run zip
npm run zip:firefox
```

## 🧪 Testing Checklist

### Chrome Testing
- [ ] Extension loads without errors
- [ ] Popup opens and displays correctly
- [ ] Theme toggle works
- [ ] Date range inputs work
- [ ] Export button triggers scraping
- [ ] Messages collected correctly
- [ ] Badge updates during scraping
- [ ] JSON export downloads
- [ ] CSV export downloads
- [ ] HTML export downloads
- [ ] Avatar embedding works
- [ ] Empty chat shows banner
- [ ] Date filtering works

### Firefox Testing
Run all Chrome tests plus:
- [ ] Extension loads in Firefox
- [ ] Download API works
- [ ] Storage persistence works
- [ ] Badge updates work

### Performance Comparison
- [ ] Extension size (before vs. after)
- [ ] Load time
- [ ] Scraping performance (1000+ messages)
- [ ] Memory usage

## 📦 Build Output

### Chrome (`.output/chrome-mv3/`)
- Manifest V3
- Uses `chrome.*` namespace
- Service worker background

### Firefox (`.output/firefox-mv2/`)
- Manifest V2 (auto-converted)
- Uses `browser.*` namespace (polyfilled)
- Background page (not service worker)

## 🔧 Development Notes

### Current State
This POC now uses WXT with a Svelte + TypeScript popup:
- ✅ Popup migrated to Svelte/TypeScript (`App.svelte` + `main.ts`)
- ✅ Files organized under `src/` for WXT
- ✅ Manifest moved to `wxt.config.ts`
- ⚠️ Background/content scripts still plain JS (cross-browser polyfills in place)
- ⚠️ Further refactors/TS adoption in background/content pending

### Known Issues
1. **Content Script Injection**: The manual `chrome.scripting.executeScript` fallback in `background.js:378` may need adjustment for WXT's bundled paths
2. **Data URL Size**: Large HTML exports may exceed data URL limits (~50MB in Chrome)

### Next Steps
See [../docs/WXT_MIGRATION_PLAN.md](../docs/WXT_MIGRATION_PLAN.md) for the full migration roadmap.

## 📚 Resources

- **WXT Documentation**: https://wxt.dev/
- **Migration Plan**: [../docs/WXT_MIGRATION_PLAN.md](../docs/WXT_MIGRATION_PLAN.md)
- **Original Extension**: [../](../)

## 🐛 Troubleshooting

### Build Fails
```bash
# Clear cache and rebuild
rm -rf .output .wxt node_modules
npm install
npm run build
```

### Extension Won't Load
- Check `.output/chrome-mv3/manifest.json` exists
- Verify icons copied to `.output/chrome-mv3/icons/`
- Check browser console for errors

### Hot Reload Not Working
- Ensure you're running `npm run dev` (not `npm run build`)
- WXT will automatically reload on file changes
- Check terminal for WXT server logs

---

**Status**: Proof of Concept
**Last Updated**: 2025-01-23
