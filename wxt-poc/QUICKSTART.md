# 🚀 WXT POC Quick Start

Get the WXT proof-of-concept running in under 5 minutes!

## Step 1: Install Dependencies

```bash
cd wxt-poc
npm install
```

**Expected output**: npm will install WXT and dependencies (~30 seconds)

## Step 2: Start Development Server

```bash
npm run dev
```

**Expected output**:
```
WXT 0.19.x
Building chrome-mv3 for development...
✓ Built in XXXms
```

## Step 3: Load in Chrome

1. Open Chrome and go to: `chrome://extensions/`
2. Enable "**Developer mode**" (toggle in top-right corner)
3. Click "**Load unpacked**"
4. Navigate to: `wxt-poc/.output/chrome-mv3/`
5. Click "**Select Folder**"
6. Pin the extension to your toolbar

## Step 4: Test the Extension

1. Go to `teams.microsoft.com`
2. Open any chat conversation
3. Click the extension icon
4. Try exporting a chat!

---

## 🎯 That's It!

You're now running the WXT version of Teams Chat Exporter with:
- ✅ Hot module reload (edit code → auto-reload)
- ✅ Same functionality as original
- ✅ Cross-browser support ready

## 📝 What Changed?

Visually: **Nothing!** The extension works exactly the same.

Under the hood:
- Files reorganized into WXT structure
- Manifest extracted to `wxt.config.ts`
- Build system powered by Vite
- Auto-reload on code changes

## 🔥 Try Hot Reload

1. Keep `npm run dev` running
2. Edit `src/entrypoints/popup/index.html` (change the title)
3. Watch the extension auto-reload!

## 🦊 Test on Firefox

```bash
npm run dev:firefox
```

Then load from `.output/firefox-mv2/` in Firefox's `about:debugging`

---

## Next: Full Migration

See [../docs/WXT_MIGRATION_PLAN.md](../docs/WXT_MIGRATION_PLAN.md) for the complete migration plan.
