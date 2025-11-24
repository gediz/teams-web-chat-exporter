# WXT Deployment Guide - Teams Chat Exporter

Complete guide for building, testing, and deploying to browser stores.

---

## Table of Contents
- [Quick Reference](#quick-reference)
- [Build Commands](#build-commands)
- [Development Testing](#development-testing)
- [Store Deployment](#store-deployment)
- [Firefox Permanent Install](#firefox-permanent-install)
- [Updating Existing Store Listings](#updating-existing-store-listings)
- [Version Management](#version-management)

---

## Quick Reference

```bash
# Development
npm run dev              # Chrome dev build (manual reload)
npm run dev:firefox      # Firefox dev build (auto-reload)

# Production builds
npm run build            # Build Chrome only
npm run build:firefox    # Build Firefox only

# Store-ready ZIPs
npm run zip              # Create chrome.zip
npm run zip:firefox      # Create firefox.zip
```

---

## Build Commands

### 1. Development Builds

**Chrome Development:**
```bash
npm run dev
```
- Output: `.output/chrome-mv3/`
- Hot reload: Disabled (to prevent Chrome throttling)
- Manual reload: Required after code changes
- **Use for**: Active development

**Firefox Development:**
```bash
npm run dev:firefox
```
- Output: `.output/firefox-mv2/`
- Hot reload: Enabled (works well on Firefox)
- **Use for**: Active development

---

### 2. Production Builds

**Build for Chrome:**
```bash
npm run build
```
- Output: `.output/chrome-mv3/`
- Minified: Yes
- Source maps: No
- **Use for**: Final testing and store upload

**Build for Firefox:**
```bash
npm run build:firefox
```
- Output: `.output/firefox-mv2/`
- Minified: Yes
- Manifest: Converted to V2 automatically
- **Use for**: Final testing and store upload

**Build for All Browsers:**
```bash
npm run build && npm run build:firefox
```
- Builds both Chrome and Firefox versions
- **Use for**: Release preparation

---

### 3. Store-Ready ZIP Files

**Create Chrome ZIP:**
```bash
npm run zip
```
- Creates: `.output/teams-chat-exporter-1.1.0-chrome.zip`
- Contains: Minified production build
- Ready for: Chrome Web Store upload

**Create Firefox ZIP:**
```bash
npm run zip:firefox
```
- Creates: `.output/teams-chat-exporter-1.1.0-firefox.zip`
- Contains: Minified MV2 build
- Ready for: Firefox Add-ons upload

**Create All ZIPs:**
```bash
npm run zip && npm run zip:firefox
```

---

## Development Testing

### Chrome Testing

**Load Extension:**
1. Run `npm run build` (or `npm run dev`)
2. Open `chrome://extensions/`
3. Enable "Developer mode" (top-right toggle)
4. Click "Load unpacked"
5. Select `.output/chrome-mv3/`

**Update Extension:**
1. Make code changes
2. Run `npm run build` again
3. Click refresh icon on extension card in `chrome://extensions/`

**Troubleshooting:**
- **"Extension reloaded too frequently"**: Use `npm run build` instead of `npm run dev`
- **Changes not appearing**: Hard refresh the extension (remove and re-add)

---

### Firefox Testing

#### Temporary Install (Removed on Restart)

**Load Extension:**
1. Run `npm run dev:firefox` (or `npm run build:firefox`)
2. Open `about:debugging#/runtime/this-firefox`
3. Click "Load Temporary Add-on..."
4. Navigate to `.output/firefox-mv2/`
5. Select `manifest.json`

**Update Extension:**
- With `npm run dev:firefox`: Auto-reloads on changes
- With `npm run build:firefox`: Click "Reload" button in debugging page

**Limitation**: Extension removed when Firefox closes

---

#### Permanent Install (Survives Restarts)

**Option 1: Self-Sign (Recommended for Development)**

```bash
# Install web-ext globally
npm install -g web-ext

# Build Firefox version
npm run build:firefox

# Sign the extension (requires Firefox Add-ons account)
cd .output/firefox-mv2
web-ext sign --api-key=YOUR_KEY --api-secret=YOUR_SECRET
```

This creates a signed `.xpi` file that can be installed permanently.

**Get API credentials:**
1. Go to https://addons.mozilla.org/developers/addon/api/key/
2. Generate API credentials
3. Use in `web-ext sign` command

**Option 2: Disable Signature Verification (Firefox Developer/Nightly Only)**

⚠️ **Only works on Firefox Developer Edition or Nightly**

1. Open `about:config`
2. Search: `xpinstall.signatures.required`
3. Set to `false`
4. Install `.output/firefox-mv2/` as temporary add-on
5. Will persist across restarts

**Option 3: Use Firefox Developer Edition**

Firefox Developer Edition allows unsigned extensions:
1. Download from https://www.mozilla.org/firefox/developer/
2. Install extension as temporary add-on
3. Enable persistence in `about:config`:
   - Set `extensions.experiments.enabled` to `true`
   - Set `xpinstall.signatures.required` to `false`

---

## Store Deployment

### Chrome Web Store

#### Initial Upload (If Not Published Yet)

1. **Build:**
   ```bash
   npm run zip
   ```

2. **Go to Chrome Web Store Developer Dashboard:**
   - https://chrome.google.com/webstore/devconsole

3. **Create New Item:**
   - Click "New Item"
   - Upload `.output/teams-chat-exporter-1.1.0-chrome.zip`
   - Fill out store listing:
     - Name: Teams Chat Exporter
     - Description: (from README)
     - Screenshots: (capture from extension)
     - Category: Productivity
     - Language: English

4. **Submit for Review:**
   - Review can take 1-3 days
   - Check email for approval/rejection

#### Update Existing Listing

Since you already have it on Chrome Web Store:

1. **Build new version:**
   ```bash
   npm run zip
   ```

2. **Update version in `wxt.config.ts`:**
   ```typescript
   version: '1.2.0', // Increment version
   ```

3. **Rebuild:**
   ```bash
   npm run zip
   ```

4. **Upload to existing listing:**
   - Go to https://chrome.google.com/webstore/devconsole
   - Find "Teams Chat Exporter"
   - Click "Package" tab
   - Click "Upload New Package"
   - Select `.output/teams-chat-exporter-1.2.0-chrome.zip`
   - Update "What's New" section
   - Click "Submit for Review"

**Notes:**
- Version must be higher than current published version
- Updates typically reviewed faster than new submissions (hours, not days)

---

### Firefox Add-ons (AMO)

#### Prerequisites

**Set a unique Add-on ID** in `wxt.config.ts`:

```typescript
browser_specific_settings: {
  gecko: {
    id: 'teams-chat-exporter@yourdomain.com', // Change to your email/domain
    strict_min_version: '109.0',
    data_collection_permissions: {
      required: ["none"], // This extension doesn't collect data
    },
  },
}
```

**Important**: Replace `yourdomain.com` with:
- Your email: `teams-chat-exporter@yourname.com`
- Your domain: `teams-chat-exporter@github.io`
- Format: Must be email-like (has `@`)

#### Initial Upload

1. **Build:**
   ```bash
   npm run zip:firefox
   ```

2. **Create Firefox Add-ons Account:**
   - Go to https://addons.mozilla.org/developers/
   - Sign up or log in

3. **Submit New Add-on:**
   - Click "Submit a New Add-on"
   - Select "On this site"
   - Upload `.output/teams-chat-exporter-1.1.0-firefox.zip`
   - Select "Firefox" as target
   - Choose distribution channel:
     - **Listed**: Public on AMO (recommended)
     - **Unlisted**: Get signed XPI for self-distribution

4. **Fill out Listing Information:**
   - Name: Teams Chat Exporter
   - Summary: (short description)
   - Description: (from README)
   - Categories: Productivity, Social & Communication
   - Tags: teams, microsoft, chat, export
   - Screenshots: (capture from Firefox)
   - License: MIT (or your license)

5. **Review Process:**
   - **Automated review**: Minutes to hours
   - **Manual review** (if flagged): 1-7 days
   - Check email for status

#### Update Existing Listing

1. **Build new version:**
   ```bash
   npm run zip:firefox
   ```

2. **Update version in `wxt.config.ts`:**
   ```typescript
   version: '1.2.0',
   ```

3. **Rebuild:**
   ```bash
   npm run zip:firefox
   ```

4. **Upload new version:**
   - Go to https://addons.mozilla.org/developers/addons
   - Find "Teams Chat Exporter"
   - Click "Upload New Version"
   - Upload `.output/teams-chat-exporter-1.2.0-firefox.zip`
   - Describe changes in "Version Notes"
   - Submit for review

**Notes:**
- Firefox reviews source code - be prepared to explain obfuscated code
- If using external libraries, document them
- Provide test credentials if app requires login

---

#### Notes to Reviewer (for Firefox Add-ons Submission)

When submitting to Firefox Add-ons, you'll be asked: **"Is there anything our reviewers should bear in mind when reviewing this add-on?"**

Copy and paste the following into the "Notes to Reviewer" field:

```
# Teams Chat Exporter - Build Instructions

## Overview
This extension is built using WXT (https://wxt.dev/), a modern framework for cross-browser extensions.
The source code is available at: https://github.com/gediz/teams-web-chat-exporter

## Step-by-Step Build Instructions

### Prerequisites
- Node.js 22.x or higher
- npm (comes with Node.js)

### Build Steps

1. Clone the repository:
   git clone https://github.com/gediz/teams-web-chat-exporter.git
   cd teams-web-chat-exporter/wxt-poc

2. Install dependencies:
   npm install

3. Build the Firefox version:
   npm run build:firefox

4. The built extension will be in: .output/firefox-mv2/

5. (Optional) Create distribution ZIP:
   npm run zip:firefox
   Output: .output/teams-chat-exporter-1.1.0-firefox.zip

### Verification

To verify the build matches the submitted ZIP:
- The build is deterministic (same source = same output)
- Compare .output/firefox-mv2/ contents with the uploaded ZIP
- All source code is in the repository under wxt-poc/src/entrypoints/

### Build System Explanation

- WXT automatically converts Manifest V3 code to Firefox's Manifest V2
- Browser API compatibility is handled by WXT's polyfills
- Code is bundled and minified by Vite (WXT's underlying build tool)
- No obfuscation is used, only standard minification

### Dependencies

All dependencies are listed in package.json:
- wxt: ^0.19.0 (build framework, dev dependency)
- Svelte + @sveltejs/vite-plugin-svelte + svelte-check (dev-only; popup UI)
- No runtime dependencies (extension uses only browser APIs)

### Architecture

- src/entrypoints/popup/: Extension popup UI (Svelte + TypeScript; App.svelte + main.ts)
- src/entrypoints/background.ts: Background service worker (message handling, downloads)
- src/entrypoints/content.ts: Content script (scrapes Teams web page DOM)
- wxt.config.ts: Build configuration and manifest settings
- src/public/icons/: Extension icons

### Data Collection

This extension does NOT collect or transmit any user data:
- All exports are saved locally to the user's device
- No analytics, telemetry, or tracking
- No external network requests (except Teams page access)
- See browser_specific_settings.gecko.data_collection_permissions in manifest

### Testing

No special test credentials are required. To test:
1. Open https://teams.microsoft.com/ and log in with any Microsoft Teams account
2. Navigate to a chat conversation
3. Click the extension icon
4. Configure export options and click "Export current chat"
5. Verify downloaded file (JSON/CSV/HTML) contains chat messages
```

**Tips for Submission:**
- Include the GitHub repository URL when uploading source code
- The WXT build system is widely used and recognized by Mozilla reviewers
- If reviewers have questions about the build process, point them to WXT documentation
- The extension code is straightforward - no complex build steps or custom tooling

---

### Edge Add-ons (Optional)

Edge uses Chromium, so the Chrome build works:

1. **Build:**
   ```bash
   npm run zip
   ```

2. **Go to Edge Partner Center:**
   - https://partner.microsoft.com/dashboard/microsoftedge

3. **Submit:**
   - Click "New Extension"
   - Upload Chrome ZIP (same as Chrome Web Store)
   - Fill out similar metadata

**Note**: Edge has smaller user base but easy to support since it's the same build as Chrome.

---

## Updating Existing Store Listings

### Your Current Situation

You mentioned you already have the extension on Chrome Web Store. Here's how to update it to the WXT version:

#### Step 1: Verify WXT Build Works

```bash
npm run build
cd .output/chrome-mv3
# Manually test this version thoroughly
```

#### Step 2: Update Version Number

Edit `wxt-poc/wxt.config.ts`:
```typescript
version: '1.1.0', // Change to '1.2.0' or higher
```

#### Step 3: Create Store-Ready ZIP

```bash
npm run zip
```

#### Step 4: Upload to Chrome Web Store

1. Go to https://chrome.google.com/webstore/devconsole
2. Find your existing "Teams Chat Exporter" listing
3. Click the listing
4. Go to "Package" tab
5. Click "Upload New Package"
6. Select `.output/teams-chat-exporter-1.2.0-chrome.zip`
7. Update release notes:
   ```
   v1.2.0 - Built with WXT Framework
   - Added Firefox support
   - Improved cross-browser compatibility
   - Enhanced build system with Vite
   - All existing features preserved
   ```
8. Submit for review

#### Step 5: Monitor Review

- Chrome usually reviews updates within 24-48 hours
- Check email for approval
- Once approved, users auto-update

---

## Version Management

### Semantic Versioning

Use semantic versioning: `MAJOR.MINOR.PATCH`

**Examples:**
- `1.0.0` → `1.0.1`: Bug fixes (PATCH)
- `1.0.1` → `1.1.0`: New features, backward compatible (MINOR)
- `1.1.0` → `2.0.0`: Breaking changes (MAJOR)

### Update Version in `wxt.config.ts`

```typescript
export default defineConfig({
  manifest: {
    version: '1.2.0', // Update here
    // ...
  },
});
```

### Release Checklist

Before each release:

- [ ] Update version in `wxt.config.ts`
- [ ] Test on Chrome (`npm run build` + manual test)
- [ ] Test on Firefox (`npm run build:firefox` + manual test)
- [ ] Update CHANGELOG.md (if you have one)
- [ ] Create git tag: `git tag v1.2.0`
- [ ] Build release: `npm run zip && npm run zip:firefox`
- [ ] Upload to Chrome Web Store
- [ ] Upload to Firefox Add-ons
- [ ] Push tag: `git push origin v1.2.0`

---

## Build Output Structure

After running builds:

```
.output/
├── chrome-mv3/              # Chrome build
│   ├── manifest.json        # MV3 manifest
│   ├── popup.html           # Popup HTML (Svelte entry)
│   ├── chunks/popup-*.js    # Bundled & minified popup code
│   ├── assets/popup-*.css   # Popup CSS
│   ├── background.js        # Bundled & minified
│   ├── content.js           # Bundled & minified
│   └── icons/
├── firefox-mv2/             # Firefox build
│   ├── manifest.json        # MV2 manifest (auto-converted)
│   ├── popup.html           # Popup HTML (Svelte entry)
│   ├── chunks/popup-*.js    # Popup bundle
│   ├── assets/popup-*.css   # Popup CSS
│   ├── background.js        # Background page (not service worker)
│   ├── content.js
│   └── icons/
└── *.zip                    # Store-ready ZIPs
```

**Key differences:**
- **Chrome**: Manifest V3, service worker
- **Firefox**: Manifest V2, background page, `browserAction`

WXT handles all conversions automatically!

---

## Troubleshooting

### Build Errors

**"Module not found"**
```bash
rm -rf node_modules .output
npm install
npm run build
```

**"TypeScript errors"**
- WXT expects TypeScript config
- Ignore warnings for `.js` files
- Consider migrating to `.ts` gradually

### Store Rejections

**Chrome: "Code obfuscation detected"**
- Chrome may flag minified code
- Solution: Include source maps or unminified build
- In `wxt.config.ts`:
  ```typescript
  build: {
    sourcemap: true,
  }
  ```

**Firefox: "Requires source code review"**
- Firefox reviews source for security
- Solution: Include link to GitHub repo in submission notes
- Example: "Source code: https://github.com/gediz/teams-web-chat-exporter"

### Size Limits

**Chrome Web Store**: 128 MB max (you're safe at ~100KB)
**Firefox Add-ons**: 200 MB max

If you exceed limits:
- Remove unused dependencies
- Optimize images
- Use external hosting for large assets

---

## Advanced: CI/CD Automation

For automated releases, add to `.github/workflows/release.yml`:

```yaml
name: Release
on:
  push:
    tags:
      - 'v*'

jobs:
  build:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v3
      - uses: actions/setup-node@v3
        with:
          node-version: 22
      - run: npm install
      - run: npm run zip && npm run zip:firefox
      - uses: actions/upload-artifact@v3
        with:
          name: extension-builds
          path: .output/*.zip
```

This auto-builds ZIPs when you push a version tag.

---

## Summary

### For Your Current Workflow:

1. **Develop**: `npm run dev` (Chrome) or `npm run dev:firefox` (Firefox)
2. **Test**: Load from `.output/chrome-mv3/` or `.output/firefox-mv2/`
3. **Release**:
   ```bash
   # Update version in wxt.config.ts
   npm run zip && npm run zip:firefox
   # Upload ZIPs to stores
   ```

### Store Upload URLs:

- **Chrome**: https://chrome.google.com/webstore/devconsole
- **Firefox**: https://addons.mozilla.org/developers/addons
- **Edge**: https://partner.microsoft.com/dashboard/microsoftedge

---

**Questions?** Check WXT docs: https://wxt.dev/guide/publishing.html

**Last Updated**: 2025-01-23
