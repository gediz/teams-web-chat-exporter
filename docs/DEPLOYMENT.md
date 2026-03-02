# Deployment

## Build artifacts

- Chrome/Edge build: `.output/chrome-mv3/`
- Firefox build: `.output/firefox-mv2/`

## Commands

```bash
# Build
npm run build
npm run build:firefox

# Zip packages
npm run zip
npm run zip:firefox
```

## Version update

Update version in both files before release:

- `package.json`
- `wxt.config.ts`

## Store uploads

- Chrome Web Store: upload Chrome zip (`npm run zip` output)
- Edge Add-ons: use same Chrome zip
- Firefox Add-ons (AMO): upload Firefox zip (`npm run zip:firefox` output)

## AMO reviewer notes

If Firefox review asks for build steps or data collection info, paste this:

```
Source: https://github.com/gediz/teams-web-chat-exporter
Requires: Node.js LTS, npm

Build steps:
1. npm install
2. npm run build:firefox
3. Output: .output/firefox-mv2/

Data collection: None. All exports are saved locally. No data is transmitted.
```

## Release checklist

1. Update versions.
2. Run `npm run check`.
3. Build both targets.
4. Create both zip packages.
5. Verify install in target browsers.
6. Upload to stores.
7. Tag and publish release notes.
