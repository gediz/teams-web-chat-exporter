# Deployment

## Build artifacts

- Chrome/Edge build: `.output/chrome-mv3/`
- Firefox build: `.output/firefox-mv2/`

## Commands

```bash
# Build
pnpm build
pnpm build:firefox

# Zip packages
pnpm zip
pnpm zip:firefox
```

## Version update

Update version in both files before release:

- `package.json`
- `wxt.config.ts`

## Store uploads

- Chrome Web Store: upload Chrome zip (`pnpm zip` output)
- Edge Add-ons: use same Chrome zip
- Firefox Add-ons (AMO): upload Firefox zip (`pnpm zip:firefox` output)

## AMO reviewer notes

If Firefox review asks for build steps or data collection info, paste this:

```
Source: https://github.com/gediz/teams-web-chat-exporter
Requires: Node.js 24+, pnpm 10+

Build steps:
1. pnpm install
2. pnpm build:firefox
3. Output: .output/firefox-mv2/

Data collection: None. The extension reads messages from the Teams Chat Service
API and Microsoft Graph API using the user's existing session tokens. All
exported data is saved locally. No data is sent to any third-party server or
to the extension developer.
```

## Release checklist

1. Update versions in both `package.json` and `wxt.config.ts`.
2. Run `pnpm check`.
3. Build both targets (`pnpm build` and `pnpm build:firefox`).
4. Create both zip packages (`pnpm zip` and `pnpm zip:firefox`).
5. Verify install in target browsers.
6. Upload to stores.
7. Tag and publish release notes.
