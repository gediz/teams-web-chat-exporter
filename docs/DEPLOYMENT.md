# Deployment

## Build artifacts

- Chrome build: `.output/chrome-mv3/`
- Edge build: `.output/edge-mv3/`
- Firefox build: `.output/firefox-mv2/`
- Safari build: `.output/safari-mv2/`

## Commands

```bash
# Build
pnpm build
pnpm build:edge
pnpm build:firefox
pnpm build:safari

# Zip packages
pnpm zip
pnpm zip:edge
pnpm zip:firefox
pnpm zip:safari
```

## Version update

Update version in both files before release:

- `package.json`
- `wxt.config.ts`

## Store uploads

- Chrome Web Store: upload Chrome zip (`pnpm zip` output).
- Microsoft Edge Add-ons: upload Edge zip (`pnpm zip:edge` output).
- Firefox Add-ons (AMO): upload Firefox zip (`pnpm zip:firefox` output).
- Safari (macOS / iOS App Store): the Safari zip (`pnpm zip:safari` output) is a web-extension folder. To install or distribute, wrap it with Xcode's Safari Web Extension Converter (`xcrun safari-web-extension-converter`) and build the resulting `.app` / `.appex`. Apple Developer account required for distribution.

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
3. Build all targets (`pnpm build`, `pnpm build:edge`, and `pnpm build:firefox`).
4. Create all zip packages (`pnpm zip`, `pnpm zip:edge`, and `pnpm zip:firefox`).
5. Verify install in target browsers.
6. Upload to stores.
7. Tag and publish release notes.
