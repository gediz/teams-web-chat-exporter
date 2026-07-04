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

Then update `CHANGELOG.md`. The entry must cover **every** user-facing change
since the previous tag, not just the last thing worked on. List the full commit
range and check each one is either reflected in the entry or is internal-only:

```sh
git log --oneline v<previous>..HEAD
```

Internal commits (docs, i18n string files, dead-code removal, dev-only tweaks)
can be left out; every feature or fix a user would notice must appear.

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
2. Update `CHANGELOG.md` to cover **all** user-facing changes since the previous
   tag. Verify against `git log --oneline v<previous>..HEAD` so nothing is
   missed (this list is the single source of truth for what shipped).
3. Run `pnpm check`.
4. Build all targets (`pnpm build`, `pnpm build:edge`, and `pnpm build:firefox`).
5. Create all zip packages (`pnpm zip`, `pnpm zip:edge`, and `pnpm zip:firefox`).
6. Verify install in target browsers.
7. Upload to stores.
8. Tag and publish release notes. The tag must point at the commit that carries
   the complete `CHANGELOG.md`.

The version-bump commit (the one that raises the version in `package.json` and
`wxt.config.ts` and carries the complete `CHANGELOG.md`) must be the **tip** of
the branch at release time, and the tag points at it. Any docs or chore commits
(comment cleanups, this checklist, and the like) land **before** it, not after,
so the last commit on the branch is always the release itself.
