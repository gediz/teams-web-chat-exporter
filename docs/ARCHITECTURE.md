# Architecture

Technical summary of how the extension works.

## Stack

- WXT
- TypeScript
- Svelte 5 (popup UI)

## Project structure

```
src/
├── entrypoints/
│   ├── popup/              # Svelte UI
│   │   ├── App.svelte      # Main popup component
│   │   ├── components/     # UI components (ExportButton, FormatSection,
│   │   │                   #   IncludeSection, DateRangeSection, TargetSection,
│   │   │                   #   SettingsPage, HistoryPage, HeaderActions,
│   │   │                   #   OnboardingOverlay)
│   │   └── popup.css       # Styles
│   ├── background.ts       # Service worker / background script
│   └── content.ts          # Content script (scraper)
├── background/             # Background script modules
│   ├── builders.ts         # toCSV, toHTML + download-URL helpers
│   ├── download.ts         # Build pipeline, JSON/TXT serialization,
│   │                       #   single-format + HTML.zip + bundle.zip paths
│   ├── pdf.ts              # PDF builder (pdf-lib + HarfBuzz font subsets
│   │                       #   + Twemoji rasterization)
│   ├── font-subset.ts      # WASM wrapper around HarfBuzz `hb-subset`
│   │                       #   (drops unused glyphs + GSUB/GPOS/GDEF/kern)
│   └── zip.ts              # ZIP archive creation (via fflate)
├── content/                # Content script modules
│   ├── api-client.ts       # Teams Chat Service API client (MSAL tokens,
│   │                       #   discovery, pagination, Graph photos)
│   ├── api-converter.ts    # API response → ExportMessage + Giphy parser
│   ├── scroll.ts           # DOM auto-scroll message collection
│   ├── reactions.ts        # Reaction + reactor parsing
│   ├── replies.ts          # Reply context parsing
│   ├── attachments.ts      # Attachment extraction (files, images, links,
│   │                       #   previews) + domain-allowlisted image fetch
│   ├── text.ts             # Text extraction with emoji/mention handling
│   └── title.ts            # Chat/channel title extraction
├── utils/                  # Shared utilities
│   ├── time.ts             # Time formatting and parsing
│   ├── text.ts             # Text processing (normalize, score, escape)
│   ├── dom.ts              # DOM query helpers ($, $$)
│   ├── options.ts          # Settings persistence + Options type + storage keys
│   ├── messaging.ts        # Runtime message wrapper
│   ├── badge.ts            # Extension icon badge manager
│   ├── messages.ts         # Day dividers, filename sanitization, merge helpers
│   ├── avatars.ts          # Avatar ID extraction from URLs
│   └── teams-urls.ts       # Teams URL match patterns (manifest + runtime)
├── types/                  # TypeScript types
│   ├── messaging.ts        # IPC message types (popup/background/content)
│   └── shared.ts           # Domain types (ExportMessage, ScrapeOptions, etc.)
├── i18n/                   # Internationalization
│   ├── i18n.ts             # t() helper + locale loader
│   └── locales/            # 24 translation files (keys kept in parity)
└── public/                 # Static assets served by the extension
    ├── icons/              # Extension icons (16, 32, 48, 128px)
    ├── fonts/              # Bundled Noto TTFs (Regular, Bold, SC)
    ├── twemoji/            # Vendored Twemoji SVGs (copied in at postinstall
    │                       #   from @twemoji/svg; gitignored)
    └── wasm/               # hb-subset.wasm copied from harfbuzzjs at
                            #   postinstall (gitignored)
```

`scripts/vendor-twemoji.mjs` copies Twemoji SVGs and `hb-subset.wasm` into `src/public/` after `pnpm install`. Both subtrees are gitignored.

## Component communication

```
┌─────────────┐
│   Popup     │  User configures export
│  (Svelte)   │  Sends START_EXPORT
└──────┬──────┘
       │
       ↓ chrome.runtime.sendMessage
┌─────────────┐
│ Background  │  Orchestrates export
│  (Service   │  Sends SCRAPE_TEAMS to content
│   Worker /  │  Builds export files (single, HTML.zip, or bundle.zip)
│   Event     │  Appends a HistoryEntry on success/cancel
│   Page)     │
└──────┬──────┘
       │
       ↓ chrome.tabs.sendMessage
┌─────────────┐
│  Content    │  1. Tries API fetch (fast, complete)
│  Script     │  2. Falls back to DOM scroll if API fails
│             │  Streams messages back via runtime port
└─────────────┘
```

All end-user export runs flow through `START_EXPORT`; the `START_EXPORT_ZIP` message type was removed in the multi-format migration.

## Runtime flow

1. Popup sends `START_EXPORT` to background with `formats: OptionFormat[]` plus include/avatar/download/pdf flags.
2. Background checks tab context and injects the content script if needed.
3. Background asks the content script to scrape Teams (`SCRAPE_TEAMS`).
4. Content script tries API-based fetch first:
   - Reads MSAL tokens from localStorage (IC3, Skype, Graph). Handles MSAL Browser v4+ encrypted cache entries. Supports commercial and GCC High endpoints.
   - Discovers the chat service endpoint via the Teams authz API.
   - Fetches messages page by page from the chat service `/messages` endpoint (cap: `MAX_PAGES = 500` in `api-client.ts`, retries on 429 with exponential backoff + adaptive inter-page delay after page 20).
   - Resolves unknown user MRIs (forwarded senders, reactors) via Microsoft Graph when a Graph token is available.
   - Converts API response objects to ExportMessage format (`api-converter.ts`).
   - Fetches inline image/audio data from AMS URLs (max 5 MB per file, 6 concurrent fetches, retries 408/410/429/5xx with exponential backoff).
   - Fetches avatar photos via Microsoft Graph API when `embedAvatars` is on for a format that renders them.
5. If API fails at any step, content script falls back to DOM scroll mode:
   - Scrolls through the message list to load history.
   - Auto-clicks "See More" buttons on truncated messages and "Show hidden history" buttons when present.
   - Extracts messages from the DOM as they appear.
   - Deduplicates by message ID across scroll passes.
6. Content script streams results to background in chunks through a `chrome.runtime.Port` (port name prefix `scrape-result:`, 30-second connection safety timeout). Streaming avoids Chrome's 64 MiB single-message limit.
7. Background builds output based on the selected formats:
   - Single text format (JSON/CSV/HTML/TXT) → one file.
   - HTML + `downloadImages` OR HTML + `avatarMode: 'files'` + `embedAvatars` → `HTML.zip` with `images/` and/or `avatars/` folders.
   - PDF → `.pdf` produced by `src/background/pdf.ts` (pdf-lib + HarfBuzz subsets + Twemoji rasterization).
   - Multiple formats → `bundle.zip` containing each format side by side, with shared `images/` / `avatars/` folders when relevant.
8. After a successful download (or cancel) background appends a `HistoryEntry` to `chrome.storage.local` under `teamsExporterHistory` and honors `afterExport: 'show' | 'manual'` to auto-show the file in its folder.
9. Background sends progress/status updates back to popup via `EXPORT_STATUS` and forwards scrape progress as `EXPORT_PROGRESS`. The current `activeExports` map is mirrored to `teamsExporterActiveExports` in storage so a reopened popup can rehydrate its state without waiting for a message round-trip.

## Key data types

- `Options`: user settings, stored in `chrome.storage.local` (`src/utils/options.ts`)
- `ExportMessage`: normalized message shape (`src/types/shared.ts`)
- `ScrapeOptions` / `BuildOptions`: scrape and output settings (`src/types/shared.ts`)

### Options fields

Authoritative definition: `Options` in `src/utils/options.ts`.

| Field | Type |
|-------|------|
| `lang` | `string?` |
| `startAt` / `endAt` | `string` (local input format) |
| `startAtISO` / `endAtISO` | `string` (ISO) |
| `exportTarget` | `'chat' \| 'team'` |
| `formats` | `OptionFormat[]` — non-empty; values from `'json' \| 'csv' \| 'html' \| 'txt' \| 'pdf'` |
| `includeReplies` | `boolean` |
| `includeReactions` | `boolean` |
| `includeSystem` | `boolean` |
| `embedAvatars` | `boolean` |
| `downloadImages` | `boolean` |
| `showHud` | `boolean` |
| `theme` | `'light' \| 'dark'` |
| `afterExport` | `'manual' \| 'show'` |
| `avatarMode` | `'inline' \| 'files'` (HTML only) |
| `pdfPageSize` | `'a4' \| 'letter'` |
| `pdfBodyFontSize` | `number` (clamped to `[8, 16]`) |
| `pdfShowPageNumbers` | `boolean` |
| `pdfIncludeAvatars` | `boolean` (per-format override) |
| `onboardingDismissed` | `boolean` |

Legacy storage with a singular `format` field is migrated into `formats: [format]` by `normalizeOptions` so existing installs keep their choice.

### ExportMessage fields

Authoritative definition: `ExportMessage` in `src/types/shared.ts`.

| Field | Type |
|-------|------|
| `id` | `string?` |
| `threadId` | `string \| null?` |
| `author` | `string?` |
| `timestamp` | `string?` |
| `text` | `string?` |
| `contentHtml` | `string?` (raw HTML, API mode only) |
| `messageType` | `string?` (e.g. `"Text"`, `"RichText/Html"`, `"Event/Call"`) |
| `edited` | `boolean?` |
| `system` | `boolean?` |
| `forwarded` | `ForwardContext?` (`originalAuthor`, `originalTimestamp`, `originalMessageId`, `originalThreadId`, `originalText`) |
| `importance` | `string?` (code checks for `"urgent"` and `"high"`) |
| `subject` | `string?` (channel post subject line) |
| `isOwn` | `boolean?` (author is the current Teams user — drives own-msg styling) |
| `avatar` | `string \| null?` |
| `avatarId` | `string?` |
| `avatarUrl` | `string?` |
| `reactions` | `Reaction[]` — `{ emoji, count, reactors?, self? }` |
| `attachments` | `Attachment[]` — `{ href?, label?, type?, size?, owner?, metaText?, dataUrl?, kind? }` |
| `tables` | `string[][][]` |
| `replyTo` | `ReplyContext \| null?` — `{ author, timestamp, text, id? }` |
| `mentions` | `Array<{ name, mri? }>` |
| `systemAttendees` | `string[]?` (participant names for call/meeting system messages) |
| `recordingDetails` | `RecordingDetails?` (populated for `RichText/Media_CallRecording`) |

## Status and persistence

- Background tracks active exports per tab in an in-memory `activeExports` map.
- Export-status phases the background emits as `EXPORT_STATUS`: `starting`, `scrape:start`, `scrape:complete`, `build`, `complete`, `empty`, `error`, `cancelling`, `cancelled`.
- Scrape-progress phases forwarded from the content script as `EXPORT_PROGRESS`: `api-fetch`, `scroll`, `extract`, `images`, `avatars`, `build`.
- Storage keys (all in `chrome.storage.local`, declared in `src/utils/options.ts`):
  - `teamsExporterOptions` — user settings.
  - `teamsExporterLastError` — most recent error, so the popup can recover context.
  - `teamsExporterHistory` — array of `HistoryEntry` for the History page.
  - `teamsExporterHistoryViewedAt` — timestamp used to compute the "new since last visit" count.
  - `teamsExporterLastPage` — `'main' | 'settings' | 'history'`, restored on popup reopen.
  - `teamsExporterActiveExports` — mirror of the background's `activeExports` map so the popup can rehydrate its export-button state on first render without waiting for the `GET_EXPORT_STATUS` round-trip.

## Supported Teams environments

- Commercial: `teams.microsoft.com`, `cloud.microsoft`, `teams.live.com`
- GCC High: `teams.microsoft.us`
- Microsoft Defender for Cloud Apps proxy: `.mcas.ms` suffix on any of the above

Full pattern list is in `src/utils/teams-urls.ts`.

## Browser notes

- Chrome/Edge build target: MV3 output (`.output/chrome-mv3/`)
- Firefox build target: MV2 output (`.output/firefox-mv2/`)
- Background code uses `browser.action` when available, falls back to `browser.browserAction` (Firefox MV2).
- Download URL creation (`textToDownloadUrl`, `binaryToDownloadUrl` in `src/background/builders.ts`) uses `URL.createObjectURL` (blob URLs) when available. In Chrome MV3 service workers where blob URLs are not supported, it falls back to base64 data URLs.

## Large export behavior

- With API mode, exports are typically fast regardless of message count. DOM scroll mode is slower because Teams loads history progressively.
- In DOM scroll mode, scrolling stops when no new messages appear for several consecutive passes. Default thresholds: 12 passes without loading indicators, 20 with loading indicators. Team channel mode uses higher thresholds (30/35). If Teams stops loading older history, only what was loaded gets exported.
- Inline image/audio data can use significant memory for media-heavy chats. This data is streamed in chunks to avoid Chrome's 64 MiB single-message limit.
- For HTML exports, if total embedded data exceeds 5 MB or the export has 500+ messages with `downloadImages` enabled, the build pipeline auto-upgrades to ZIP output to avoid V8 string length limits.

## PDF specifics

- Fonts: bundled Noto Sans Regular/Bold/SC are shipped in `src/public/fonts/`. Before embedding, each font is subsetted at runtime with HarfBuzz (via `harfbuzzjs` + `hb-subset.wasm` loaded through `src/background/font-subset.ts`) to exactly the codepoints the document uses. The GSUB/GPOS/GDEF/kern tables are dropped from the subset so pdf-lib's per-codepoint drawing matches per-codepoint measurement (prevents the "confl uence"-style ligature gap).
- Emoji: Twemoji SVGs are vendored from `@twemoji/svg` at postinstall into `src/public/twemoji/`. The PDF builder rasterizes only the emoji actually used in the document. Firefox MV2 background takes an `<img>` + OffscreenCanvas path; Chrome MV3 service worker uses `createImageBitmap(svgBlob, { resizeWidth, resizeHeight })`.
- URLs: detected `http(s)` substrings in message body, reply quotes, forwarded bodies, and single-line attachment labels become PDF `/Annot /Link` annotations.

## Known quirks

- Content script injection in `background.ts` uses a hardcoded filename `content.js`. If WXT changes its output naming, this will break.
- In HTML.zip and bundle.zip exports, images and GIFs are extracted to the `images/` folder. Audio (voice messages) stays inline as base64 in the HTML file.
- Video attachments store the video URL in the `owner` field and the thumbnail URL in `href`. The HTML builder reads `owner` to render the video link.
- `indexedDB.databases()` is not available in Firefox < 126. When unavailable, conversation ID lookup falls back to URL parsing and DOM attributes. A per-URL cache in `extractConversationId` remembers the last successful ID so repeat exports on the same tab don't re-hit IDB.
- Chat/channel context detection (`checkChatContext`) only matches `data-tid` attributes, never visible text — this is deliberate so non-English Teams UIs (French "Conversation", Japanese, etc.) work the same as English.
- Image fetching only allows Microsoft-owned domains (list in `src/content/attachments.ts`). Giphy GIF URLs from message content HTML are public and fetched without auth via the Giphy image parser in `src/content/api-converter.ts`.