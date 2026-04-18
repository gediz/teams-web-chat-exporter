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
│   │   ├── components/     # UI components
│   │   └── popup.css       # Styles
│   ├── background.ts       # Service worker
│   └── content.ts          # Content script (scraper)
├── background/             # Background script modules
│   ├── builders.ts         # Export format builders (CSV, HTML) + download URL helpers
│   ├── download.ts         # Export build pipeline + download logic
│   └── zip.ts              # ZIP archive creation
├── content/                # Content script modules
│   ├── api-client.ts       # Teams Chat Service API client (token, discovery, fetch)
│   ├── api-converter.ts    # API response to ExportMessage conversion
│   ├── scroll.ts           # DOM auto-scroll message collection
│   ├── reactions.ts        # Reaction parsing
│   ├── replies.ts          # Reply context parsing
│   ├── attachments.ts      # Attachment extraction (files, images, links, previews)
│   ├── text.ts             # Text extraction with emoji/mention handling
│   └── title.ts            # Chat/channel title extraction
├── utils/                  # Shared utilities
│   ├── time.ts             # Time formatting and parsing
│   ├── text.ts             # Text processing (normalize, score, escape)
│   ├── dom.ts              # DOM query helpers ($, $$)
│   ├── options.ts          # Settings persistence + Options type
│   ├── messaging.ts        # Runtime message wrapper
│   ├── badge.ts            # Extension icon badge manager
│   ├── messages.ts         # Day dividers, filename sanitization, merge helpers
│   ├── avatars.ts          # Avatar ID extraction from URLs
│   └── teams-urls.ts       # Teams URL match patterns (manifest + runtime)
├── types/                  # TypeScript types
│   ├── messaging.ts        # IPC message types (popup/background/content)
│   └── shared.ts           # Domain types (ExportMessage, ScrapeOptions, etc.)
├── i18n/                   # Internationalization
│   └── locales/            # 24 translation files
└── public/                 # Static assets
    └── icons/              # Extension icons (16, 32, 48, 128px)
```

## Component communication

```
┌─────────────┐
│   Popup     │  User configures export
│  (Svelte)   │  Sends START_EXPORT or START_EXPORT_ZIP
└──────┬──────┘
       │
       ↓ chrome.runtime.sendMessage
┌─────────────┐
│ Background  │  Orchestrates export
│  (Service   │  Sends SCRAPE_TEAMS to content
│   Worker)   │  Builds export files
└──────┬──────┘
       │
       ↓ chrome.tabs.sendMessage
┌─────────────┐
│  Content    │  1. Tries API fetch (fast, complete)
│  Script     │  2. Falls back to DOM scroll if API fails
│             │  Streams messages back via runtime port
└─────────────┘
```

## Runtime flow

1. Popup sends `START_EXPORT` (or `START_EXPORT_ZIP`) to background.
2. Background checks tab context and ensures content script is injected.
3. Background asks content script to scrape Teams (`SCRAPE_TEAMS`).
4. Content script tries API-based fetch first:
   - Reads MSAL tokens from localStorage (IC3, Skype, Graph). Supports both commercial and GCC High endpoints.
   - Discovers the chat service endpoint via the Teams authz API.
   - Fetches messages page by page from the chat service `/messages` endpoint (max 500 pages, retries on 429 with exponential backoff).
   - Converts API response objects to ExportMessage format.
   - Fetches inline image/audio data from AMS URLs (max 5 MB per file, 6 concurrent fetches, retries 408/410/429/5xx with exponential backoff).
   - Fetches avatar photos via Microsoft Graph API.
5. If API fails at any step, content script falls back to DOM scroll mode:
   - Scrolls through the message list to load history.
   - Auto-clicks "See More" buttons on truncated messages and "Show hidden history" buttons when present.
   - Extracts messages from the DOM as they appear.
   - Deduplicates by message ID across scroll passes.
6. Content script streams results to background in chunks through a runtime port (30-second connection timeout). Streaming avoids Chrome's 64 MiB single-message limit.
7. Background builds output (JSON, CSV, HTML, or TXT) and triggers download.
8. Background sends progress/status updates back to popup via `EXPORT_STATUS`.

## Key data types

- `Options`: user settings, stored in `chrome.storage.local` (`src/utils/options.ts`)
- `ExportMessage`: normalized message shape (`src/types/shared.ts`)
- `ScrapeOptions` / `BuildOptions`: scrape and output settings (`src/types/shared.ts`)

### Options fields

| Field | Type |
|-------|------|
| `lang` | `string?` |
| `startAt` / `endAt` | `string` (local input format) |
| `startAtISO` / `endAtISO` | `string` (ISO) |
| `exportTarget` | `'chat' \| 'team'` |
| `format` | `'json' \| 'csv' \| 'html' \| 'txt'` |
| `includeReplies` | `boolean` |
| `includeReactions` | `boolean` |
| `includeSystem` | `boolean` |
| `embedAvatars` | `boolean` |
| `downloadImages` | `boolean` |
| `showHud` | `boolean` |
| `theme` | `'light' \| 'dark'` |

### ExportMessage fields

| Field | Type |
|-------|------|
| `id` | `string?` |
| `threadId` | `string?` |
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
| `avatar` | `string?` |
| `avatarUrl` | `string?` |
| `avatarId` | `string?` |
| `reactions` | `Array<{ emoji, count, reactors?, self? }>` |
| `attachments` | `Array<{ href?, label?, type?, size?, owner?, metaText?, dataUrl?, kind? }>` |
| `tables` | `string[][][]` |
| `replyTo` | `{ author, timestamp, text, id? }?` |
| `mentions` | `Array<{ name, mri? }>` |

## Status and persistence

- Background tracks active exports per tab in an `activeExports` map.
- Export phases sent to popup: `starting`, `scrape:start`, `scrape:complete`, `build`, `complete`, `empty`, `error`.
- Options are saved under `teamsExporterOptions` in `chrome.storage.local`.
- Last error is saved under `teamsExporterLastError` for popup recovery.

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

## Known quirks

- Content script injection in `background.ts` uses a hardcoded filename `content.js`. If WXT changes its output naming, this will break.
- In ZIP exports, images and GIFs are extracted to the `images/` folder. Audio (voice messages) stays inline as base64 in the HTML file.
- Video attachments store the video URL in the `owner` field and the thumbnail URL in `href`. The HTML builder reads `owner` to render the video link.
- `indexedDB.databases()` is not available in Firefox < 126. When unavailable, conversation ID lookup falls back to URL parsing and DOM attributes.
- Image fetching only allows Microsoft-owned domains (list in `src/content/attachments.ts`). Giphy GIFs are fetched without auth.