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
│   ├── builders.ts         # Export format builders
│   ├── download.ts         # Download handler
│   └── zip.ts              # ZIP archive creation
├── content/                # Content script modules
│   ├── scroll.ts           # Auto-scroll logic
│   ├── reactions.ts        # Reaction parsing
│   ├── replies.ts          # Reply parsing
│   ├── attachments.ts      # Attachment parsing
│   ├── text.ts             # Text extraction
│   └── title.ts            # Chat/channel title extraction
├── utils/                  # Shared utilities
│   ├── time.ts             # Time formatting
│   ├── text.ts             # Text processing
│   ├── dom.ts              # DOM helpers
│   ├── options.ts          # Settings persistence + Options type
│   ├── messaging.ts        # Chrome messaging
│   ├── badge.ts            # Badge updates
│   ├── messages.ts         # Message utilities
│   └── avatars.ts          # Avatar processing
├── types/                  # TypeScript types
│   ├── messaging.ts        # Message types
│   └── shared.ts           # Shared types (ExportMessage, etc.)
├── i18n/                   # Internationalization
│   └── locales/            # 24 translation files
└── public/                 # Static assets
    └── icons/              # Extension icons
```

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
│   Worker)   │  Builds export files
└──────┬──────┘
       │
       ↓ chrome.tabs.sendMessage
┌─────────────┐
│  Content    │  Scrapes Teams DOM
│  Script     │  Auto-scrolls to load history
│             │  Returns messages
└─────────────┘
```

## Runtime flow

1. Popup sends `START_EXPORT` (or `START_EXPORT_ZIP`) to background.
2. Background checks tab context and ensures content script is ready.
3. Background asks content script to scrape Teams (`SCRAPE_TEAMS`).
4. Content script streams scrape data to background in chunks through a runtime port.
5. Background builds output (`json/csv/html/txt`) and starts download.
6. Background sends progress/status back to popup with `EXPORT_STATUS`.

Streaming is used to avoid message-size limits on single runtime messages.

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
| `edited` | `boolean?` |
| `system` | `boolean?` |
| `avatar` | `string?` |
| `avatarUrl` | `string?` |
| `avatarId` | `string?` |
| `reactions` | `Array<{ emoji, count, reactors?, self? }>` |
| `attachments` | `Array<{ href?, label?, type?, size?, owner?, metaText?, dataUrl?, kind? }>` |
| `tables` | `string[][][]` |
| `replyTo` | `{ author, timestamp, text, id? }?` |

## Status and persistence

- Background tracks active exports per tab in memory.
- Popup receives status updates like `starting`, `scrape:start`, `scrape:complete`, `build`, `complete`, `empty`, and `error`.
- Options are saved under `teamsExporterOptions` in local extension storage.
- Last error is saved under `teamsExporterLastError` for popup recovery.

## Browser notes

- Chrome/Edge build target: MV3 output (`.output/chrome-mv3/`)
- Firefox build target: MV2 output (`.output/firefox-mv2/`)
- Background code uses `browser.*` when available, otherwise `chrome.*`
- Firefox MV2 uses `browserAction`; Chrome MV3 uses `action`
- Downloads use browser-specific URL creation logic in `src/background/builders.ts` (`textToDownloadUrl`, `binaryToDownloadUrl`)
- Chrome MV3 service worker lacks `URL.createObjectURL`, so downloads use base64 data URLs
- Firefox uses blob URLs for downloads (no data URL size issues)

## Large export behavior

- Exporting 10,000+ messages spanning a year or more can take 30–60 minutes. Teams gets slower loading older history.
- Scroll stops when no new messages appear for several consecutive passes (default: 12–20 passes depending on loading signals). If Teams has more history but stops loading it, only what was loaded gets exported.
- Inline image data (attachment previews) can use 100MB+ of memory for image-heavy chats. This data is streamed in chunks to avoid Chrome's 64MiB single-message limit.
- Teams may cap rendered messages at ~750 in the DOM. The scroll engine works around this with deduplication and incremental aggregation.

## Known quirks

- Content script injection in background.ts uses a hardcoded filename `content.js`. If WXT changes its output naming, this will break.
