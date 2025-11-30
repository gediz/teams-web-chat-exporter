# Architecture

Technical overview of the Teams Chat Exporter extension.

## Technology Stack

- **Framework**: WXT 0.19+ (Vite-based)
- **Popup UI**: Svelte 5 + TypeScript
- **Background/Content**: TypeScript
- **Styling**: CSS (scoped in Svelte components)
- **I18N**: Custom implementation (14 languages)

## Project Structure

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
│   └── download.ts         # Download handler
├── content/                # Content script modules
│   ├── scroll.ts           # Auto-scroll logic
│   ├── reactions.ts        # Reaction parsing
│   ├── replies.ts          # Reply parsing
│   ├── attachments.ts      # Attachment parsing
│   └── text.ts             # Text extraction
├── utils/                  # Shared utilities
│   ├── time.ts             # Time formatting
│   ├── text.ts             # Text processing
│   ├── dom.ts              # DOM helpers
│   ├── options.ts          # Settings persistence
│   ├── messaging.ts        # Chrome messaging
│   ├── badge.ts            # Badge updates
│   └── messages.ts         # Message utilities
├── types/                  # TypeScript types
│   ├── messaging.ts        # Message types
│   └── shared.ts           # Shared types
├── i18n/                   # Internationalization
│   └── locales/            # Translation files
└── public/                 # Public assets
    └── icons/              # Extension icons
```

## Component Communication

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

## Message Flow

### Export Process

1. **User clicks "Export"** in popup
2. **Popup** sends `START_EXPORT` to background with options
3. **Background** sends `SCRAPE_TEAMS` to content script
4. **Content script**:
   - Validates Teams chat is open
   - Auto-scrolls to load all messages
   - Extracts messages, reactions, replies
   - Returns aggregated data
5. **Background**:
   - Builds export file (JSON/CSV/HTML/Text)
   - Triggers download via Chrome Downloads API
   - Sends `EXPORT_STATUS` updates to popup
6. **Popup** displays progress and completion

## Data Structures

### Options
```typescript
{
  lang: string,              // UI language
  startAt: string,           // Date range start (YYYY-MM-DD)
  endAt: string,             // Date range end (YYYY-MM-DD)
  format: 'json' | 'csv' | 'html' | 'txt',
  includeReplies: boolean,
  includeReactions: boolean,
  includeSystem: boolean,
  embedAvatars: boolean,     // HTML only
  showHud: boolean,          // In-page progress overlay
  theme: 'light' | 'dark'
}
```

### Message
```typescript
{
  id: string,
  author: string,
  timestamp: string,
  text: string,
  edited: boolean,
  avatar: string | null,
  reactions: Array<{ emoji: string, count: number }>,
  attachments: Array<{ href: string, label: string, type?: string }>,
  replyTo: { author: string, text: string } | null,
  system: boolean
}
```

## Browser Compatibility

### Chrome/Edge (Manifest V3)
- Uses `chrome.*` namespace
- Service worker background
- Native Downloads API (data URLs)

### Firefox (Manifest V2)
- Uses `browser.*` namespace (polyfilled)
- Background page (not service worker)
- Blob URL fallback for downloads
- Badge uses `browserAction` instead of `action`

### WXT Polyfills
WXT automatically handles cross-browser differences:
- API namespace (`chrome` vs `browser`)
- Manifest version conversion
- Storage API compatibility
- Messaging Promise handling

## Build Pipeline

1. **Development**: `npm run dev` / `npm run dev:firefox`
   - Vite dev server with HMR
   - Source maps enabled
   - No minification

2. **Production**: `npm run build` / `npm run build:firefox`
   - TypeScript compilation
   - Svelte component bundling
   - Tree-shaking and minification
   - Source maps (external)

3. **Output**:
   - Chrome: `.output/chrome-mv3/`
   - Firefox: `.output/firefox-mv2/`

## Storage

Uses `chrome.storage.local` for:
- User preferences (options)
- Last error message (for popup reconnection)

Data is persisted across sessions but not synced across devices.

## Performance Considerations

- **Scraping**: Sequential DOM parsing (single-threaded)
- **Memory**: ~5-10 MB for 1000 messages
- **Large exports**: HTML with embedded avatars can exceed 50 MB
- **Optimization**: Blob URLs in Firefox prevent data URL size limits
