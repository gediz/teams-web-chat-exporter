// One reactor of a given emoji. `name` is the resolved display name
// (falls back to MRI fragment if resolution failed). `avatarId` is set
// when we already have the reactor's profile photo — it points into
// meta.avatars so the HTML builder can render the dot stack without
// re-fetching. `self` marks the current user so the chip can highlight
// "You" differently.
export type ReactorInfo = { name: string; avatarId?: string; self?: boolean };

// A single emoji reaction on a message. `reactors` now carries rich
// per-reactor info for the v9 reactor chip (avatar stack + popover).
// Legacy data on disk that stored `reactors: string[]` is still
// readable — the HTML builder defensively handles both shapes.
export type Reaction = { emoji: string; count: number; reactors?: ReactorInfo[]; self?: boolean };

export type Attachment = {
  href?: string;
  label?: string;
  type?: string | null;
  size?: string | null;
  owner?: string | null;
  metaText?: string | null;
  dataUrl?: string;
  kind?: 'preview';
};

export type ReplyContext = {
  author: string;
  timestamp: string;
  text: string;
  id?: string;
};

export type ForwardContext = {
  originalAuthor?: string;      // Who wrote the original message
  originalTimestamp?: string;    // When the original was sent
  originalMessageId?: string;   // ID of the original message
  originalThreadId?: string;    // Conversation the original came from
  originalText?: string;        // Content of the original forwarded message
};

/** Aggregated metadata for a meeting recording artifact. */
export type RecordingDetails = {
  title?: string;
  callId?: string;
  amsDocumentId?: string;
  thumbnailUrl?: string;        // initially the asyncgw URL; becomes data: URL after fetch
  playUrl?: string;             // SharePoint share link
  transcriptUrl?: string;       // asyncgw transcript view (needs Bearer)
  videoUrl?: string;            // asyncgw video view (needs Bearer)
  rosterUrl?: string;           // asyncgw rosterevents view (needs Bearer)
  meetingStart?: string;        // ISO — paired Meeting-started event composetime
  meetingEnd?: string;          // ISO — paired Meeting-ended event composetime
  durationSec?: number;         // duration in seconds (from paired meeting-ended <duration>)
  organizerUpn?: string;        // organizer email (from paired meeting events)
  meetingType?: string;         // Recurring / Scheduled (from paired meeting events)
  meetingSubject?: string;      // subject (from paired meeting events)
  attendees?: string[];         // attendees (from paired meeting-ended event)
};

export type ExportMessage = {
  id?: string;
  threadId?: string | null;
  author?: string;
  timestamp?: string;
  text?: string;
  contentHtml?: string;           // Raw HTML content (API mode only)
  messageType?: string;           // e.g. "Text", "RichText/Html", "Event/Call"
  edited?: boolean;
  system?: boolean;
  forwarded?: ForwardContext;     // Forward context with original author info
  importance?: string;            // "normal", "urgent", etc.
  subject?: string;               // Subject line (channel posts)
  // True when this message was authored by the current Teams user. The
  // HTML builder styles own-messages with a distinct accent so scanning
  // a conversation visually separates "me" from "them" (issue #20).
  isOwn?: boolean;
  avatar?: string | null;
  avatarId?: string;
  avatarUrl?: string;
  reactions?: Reaction[];
  attachments?: Attachment[];
  tables?: string[][][];
  replyTo?: ReplyContext | null;
  mentions?: Array<{ name: string; mri?: string }>;  // @mentions in the message
  systemAttendees?: string[];     // participant display names for call/meeting system messages
  recordingDetails?: RecordingDetails; // populated for RichText/Media_CallRecording messages
};

export type ExportMeta = {
  title?: string | null;
  startAt?: string | null;
  endAt?: string | null;
  timeRange?: string | null;
  avatars?: Record<string, string>; // Map of avatarId -> base64 data URL
  // Teams conversation id for the scraped chat (when the content script could
  // resolve it). Used by the background to pin the persisted outcome snapshot
  // to a specific conversation, not just a tab.
  conversationId?: string;
  [key: string]: unknown;
};

export type ScrapeOptions = {
  startAt?: string | null;
  endAt?: string | null;
  startAtISO?: string | null;
  endAtISO?: string | null;
  includeReplies?: boolean;
  includeReactions?: boolean;
  includeSystem?: boolean;
  showHud?: boolean;
  exportTarget?: 'chat' | 'team';
  // Downstream build options surfaced to the scraper so it can skip
  // expensive fetches none of the selected formats will actually use:
  //   txt/csv never render images or avatars
  //   json/html without embedAvatars don't need avatar photos
  //   html without downloadImages doesn't need inline image blobs
  // The union across `formats` decides — fetch when ANY selected format wants it.
  formats?: ('json' | 'csv' | 'html' | 'txt' | 'pdf')[];
  embedAvatars?: boolean;
  downloadImages?: boolean;
};

export type BuildOptions = {
  // Selected output formats. 1 -> single file (or HTML.zip when html + downloadImages).
  // 2+ -> all built and packaged together as bundle.zip.
  formats?: ('json' | 'csv' | 'html' | 'txt' | 'pdf')[];
  saveAs?: boolean;
  embedAvatars?: boolean;
  downloadImages?: boolean;
  // 'files' forces HTML-only exports to zip so the avatars/ folder can
  // sit alongside the HTML. 'inline' keeps avatars as base64 in the
  // HTML (single self-contained file).
  avatarMode?: 'inline' | 'files';
  // PDF layout preferences — only read when 'pdf' is in formats.
  pdfPageSize?: 'a4' | 'letter';
  pdfBodyFontSize?: number;
  pdfShowPageNumbers?: boolean;
  pdfIncludeAvatars?: boolean;
  // User's "After export" preference. Drives auto-open / auto-show on success.
  afterExport?: 'manual' | 'show';
};

export type ScrapeResult = {
  messages: ExportMessage[];
  meta?: ExportMeta;
};

export type ExportStatusPayload = {
  tabId?: number;
  phase?: string;
  messages?: number;
  messagesExtracted?: number;
  filename?: string;
  // Set on the 'complete' status so the popup can wire the Open/Show action
  // buttons to chrome.downloads.open(id) / .show(id).
  downloadId?: number;
  // The user's "After export" preference, forwarded to the popup
  // primarily for UI state (it no longer gates auto-open — the only
  // auto-action is 'show', handled by the service worker because
  // downloads.show() doesn't require a user gesture).
  afterExport?: 'manual' | 'show';
  error?: string;
  message?: string;
  startedAt?: number | string;
};

// What the popup shows in the ExportButton's right zone after an export
// finishes. For 'success', the action buttons render when downloadId is set;
// otherwise the sticky "primary/secondary" tile falls back.
export type ExportOutcome = {
  kind: 'success' | 'cancelled';
  primary: string;
  secondary: string;
  downloadId?: number;
  // When true, the post-export actions swap roles: 'Show in folder' becomes
  // the primary action (safer across platforms for .zip).
  isZip?: boolean;
};

// One row in the export history. Written on phase='complete'|'cancelled'
// and rendered by the HistoryPage. Metadata only — no message content,
// no avatars, no file bytes.
export type HistoryEntry = {
  // UUID; used as the React-style key when rendering and as the target
  // when the user removes a single entry.
  id: string;
  tabId: number;
  kind: 'success' | 'cancelled';
  // Teams conversation id (from the content script's resolver). Optional;
  // present helps verify "this entry belongs to the chat I'm looking at."
  convId?: string;
  // chrome.downloads id — drives the Open / Show actions.
  downloadId?: number;
  filename?: string;
  // Chat / channel display name from the scrape meta. Shown in the row's
  // secondary line so users can tell entries apart at a glance.
  title?: string;
  // Selected build format(s). `formats` is the source of truth for new
  // entries; `format` is preserved for back-compat with rows already on
  // disk (pre-multi-format release). HistoryPage prefers `formats` then
  // falls back to `format`.
  formats?: ('json' | 'csv' | 'html' | 'txt' | 'pdf')[];
  format?: 'json' | 'csv' | 'html' | 'txt' | 'pdf';
  isZip?: boolean;
  messageCount?: number;
  elapsedMs?: number;
  savedAt: number;        // Date.now() when the entry was written
  // Tri-state file-existence record:
  //   undefined  - never verified (or unknown — render as available)
  //   true       - confirmed present on disk
  //   false      - confirmed missing (Open failed, or onChanged exists=false)
  // Persisted so missing state survives popup close/reopen, especially
  // important on Firefox where downloads.search() returns stale 'exists:
  // true' for files deleted outside the browser.
  fileExists?: boolean;
};

export type ActiveExportInfo = {
  startedAt?: number;
  lastStatus?: ExportStatusPayload;
  phase?: string;
  completedAt?: number;
};

export type AggregatedItem = {
  message?: ExportMessage;
  orderKey: number;
  tsMs: number | null;
  kind: 'message' | 'system-control' | 'day-divider';
  timeLabel?: string;
  anchorTs?: number;
};

export type OrderContext = {
  lastTimeMs: number | null;
  yearHint: number | null;
  seqBase: number;
  lastAuthor: string | null;
  lastId: string | null;
  seq: number;
  systemCursor: number;
};
