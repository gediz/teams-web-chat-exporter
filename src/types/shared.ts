// One reactor of a given emoji. `name` is the resolved display name
// (falls back to MRI fragment if resolution failed). `avatarId` is set
// when we already have the reactor's profile photo — it points into
// meta.avatars so the HTML builder can render the dot stack without
// re-fetching. `self` marks the current user so the chip can highlight
// "You" differently. `uuid` is TRANSIENT — set during conversion so the
// content script can resolve a reactor's avatar by stable identity, then
// stripped before the message is serialized to the export.
export type ReactorInfo = { name: string; avatarId?: string; self?: boolean; uuid?: string };

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
  // Stable SharePoint file GUID (== the file's site-scoped listItemUniqueId).
  // Used as the download.aspx `UniqueId` when streaming renderable-markup
  // attachments to disk. Non-sensitive; kept in output. Renderers ignore it.
  itemid?: string;
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

// A merged-cell region inside a table. Field names match Handsontable's
// `mergeCells` format (https://handsontable.com/docs/javascript-data-grid/merge-cells/):
// `row`/`col` are the anchor (top-left) cell, `rowspan`/`colspan` the extents.
export type MergeRegion = { row: number; col: number; rowspan: number; colspan: number };

// A parsed table. `rows` is a dense grid where a spanned cell's value is
// REPEATED into every position it covers (so each row is self-contained, the
// pandas read_html / tidy-data convention); `merges` records each span once on
// its anchor (the Excel / Google Sheets / Handsontable convention) so the merge
// is losslessly recoverable. `headerRowCount` is the number of leading header
// rows (from <thead> or a leading all-<th> row).
export type TableData = {
  columns: number;
  headerRowCount: number;
  rows: string[][];
  merges: MergeRegion[];
};

// One block of a message body, kept in original order so text and tables
// interleave faithfully (a message with several tables and text between them
// stays readable). Built by src/content/table-model.ts in API mode.
export type BodyBlock =
  | { type: 'text'; text: string }
  | { type: 'table'; table: TableData };

export type ExportMessage = {
  id?: string;
  threadId?: string | null;
  author?: string;
  timestamp?: string;
  text?: string;
  contentHtml?: string;           // Raw HTML content (API mode only)
  messageType?: string;           // e.g. "Text", "RichText/Html", "Event/Call"
  edited?: boolean;
  // True for a deleted-for-everyone message kept as a "[message deleted]"
  // tombstone (text is set to the placeholder). Mirrors `edited`.
  deleted?: boolean;
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
  tables?: string[][][];          // legacy DOM-scrape tables (flat cell text, no spans)
  // Ordered text/table blocks parsed from contentHtml in API mode. When set,
  // table-aware renderers (HTML, PDF, TXT) use this instead of the flat `text`
  // so real tables render with structure. Derived into the JSON `tables` field
  // and stripped before JSON serialization (it is a render aid, not output).
  bodyBlocks?: BodyBlock[];
  replyTo?: ReplyContext | null;
  mentions?: Array<{ name: string; mri?: string }>;  // @mentions in the message
  systemAttendees?: string[];     // participant display names for call/meeting system messages
  recordingDetails?: RecordingDetails; // populated for RichText/Media_CallRecording messages
};

// One member of the chat roster, as surfaced in the export's participant
// list. Reflects CURRENT membership (the roster API), not every historical
// actor — someone who reacted/posted then left appears in the messages but
// not here. `name` is the best-resolved display name (Graph displayName, else
// the roster's friendlyName). `external` flags members named only via the
// roster's friendlyName cache (typically federated/guest users Graph can't read).
export type Participant = {
  name: string;
  mri?: string;
  external?: boolean;
};

export type ExportMeta = {
  title?: string | null;
  startAt?: string | null;
  endAt?: string | null;
  timeRange?: string | null;
  avatars?: Record<string, string>; // Map of avatarId -> base64 data URL
  // Current chat roster (resolved names), and the roster API's canonical
  // member count (may exceed participants.length when enumeration is capped).
  participants?: Participant[];
  memberCount?: number;
  // Teams conversation id for the scraped chat (when the content script could
  // resolve it). Used by the background to pin the persisted outcome snapshot
  // to a specific conversation, not just a tab.
  conversationId?: string;
  // Set by the content script when the scrape is known to be incomplete.
  // Drives the -PARTIAL filename suffix, in-file warning banner, and
  // history kind='partial'. Two known causes:
  //   'network'    — a NetworkError was caught during the API path or
  //                  during message hydration; some pages may not have
  //                  been fetched.
  //   'truncation' — DOM-scroll fallback ended with messages still
  //                  hydrating ("hydration pending after retries"
  //                  fired) — Teams' UI couldn't load the full content
  //                  of some visible messages.
  // When both apply, 'network' wins (higher-confidence root cause).
  partial?: { reason: 'network' | 'truncation' };
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
  // Fetch the full-resolution AMS image view (imgpsh_fullsize) instead of the
  // downscaled display view (imgo, ~1280px cap). Larger but the true original.
  fullResImages?: boolean;
  // Explicit conversation ID chosen by the popup's ConversationPicker.
  // When set, the API scraper skips DOM/IDB extraction entirely and
  // fetches this conversation directly. When absent, falls back to the
  // legacy "guess the currently-visible chat" path via extractConversationId.
  conversationId?: string | null;
  // Display name the picker resolved for the chosen conversation. The
  // content script uses this for meta.title (and thus the export filename)
  // when set — otherwise the DOM-derived chat title leaks in, which shows
  // the *currently visible* chat's name instead of the *selected* one.
  conversationTitle?: string | null;
  // Multi-chat bundle export sets this true so the content script
  // refuses to fall back to DOM scrolling when the API path fails. DOM
  // scroll always operates on whichever chat is *currently visible* in
  // the user's tab, which in bundle mode is whichever chat the user
  // happened to leave open — not the chat we're trying to export. The
  // background loop records that as a per-chat failure (FAILURES.txt
  // line) and moves on to the next chat.
  noDomFallback?: boolean;
};

export type BuildOptions = {
  // Selected output formats. 1 -> single file (or HTML.zip when html + downloadImages).
  // 2+ -> all built and packaged together as bundle.zip.
  formats?: ('json' | 'csv' | 'html' | 'txt' | 'pdf')[];
  saveAs?: boolean;
  embedAvatars?: boolean;
  downloadImages?: boolean;
  // Prepend each saved inline image's message share time (UTC) to its
  // images/ filename. Off by default; only affects the HTML/bundle image
  // folder, for cataloguing images pulled out of the export.
  imageFilenameDate?: boolean;
  // Stream non-image SharePoint document attachments to disk via
  // chrome.downloads (the "Files" toggle). Independent of the selected
  // formats; runs after the main export is saved.
  downloadFiles?: boolean;
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
  // Build-phase progress: how many messages have been written to the
  // output so far (PDF page emit / HTML chunk). Drives the segment-4
  // determinate fill on the export button.
  messagesBuilt?: number;
  messagesTotal?: number;
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
  // Multi-chat bundle context — present on every status broadcast that
  // happens inside a START_BUNDLE_EXPORT loop. Lets the popup show
  // "Chat 3 of 12: <name>" alongside the per-chat phase.
  bundleCurrentChat?: number;
  bundleTotalChats?: number;
  bundleChatName?: string;
  bundleSuccessCount?: number;
  bundleFailedCount?: number;
  // Attachment-download phase ('downloading-files'): how many document
  // attachments have been processed (resolved + handed to chrome.downloads)
  // out of the total found for this export. Drives the "Downloading files:
  // N / M" progress line. Only present while downloadFiles is on.
  filesDone?: number;
  filesTotal?: number;
  // Final attachment-download tally, attached to the 'complete' status when
  // downloadFiles is on. `saved` = downloaded to disk; `links` = kept as a
  // link (cross-tenant / consumer / unresolved); `failed` = resolve or
  // download error. Drives the post-export summary line.
  filesSummary?: { total: number; saved: number; links: number; failed: number };
};

// One row in the export history. Written on phase='complete'|'cancelled'
// and rendered by the HistoryPage. Metadata only — no message content,
// no avatars, no file bytes.
// A user-saved set of conversation ids for one-click batch re-selection.
// `convIds` may reference chats not present in the current account/list
// (resolved best-effort on apply); the picker reports how many matched.
export type SavedPreset = {
  id: string;          // crypto.randomUUID()
  name: string;
  convIds: string[];
  createdAt: number;
  updatedAt: number;
};

export type HistoryEntry = {
  // UUID; used as the React-style key when rendering and as the target
  // when the user removes a single entry.
  id: string;
  tabId: number;
  // 'success'   — the export produced the file the user asked for.
  // 'cancelled' — user stopped mid-run; nothing was saved (no file on disk).
  // 'failed'    — bundle export where 0 chats succeeded; FAILURES.txt was
  //               saved directly (no .zip wrapper since there's nothing
  //               to wrap). The file IS on disk, so Open/Show still work.
  // 'partial'   — file IS on disk, but the scrape detected incomplete data
  //               (NetworkError mid-export, or DOM-scroll truncation
  //               leaving messages without full content). The output gets
  //               a -PARTIAL filename suffix and an in-file warning banner.
  //               Distinct from 'cancelled' (cancelled writes no file at
  //               all) and from 'failed' (failed = nothing usable).
  kind: 'success' | 'cancelled' | 'failed' | 'partial';
  // For kind === 'partial', the dominant cause. 'network' = a NetworkError
  // was observed during scraping; 'truncation' = DOM scroll ended with
  // messages still hydrating. Used by the popup history label and helps
  // future bug reports distinguish the cause without re-asking the user.
  partialReason?: 'network' | 'truncation';
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

// Compact shape returned by listConversations() — just what the popup's
// picker needs to render an entry and hand an id back for export.
//
//   - 'chat'    : OneToOneChat (no topic — name derived from last sender)
//   - 'group'   : Chat with multiple members (has topic)
//   - 'meeting' : Meeting-scoped chat (has topic = meeting subject)
//   - 'channel' : TeamsStandardChannel (has topic = channel name)
//
// System-virtual conversations (StreamOfNotes / Mentions / CallLogs /
// Saved / Drafts / Notifications) are filtered out upstream — they
// never appear here. If classification fails for a new Teams product
// thread type we haven't seen yet, the conversation is also dropped.
export type ConversationKind = 'chat' | 'group' | 'meeting' | 'channel';
export type ConversationSummary = {
  id: string;              // e.g. "19:xxxx@unq.gbl.spaces"
  kind: ConversationKind;
  // Primary label. For 1:1 chats: the other person's display name
  // (resolved via Graph MRI lookup). For groups / meetings / channels:
  // `threadProperties.topic`.
  //
  // Empty string ("") is a sentinel meaning "could not resolve a name"
  // — the popup renders a locale-aware placeholder based on `kind` +
  // `isSelfChat` + `groupMembers`, keeping all user-facing text out of
  // the content script (which has no access to the popup's i18n).
  name: string;
  // Optional member-name list for unnamed group chats. When set and
  // non-empty, the popup formats it locale-aware with Intl.ListFormat
  // and appends a "+N more" remainder from `groupExtraMembers`.
  groupMembers?: string[];
  groupExtraMembers?: number;
  // Already-formatted secondary hint when locale-independent (meeting
  // date — Intl.DateTimeFormat picks up the user locale automatically;
  // channel's parent-team name — a display string from the tenant).
  // The popup may add locale-aware suffixes of its own (self-chat
  // label, "+N more" from groupExtraMembers) alongside this.
  subtitle?: string;
  // True for the "chat with yourself" thread. Popup adds the "(You)"
  // suffix + "Self-chat" subtitle in the active UI language.
  isSelfChat?: boolean;
  lastActivity?: string;   // ISO timestamp (for sort + UX hint)
  // Picker folder filter — ids of every Folder this conversation appears
  // in. A chat can be in multiple folders (Favorites overlaps with user
  // folders). Only Favorites + UserDefined folder ids land here; system
  // computations (MeetingChats, MutedChats, RecentChats, TeamsAndChannels,
  // QuickViews) are stripped upstream.
  folderIds?: string[];
};

// Compact shape for the picker's folder-filter rail. Mirrors the Teams
// `folders` IDB store but stripped down to just what the popup renders.
// Only Favorites + user-created folders are emitted; system-computed
// folders are filtered out at the source (see listConversationsFromIdb).
export type FolderSummary = {
  id: string;
  name: string;
  // 'favorites' for the system-curated star folder, 'user' for folders
  // the user created themselves. Distinguishes the favorite icon (★) from
  // the generic folder icon in the rail.
  kind: 'favorites' | 'user';
  // Number of conversations currently in this folder. Empty folders are
  // dropped before this point, so count is always >= 1 in practice.
  count: number;
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
