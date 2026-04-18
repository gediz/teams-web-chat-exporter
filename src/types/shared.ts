export type Reaction = { emoji: string; count: number; reactors?: string[]; self?: boolean };

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
};

export type BuildOptions = {
  format?: 'json' | 'csv' | 'html' | 'txt';
  saveAs?: boolean;
  embedAvatars?: boolean;
  downloadImages?: boolean;
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
  error?: string;
  message?: string;
  startedAt?: number | string;
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
