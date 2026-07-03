import type { BuildOptions, ConversationSummary, ExportMessage, ExportStatusPayload, FolderSummary, ScrapeOptions } from './shared';

export type PingSWRequest = { type: 'PING_SW' };
export type PingSWResponse = { ok: boolean; now: number };

export type GetExportStatusRequest = { type: 'GET_EXPORT_STATUS'; tabId?: number | null };
export type GetExportStatusResponse = { active: boolean; info?: { startedAt?: number | string; lastStatus?: ExportStatusPayload } };

export type StartExportRequest = {
  type: 'START_EXPORT';
  data: {
    tabId?: number | null;
    scrapeOptions: ScrapeOptions;
    buildOptions: BuildOptions;
  };
};
export type StartExportResponse = {
  ok?: boolean;
  filename?: string;
  downloadId?: number;
  messages?: number;
  cancelled?: boolean;
  error?: string;
  code?: string;
};

// Multi-chat bundle export. The popup sends this when the picker has 2+
// conversations selected. The SW loops over each id, runs the standard
// scrape pipeline, packs everything into a single outer zip with per-
// chat subfolders + FAILURES.txt for any chat that errored. scrapeOptions
// MUST NOT carry `conversationId` / `conversationTitle` here — the SW
// injects those per-iteration from the `conversations` array.
export type StartBundleExportRequest = {
  type: 'START_BUNDLE_EXPORT';
  data: {
    tabId?: number | null;
    conversations: Array<{ id: string; title: string }>;
    scrapeOptions: ScrapeOptions;
    buildOptions: BuildOptions;
  };
};
export type StartBundleExportResponse = {
  ok?: boolean;
  filename?: string;
  downloadId?: number;
  totalChats?: number;
  successChats?: number;
  failedChats?: number;
  cancelled?: boolean;
  error?: string;
  code?: string;
};

// Low-level "I have a payload, write a file" entry. Single-format only —
// multi-format end-user exports go through START_EXPORT (which routes to
// buildAndDownloadBundle in the SW). If a future caller needs bundle
// output via this path, extend the handler too, not just this type.
export type BuildAndDownloadRequest = {
  type: 'BUILD_AND_DOWNLOAD';
  data: {
    messages?: ExportMessage[];
    meta?: Record<string, unknown>;
    format?: 'json' | 'csv' | 'html' | 'txt' | 'pdf';
    saveAs?: boolean;
    embedAvatars?: boolean;
    downloadImages?: boolean;
  };
};

// Background-mediated fetch for credentialed cross-origin requests (Firefox content
// scripts can't send page cookies cross-origin, but background scripts can with
// matching host_permissions).
export type FetchBlobRequest = {
  type: 'FETCH_BLOB';
  url: string;
  bearerToken?: string;
  maxBytes?: number;
  minBytes?: number;
};
// Direct upstream fetch — used by the image-fetch-fallback feature when
// Teams' urlp/AMS proxy returns a permanent-shaped failure (4xx). No
// auth headers; relies on the user having granted <all_urls> via the
// "Image fetch fallback" toggle in Settings. Same response shape as
// FETCH_BLOB so the call site can treat them uniformly.
export type FetchBlobDirectRequest = {
  type: 'FETCH_BLOB_DIRECT';
  url: string;
  maxBytes?: number;
  minBytes?: number;
};
// Image-fetch-fallback feature gate. Content scripts can't reliably
// access the permissions API in Firefox MV2 — this message routes the
// check to the background, which has full API access. Returns the AND
// of (a) user toggled ON in Settings AND (b) <all_urls> permission is
// currently granted. Content asks once at scrape start to decide
// whether to bother sending FETCH_BLOB_DIRECT on proxy failures.
export type FallbackStatusRequest = { type: 'FALLBACK_STATUS' };
export type FallbackStatusResponse = { enabled: boolean };
export type StopExportRequest = { type: 'STOP_EXPORT'; tabId?: number | null };
export type StopExportResponse = { ok: boolean; error?: string };

// NOTE: Open/Show actions on a finished download are handled by the popup
// calling chrome.downloads.open / .show directly. Routing through the
// service worker broke the user-activation requirement on downloads.open
// in MV3. See openSavedDownload in App.svelte.

// Ask the content script to fetch the user's conversation list via the
// Teams chat service API. Popup uses the result to populate its picker.
// Returns either the sorted list or a failure reason — typically "no
// valid IC3 token" (user hasn't authenticated to Teams in this tab) or
// a network/HTTP error from the chat service itself.
export type ListConversationsRequest = { type: 'LIST_CONVERSATIONS'; tabId?: number | null };
export type ListConversationsResponse =
  | { ok: true; conversations: ConversationSummary[]; folders?: FolderSummary[] }
  | { ok: false; error: string };

// Fast IDB-only variant: returns ~instantly with topic + last-sender
// names but without Graph / roster enrichment. Used to paint the
// picker on cold load without waiting for the full pipeline.
export type ListConversationsQuickRequest = { type: 'LIST_CONVERSATIONS_QUICK'; tabId?: number | null };
export type ListConversationsQuickResponse = ListConversationsResponse;

// Diagnostics, Layer 1 (passive snapshot). Popup pulls the background's
// log tail. Content script returns its log tail plus an IDB shape probe
// (database and store names, row counts only).
// Combined log buffer entry. Lives in BG, populated by:
//   - BG's own console wrap (src: 'bg')
//   - Content scripts forwarding their captures (src: 'content')
// The popup pulls the merged buffer and the report renderer splits
// back into LOGS_BACKGROUND / LOGS_CONTENT by src.
export type DiagLogEntry = { src: 'bg' | 'content'; ts: number; level: string; line: string };

export type GetDiagnosticsBgRequest = { type: 'GET_DIAGNOSTICS_BG' };
export type GetDiagnosticsBgResponse = {
  entries: DiagLogEntry[];
  // Bytes currently on disk. Null when persistence is on but the
  // browser does not implement getBytesInUse (older Firefox WebExt
  // polyfill). The UI renders 'unknown size' for null instead of a
  // misleading '0 B'.
  bytesUsed: number | null;
  persistEnabled: boolean;
  // Captures the most recent persistence write failure (quota
  // exceeded, profile corrupt, etc.). Surfaced so a user with the
  // toggle ON but bytesUsed === 0 sees a clear hint, instead of
  // wondering why nothing is being saved.
  lastFlushError: { ts: number; reason: string } | null;
};

// Content script forwards its captures to BG as small batched arrays.
// Fire-and-forget; BG wakes on receive and appends.
export type DiagLogForwardRequest = {
  type: 'DIAG_LOG_FORWARD';
  entries: { ts: number; level: string; line: string }[];
};

// Wipes both the in-memory log buffer and the persisted storage key.
// Fired from the Diagnostics page's "Clear logs" button.
export type ClearDiagnosticsLogsRequest = { type: 'CLEAR_DIAGNOSTICS_LOGS' };
export type ClearDiagnosticsLogsResponse = { ok: boolean };

// Toggles whether BG persists the combined log buffer to
// chrome.storage.local. Off by default (zero disk footprint for the
// 99.9% of users who never engage with diagnostics).
export type SetDiagLogPersistRequest = { type: 'SET_DIAG_LOG_PERSIST'; enabled: boolean };
export type SetDiagLogPersistResponse = { ok: boolean };

// Toggles whether the console-only [export-stats] line includes per-chat
// detail (titles/convIds/per-chat stage split). Off by default; aggregates
// only. Opt-in from the Diagnostics page for local perf debugging.
export type SetDiagVerboseStatsRequest = { type: 'SET_DIAG_VERBOSE_STATS'; enabled: boolean };
export type SetDiagVerboseStatsResponse = { ok: boolean };

export type DiagnosticsProbeResult = {
  name: string;
  status: 'pass' | 'fail' | 'skipped';
  detail?: string;
  ms: number;
};
export type RunProbesBgRequest = { type: 'RUN_PROBES_BG' };
export type RunProbesBgResponse =
  | { ok: true; results: DiagnosticsProbeResult[]; totalMs: number }
  | { ok: false; reason: string };
export type RunProbesContentRequest = { type: 'RUN_PROBES_CONTENT' };
export type RunProbesContentResponse =
  | { ok: true; results: DiagnosticsProbeResult[]; totalMs: number }
  | { ok: false; reason: string };

// File-access probe (Diagnostics page). Two hops:
//   popup -> content:  RESOLVE_SHARE_FILE resolves a SharePoint file URL via
//                      the shares API (must run with the Teams page origin).
//   popup -> bg:       PROBE_FILE_DOWNLOAD hands the resolved pre-auth
//                      downloadUrl to chrome.downloads and reports the settled
//                      outcome. The downloadUrl embeds a short-lived token —
//                      transported over internal messaging only, never logged.
// Diagnostics: dump the raw field names of the open chat's file attachments
// (values never leave the page), to check for a sharing-link field.
export type DumpFileFieldsRequest = { type: 'DUMP_FILE_FIELDS' };
export type DumpFileFieldsResponse =
  | { ok: true; messages: number; fileRecords: number; keys: string[]; linkFields: string[] }
  | { ok: false; error: string };
// Salvage tool (Diagnostics page): resolve a list of file links at once.
export type BatchResolveHrefsRequest = { type: 'BATCH_RESOLVE_HREFS'; hrefs: string[] };
export type BatchResolveHrefsResponse = {
  ok: boolean;
  error?: string;
  results?: Array<{ href: string; name: string; ok: boolean; downloadUrl?: string; blocksDownload?: boolean; error?: string; via?: string }>;
};
export type ResolveShareFileRequest = { type: 'RESOLVE_SHARE_FILE'; href: string };
export type ResolveShareFileResponse = {
  ok: boolean;
  status: number;
  name?: string;
  mimeType?: string;
  downloadUrl?: string;
  blocksDownload?: boolean;
  allowEdit?: boolean;
  readOnly?: boolean;
  itemId?: string;
  error?: string;
  // Whether the resolve used the matched file's sharing link or fell back to
  // the pasted raw URL; matchedName is the file record we matched.
  via?: 'shareUrl' | 'rawUrl';
  matchedName?: string;
  // Credentials mode the page-world helper used (diagnostic).
  mode?: string;
};
// Files-phase resolver RPC (background -> content): resolve a file's sharing
// link to a pre-authenticated download URL. The downloadUrl is short-lived and
// must never be logged or persisted.
export type DownloadResolveShareRequest = { type: 'DOWNLOAD_RESOLVE_SHARE'; shareUrl: string };
export type DownloadResolveShareResponse = { ok: boolean; downloadUrl?: string; blocksDownload?: boolean; error?: string };
// Salvage tool runs in the background (survives the popup closing on file pick).
export type SalvageLinksRequest = { type: 'SALVAGE_LINKS'; hrefs: string[]; tabId: number };
export type SalvageLinksResponse = { ok: boolean; started?: number; error?: string };
export type ProbeFileDownloadRequest = { type: 'PROBE_FILE_DOWNLOAD'; url: string; name?: string };
export type ProbeFileDownloadResponse =
  | { ok: true; outcome: string; mime?: string; bytes?: number; filename?: string }
  | { ok: false; error: string };

export type GetDiagnosticsContentRequest = { type: 'GET_DIAGNOSTICS_CONTENT' };
export type DiagnosticsIdbStore = { name: string; count: number; error?: string };
export type DiagnosticsIdbDatabase = {
  name: string;
  version: number;
  status: 'opened' | 'blocked' | 'error';
  reason?: string;
  stores: DiagnosticsIdbStore[];
};
// Content no longer owns log entries (BG holds the merged buffer
// populated via DIAG_LOG_FORWARD). This message now returns only the
// IDB shape probe and per-content forwarding stats so a silent
// forwarding failure surfaces in the report.
export type DiagnosticsForwardingStats = {
  lostBatches: number;
  lostEntries: number;
  lastError: string | null;
};
export type GetDiagnosticsContentResponse = {
  idbShape:
    | { available: true; databases: DiagnosticsIdbDatabase[] }
    | { available: false; reason: string };
  forwardingStats: DiagnosticsForwardingStats;
};

export type ScrapeProgressMessage = {
  type: 'SCRAPE_PROGRESS';
  payload: {
    phase?: string;
    passes?: number;
    messagesVisible?: number;
    aggregated?: number;
    seen?: number;
    filteredSeen?: number;
    [key: string]: unknown;
  };
};
export type ExportStatusMessage = { type: 'EXPORT_STATUS' } & ExportStatusPayload;
export type ExportStatusUpdateMessage = { type: 'EXPORT_STATUS_UPDATE'; payload: ExportStatusPayload };

export type RuntimeRequest =
  | PingSWRequest
  | GetExportStatusRequest
  | StartExportRequest
  | StartBundleExportRequest
  | StopExportRequest
  | BuildAndDownloadRequest
  | ListConversationsRequest
  | ListConversationsQuickRequest
  | GetDiagnosticsBgRequest
  | RunProbesBgRequest
  | ProbeFileDownloadRequest
  | SalvageLinksRequest
  | ClearDiagnosticsLogsRequest
  | SetDiagLogPersistRequest
  | SetDiagVerboseStatsRequest;

export type BackgroundIncomingMessage =
  | PingSWRequest
  | GetExportStatusRequest
  | StartExportRequest
  | StartBundleExportRequest
  | StopExportRequest
  | BuildAndDownloadRequest
  | ExportStatusUpdateMessage
  | ScrapeProgressMessage
  | FetchBlobRequest
  | FetchBlobDirectRequest
  | FallbackStatusRequest
  | ListConversationsRequest
  | ListConversationsQuickRequest
  | GetDiagnosticsBgRequest
  | RunProbesBgRequest
  | ProbeFileDownloadRequest
  | SalvageLinksRequest
  | DiagLogForwardRequest
  | ClearDiagnosticsLogsRequest
  | SetDiagLogPersistRequest
  | SetDiagVerboseStatsRequest;

export type RuntimeResponse<T extends RuntimeRequest> =
  T extends PingSWRequest ? PingSWResponse :
  T extends GetExportStatusRequest ? GetExportStatusResponse :
  T extends StartExportRequest ? StartExportResponse :
  T extends StartBundleExportRequest ? StartBundleExportResponse :
  T extends StopExportRequest ? StopExportResponse :
  T extends ListConversationsRequest ? ListConversationsResponse :
  T extends ListConversationsQuickRequest ? ListConversationsQuickResponse :
  T extends GetDiagnosticsBgRequest ? GetDiagnosticsBgResponse :
  T extends RunProbesBgRequest ? RunProbesBgResponse :
  T extends ProbeFileDownloadRequest ? ProbeFileDownloadResponse :
  T extends ClearDiagnosticsLogsRequest ? ClearDiagnosticsLogsResponse :
  T extends SetDiagLogPersistRequest ? SetDiagLogPersistResponse :
  T extends SetDiagVerboseStatsRequest ? SetDiagVerboseStatsResponse :
  unknown;
