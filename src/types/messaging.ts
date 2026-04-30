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
  | ListConversationsQuickRequest;

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
  | ListConversationsQuickRequest;

export type RuntimeResponse<T extends RuntimeRequest> =
  T extends PingSWRequest ? PingSWResponse :
  T extends GetExportStatusRequest ? GetExportStatusResponse :
  T extends StartExportRequest ? StartExportResponse :
  T extends StartBundleExportRequest ? StartBundleExportResponse :
  T extends StopExportRequest ? StopExportResponse :
  T extends ListConversationsRequest ? ListConversationsResponse :
  T extends ListConversationsQuickRequest ? ListConversationsQuickResponse :
  unknown;
