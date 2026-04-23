import type { BuildOptions, ExportMessage, ExportStatusPayload, ScrapeOptions } from './shared';

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

// START_EXPORT_ZIP was removed in the multi-format migration. The popup no
// longer pre-routes by format; the service worker decides single-file vs
// HTML.zip vs bundle.zip from `buildOptions.formats` + `downloadImages`.

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
export type FetchBlobResponse =
  | { ok: true; dataUrl: string; size: number }
  | { ok: false; cancelled?: boolean; status?: number; statusText?: string; error?: string; sizeReason?: number };

export type StopExportRequest = { type: 'STOP_EXPORT'; tabId?: number | null };
export type StopExportResponse = { ok: boolean; error?: string };

// NOTE: Open/Show actions on a finished download are handled by the popup
// calling chrome.downloads.open / .show directly. Routing through the
// service worker broke the user-activation requirement on downloads.open
// in MV3. See openSavedDownload in App.svelte.

// Ask the content script for the current Teams conversation id. Sent to a
// specific tab via `chrome.tabs.sendMessage` (not runtime.sendMessage), since
// the response depends on per-tab DOM/IndexedDB state.
// Returns `null` when the page isn't showing a conversation yet (e.g. Teams
// landing) or when extraction fails.
export type GetConvIdRequest = { type: 'GET_CONV_ID' };
export type GetConvIdResponse = { convId: string | null };

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
  | StopExportRequest
  | BuildAndDownloadRequest;

export type BackgroundIncomingMessage =
  | PingSWRequest
  | GetExportStatusRequest
  | StartExportRequest
  | StopExportRequest
  | BuildAndDownloadRequest
  | ExportStatusUpdateMessage
  | ScrapeProgressMessage
  | FetchBlobRequest;
export type PopupIncomingMessage = ExportStatusMessage | ScrapeProgressMessage;

export type RuntimeResponse<T extends RuntimeRequest> =
  T extends PingSWRequest ? PingSWResponse :
  T extends GetExportStatusRequest ? GetExportStatusResponse :
  T extends StartExportRequest ? StartExportResponse :
  T extends StopExportRequest ? StopExportResponse :
  unknown;
