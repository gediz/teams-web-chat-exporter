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
export type StartExportResponse = { filename?: string; messages?: number; error?: string; code?: string };

export type StartExportZipRequest = {
  type: 'START_EXPORT_ZIP';
  data: {
    tabId?: number | null;
    scrapeOptions: ScrapeOptions;
    buildOptions: BuildOptions;
  };
};
export type StartExportZipResponse = { ok?: boolean; filename?: string; downloadId?: number; error?: string; code?: string };

export type BuildAndDownloadRequest = {
  type: 'BUILD_AND_DOWNLOAD';
  data: {
    messages?: ExportMessage[];
    meta?: Record<string, unknown>;
    format?: 'json' | 'csv' | 'html' | 'txt';
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
  | StartExportZipRequest
  | StopExportRequest
  | BuildAndDownloadRequest;

export type BackgroundIncomingMessage =
  | PingSWRequest
  | GetExportStatusRequest
  | StartExportRequest
  | StartExportZipRequest
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
  T extends StartExportZipRequest ? StartExportZipResponse :
  T extends StopExportRequest ? StopExportResponse :
  unknown;
