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

export type BuildAndDownloadRequest = {
  type: 'BUILD_AND_DOWNLOAD';
  data: {
    messages?: ExportMessage[];
    meta?: Record<string, unknown>;
    format?: 'json' | 'csv' | 'html';
    saveAs?: boolean;
    embedAvatars?: boolean;
  };
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

export type RuntimeRequest = PingSWRequest | GetExportStatusRequest | StartExportRequest | BuildAndDownloadRequest;

export type BackgroundIncomingMessage = PingSWRequest | GetExportStatusRequest | StartExportRequest | BuildAndDownloadRequest | ScrapeProgressMessage;
export type PopupIncomingMessage = ExportStatusMessage | ScrapeProgressMessage;

export type RuntimeResponse<T extends RuntimeRequest> =
  T extends PingSWRequest ? PingSWResponse :
  T extends GetExportStatusRequest ? GetExportStatusResponse :
  T extends StartExportRequest ? StartExportResponse :
  unknown;
