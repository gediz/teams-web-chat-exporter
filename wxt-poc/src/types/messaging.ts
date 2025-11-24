import type { BuildOptions, ExportStatusPayload, ScrapeOptions } from './shared';

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

export type RuntimeRequest = PingSWRequest | GetExportStatusRequest | StartExportRequest;

export type RuntimeResponse<T extends RuntimeRequest> =
  T extends PingSWRequest ? PingSWResponse :
  T extends GetExportStatusRequest ? GetExportStatusResponse :
  T extends StartExportRequest ? StartExportResponse :
  unknown;
