// Offscreen document entry point.
//
// Hosted by chrome.offscreen for Chrome/Edge MV3 service workers that need
// DOM access the worker lacks. It does two jobs:
//   1. SVG rasterization (OFFSCREEN_RASTERIZE_SVG) for the PDF emoji path.
//   2. Minting blob: URLs (OFFSCREEN_MINT_BLOB_URL) for large downloads.
//      The SW cannot call URL.createObjectURL; for a large export, a
//      base64 data: URL exceeds V8's max string length and throws
//      "Invalid string length" (issue #27). The SW parks the export Blob
//      in IndexedDB and this document reads it back and mints a blob: URL.
//
// This document uses only chrome.runtime plus DOM/Blob/URL/IndexedDB APIs.
// The offscreen document API restricts the extension surface to
// chrome.runtime.

import { rasterizeSvgInDom } from '../../utils/svg-rasterize';
import { takeTransferBlob } from '../../utils/blob-transfer';

type RasterizeRequest = {
  type: 'OFFSCREEN_RASTERIZE_SVG';
  svgText: string;
  sizePx: number;
};

type MintBlobUrlRequest = {
  type: 'OFFSCREEN_MINT_BLOB_URL';
  key: string;
};

type RevokeBlobUrlRequest = {
  type: 'OFFSCREEN_REVOKE_BLOB_URL';
  url: string;
};

chrome.runtime.onMessage.addListener((
  msg: unknown,
  _sender: chrome.runtime.MessageSender,
  sendResponse: (response: unknown) => void,
): boolean => {
  if (isRasterizeRequest(msg)) {
    rasterizeSvgInDom(msg.svgText, msg.sizePx)
      .then(bytes => {
        if (bytes) {
          // chrome.runtime messages go through JSON serialization. Convert
          // the Uint8Array to a plain array so the bytes survive the trip.
          sendResponse({ ok: true, bytes: Array.from(bytes) });
        } else {
          sendResponse({ ok: false, error: 'rasterize returned null' });
        }
      })
      .catch((err: unknown) => {
        const message = err instanceof Error ? err.message : String(err);
        sendResponse({ ok: false, error: message });
      });
    // Keep the message channel open for the async response.
    return true;
  }

  if (isMintBlobUrlRequest(msg)) {
    // Read the staged Blob back out of IndexedDB and mint a blob: URL.
    // The URL stays valid as long as this document lives (it is never
    // closed), which covers the chrome.downloads.download queueing window.
    takeTransferBlob(msg.key)
      .then(blob => {
        if (!blob) {
          sendResponse({ ok: false, error: 'transfer blob not found' });
          return;
        }
        const url = URL.createObjectURL(blob);
        sendResponse({ ok: true, url });
      })
      .catch((err: unknown) => {
        const message = err instanceof Error ? err.message : String(err);
        sendResponse({ ok: false, error: message });
      });
    return true;
  }

  if (isRevokeBlobUrlRequest(msg)) {
    try { URL.revokeObjectURL(msg.url); } catch { /* already gone */ }
    sendResponse({ ok: true });
    return false;
  }

  return false;
});

function isRasterizeRequest(msg: unknown): msg is RasterizeRequest {
  if (!msg || typeof msg !== 'object') return false;
  const m = msg as Record<string, unknown>;
  return (
    m.type === 'OFFSCREEN_RASTERIZE_SVG' &&
    typeof m.svgText === 'string' &&
    typeof m.sizePx === 'number'
  );
}

function isMintBlobUrlRequest(msg: unknown): msg is MintBlobUrlRequest {
  if (!msg || typeof msg !== 'object') return false;
  const m = msg as Record<string, unknown>;
  return m.type === 'OFFSCREEN_MINT_BLOB_URL' && typeof m.key === 'string';
}

function isRevokeBlobUrlRequest(msg: unknown): msg is RevokeBlobUrlRequest {
  if (!msg || typeof msg !== 'object') return false;
  const m = msg as Record<string, unknown>;
  return m.type === 'OFFSCREEN_REVOKE_BLOB_URL' && typeof m.url === 'string';
}
