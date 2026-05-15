// Offscreen document entry point.
//
// Hosted by chrome.offscreen for Chrome/Edge MV3 service workers that need
// DOM access for SVG rasterization. Listens for OFFSCREEN_RASTERIZE_SVG
// messages from the SW, rasterizes via the shared svg-rasterize utility,
// and posts the PNG bytes back.
//
// This document does NOT use chrome.action, chrome.tabs, or any other
// extension API beyond chrome.runtime — the offscreen document API
// deliberately restricts surface to chrome.runtime only.

import { rasterizeSvgInDom } from '../../utils/svg-rasterize';

type RasterizeRequest = {
  type: 'OFFSCREEN_RASTERIZE_SVG';
  svgText: string;
  sizePx: number;
};

type RasterizeResponse =
  | { ok: true; bytes: number[] }
  | { ok: false; error?: string };

chrome.runtime.onMessage.addListener((
  msg: unknown,
  _sender: chrome.runtime.MessageSender,
  sendResponse: (response: RasterizeResponse) => void,
): boolean => {
  if (!isRasterizeRequest(msg)) return false;
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
