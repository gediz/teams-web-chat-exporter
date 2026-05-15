// Service-worker-side client for the offscreen rasterization document.
//
// Chrome/Edge MV3 service workers have no DOM, so SVG decoding via
// createImageBitmap throws InvalidStateError for most real-world SVGs.
// The offscreen document hosts a DOM page on the SW's behalf and
// rasterizes via HTMLImageElement + OffscreenCanvas (proven path).
//
// This module is callable only from contexts with chrome.offscreen
// available (MV3 service workers on Chrome 109+ / Edge 109+). Firefox
// MV2 doesn't have chrome.offscreen and doesn't need it (background
// page already has a DOM).

const OFFSCREEN_URL = 'offscreen.html';

// Tracks an in-flight document creation so concurrent rasterizeViaOffscreen
// callers don't all try to create the document simultaneously.
let creationInFlight: Promise<void> | null = null;

async function offscreenDocumentExists(): Promise<boolean> {
  // chrome.runtime.getContexts is the canonical existence check (Chrome 116+).
  // We only target Chrome 109+ for the API, so fall back to hasDocument()
  // for the 109-115 window if getContexts isn't available.
  const r = chrome.runtime as unknown as {
    getContexts?: (filter: { contextTypes: string[] }) => Promise<unknown[]>;
  };
  if (typeof r.getContexts === 'function') {
    const contexts = await r.getContexts({ contextTypes: ['OFFSCREEN_DOCUMENT'] });
    return contexts.length > 0;
  }
  const o = chrome.offscreen as unknown as { hasDocument?: () => Promise<boolean> };
  if (typeof o.hasDocument === 'function') {
    return await o.hasDocument();
  }
  // No way to check existence; create and let "already exists" error surface.
  return false;
}

async function ensureOffscreenDocument(): Promise<void> {
  if (await offscreenDocumentExists()) return;
  if (!creationInFlight) {
    creationInFlight = chrome.offscreen.createDocument({
      url: OFFSCREEN_URL,
      reasons: [chrome.offscreen.Reason.BLOBS],
      justification: 'Decode SVG emoji images for PDF rendering. Service workers lack DOM access required to rasterize SVG via HTMLImageElement.',
    }).catch(err => {
      // "Only a single offscreen document may be created" is a benign race
      // when concurrent callers hit ensure() at the same time. Treat as success.
      const message = err instanceof Error ? err.message : String(err);
      if (!/only a single offscreen document/i.test(message)) {
        throw err;
      }
    });
  }
  try {
    await creationInFlight;
  } finally {
    creationInFlight = null;
  }
}

// Rasterize an SVG to PNG bytes via the offscreen document. Retries once
// if the first attempt fails: the offscreen document may have been torn
// down alongside a previous SW death, in which case recreate-and-retry
// succeeds. Returns null on persistent failure.
export async function rasterizeViaOffscreen(
  svgText: string,
  sizePx: number,
): Promise<Uint8Array | null> {
  for (let attempt = 0; attempt < 2; attempt++) {
    try {
      await ensureOffscreenDocument();
      const response = await chrome.runtime.sendMessage({
        type: 'OFFSCREEN_RASTERIZE_SVG',
        svgText,
        sizePx,
      }) as { ok: boolean; bytes?: number[]; error?: string } | undefined;

      if (response && response.ok && Array.isArray(response.bytes)) {
        return new Uint8Array(response.bytes);
      }
      // Response shape was wrong or ok:false. No point retrying; the offscreen
      // doc handled the request but couldn't produce bytes (e.g., bad SVG).
      return null;
    } catch (err) {
      // SendMessage threw — usually means the offscreen doc is gone (SW
      // restart killed it) or never came up. Reset state and retry once.
      if (attempt === 0) {
        creationInFlight = null;
        continue;
      }
      console.warn('[offscreen-client] rasterize failed after retry:', err);
      return null;
    }
  }
  return null;
}
