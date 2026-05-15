// SVG-to-PNG rasterization via HTMLImageElement + OffscreenCanvas.
//
// This module lives outside the background script so it can be shared by:
//   1. The Firefox MV2 background script (which has DOM access directly).
//   2. The Chrome / Edge MV3 offscreen document (hosts a DOM page on the
//      SW's behalf so MV3 service workers can rasterize SVGs).
//
// Why a shared module: service workers have no `Image` or `document`,
// so `createImageBitmap` is the only async image decoder available. In
// Chromium MV3 service workers, `createImageBitmap` on SVG blobs throws
// `InvalidStateError` for many real-world SVGs (verified against the
// Twemoji set). The HTMLImageElement path used by Firefox MV2 is rock-
// solid; the offscreen document gives Chromium SWs access to it too.

// Twemoji SVGs carry only `viewBox="0 0 36 36"` with no width/height
// attributes. Firefox's createImageBitmap silently fails to rasterize
// such "viewport-only" SVGs even when resizeWidth/Height are passed.
// Stamp explicit width/height onto the root <svg> tag so every decoder
// can size the bitmap reliably. Idempotent — strips any existing
// width/height attributes before adding the requested size.
export function prepareSvgText(svgText: string, sizePx: number): string {
  return svgText.replace(/<svg\b([^>]*)>/i, (_m, attrs: string) => {
    const stripped = attrs
      .replace(/\s+width="[^"]*"/i, '')
      .replace(/\s+height="[^"]*"/i, '');
    return `<svg${stripped} width="${sizePx}" height="${sizePx}">`;
  });
}

// Rasterize an SVG to PNG bytes via the DOM image-decoder pipeline.
// Requires a context where `Image` and `OffscreenCanvas` are available
// (Firefox MV2 background, popup window, or an offscreen document on
// Chromium MV3). Returns null on decode/draw failure.
export async function rasterizeSvgInDom(svgText: string, sizePx: number): Promise<Uint8Array | null> {
  const prepared = prepareSvgText(svgText, sizePx);
  const dataUrl = 'data:image/svg+xml;base64,' + btoa(unescape(encodeURIComponent(prepared)));
  const img = new Image();
  img.width = sizePx;
  img.height = sizePx;
  const loaded = new Promise<boolean>(resolve => {
    img.onload = () => resolve(true);
    img.onerror = () => resolve(false);
  });
  img.src = dataUrl;
  const ok = await loaded;
  if (!ok) return null;
  const canvas = new OffscreenCanvas(sizePx, sizePx);
  const ctx = canvas.getContext('2d');
  if (!ctx) return null;
  try {
    ctx.drawImage(img, 0, 0, sizePx, sizePx);
  } catch {
    return null;
  }
  const pngBlob = await canvas.convertToBlob({ type: 'image/png' });
  return new Uint8Array(await pngBlob.arrayBuffer());
}
