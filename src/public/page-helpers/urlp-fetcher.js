// Page-context helper for cookie-authenticated fetches that the
// extension's content-script context cannot make on its own.
//
// Why: Firefox content scripts (and Chrome with 3rd-party cookie
// partitioning) make fetches through the extension's network stack,
// not the page's. Cookies set on the asyncgw domain by Teams' login
// flow (authtoken_asm_urlp, skypetoken_asm) are partitioned to the
// teams.cloud.microsoft top-level origin; fetches from any other
// partition key (the extension's) won't see them, so the URL-image
// proxy returns 401. Running fetch() from the page's MAIN world
// shares the page's partition key and the cookies attach normally.
//
// Loaded by the content script via a `<script src=runtime.getURL(…)>`
// element, which means we run with the Teams page's exact origin and
// cookie context. Communication is window.postMessage-based RPC so
// we don't have to share JS-world identity with the content script.
//
// Protocol:
//   request : { type: 'tce-urlp-fetch',  id: string, url: string }
//   response: { type: 'tce-urlp-result', id: string, ok: boolean,
//               status?: number, statusText?: string,
//               mime?: string, bytes?: ArrayBuffer, error?: string }
//   ready   : { type: 'tce-urlp-helper-ready' } — emitted once on load.
//
// All responses go to window with a wildcard targetOrigin so the
// content script (which shares window via Xray-wrapped DOM) can
// receive them. We check e.source === window in the listener to
// avoid acting on messages from other windows or iframes.

(function () {
  'use strict';

  // Idempotent: a second injection should be a no-op.
  if (window.__teamsExporterUrlpHelperLoaded) return;
  window.__teamsExporterUrlpHelperLoaded = true;

  window.addEventListener('message', async function (e) {
    if (e.source !== window) return;
    var data = e.data;
    if (
      !data
      || data.type !== 'tce-urlp-fetch'
      || typeof data.id !== 'string'
      || typeof data.url !== 'string'
    ) return;

    var reply;
    try {
      var resp = await fetch(data.url, { credentials: 'include' });
      if (!resp.ok) {
        reply = {
          type: 'tce-urlp-result',
          id: data.id,
          ok: false,
          status: resp.status,
          statusText: resp.statusText || '',
        };
      } else {
        var buf = await resp.arrayBuffer();
        var ct = resp.headers.get('content-type') || '';
        var mime = ct.split(';')[0].trim() || 'image/jpeg';
        reply = {
          type: 'tce-urlp-result',
          id: data.id,
          ok: true,
          mime: mime,
          bytes: buf,
        };
      }
    } catch (err) {
      reply = {
        type: 'tce-urlp-result',
        id: data.id,
        ok: false,
        error: String((err && err.message) || err),
      };
    }

    // Transfer the ArrayBuffer to avoid copying ~MB-sized image bytes.
    var transfer = reply.bytes ? [reply.bytes] : [];
    window.postMessage(reply, '*', transfer);
  });

  // Signal readiness so the content script can resolve its
  // ensureUrlpHelperLoaded() promise.
  window.postMessage({ type: 'tce-urlp-helper-ready' }, '*');
})();
