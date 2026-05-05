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
//   cancel  : { type: 'tce-urlp-cancel' } — aborts every in-flight
//             fetch this helper started in the current export. Used
//             when the user clicks Stop while many slow helper calls
//             are pending; without it we wait on the 30s content-side
//             timeout for each one.
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

  // Track every in-flight fetch's AbortController by request id so we
  // can cancel them all in response to a tce-urlp-cancel message.
  var inflight = new Map();

  window.addEventListener('message', async function (e) {
    if (e.source !== window) return;
    var data = e.data;
    if (!data) return;

    if (data.type === 'tce-urlp-cancel') {
      inflight.forEach(function (controller) {
        try { controller.abort(); } catch (_) { /* ignore */ }
      });
      inflight.clear();
      return;
    }

    if (
      data.type !== 'tce-urlp-fetch'
      || typeof data.id !== 'string'
      || typeof data.url !== 'string'
    ) return;

    var controller = new AbortController();
    inflight.set(data.id, controller);

    var reply;
    try {
      var resp = await fetch(data.url, {
        credentials: 'include',
        signal: controller.signal,
      });
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
    } finally {
      inflight.delete(data.id);
    }

    // Transfer the ArrayBuffer to avoid copying ~MB-sized image bytes.
    var transfer = reply && reply.bytes ? [reply.bytes] : [];
    window.postMessage(reply, '*', transfer);
  });

  // Signal readiness so the content script can resolve its
  // ensureUrlpHelperLoaded() promise.
  window.postMessage({ type: 'tce-urlp-helper-ready' }, '*');
})();
