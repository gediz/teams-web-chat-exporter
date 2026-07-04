// Page-context helper for the SharePoint sharing-link resolve.
//
// Why (same reason as urlp-fetcher.js): the resolve is UNCREDENTIALED and
// authenticated by a Bearer token, not a cookie (see the body comment: the
// endpoint answers Access-Control-Allow-Origin: '*', which forbids
// credentials). It must run from the Teams page origin: a content-script fetch
// runs through the EXTENSION's network partition / foreign origin and fails,
// while the page MAIN world sends the exact page Origin SharePoint accepts and
// the call returns 200 with the pre-authenticated downloadUrl. This helper runs
// the identical request from the page world.
//
// Loaded by the content script via <script src=runtime.getURL(…)>, so it runs
// with the Teams page's exact origin and cookie context. RPC via
// window.postMessage.
//
// Protocol:
//   request : { type: 'tce-share-fetch', id, url, headers }
//   response: { type: 'tce-share-result', id, ok, status, json?, error? }
//   ready   : { type: 'tce-share-helper-ready' }

(function () {
  'use strict';
  if (window.__teamsExporterShareHelperLoaded) return;
  window.__teamsExporterShareHelperLoaded = true;

  window.addEventListener('message', async function (e) {
    if (e.source !== window) return;
    var data = e.data;
    if (!data || data.type !== 'tce-share-fetch' || typeof data.id !== 'string' || typeof data.url !== 'string') return;

    // The shares endpoint answers with Access-Control-Allow-Origin: '*',
    // which the browser forbids combining with credentials — so a
    // credentialed request is CORS-blocked ("Failed to fetch"). Teams' own
    // working call is therefore uncredentialed; the page world's value is the
    // correct page Origin, not the cookie. Try 'omit' first (matches Teams),
    // fall back to 'include' in case a tenant returns a non-wildcard ACAO.
    async function attempt(mode) {
      var resp = await fetch(data.url, { method: 'GET', credentials: mode, headers: data.headers || {} });
      if (!resp.ok) return { ok: false, status: resp.status, mode: mode };
      var json = null;
      try { json = await resp.json(); } catch (_) { json = null; }
      return { ok: true, status: resp.status, json: json, mode: mode };
    }
    var reply;
    try {
      var r;
      try { r = await attempt('omit'); }
      catch (e1) { r = await attempt('include'); }
      reply = { type: 'tce-share-result', id: data.id, ok: r.ok, status: r.status, json: r.json, mode: r.mode };
    } catch (err) {
      reply = { type: 'tce-share-result', id: data.id, ok: false, status: 0, error: String((err && err.message) || err) };
    }
    window.postMessage(reply, '*');
  });

  window.postMessage({ type: 'tce-share-helper-ready' }, '*');
})();
