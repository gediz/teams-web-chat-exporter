// SharePoint sharing-link resolver — the same two-step flow Teams' own web
// client uses to open/download file attachments (verified request-by-request
// against three HAR captures of manual downloads from a real tenant).
//
// A file's absolute SharePoint URL is encoded into a share token
// ('u!' + base64url(url)) and resolved via
//   GET https://<file-host>/_api/v2.0/shares/<token>?select=...
// with `Prefer: redeemSharingLink,getShortLivedDownloadUrl`. The token itself
// is the capability: the captured requests carry NO cookies and NO
// Authorization header, and succeed even when the browser has no session on
// the file's host (e.g. a migrated-tenant OneDrive) — exactly the case where
// the raw-URL chrome.downloads transport comes back with a request-access
// page instead of the file.
//
// AUTH (verified via CDP network capture): Teams' working resolve carries an
// Authorization: Bearer token scoped to the file's SharePoint resource and
// targets the /_api/v2.0/shares/<token>/driveItem/ sub-path. A request without
// that Bearer 401s even with the correct sharing-link token and Teams' exact
// other headers (the shares endpoint answers ACAO:* so the request is
// uncredentialed — no cookie involved either way). We extract the same
// SharePoint token from the MSAL cache (getSharePointToken) and send it.
//
// The fetch runs through a page-world helper (public/page-helpers/
// share-resolver.js), mirroring the image pipeline: a content-script fetch is
// subject to the extension's network context and returned 401/failed here,
// while the page world matches Teams' exact request context.
//
// The response is the file's driveItem: `@content.downloadUrl` (a short-lived
// PRE-AUTHENTICATED URL that chrome.downloads can fetch with no cookie
// dependency) plus `currentUserRole.blocksDownload` — an upfront access
// oracle this project previously believed unobtainable (see
// tce-debug/ACCESS_DETECTION_FINDINGS.md; that conclusion held only for the
// background's chrome-extension:// origin, which SharePoint rejects
// server-side. This module runs in the CONTENT SCRIPT, whose fetches carry
// the Teams page origin that SharePoint's CORS allows — proven by the HARs).
//
// SECURITY: href comes from message content (attacker-influenceable), so it
// is gated to the SharePoint host family before any request is made — same
// gate as the download feature. The returned downloadUrl embeds a short-lived
// bearer-like token: pass it to chrome.downloads, never log it.

import { isSharePointFileHost } from '../utils/teams-urls';
import { getSharePointToken } from './api-client';

export interface ShareResolveResult {
  ok: boolean;
  // HTTP status of the shares call; 0 = request never completed (network
  // error, CORS rejection) and `error` carries the reason.
  status: number;
  name?: string;
  mimeType?: string;
  // Short-lived pre-authenticated download URL. Contains a token: do not log.
  downloadUrl?: string;
  // Access oracle. When true the server will refuse the bytes; when the
  // whole resolve fails with 403/404 the file is inaccessible or gone.
  blocksDownload?: boolean;
  allowEdit?: boolean;
  readOnly?: boolean;
  // Stable file GUID (listItemUniqueId) — the download.aspx UniqueId.
  itemId?: string;
  error?: string;
  // Which credentials mode the page-world helper used (diagnostic).
  mode?: string;
}

// Mirrors the select list Teams sends (from the HAR), so SharePoint treats
// the request identically. `@microsoft.graph.downloadUrl` arrives in the
// response as `@content.downloadUrl` on this endpoint.
const SHARES_SELECT = '@microsoft.graph.downloadUrl,file,id,webDavUrl,sharepointIds,eTag,name,currentUserRole,parentReference';

/** Graph sharing-token encoding: 'u!' + unpadded base64url of the URL. */
export function encodeShareToken(url: string): string {
  const b64 = btoa(unescape(encodeURIComponent(url)));
  return 'u!' + b64.replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
}

// ── Page-world helper injection + RPC ───────────────────────────────────────
// Mirrors ensureUrlpHelperLoaded in content.ts. Lazy + idempotent.

const SHARE_REQ = 'tce-share-fetch';
const SHARE_RES = 'tce-share-result';
const SHARE_READY = 'tce-share-helper-ready';

let shareHelperPromise: Promise<'ready' | 'error'> | null = null;
function ensureShareHelperLoaded(): Promise<'ready' | 'error'> {
  if (shareHelperPromise) return shareHelperPromise;
  shareHelperPromise = new Promise(resolve => {
    let done = false;
    const finish = (s: 'ready' | 'error') => { if (done) return; done = true; window.removeEventListener('message', onReady); resolve(s); };
    function onReady(e: MessageEvent) {
      if (e.source !== window) return;
      if ((e.data as { type?: string } | null)?.type === SHARE_READY) finish('ready');
    }
    window.addEventListener('message', onReady);
    try {
      const script = document.createElement('script');
      // wxt exposes browser.runtime.getURL; content modules can use the global.
      script.src = (globalThis as { chrome?: { runtime: { getURL(p: string): string } }; browser?: { runtime: { getURL(p: string): string } } })
        .browser?.runtime.getURL('page-helpers/share-resolver.js')
        ?? (globalThis as { chrome: { runtime: { getURL(p: string): string } } }).chrome.runtime.getURL('page-helpers/share-resolver.js');
      script.onload = () => script.remove();
      script.onerror = () => finish('error');
      (document.head || document.documentElement).appendChild(script);
    } catch { finish('error'); }
    setTimeout(() => finish('error'), 5_000);
  });
  return shareHelperPromise;
}

let shareRpcSeq = 0;
function pageWorldGet(url: string, headers: Record<string, string>): Promise<{ ok: boolean; status: number; json?: Record<string, unknown>; error?: string; mode?: string }> {
  return new Promise(async resolve => {
    const ready = await ensureShareHelperLoaded();
    if (ready !== 'ready') { resolve({ ok: false, status: 0, error: 'share helper failed to load' }); return; }
    const id = `share-${++shareRpcSeq}`;
    let settled = false;
    const finish = (r: { ok: boolean; status: number; json?: Record<string, unknown>; error?: string; mode?: string }) => {
      if (settled) return; settled = true; window.removeEventListener('message', onMsg); resolve(r);
    };
    function onMsg(e: MessageEvent) {
      if (e.source !== window) return;
      const d = e.data as { type?: string; id?: string; ok?: boolean; status?: number; json?: Record<string, unknown>; error?: string; mode?: string } | null;
      if (d?.type !== SHARE_RES || d.id !== id) return;
      finish({ ok: !!d.ok, status: d.status ?? 0, json: d.json, error: d.error, mode: d.mode });
    }
    window.addEventListener('message', onMsg);
    window.postMessage({ type: SHARE_REQ, id, url, headers }, '*');
    setTimeout(() => finish({ ok: false, status: 0, error: 'share resolve timeout' }), 20_000);
  });
}

/**
 * Resolve a SharePoint file URL to its driveItem (downloadUrl + access
 * oracle). Must run in a context whose fetch carries the Teams page origin
 * (content script / page world) — SharePoint rejects the extension origin.
 */
export async function resolveShareFile(href: string): Promise<ShareResolveResult> {
  if (!isSharePointFileHost(href)) {
    return { ok: false, status: 0, error: 'not a SharePoint file host' };
  }
  let host: string;
  try { host = new URL(href).host; } catch { return { ok: false, status: 0, error: 'invalid URL' }; }

  // Teams authenticates the resolve with a SharePoint-scoped Bearer token and
  // hits the /driveItem/ sub-path (verified via CDP: its 200 request carries
  // Authorization + /driveItem/; our tokenless base-path request 401s). Match
  // both. The token is uncredentialed-with-Bearer, so no cookie/partition
  // dependency — this is why it works when the raw-URL session cookie is stale.
  const bearer = await getSharePointToken(host);
  const endpoint = `https://${host}/_api/v2.0/shares/${encodeShareToken(href)}/driveItem/?select=${encodeURIComponent(SHARES_SELECT)}`;
  // NOTE: we intentionally drop Teams' `manualRedirect` from Prefer. With it,
  // SharePoint returns @content.downloadUrl as a bare download.aspx?UniqueId
  // form that still needs the session cookie / a POSTed access_token (Teams
  // drives that itself). Without it, getShortLivedDownloadUrl yields a fully
  // pre-authenticated (token-bearing) URL that chrome.downloads can GET on its
  // own — which is the whole point of resolving.
  const headers: Record<string, string> = {
    'Accept': 'application/json',
    'Prefer': 'redeemSharingLink,getShortLivedDownloadUrl',
    'application': 'Teams_Web',
  };
  if (bearer) headers['Authorization'] = `Bearer ${bearer}`;
  const res = await pageWorldGet(endpoint, headers);
  if (!res.ok || !res.json) {
    return {
      ok: false,
      status: res.status,
      error: (res.error || `shares API returned ${res.status}`) + (bearer ? '' : ' (no SharePoint token found)'),
      mode: res.mode,
    };
  }
  const j = res.json;

  const role = (j.currentUserRole ?? {}) as Record<string, unknown>;
  const spIds = (j.sharepointIds ?? {}) as Record<string, unknown>;
  const file = (j.file ?? {}) as Record<string, unknown>;
  const downloadUrl = (j['@content.downloadUrl'] ?? j['@microsoft.graph.downloadUrl']) as string | undefined;
  return {
    ok: true,
    status: res.status,
    name: typeof j.name === 'string' ? j.name : undefined,
    mimeType: typeof file.mimeType === 'string' ? file.mimeType : undefined,
    downloadUrl: typeof downloadUrl === 'string' ? downloadUrl : undefined,
    blocksDownload: typeof role.blocksDownload === 'boolean' ? role.blocksDownload : undefined,
    allowEdit: typeof role.allowEdit === 'boolean' ? role.allowEdit : undefined,
    readOnly: typeof role.readOnly === 'boolean' ? role.readOnly : undefined,
    itemId: typeof spIds.listItemUniqueId === 'string' ? spIds.listItemUniqueId : undefined,
    mode: res.mode,
  };
}
