// Document attachment download (the "Files" toggle).
//
// Streams non-image SharePoint/OneDrive-for-Business file attachments
// (PDF/DOCX/XLSX/ZIP/video/...) to disk via chrome.downloads, into an
// `attachments/` folder beside the export. Inline images are handled
// separately by downloadImages; this is the paperclip / shared-file documents
// that are otherwise kept as links.
//
// Mechanism (verified live against a real tenant): the converter stores a
// document's absolute SharePoint URL (the file's webDavUrl/objectUrl) as
// Attachment.href. SharePoint serves that URL's raw bytes to an authenticated
// caller (HTTP 200, the file content), so we hand it straight to
// chrome.downloads.download. The browser fetches it in its OWN network context,
// sending the user's first-party SharePoint cookies and following any redirect,
// streaming server->disk with zero extension-held bytes. Most types serve raw
// bytes directly (`?download=1` force-download hint); renderable-markup types
// (html/svg/xml/...) are routed through the _layouts/download.aspx handler by
// the file's UniqueId so SharePoint's active-content sandbox interstitial isn't
// saved in their place. See toDownloadUrl.
//
// A file the user cannot open can't be detected upfront (no reliable access
// oracle exists — see debug/ACCESS_DETECTION_FINDINGS.md); it comes down as a
// "request access" HTML page saved under the real file's name. For non-markup
// files that junk is cleaned up after the fact by verifyDownloadOnChanged
// (below); markup is ambiguous (a real .html is also text/html) so a genuinely
// inaccessible markup file is the one residual junk case.
//
// Why not the `/_api/v2.0/shares` resolver: it works same-origin but rejects
// the extension's cross-origin background fetch (foreign chrome-extension://
// origin), so every resolve returned null (0 saved) in testing. The direct
// chrome.downloads path avoids it: downloads run in the browser context, not
// the extension origin, so there is no foreign-origin rejection and no
// credentialed extension-origin fetch at all.
//
// SECURITY: href is attacker-influenceable (it comes from message content), so
// it is gated to the SharePoint host family before it is ever handed to
// chrome.downloads. The browser only sends first-party cookies for the host it
// actually fetches, so a non-SharePoint href could not exfiltrate the
// SharePoint session even if the gate were bypassed; the gate additionally
// stops the feature from saving arbitrary attacker-hosted content into the
// export folder.

import type { ExportMessage } from '../types/shared';
import { isSharePointFileHost } from '../utils/teams-urls';
import { sanitizeFileName } from '../utils/messages';

// The subset of chrome.downloads we use. Mirrors the deps-injection style in
// download.ts so the module stays unit-testable without a real browser.
// search/removeFile/erase back the post-download request-access cleanup.
type DownloadsApi = Pick<typeof chrome.downloads, 'download' | 'search' | 'removeFile' | 'erase'>;

export interface AttachmentDownloadDeps {
  downloads: DownloadsApi;
  // Called after each candidate is processed. `done` counts processed
  // candidates; `total` is the candidate count.
  onProgress?: (done: number, total: number) => void;
  log?: (msg: string) => void;
}

export interface AttachmentDownloadSummary {
  total: number;    // SharePoint document candidates found (deduped by href)
  saved: number;    // handed to chrome.downloads (fire-and-forget; the browser
                    // streams server->disk in its own context)
  links: number;    // skipped without a download attempt (host gate) — kept as a link
  failed: number;   // chrome.downloads.download threw synchronously
  failures: Array<{ name: string; reason: string }>;
}

export interface AttachmentCandidate {
  href: string;        // the file's absolute SharePoint URL (download source)
  name: string;        // best-known display name (from Attachment.label)
  chatFolder: string;  // per-chat subfolder under the export root ('' for single-chat)
  itemid?: string;     // stable file GUID (== listItemUniqueId); the download.aspx UniqueId for markup types
}

// fileType values the converter routes to an AMS preview URL (so their href is
// NOT a SharePoint file URL). Skipped here: inline images are handled by the
// separate downloadImages path and would otherwise be double-fetched.
const IMAGE_FILETYPE = /^(png|jpe?g|gif|webp|bmp|svg|ico|tiff?|heic|heif|avif)$/i;

/** Last path segment of a URL, percent-decoded, as a filename fallback. */
function fileNameFromUrl(url: string): string {
  try {
    const seg = new URL(url).pathname.split('/').filter(Boolean).pop() || '';
    return decodeURIComponent(seg);
  } catch {
    return '';
  }
}

function errMsg(e: unknown): string {
  if (e && typeof e === 'object' && 'message' in e) return String((e as { message: unknown }).message);
  return String(e);
}

// Renderable-markup types that SharePoint serves through an auth/sandbox
// "Working..." interstitial (a form POST to /_forms/default.aspx) when the
// request looks like a navigation — which is how chrome.downloads fetches. For
// these, chrome.downloads would save the interstitial page, not the file.
const SANDBOXED_EXT = /\.(html?|xhtml|shtml|svgz?|xml|mht|mhtml|aspx?)$/i;

/** True when the download URL routes through the _layouts download handler. */
function isDownloadAspxUrl(url: string): boolean {
  return url.includes('/_layouts/15/download.aspx');
}

/**
 * Site-collection base for a SharePoint file URL: `https://<host>/personal/<u>`,
 * `/sites/<s>`, or `/teams/<t>`. The download.aspx `UniqueId` is site-collection
 * scoped, so the handler must be addressed on the file's OWN site, not the host
 * root — verified live: the root-host form returns a generic page for personal
 * (OneDrive) sites, while the site-collection form streams the real file. Falls
 * back to the host root for layouts we don't recognize.
 */
function siteCollectionBase(u: URL): string {
  const m = u.pathname.match(/^(\/(?:personal|sites|teams)\/[^/]+)/i);
  return `https://${u.host}${m ? m[1] : ''}`;
}

/**
 * Map a file's absolute SharePoint URL to the URL chrome.downloads should fetch.
 *
 * Most types serve their raw bytes directly to an authenticated caller, so we
 * append `?download=1` (SharePoint's force-download hint) and hand them over as
 * is — the proven path that pulls PDFs/zips/Office/video correctly.
 *
 * Renderable-markup types ({@link SANDBOXED_EXT}) instead hit the active-content
 * sandbox interstitial. Route those through the `_layouts/download.aspx` handler
 * addressed by the file's UniqueId on its own site collection, with the native
 * `Translate=false&ApiVersion=2.0` parameters. This is SharePoint's own download
 * path (verified against a captured HAR and a live probe): it forces a
 * server-side attachment download with just the FedAuth cookie. Falls back to the
 * absolute-path SourceUrl form when no UniqueId is known.
 */
function toDownloadUrl(href: string, itemid?: string): string {
  let u: URL;
  try { u = new URL(href); } catch { return href; }
  if (SANDBOXED_EXT.test(u.pathname)) {
    if (itemid) {
      return `${siteCollectionBase(u)}/_layouts/15/download.aspx?UniqueId=${encodeURIComponent(itemid)}&Translate=false&ApiVersion=2.0`;
    }
    return `https://${u.host}/_layouts/15/download.aspx?SourceUrl=${encodeURIComponent(href)}`;
  }
  return href + (href.includes('?') ? '&' : '?') + 'download=1';
}

/** Collect deduped SharePoint document attachments across all messages. */
export function collectDocumentAttachments(messages: ExportMessage[]): Array<{ href: string; name: string; itemid?: string }> {
  const seen = new Map<string, { href: string; name: string; itemid?: string }>();
  const out: Array<{ href: string; name: string; itemid?: string }> = [];
  for (const m of messages) {
    const atts = m.attachments;
    if (!atts) continue;
    for (const a of atts) {
      const href = a.href;
      if (!href) continue;
      // Skip inline-image types (handled by downloadImages) and anything not on
      // a SharePoint file host (the only family we download from).
      if (IMAGE_FILETYPE.test((a.type || '').toString())) continue;
      if (!isSharePointFileHost(href)) continue;
      const dup = seen.get(href);
      if (dup) {
        // Same file twice: keep the first, but adopt an itemid if the first
        // occurrence lacked one (so the join key is never lost to dedup order).
        if (!dup.itemid && a.itemid) dup.itemid = a.itemid;
        continue;
      }
      const entry = { href, name: (a.label || '').toString(), itemid: a.itemid };
      seen.set(href, entry);
      out.push(entry);
    }
  }
  return out;
}

/**
 * Build a Downloads-relative file path:
 *   <exportFolder>/[<chatFolder>/]attachments/<name>
 * Only the server-provided file `name` is sanitized here (sanitizeFileName
 * keeps the extension while defending against path traversal and reserved
 * names). exportFolder and chatFolder are already filesystem-safe upstream
 * (computeBaseName / pickBundleFolderName), so they are NOT re-sanitized:
 * doing so would re-apply title truncation and chop the export's stamp.
 */
export function buildAttachmentPath(exportFolder: string, chatFolder: string, name: string): string {
  const safeName = sanitizeFileName(name);
  return [exportFolder, chatFolder, 'attachments', safeName]
    .filter(s => s && s.length)
    .join('/');
}

/**
 * Write a FAILURES.txt listing files that were shared in the chat but could not
 * be downloaded (the user has no access). Goes in the same attachments/ folder,
 * via a data: URL through chrome.downloads. `overwrite` so each call replaces
 * the file with the full accumulated list (the post-download verifier calls
 * this incrementally as junk pages are discovered). The previous write's history
 * row is erased first (erase drops the shelf record, not the file) so repeated
 * rewrites leave a single FAILURES.txt entry. Each line carries the file's link
 * so the user can open it in Teams.
 */
async function writeFailuresFile(
  downloads: DownloadsApi,
  exportFolder: string,
  chatFolder: string,
  entries: Array<{ name: string; href: string }>,
  key: string,
): Promise<void> {
  const body =
    'These files were shared in this chat but could not be downloaded because\n' +
    'your account does not have access to them. Open each one in Teams using its link.\n\n' +
    entries.map(e => `${e.name}\n${e.href}\n`).join('\n');
  const url = 'data:text/plain;charset=utf-8,' + encodeURIComponent(body);
  const prev = failuresDownloadId.get(key);
  if (prev !== undefined) { try { await Promise.resolve(downloads.erase({ id: prev })); } catch { /* noop */ } }
  try {
    const id = await Promise.resolve(downloads.download({
      url,
      filename: buildAttachmentPath(exportFolder, chatFolder, 'FAILURES.txt'),
      saveAs: false,
      conflictAction: 'overwrite',
    }));
    if (typeof id === 'number') failuresDownloadId.set(key, id);
  } catch { /* best-effort; the file just stays out of the export */ }
}

// ── Post-download request-access verification ──────────────────────────────
//
// Fire-and-forget downloads can't inspect the response, so a file the user has
// no access to comes down as SharePoint's "request access" HTML page saved
// under the real file's name. There is no reliable way to know this upfront
// (see debug/ACCESS_DETECTION_FINDINGS.md), so we verify after the fact: when a
// NON-markup download completes with Content-Type text/html, it is that junk
// page (a real pdf/xlsx/zip/video is never text/html). Remove it from disk and
// history and list the file, with its link, in the chat's FAILURES.txt.
//
// Markup types (html/svg/xml) are deliberately NOT verified: routed through
// download.aspx they legitimately arrive as text/html, so the signal is
// ambiguous for them. Their request-access pages are the one junk case that
// still slips through.
//
// State is in-memory and best-effort. The stream of onChanged events keeps the
// MV3 worker alive while downloads land; if a long gap evicts it, late checks
// are simply skipped (the file just stays — no upfront data loss).

interface VerifyMeta { name: string; href: string; chatFolder: string; exportFolder: string; }
const pendingVerify = new Map<number, VerifyMeta>();
const verifiedFailures = new Map<string, Array<{ name: string; href: string }>>();
const failuresDownloadId = new Map<string, number>(); // folder key → last FAILURES.txt download id

/** Clear verifier state at the start of an export's Files phase. */
function resetVerifyState(): void {
  pendingVerify.clear();
  verifiedFailures.clear();
  failuresDownloadId.clear();
}

/**
 * chrome.downloads.onChanged handler. Confirms a completed download is the real
 * file, not a request-access page. Register once at the service-worker top level
 * (so the event can wake the worker after the export's await chain has ended):
 *
 *   downloads.onChanged.addListener(d => void verifyDownloadOnChanged({ downloads, log }, d));
 */
export async function verifyDownloadOnChanged(
  deps: { downloads: DownloadsApi; log?: (m: string) => void },
  delta: { id: number; state?: { current?: string } },
): Promise<void> {
  if (delta.state?.current !== 'complete') return;
  const meta = pendingVerify.get(delta.id);
  if (!meta) return;                 // not one of ours, or already handled
  pendingVerify.delete(delta.id);

  let item: chrome.downloads.DownloadItem | undefined;
  try {
    const r = await deps.downloads.search({ id: delta.id });
    item = Array.isArray(r) ? r[0] : undefined;
  } catch { return; }
  // A non-markup file that arrived as anything but text/html is the real file.
  if (!item || (item.mime || '').toLowerCase() !== 'text/html') return;

  // Request-access page saved under the file's name: delete file + history row.
  try { await deps.downloads.removeFile(delta.id); } catch { /* may already be gone */ }
  try { await deps.downloads.erase({ id: delta.id }); } catch { /* noop */ }

  const key = `${meta.exportFolder} ${meta.chatFolder}`;
  const list = verifiedFailures.get(key) || [];
  list.push({ name: meta.name, href: meta.href });
  verifiedFailures.set(key, list);
  deps.log?.(`attachment no-access: removed request-access page for ${meta.name}`);
  await writeFailuresFile(deps.downloads, meta.exportFolder, meta.chatFolder, list, key);
}

/**
 * Hand a set of SharePoint document attachments to chrome.downloads. Returns a
 * tally. No-op safe: an empty `items` returns a zero summary.
 *
 * FIRE-AND-FORGET, and it must stay that way. On a remote URL, download() does
 * NOT resolve until the transfer actually starts (a download-manager slot plus a
 * network round-trip) — about a second each. Awaiting it therefore serialized
 * dispatch at the network rate (~1 file/s), and a few-hundred-file set ran longer
 * than the MV3 service worker stayed alive: it was evicted mid-phase, so the
 * export never reached its terminal state and no history row was written. Here we
 * call download() back-to-back WITHOUT awaiting, so every file is queued in
 * milliseconds and the export completes immediately; the browser then streams
 * each one server->disk in its own context, outliving the worker.
 *
 * The id-promise is not awaited but IS observed: a non-markup download is
 * recorded for post-completion verification (see verifyDownloadOnChanged), and a
 * rejection is logged so it isn't an unhandled promise.
 */
export async function downloadAttachments(
  items: AttachmentCandidate[],
  exportFolder: string,
  deps: AttachmentDownloadDeps,
  signal: AbortSignal,
): Promise<AttachmentDownloadSummary> {
  const summary: AttachmentDownloadSummary = { total: items.length, saved: 0, links: 0, failed: 0, failures: [] };
  if (!items.length) return summary;

  resetVerifyState();
  deps.log?.(`attachment phase: ${items.length} candidate(s)`);

  let viaAspx = 0;  // queued files routed through the download.aspx handler (markup types)
  for (let i = 0; i < items.length; i++) {
    if (signal.aborted) break;
    const item = items[i];
    // Defense-in-depth: only SharePoint-family hosts reach chrome.downloads.
    // collectDocumentAttachments already gates on this, so this is a
    // belt-and-suspenders check, not the primary filter.
    if (!isSharePointFileHost(item.href)) {
      summary.links++;
    } else {
      const name = item.name || fileNameFromUrl(item.href) || item.href;
      const rawName = item.name || fileNameFromUrl(item.href) || 'attachment';
      const url = toDownloadUrl(item.href, item.itemid);
      const viaDownloadAspx = isDownloadAspxUrl(url);
      try {
        // Fire, do NOT await (see the function note). The download is registered
        // the instant this is called; we observe the returned id only to enroll
        // non-markup files for the request-access verification sweep, and attach
        // a catch so a rejection isn't an unhandled promise.
        Promise.resolve(deps.downloads.download({
          url,
          filename: buildAttachmentPath(exportFolder, item.chatFolder, rawName),
          saveAs: false,
          conflictAction: 'uniquify',
        }))
          .then(id => {
            // Markup (download.aspx-routed) legitimately arrives as text/html, so
            // it can't be verified by Content-Type; only enroll the rest.
            if (typeof id === 'number' && !viaDownloadAspx) {
              pendingVerify.set(id, { name, href: item.href, chatFolder: item.chatFolder, exportFolder });
            }
          })
          .catch(e => deps.log?.(`attachment download error: ${name}: ${errMsg(e)}`));
        summary.saved++;
        if (viaDownloadAspx) viaAspx++;
      } catch (e) {
        // Synchronous throw (e.g. a malformed options object) — rare.
        summary.failed++;
        summary.failures.push({ name, reason: errMsg(e) });
      }
    }
    const processed = summary.saved + summary.links + summary.failed;
    deps.onProgress?.(processed, items.length);
    if (processed % 50 === 0 || processed === items.length) deps.log?.(`attachment dispatch: ${processed}/${items.length}`);
  }

  deps.log?.(`attachment download: ${summary.saved} queued (${viaAspx} via download.aspx), ${summary.links} link(s), ${summary.failed} failed of ${summary.total}`);
  return summary;
}

/**
 * Single-chat convenience: collect this chat's document attachments and stream
 * them under `<exportFolder>/attachments/`.
 */
export async function downloadChatAttachments(
  messages: ExportMessage[],
  exportFolder: string,
  deps: AttachmentDownloadDeps,
  signal: AbortSignal,
): Promise<AttachmentDownloadSummary> {
  const items = collectDocumentAttachments(messages).map(c => ({ ...c, chatFolder: '' }));
  return downloadAttachments(items, exportFolder, deps, signal);
}

/**
 * Counts-only projection for the status wire. Drops the `failures` array (which
 * carries attachment file names) so it is never broadcast over runtime
 * messaging or persisted into the active-export snapshot in chrome.storage.
 */
export function toFilesSummaryWire(s: AttachmentDownloadSummary): { total: number; saved: number; links: number; failed: number } {
  return { total: s.total, saved: s.saved, links: s.links, failed: s.failed };
}
