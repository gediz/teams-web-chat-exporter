// Document attachment download (the "Files" toggle).
//
// Streams non-image SharePoint/OneDrive-for-Business file attachments
// (PDF/DOCX/XLSX/ZIP/video/...) to disk via chrome.downloads, into an
// `attachments/` folder INSIDE the export's package folder (the builders save
// the export itself into the same folder in package mode, so the whole export
// is one folder in Downloads). Inline images are handled separately by
// downloadImages; this is the paperclip / shared-file documents that are
// otherwise kept as links.
//
// Mechanism (verified live against a real tenant): the converter stores a
// document's absolute SharePoint URL (the file's webDavUrl/objectUrl) as
// Attachment.href. SharePoint serves that URL's raw bytes to an authenticated
// caller (HTTP 200, the file content), so we hand it straight to
// chrome.downloads.download. The browser fetches it in its OWN network context,
// sending the user's first-party SharePoint cookies and following any redirect,
// streaming server->disk with zero extension-held bytes. Most types serve raw
// bytes directly (`?download=1` force-download hint); renderable-markup types
// (html/svg/xml/...) go through the _layouts/download.aspx handler by UniqueId so
// the active-content sandbox interstitial isn't saved in their place. (Routing
// EVERY type through download.aspx, as Teams' web UI does, was tried and reverted:
// it can't reach sharing-link or other-owner files without the CORS-walled shares
// API.) See toDownloadUrl.
//
// A file the user cannot open can't be detected upfront (no reliable access
// oracle exists — see tce-debug/ACCESS_DETECTION_FINDINGS.md); it comes down as a
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
import { revokeDownloadUrl, textToDownloadUrl } from './builders';

// The subset of chrome.downloads we use. Mirrors the deps-injection style in
// download.ts so the module stays unit-testable without a real browser.
// search/removeFile/erase back the post-download request-access cleanup;
// cancel backs the STOP path (stop the actual transfers, not just dispatch).
type DownloadsApi = Pick<typeof chrome.downloads, 'download' | 'search' | 'removeFile' | 'erase' | 'cancel'>;

// Result of resolving a file's sharing link to a pre-authenticated download
// URL (via the content-script shares resolver). `downloadUrl` is short-lived
// and self-authenticating — chrome.downloads can fetch it with no session
// cookie. `blocksDownload` is the upfront access oracle.
export interface ShareResolveOutcome {
  downloadUrl?: string;
  blocksDownload?: boolean;
}

export interface AttachmentDownloadDeps {
  downloads: DownloadsApi;
  // Called as candidates SETTLE (link kept, dispatch failed, or the download
  // reached a terminal state and was verified). `done` counts settled
  // candidates — files whose outcome is known — not dispatches, so the
  // popup's progress is real and monotonic. `total` is the candidate count.
  onProgress?: (done: number, total: number) => void;
  log?: (msg: string) => void;
  // Optional: resolve a file's sharing link to a pre-authenticated download
  // URL (background wires this to a content-script RPC). When present and a
  // candidate has a shareUrl, the resolved URL is used instead of the raw
  // SharePoint path, which downloads even when the host session cookie is
  // stale. Returns null on any failure, and the raw-URL path is used as a
  // fallback so nothing regresses.
  resolveShare?: (shareUrl: string) => Promise<ShareResolveOutcome | null>;
}

export interface AttachmentDownloadSummary {
  total: number;     // SharePoint document candidates found (deduped by href)
  saved: number;     // downloads that COMPLETED and passed verification
                     // (request-access junk pages are verified out, see below)
  links: number;     // skipped without a download attempt (host gate) — kept as a link
  failed: number;    // dispatch failures (sync or async) + verified failures
                     // (no-access pages, server-interrupted transfers)
  cancelled: number; // stopped before finishing (user cancel / STOP_EXPORT)
  failures: Array<{ name: string; reason: string }>;
}

export interface AttachmentCandidate {
  href: string;        // the file's absolute SharePoint URL (download source)
  name: string;        // best-known display name (from Attachment.label)
  chatFolder: string;  // per-chat subfolder under the export root ('' for single-chat)
  itemid?: string;     // stable file GUID (== listItemUniqueId); the download.aspx UniqueId for markup types
  shareUrl?: string;   // SharePoint sharing link, for the resolver (see resolveShare)
  resolvedUrl?: string; // set by the resolve phase: pre-authenticated download URL
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
 * True when the file is itself renderable markup, so its bytes legitimately
 * arrive as text/html — a text/html download of it is the real file, not a junk
 * page. Keyed off the file's own URL (extension), so the verifier classifies it
 * by file type, not by which download URL form we happened to build.
 */
function isMarkupFile(href: string): boolean {
  try { return SANDBOXED_EXT.test(new URL(href).pathname); } catch { return false; }
}

/**
 * Map a file's absolute SharePoint URL to the URL chrome.downloads should fetch.
 *
 * Most types serve their raw bytes directly to an authenticated caller, so we
 * append `?download=1` (SharePoint's force-download hint) and hand them over as
 * is — the proven path that pulls PDFs/zips/Office/video correctly, including
 * files reached through a sharing link (where ?download=1 lets SharePoint do its
 * own redirect).
 *
 * Renderable-markup types ({@link SANDBOXED_EXT}) instead hit the active-content
 * sandbox interstitial, so route those through the `_layouts/download.aspx`
 * handler addressed by the file's UniqueId on its own site collection, with the
 * native `Translate=false&ApiVersion=2.0` parameters. Falls back to the
 * absolute-path SourceUrl form when no UniqueId is known.
 *
 * NOTE: download.aspx?UniqueId was tried for EVERY type (it's what Teams' web UI
 * uses) and REGRESSED — it can't reach sharing-link hrefs (siteCollectionBase
 * can't derive the site) or other-owner files (no site permission); Teams only
 * works there by resolving each file through the CORS-walled /_api/v2.0/shares
 * API first. So markup-only download.aspx, `?download=1` for the rest, stays.
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
export function collectDocumentAttachments(messages: ExportMessage[]): Array<{ href: string; name: string; itemid?: string; shareUrl?: string }> {
  const seen = new Map<string, { href: string; name: string; itemid?: string; shareUrl?: string }>();
  const out: Array<{ href: string; name: string; itemid?: string; shareUrl?: string }> = [];
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
        // Same file twice: keep the first, but adopt an itemid/shareUrl if the
        // first occurrence lacked one (so the join key is never lost to order).
        if (!dup.itemid && a.itemid) dup.itemid = a.itemid;
        if (!dup.shareUrl && a.shareUrl) dup.shareUrl = a.shareUrl;
        continue;
      }
      const entry = { href, name: (a.label || '').toString(), itemid: a.itemid, shareUrl: a.shareUrl };
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
 * Write a FAILURES.txt for this chat's attachments folder, in two sections:
 *   - files that could not be downloaded (no access, moved/deleted, an interrupted
 *     transfer) — the system couldn't retrieve them;
 *   - files cancelled mid-download — stopped by the user (or an aborted export),
 *     so no file landed, listed separately so a cancel is never silently lost.
 * Either section may be empty; the file is written whenever at least one entry
 * exists. Each line carries the file's link so the user can open it in Teams.
 *
 * `overwrite` so each (debounced) call replaces the file with the full accumulated
 * lists. The previous write's history row is erased first (erase drops the shelf
 * record, not the file) so repeated rewrites leave a single FAILURES.txt entry.
 *
 * The URL comes from textToDownloadUrl, which yields a blob: URL where
 * createObjectURL exists (the Firefox background page) and a data: URL fallback
 * in the Chrome MV3 worker (which lacks createObjectURL). A bare data: URL is
 * rejected by Firefox's download path, which is why FAILURES.txt silently failed
 * to write there.
 */
async function writeFailuresFile(
  downloads: DownloadsApi,
  exportFolder: string,
  chatFolder: string,
  failures: Array<{ name: string; href: string }>,
  cancels: Array<{ name: string; href: string }>,
  key: string,
): Promise<void> {
  const list = (entries: Array<{ name: string; href: string }>) =>
    entries.map(e => `${e.name}\n${e.href}\n`).join('\n');
  const sections: string[] = [];
  if (failures.length) {
    sections.push(
      'These files were shared in this chat but could not be downloaded. You may not\n' +
      'have access, the file was moved or deleted, or the download was interrupted.\n' +
      'Open each one in Teams using its link to check.\n\n' +
      list(failures));
  }
  if (cancels.length) {
    sections.push(
      '--- Cancelled during download (not saved) ---\n' +
      'These were stopped before finishing; re-export to retrieve them.\n\n' +
      list(cancels));
  }
  if (!sections.length) return;
  const url = textToDownloadUrl(sections.join('\n'), 'text/plain;charset=utf-8');
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
    // Revoke after the download has read the (tiny) blob. No-op for data: URLs.
    setTimeout(() => revokeDownloadUrl(url), 60_000);
  } catch { revokeDownloadUrl(url); /* best-effort; the file just stays out of the export */ }
}

// ── Post-download verification (could-not-download -> FAILURES.txt) ─────────
//
// Fire-and-forget downloads can't inspect the response, so we verify after the
// fact via downloads.onChanged. A file the user can't retrieve shows up two ways:
//   - NON-markup that COMPLETES as text/html: SharePoint's "request access" HTML
//     page saved under the file's name (a real pdf/zip/video is never text/html).
//     Remove it from disk + history.
//   - A download the server INTERRUPTS with a concrete reason (no access, deleted,
//     a network drop): no real file landed, listed under FAILURES.txt's failures
//     section. An interrupt with NO reason is a cancel (the user or an aborted
//     export), listed under FAILURES.txt's separate "cancelled" section instead.
// Either way the file, with its link, is recorded so a missing file never leaves
// zero trace.
//
// Markup that COMPLETES is kept: routed through download.aspx it legitimately
// arrives as text/html, so completion means the real file (an inaccessible markup
// file is interrupted with a JSON error instead, caught by the interrupted path).
//
// State is in-memory and best-effort. The stream of onChanged events keeps the
// MV3 worker alive while downloads land; if a long gap evicts it, late checks
// are simply skipped (no upfront data loss).

// How a settled download is classified for the run's tally. 'saved' is a
// clean verified complete; 'failed' covers no-access pages, server interrupts,
// and downloads that vanished; 'cancelled' is a user/STOP cancel.
export type SettleKind = 'saved' | 'failed' | 'cancelled';

interface VerifyMeta {
  name: string; href: string; chatFolder: string; exportFolder: string; markup?: boolean;
  // Set by the dispatching run so it can tally outcomes and know when every
  // download settled. `reason` carries the InterruptReason on a failure so the
  // run can decide whether it is worth a retry. Optional: entries without it
  // (a run that already returned) still get the junk-page cleanup below.
  onSettled?: (kind: SettleKind, reason?: string) => void;
}
// All maps below are keyed so concurrent per-tab exports can't clobber each
// other: pendingVerify by (unique) download id, the rest by the folder key
// `${exportFolder} ${chatFolder}` (exportFolder embeds a per-export
// timestamp). There is deliberately NO blanket reset between runs — an
// earlier version cleared everything at each Files-phase start, which wiped
// a concurrent export's pending verifications.
const pendingVerify = new Map<number, VerifyMeta>();
const verifiedFailures = new Map<string, Array<{ name: string; href: string }>>(); // could-not-download
const verifiedCancels = new Map<string, Array<{ name: string; href: string }>>();  // stopped before finishing
const failuresDownloadId = new Map<string, number>(); // folder key → last FAILURES.txt download id
const failuresWriteTimer = new Map<string, ReturnType<typeof setTimeout>>(); // folder key → pending write

// Failures surface in bursts (many downloads settle within a few hundred ms), so
// the FAILURES.txt write is debounced. Writing per-failure dispatched one download
// each; overlapping downloads to the same filename can't silently overwrite a file
// that an earlier one still holds open, so Chrome falls back to a Save-As prompt —
// the user saw several. Coalescing the burst into one trailing write removes both
// the prompt storm and the race. The window must outlast a settle burst but stay
// well under the MV3 service-worker idle timeout so the trailing write always runs.
const FAILURES_WRITE_DEBOUNCE_MS = 1500;

// While an active Files phase is running, FAILURES.txt writes are suppressed so
// the file is not created mid-phase (confusing, and it would show entries a
// retry is about to remove). The in-memory lists still update; the run flushes
// once at the end. Late verifier events (after the run returns) are not
// suppressed, so an eviction-and-late-settle still persists its record.
let failuresWriteSuppressed = false;

/** Write FAILURES.txt once for every folder under exportFolder that has any
 *  recorded failure/cancel. Called at the true end of a Files phase. */
async function flushFailures(deps: { downloads: DownloadsApi }, exportFolder: string): Promise<void> {
  const folders = new Set<string>();
  for (const key of verifiedFailures.keys()) if (key.startsWith(exportFolder + ' ')) folders.add(key.slice(exportFolder.length + 1));
  for (const key of verifiedCancels.keys()) if (key.startsWith(exportFolder + ' ')) folders.add(key.slice(exportFolder.length + 1));
  for (const chatFolder of folders) {
    const key = `${exportFolder} ${chatFolder}`;
    await writeFailuresFile(deps.downloads, exportFolder, chatFolder, verifiedFailures.get(key) || [], verifiedCancels.get(key) || [], key);
  }
}

/**
 * (Re)schedule the debounced FAILURES.txt write for a chat folder. The in-memory
 * lists update immediately at the call site; this coalesces a burst of records
 * into a single download once it goes quiet (see FAILURES_WRITE_DEBOUNCE_MS).
 */
function scheduleFailuresWrite(deps: { downloads: DownloadsApi }, meta: VerifyMeta): void {
  if (failuresWriteSuppressed) return; // flushed once at phase end instead
  const key = `${meta.exportFolder} ${meta.chatFolder}`;
  const prev = failuresWriteTimer.get(key);
  if (prev) clearTimeout(prev);
  failuresWriteTimer.set(key, setTimeout(() => {
    failuresWriteTimer.delete(key);
    void writeFailuresFile(
      deps.downloads, meta.exportFolder, meta.chatFolder,
      verifiedFailures.get(key) || [], verifiedCancels.get(key) || [], key,
    );
  }, FAILURES_WRITE_DEBOUNCE_MS));
}

/** Record a file the system could not download into this chat's FAILURES.txt. */
function recordFailure(deps: { downloads: DownloadsApi; log?: (m: string) => void }, meta: VerifyMeta): void {
  const key = `${meta.exportFolder} ${meta.chatFolder}`;
  const list = verifiedFailures.get(key) || [];
  list.push({ name: meta.name, href: meta.href });
  verifiedFailures.set(key, list);
  scheduleFailuresWrite(deps, meta);
}

/** Remove a previously-recorded failure for a file (by href), used when a retry
 *  of a transiently-failed download eventually succeeds so its stale entry does
 *  not remain in FAILURES.txt. chatFolder is unknown here, so all folder lists
 *  are checked. */
function removeRecordedFailure(item: { href: string }, exportFolder: string): void {
  for (const [key, list] of verifiedFailures) {
    if (!key.startsWith(exportFolder + ' ')) continue;
    const next = list.filter(e => e.href !== item.href);
    if (next.length !== list.length) verifiedFailures.set(key, next);
  }
}

/** Record a file stopped mid-download (user cancel or aborted export) into the
 *  separate "cancelled" section, so a cancel is logged without being mislabelled
 *  a failure. */
function recordCancelled(deps: { downloads: DownloadsApi; log?: (m: string) => void }, meta: VerifyMeta): void {
  const key = `${meta.exportFolder} ${meta.chatFolder}`;
  const list = verifiedCancels.get(key) || [];
  list.push({ name: meta.name, href: meta.href });
  verifiedCancels.set(key, list);
  scheduleFailuresWrite(deps, meta);
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
  const st = delta.state?.current;
  if (st !== 'complete' && st !== 'interrupted') return;
  const meta = pendingVerify.get(delta.id);
  if (!meta) return;                 // not one of ours, or already handled
  pendingVerify.delete(delta.id);
  // Every path below settles this download exactly once so the dispatching
  // run's tally always adds up to its candidate count.
  const settle = (kind: SettleKind, reason?: string) => { try { meta.onSettled?.(kind, reason); } catch { /* run already gone */ } };

  let item: chrome.downloads.DownloadItem | undefined;
  try {
    const r = await deps.downloads.search({ id: delta.id });
    item = Array.isArray(r) ? r[0] : undefined;
  } catch { settle('failed'); return; }
  if (!item) { settle('failed'); return; } // erased from history — no file landed

  // Classify by the item's ACTUAL state, not the delta's: the settlement
  // poll feeds synthetic deltas that can lag or guess wrong (an id absent
  // from one batch-search snapshot may be alive and well). A still-running
  // download is not settled at all — re-arm it and wait for the real event.
  const actual = item.state === 'complete' || item.state === 'interrupted' ? item.state : undefined;
  if (!actual) {
    pendingVerify.set(delta.id, meta);
    return;
  }

  if (actual === 'interrupted') {
    // The download left no real file. A genuine "could not download" always
    // carries a concrete InterruptReason (NETWORK_FAILED, SERVER_BAD_CONTENT, a
    // cross-owner SERVER_FAILED, a network drop mid-transfer): those we record.
    // An interrupt with NO reason means the user/client stopped it (a manual
    // cancel or an aborted export), which is not a failure: it goes in the
    // separate "cancelled" section instead. The two engines encode a manual
    // cancel differently: Chrome sets error to USER_CANCELED, Firefox leaves
    // error empty. Treat both as a cancel. (canResume=true is NOT a promise it
    // will finish — a 44MB transfer that NETWORK_FAILED part-way had
    // canResume=true yet never resumed and left no file, so a real error is still
    // recorded.) We do NOT removeFile/erase: a genuine partial may yet auto-resume
    // and deleting it would break that.
    if (!item.error || item.error === 'USER_CANCELED') {
      deps.log?.(`attachment cancelled: ${meta.name}`);
      recordCancelled(deps, meta);
      settle('cancelled');
      return;
    }
    deps.log?.(`attachment no-access: ${meta.name} (${item.error})`);
    recordFailure(deps, meta);
    settle('failed', item.error);
    return;
  }

  // Completed. A NON-markup file that arrived as text/html is a request-access
  // page saved under the file's name (a real pdf/zip/video is never text/html).
  // Markup is legitimately text/html, so a completed markup download is kept.
  if (meta.markup || (item.mime || '').toLowerCase() !== 'text/html') { settle('saved'); return; }
  try { await deps.downloads.removeFile(delta.id); } catch { /* may already be gone */ }
  try { await deps.downloads.erase({ id: delta.id }); } catch { /* noop */ }
  deps.log?.(`attachment no-access: removed request-access page for ${meta.name}`);
  recordFailure(deps, meta);
  // A request-access page is a permanent access failure, not a transient one.
  settle('failed', 'REQUEST_ACCESS_PAGE');
}

// Interrupt reasons worth retrying: transient network/server conditions where a
// second attempt (after backoff, and after a bandwidth-hogging sibling
// download finishes) plausibly succeeds. NOT retried: access denials
// (SERVER_UNAUTHORIZED/FORBIDDEN), the request-access page, or a user cancel.
function isTransientReason(reason?: string): boolean {
  if (!reason) return false;
  // Network-layer conditions where a retry after backoff plausibly succeeds
  // (this is the NETWORK_FAILED-during-a-burst case seen live). Deliberately
  // NOT SERVER_BAD_CONTENT (a bad response body, usually a permanent error
  // page) nor SERVER_UNAUTHORIZED/FORBIDDEN (access denials).
  return /NETWORK_FAILED|NETWORK_TIMEOUT|NETWORK_DISCONNECTED|NETWORK_SERVER_DOWN|SERVER_FAILED|SERVER_UNREACHABLE|SERVER_NO_RANGE|CONNECTION|TIMED_OUT|CRASH/i.test(reason);
}

/**
 * Hand a set of SharePoint document attachments to chrome.downloads, then wait
 * for every one of them to SETTLE (terminal state observed and verified), so
 * the caller's completion broadcast / history row / auto-show all describe an
 * export whose files are really done. Returns the settled tally. No-op safe:
 * an empty `items` returns a zero summary.
 *
 * Two dispatch modes:
 *   - Chrome (default): fire-and-forget. download() does not resolve until the
 *     transfer starts (~1s each), so awaiting it serially outran the MV3 worker's
 *     lifetime and evicted it mid-phase. We queue every file in milliseconds and
 *     let Chrome's per-host concurrency cap pace them; the id is observed only to
 *     enroll the download for verification and settlement tallying.
 *   - Firefox (opts.maxConcurrent): bounded pool. Firefox does NOT cap concurrent
 *     downloads, so firing all at once made SharePoint throttle the burst and
 *     return an error page for every file. Each worker awaits its download's
 *     terminal state before pulling the next, capping in-flight transfers.
 *     Firefox's persistent background page makes the long wait safe.
 *
 * Settlement: after dispatch, a wait loop runs until every candidate is
 * accounted for (link / dispatch failure / verified saved / failed /
 * cancelled). Outcomes arrive through the top-level downloads.onChanged
 * verifier (VerifyMeta.onSettled); a slow batched downloads.search backs up
 * missed events and doubles as the MV3 keep-alive. On abort (STOP_EXPORT) the
 * run cancels its still-running downloads — each lands as a reason-less
 * interrupt that the verifier records in FAILURES.txt's cancelled section.
 */
export async function downloadAttachments(
  items: AttachmentCandidate[],
  exportFolder: string,
  deps: AttachmentDownloadDeps,
  signal: AbortSignal,
  opts: { maxConcurrent?: number } = {},
): Promise<AttachmentDownloadSummary> {
  const summary: AttachmentDownloadSummary = { total: items.length, saved: 0, links: 0, failed: 0, cancelled: 0, failures: [] };
  if (!items.length) return summary;
  // Suppress mid-phase FAILURES.txt writes; flush once at the end (below).
  failuresWriteSuppressed = true;

  deps.log?.(`attachment phase: ${items.length} candidate(s)` + (opts.maxConcurrent ? `, max ${opts.maxConcurrent} concurrent` : ''));

  let viaAspx = 0;    // queued files routed through the download.aspx handler
  let dispatched = 0; // candidates handed to chrome.downloads (log cadence only)

  // Run-local settlement accounting. Every candidate settles exactly once:
  // links and dispatch failures settle at dispatch time; enrolled downloads
  // settle when the verifier classifies their terminal state. onProgress
  // reports settled counts, so the popup's numbers are real and monotonic.
  const expected = items.length;
  let settledCount = 0;
  const pendingIds = new Set<number>(); // this run's enrolled, not-yet-settled ids
  let notifySettle: (() => void) | undefined; // wakes the settlement loop early
  // Set just before this function returns. Late arrivals (a dispatch promise
  // resolving after a stall-break/abort, a transfer settling afterwards) must
  // not mutate the escaped summary or re-broadcast 'downloading-files' AFTER
  // the export flow's terminal 'complete' — a late broadcast would re-insert
  // a non-terminal phase into activeExports and block the tab's next export.
  let runEnded = false;
  const report = () => {
    if (runEnded) return;
    deps.onProgress?.(settledCount, expected);
    if (settledCount % 50 === 0 || settledCount === expected) deps.log?.(`attachment settle: ${settledCount}/${expected}`);
  };
  const settledOne = () => {
    if (runEnded) return;
    settledCount++;
    report();
    notifySettle?.();
  };
  // Retry state. A transient download failure (NETWORK_FAILED etc., common when
  // a huge sibling download saturates the connection) is re-dispatched instead
  // of settling: the item stays "pending" so the settlement wait naturally
  // covers the retry, and no accounting has to be un-done. The verifier already
  // recorded the failed attempt in FAILURES.txt; on eventual success we remove
  // that record. Access denials and cancels are never retried.
  const enrolledItem = new Map<number, AttachmentCandidate>(); // id -> candidate (for retry)
  const retryCount = new Map<string, number>();                 // href -> attempts used
  const MAX_RETRIES = 2;
  const RETRY_BACKOFF_MS = 3_000;
  const keyOf = (it: AttachmentCandidate) => it.href;

  const onDownloadSettled = (id: number, kind: SettleKind, reason?: string) => {
    if (runEnded) { pendingIds.delete(id); enrolledItem.delete(id); return; }
    if (!pendingIds.delete(id)) return; // event+poll double delivery guard
    const item = enrolledItem.get(id);
    enrolledItem.delete(id);
    if (kind === 'failed' && item && !signal.aborted && isTransientReason(reason)
        && (retryCount.get(keyOf(item)) ?? 0) < MAX_RETRIES) {
      // Re-dispatch after a backoff; do NOT settle (item stays pending).
      retryCount.set(keyOf(item), (retryCount.get(keyOf(item)) ?? 0) + 1);
      deps.log?.(`attachment retry (${reason}): ${item.name} [attempt ${retryCount.get(keyOf(item))}]`);
      scheduleRetry(item);
      return;
    }
    if (kind === 'saved') {
      summary.saved++;
      // A prior attempt of a retried file left a stale FAILURES.txt entry.
      if (item && (retryCount.get(keyOf(item)) ?? 0) > 0) removeRecordedFailure(item, exportFolder);
    } else if (kind === 'cancelled') summary.cancelled++;
    else summary.failed++;
    settledOne();
  };

  // Re-resolve + re-dispatch a transiently-failed item after a backoff. Any
  // failure to even re-dispatch settles it as failed (never leaves it pending).
  const scheduleRetry = (item: AttachmentCandidate) => {
    const attempt = retryCount.get(keyOf(item)) ?? 1;
    setTimeout(async () => {
      if (runEnded || signal.aborted) { if (!runEnded) { summary.failed++; settledOne(); } return; }
      item.resolvedUrl = undefined; // force a fresh pre-authenticated URL
      let spec: NonNullable<ReturnType<typeof prep>> | undefined;
      try {
        const p = await prepAsync(item);
        if (p.kind !== 'spec') { summary.failed++; settledOne(); return; }
        spec = p.spec;
      } catch { summary.failed++; settledOne(); return; }
      if (runEnded || signal.aborted) { if (!runEnded) { summary.failed++; settledOne(); } return; }
      let id: number | undefined;
      try {
        id = await Promise.resolve(deps.downloads.download({ url: spec.url, filename: spec.filename, saveAs: false, conflictAction: 'uniquify' }));
      } catch (e) { dispatchFailed(spec.name, errMsg(e)); return; }
      if (typeof id !== 'number') { dispatchFailed(spec.name, 'no download id'); return; }
      enroll(id, item, spec.name, spec.markup);
    }, RETRY_BACKOFF_MS * attempt);
  };
  const dispatchFailed = (name: string, reason: string) => {
    if (runEnded) return;
    summary.failed++;
    summary.failures.push({ name, reason });
    settledOne();
  };

  // Broadcast the phase at 0/N before any dispatch: the popup shows honest
  // progress from the first moment, and the STOP handler's Files-phase gate
  // (which keys on phase === 'downloading-files') is armed for the WHOLE
  // phase, not just from the first settlement.
  report();

  // Captured before dispatch so the settlement loop's batched search covers
  // every download of this run — on the Firefox pool path dispatch itself
  // takes minutes, so capturing at loop start would miss early downloads.
  const startedAfter = new Date(Date.now() - 60_000).toISOString();

  // Resolve is done JUST-IN-TIME per item (inside prepAsync, right before that
  // item's download() call), never all-upfront — the pre-authenticated URL is
  // short-lived, so it must be fetched seconds before use, not minutes.
  const needResolve = Boolean(deps.resolveShare) && items.some(i => i.shareUrl);

  // Per-item prep (raw-URL transport, no resolve). Returns the download spec, or
  // null for a host-gated link. `markup` is the file TYPE (so the verifier knows
  // text/html is the real file), independent of the download URL form.
  const prep = (item: AttachmentCandidate) => {
    if (!isSharePointFileHost(item.href)) return null;
    const name = item.name || fileNameFromUrl(item.href) || item.href;
    const rawName = item.name || fileNameFromUrl(item.href) || 'attachment';
    const url = item.resolvedUrl || toDownloadUrl(item.href, item.itemid);
    const viaDownloadAspx = isDownloadAspxUrl(url);
    const markup = isMarkupFile(item.href);
    return { name, url, viaDownloadAspx, markup, filename: buildAttachmentPath(exportFolder, item.chatFolder, rawName) };
  };

  // Async prep: resolve the sharing link to a fresh pre-authenticated URL right
  // before dispatch. Returns a kind:
  //   'link' — host-gated, kept as a link (counted as such);
  //   'spec' — a download spec (resolved URL when available, else raw URL).
  // Any resolve failure leaves the raw-URL spec, so nothing regresses.
  type Prepped =
    | { kind: 'link' }
    | { kind: 'spec'; spec: NonNullable<ReturnType<typeof prep>> };
  const prepAsync = async (item: AttachmentCandidate): Promise<Prepped> => {
    const base = prep(item);
    if (!base) return { kind: 'link' };
    if (needResolve && deps.resolveShare && item.shareUrl && !item.resolvedUrl) {
      try {
        const r = await deps.resolveShare(item.shareUrl);
        // Adopt the resolved pre-auth URL only when the role does NOT block
        // download AND the URL is a gated SharePoint host (the URL is
        // server-controlled, so gate it like the raw href). blocksDownload is
        // NOT a hard skip: its meaning ("server refuses" vs "UI hides button")
        // is uncertain across clouds, so a blocked file still falls through to
        // the raw-URL path, which may work via the session cookie — never worse
        // than not resolving at all.
        if (r?.downloadUrl && r.blocksDownload !== true && isSharePointFileHost(r.downloadUrl)) {
          item.resolvedUrl = r.downloadUrl;
          const url = r.downloadUrl;
          return { kind: 'spec', spec: { ...base, url, viaDownloadAspx: isDownloadAspxUrl(url) } };
        }
      } catch { /* leave the raw-URL spec */ }
    }
    return { kind: 'spec', spec: base };
  };
  const enroll = (id: number, item: AttachmentCandidate, name: string, markup: boolean) => {
    // pendingIds BEFORE pendingVerify: the terminal event can fire the moment
    // the entry exists, and onDownloadSettled only counts ids it knows.
    pendingIds.add(id);
    enrolledItem.set(id, item);
    pendingVerify.set(id, {
      name, href: item.href, chatFolder: item.chatFolder, exportFolder, markup,
      onSettled: (kind, reason) => onDownloadSettled(id, kind, reason),
    });
    if (signal.aborted) {
      // A STOP raced this enrollment: stop the transfer now; the verifier
      // records the interrupt in the cancelled section.
      try { void Promise.resolve(deps.downloads.cancel(id)).catch(() => { /* already terminal */ }); } catch { /* noop */ }
    }
  };
  const logDispatch = () => {
    dispatched++;
    if (dispatched % 50 === 0 || dispatched === items.length) deps.log?.(`attachment dispatch: ${dispatched}/${items.length}`);
  };

  // Fire-and-forget dispatch of one prepared spec (Chrome). Self-accounting via
  // enroll/dispatchFailed; never awaits the download.
  const fireAndForget = (item: AttachmentCandidate, spec: NonNullable<ReturnType<typeof prep>>) => {
    try {
      Promise.resolve(deps.downloads.download({ url: spec.url, filename: spec.filename, saveAs: false, conflictAction: 'uniquify' }))
        .then(id => {
          if (typeof id === 'number') { enroll(id, item, spec.name, spec.markup); if (spec.viaDownloadAspx) viaAspx++; }
          else dispatchFailed(spec.name, 'no download id');
        })
        .catch(e => {
          // An async rejection is a real failure: never enrolled, so without
          // this the file would silently vanish from every tally bucket.
          deps.log?.(`attachment download error: ${spec.name}: ${errMsg(e)}`);
          dispatchFailed(spec.name, errMsg(e));
        });
    } catch (e) { dispatchFailed(spec.name, errMsg(e)); }
    logDispatch();
  };

  if (opts.maxConcurrent && opts.maxConcurrent > 0) {
    // BOUNDED CONCURRENCY (Firefox): cap in-flight transfers so we never burst
    // SharePoint. A worker resolves (JIT), dispatches, then awaits its download's
    // terminal state before the next — so the resolved URL is seconds old.
    let cursor = 0;
    const worker = async (): Promise<void> => {
      while (cursor < items.length && !signal.aborted) {
        const item = items[cursor++];
        const p = await prepAsync(item);
        if (p.kind === 'link') { summary.links++; settledOne(); continue; }
        const spec = p.spec;
        let id: number | undefined;
        try {
          id = await Promise.resolve(deps.downloads.download({ url: spec.url, filename: spec.filename, saveAs: false, conflictAction: 'uniquify' }));
        } catch (e) { dispatchFailed(spec.name, errMsg(e)); continue; }
        if (typeof id !== 'number') { dispatchFailed(spec.name, 'no download id'); continue; }
        enroll(id, item, spec.name, spec.markup);
        if (spec.viaDownloadAspx) viaAspx++;
        logDispatch();
        await awaitTerminal(deps, id, signal);
      }
    };
    await Promise.all(Array.from({ length: opts.maxConcurrent }, () => worker()));
  } else if (needResolve) {
    // CHROME WITH RESOLVE: a bounded pool resolves each URL just-in-time, then
    // fires its download without awaiting. Bounds resolve concurrency (off the
    // shares rate limit) while keeping downloads fire-and-forget and URLs fresh.
    let cursor = 0;
    const worker = async (): Promise<void> => {
      while (cursor < items.length && !signal.aborted) {
        const item = items[cursor++];
        const p = await prepAsync(item);
        if (p.kind === 'link') { summary.links++; settledOne(); continue; }
        fireAndForget(item, p.spec);
      }
    };
    await Promise.all(Array.from({ length: RESOLVE_CONCURRENCY }, () => worker()));
  } else {
    // FIRE-AND-FORGET (Chrome, no resolve): queue every download without awaiting;
    // Chrome paces them via its per-host cap. The settlement loop below is the
    // only wait — no barrier that could sit idle while Chrome's queue drains.
    for (let i = 0; i < items.length; i++) {
      if (signal.aborted) break;
      const spec = prep(items[i]);
      if (!spec) { summary.links++; settledOne(); continue; }
      fireAndForget(items[i], spec);
    }
  }

  await awaitSettlement({ deps, signal, startedAfter, expected: () => expected, settled: () => settledCount, pendingIds, setNotify: fn => { notifySettle = fn; } });

  if (signal.aborted && pendingIds.size) {
    // STOP: actually stop the transfers (they would otherwise keep running in
    // the browser). Each cancel settles as a reason-less interrupt, which the
    // verifier files under FAILURES.txt's cancelled section. A short bounded
    // drain lets those interrupts land so the summary's cancelled bucket is
    // populated before it escapes; STOP stays responsive either way.
    deps.log?.(`attachment stop: cancelling ${pendingIds.size} in-flight download(s)`);
    for (const id of pendingIds) {
      try { void Promise.resolve(deps.downloads.cancel(id)).catch(() => { /* already terminal */ }); } catch { /* noop */ }
    }
    const drainDeadline = Date.now() + 3_000;
    while (pendingIds.size && Date.now() < drainDeadline) {
      await new Promise<void>(res => { notifySettle = res; setTimeout(res, 200); });
      notifySettle = undefined;
    }
  }
  if (signal.aborted) {
    // Candidates whose outcome the run no longer knows — never dispatched
    // (the loops break on abort), not yet enrolled, or a cancel that didn't
    // land within the drain — were all stopped by the user: tally them as
    // cancelled so the buckets account for every candidate. Callbacks are
    // stripped below, so a late-landing settle can't double-count them.
    const unaccounted = expected - settledCount;
    if (unaccounted > 0) { summary.cancelled += unaccounted; settledCount += unaccounted; }
  }

  // End of run: freeze the escaped summary and drop the progress callbacks of
  // anything still unsettled (stall-break leftovers, undrained cancels). The
  // top-level verifier still does their FAILURES.txt bookkeeping from the
  // remaining meta; only this run's tally stops listening.
  runEnded = true;
  for (const id of pendingIds) {
    const meta = pendingVerify.get(id);
    if (meta?.onSettled) pendingVerify.set(id, { ...meta, onSettled: undefined });
  }

  // Flush FAILURES.txt once, now that retries are done and the failure lists
  // are final. Late verifier events (after this) write normally again.
  failuresWriteSuppressed = false;
  await flushFailures(deps, exportFolder);

  deps.log?.(`attachment downloads settled: ${summary.saved} saved, ${summary.links} link(s), ${summary.failed} failed, ${summary.cancelled} cancelled of ${summary.total}` + (viaAspx ? ` (${viaAspx} via download.aspx)` : ''));
  return summary;
}

// Resolve concurrency. Small on purpose: the shares API is rate-limited and a
// burst of hundreds of resolves risks throttling (which would 429 the whole
// set). Resolve is a fast metadata call, so a low cap still drains quickly.
const RESOLVE_CONCURRENCY = 4;

// Settlement wait tuning. The poll is a backstop for missed onChanged events
// and the MV3 keep-alive; the verifier's onSettled callbacks (via notify) do
// the real-time accounting. The stall window only trips when NOTHING moved —
// no settlement and no byte progress across every pending transfer — for its
// whole duration; then we stop blocking completion (transfers keep running in
// the browser and the top-level verifier still handles them as they finish).
const SETTLE_POLL_MS = 2_500;
const SETTLE_STALL_MS = 180_000;

async function awaitSettlement(run: {
  deps: AttachmentDownloadDeps;
  signal: AbortSignal;
  // Captured by the caller BEFORE dispatch, so the batch covers downloads
  // started at any point of the run (Firefox's bounded pool dispatches over
  // minutes; a loop-start capture would miss its early downloads).
  startedAfter: string;
  expected: () => number;
  settled: () => number;
  pendingIds: Set<number>;
  setNotify: (fn: (() => void) | undefined) => void;
}): Promise<void> {
  const { deps, signal, startedAfter, pendingIds } = run;
  const bytesSeen = new Map<number, number>();
  let lastActivityAt = Date.now();
  let lastSettled = run.settled();

  while (run.settled() < run.expected() && !signal.aborted) {
    await new Promise<void>(res => {
      run.setNotify(res);
      setTimeout(res, SETTLE_POLL_MS);
    });
    run.setNotify(undefined);
    if (run.settled() >= run.expected() || signal.aborted) break;

    // Batched poll: ONE search per tick (per-id polling does not scale to
    // hundreds of files) filtered against this run's pending ids. A failed
    // search falls through with an empty snapshot — the stall clock below
    // still runs, so a permanently broken search can't loop forever.
    let found: chrome.downloads.DownloadItem[] = [];
    try {
      const r = await Promise.resolve(deps.downloads.search({ startedAfter, limit: 0 }));
      found = Array.isArray(r) ? r : [];
    } catch { /* fall through with found = [] */ }

    let progress = false;
    const seen = new Set<number>();
    for (const it of found) {
      if (!pendingIds.has(it.id)) continue;
      seen.add(it.id);
      if (it.state === 'complete' || it.state === 'interrupted') {
        // Missed onChanged event — feed the verifier a synthetic delta.
        // Idempotent: verifyDownloadOnChanged deletes from pendingVerify
        // before any await, so a racing real event can't double-process; it
        // also re-checks the item's real state, so a stale snapshot can't
        // misclassify.
        void verifyDownloadOnChanged(deps, { id: it.id, state: { current: it.state } });
      } else if (it.paused) {
        progress = true; // parked on the user, not stalled
      } else {
        const prev = bytesSeen.get(it.id) ?? -1;
        const cur = it.bytesReceived || 0;
        if (cur > prev) { bytesSeen.set(it.id, cur); progress = true; }
      }
    }
    // Ids absent from the batch snapshot: usually an enrollment that raced
    // the search, sometimes an erased download. CONFIRM per id before doing
    // anything — feeding a blind synthetic interrupt here would false-cancel
    // a live download.
    for (const id of [...pendingIds]) {
      if (seen.has(id)) continue;
      let item: chrome.downloads.DownloadItem | undefined;
      let searched = false;
      try {
        const r = await Promise.resolve(deps.downloads.search({ id }));
        item = Array.isArray(r) ? r[0] : undefined;
        searched = true;
      } catch { /* leave it for the next tick */ }
      if (!searched) continue;
      if (!item || item.state === 'complete' || item.state === 'interrupted') {
        // Terminal, or erased from history (the verifier's missing-item path
        // settles erased ids as failures).
        void verifyDownloadOnChanged(deps, { id, state: { current: item?.state ?? 'interrupted' } });
      } else if (item.paused) {
        progress = true;
      } else {
        const prev = bytesSeen.get(id) ?? -1;
        const cur = item.bytesReceived || 0;
        if (cur > prev) { bytesSeen.set(id, cur); progress = true; }
      }
    }

    if (run.settled() > lastSettled) { lastSettled = run.settled(); progress = true; }
    if (progress) {
      lastActivityAt = Date.now();
    } else if (Date.now() - lastActivityAt > SETTLE_STALL_MS) {
      deps.log?.(`attachment settlement stalled: ${run.expected() - run.settled()} of ${run.expected()} still pending; not blocking completion`);
      break;
    }
  }
  run.setNotify(undefined);
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
  opts: { maxConcurrent?: number } = {},
): Promise<AttachmentDownloadSummary> {
  const items = collectDocumentAttachments(messages).map(c => ({ ...c, chatFolder: '' }));
  return downloadAttachments(items, exportFolder, deps, signal, opts);
}

/**
 * Poll a download until it leaves 'in_progress' (complete/interrupted), so a
 * bounded worker can pace its next dispatch. Best-effort: returns on abort, on a
 * missing item, or after a long cap. A user cancel or server error ends the wait.
 */
async function awaitTerminal(deps: { downloads: DownloadsApi }, id: number, signal: AbortSignal): Promise<void> {
  for (let n = 0; n < 2400 && !signal.aborted; n++) {
    let item: chrome.downloads.DownloadItem | undefined;
    try { const r = await deps.downloads.search({ id }); item = Array.isArray(r) ? r[0] : undefined; }
    catch { return; }
    if (!item || item.state !== 'in_progress') return;
    await new Promise<void>(res => setTimeout(res, 500));
  }
}

/**
 * Counts-only projection for the status wire. Drops the `failures` array (which
 * carries attachment file names) so it is never broadcast over runtime
 * messaging or persisted into the active-export snapshot in chrome.storage.
 */
export function toFilesSummaryWire(s: AttachmentDownloadSummary): { total: number; saved: number; links: number; failed: number; cancelled: number } {
  return { total: s.total, saved: s.saved, links: s.links, failed: s.failed, cancelled: s.cancelled };
}
