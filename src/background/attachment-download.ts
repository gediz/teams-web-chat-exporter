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

interface VerifyMeta { name: string; href: string; chatFolder: string; exportFolder: string; markup?: boolean; }
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

/** Clear verifier state at the start of an export's Files phase. */
function resetVerifyState(): void {
  pendingVerify.clear();
  verifiedFailures.clear();
  verifiedCancels.clear();
  failuresDownloadId.clear();
  for (const t of failuresWriteTimer.values()) clearTimeout(t);
  failuresWriteTimer.clear();
}

/**
 * (Re)schedule the debounced FAILURES.txt write for a chat folder. The in-memory
 * lists update immediately at the call site; this coalesces a burst of records
 * into a single download once it goes quiet (see FAILURES_WRITE_DEBOUNCE_MS).
 */
function scheduleFailuresWrite(deps: { downloads: DownloadsApi }, meta: VerifyMeta): void {
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

  let item: chrome.downloads.DownloadItem | undefined;
  try {
    const r = await deps.downloads.search({ id: delta.id });
    item = Array.isArray(r) ? r[0] : undefined;
  } catch { return; }
  if (!item) return;

  if (st === 'interrupted') {
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
      return;
    }
    deps.log?.(`attachment no-access: ${meta.name} (${item.error})`);
    recordFailure(deps, meta);
    return;
  }

  // Completed. A NON-markup file that arrived as text/html is a request-access
  // page saved under the file's name (a real pdf/zip/video is never text/html).
  // Markup is legitimately text/html, so a completed markup download is kept.
  if (meta.markup || (item.mime || '').toLowerCase() !== 'text/html') return;
  try { await deps.downloads.removeFile(delta.id); } catch { /* may already be gone */ }
  try { await deps.downloads.erase({ id: delta.id }); } catch { /* noop */ }
  deps.log?.(`attachment no-access: removed request-access page for ${meta.name}`);
  recordFailure(deps, meta);
}

/**
 * Hand a set of SharePoint document attachments to chrome.downloads. Returns a
 * tally. No-op safe: an empty `items` returns a zero summary.
 *
 * Two dispatch modes:
 *   - Chrome (default): fire-and-forget. download() does not resolve until the
 *     transfer starts (~1s each), so awaiting it serially outran the MV3 worker's
 *     lifetime and evicted it mid-phase. We queue every file in milliseconds and
 *     let Chrome's per-host concurrency cap pace them; the id is observed only to
 *     enroll the download for post-completion verification.
 *   - Firefox (opts.maxConcurrent): bounded pool. Firefox does NOT cap concurrent
 *     downloads, so firing all at once made SharePoint throttle the burst and
 *     return an error page for every file. Each worker awaits its download's
 *     terminal state before pulling the next, capping in-flight transfers.
 *     Firefox's persistent background page makes the long wait safe.
 */
export async function downloadAttachments(
  items: AttachmentCandidate[],
  exportFolder: string,
  deps: AttachmentDownloadDeps,
  signal: AbortSignal,
  opts: { maxConcurrent?: number } = {},
): Promise<AttachmentDownloadSummary> {
  const summary: AttachmentDownloadSummary = { total: items.length, saved: 0, links: 0, failed: 0, failures: [] };
  if (!items.length) return summary;

  resetVerifyState();
  deps.log?.(`attachment phase: ${items.length} candidate(s)` + (opts.maxConcurrent ? `, max ${opts.maxConcurrent} concurrent` : ''));

  let viaAspx = 0;  // queued files routed through the download.aspx handler
  const report = () => {
    const processed = summary.saved + summary.links + summary.failed;
    deps.onProgress?.(processed, items.length);
    if (processed % 50 === 0 || processed === items.length) deps.log?.(`attachment dispatch: ${processed}/${items.length}`);
  };

  // Per-item prep shared by both dispatch paths. Returns the download spec, or
  // null for a host-gated link (the caller counts it as a link). `markup` is the
  // file TYPE (so the verifier knows text/html is the real file), independent of
  // the download URL form (markup hits download.aspx, everything else ?download=1).
  const prep = (item: AttachmentCandidate) => {
    if (!isSharePointFileHost(item.href)) return null;
    const name = item.name || fileNameFromUrl(item.href) || item.href;
    const rawName = item.name || fileNameFromUrl(item.href) || 'attachment';
    const url = toDownloadUrl(item.href, item.itemid);
    const viaDownloadAspx = isDownloadAspxUrl(url);
    const markup = isMarkupFile(item.href);
    return { name, url, viaDownloadAspx, markup, filename: buildAttachmentPath(exportFolder, item.chatFolder, rawName) };
  };
  const enroll = (id: number, item: AttachmentCandidate, name: string, markup: boolean) => {
    pendingVerify.set(id, { name, href: item.href, chatFolder: item.chatFolder, exportFolder, markup });
  };

  if (opts.maxConcurrent && opts.maxConcurrent > 0) {
    // BOUNDED CONCURRENCY (Firefox): cap in-flight transfers so we never burst
    // SharePoint. A worker awaits its download's terminal state before the next.
    let cursor = 0;
    const worker = async (): Promise<void> => {
      while (cursor < items.length && !signal.aborted) {
        const item = items[cursor++];
        const spec = prep(item);
        if (!spec) { summary.links++; report(); continue; }
        let id: number | undefined;
        try {
          id = await Promise.resolve(deps.downloads.download({ url: spec.url, filename: spec.filename, saveAs: false, conflictAction: 'uniquify' }));
        } catch (e) { summary.failed++; summary.failures.push({ name: spec.name, reason: errMsg(e) }); report(); continue; }
        if (typeof id === 'number') { enroll(id, item, spec.name, spec.markup); summary.saved++; if (spec.viaDownloadAspx) viaAspx++; }
        report();
        if (typeof id === 'number') await awaitTerminal(deps, id, signal);
      }
    };
    await Promise.all(Array.from({ length: opts.maxConcurrent }, () => worker()));
  } else {
    // FIRE-AND-FORGET (Chrome): queue all downloads without awaiting; Chrome paces
    // them via its per-host concurrency cap. The id is observed only to enroll.
    for (let i = 0; i < items.length; i++) {
      if (signal.aborted) break;
      const spec = prep(items[i]);
      if (!spec) { summary.links++; report(); continue; }
      const item = items[i];
      try {
        Promise.resolve(deps.downloads.download({ url: spec.url, filename: spec.filename, saveAs: false, conflictAction: 'uniquify' }))
          .then(id => { if (typeof id === 'number') enroll(id, item, spec.name, spec.markup); })
          .catch(e => deps.log?.(`attachment download error: ${spec.name}: ${errMsg(e)}`));
        summary.saved++;
        if (spec.viaDownloadAspx) viaAspx++;
      } catch (e) { summary.failed++; summary.failures.push({ name: spec.name, reason: errMsg(e) }); }
      report();
    }
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
export function toFilesSummaryWire(s: AttachmentDownloadSummary): { total: number; saved: number; links: number; failed: number } {
  return { total: s.total, saved: s.saved, links: s.links, failed: s.failed };
}
