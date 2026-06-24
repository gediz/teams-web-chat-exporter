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
// (html/svg/xml/...) are routed through the _layouts/download.aspx handler so
// SharePoint's active-content sandbox interstitial isn't saved in their place.
// See toDownloadUrl.
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
type DownloadsApi = Pick<typeof chrome.downloads, 'download'>;

export interface AttachmentDownloadDeps {
  downloads: DownloadsApi;
  // Called after each candidate is processed. `done` counts processed
  // candidates; `total` is the candidate count.
  onProgress?: (done: number, total: number) => void;
  log?: (msg: string) => void;
}

export interface AttachmentDownloadSummary {
  total: number;   // SharePoint document candidates found (deduped by href)
  saved: number;   // handed to chrome.downloads (fire-and-forget; the browser
                   // streams server->disk in its own context)
  links: number;   // skipped without a download attempt (host gate) — kept as a link
  failed: number;  // chrome.downloads.download threw synchronously
  failures: Array<{ name: string; reason: string }>;
}

export interface AttachmentCandidate {
  href: string;        // the file's absolute SharePoint URL (download source)
  name: string;        // best-known display name (from Attachment.label)
  chatFolder: string;  // per-chat subfolder under the export root ('' for single-chat)
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

/**
 * Map a file's absolute SharePoint URL to the URL chrome.downloads should fetch.
 *
 * Most types serve their raw bytes directly to an authenticated caller, so we
 * append `?download=1` (SharePoint's force-download hint) and hand them over as
 * is — the proven path that pulls PDFs/zips/Office/video correctly.
 *
 * Renderable-markup types ({@link SANDBOXED_EXT}) instead hit the active-content
 * sandbox interstitial. Route those through the `_layouts/download.aspx` handler
 * (addressed by SourceUrl), which forces a server-side attachment download (no
 * rendering, no sandbox) so the real file streams with just the FedAuth cookie.
 * Root-level `_layouts` plus an absolute SourceUrl redirects to the file's own
 * web, so no site-path parsing is needed.
 */
function toDownloadUrl(href: string): string {
  let u: URL;
  try { u = new URL(href); } catch { return href; }
  if (SANDBOXED_EXT.test(u.pathname)) {
    return `https://${u.host}/_layouts/15/download.aspx?SourceUrl=${encodeURIComponent(href)}`;
  }
  return href + (href.includes('?') ? '&' : '?') + 'download=1';
}

/** Collect deduped SharePoint document attachments across all messages. */
export function collectDocumentAttachments(messages: ExportMessage[]): Array<{ href: string; name: string }> {
  const seen = new Set<string>();
  const out: Array<{ href: string; name: string }> = [];
  for (const m of messages) {
    const atts = m.attachments;
    if (!atts) continue;
    for (const a of atts) {
      const href = a.href;
      if (!href || seen.has(href)) continue;
      // Skip inline-image types (handled by downloadImages) and anything not on
      // a SharePoint file host (the only family we download from).
      if (IMAGE_FILETYPE.test((a.type || '').toString())) continue;
      if (!isSharePointFileHost(href)) continue;
      seen.add(href);
      out.push({ href, name: (a.label || '').toString() });
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
 * Trade-off: because we no longer await each download, we can't inspect its
 * response, so a file the user cannot open still saves the SharePoint "request
 * access" HTML page (it surfaces in the browser's own download manager). Deleting
 * those required awaiting completion, which is exactly what broke reliability.
 */
export async function downloadAttachments(
  items: AttachmentCandidate[],
  exportFolder: string,
  deps: AttachmentDownloadDeps,
  signal: AbortSignal,
): Promise<AttachmentDownloadSummary> {
  const summary: AttachmentDownloadSummary = { total: items.length, saved: 0, links: 0, failed: 0, failures: [] };
  if (!items.length) return summary;

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
      const rawName = item.name || fileNameFromUrl(item.href) || 'attachment';
      const url = toDownloadUrl(item.href);
      const name = item.name || fileNameFromUrl(item.href) || item.href;
      try {
        // Fire, do NOT await (see the function note). The download is registered
        // the instant this is called; the returned id-promise is ignored, but we
        // attach a catch so a rejection isn't an unhandled promise.
        Promise.resolve(deps.downloads.download({
          url,
          filename: buildAttachmentPath(exportFolder, item.chatFolder, rawName),
          saveAs: false,
          conflictAction: 'uniquify',
        })).catch(e => deps.log?.(`attachment download error: ${name}: ${errMsg(e)}`));
        summary.saved++;
        if (url.includes('/_layouts/15/download.aspx')) viaAspx++;
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
