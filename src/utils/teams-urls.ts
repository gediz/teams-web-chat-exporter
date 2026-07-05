/**
 * Single source of truth for all Teams URL patterns.
 *
 * Used by:
 *  - wxt.config.ts          (host_permissions)
 *  - content.ts             (content-script matches)
 *  - background.ts          (runtime URL check)
 *  - popup/App.svelte       (runtime URL check)
 *
 * Two distinct concerns:
 *  - {@link TEAMS_MATCH_PATTERNS}: web-app origins where the content script RUNS.
 *  - {@link API_FETCH_PATTERNS}:  upstream service origins the extension FETCHES
 *    from (Graph for users/photos, AMS for inline images). Required as
 *    host_permissions in Firefox so content-script fetches aren't blocked by
 *    the extension-origin CORS policy; harmless in Chrome.
 *
 * To support a new domain or proxy suffix, add it to the arrays below.
 */

/** Teams web-app domains. `hasSubdomains` adds a `*.` wildcard prefix to match patterns. */
const TEAMS_DOMAINS = [
  { domain: 'teams.microsoft.com', hasSubdomains: true },
  { domain: 'teams.microsoft.us', hasSubdomains: true },
  { domain: 'cloud.microsoft', hasSubdomains: true },
  { domain: 'teams.live.com', hasSubdomains: false },
] as const;

/** Proxy suffixes appended to Teams domains (e.g. Microsoft Defender for Cloud Apps). */
const PROXY_SUFFIXES = ['', '.mcas.ms'] as const;

/**
 * Match-patterns for manifest `host_permissions` and content-script `matches`.
 * Derived from {@link TEAMS_DOMAINS} × {@link PROXY_SUFFIXES}.
 */
export const TEAMS_MATCH_PATTERNS: string[] = TEAMS_DOMAINS.flatMap(
  ({ domain, hasSubdomains }) =>
    PROXY_SUFFIXES.map(suffix => {
      const host = hasSubdomains ? `*.${domain}${suffix}` : `${domain}${suffix}`;
      return `https://${host}/*`;
    }),
);

/**
 * Upstream service origins the extension fetches from (NOT where it runs).
 *  - graph.microsoft.com / .us : Graph API for user resolution and profile photos.
 *  - *.asm.skype.com           : Skype AMS image CDN (inline images, link previews).
 *  - SharePoint hosts          : Needed to fetch the binary of image files
 *                                attached via paperclip / drag-drop. Teams
 *                                stores these on the user's SharePoint /
 *                                OneDrive, not on AMS. Without this the export
 *                                only captures the file metadata, leaving the
 *                                image missing from the exported HTML / PDF /
 *                                zip. Graph's /shares endpoint can't substitute
 *                                because every Graph file-content endpoint
 *                                redirects to a SharePoint host for the
 *                                actual bytes.
 *
 *    Microsoft cloud → SharePoint host:
 *      Worldwide / commercial / GCC : <tenant>.sharepoint.com
 *      US GCC High                  : <tenant>.sharepoint.us
 *      US DoD                       : <tenant>.sharepoint-mil.us
 *      Office 365 China (21Vianet)  : <tenant>.sharepoint.cn
 *      Microsoft Cloud Deutschland  : retired in 2021, omitted
 *
 *  Not listed: my.microsoftpersonalcontent.com (consumer SharePoint /
 *  OneDrive Personal Content). Teams Free stores paperclip uploads
 *  there for personal accounts, but the host returns a 302 to
 *  login.live.com when accessed cross-origin — Microsoft's auth flow
 *  requires interactive sign-in. Even with host_permission and a
 *  background-context fetch, the redirect-follow fails because
 *  login.live.com has no CORS headers for any non-interactive caller.
 *  Confirmed via probe + console capture in 2026-04. The user has to
 *  click the link in the rendered HTML to open the file in OneDrive
 *  manually.
 */
export const API_FETCH_PATTERNS: string[] = [
  'https://graph.microsoft.com/*',
  'https://graph.microsoft.us/*',
  'https://*.asm.skype.com/*',
  'https://*.sharepoint.com/*',
  'https://*.sharepoint.us/*',
  'https://*.sharepoint-mil.us/*',
  'https://*.sharepoint.cn/*',
  // Teams Free messaging service. Discovered chat service URL is
  // `https://msgapi.teams.live.com` — a subdomain of teams.live.com,
  // not covered by the bare-domain content-script match. Listed here
  // (not in TEAMS_MATCH_PATTERNS) so the content script doesn't get
  // injected on these non-UI hosts.
  'https://*.teams.live.com/*',
];

/** Regex that matches any Teams web-app URL (including proxy suffixes). */
const escapeDot = (s: string) => s.replace(/\./g, '\\.');
const domainAlts = TEAMS_DOMAINS.map(({ domain }) => escapeDot(domain)).join('|');
const suffixAlts = PROXY_SUFFIXES.filter(Boolean).map(escapeDot).join('|');

const TEAMS_URL_REGEX = new RegExp(
  `^https://(.*\\.)?(${domainAlts})(${suffixAlts})?/`,
  'i',
);

/** Returns `true` when the given URL belongs to a Teams web-app host. */
export function isTeamsUrl(url?: string | null): boolean {
  return TEAMS_URL_REGEX.test(url || '');
}

/**
 * Microsoft-owned host suffixes the extension is allowed to attach the user's
 * live messaging/Skype token to. Deliberately broad: a too-narrow list would
 * risk false negatives on regional chat-service / SharePoint hosts we cannot
 * all enumerate, while any non-Microsoft host (e.g. an exfiltration target
 * injected via a poisoned `backwardLink` or `chatServiceUrl`) is still
 * rejected. The leading-dot check (`.endsWith('.' + s)`) prevents a
 * look-alike like `evilmicrosoft.com` from matching `microsoft.com`.
 */
const MS_API_HOST_SUFFIXES = [
  'microsoft.com',   // teams / graph / *.ng.msg.teams.microsoft.com, etc.
  'microsoft.us',    // GCC High / DoD
  'cloud.microsoft',
  'teams.live.com',  // Teams Free
  'skype.com',       // *.asm.skype.com (AMS image CDN)
  'office.com',      // *.ic3.teams.office.com
  'office365.com', 'office365.us',  // GCC / GCC-High ic3 + chat-service hosts (matches attachments.ts allow-list)
  'lync.com',        // legacy infra hosts
  'sharepoint.com', 'sharepoint.us', 'sharepoint-mil.us', 'sharepoint.cn',
] as const;

/**
 * Returns `true` when `url` is https and its host is a Microsoft-owned API
 * host the extension may send credentialed (token-bearing) requests to.
 * Used by the api-client messaging guard so the live token is never attached
 * to a non-Microsoft origin.
 */
export function isMicrosoftApiHost(url?: string | null): boolean {
  if (!url) return false;
  let host: string;
  try {
    const u = new URL(url);
    if (u.protocol !== 'https:') return false;
    host = u.hostname.toLowerCase().replace(/\.$/, '');  // tolerate absolute-DNS trailing dot
  } catch {
    return false;
  }
  return MS_API_HOST_SUFFIXES.some(s => host === s || host.endsWith('.' + s));
}

/**
 * SharePoint / OneDrive-for-Business file host suffixes, across all national
 * clouds. Deliberately NARROWER than {@link MS_API_HOST_SUFFIXES}: this gates
 * which attachment URLs the document downloader will run the credentialed
 * `/_api/v2.0/shares` resolver against, so it must be the SharePoint family
 * only — never teams.* / graph.* / skype.*. Consumer
 * `my.microsoftpersonalcontent.com` is intentionally absent (it 302s to
 * login.live.com with no CORS; see API_FETCH_PATTERNS), so those stay
 * link-only.
 */
const SHAREPOINT_FILE_HOST_SUFFIXES = [
  'sharepoint.com',
  'sharepoint.us',
  'sharepoint-mil.us',
  'sharepoint.cn',
] as const;

/**
 * Returns `true` when `url` is an https SharePoint/OneDrive-for-Business file
 * host (any national cloud). Used to gate the document-download shares resolver
 * before a credentialed (cookie-bearing) fetch is made. The leading-dot check
 * (`.endsWith('.' + s)`) prevents a look-alike like `evilsharepoint.com` from
 * matching `sharepoint.com`.
 */
export function isSharePointFileHost(url?: string | null): boolean {
  if (!url) return false;
  let host: string;
  try {
    const u = new URL(url);
    if (u.protocol !== 'https:') return false;
    host = u.hostname.toLowerCase().replace(/\.$/, '');
  } catch {
    return false;
  }
  return SHAREPOINT_FILE_HOST_SUFFIXES.some(s => host === s || host.endsWith('.' + s));
}

/**
 * Site-collection base for a SharePoint file URL: `https://<host>/personal/<u>`,
 * `/sites/<s>`, or `/teams/<t>` (falls back to the host root). Site-collection
 * scoped endpoints (the drive-item API, the download.aspx handler) must be
 * addressed on the file's own site, not the host root.
 */
export function siteCollectionBase(u: URL): string {
  const m = u.pathname.match(/^(\/(?:personal|sites|teams)\/[^/]+)/i);
  return `https://${u.host}${m ? m[1] : ''}`;
}
