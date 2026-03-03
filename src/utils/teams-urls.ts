/**
 * Single source of truth for all Teams URL patterns.
 *
 * Used by:
 *  - wxt.config.ts          (host_permissions)
 *  - content.ts             (content-script matches)
 *  - background.ts          (runtime URL check)
 *  - popup/App.svelte       (runtime URL check)
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
