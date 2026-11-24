/**
 * What to do with an API-mode scrape result.
 *
 * - 'messages'      : some messages survived the filters; export them.
 * - 'empty'         : the API returned no messages at all (most often a legacy
 *                     Skype-imported 1:1 with no migrated history); export empty.
 * - 'none-in-range' : the API had messages but none in the requested date
 *                     range, and there is no DOM fallback (multi-chat bundle).
 *                     Nothing new, not a failure; export empty and flag it.
 * - 'dom-fallback'  : the API had messages but the filters dropped them all
 *                     (a system-only chat, say); single-chat mode retries by
 *                     DOM scroll.
 */
export type ApiScrapeOutcome = 'messages' | 'empty' | 'none-in-range' | 'dom-fallback';

export function apiScrapeOutcome(converted: number, raw: number, noDomFallback: boolean): ApiScrapeOutcome {
  if (converted > 0) return 'messages';
  if (raw === 0) return 'empty';
  return noDomFallback ? 'none-in-range' : 'dom-fallback';
}
