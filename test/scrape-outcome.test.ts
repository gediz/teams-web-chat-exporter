// Behavior lock for apiScrapeOutcome (src/utils/scrape-outcome.ts): how an
// API-mode scrape result is handled, including a dated multi-chat export where
// a chat has nothing new in range (it used to be reported as a failure).
import { test, expect } from 'vitest';
import { apiScrapeOutcome } from '../src/utils/scrape-outcome.ts';

test('messages that survive the filters are exported', () => {
  expect(apiScrapeOutcome(3, 10, true)).toBe('messages');
  expect(apiScrapeOutcome(3, 10, false)).toBe('messages');
});

test('an API result with no messages at all is an empty export', () => {
  expect(apiScrapeOutcome(0, 0, true)).toBe('empty');
  expect(apiScrapeOutcome(0, 0, false)).toBe('empty');
});

test('a bundle chat with messages but none in range is nothing new, not a failure', () => {
  expect(apiScrapeOutcome(0, 25, true)).toBe('none-in-range');
});

test('a single-chat export still retries by DOM scroll when the filters drop everything', () => {
  expect(apiScrapeOutcome(0, 25, false)).toBe('dom-fallback');
});
