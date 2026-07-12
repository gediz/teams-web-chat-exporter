// Behavior lock for pure helpers in src/content/api-converter.ts.
// Covers three correctness fixes landed alongside these tests:
//   - CB1: a video/audio duration over 1h keeps its hours field (was dropped)
//   - CB2: an unresolved reactor shows "(unknown user)", not a raw id slice
//   - CB9: preview-card dedup only merges a real prefix-up-to-separator, not any
//          mid-string substring (which used to delete distinct cards)
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { formatDuration, decorateReactions, previewTitlesDuplicate } from '../src/content/api-converter.ts';

test('CB1: formatDuration keeps hours for durations over 1h', () => {
  assert.equal(formatDuration('PT1H5M23S'), '1:05:23');
  assert.equal(formatDuration('PT2H'), '2:00:00');
  assert.equal(formatDuration('PT2H15M48S'), '2:15:48');
  // Sub-hour shape is unchanged (m:ss, no leading hours).
  assert.equal(formatDuration('PT12S'), '0:12');
  assert.equal(formatDuration('PT1M30S'), '1:30');
});

test('CB2: an unresolved reactor is "(unknown user)", never a raw id slice', () => {
  const map = new Map([['8:orgid:known-guid', 'Alice']]);
  const raw = [{ emoji: '👍', count: 2, rawReactors: [
    { mri: '8:orgid:known-guid' },
    { mri: '8:orgid:0123456789abcdef' }, // not in the resolution map
  ] }];
  const [reaction] = decorateReactions(raw, map, null, null);
  const names = reaction.reactors.map((r) => r.name);
  assert.deepEqual(names, ['Alice', '(unknown user)']);
  assert.ok(!names.some((n) => /^[0-9a-f]{8}$/i.test(n)), 'no 8-hex id slice as a name');
});

test('CB9: preview dedup merges a prefix-to-separator, not any substring', () => {
  // The documented duplicate: a card title is a prefix of the full preview title.
  assert.equal(previewTitlesDuplicate('title', 'title - author - site'), true);
  assert.equal(previewTitlesDuplicate('title - author - site', 'title'), true, 'order-independent');
  assert.equal(previewTitlesDuplicate('report', 'report'), true, 'identical');
  // The bug: a distinct title that merely appears mid-string must NOT be merged.
  assert.equal(previewTitlesDuplicate('report', 'annual report 2024'), false);
  // A plain trailing-space continuation is a distinct card, not a duplicate.
  assert.equal(previewTitlesDuplicate('report', 'report 2024'), false);
});
