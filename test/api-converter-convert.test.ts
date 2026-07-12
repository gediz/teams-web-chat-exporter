// Behavior lock for convertApiMessages (src/content/api-converter.ts) — the
// conversion spine that turns the raw Teams API payload into the ExportMessage[]
// every format renders. Fixtures are SYNTHETIC (no real OIDs/thread-ids/names/
// message text) per repo policy. Covers the orchestration convertApiMessages owns:
// per-message conversion, the newest-first -> chronological re-sort, and the
// duplicate-forward-row dedup.
import { test, expect } from 'vitest';
import { convertApiMessages } from '../src/content/api-converter.ts';

const opts = {} as never;
const richText = (o: Record<string, unknown>) => ({ messagetype: 'RichText/Html', ...o } as never);

test('converts a basic rich-text message to author + text', () => {
  const out = convertApiMessages([
    richText({ content: '<div>hello world</div>', imdisplayname: 'Alice', composetime: '2024-01-01T10:00:00Z' }),
  ], opts);
  expect(out).toHaveLength(1);
  expect(out[0].author).toBe('Alice');
  expect(out[0].text).toContain('hello world');
});

test('orders by timestamp even when input array order disagrees (reverse+stable-sort, not just reverse)', () => {
  // Deliberately scrambled: the array order is NOT reverse-chronological, so a plain
  // reverse() would mis-order these (yield first, third, second). Only the stable
  // timestamp sort the converter applies after reverse() produces the right order —
  // this is the re-ranked-tombstone case the sort exists for.
  const out = convertApiMessages([
    richText({ content: '<div>second</div>', imdisplayname: 'A', composetime: '2024-01-01T10:05:00Z' }),
    richText({ content: '<div>third</div>', imdisplayname: 'A', composetime: '2024-01-01T10:10:00Z' }),
    richText({ content: '<div>first</div>', imdisplayname: 'A', composetime: '2024-01-01T10:00:00Z' }),
  ], opts);
  expect(out.map((m) => m.text.trim())).toEqual(['first', 'second', 'third']);
});

test('collapses the duplicate row Teams emits for one forward (same clientmessageid)', () => {
  const fwd = (cid: string) => richText({
    content: '<div>forwarded body</div>',
    imdisplayname: 'Bob',
    composetime: '2024-01-01T10:00:00Z',
    clientmessageid: cid,
    properties: { forwardTemplateId: 'ft1' },
  });
  // Same clientmessageid = Teams' duplicate row for ONE forward -> collapse to one.
  expect(convertApiMessages([fwd('cm1'), fwd('cm1')], opts)).toHaveLength(1);
  // Distinct clientmessageids = two genuine forwards -> both kept.
  expect(convertApiMessages([fwd('cm1'), fwd('cm2')], opts)).toHaveLength(2);
});
