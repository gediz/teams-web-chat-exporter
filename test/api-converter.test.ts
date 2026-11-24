// Behavior lock for pure helpers in src/content/api-converter.ts.
// Covers three correctness fixes landed alongside these tests:
//   - CB1: a video/audio duration over 1h keeps its hours field (was dropped)
//   - CB2: an unresolved reactor shows "(unknown user)", not a raw id slice
//   - CB9: preview-card dedup only merges a real prefix-up-to-separator, not any
//          mid-string substring (which used to delete distinct cards)
// Also covers the app-card text walker, <Swift> card decoding and onbehalfof
// parsing.
import { test } from 'vitest';
import assert from 'node:assert/strict';
import { formatDuration, decorateReactions, previewTitlesDuplicate, collectAdaptiveText, decodeSwiftCards, parseOnBehalfOf, cardsToText } from '../src/content/api-converter.ts';

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

test('collectAdaptiveText reads nested card elements in document order', () => {
  const body = [
    { type: 'ColumnSet', columns: [
      { type: 'Column', items: [{ type: 'TextBlock', text: 'Daily check-in' }] },
      { type: 'Column', items: [{ type: 'Image', url: 'https://example.invalid/icon.png' }] },
    ] },
    { type: 'Container', items: [
      { type: 'TextBlock', text: 'What did you work on?' },
      { type: 'Container', items: [{ type: 'TextBlock', text: 'Fixed the login page' }] },
    ] },
    { type: 'FactSet', facts: [{ title: 'Status', value: 'Done' }, { title: '', value: 'Untitled fact' }] },
    { type: 'RichTextBlock', inlines: [{ type: 'TextRun', text: 'Rich ' }, 'tail'] },
    { type: 'Table', rows: [
      { cells: [{ items: [{ type: 'TextBlock', text: 'Cell A' }] }, { items: [{ type: 'TextBlock', text: 'Cell B' }] }] },
    ] },
    { type: 'TextBlock', text: 'Due {{DATE(2024-01-02T03:04:05Z, SHORT)}}' },
  ];
  assert.deepEqual(collectAdaptiveText(body), [
    'Daily check-in',
    'What did you work on?',
    'Fixed the login page',
    'Status: Done',
    'Untitled fact',
    'Rich tail',
    'Cell A | Cell B',
    'Due 2024-01-02',
  ]);
});

test('collectAdaptiveText skips buttons and show/hide toggles but keeps hidden text', () => {
  const body = [
    { type: 'Container', isVisible: false, items: [{ type: 'TextBlock', text: 'Full answer' }] },
    { type: 'Container', selectAction: { type: 'Action.ToggleVisibility' }, items: [{ type: 'TextBlock', text: 'See more' }] },
    { type: 'TextBlock', text: 'See less', selectAction: { type: 'Action.ToggleVisibility' } },
    { type: 'ActionSet', actions: [{ type: 'Action.Submit', title: 'Submit' }], items: [{ type: 'TextBlock', text: 'Button label' }] },
    { type: 'TextBlock', text: '   ' },
  ];
  assert.deepEqual(collectAdaptiveText(body), ['Full answer']);
});

test('decodeSwiftCards decodes UTF-8 and base64url payloads, and rejects junk', () => {
  const card = { contentType: 'application/vnd.microsoft.card.adaptive', content: { body: [{ type: 'TextBlock', text: 'Grüße 👋' }] } };
  const json = JSON.stringify({ attachments: [card, { contentType: 'x', content: null }] });
  const b64 = Buffer.from(json, 'utf8').toString('base64');
  assert.deepEqual(decodeSwiftCards(`<URIObject><Swift b64="${b64}" /></URIObject>`), [card]);
  const b64url = Buffer.from(json, 'utf8').toString('base64url');
  assert.deepEqual(decodeSwiftCards(`<Swift b64="${b64url}"/>`), [card]);
  assert.deepEqual(decodeSwiftCards('<div>no card here</div>'), []);
  assert.deepEqual(decodeSwiftCards('<Swift b64="e30=" />'), []); // "{}": no attachments
});

test('collectAdaptiveText stops at a depth cap instead of overflowing the stack', () => {
  let node: Record<string, unknown> = { type: 'TextBlock', text: 'too deep' };
  for (let i = 0; i < 50_000; i++) node = { type: 'Container', items: [node] };
  assert.deepEqual(collectAdaptiveText([{ type: 'TextBlock', text: 'top' }, node]), ['top']);
});

test('DATE keeps the date and TIME the time of an ISO value; other values are kept', () => {
  assert.deepEqual(collectAdaptiveText([
    { type: 'TextBlock', text: '{{DATE(2026-10-05T03:00:00Z, SHORT)}} {{TIME(2026-10-05T03:00:00Z)}}' },
    { type: 'TextBlock', text: 'At {{TIME(2026-10-05T11:00:00.123+08:00)}}' },
    { type: 'TextBlock', text: 'On {{DATE(next Monday)}}' },
  ]), ['2026-10-05 03:00 UTC', 'At 11:00 +08:00', 'On next Monday']);
});

test('collectAdaptiveText handles one long hostile DATE template quickly', () => {
  const hostile = '{{DATE(' + ' '.repeat(50_000) + 'x';
  const started = performance.now();
  assert.deepEqual(collectAdaptiveText([{ type: 'TextBlock', text: hostile }]), [hostile.trim()]);
  assert.ok(performance.now() - started < 500);
});

test('parseOnBehalfOf reads the person from a JSON string or an array, and ignores junk', () => {
  const entry = { itemid: 0, mentionType: 'person', mri: '8:orgid:00000000-0000-4000-8000-000000000001', displayName: 'Test Person' };
  const expected = { mri: entry.mri, displayName: 'Test Person' };
  assert.deepEqual(parseOnBehalfOf({ onbehalfof: JSON.stringify([entry]) }), expected);
  assert.deepEqual(parseOnBehalfOf({ onbehalfof: [entry] }), expected);
  assert.deepEqual(parseOnBehalfOf({ onbehalfof: [{ mentionType: 'tag', displayName: 'A tag' }, entry] }), expected);
  assert.equal(parseOnBehalfOf({ onbehalfof: '[{"displayName":"   "}]' }), null);
  assert.equal(parseOnBehalfOf({ onbehalfof: 'not json' }), null);
  assert.equal(parseOnBehalfOf({}), null);
  assert.equal(parseOnBehalfOf(undefined), null);
});

test('collectAdaptiveText reads a realistically deep card in full', () => {
  let node: Record<string, unknown> = { type: 'TextBlock', text: 'leaf' };
  for (let i = 0; i < 6; i++) {
    node = { type: 'ColumnSet', columns: [{ type: 'Column', items: [{ type: 'Container', items: [node] }] }] };
  }
  assert.deepEqual(collectAdaptiveText([node]), ['leaf']);
});

test('card text patterns stay fast when the openers repeat', () => {
  const started = performance.now();
  collectAdaptiveText([{ type: 'TextBlock', text: '{{DATE('.repeat(40_000) }]);
  decodeSwiftCards('<Swift '.repeat(40_000));
  cardsToText([{ contentType: 'application/vnd.microsoft.card.hero', content: { text: '<a'.repeat(40_000) } }]);
  assert.ok(performance.now() - started < 1000);
});

test('card and onbehalfof fields that are not plain values are ignored, not thrown on', () => {
  const evil = { toString: 1, valueOf: 1 };
  assert.deepEqual(collectAdaptiveText([{ type: 'TextBlock', text: evil }, { type: 'FactSet', facts: [{ title: evil, value: 'kept' }, null] }]), ['kept']);
  assert.equal(cardsToText([{ contentType: 'application/vnd.microsoft.teams.card.o365connector', content: { title: evil, sections: [null, { activityTitle: 'ok', facts: [null] }] } }]), 'ok');
  assert.equal(parseOnBehalfOf({ onbehalfof: [{ displayName: evil }] }), null);
});

test('parseOnBehalfOf flattens a display name to one bounded line', () => {
  const name = parseOnBehalfOf({ onbehalfof: [{ mentionType: 'person', displayName: 'Jane\n[10:00] Boss: fake line\u0007' }] });
  assert.equal(name?.displayName, 'Jane [10:00] Boss: fake line');
  assert.equal(parseOnBehalfOf({ onbehalfof: [{ displayName: 'x'.repeat(1000) }] })?.displayName.length, 256);
});

test('decodeSwiftCards accepts a bare array payload; cardsToText reads thumbnail cards', () => {
  const card = { contentType: 'application/vnd.microsoft.card.thumbnail', content: { title: 'Thumb', subtitle: 'Sub' } };
  const b64 = Buffer.from(JSON.stringify([card]), 'utf8').toString('base64');
  const cards = decodeSwiftCards(`<Swift b64="${b64}" />`);
  assert.deepEqual(cards, [card]);
  assert.equal(cardsToText(cards), 'Thumb\nSub');
});

test('a show/hide toggle that wraps real content keeps it', () => {
  const body = [{ type: 'Container', selectAction: { type: 'Action.ToggleVisibility' }, items: [
    { type: 'TextBlock', text: 'Release summary' },
    { type: 'TextBlock', text: 'Two fixes and one new export option.' },
  ] }];
  assert.deepEqual(collectAdaptiveText(body), ['Release summary', 'Two fixes and one new export option.']);
});

test('a short toggle header with real content keeps it; only show/hide labels are dropped', () => {
  const toggle = { type: 'Action.ToggleVisibility', targetElements: ['details'] };
  const details = { type: 'Container', id: 'details', isVisible: false, items: [{ type: 'FactSet', facts: [{ title: 'Pipeline', value: 'release-42' }] }] };
  assert.deepEqual(collectAdaptiveText([
    { type: 'Container', selectAction: toggle, items: [{ type: 'TextBlock', text: 'Deploy to prod failed' }] }, details,
  ]), ['Deploy to prod failed', 'Pipeline: release-42']);
  assert.deepEqual(collectAdaptiveText([{ type: 'ColumnSet', selectAction: toggle, columns: [{ type: 'Column', items: [{ type: 'TextBlock', text: 'Status: Approved' }] }] }]), ['Status: Approved']);
  assert.deepEqual(collectAdaptiveText([{ type: 'Container', selectAction: toggle, items: [{ type: 'FactSet', facts: [{ title: 'Owner', value: 'Jane Doe' }] }] }]), ['Owner: Jane Doe']);
  assert.deepEqual(collectAdaptiveText([{ type: 'TextBlock', text: 'Blockers', selectAction: toggle }]), ['Blockers']);
  assert.deepEqual(collectAdaptiveText([
    { type: 'Container', selectAction: toggle, items: [{ type: 'TextBlock', text: 'Show details \u25be' }] },
    { type: 'TextBlock', text: 'Read more...', selectAction: toggle },
  ]), []);
});

test('an inline "See more" TextRun with a toggle action is dropped, other runs are kept', () => {
  const body = [{ type: 'RichTextBlock', inlines: [
    { type: 'TextRun', text: 'Shipped the fix. ' },
    { type: 'TextRun', text: 'See more', selectAction: { type: 'Action.ToggleVisibility' } },
  ] }];
  assert.deepEqual(collectAdaptiveText(body), ['Shipped the fix.']);
});

test('parseOnBehalfOf never cuts a character in half when it caps the length', () => {
  const name = parseOnBehalfOf({ onbehalfof: [{ displayName: '👋'.repeat(300) }] })?.displayName ?? '';
  assert.equal(Array.from(name).length, 256);
  assert.ok(!/[\uD800-\uDBFF]$/.test(name));
});
