// Behavior lock for toCSV in src/background/builders.ts.
//
// Covers the format contract that spreadsheet/programmatic consumers rely on:
//   - the column set and order (new: reply_to, reactions, is_own, message_type)
//   - one cell per header column on every data row (no dropped/extra push)
//   - the leading UTF-8 BOM (so Excel reads emoji/CJK/Turkish, not mojibake)
//   - RFC 4180 quoting + the OWASP CSV-injection guard (all five triggers)
//   - the human reactions summary (every reactor by name, no "+N")
import { test } from 'vitest';
import assert from 'node:assert/strict';
import { toCSV } from '../src/background/builders.ts';

const HEADER = [
  'id', 'author', 'timestamp', 'text', 'edited', 'deleted', 'system', 'subject',
  'importance', 'mentions', 'reactions_json', 'attachments_json', 'forwarded',
  'reply_to', 'reactions', 'is_own', 'message_type',
];

// Header line with the BOM stripped, split into column names.
const headerCols = (csv) => csv.replace(/^﻿/, '').split('\n')[0].split(',');

// Parse one CSV line (every cell double-quoted, no embedded newlines in these
// test rows) into its unquoted field values.
const cells = (line) => Array.from(line.matchAll(/"((?:[^"]|"")*)"/g), (m) => m[1].replace(/""/g, '"'));
const col = (name) => HEADER.indexOf(name);
// Data rows (BOM + header stripped) parsed into field arrays. Only safe for
// fixtures whose cells contain no literal newline (all rows below qualify).
const dataRows = (csv) => csv.replace(/^﻿/, '').split('\n').slice(1).map(cells);

test('prepends a UTF-8 BOM so Excel opens it as UTF-8', () => {
  const csv = toCSV([{ id: '1', author: 'A', text: 'hi 👍' }]);
  assert.equal(csv.codePointAt(0), 0xfeff);
  // The BOM sits before everything, including a partial-export banner.
  const withBanner = toCSV([{ id: '1' }], { partial: { reason: 'network' } });
  assert.equal(withBanner.codePointAt(0), 0xfeff);
  assert.ok(withBanner.replace(/^﻿/, '').startsWith('# WARNING'));
});

test('emits the full column set in a stable order', () => {
  assert.deepEqual(headerCols(toCSV([])), HEADER);
});

test('every data row has exactly one cell per header column', () => {
  // A dropped or extra row.push() would desync the columns; lock the count.
  const rows = dataRows(toCSV([
    { id: '1', author: 'A', text: 'plain' },
    { id: '2', author: 'B', text: 'y', isOwn: true, messageType: 'Text',
      reactions: [{ emoji: '👍', count: 1, reactors: [{ name: 'Z' }] }],
      replyTo: { author: 'J', timestamp: '', text: 'q' },
      forwarded: { originalAuthor: 'Ext', originalText: 'FYI' } },
  ]));
  assert.equal(rows.length, 2);
  for (const r of rows) assert.equal(r.length, HEADER.length, `${HEADER.length} cells per row`);
});

test('reply_to carries "author: text", empty when there is no quoted text', () => {
  const rows = dataRows(toCSV([
    { id: '1', author: 'Mary', text: 'hi', replyTo: { author: 'John', timestamp: '', text: 'Hello' } },
    { id: '2', author: 'Al', text: 'plain' },
    // placeholder replyTo with empty text must NOT produce a reply_to value
    { id: '3', author: 'Zoe', text: 'x', replyTo: { author: 'X', timestamp: '', text: '' } },
  ]));
  const rt = col('reply_to');
  assert.equal(rows[0][rt], 'John: Hello', 'reply carries author: text');
  assert.equal(rows[1][rt], '', 'non-reply row has empty reply_to');
  assert.equal(rows[2][rt], '', 'empty-text placeholder yields empty reply_to');
});

test('reply_to with no author omits the prefix', () => {
  const [row] = dataRows(toCSV([{ id: '1', author: 'Al', text: 'ok', replyTo: { author: '', timestamp: '', text: 'context only' } }]));
  assert.equal(row[col('reply_to')], 'context only');
});

test('RFC 4180 quoting doubles embedded quotes and keeps newlines literal', () => {
  const csv = toCSV([{ id: '1', author: 'A', text: 'he said "hi"\nbye' }]);
  assert.ok(csv.includes('"he said ""hi""\nbye"'), 'quotes doubled, newline kept in-quote');
});

test('CSV-injection guard prefixes an apostrophe to all five formula triggers', () => {
  const csv = toCSV([{ id: '1', author: '=1+2', text: '@cmd', subject: '-x', importance: '+y' }]);
  assert.ok(csv.includes('"\'=1+2"'), 'leading = guarded');
  assert.ok(csv.includes('"\'@cmd"'), 'leading @ guarded');
  assert.ok(csv.includes('"\'-x"'), 'leading - guarded');
  assert.ok(csv.includes('"\'+y"'), 'leading + guarded');
  // Tab and carriage-return also trigger. Use author, which (unlike text) is
  // not \r->\n normalized, so a leading \r survives to reach the guard.
  assert.ok(toCSV([{ id: '1', author: '\tx' }]).includes('"\'\tx"'), 'leading tab guarded');
  assert.ok(toCSV([{ id: '1', author: '\rx' }]).includes('"\'\rx"'), 'leading CR guarded');
  // A benign cell is left untouched (no spurious apostrophe).
  const [row] = dataRows(toCSV([{ id: '1', author: 'Normal', text: 'hello' }]));
  assert.equal(row[col('author')], 'Normal', 'benign cell not prefixed');
});

test('reactions column lists every reactor by name, never "+N"', () => {
  const rows = dataRows(toCSV([
    { id: '1', author: 'A', text: 't', reactions: [
      { emoji: '👍', count: 2, reactors: [{ name: 'Alice' }, { self: true }] },
      { emoji: '❤️', count: 1, reactors: [{ name: 'Carol' }] },
    ] },
    // DOM-fallback shape: count (5) exceeds the named reactors — still no "+3"
    { id: '2', author: 'B', text: 't', reactions: [{ emoji: '🔥', count: 5, reactors: [{ name: 'Alice' }, { name: 'Bob' }] }] },
    // no names resolved at all -> "×count" fallback; single unnamed -> bare emoji
    { id: '3', author: 'C', text: 't', reactions: [{ emoji: '🎉', count: 5, reactors: [] }] },
    { id: '4', author: 'D', text: 't', reactions: [{ emoji: '👀', count: 1 }] },
  ]));
  const rc = col('reactions');
  // Assert the reactions CELL exactly (not a whole-doc substring, which the
  // reactions_json column would also satisfy).
  assert.equal(rows[0][rc], '👍 Alice, You · ❤️ Carol', 'self as You, names joined');
  assert.equal(rows[1][rc], '🔥 Alice, Bob', 'lists known names, no +3');
  assert.equal(rows[2][rc], '🎉 ×5', 'no names -> ×count');
  assert.equal(rows[3][rc], '👀', 'single unnamed -> bare emoji, no ×1');
});

test('is_own and message_type columns reflect the message', () => {
  const rows = dataRows(toCSV([
    { id: '1', author: 'A', text: 't', isOwn: true, messageType: 'Event/Call' },
    { id: '2', author: 'B', text: 't' }, // isOwn absent -> false, messageType absent -> empty
  ]));
  assert.equal(rows[0][col('is_own')], 'true');
  assert.equal(rows[0][col('message_type')], 'Event/Call');
  assert.equal(rows[1][col('is_own')], 'false', 'absent isOwn -> false');
  assert.equal(rows[1][col('message_type')], '', 'absent messageType -> empty');
});

test('surfaces the failed-image summary as a leading # comment', () => {
  // Failure-transparency banner (builders.ts imagesBanner). Reasons are ordered
  // most-common first. Without this a dropped/reworded banner would pass silently.
  const csv = toCSV([{ id: '1', author: 'A', text: 'hi' }], {
    attachmentStats: { total: 5, failed: 3, byReason: { expired: 2, 'sign-in': 1 } },
  });
  assert.equal(csv.codePointAt(0), 0xfeff, 'BOM stays first');
  assert.ok(
    csv.replace(/^﻿/, '').includes('# 3 of 5 images could not be included (2 expired, 1 sign-in).'),
    'CSV carries the failed-image summary banner as a # comment',
  );
});
