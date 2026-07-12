// Behavior lock for toPlainText in src/background/download.ts.
//
// The headline case: a quoted reply must render the quote ABOVE the body
// ("quote, then response"), matching the HTML and PDF exports. It used to be
// appended after the body, which inverted the reading order. Also locks the
// full reactor list (no "+N") and the shapes that must NOT change (plain and
// forwarded messages).
import { test } from 'vitest';
import assert from 'node:assert/strict';
import { toPlainText } from '../src/background/download.ts';

const reply = { author: 'Mary', timestamp: '10:03', text: 'hi', replyTo: { author: 'John', timestamp: '', text: 'Hello' } };
const plain = { author: 'John', timestamp: '10:01', text: 'Hello' };

test('renders the reply quote ABOVE the body (header, quote, indented body)', () => {
  assert.equal(toPlainText([reply]), '[10:03] Mary:\n  > John: Hello\n    hi');
});

test('the quote precedes the reply body in document order', () => {
  const out = toPlainText([plain, reply]);
  assert.ok(out.indexOf('> John: Hello') < out.indexOf('    hi'), 'quote comes before the body');
});

test('a plain (non-reply) message keeps the compact single line', () => {
  assert.equal(toPlainText([plain]), '[10:01] John: Hello');
});

test('a multi-line reply body is uniformly 4-space indented under the quote', () => {
  const m = { author: 'X', timestamp: '10:00', text: 'a\nb', replyTo: { author: 'John', timestamp: '', text: 'q' } };
  assert.equal(toPlainText([m]), '[10:00] X:\n  > John: q\n    a\n    b');
});

test('a reply whose quote has no author omits the "author:" prefix', () => {
  const m = { author: 'Al', timestamp: '10:05', text: 'ok', replyTo: { author: '', timestamp: '', text: 'context only' } };
  assert.equal(toPlainText([m]), '[10:05] Al:\n  > context only\n    ok');
});

test('a placeholder replyTo with empty text produces no quote line', () => {
  const m = { author: 'Zoe', timestamp: '10:12', text: 'plain', replyTo: { author: 'X', timestamp: '', text: '' } };
  assert.equal(toPlainText([m]), '[10:12] Zoe: plain');
});

test('reactions list every reactor by name, never "+N", and ride the message line', () => {
  const m = { author: 'A', timestamp: '10:07', text: 'agreed',
    reactions: [{ emoji: '🔥', count: 5, reactors: [{ name: 'Alice' }, { name: 'Bob' }] }] };
  const out = toPlainText([m]);
  assert.equal(out, '[10:07] A: agreed  [🔥 Alice, Bob]');
  assert.ok(!out.includes('+3'), 'no "+N"');
});

test('forwarded messages (no reply) keep their original shape', () => {
  const withText = { author: 'Kim', timestamp: '10:09', text: 'see below', forwarded: { originalAuthor: 'Ext', originalText: 'FYI' } };
  const noText = { author: 'Kim', timestamp: '10:08', text: '', forwarded: { originalAuthor: 'Ext', originalText: 'FYI' } };
  assert.equal(toPlainText([withText]), '[10:09] Kim: see below\n  [forwarded from Ext]: FYI');
  assert.equal(toPlainText([noText]), '[10:08] Kim [forwarded from Ext]: FYI');
});

test('a forwarded message that is also a reply puts the quote before the forward', () => {
  const m = { author: 'Kim', timestamp: '10:10', text: 'relevant',
    forwarded: { originalAuthor: 'Ext', originalText: 'FYI' },
    replyTo: { author: 'John', timestamp: '', text: 'any updates?' } };
  const out = toPlainText([m]);
  assert.ok(out.indexOf('> John: any updates?') < out.indexOf('[forwarded from Ext]'), 'reply precedes forward');
});

test('renders the failed-image summary line', () => {
  // Failure-transparency banner in TXT (download.ts toPlainText). Reasons are
  // ordered most-common first; a dropped/reworded line would otherwise pass.
  const out = toPlainText([plain], {
    attachmentStats: { total: 5, failed: 3, byReason: { expired: 2, 'sign-in': 1 } },
  });
  assert.ok(
    out.includes('3 of 5 images could not be included (2 expired, 1 sign-in).'),
    'TXT carries the failed-image summary line',
  );
});
