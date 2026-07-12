// Behavior lock for the pure helpers in src/background/pdf.ts (the biggest
// coverage sink). These carry the link-tail trimming, timestamp/URI encoding,
// CJK/KR font-routing, image-data decoding, and reactor-label logic that shape
// visible PDF output — testable in isolation without touching pdf-lib.
import { test, expect } from 'vitest';
import {
  trimTrailingPunct, formatTimestamp, cpRoutesToKr, cpRoutesToCjk,
  dataUrlToBytes, encodeUriForPdf, reactorInitials, reactorNames,
} from '../src/background/pdf.ts';

test('trimTrailingPunct strips trailing sentence punctuation and brackets', () => {
  expect(trimTrailingPunct('https://x.com/page).')).toBe('https://x.com/page');
  expect(trimTrailingPunct('https://x.com/a,')).toBe('https://x.com/a');
  expect(trimTrailingPunct('https://x.com/a')).toBe('https://x.com/a');
});

test('formatTimestamp yields YYYY-MM-DD HH:MM, empty for none, echoes an unparseable input', () => {
  expect(formatTimestamp('')).toBe('');
  expect(formatTimestamp(undefined)).toBe('');
  expect(formatTimestamp('not-a-date')).toBe('not-a-date');
  // Exact value is local-timezone dependent; lock the shape.
  expect(formatTimestamp('2024-03-05T09:07:00Z')).toMatch(/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$/);
});

test('cpRoutesToKr matches Hangul only', () => {
  expect(cpRoutesToKr(0xAC00)).toBe(true); // 가
  expect(cpRoutesToKr(0x1100)).toBe(true); // Hangul Jamo
  expect(cpRoutesToKr(0x4E00)).toBe(false); // CJK 一
  expect(cpRoutesToKr(0x41)).toBe(false); // 'A'
});

test('cpRoutesToCjk matches CJK/Kana but not Hangul or Latin', () => {
  expect(cpRoutesToCjk(0x4E00)).toBe(true); // 一
  expect(cpRoutesToCjk(0x3042)).toBe(true); // あ hiragana
  expect(cpRoutesToCjk(0xAC00)).toBe(false); // Hangul routes to KR, checked first
  expect(cpRoutesToCjk(0x41)).toBe(false);
});

test('dataUrlToBytes decodes an image data URL and rejects non-image/invalid input', () => {
  const r = dataUrlToBytes('data:image/png;base64,SGk='); // base64 of "Hi"
  expect(r?.mime).toBe('image/png');
  expect([...(r?.bytes ?? [])]).toEqual([72, 105]);
  expect(dataUrlToBytes('data:text/plain;base64,SGk=')).toBe(null); // only image/* matches
  expect(dataUrlToBytes('nope')).toBe(null);
});

test('encodeUriForPdf escapes non-ASCII, spaces, and PDF-special parens/backslash', () => {
  expect(encodeUriForPdf('https://x.com/a b')).toBe('https://x.com/a%20b');
  expect(encodeUriForPdf('https://x.com/(x)')).toBe('https://x.com/\\(x\\)');
  expect(encodeUriForPdf('https://x.com/é')).toBe('https://x.com/%C3%A9'); // é -> percent-encoded
  expect(encodeUriForPdf('https://x.com/plain')).toBe('https://x.com/plain');
});

test('reactorInitials returns up to two uppercase initials, "?" as fallback', () => {
  expect(reactorInitials('Alice Bob')).toBe('AB');
  expect(reactorInitials('alice')).toBe('A');
  expect(reactorInitials('')).toBe('?');
});

test('reactorNames joins reactor names, "You" for self, empty for none', () => {
  expect(reactorNames({ emoji: '👍', count: 2, reactors: [{ name: 'Alice' }, { name: 'Bob', self: true }] } as never)).toBe('Alice, You');
  expect(reactorNames({ emoji: '👍', count: 0 } as never)).toBe('');
});
