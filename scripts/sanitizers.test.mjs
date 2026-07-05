// Behavior lock for the filename sanitizers in src/utils/messages.ts.
//
// Written BEFORE the 2026-07 lint cleanup rewrote the invisible-character
// regex literals (U+200B..U+FEFF etc.) as \uXXXX escapes, so a pass after
// that rewrite proves the escaped form strips exactly the same characters.
// Keep these cases in sync with the ranges documented above the regexes:
//   U+200B-U+200F zero-widths, U+202A-U+202E bidi overrides,
//   U+2060-U+2064 word joiner + invisible operators, U+FEFF BOM.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { sanitizeBase, sanitizeFileName } from '../src/utils/messages.ts';

test('sanitizeBase strips zero-width and bidi characters entirely', () => {
  // One character from the edge of every documented range.
  assert.equal(sanitizeBase('Jane\u200BDoe'), 'JaneDoe'); // zero-width space (range start)
  assert.equal(sanitizeBase('Jane\u200FDoe'), 'JaneDoe'); // RLM (range end)
  assert.equal(sanitizeBase('Jane\u202ADoe'), 'JaneDoe'); // LRE (range start)
  assert.equal(sanitizeBase('Jane\u202EDoe'), 'JaneDoe'); // RLO (range end)
  assert.equal(sanitizeBase('Jane\u2060Doe'), 'JaneDoe'); // word joiner (range start)
  assert.equal(sanitizeBase('Jane\u2064Doe'), 'JaneDoe'); // invisible plus (range end)
  assert.equal(sanitizeBase('Jane\uFEFFDoe'), 'JaneDoe'); // BOM
  // Stacked invisibles collapse to nothing, not dashes.
  assert.equal(sanitizeBase('Jane\u200B\u202E\uFEFFDoe'), 'JaneDoe');
});

test('sanitizeBase leaves neighbouring characters alone', () => {
  // NBSP and hair space are \s, so they collapse to a single space instead.
  assert.equal(sanitizeBase('Jane\u00A0Doe'), 'Jane Doe');
  assert.equal(sanitizeBase('Jane\u200ADoe'), 'Jane Doe');
  // Control chars still become dashes (separate regex, unchanged).
  assert.equal(sanitizeBase('Jane\x01Doe'), 'Jane-Doe');
});

test('sanitizeFileName strips the same invisibles but keeps the extension', () => {
  assert.equal(sanitizeFileName('re\u200Bport\u202E.pdf'), 'report.pdf');
  assert.equal(sanitizeFileName('\uFEFFnotes\u2060.txt'), 'notes.txt');
  assert.equal(sanitizeFileName('plan\u00A0v2.docx'), 'plan v2.docx'); // NBSP -> space
});
