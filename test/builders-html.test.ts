// Behavior lock for toHTML (src/background/builders.ts) — the default human-facing
// export format. Focuses on the high-value, low-brittleness guarantees: HTML
// escaping (a prior real bug in this repo was double-escaping), the two warning
// banners (partial export + failed images), and basic message rendering.
import { test, expect } from 'vitest';
import { toHTML } from '../src/background/builders.ts';

const html = (rows: unknown[], meta: unknown = {}) => toHTML(rows as never, meta as never).join('');
const at = '2024-01-01T10:00:00Z';

test('toHTML escapes HTML in message text, and never double-escapes', () => {
  const out = html([{ id: '1', author: 'Alice', timestamp: at, text: '<script>alert(1)</script> & "hi"' }]);
  expect(out).toContain('&lt;script&gt;');
  expect(out).toContain('&amp;');
  expect(out).not.toContain('<script>alert'); // the raw tag from message text must not survive
  expect(out).not.toContain('&amp;amp;');      // not double-escaped
});

test('toHTML renders the author (escaped) and the body text', () => {
  const out = html([{ id: '1', author: 'Bob <admin>', timestamp: at, text: 'hello there' }]);
  expect(out).toContain('Bob &lt;admin&gt;');
  expect(out).toContain('hello there');
});

test('toHTML shows the partial-export warning banner when meta.partial is set', () => {
  const out = html([{ id: '1', author: 'A', timestamp: at, text: 'x' }], { partial: { reason: 'network' } });
  expect(out).toContain('partial-warning');
  expect(out).toContain('Export may be incomplete');
});

test('toHTML shows the failed-images banner from meta.attachmentStats', () => {
  const out = html(
    [{ id: '1', author: 'A', timestamp: at, text: 'x' }],
    { attachmentStats: { total: 2, failed: 1, byReason: { expired: 1 } } },
  );
  expect(out).toContain('images-warning');
  expect(out).toContain('1 of 2 images could not be included (1 expired).');
});

test('toHTML with no messages still returns a valid (non-empty) document', () => {
  const out = html([]);
  expect(out).toContain('<'); // produces markup, does not throw
});
