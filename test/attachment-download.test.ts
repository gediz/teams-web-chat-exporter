// Behavior lock for the pure helpers in src/background/attachment-download.ts.
// These govern WHICH files get written (deny-list, host gate, dedup) and under
// WHAT on-disk path/name — a bug here drops files the user wanted or writes
// outside the export folder.
import { test, expect } from 'vitest';
import {
  parseExtDenyList,
  isDeniedByType,
  buildAttachmentPath,
  attachmentDatePrefix,
  collectDocumentAttachments,
  toFilesSummaryWire,
} from '../src/background/attachment-download.ts';

test('parseExtDenyList splits, trims, strips leading dots, lowercases; keeps multi-dot', () => {
  expect([...parseExtDenyList('pdf, .ZIP ,exe')].sort()).toEqual(['exe', 'pdf', 'zip']);
  expect([...parseExtDenyList('..tar.gz')]).toEqual(['tar.gz']); // multi-dot stays one token
  expect(parseExtDenyList('').size).toBe(0);
  expect(parseExtDenyList(undefined).size).toBe(0);
});

test('isDeniedByType matches by extension case-insensitively; no-extension always allowed', () => {
  const deny = parseExtDenyList('pdf, tar.gz');
  expect(isDeniedByType('Report.PDF', deny)).toBe(true);
  expect(isDeniedByType('archive.tar.gz', deny)).toBe(true);
  expect(isDeniedByType('archive.gz', deny)).toBe(false); // .gz alone is not denied
  expect(isDeniedByType('noextension', deny)).toBe(false);
  expect(isDeniedByType('x.pdf', new Set())).toBe(false); // empty deny denies nothing
  expect(isDeniedByType('', deny)).toBe(false);
});

test('buildAttachmentPath assembles <folder>/[<chat>/]attachments/<name>, drops empty segments, sanitizes the name', () => {
  expect(buildAttachmentPath('Export_2026', '', 'f.pdf')).toBe('Export_2026/attachments/f.pdf');
  expect(buildAttachmentPath('Export_2026', 'Chat A', 'f.pdf')).toBe('Export_2026/Chat A/attachments/f.pdf');
  // A path separator in the name must NOT inject a new directory level.
  const p = buildAttachmentPath('base', '', 'a/b:1.pdf');
  expect(p.startsWith('base/attachments/')).toBe(true);
  expect(p.slice('base/attachments/'.length)).not.toMatch(/[/\\:]/);
});

test('attachmentDatePrefix emits a compact UTC stamp, or "" for invalid/out-of-range dates', () => {
  expect(attachmentDatePrefix('2026-03-27T12:50:05Z')).toBe('20260327T125005Z__');
  expect(attachmentDatePrefix('not-a-date')).toBe('');
  expect(attachmentDatePrefix('1970-01-01T00:00:00Z')).toBe(''); // before fflate/DOS floor 1980
});

test('collectDocumentAttachments dedups by href, skips images + non-SharePoint hosts, adopts a later itemid', () => {
  const msgs = [
    { attachments: [
      { href: 'https://contoso.sharepoint.com/sites/x/report.pdf', label: 'report.pdf', type: 'pdf' },
      { href: 'https://contoso.sharepoint.com/sites/x/pic.png', label: 'pic.png', type: 'png' }, // image -> skip
      { href: 'https://example.com/other.pdf', label: 'other.pdf', type: 'pdf' }, // non-SharePoint -> skip
      { label: 'no-href.pdf', type: 'pdf' }, // no href -> skip
    ] },
    { attachments: [
      { href: 'https://contoso.sharepoint.com/sites/x/report.pdf', label: 'report.pdf', type: 'pdf', itemid: 'guid-1' },
    ] },
  ];
  const out = collectDocumentAttachments(msgs as never);
  expect(out).toHaveLength(1);
  expect(out[0].href).toBe('https://contoso.sharepoint.com/sites/x/report.pdf');
  expect(out[0].name).toBe('report.pdf');
  expect(out[0].itemid).toBe('guid-1'); // adopted from the second (deduped) occurrence
});

test('toFilesSummaryWire projects the count fields and drops the failures array', () => {
  const wire = toFilesSummaryWire({
    total: 5, saved: 3, links: 1, failed: 1, cancelled: 0, skipped: 0,
    failures: [{ name: 'x', reason: 'no access' }],
  });
  expect(wire).toEqual({ total: 5, saved: 3, links: 1, failed: 1, cancelled: 0, skipped: 0 });
  expect('failures' in wire).toBe(false);
});
