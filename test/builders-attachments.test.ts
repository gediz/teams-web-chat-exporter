// Behavior lock for the attachment-summary + failure-transparency helpers in
// src/background/builders.ts. These drive the TXT/CSV attachment labels and the
// "N of M images could not be included" banner + IMAGES_FAILED.txt manifest that
// users rely on to know what did NOT make it into the export.
import { test, expect } from 'vitest';
import {
  summarizeAttachments,
  collectAttachmentOutcomes,
  imageSummaryLine,
  buildImagesFailedManifest,
  removeAvatars,
} from '../src/background/builders.ts';

const msg = (attachments: unknown[]) => ({ attachments } as never);

test('summarizeAttachments labels image/video/audio/file/link and appends the failure reason', () => {
  expect(summarizeAttachments(msg([{ type: 'png', label: 'photo.png' }]))).toBe('[image: photo.png]');
  expect(summarizeAttachments(msg([{ type: 'png', label: '' }]))).toBe('[image]');
  expect(summarizeAttachments(msg([{ type: 'mp4', label: 'clip.mp4' }]))).toBe('[video: clip.mp4]');
  expect(summarizeAttachments(msg([{ type: 'mp3', label: '' }]))).toBe('[audio]');
  expect(summarizeAttachments(msg([{ type: 'pdf', label: 'doc.pdf' }]))).toBe('[file: doc.pdf]');
  expect(summarizeAttachments(msg([{ kind: 'preview', metaText: 'Cool Site\nhttps://x' }]))).toBe('[link: Cool Site]');
  expect(summarizeAttachments(msg([{ type: 'png', label: 'a.png', failReason: 'expired' }]))).toBe('[image: a.png] (expired)');
  expect(summarizeAttachments(msg([]))).toBe('');
});

test('summarizeAttachments caps at 3 with a "+N more", and skipPreviews drops link cards', () => {
  const five = summarizeAttachments(msg([
    { type: 'pdf', label: '1.pdf' }, { type: 'pdf', label: '2.pdf' },
    { type: 'pdf', label: '3.pdf' }, { type: 'pdf', label: '4.pdf' }, { type: 'pdf', label: '5.pdf' },
  ]));
  expect(five).toBe('[file: 1.pdf] [file: 2.pdf] [file: 3.pdf] [+2 more]');
  const skipped = summarizeAttachments(
    msg([{ kind: 'preview', metaText: 'Site\nurl' }, { type: 'pdf', label: 'x.pdf' }]),
    { skipPreviews: true },
  );
  expect(skipped).toBe('[file: x.pdf]');
});

test('collectAttachmentOutcomes dedups by href and lets a success cancel a failure of the same URL', () => {
  const s = collectAttachmentOutcomes([
    msg([
      { href: 'h1', failReason: 'expired' },
      { href: 'h2', failReason: 'sign-in' },
      { href: 'h3', dataUrl: 'data:ok' },
    ]),
    msg([
      { href: 'h1', failReason: 'expired' }, // dup of a failure -> counted once
      { href: 'h2', dataUrl: 'data:ok' },    // h2 succeeded here -> cancels its failure
    ]),
  ]);
  expect(s.failed).toBe(1);            // only h1 remains failed
  expect(s.byReason).toEqual({ expired: 1 });
  expect(s.total).toBe(3);             // ok: h2,h3 + failed: h1
});

test('imageSummaryLine formats count + reasons (most common first), pluralizes, and is empty when nothing failed', () => {
  expect(imageSummaryLine({ total: 5, failed: 3, byReason: { expired: 2, 'sign-in': 1 } }))
    .toBe('3 of 5 images could not be included (2 expired, 1 sign-in).');
  expect(imageSummaryLine({ total: 1, failed: 1, byReason: { removed: 1 } }))
    .toBe('1 of 1 image could not be included (1 removed).');
  expect(imageSummaryLine(undefined)).toBe('');
  expect(imageSummaryLine({ total: 5, failed: 0, byReason: {} })).toBe('');
});

test('buildImagesFailedManifest emits a legend + one TSV row per distinct failure, empty when none', () => {
  const man = buildImagesFailedManifest([
    { author: 'Alice', timestamp: '2024-01-01T10:00:00Z', attachments: [
      { href: 'https://x.com/a.png', failReason: 'expired', label: 'a.png' },
    ] },
  ] as never);
  expect(man).toContain('# Images and file attachments');
  expect(man).toContain('unavailable = not embeddable'); // legend covers the fallthrough reason
  const rows = man.split('\n').filter((l) => l && !l.startsWith('#'));
  expect(rows).toHaveLength(1);
  expect(rows[0].split('\t')).toEqual(['Alice', '2024-01-01T10:00:00Z', 'expired', 'x.com', 'a.png']);
  // No failures -> empty string (no file written).
  expect(buildImagesFailedManifest([msg([{ href: 'h', dataUrl: 'd' }])])).toBe('');
});

test('removeAvatars strips avatar/avatarId and leaves avatar-free messages untouched', () => {
  const out = removeAvatars([
    { author: 'A', avatarId: 'av1', text: 'hi' },
    { author: 'B', text: 'yo' },
  ] as never);
  expect('avatarId' in out[0]).toBe(false);
  expect(out[0].author).toBe('A');
  expect(out[1]).toEqual({ author: 'B', text: 'yo' });
});
