// Behavior lock for src/background/zip.ts — the streaming zip that every 2+-format
// and multi-chat bundle export ships as. A CRC, framing, store-vs-deflate, or path
// bug here silently corrupts every user's bundle.zip, so the core guarantee is a
// byte-identical round-trip: build -> re-parse with fflate's independent unzipSync
// -> every entry's bytes and path come back unchanged.
import { test, expect } from 'vitest';
import { unzipSync } from 'fflate';
import { buildZipAsync, createZipStream } from '../src/background/zip.ts';

const enc = (s: string) => new TextEncoder().encode(s);
const zipBytes = async (blob: Blob) => new Uint8Array(await blob.arrayBuffer());

test('buildZipAsync round-trips every entry byte-identically and preserves nested paths', async () => {
  const files = [
    { path: 'messages.txt', data: enc('hello world\nsecond line') },
    { path: 'images/photo.png', data: new Uint8Array([0, 1, 2, 3, 250, 251, 252, 255]) },
    { path: 'a\\b\\c.json', data: enc('{"k":1}') }, // backslashes normalize to forward slashes
  ];
  const out = unzipSync(await zipBytes(await buildZipAsync(files)));
  expect(Object.keys(out).sort()).toEqual(['a/b/c.json', 'images/photo.png', 'messages.txt']);
  expect(out['messages.txt']).toEqual(files[0].data);
  expect(out['images/photo.png']).toEqual(files[1].data);
  expect(out['a/b/c.json']).toEqual(files[2].data);
});

test('a .txt is deflated but a .png is stored (ALREADY_COMPRESSED_RE routing), both round-trip', async () => {
  const payload = enc('A'.repeat(2000)); // highly compressible, so deflate is visibly smaller
  const asText = await zipBytes(await buildZipAsync([{ path: 'x.txt', data: payload }]));
  const asImage = await zipBytes(await buildZipAsync([{ path: 'x.png', data: payload }]));
  // Text is deflated (tiny) while the image extension is STORED (payload + framing).
  expect(asText.length).toBeLessThan(asImage.length);
  expect(asImage.length).toBeGreaterThanOrEqual(payload.length);
  // Routing must not corrupt content: both extract back to the original bytes.
  expect(unzipSync(asText)['x.txt']).toEqual(payload);
  expect(unzipSync(asImage)['x.png']).toEqual(payload);
});

test('createZipStream builds incrementally (add/finish) and round-trips', async () => {
  const s = createZipStream('t');
  await s.add('one.txt', enc('one'));
  await s.add('dir/two.csv', enc('a,b\n1,2'));
  const out = unzipSync(await zipBytes(await s.finish()));
  expect(out['one.txt']).toEqual(enc('one'));
  expect(out['dir/two.csv']).toEqual(enc('a,b\n1,2'));
});

test('an empty archive finishes cleanly', async () => {
  const out = unzipSync(await zipBytes(await buildZipAsync([])));
  expect(Object.keys(out)).toEqual([]);
});
