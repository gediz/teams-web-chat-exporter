import { test, expect } from 'vitest';
import { fakeBrowser } from 'wxt/testing';

// Proves the extension-API side of the harness end to end: WxtVitest stubs the
// chrome/browser globals with fakeBrowser, storage.local is a real in-memory
// store, and setup.ts's beforeEach reset isolates specs. This is the seam the
// background/RF4 tests will build on.
test('fakeBrowser storage.local round-trips', async () => {
  await fakeBrowser.storage.local.set({ hello: 'world' });
  expect(await fakeBrowser.storage.local.get('hello')).toEqual({ hello: 'world' });
});

test('the chrome global is wired to fakeBrowser (and reset between tests)', async () => {
  // A fresh reset means nothing set by the previous test survives here.
  expect(await chrome.storage.local.get('hello')).toEqual({});
  await chrome.storage.local.set({ n: 1 });
  expect(await chrome.storage.local.get('n')).toEqual({ n: 1 });
});
