// Behavior lock for the messaging-fetch network resilience in
// src/content/api-client.ts (the "one blip loses a whole chat" fix):
//   - fetchPageWithRetry retries the timeout/network class ONCE (was zero),
//     bounded (no infinite loop), then throws a NetworkError so downstream
//     partial-network flagging still fires.
//   - fetchAllMessages keeps the pages already fetched when the network drops
//     mid-pagination (a partial chat) instead of discarding the whole chat.
//
// The retry backoff (NETWORK_RETRY_BACKOFF_MS = 2000) is driven with fake timers
// so these run in milliseconds instead of blocking on real sleeps. The internal
// 30s AbortSignal.timeout never fires: the fetch mock ignores the signal and
// resolves/throws synchronously, and we only advance past the 2s backoff.
import { test, vi } from 'vitest';
import assert from 'node:assert/strict';
import { fetchPageWithRetry, fetchAllMessages } from '../src/content/api-client.ts';

const MS_URL = 'https://apac.ng.msg.teams.microsoft.com/v1/x';

// Minimal Response-like stub for a 200 JSON page.
function jsonResponse(body) {
  return {
    status: 200,
    ok: true,
    headers: { get: () => null },
    clone: () => ({ text: async () => '' }),
    json: async () => body,
  };
}

test('fetchPageWithRetry retries once on a transient network throw, then succeeds', async () => {
  vi.useFakeTimers();
  const orig = global.fetch;
  let calls = 0;
  global.fetch = async () => {
    calls++;
    if (calls === 1) throw new TypeError('NetworkError when attempting to fetch resource.');
    return jsonResponse({ messages: [{ id: 'm1' }] });
  };
  try {
    const p = fetchPageWithRetry(MS_URL, 'tok');
    await vi.advanceTimersByTimeAsync(2500); // flush the one retry backoff (2000ms)
    const data = await p;
    assert.equal(calls, 2, 'fetched twice — one retry after the blip');
    assert.equal(data.messages.length, 1);
  } finally { global.fetch = orig; vi.useRealTimers(); }
});

test('fetchPageWithRetry gives up after the bounded network retry (no infinite loop)', async () => {
  vi.useFakeTimers();
  const orig = global.fetch;
  let calls = 0;
  global.fetch = async () => { calls++; throw new TypeError('Failed to fetch'); };
  try {
    const p = fetchPageWithRetry(MS_URL, 'tok');
    const settled = p.then(() => null, (e) => e); // capture the rejection so it's never "unhandled"
    await vi.advanceTimersByTimeAsync(2500);
    const err = await settled;
    assert.ok(err && /NetworkError|Failed to fetch/i.test(String(err.message)), 'threw a network error');
    assert.equal(calls, 2, 'initial attempt + exactly one retry, then throw');
  } finally { global.fetch = orig; vi.useRealTimers(); }
});

test('fetchAllMessages keeps the fetched pages when the network drops mid-pagination', async () => {
  vi.useFakeTimers();
  const orig = global.fetch;
  let calls = 0;
  global.fetch = async () => {
    calls++;
    if (calls === 1) {
      return jsonResponse({
        messages: [{ id: 'a' }, { id: 'b' }],
        _metadata: { backwardLink: 'https://apac.ng.msg.teams.microsoft.com/v1/next' },
      });
    }
    throw new TypeError('Failed to fetch'); // page 2 dies for good (initial + retry)
  };
  try {
    const config = { chatServiceUrl: 'https://apac.ng.msg.teams.microsoft.com', ic3Token: 'tok' };
    const p = fetchAllMessages(config, 'conv1');
    await vi.advanceTimersByTimeAsync(3000); // inter-page delay (150ms) + retry backoff (2000ms)
    const msgs = await p;
    assert.equal(msgs.length, 2, 'the first page is kept, not discarded');
  } finally { global.fetch = orig; vi.useRealTimers(); }
});
