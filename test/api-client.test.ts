// Behavior lock for the messaging-fetch network resilience in
// src/content/api-client.ts (the "one blip loses a whole chat" fix):
//   - fetchPageWithRetry retries the timeout/network class ONCE (was zero),
//     bounded (no infinite loop), then throws a NetworkError so downstream
//     partial-network flagging still fires.
//   - fetchAllMessages keeps the pages already fetched when the network drops
//     mid-pagination (a partial chat) instead of discarding the whole chat.
import { test } from 'vitest';
import assert from 'node:assert/strict';
import { fetchPageWithRetry, fetchAllMessages } from '../src/content/api-client.ts';

// The token-refresh path (getIc3Token) reads localStorage; in node there is
// none, so stub an empty one — it returns null and fetchAllMessages falls back
// to config.ic3Token, which is all these transport tests exercise.
globalThis.localStorage = globalThis.localStorage || { getItem: () => null, setItem() {}, removeItem() {} };

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
  const orig = global.fetch;
  let calls = 0;
  global.fetch = async () => {
    calls++;
    if (calls === 1) throw new TypeError('NetworkError when attempting to fetch resource.');
    return jsonResponse({ messages: [{ id: 'm1' }] });
  };
  try {
    const data = await fetchPageWithRetry(MS_URL, 'tok');
    assert.equal(calls, 2, 'fetched twice — one retry after the blip');
    assert.equal(data.messages.length, 1);
  } finally { global.fetch = orig; }
});

test('fetchPageWithRetry gives up after the bounded network retry (no infinite loop)', async () => {
  const orig = global.fetch;
  let calls = 0;
  global.fetch = async () => { calls++; throw new TypeError('Failed to fetch'); };
  try {
    await assert.rejects(
      fetchPageWithRetry(MS_URL, 'tok'),
      (e) => /NetworkError|Failed to fetch/i.test(String(e && e.message)),
    );
    assert.equal(calls, 2, 'initial attempt + exactly one retry, then throw');
  } finally { global.fetch = orig; }
});

test('fetchAllMessages keeps the fetched pages when the network drops mid-pagination', async () => {
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
    const msgs = await fetchAllMessages(config, 'conv1');
    assert.equal(msgs.length, 2, 'the first page is kept, not discarded');
  } finally { global.fetch = orig; }
});
