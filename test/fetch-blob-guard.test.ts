// Regression lock for the FETCH_BLOB bearer-token host guard
// (background.ts handleFetchBlobMessage): the handler attaches
// Authorization: Bearer only when isMicrosoftApiHost(url) passes.
//
// The allowed list pins every URL shape the FETCH_BLOB call sites send
// today (content.ts fetchImageAsDataUrl via transformImageUrlToProxy,
// and the two Graph photo fetchers). If a future change makes one of
// these fail the guard, images silently degrade to cookie-only fetches,
// so keep this list in sync with the call sites.
//
// Runs via `node --test scripts/` (part of `pnpm check`). Node 24
// strips the types from the .ts import natively.
import { test } from 'vitest';
import assert from 'node:assert/strict';
import { isMicrosoftApiHost } from '../src/utils/teams-urls.ts';

const UUID = '01234567-89ab-cdef-0123-456789abcdef';

// URL shapes the call sites send with a token attached today.
const allowed = [
    // AMS-direct object, rewritten by transformImageUrlToProxy()
    `https://eu-prod.asyncgw.teams.microsoft.com/v1/${UUID}/objects/0-abc-d1-xyz/views/imgo?v=1`,
    // urlp thumbnail proxy: a foreign upstream URL rides along only as an
    // encoded query param; the request host stays Microsoft-owned.
    `https://na-prod.asyncgw.teams.microsoft.com/urlp/v1/${UUID}/url/image/Thumbnail?url=${encodeURIComponent('https://evil.example/cat.png')}`,
    // Graph profile-photo fetches (author avatars + reactor photos)
    `https://graph.microsoft.com/v1.0/users/${UUID}/photo/$value`,
    // Raw AMS host (skype.com suffix), in case a caller ever skips the rewrite
    'https://us-api.asm.skype.com/v1/objects/abc/views/imgo',
    // Absolute-DNS trailing dot is tolerated by design
    'https://graph.microsoft.com./v1.0/x',
];

// Shapes a poisoned URL could take. None may carry the token.
const refused = [
    'https://evil.example/exfil',
    'https://evilmicrosoft.com/',                       // look-alike without a dot boundary
    'https://microsoft.com.evil.example/',              // MS name as attacker subdomain
    'https://graph.microsoft.com.attacker.example/',
    'http://graph.microsoft.com/v1.0/x',                // https only
    'https://graph.microsoft.com@evil.example/photo',   // userinfo trick: real host is evil.example
    'ftp://microsoft.com/x',
    'data:text/html,x',
    'not a url',
    '',
];

test('bearer allowed for every URL shape the call sites send', () => {
    for (const url of allowed) {
        assert.equal(isMicrosoftApiHost(url), true, `should allow: ${url}`);
    }
});

test('bearer refused for non-Microsoft or non-https URLs', () => {
    for (const url of refused) {
        assert.equal(isMicrosoftApiHost(url), false, `should refuse: ${url || '(empty string)'}`);
    }
});

test('bearer refused for null and undefined', () => {
    assert.equal(isMicrosoftApiHost(null), false);
    assert.equal(isMicrosoftApiHost(undefined), false);
});
