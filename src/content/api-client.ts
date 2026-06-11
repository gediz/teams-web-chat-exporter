/**
 * Teams API Client
 *
 * Fetches messages directly from the Teams chat service API.
 * Uses MSAL tokens from localStorage and service discovery via authz endpoint.
 * Falls back gracefully — any failure should trigger DOM scroll fallback.
 */

/* eslint-disable no-console */

import type { ConversationSummary, FolderSummary } from '../types/shared';

// ── Types ──────────────────────────────────────────────────────────────

export type TeamsApiConfig = {
  chatServiceUrl: string;
  userRegion: string;
  ic3Token: string;
};

export type TeamsApiMessage = {
  sequenceId?: number;
  conversationid?: string;
  conversationLink?: string;
  contenttype?: string;
  type?: string;
  messagetype?: string;
  id?: string;
  clientmessageid?: string;
  version?: string;
  content?: string;
  from?: string;
  imdisplayname?: string;
  prioritizeImDisplayName?: boolean;
  composetime?: string;
  originalarrivaltime?: string;
  amsreferences?: string;
  fromDisplayNameInToken?: string;
  fromGivenNameInToken?: string;
  fromFamilyNameInToken?: string;
  properties?: Record<string, unknown>;
};

type AuthzResponse = {
  region?: string;
  userRegion?: string;
  regionGtms?: {
    chatService?: string;
    [key: string]: unknown;
  };
  // Teams Free shape (teams.live.com authz/consumer response): the
  // messaging token sits in a nested object at `skypeToken.skypetoken`
  // (note the lowercase inner field). The same response also carries
  // `expiresIn`, the user's `skypeid` / `signinname`, and a flag.
  skypeToken?: {
    skypetoken?: string;
    expiresIn?: number;
    skypeid?: string;
    signinname?: string;
    isBusinessTenant?: boolean;
  };
  // Older / hypothetical Work-School shape, kept for forward
  // compatibility — never observed in current responses but cheap to
  // probe before falling back to IC3 cache.
  tokens?: {
    skypeToken?: string;
    expiresIn?: number;
    tokenType?: string;
  };
  [key: string]: unknown;
};

type MessagesResponse = {
  messages?: TeamsApiMessage[];
  tenantId?: string;
  _metadata?: {
    backwardLink?: string;
    syncState?: string;
    lastCompleteSegmentStartTime?: number;
    lastCompleteSegmentEndTime?: number;
  };
  errorCode?: number;
  message?: string;
};

export type FetchProgress = {
  phase: 'discover' | 'fetch';
  page: number;
  messagesSoFar: number;
};

// ── Token Extraction ───────────────────────────────────────────────────

// MSAL Browser v4+ encrypts cache entries by default. Each entry is stored as
// { id, nonce, data, lastUpdatedAt }. The session cookie msal.cache.encryption
// holds the base key; per-entry AES-GCM keys are derived via HKDF-SHA256 with
// salt=nonce and info=clientId. The IV is always 12 zero bytes (per-entry
// uniqueness comes from the unique HKDF salt).
// See: https://github.com/AzureAD/microsoft-authentication-library-for-js
//      /blob/dev/lib/msal-browser/src/crypto/BrowserCrypto.ts

type EncryptedEntry = { id?: string; nonce?: string; data?: string };
type DecryptedAccessToken = { secret?: string; expiresOn?: string };

const b64uToBytes = (b: string): Uint8Array<ArrayBuffer> => {
  const s = b.replace(/-/g, '+').replace(/_/g, '/');
  const p = s + '='.repeat((4 - s.length % 4) % 4);
  const bin = atob(p);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
};

let cachedBaseKey: { raw: string; key: CryptoKey } | null = null;

async function getMsalBaseKey(): Promise<CryptoKey | null> {
  const m = document.cookie.match(/msal\.cache\.encryption=([^;]+)/);
  if (!m) return null;
  let parsed: { key?: string };
  try { parsed = JSON.parse(decodeURIComponent(m[1])); } catch { return null; }
  if (!parsed.key) return null;
  if (cachedBaseKey?.raw === parsed.key) return cachedBaseKey.key;
  try {
    const key = await crypto.subtle.importKey('raw', b64uToBytes(parsed.key), 'HKDF', false, ['deriveKey']);
    cachedBaseKey = { raw: parsed.key, key };
    return key;
  } catch { return null; }
}

// MSAL key shape: msal.2|<accountId>|<env>|accesstoken|<clientId>|<realm>|<target>|
// Context for HKDF info is the clientId for accesstoken entries, empty otherwise.
function getEncryptionContext(lsKey: string): string {
  const parts = lsKey.split('|');
  return parts[3] === 'accesstoken' && parts[4] ? parts[4] : '';
}

async function decryptEntry(baseKey: CryptoKey, entry: EncryptedEntry, context: string): Promise<DecryptedAccessToken | null> {
  if (!entry.nonce || !entry.data) return null;
  try {
    const derivedKey = await crypto.subtle.deriveKey(
      { name: 'HKDF', hash: 'SHA-256', salt: b64uToBytes(entry.nonce), info: new TextEncoder().encode(context) },
      baseKey,
      { name: 'AES-GCM', length: 256 },
      false,
      ['decrypt'],
    );
    const decrypted = await crypto.subtle.decrypt(
      { name: 'AES-GCM', iv: new Uint8Array(12) },
      derivedKey,
      b64uToBytes(entry.data),
    );
    return JSON.parse(new TextDecoder().decode(decrypted));
  } catch { return null; }
}

/** Find a valid (non-expired) MSAL access token matching the scope pattern. */
async function findValidToken(scopePattern: string): Promise<string | null> {
  const now = Math.floor(Date.now() / 1000);
  let baseKey: CryptoKey | null = null;
  for (let i = 0; i < localStorage.length; i++) {
    const key = localStorage.key(i);
    if (!key || !key.includes('accesstoken') || !key.includes(scopePattern)) continue;
    try {
      const entry = JSON.parse(localStorage.getItem(key) || '');
      // Plain entry (older MSAL or KMSI sessions)
      if (entry.secret && Number(entry.expiresOn) > now) return entry.secret;
      // Encrypted entry (MSAL Browser v4+ default)
      if (entry.data && entry.nonce) {
        if (!baseKey) baseKey = await getMsalBaseKey();
        if (!baseKey) continue;
        const decrypted = await decryptEntry(baseKey, entry, getEncryptionContext(key));
        if (decrypted?.secret && Number(decrypted.expiresOn) > now) return decrypted.secret;
      }
    } catch { /* skip malformed entries */ }
  }
  return null;
}

/** Get a valid ic3 token for the message API. Re-reads localStorage each call. */
export async function getIc3Token(): Promise<string | null> {
  // Standard commercial: ic3.teams.office.com
  // GCC High: ic3.teams.office365.us or chatsvcagg
  // Teams Free: mtsvc.fl.teams.microsoft.com (different consumer service)
  return await findValidToken('ic3.teams.office.com')
    || await findValidToken('ic3.teams.office365.us')
    || await findValidToken('chatsvcagg')
    || await findValidToken('mtsvc.fl.teams.microsoft.com');
}

/** Get a valid Skype API token for the authz discovery endpoint. */
export async function getSkypeToken(): Promise<string | null> {
  // Pattern broadened from `api.spaces.skype` to `spaces.skype` so it
  // matches both work/school (`api.spaces.skype.com`) and Teams Free
  // (`api.fl.spaces.skype.com` — the consumer service has a `.fl.`
  // infix that breaks the older substring check).
  return findValidToken('spaces.skype');
}

/** Get a valid Graph API token for user resolution. */
export async function getGraphToken(): Promise<string | null> {
  // Standard: graph.microsoft.com, GCC High: graph.microsoft.us
  return await findValidToken('graph.microsoft.com')
    || await findValidToken('graph.microsoft.us');
}

/**
 * Extract the user's UUID from an MSAL access token cache key.
 * Cache key format: msal.2|<userUuid>.<tenantId>|<env>|accesstoken|...
 *
 * In multi-account sessions MSAL may store tokens for multiple accounts; prefer
 * entries for Teams-related scopes so we identify the account actually using
 * the Teams page, not some other signed-in tenant.
 */
export function getCurrentUserUuid(): string | null {
  const UUID_RE = /^([a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12})/i;
  const teamsScopes = ['ic3.teams.office.com', 'spaces.skype', 'graph.microsoft.com', 'mtsvc.fl.teams.microsoft.com'];
  const findWithFilter = (predicate: (k: string) => boolean): string | null => {
    for (let i = 0; i < localStorage.length; i++) {
      const key = localStorage.key(i);
      if (!key || !key.includes('accesstoken') || !predicate(key)) continue;
      const m = (key.split('|')[1] || '').match(UUID_RE);
      if (m) return m[1];
    }
    return null;
  };
  // Prefer a Teams-scope access token; fall back to any access token.
  return findWithFilter(k => teamsScopes.some(s => k.includes(s)))
      ?? findWithFilter(() => true);
}

// ── User Resolution via Graph API ──────────────────────────────────────

/**
 * Extract a UUID from an MRI string.
 * Handles: "8:orgid:{uuid}", "gid:{uuid}", full URLs containing a UUID.
 */
function extractUuid(mri: string): string | null {
  const match = mri.match(/([a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12})/i);
  return match ? match[1] : null;
}

/**
 * Resolve a set of MRIs to display names via Microsoft Graph API.
 * Accepts MRIs in any format (8:orgid:{uuid}, gid:{uuid}, full URL).
 * Returns a map of original MRI → displayName.
 */
// Module-scope memo for Graph user resolutions. Prevents re-querying
// the same UUIDs every popup-open (40+ MRIs × N opens × ~1 s = real
// network waste, plus 404 spam in the console for federated users
// who aren't in this tenant). Negative results are remembered too so
// we don't keep firing the same 404s. Cache lifetime is the content-
// script process — a Teams tab reload flushes it.
const _graphUserCache = new Map<string, string | null>(); // uuid → name | null (404)

export async function resolveUserNames(
  mris: string[],
  graphToken: string,
): Promise<Map<string, string>> {
  const result = new Map<string, string>();
  if (!mris.length || !graphToken) return result;

  // Deduplicate UUIDs and map back to original MRIs
  const uuidToMris = new Map<string, string[]>();
  for (const mri of mris) {
    const uuid = extractUuid(mri);
    if (!uuid) continue;
    const existing = uuidToMris.get(uuid) || [];
    existing.push(mri);
    uuidToMris.set(uuid, existing);
  }

  // Drain cached results first; keep only uncached UUIDs for network.
  const toFetch = new Map<string, string[]>();
  for (const [uuid, originalMris] of uuidToMris) {
    if (_graphUserCache.has(uuid)) {
      const cached = _graphUserCache.get(uuid);
      if (cached) {
        for (const mri of originalMris) result.set(mri, cached);
        result.set(`8:orgid:${uuid}`, cached);
        result.set(`gid:${uuid}`, cached);
      }
    } else {
      toFetch.set(uuid, originalMris);
    }
  }
  if (toFetch.size === 0) return result;

  // Resolve each UUID via Graph. Sequential resolution made the picker
  // wait 100s on accounts with lots of 1:1 chats — IDB-source surfaces
  // ~3× more conversations than the API list did, and most are 1:1s
  // needing a name. Bounded-concurrency parallelism keeps wall-time
  // down to ceil(N/CONCURRENCY) * RTT instead of N * RTT.
  const CONCURRENCY = 10;
  const errors: string[] = [];
  const entries = [...toFetch.entries()];
  for (let i = 0; i < entries.length; i += CONCURRENCY) {
    const batch = entries.slice(i, i + CONCURRENCY);
    await Promise.all(batch.map(async ([uuid, originalMris]) => {
      try {
        const resp = await fetch(
          `https://graph.microsoft.com/v1.0/users/${uuid}?$select=displayName`,
          {
            headers: { 'Authorization': `Bearer ${graphToken}` },
            signal: AbortSignal.timeout(5_000),
          },
        );
        if (!resp.ok) {
          // 404 = federated user not in this tenant. Remember the
          // miss so we don't refetch on the next popup open.
          if (resp.status === 404) _graphUserCache.set(uuid, null);
          if (errors.length < 3) errors.push(`HTTP ${resp.status} for ${uuid.slice(0,8)}`);
          return;
        }
        const data = await resp.json();
        const name = data.displayName;
        if (name) {
          _graphUserCache.set(uuid, name);
          for (const mri of originalMris) result.set(mri, name);
          result.set(`8:orgid:${uuid}`, name);
          result.set(`gid:${uuid}`, name);
        } else {
          _graphUserCache.set(uuid, null);
        }
      } catch (e) {
        if (errors.length < 3) errors.push(String(e));
      }
    }));
  }
  if (errors.length > 0 && result.size === 0) {
    // Logged at info, not warn: a 404 here just means the user is
    // federated/guest and not in this tenant's directory (expected), and
    // names fall back to other sources. Keeping it off the warn/error
    // channel stops it surfacing in the extension's chrome://extensions
    // Errors panel, while still logging to the page console for debugging.
    console.info(`[API] Graph name resolution: ${uuidToMris.size} user(s) unresolved (often federated/guest 404s, expected) — ${errors.join('; ')}`);
  }

  return result;
}

// ── SharePoint file fetch ────────────────────────────────────────────
//
// Files attached via paperclip / drag-drop in Teams (as opposed to
// pasted-as-image) end up on the user's SharePoint OneDrive under
// "Microsoft Teams Chat Files/". The chat message references the file
// by its SharePoint URL. To embed the file's bytes in an export we
// fetch the URL directly with the user's session cookies — they're
// already signed into SharePoint via Teams' Microsoft 365 SSO.
//
// Requires *.sharepoint.com / *.sharepoint.us in host_permissions
// (declared in src/utils/teams-urls.ts). Going through Microsoft
// Graph's /shares endpoint instead doesn't help: every Graph file-
// content endpoint returns a 302 redirect to a SharePoint host, and
// the redirect-follow fails without the same host_permission.

export type SharePointFetchResult =
  | { ok: true; bytes: Uint8Array; mime: string }
  | { ok: false; status: number; statusText: string; reason: string };

// Sniff whether a fetched payload is actually an image. SharePoint/OneDrive
// serve a "request access" / sign-in HTML page at HTTP 200 when the user lacks
// access; without this check that HTML gets saved with an image extension and
// renders as a broken <img>. Magic bytes are authoritative — Content-Type is
// unreliable (SharePoint often serves images as application/octet-stream). A
// genuine SVG (text, no binary magic) is allowed; any other non-raster body
// (the auth/HTML pages) is rejected. Only ever called on image-file fetches.
function isImagePayload(bytes: Uint8Array): boolean {
  const b = bytes;
  const has = (sig: number[], off = 0) => sig.every((v, i) => b[off + i] === v);
  if (has([0xff, 0xd8, 0xff])) return true;                                          // JPEG
  if (has([0x89, 0x50, 0x4e, 0x47])) return true;                                    // PNG
  if (has([0x47, 0x49, 0x46, 0x38])) return true;                                    // GIF87a/89a
  if (has([0x52, 0x49, 0x46, 0x46]) && has([0x57, 0x45, 0x42, 0x50], 8)) return true; // RIFF....WEBP
  if (has([0x42, 0x4d])) return true;                                                // BMP
  if (has([0x49, 0x49, 0x2a, 0x00]) || has([0x4d, 0x4d, 0x00, 0x2a])) return true;   // TIFF
  if (has([0x66, 0x74, 0x79, 0x70], 4)) {                                            // ....ftyp (HEIC/HEIF/AVIF)
    const brand = String.fromCharCode(b[8] || 0, b[9] || 0, b[10] || 0, b[11] || 0);
    if (/heic|heix|hevc|heim|heis|mif1|msf1|avif/i.test(brand)) return true;
  }
  // Non-raster: accept only a real SVG document; reject HTML/XML auth pages.
  const head = new TextDecoder('utf-8', { fatal: false }).decode(b.subarray(0, 512)).trimStart().toLowerCase();
  return (head.startsWith('<?xml') || head.startsWith('<svg')) && head.includes('<svg');
}

/**
 * Fetch a SharePoint-hosted file's bytes using the user's session
 * cookies. Returns a result shape (never throws) so the caller can
 * log + skip cleanly. Direct content-script fetch — works because
 * `*.sharepoint.com/.us/.cn/-mil.us` are in host_permissions and the
 * work/school SharePoint hosts honour CORS for the Teams origin.
 */
export async function fetchSharePointFile(
  url: string,
  signal?: AbortSignal,
): Promise<SharePointFetchResult> {
  try {
    const resp = await fetch(url, {
      credentials: 'include',
      signal: signal || AbortSignal.timeout(20_000),
      redirect: 'follow',
    });
    if (!resp.ok) {
      return { ok: false, status: resp.status, statusText: resp.statusText, reason: `SharePoint ${resp.status}` };
    }
    const buf = await resp.arrayBuffer();
    const bytes = new Uint8Array(buf);
    if (!isImagePayload(bytes)) {
      // 200 OK but the body isn't an image — SharePoint/OneDrive returned a
      // "request access" / sign-in HTML page. Saving it would be a broken
      // image, so report a failed fetch; the caller renders a link placeholder.
      return { ok: false, status: resp.status, statusText: 'non-image body', reason: 'SharePoint returned a non-image response (access/sign-in page)' };
    }
    const mime = resp.headers.get('content-type')?.split(';')[0]?.trim() || 'application/octet-stream';
    return { ok: true, bytes, mime };
  } catch (e: any) {
    return { ok: false, status: 0, statusText: '', reason: e?.message || String(e) };
  }
}

/**
 * Collect all unresolved MRIs from a batch of API messages.
 * Returns MRIs that appear as senders (gid:), reactors, or forward originators
 * but have no display name in the message data.
 */
export function collectUnresolvedMris(messages: TeamsApiMessage[]): string[] {
  // Build a set of MRIs that already have names
  const knownMris = new Set<string>();
  for (const msg of messages) {
    if (msg.from && msg.imdisplayname) {
      knownMris.add(msg.from);
      const uuid = extractUuid(msg.from);
      if (uuid) {
        knownMris.add(`8:orgid:${uuid}`);
        knownMris.add(`gid:${uuid}`);
      }
    }
  }

  // Collect MRIs that need resolution
  const unresolved = new Set<string>();
  for (const msg of messages) {
    // Forwarded message senders (gid: with no name)
    if (msg.from?.includes('gid:') && !msg.imdisplayname && !knownMris.has(msg.from)) {
      unresolved.add(msg.from);
    }

    // Original sender in forwarded messages
    const ctx = (msg.properties as Record<string, unknown>)?.originalMessageContext;
    if (ctx) {
      try {
        const parsed = typeof ctx === 'string' ? JSON.parse(ctx) : ctx;
        const sender = (parsed as Record<string, string>).originalSender;
        if (sender && !knownMris.has(sender)) unresolved.add(sender);
      } catch { /* ignore */ }
    }

    // Unresolved reactor MRIs
    const emotions = (msg.properties as Record<string, unknown>)?.emotions;
    if (emotions) {
      try {
        const parsed = typeof emotions === 'string' ? JSON.parse(emotions) : emotions;
        if (Array.isArray(parsed)) {
          for (const e of parsed) {
            for (const u of (e.users || [])) {
              if (u.mri && !knownMris.has(u.mri)) unresolved.add(u.mri);
            }
          }
        }
      } catch { /* ignore */ }
    }
  }

  return [...unresolved];
}

// ── Service Discovery ──────────────────────────────────────────────────

// Establish Teams Free's `skypetoken_asm` session cookie on
// `.asm.skype.com` by exchanging the short Skype JWT at
//   POST https://us-api.asm.skype.com/v1/skypetokenauth
//   body: skypetoken=<JWT>
// 204 No Content. Subsequent direct GETs on `*.asm.skype.com/v1/objects/...`
// (inline images embedded in messages) need this cookie — without it the
// AMS host returns 401. Memoised per-JWT for the SW lifetime so a
// multi-chat bundle only pays the round-trip once.
//
// Not used by the chat-service messaging endpoint (that one auths via
// the `authentication: skypetoken=<JWT>` header alone). Kept narrowly
// scoped to image fetching where direct AMS calls are unavoidable —
// Teams Free has no image-proxy host equivalent to Work/School's
// `*.asyncgw.teams.microsoft.com` that we could route through instead.
const _skypeTokenAuthDone = new Set<string>();

export async function ensureSkypeTokenCookies(jwt: string, signal?: AbortSignal): Promise<boolean> {
  if (!jwt) return false;
  if (_skypeTokenAuthDone.has(jwt)) return true;
  try {
    const resp = await fetch('https://us-api.asm.skype.com/v1/skypetokenauth', {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: `skypetoken=${encodeURIComponent(jwt)}`,
      credentials: 'include',
      signal: signal || AbortSignal.timeout(15_000),
    });
    if (resp.ok || resp.status === 204) {
      _skypeTokenAuthDone.add(jwt);
      return true;
    }
    console.warn(`[API:tfl] skypetokenauth: ${resp.status} ${resp.statusText}`);
    return false;
  } catch (e: any) {
    console.warn(`[API:tfl] skypetokenauth threw: ${e?.message || String(e)}`);
    return false;
  }
}

// Teams Free encrypts cached auth state in localStorage with a 32-byte
// AES-256-CBC key it stores alongside, in the clear, at:
//   tmp.auth.v1.GLOBAL.ExportedEncryptionKey.ExportedEncryptionKey
// The key's own JSON-wrapper `comment` field calls it a "temporary
// location" — Teams effectively designed for first-party clients to
// be able to decrypt the cache. We use this to read the short-lived
// Skype JWT (`Discover.SKYPE-TOKEN`) the chat-service proxy expects.
let _cachedTeamsFreeAesKey: { raw: string; key: CryptoKey } | null = null;

async function getTeamsFreeAesKey(): Promise<CryptoKey | null> {
  for (let i = 0; i < localStorage.length; i++) {
    const k = localStorage.key(i);
    if (!k || !k.endsWith('.ExportedEncryptionKey.ExportedEncryptionKey')) continue;
    try {
      const wrapper = JSON.parse(localStorage.getItem(k) || '');
      const exported: unknown = wrapper?.item?.exportedKey;
      if (typeof exported !== 'string') continue;
      if (_cachedTeamsFreeAesKey?.raw === exported) return _cachedTeamsFreeAesKey.key;
      const key = await crypto.subtle.importKey('raw', b64uToBytes(exported), { name: 'AES-CBC' }, false, ['decrypt']);
      _cachedTeamsFreeAesKey = { raw: exported, key };
      return key;
    } catch { /* skip malformed */ }
  }
  return null;
}

type TeamsFreeEncryptedItem = { encryptedToken?: unknown; iv?: unknown };
async function decryptTeamsFreeItem(item: TeamsFreeEncryptedItem): Promise<string | null> {
  if (typeof item?.encryptedToken !== 'string' || typeof item?.iv !== 'string') return null;
  const key = await getTeamsFreeAesKey();
  if (!key) return null;
  try {
    const pt = await crypto.subtle.decrypt(
      { name: 'AES-CBC', iv: b64uToBytes(item.iv) },
      key,
      b64uToBytes(item.encryptedToken),
    );
    return new TextDecoder().decode(pt);
  } catch { return null; }
}

/**
 * Read Teams Free's cached chat-service auth: the short-lived Skype
 * JWT and the chat-service URL, decrypted from
 *   tmp.auth.v1.<cid>.Discover.SKYPE-TOKEN
 *
 * Returns null when the cache is missing, expired, or undecryptable —
 * caller falls back to the live authz POST in that case.
 *
 * The cache record carries an unencrypted `regionGtms.chatService` and
 * `expiration` (unix ms) so we can pick the right messaging endpoint
 * and skip a stale record without paying the cost of a decrypt that
 * would yield an unusable token anyway.
 */
async function getTeamsFreeChatAuth(): Promise<{ token: string; chatServiceUrl?: string; userRegion?: string } | null> {
  const now = Date.now();
  for (let i = 0; i < localStorage.length; i++) {
    const k = localStorage.key(i);
    if (!k || !k.endsWith('.Discover.SKYPE-TOKEN')) continue;
    try {
      const wrapper = JSON.parse(localStorage.getItem(k) || '');
      const item = wrapper?.item;
      if (!item || typeof item !== 'object') continue;
      if (Number(item.expiration) <= now) continue;
      const token = await decryptTeamsFreeItem(item);
      if (!token) continue;
      let chatServiceUrl: string | undefined;
      const cs = (item.regionGtms as Record<string, unknown> | undefined)?.chatService;
      if (typeof cs === 'string') chatServiceUrl = cs;
      // The cached chatService is the msgapi.teams.live.com backend; the
      // proxied teams.live.com path is what Teams' own UI uses (and what
      // works without registering a Skype endpoint first).
      if (chatServiceUrl && /msgapi\.teams\.live\.com/i.test(chatServiceUrl)) {
        chatServiceUrl = 'https://teams.live.com/api/chatsvc/consumer';
      }
      const userDetails = item.userDetails as Record<string, unknown> | undefined;
      const userRegion = typeof userDetails?.region === 'string' ? userDetails.region : '';
      return { token, chatServiceUrl, userRegion };
    } catch { /* skip */ }
  }
  return null;
}

/**
 * Read Teams Free's local profile cache (mri → displayName) from
 *   IndexedDB: Teams:profiles:react-web-client:<tenant>:<cid>:<lang>
 *     objectStore: profiles  (key path = mri)
 *
 * On Teams Free this is the only place the converter can resolve Skype
 * consumer IDs to names — Microsoft Graph doesn't know consumer
 * accounts. Gated to teams.live.com so Work/School never even reads
 * the DB (which exists with a different schema there); their existing
 * Graph-resolution path stays the sole source of names.
 *
 * Each profile carries both keys we want (`mri`, `displayName`); we
 * also store the bare-id form (e.g. `<id>` from `8:<id>`) so
 * `<part identity="<id>">` in Event/Call partlists can resolve
 * without forcing the converter to know about the `8:` prefix dance.
 *
 * Memoised at module scope with a 60-second TTL so a 14-chat bundle
 * pays one IDB roundtrip instead of 14. The TTL is short enough that
 * a long-running session picks up newly-added contacts within a
 * minute, which matches Teams' own UI refresh cadence.
 */
let _profilesCache: { map: Map<string, string>; expiresAt: number } | null = null;
let _profilesCachePromise: Promise<Map<string, string>> | null = null;

async function loadTeamsFreeProfiles(): Promise<Map<string, string>> {
  // Gate: only Teams Free has the consumer-Skype profile records this
  // function is built for. Work/School returns immediately so we don't
  // touch their (similarly named but different-schema) profiles store.
  if (typeof location === 'undefined'
      || !location.hostname.toLowerCase().includes('teams.live.com')) {
    return new Map();
  }
  const now = Date.now();
  if (_profilesCache && _profilesCache.expiresAt > now) return _profilesCache.map;
  if (_profilesCachePromise) return _profilesCachePromise;
  _profilesCachePromise = (async () => {
    const out = new Map<string, string>();
    if (!('indexedDB' in self)) return out;
    let dbs: IDBDatabaseInfo[];
    try { dbs = await indexedDB.databases(); } catch { return out; }
    const profilesDb = dbs.find(d => typeof d.name === 'string' && /^Teams:profiles:/.test(d.name));
    if (!profilesDb?.name) return out;
    let db: IDBDatabase;
    try {
      db = await new Promise<IDBDatabase>((res, rej) => {
        const req = indexedDB.open(profilesDb.name!);
        req.onsuccess = () => res(req.result);
        req.onerror = () => rej(req.error);
      });
    } catch { return out; }
    try {
      if (!db.objectStoreNames.contains('profiles')) return out;
      const records = await new Promise<Array<{ mri?: string; displayName?: string }>>((res, rej) => {
        const tx = db.transaction('profiles', 'readonly');
        const req = tx.objectStore('profiles').getAll();
        req.onsuccess = () => res(req.result || []);
        req.onerror = () => rej(req.error);
      });
      for (const p of records) {
        const mri = typeof p?.mri === 'string' ? p.mri : '';
        const name = typeof p?.displayName === 'string' ? p.displayName.trim() : '';
        if (!mri || !name) continue;
        out.set(mri, name);
        // Bare-id alias for `<part identity="X">` lookups (X has no `8:` prefix).
        if (mri.startsWith('8:')) out.set(mri.slice(2), name);
      }
    } catch { /* swallow */ } finally { db.close(); }
    return out;
  })();
  try {
    const map = await _profilesCachePromise;
    _profilesCache = { map, expiresAt: Date.now() + 60_000 };
    return map;
  } finally {
    _profilesCachePromise = null;
  }
}

/**
 * Read a Teams Free pre-cached `tmp.auth.v1.*.Discover.<key>` entry.
 * Teams Free (teams.live.com) caches its discovered service URLs in
 * localStorage entries shaped like:
 *
 *   tmp.auth.v1.<consumer-tenant-id>.Discover.SKYPE-AUTHZ-URL
 *   tmp.auth.v1.<consumer-tenant-id>.Discover.SKYPE-TOKEN
 *
 * The value is `{ hitCount, item, shouldRefresh }`. `item` is the
 * cached value (a URL string for SKYPE-AUTHZ-URL; the discovery
 * response object for richer entries). Returns null if not present.
 */
function readTmpAuthDiscover(suffix: string): unknown | null {
  for (let i = 0; i < localStorage.length; i++) {
    const key = localStorage.key(i);
    if (!key) continue;
    if (!key.startsWith('tmp.auth.v1.')) continue;
    if (!key.endsWith(`.Discover.${suffix}`)) continue;
    try {
      const entry = JSON.parse(localStorage.getItem(key) || '');
      return entry?.item ?? null;
    } catch { /* not JSON; skip */ }
  }
  return null;
}

/**
 * Detect the authz service base URL from the current Teams domain.
 *  - GCC High         : teams.microsoft.us → authsvc.teams.microsoft.us
 *  - Teams Free       : teams.live.com → URL is pre-cached in
 *                       tmp.auth.v1.*.Discover.SKYPE-AUTHZ-URL because
 *                       the consumer Skype service runs on a different
 *                       authz host than the work/school endpoint.
 *  - All others       : authsvc.teams.microsoft.com
 *
 * Returns null when the URL can't be determined (Teams Free with no
 * cached entry yet — the user hasn't completed Teams' own discovery).
 */
function getAuthzUrl(): string | null {
  const host = location.hostname.toLowerCase();
  if (host.includes('teams.microsoft.us')) {
    return 'https://authsvc.teams.microsoft.us/v1.0/authz';
  }
  if (host.includes('teams.live.com')) {
    const cached = readTmpAuthDiscover('SKYPE-AUTHZ-URL');
    return typeof cached === 'string' && cached ? cached : null;
  }
  return 'https://authsvc.teams.microsoft.com/v1.0/authz';
}

/**
 * Discover the chat service URL and region via the authz endpoint.
 * Returns the chat-service URL plus, when the response includes one,
 * a messaging token that authenticates against that chat service.
 *
 * Work/School Teams uses the IC3 token (looked up separately by the
 * caller) for messaging. Teams Free's consumer authz endpoint returns
 * its messaging token in the same response (`data.tokens.skypeToken`)
 * and the IC3 token is irrelevant — there's no IC3 service in Teams
 * Free.
 */
export async function discover(skypeToken: string): Promise<{ chatServiceUrl: string; userRegion: string; messagingToken?: string }> {
  const authzUrl = getAuthzUrl();
  if (!authzUrl) {
    throw new Error('authz failed: no authz URL available (Teams Free needs Teams app to complete its own discovery first)');
  }
  console.log(`[API] authz endpoint: ${authzUrl}`);
  const resp = await fetch(authzUrl, {
    method: 'POST',
    headers: {
      'Authorization': `Bearer ${skypeToken}`,
      'Content-Type': 'application/json',
    },
    body: '{}',
    signal: AbortSignal.timeout(10_000),
  });

  if (!resp.ok) {
    throw new Error(`authz failed: ${resp.status} ${resp.statusText}`);
  }

  const data: AuthzResponse = await resp.json();

  let chatServiceUrl = data.regionGtms?.chatService;
  const userRegion = data.userRegion || data.region || '';
  // Token-extraction order:
  //   1. Teams Free shape: data.skypeToken.skypetoken (object → string)
  //   2. Hypothetical Work-School shape: data.tokens.skypeToken
  // For Work-School the response usually doesn't carry either; the
  // caller falls back to looking up the IC3 MSAL token from localStorage.
  const messagingToken = data.skypeToken?.skypetoken || data.tokens?.skypeToken;

  // Teams Free quirk: discover() returns chatServiceUrl as
  //   https://msgapi.teams.live.com
  // — the *backend* origin. Direct calls to it require the Skype
  // registration-token flow (POST /endpoints, store Set-RegistrationToken,
  // …). Teams' own UI sidesteps that by hitting a *proxied* path on
  // the same teams.live.com origin:
  //   https://teams.live.com/api/chatsvc/consumer/v1/users/ME/...
  // The proxy bridges the auth (we just send `authentication: <token>`,
  // no registration step). Confirmed against a HAR capture of Teams'
  // own client. Substitute the URL here so fetchAllMessages naturally
  // emits the proxied form.
  if (location.hostname.toLowerCase().includes('teams.live.com')
      && chatServiceUrl
      && /msgapi\.teams\.live\.com/i.test(chatServiceUrl)) {
    chatServiceUrl = 'https://teams.live.com/api/chatsvc/consumer';
  }

  if (!chatServiceUrl || typeof chatServiceUrl !== 'string') {
    throw new Error('authz response missing chatService URL');
  }

  return { chatServiceUrl, userRegion, messagingToken };
}

// ── Conversation List ──────────────────────────────────────────────────
//
// The picker shows a list of every conversation the user can export.
// Raw conversation records from the chat service are lean — good enough
// for IDs and topic strings on named chats, but missing the context a
// user needs to disambiguate a cluttered list:
//
//   - 1:1 chats have no `topic`; the counterparty's name lives behind
//     a Graph lookup (their MRI is encoded in the conversation id).
//   - Unnamed group chats ("(no topic)") need a roster fetch to turn
//     into "A, B, C, +N" style labels.
//   - Team channels carry only `groupId` — the parent team's name
//     comes from Graph's /me/joinedTeams.
//
// Strategy: one conversations page, classify, then fire the three
// enrichment sources in parallel, batch all required MRIs into one
// Graph /users call. Cold load for a ~100-chat account: ~2–3 s,
// dominated by the Graph calls; warm (joinedTeams cached) under 1 s.

type RawConversation = {
  id?: string;
  threadProperties?: {
    topic?: string;
    productThreadType?: string;
    // Present on TeamsStandardChannel. The UUID is the parent team's
    // group id; match against Graph's /me/joinedTeams to get the name.
    groupId?: string;
  };
  lastMessage?: {
    imdisplayname?: string;
    fromDisplayNameInToken?: string;
    composetime?: string;
    originalarrivaltime?: string;
    from?: string;
  };
};

// Block-listed productThreadTypes — system-virtual threads the picker
// shouldn't expose (activity streams, drafts, call logs, etc.). Anything
// else is a real conversation the user might want to export, so we keep
// the classifier permissive: unknown types fall back to 'chat' rather
// than silently disappearing from the picker when Teams ships a new
// thread category we haven't seen.
const NON_EXPORTABLE_PREFIXES = ['StreamOf', 'Activity', 'Notification'];
const NON_EXPORTABLE_EXACT = new Set(['Saved', 'Drafts', 'CallLog', 'CallLogs', 'Mentions']);

function classifyConversationKind(ptt: string | undefined, convId?: string): ConversationSummary['kind'] | null {
  if (ptt) {
    if (NON_EXPORTABLE_EXACT.has(ptt)) return null;
    if (NON_EXPORTABLE_PREFIXES.some(p => ptt.startsWith(p))) return null;
  }
  switch (ptt) {
    case 'OneToOneChat': return 'chat';
    case 'Chat':         return 'group';
    case 'Meeting':      return 'meeting';
    case 'TeamsStandardChannel': return 'channel';
  }
  // Unknown / unspecified. Look at the id shape to guess a sensible kind.
  if (convId) {
    if (convId === '48:notes') return 'chat';
    if (convId.startsWith('19:meeting_')) return 'meeting';
    if (/^19:[a-f0-9-]{36}_[a-f0-9-]{36}@/i.test(convId)) return 'chat';
    if (/^19:[^@]+@thread\.v2$/i.test(convId)) return 'group';
    if (/^19:[^@]+@thread\.tacv2$/i.test(convId)) return 'channel';
  }
  return 'chat';
}

// Extract the OTHER party's UUID from a 1:1 chat id. Returns:
//   - string uuid when the id is a normal 1:1 (the non-self member)
//   - 'self'   for the self-chat. Teams uses a special id `48:notes`
//              for "chat with yourself" (sidebar label "<you> (You)")
//              which does NOT follow the 19:<a>_<b>@ format — hence
//              the explicit early return.
//   - null     if the id doesn't match the expected shape, or selfUuid
//              isn't available
function extractOtherPartyUuid(id: string, selfUuid: string | null): string | 'self' | null {
  if (id === '48:notes') return 'self';
  if (!selfUuid) return null;
  const m = id.match(/^19:([a-f0-9-]{36})_([a-f0-9-]{36})@/i);
  if (!m) return null;
  const a = m[1].toLowerCase();
  const b = m[2].toLowerCase();
  const self = selfUuid.toLowerCase();
  if (a === self && b === self) return 'self';
  if (a === self) return b;
  if (b === self) return a;
  return null; // Neither member matches — shouldn't happen, but be defensive.
}

// Format a UUID as an `8:orgid:<uuid>` MRI so resolveUserNames() can
// batch-resolve it alongside the MRIs that come from roster lookups.
const uuidToMri = (uuid: string): string => `8:orgid:${uuid}`;

// Teams stamps these placeholder labels onto chats whose counterparty it
// can no longer resolve (left-org user, deleted account). They look like
// real chat titles to our picker and hide which user the chat is with —
// our own resolution chain (Graph → contacts → replychain → roster) often
// still has the name cached. Treat them as "missing topic" so the chain
// runs instead. English only for now; extend per user report in other
// locales as needed.
const TEAMS_SCRUBBED_TOPICS = new Set(['unknown user', 'just me']);
const isScrubbedTopic = (topic: string | undefined | null): boolean =>
  !!topic && TEAMS_SCRUBBED_TOPICS.has(topic.toLowerCase().trim());

// Short date formatter for meeting subtitles. Same-year meetings show
// "Apr 24"; cross-year meetings include the year. Locale-aware so
// Turkish shows "24 Nis" etc.
function formatMeetingDate(iso: string | undefined): string | undefined {
  if (!iso) return undefined;
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return undefined;
  const sameYear = d.getFullYear() === new Date().getFullYear();
  return d.toLocaleDateString(undefined, sameYear
    ? { month: 'short', day: 'numeric' }
    : { year: 'numeric', month: 'short', day: 'numeric' });
}

// Module-scope cache for Graph /me/joinedTeams. The call is ~2 s on
// first fire and the result rarely changes within a session — users
// join/leave teams on a slow cadence. Caching across popup opens for
// the content-script lifetime is safe. We don't expire; SW/content
// restart flushes it.
let _joinedTeamsCache: Map<string, string> | null = null;
async function fetchJoinedTeams(graphToken: string): Promise<Map<string, string>> {
  if (_joinedTeamsCache) return _joinedTeamsCache;
  try {
    const resp = await fetch('https://graph.microsoft.com/v1.0/me/joinedTeams', {
      headers: { 'Authorization': `Bearer ${graphToken}` },
      signal: AbortSignal.timeout(15_000),
    });
    if (!resp.ok) return new Map();
    const data = (await resp.json()) as { value?: Array<{ id?: string; displayName?: string }> };
    const map = new Map<string, string>();
    for (const t of data.value || []) {
      if (t.id && t.displayName) map.set(t.id.toLowerCase(), t.displayName);
    }
    _joinedTeamsCache = map;
    return map;
  } catch {
    return new Map();
  }
}

// Member entry as returned by the chat-service roster endpoint. The
// Teams chat service caches federated contacts' display names locally
// (`friendlyName`) — that's how Teams' own UI labels external users
// without needing a Graph lookup. We use it as the second source of
// truth when Graph misses (external tenants).
type RosterMember = {
  mri: string;
  friendlyName?: string;
};

type RosterFetchResult = {
  members: RosterMember[];
  totalMemberCount?: number;
};

// Fetch the roster of a single thread. Used both for unnamed group
// chats (to assemble a member-list label) and for 1:1 chats where the
// Graph lookup missed (typically external/federated contacts).
async function fetchThreadRoster(
  config: Pick<TeamsApiConfig, 'chatServiceUrl' | 'ic3Token'>,
  threadId: string,
): Promise<RosterFetchResult> {
  try {
    const url = `${config.chatServiceUrl}/v1/threads/${encodeURIComponent(threadId)}/members`;
    const resp = await fetch(url, {
      headers: { 'Authorization': `Bearer ${config.ic3Token}` },
      signal: AbortSignal.timeout(10_000),
    });
    if (!resp.ok) return { members: [] };
    const data = (await resp.json()) as {
      totalMemberCount?: number;
      members?: Array<{ id?: string; mri?: string; friendlyName?: string; properties?: { friendlyName?: string } }>;
    };
    const out: RosterMember[] = [];
    for (const m of data.members || []) {
      const mri = m.id || m.mri || '';
      if (!mri) continue;
      const friendlyName = m.friendlyName?.trim() || m.properties?.friendlyName?.trim();
      out.push(friendlyName ? { mri, friendlyName } : { mri });
    }
    return { members: out, totalMemberCount: data.totalMemberCount };
  } catch {
    return { members: [] };
  }
}

// Shared enrichment kernel — used by the IDB conversation source path.
// Takes already-classified records and resolves names (Graph batch +
// roster fallback for federated contacts), populates subtitles, sorts
// recency-first, returns picker-ready summaries.
type ClassifiedConversation = { conv: RawConversation; kind: ConversationSummary['kind'] };

async function enrichClassifiedConversations(
  work: ClassifiedConversation[],
  config: Pick<TeamsApiConfig, 'chatServiceUrl' | 'ic3Token'>,
): Promise<ConversationSummary[]> {
  const { readContacts, buildContactNameMap, readReplychainSenders } = await import('./teams-state');
  const contactsPromise = readContacts().then(buildContactNameMap);
  const sendersByConvPromise = readReplychainSenders();
  const selfUuid = getCurrentUserUuid();
  const graphToken = await getGraphToken();

  // Identify enrichment targets in a single pass.
  const mrisToResolve = new Set<string>();       // 1:1 other-parties + roster members
  const unnamedGroupIds: string[] = [];          // groups missing a topic
  const teamGroupIds = new Set<string>();        // channel parent-team ids
  if (selfUuid) mrisToResolve.add(uuidToMri(selfUuid));  // for self-chat labels
  for (const { conv, kind } of work) {
    if (kind === 'chat') {
      const other = extractOtherPartyUuid(conv.id!, selfUuid);
      if (other && other !== 'self') mrisToResolve.add(uuidToMri(other));
    } else if (kind === 'group') {
      if (!conv.threadProperties?.topic?.trim()) unnamedGroupIds.push(conv.id!);
    } else if (kind === 'channel') {
      const gid = conv.threadProperties?.groupId;
      if (gid) teamGroupIds.add(gid.toLowerCase());
    }
  }

  // Fire the first enrichment fan-out in parallel. joinedTeams is
  // independent; group rosters provide the membership lists for the
  // unnamed-group label path.
  const [teamMap, groupRosterResults] = await Promise.all([
    teamGroupIds.size > 0 && graphToken
      ? fetchJoinedTeams(graphToken)
      : Promise.resolve(new Map<string, string>()),
    Promise.all(unnamedGroupIds.map(async (id) => ({
      id,
      roster: await fetchThreadRoster(config, id),
    }))),
  ]);

  // Index group rosters by thread + collect their MRIs into the Graph
  // batch (so member display names resolve in one /users call).
  // totalMemberCount comes from the roster API and is the canonical
  // size — including members our /members enumeration may have missed.
  const rosterByThread = new Map<string, RosterMember[]>();
  const rosterTotalByThread = new Map<string, number>();
  for (const { id, roster } of groupRosterResults) {
    rosterByThread.set(id, roster.members);
    if (typeof roster.totalMemberCount === 'number') {
      rosterTotalByThread.set(id, roster.totalMemberCount);
    }
    for (const m of roster.members) mrisToResolve.add(m.mri);
  }

  const nameMap = (graphToken && mrisToResolve.size > 0)
    ? await resolveUserNames([...mrisToResolve], graphToken)
    : new Map<string, string>();

  // Local contact cache — Teams stores display names for users we've
  // exchanged messages with, including some federated contacts whose
  // names Graph can't resolve.
  const contactsByMri = await contactsPromise;

  // Per-conversation list of recent message senders, recency-first.
  // Pulled from Teams' replychain-manager IDB store. The only durable
  // local source for federated contact names + the right basis for
  // ranking unnamed-group member display order.
  const sendersByConv = await sendersByConvPromise;

  const selfMriSubstring = selfUuid ? selfUuid.toLowerCase() : '';

  // Phase 2: 1:1 chats whose other-party Graph lookup missed are
  // typically federated/external contacts (different tenant, /users
  // returns nothing). Their display name is cached locally by Teams
  // as `friendlyName` on the thread roster — fetch those rosters
  // now, in parallel, so the picker shows real names instead of a
  // "(direct message)" placeholder.
  const externalChatsNeedingRoster: string[] = [];
  for (const { conv, kind } of work) {
    if (kind !== 'chat') continue;
    const other = extractOtherPartyUuid(conv.id!, selfUuid);
    if (!other || other === 'self') continue;
    if (!nameMap.get(uuidToMri(other))) externalChatsNeedingRoster.push(conv.id!);
  }
  if (externalChatsNeedingRoster.length > 0) {
    const externalRosters = await Promise.all(
      externalChatsNeedingRoster.map(async (id) => ({
        id,
        roster: await fetchThreadRoster(config, id),
      }))
    );
    for (const { id, roster } of externalRosters) {
      rosterByThread.set(id, roster.members);
      if (typeof roster.totalMemberCount === 'number') {
        rosterTotalByThread.set(id, roster.totalMemberCount);
      }
    }
  }

  const out: ConversationSummary[] = [];
  for (const { conv, kind } of work) {
    const topic = conv.threadProperties?.topic?.trim();
    let name = '';
    let subtitle: string | undefined;
    let isSelfChat = false;
    let groupMembers: string[] | undefined;
    let groupExtraMembers: number | undefined;

    if (kind === 'chat') {
      const other = extractOtherPartyUuid(conv.id!, selfUuid);
      if (other === 'self') {
        isSelfChat = true;
        name = (selfUuid && nameMap.get(uuidToMri(selfUuid)))
          || conv.lastMessage?.imdisplayname?.trim()
          || '';
      } else if (topic && !isScrubbedTopic(topic)) {
        // chatTitle.shortTitle was promoted into topic upstream —
        // for federated 1:1s this is Teams' own pre-computed name
        // (e.g. "Jane Doe"), already correct. Sentinels like
        // "Unknown User" are filtered above so the resolution chain
        // below gets a chance to surface a cached name instead.
        name = topic;
      } else {
        const mri = other ? uuidToMri(other) : null;
        const graphName = mri ? nameMap.get(mri) : undefined;
        if (graphName) {
          name = graphName;
        } else {
          // Resolution chain for external/federated 1:1 chats:
          //   1. Local capiv3 contacts cache (Teams stored their
          //      display name when we last interacted)
          //   2. replychain-manager — the most recent non-self
          //      sender's cached display name. The only durable
          //      source for most federated user names.
          //   3. Thread roster's friendlyName (rarely populated for
          //      federated members in the user's tenant)
          //   4. Last-message sender name, if it wasn't self
          //   5. Placeholder + date
          const cachedName = mri ? contactsByMri.get(mri) : undefined;
          if (cachedName) {
            name = cachedName;
          } else {
            const senders = sendersByConv.get(conv.id!) || [];
            const nonSelfSender = senders.find(s => !selfMriSubstring || !s.mri.toLowerCase().includes(selfMriSubstring));
            if (nonSelfSender) {
              name = nonSelfSender.name;
            } else {
              const roster = rosterByThread.get(conv.id!) || [];
              const otherMember = roster.find(m => !selfMriSubstring || !m.mri.toLowerCase().includes(selfMriSubstring));
              if (otherMember?.friendlyName) {
                name = otherMember.friendlyName;
              } else {
                const lmFrom = conv.lastMessage?.from || '';
                const lmSenderIsSelf = selfMriSubstring && lmFrom.toLowerCase().includes(selfMriSubstring);
                const lmName = conv.lastMessage?.imdisplayname?.trim();
                if (lmName && !lmSenderIsSelf) {
                  name = lmName;
                } else {
                  const when = formatMeetingDate(conv.lastMessage?.composetime || conv.lastMessage?.originalarrivaltime);
                  if (when) subtitle = when;
                }
              }
            }
          }
        }
      }
    } else if (kind === 'group') {
      if (topic && !isScrubbedTopic(topic)) {
        name = topic;
      } else {
        // Unnamed group: collect all candidate names (roster + recent
        // senders), dedupe by MRI, sort alphabetically by first name
        // to match Teams' sidebar ordering. Names come from the best
        // source available: Graph → capiv3 → replychain → roster
        // friendlyName.
        const roster = rosterByThread.get(conv.id!) || [];
        const senders = sendersByConv.get(conv.id!) || [];
        const others = selfMriSubstring
          ? roster.filter(m => !m.mri.toLowerCase().includes(selfMriSubstring))
          : roster;
        const senderNonSelf = selfMriSubstring
          ? senders.filter(s => !s.mri.toLowerCase().includes(selfMriSubstring))
          : senders;

        type Candidate = { mri: string; name: string };
        const candidatesByMri = new Map<string, Candidate>();
        const resolveName = (mri: string, ...fallbacks: (string | undefined)[]) => {
          const n = nameMap.get(mri) || contactsByMri.get(mri);
          if (n) return n;
          for (const f of fallbacks) if (f) return f;
          return undefined;
        };
        for (const m of others) {
          const senderName = senderNonSelf.find(s => s.mri === m.mri)?.name;
          const n = resolveName(m.mri, senderName, m.friendlyName);
          if (n) candidatesByMri.set(m.mri, { mri: m.mri, name: n });
        }
        // Senders we know about that aren't in the roster (e.g. left
        // members whose messages are still cached) — include them too.
        for (const s of senderNonSelf) {
          if (candidatesByMri.has(s.mri)) continue;
          const n = resolveName(s.mri, s.name);
          if (n) candidatesByMri.set(s.mri, { mri: s.mri, name: n });
        }

        const sorted = [...candidatesByMri.values()].sort((a, b) =>
          (a.name.split(/\s+/)[0] || a.name).localeCompare(b.name.split(/\s+/)[0] || b.name)
        );
        const picked = sorted.slice(0, 3).map(c => c.name);
        if (picked.length > 0) {
          groupMembers = picked;
          const apiTotal = rosterTotalByThread.get(conv.id!);
          const totalOthers = apiTotal != null
            ? Math.max(0, apiTotal - 1)
            : Math.max(others.length, candidatesByMri.size);
          const extra = Math.max(0, totalOthers - picked.length);
          if (extra > 0) groupExtraMembers = extra;
        }
      }
    } else if (kind === 'meeting') {
      if (topic) name = topic;
      const when = formatMeetingDate(conv.lastMessage?.composetime || conv.lastMessage?.originalarrivaltime);
      if (when) subtitle = when;
    } else if (kind === 'channel') {
      if (topic && !isScrubbedTopic(topic)) {
        name = topic;
      } else {
        // Channel-side analogue of the chat resolution chain. Triggers
        // when Teams' precomputed channel title is a scrubbed sentinel
        // (typically a private channel whose only-other-member was
        // removed from the directory). Surface the most-recent non-self
        // sender from the cached message authors.
        const senders = sendersByConv.get(conv.id!) || [];
        const nonSelfSender = senders.find(s => !selfMriSubstring || !s.mri.toLowerCase().includes(selfMriSubstring));
        if (nonSelfSender) name = nonSelfSender.name;
      }
      const gid = conv.threadProperties?.groupId?.toLowerCase();
      const teamName = gid ? teamMap.get(gid) : undefined;
      if (teamName) subtitle = teamName;
    }

    out.push({
      id: conv.id!,
      kind,
      name,
      subtitle,
      isSelfChat: isSelfChat || undefined,
      groupMembers,
      groupExtraMembers,
      lastActivity: conv.lastMessage?.composetime || conv.lastMessage?.originalarrivaltime,
    });
  }

  // If the API didn't return `48:notes`, probe the conversations
  // endpoint directly for that specific id. The list endpoint is
  // inconsistent across tenants — sometimes self-chat is present,
  // sometimes not — but the single-conversation GET works when the
  // thread exists. Keeps the path API-only (no DOM scrape).
  if (!out.some(c => c.id === '48:notes')) {
    try {
      const r = await fetch(`${config.chatServiceUrl}/v1/users/ME/conversations/48:notes`, {
        headers: { 'Authorization': `Bearer ${config.ic3Token}` },
        signal: AbortSignal.timeout(5_000),
      });
      if (r.ok) {
        const data = (await r.json()) as RawConversation;
        // We accept any 2xx as "self-chat exists". The returned record
        // may or may not have a lastMessage; that's fine — we know the
        // id is valid and the UI gets a real row.
        const selfName = (selfUuid && nameMap.get(uuidToMri(selfUuid))) || '';
        out.unshift({
          id: '48:notes',
          kind: 'chat',
          name: selfName,               // empty when Graph missed; picker localises
          isSelfChat: true,
          lastActivity: data.lastMessage?.composetime
            || data.lastMessage?.originalarrivaltime
            || new Date().toISOString(),
        });
      }
    } catch { /* 404 / network / timeout — user has no self-chat or it's not reachable */ }
  }

  // Recency-first. Missing timestamps sink to the bottom.
  out.sort((a, b) => (b.lastActivity || '').localeCompare(a.lastActivity || ''));
  return out;
}

/**
 * IDB-backed sister of {@link listConversations}. Reads from Teams' own
 * client-side IndexedDB store (Teams:conversation-manager) instead of
 * the chat-service API, which:
 *   - returns the full local set (includes meeting-derived chats and
 *     other niche product types the API omits)
 *   - is consistent with what the user sees in the sidebar
 *   - costs no network (instant)
 *
 * Enrichment (Graph name resolution, unnamed-group roster, federated
 * `friendlyName`, channel team-name) is shared with the API path —
 * both call the same helpers further down.
 */
/**
 * Quick IDB-only render of the conversation list — no Graph, no
 * roster fetches. Used for the first paint of the picker so the UI
 * is responsive in <100 ms; the popup follows up with the full
 * enriched listConversationsFromIdb to fill in resolved names.
 *
 * Names come from `threadProperties.topic` (groups / meetings /
 * channels) or `lastMessage.imdisplayname` (1:1s where the other
 * party sent last). Where neither is usable the row stays
 * "(direct message)" until the full enrichment lands.
 */
export async function listConversationsFromIdbQuick(): Promise<ListConversationsResult> {
  const { readConversationList, readFolders } = await import('./teams-state');
  const [records, rawFolders] = await Promise.all([
    readConversationList(),
    readFolders(),
  ]);
  if (records.length === 0) return { conversations: [], folders: [] };
  // Folder index here too — the quick path renders the full picker UI
  // including the folder rail, so we need the folders ready on first
  // paint. Cheap (one IDB read, no network).
  const { folders, folderIdsByConv } = buildFolderIndex(rawFolders);

  const selfUuid = getCurrentUserUuid();
  const selfMriSubstring = selfUuid ? selfUuid.toLowerCase() : '';

  const out: ConversationSummary[] = [];
  for (const r of records) {
    if (!r.id) continue;
    if (r.threadProperties?.isDeleted === true) continue;
    if (r.id !== '48:notes' && r.id.startsWith('48:')) continue;
    if (r.threadProperties?.threadType === 'streamofnotes' && r.id !== '48:notes') continue;

    const ptt = idbThreadTypeToProductThreadType(r.threadProperties?.threadType, r.id);
    const kind = classifyConversationKind(ptt, r.id);
    if (!kind) continue;

    const isSelfChat = r.id === '48:notes';
    // Teams pre-computes the sidebar title and stores it in
    // chatTitle.shortTitle — exact "Alice, Bob, Carol, +N" / "Jane
    // Doe" string. Use it when available; fall back to topic; then
    // to last-message sender. Skips all the Graph/roster work.
    // Scrubbed sentinels ("Unknown User" / "Just me") are treated as
    // missing — better to show "(direct message)" briefly until the
    // full enrichment surfaces a real name than to lock in a label
    // that hides which person the chat is with.
    const rawTitle = r.chatTitle?.shortTitle?.trim() || r.threadProperties?.topic?.trim() || '';
    let name = isScrubbedTopic(rawTitle) ? '' : rawTitle;
    if (!name && kind === 'chat' && !isSelfChat) {
      const lmFrom = r.lastMessage?.from || '';
      const lmSenderIsSelf = selfMriSubstring && lmFrom.toLowerCase().includes(selfMriSubstring);
      const lmName = r.lastMessage?.imdisplayname?.trim();
      if (lmName && !lmSenderIsSelf) name = lmName;
    }

    const folderIds = folderIdsByConv.get(r.id);
    out.push({
      id: r.id,
      kind,
      name,
      isSelfChat: isSelfChat || undefined,
      folderIds: folderIds && folderIds.length > 0 ? folderIds : undefined,
      lastActivity: r.lastMessage?.composetime
        || r.lastMessage?.originalarrivaltime
        || (r.lastMessageTimeUtc ? new Date(r.lastMessageTimeUtc).toISOString() : undefined),
    });
  }
  out.sort((a, b) => (b.lastActivity || '').localeCompare(a.lastActivity || ''));
  return { conversations: out, folders };
}

// Result of an IDB-backed conversation list fetch. Folders may be empty
// if the user has no Favorites and no UserDefined folders, OR when this
// is the API-fallback path (no folder source).
export type ListConversationsResult = {
  conversations: ConversationSummary[];
  folders: FolderSummary[];
};

// Build the picker-side FolderSummary list and a per-conversation index
// from raw Teams folder records. Strips system-computed folders the
// picker shouldn't expose (MeetingChats / MutedChats / RecentChats /
// TeamsAndChannels / QuickViews) — those duplicate the kind filter or
// aren't real conversations. Only Favorites + UserDefined survive.
//
// SAFETY NOTE: TeamsAndChannels in particular may behave subtly different
// from `kind === 'channel'` for users with team-channel access (we can't
// fully test it from a chats-only account). Hidden by default keeps that
// risk contained. If a future contributor wants to expose system folders
// as an opt-in, they should verify TeamsAndChannels content against the
// kind=channel filter before turning it on.
// System folder names Teams writes automatically — these get filtered
// out because they either duplicate the kind filter (MeetingChats /
// TeamsAndChannels), provide no extra info (RecentChats — picker is
// already recency-sorted), surface a state we don't filter on
// (MutedChats), or contain pseudo-IDs that aren't real conversations
// (QuickViews holds slice-activities-mentions and the like).
//
// Match by NAME rather than folderType: Teams' folderType for user-
// created folders varies (UserDefined, CustomChatList, etc. across
// tenants and versions). Name is the stable contract.
const SYSTEM_FOLDER_NAMES = new Set([
  'MeetingChats',
  'MutedChats',
  'QuickViews',
  'RecentChats',
  'TeamsAndChannels',
]);

function buildFolderIndex(rawFolders: Array<{
  id: string;
  name?: string;
  folderType?: string;
  isHidden?: boolean;
  conversations?: Array<{ id: string }>;
}>): { folders: FolderSummary[]; folderIdsByConv: Map<string, string[]> } {
  const folders: FolderSummary[] = [];
  const folderIdsByConv = new Map<string, string[]>();
  for (const f of rawFolders) {
    if (!f.id || !f.name) continue;
    if (f.isHidden) continue;
    // Drop system-curated folders by NAME — see comment above for why
    // name beats folderType. Anything else is either Favorites (the
    // special star folder) or a user-created folder.
    if (SYSTEM_FOLDER_NAMES.has(f.name)) continue;
    const isFavorite = f.folderType === 'Favorites' || f.name === 'Favorites';
    // Defensively skip slice-* member ids in case Teams ever inlines
    // QuickViews-style pseudo-conversations into another folder.
    const convIds = (f.conversations || [])
      .map(c => c?.id)
      .filter((id): id is string => !!id && !id.startsWith('slice-'));
    if (convIds.length === 0) continue; // hide empty folders
    folders.push({
      id: f.id,
      name: f.name,
      kind: isFavorite ? 'favorites' : 'user',
      count: convIds.length,
    });
    for (const cid of convIds) {
      const list = folderIdsByConv.get(cid) || [];
      list.push(f.id);
      folderIdsByConv.set(cid, list);
    }
  }
  // Render order: Favorites first (it's the system-blessed "important"
  // folder), then user folders alphabetically. The popup shows them in
  // this order in the rail.
  folders.sort((a, b) => {
    if (a.kind !== b.kind) return a.kind === 'favorites' ? -1 : 1;
    return a.name.localeCompare(b.name);
  });
  return { folders, folderIdsByConv };
}

export async function listConversationsFromIdb(
  config: Pick<TeamsApiConfig, 'chatServiceUrl' | 'ic3Token'>,
  _opts: { includeHidden?: boolean } = {},
): Promise<ListConversationsResult> {
  const { readConversationList, readFolders } = await import('./teams-state');
  const [records, rawFolders] = await Promise.all([
    readConversationList(),
    readFolders(),
  ]);
  if (records.length === 0) return { conversations: [], folders: [] };

  // Don't filter by threadProperties.hidden. The flag turns out to
  // mean something other than "user removed from sidebar" — Teams
  // routinely sets it true on active 1:1 chats whether or not they're
  // pinned to a folder. Filtering on it dropped real recent chats from
  // the picker even though they were sitting at the top of the user's
  // sidebar. Recency sort + search keeps the full list navigable.
  const { folders, folderIdsByConv } = buildFolderIndex(rawFolders);

  // Map IDB records to the RawConversation shape the rest of the
  // pipeline already understands. The two shapes overlap; we mostly
  // just relabel fields. `productThreadType` doesn't exist in IDB
  // (the field is `threadType`) — pass it through as a synthetic
  // value so the classifier's ID-shape fallback fires reliably.
  type Work = { conv: RawConversation; kind: ConversationSummary['kind'] };
  const work: Work[] = [];
  for (const r of records) {
    if (!r.id) continue;
    // Drop conversations the user explicitly deleted (different from
    // hidden) so they don't reappear in the picker.
    if (r.threadProperties?.isDeleted === true) continue;
    // Filter Teams' internal system pseudo-chats — Activity feed,
    // Mentions, Call logs, Saved/Starred messages, Drafts, etc. They
    // share the conversation-manager store but aren't exportable
    // conversations. The `48:` prefix marks Skype-protocol special
    // threads; the only one we keep is `48:notes` (self-chat).
    if (r.id !== '48:notes' && r.id.startsWith('48:')) continue;
    if (r.threadProperties?.threadType === 'streamofnotes' && r.id !== '48:notes') continue;
    // Promote chatTitle.shortTitle into the topic slot — Teams already
    // computed the exact sidebar string (covers unnamed groups, 1:1s
    // with federated contacts, etc.), so the downstream enrichment
    // never has to reconstruct it from rosters / Graph / replychain.
    const idbTitle = r.chatTitle?.shortTitle?.trim();
    const conv: RawConversation = {
      id: r.id,
      threadProperties: {
        topic: r.threadProperties?.topic?.trim() || idbTitle || undefined,
        productThreadType: idbThreadTypeToProductThreadType(r.threadProperties?.threadType, r.id),
        groupId: r.threadProperties?.groupId,
      },
      lastMessage: r.lastMessage
        ? {
            imdisplayname: r.lastMessage.imdisplayname,
            fromDisplayNameInToken: r.lastMessage.fromDisplayNameInToken,
            composetime: r.lastMessage.composetime,
            originalarrivaltime: r.lastMessage.originalarrivaltime,
            from: r.lastMessage.from,
          }
        : undefined,
    };
    const kind = classifyConversationKind(conv.threadProperties?.productThreadType, r.id);
    if (!kind) continue;
    work.push({ conv, kind });
  }

  const conversations = await enrichClassifiedConversations(work, config);
  // Stamp folderIds onto each conversation. The picker filters by these.
  for (const c of conversations) {
    const ids = folderIdsByConv.get(c.id);
    if (ids && ids.length > 0) c.folderIds = ids;
  }
  return { conversations, folders };
}

// Map IDB's `threadType` field to the chat-service `productThreadType`
// the classifier expects. IDB uses lowercase short names; the service
// uses camelCase long names. When the threadType is missing or unknown
// we leave productThreadType undefined and let classifyConversationKind
// fall through to its ID-shape inference path.
function idbThreadTypeToProductThreadType(t: string | undefined, id: string): string | undefined {
  switch (t) {
    case 'chat':
      // IDB collapses 1:1 and group under 'chat'. Disambiguate by id.
      return /^19:[a-f0-9-]{36}_[a-f0-9-]{36}@/i.test(id) || id === '48:notes'
        ? 'OneToOneChat' : 'Chat';
    case 'meeting': return 'Meeting';
    case 'channel': return 'TeamsStandardChannel';
    default: return undefined;
  }
}


// ── Active Conversation ID ─────────────────────────────────────────────

/**
 * Return the conversation id of whatever chat the Teams tab is currently
 * showing. Reads `tmp.session.<selfUuid>-mainWindowNavHistory[index]` —
 * Teams writes the active conversation id there synchronously on every
 * sidebar click, so the lookup is locale-independent, has no DOM-race
 * window, and returns null cleanly when the user is on a non-chat view
 * (Activity, Calendar, app landing).
 *
 * Replaces a long-running DOM-heuristic implementation that combined
 * document.title parsing, sidebar [data-tabster] matching, IDB
 * replychain lookup via visible data-mid values, URL/hash regex, and
 * a chat-pane data-attribute scan — all of which were either
 * locale-fragile, race-prone, or aged out as Teams' UI changed. The
 * sessionStorage path is the load-bearing finding from the IDB-source
 * pivot (see docs/TEAMS_INTERNALS.md, "Active conversation ID").
 */
export async function extractConversationId(): Promise<string | null> {
  const selfUuid = getCurrentUserUuid();
  if (!selfUuid) return null;
  const { readActiveConversationId } = await import('./teams-state');
  return readActiveConversationId(selfUuid);
}

// ── Paginated Message Fetcher ──────────────────────────────────────────

const INITIAL_DELAY_MS = 150;
const MAX_PAGES = 500; // safety limit (~100k messages)
const MAX_RETRIES = 5;

const sleep = (ms: number) => new Promise(resolve => setTimeout(resolve, ms));

// 403 is officially "forbidden" but Teams returns it in transient
// situations too — token-rotation race, per-resource throttling that
// looks like a 429 but isn't tagged as one, brief auth-state churn
// during MSAL refresh. So we retry 403 (and 5xx) on a tight budget:
// 3 attempts at 1s/2s/4s = ~7s worst case per failing chat. If still
// failing after that, the chat is genuinely inaccessible and the
// caller decides what to do (in multi-chat mode: record a per-chat
// failure rather than fall back to DOM scroll, which would scrape
// whichever chat happens to be active in the user's tab).
const SHORT_RETRIES = 3;
const SHORT_RETRY_BACKOFF_MS = [1000, 2000, 4000];

/**
 * Fetch a single page with retry logic for transient errors.
 * Returns the parsed response or throws on non-retryable errors.
 */
// Build the request headers for a messaging API call.
//
//  - Work / School (`*.asm.skype.com`, `*.ic3.teams.office.com`, …):
//        Authorization: Bearer <token>
//
//  - Teams Free (proxied via `teams.live.com/api/chatsvc/consumer/...`):
//    the consumer chat service expects the raw Skype token in a
//    lowercase `authentication` header — NOT `Authorization`, and NO
//    `Bearer` prefix. Plus a few client-identification headers that
//    the consumer service appears to require (or at least Teams' own
//    UI always sends them; verified from HAR capture of the live
//    request that returns 200).
function buildMessagingHeaders(url: string, token: string): Record<string, string> {
  if (/teams\.live\.com\/api\/chatsvc\/consumer\b/i.test(url)) {
    // The proxy expects the literal scheme `skypetoken=<JWT>` in the
    // lowercase `authentication` header — verified against a Teams UI
    // HAR capture. The other four headers identify the client (Teams
    // Free / SfL surface); the proxy 401s without them.
    return {
      'authentication': `skypetoken=${token}`,
      'behavioroverride': 'redirectAs404',
      'ms-ic3-product': 'tfl',
      'ms-ic3-additional-product': 'Sfl',
      'x-ms-request-priority': '0',
    };
  }
  return { 'Authorization': `Bearer ${token}` };
}

// Combine two AbortSignals so the resulting signal aborts when EITHER
// upstream signal aborts. AbortSignal.any() does this natively but
// only landed in Firefox 124 / Chrome 116; we target Firefox 109+
// so we ship the manual equivalent.
function combineAbortSignals(a: AbortSignal | undefined, b: AbortSignal): AbortSignal {
  if (!a) return b;
  if (a.aborted) return a;
  if (b.aborted) return b;
  const ctrl = new AbortController();
  const onAbort = () => ctrl.abort();
  a.addEventListener('abort', onAbort, { once: true });
  b.addEventListener('abort', onAbort, { once: true });
  return ctrl.signal;
}

async function fetchPageWithRetry(
  url: string,
  token: string,
  signal?: AbortSignal,
): Promise<MessagesResponse> {
  // Teams Free's chatsvc/consumer proxy needs the user's session
  // cookies to bridge auth into the Skype backend (a 401 otherwise).
  // Work/School's *.asm.skype.com endpoint doesn't need cookies — the
  // Bearer header alone authenticates — so we omit `credentials` there
  // to keep the pre-v1.4 request shape exactly intact and avoid any
  // chance of stale third-party Skype cookies confusing the server.
  const isTeamsFreeProxy = /teams\.live\.com\/api\/chatsvc\/consumer\b/i.test(url);
  for (let attempt = 0; attempt <= MAX_RETRIES; attempt++) {
    const headers = buildMessagingHeaders(url, token);
    // Always combine the caller's signal with an internal 30 s
    // timeout. Earlier we did `signal || AbortSignal.timeout(...)`
    // which silently dropped the timeout whenever the caller passed
    // a signal — the result was that an export could hang forever
    // if wifi died mid-fetch (Firefox doesn't reliably throw
    // NetworkError when a connection silently drops; the request
    // queues until the OS gives up, which can be many minutes).
    // With the combined signal, fetch reliably aborts at 30 s and
    // the throw bubbles to apiScrape which surfaces the partial
    // signal.
    const timeoutSignal = AbortSignal.timeout(30_000);
    const init: RequestInit = {
      headers,
      signal: combineAbortSignals(signal, timeoutSignal),
    };
    if (isTeamsFreeProxy) init.credentials = 'include';
    let resp: Response;
    try {
      resp = await fetch(url, init);
    } catch (err) {
      // If the user's own signal aborted, propagate as cancellation —
      // apiScrape's catch re-throws AbortError so content.ts treats
      // it as a user cancel rather than as an API failure.
      if (signal?.aborted) throw err;
      // Our internal timeout fired and the fetch threw AbortError.
      // Convert to a TypeError with a NetworkError message so
      // apiScrape's isNetworkError() classifier picks it up — that's
      // what flips the export to partial=network downstream.
      if (timeoutSignal.aborted) {
        throw new TypeError(`NetworkError: fetch timed out after 30s (no network?) for ${url.slice(0, 200)}`);
      }
      throw err;
    }

    // Diagnostic on 401: dump request shape (header names + value
    // lengths, never values) and response body (truncated) so we can
    // compare against a known-working HAR capture.
    if (resp.status === 401) {
      const headerSummary: Record<string, string> = {};
      for (const [k, v] of Object.entries(headers)) {
        headerSummary[k] = `<len=${(v as string).length}>`;
      }
      const wwwAuth = resp.headers.get('www-authenticate');
      let body = '';
      try { body = (await resp.clone().text()).slice(0, 400); } catch { /* ignore */ }
      console.warn(`[API] 401 from messaging`,
        JSON.stringify({ url, headers: headerSummary, wwwAuthenticate: wwwAuth, bodySnippet: body }));
    }

    if (resp.status === 429) {
      if (attempt >= MAX_RETRIES) {
        throw new Error('Rate limited after max retries');
      }
      // Floor the wait at the exponential backoff. Teams sometimes
      // returns Retry-After: 0 (a "back off" signal, not "go now");
      // honouring 0 literally would burn every retry slot in a hot
      // loop. Use whichever is larger of the header and the backoff.
      const backoffWait = Math.min(2000 * Math.pow(2, attempt), 30_000); // 2s, 4s, 8s, 16s, 30s
      const retryAfter = resp.headers.get('Retry-After');
      const retryAfterMs = retryAfter ? parseInt(retryAfter, 10) * 1000 : 0;
      const waitMs = Math.max(backoffWait, Number.isFinite(retryAfterMs) ? retryAfterMs : 0);
      console.log(`[API] Rate limited (429), retrying in ${waitMs}ms (attempt ${attempt + 1}/${MAX_RETRIES})`);
      await sleep(waitMs);
      continue;
    }

    if (resp.status === 403 || (resp.status >= 500 && resp.status < 600)) {
      if (attempt >= SHORT_RETRIES - 1) {
        throw new Error(`Messages API error: ${resp.status} ${resp.statusText}`);
      }
      const waitMs = SHORT_RETRY_BACKOFF_MS[attempt] ?? 4000;
      console.log(`[API] Transient ${resp.status}, retrying in ${waitMs}ms (attempt ${attempt + 1}/${SHORT_RETRIES})`);
      await sleep(waitMs);
      continue;
    }

    if (!resp.ok) {
      throw new Error(`Messages API error: ${resp.status} ${resp.statusText}`);
    }

    try {
      return await resp.json();
    } catch {
      // JSON parse error at end of history — return empty
      console.log(`[API] JSON parse error, treating as end of history`);
      return { messages: [] };
    }
  }

  throw new Error('Unreachable');
}

/**
 * Fetch all messages for a conversation using the Teams chat service API.
 *
 * Paginates via backwardLink (newest → oldest) until history is exhausted, or
 * until the page crosses `startAtISO` (date-bound exports stop early).
 * Re-reads the ic3 token from localStorage before each page to handle refresh.
 * Retries on 429 with exponential backoff.
 */
export async function fetchAllMessages(
  config: TeamsApiConfig,
  conversationId: string,
  onProgress?: (p: FetchProgress) => void,
  signal?: AbortSignal,
  startAtISO?: string | null,
): Promise<TeamsApiMessage[]> {
  const startAtMs = startAtISO ? Date.parse(startAtISO) : NaN;
  const allMessages: TeamsApiMessage[] = [];
  const encodedConvId = encodeURIComponent(conversationId);
  let nextUrl: string | null =
    `${config.chatServiceUrl}/v1/users/ME/conversations/${encodedConvId}/messages?pageSize=200&startTime=1&view=msnp24Equivalent%7CsupportsMessageProperties`;
  let page = 0;
  let delayMs = INITIAL_DELAY_MS;

  while (nextUrl && page < MAX_PAGES) {
    if (signal?.aborted) break;
    page++;

    // Re-read token before each page so a mid-export refresh picks up
    // the new value. Source depends on the chat service in use:
    //   - Teams Free: decrypt the cached SKYPE-TOKEN entry. The IC3
    //     token MUST NOT be substituted here — the chatsvc/consumer
    //     proxy 401s on it (it's a different audience).
    //   - Work/School: re-read the IC3 MSAL access token.
    let token = config.ic3Token;
    if (config.chatServiceUrl.includes('teams.live.com/api/chatsvc/consumer')) {
      const fresh = await getTeamsFreeChatAuth();
      if (fresh?.token) token = fresh.token;
    } else {
      token = (await getIc3Token()) || config.ic3Token;
    }

    const data = await fetchPageWithRetry(nextUrl, token, signal);

    if (data.errorCode) {
      throw new Error(`Messages API error ${data.errorCode}: ${data.message}`);
    }

    const messages = data.messages || [];
    if (messages.length === 0 && !data._metadata?.backwardLink) break;

    allMessages.push(...messages);

    onProgress?.({
      phase: 'fetch',
      page,
      messagesSoFar: allMessages.length,
    });

    // Early stop for date-bound exports: pagination walks newest → oldest, so
    // once any message in this page is older than startAtMs, all subsequent
    // pages are entirely older — no more in-range messages to fetch.
    if (Number.isFinite(startAtMs) && messages.some(m => {
      const ts = Date.parse(m.composetime || m.originalarrivaltime || '');
      return Number.isFinite(ts) && ts < startAtMs;
    })) {
      console.log(`[API] Reached startAt boundary at page ${page}, stopping pagination (${allMessages.length} messages fetched)`);
      break;
    }

    nextUrl = data._metadata?.backwardLink || null;

    if (nextUrl) {
      // Gradually increase delay for large conversations to avoid rate limiting
      if (page > 20) delayMs = Math.min(delayMs + 50, 500);
      await sleep(delayMs);
    }
  }

  return allMessages;
}

// ── High-Level Orchestrator ────────────────────────────────────────────

/**
 * Full API scrape pipeline: discover → extract conversation → fetch all messages.
 * Returns null if any step fails (caller should fall back to DOM scroll).
 */
// Side channel: the most recent apiScrape failure. The function still
// returns null on failure (preserves the simple "if (result)" call shape
// at the one call site), but the caller can read this AFTER seeing null
// to classify the failure. Set on every apiScrape entry; cleared on
// success. Used by content.ts to distinguish a network failure (mark
// the export as 'partial' with reason='network') from other API errors
// (token / conversation-id / etc., where DOM fallback is still useful).
type ApiScrapeFailure = {
  reason: 'no-skype-token' | 'no-messaging-token' | 'no-conv-id' | 'thrown';
  error?: unknown;
  isNetworkError: boolean;
};
let lastApiScrapeFailure: ApiScrapeFailure | null = null;
export function getLastApiScrapeFailure(): ApiScrapeFailure | null { return lastApiScrapeFailure; }

// Browsers don't have a reliable "is this a network error" flag — we
// pattern-match on the canonical error shapes Firefox and Chromium
// produce when fetch() can't reach the server (DNS fail, offline, TCP
// reset, certificate, etc.). Both throw TypeError; the message strings
// differ slightly per browser but are stable across versions.
function isNetworkError(err: unknown): boolean {
  if (!err || typeof err !== 'object') return false;
  if ((err as { name?: string }).name !== 'TypeError') return false;
  const msg = String((err as { message?: string }).message || '');
  return /NetworkError|Failed to fetch|Load failed|networkerror/i.test(msg);
}

export async function apiScrape(
  onProgress?: (p: FetchProgress) => void,
  options?: { signal?: AbortSignal; startAtISO?: string | null; conversationId?: string | null },
): Promise<{ messages: TeamsApiMessage[]; conversationId: string; userId: string | null; userRegion: string; ic3Token: string } | null> {
  const signal = options?.signal;
  const startAtISO = options?.startAtISO ?? null;
  const explicitConvId = options?.conversationId ?? null;
  // Reset the side-channel; success leaves this null so a stale failure
  // from a previous call doesn't leak into the next.
  lastApiScrapeFailure = null;
  try {
    // Step 1+2: Discover chat service & get a messaging token.
    onProgress?.({ phase: 'discover', page: 0, messagesSoFar: 0 });

    // Teams Free fast path: the consumer surface caches a fresh,
    // already-decryptable Skype JWT and the chat-service URL together
    // in localStorage. Reading it skips an authz round-trip AND avoids
    // the trap that the live authz response's `skypeToken.skypetoken`
    // is a different (longer, wrong-audience) token than the
    // proxy-acceptable one stored in the cache.
    let chatServiceUrl: string;
    let userRegion: string;
    let messagingAuthToken: string;
    const isTeamsFree = location.hostname.toLowerCase().includes('teams.live.com');
    const cachedTeamsFree = isTeamsFree ? await getTeamsFreeChatAuth() : null;
    if (cachedTeamsFree?.token && cachedTeamsFree.chatServiceUrl) {
      chatServiceUrl = cachedTeamsFree.chatServiceUrl;
      userRegion = cachedTeamsFree.userRegion || '';
      messagingAuthToken = cachedTeamsFree.token;
      console.log(`[API:tfl] auth from cache (chatService: ${chatServiceUrl})`);
    } else {
      // Work/School (and Teams Free fallback when cache is stale or
      // missing): authenticate against the discover endpoint.
      const skypeToken = await getSkypeToken();
      if (!skypeToken) {
        console.log('[API] No valid Skype token found, falling back to DOM');
        lastApiScrapeFailure = { reason: 'no-skype-token', isNetworkError: false };
        return null;
      }
      const discovered = await discover(skypeToken);
      chatServiceUrl = discovered.chatServiceUrl;
      userRegion = discovered.userRegion;
      // Token-source priority:
      //  1. authz-response messagingToken (some tenants return one)
      //  2. IC3 MSAL access token (standard Work/School path)
      const fallback = discovered.messagingToken || (await getIc3Token()) || undefined;
      if (!fallback) {
        console.log('[API] No valid messaging token found, falling back to DOM');
        lastApiScrapeFailure = { reason: 'no-messaging-token', isNetworkError: false };
        return null;
      }
      messagingAuthToken = fallback;
      console.log(`[API] Discovered chat service: ${chatServiceUrl} (region: ${userRegion})`);
    }

    // Step 3: Conversation ID — prefer the one the popup picker handed us.
    // Fall back to DOM/IDB extraction only when the caller didn't supply
    // one (legacy path kept so in-page automation / older callers still
    // work). The picker flow is what fixes the stale-cache / chat-switch
    // class of bugs — the user tells us which chat, no guessing.
    const conversationId = explicitConvId || (await extractConversationId());
    if (!conversationId) {
      console.log('[API] Could not extract conversation ID, falling back to DOM');
      lastApiScrapeFailure = { reason: 'no-conv-id', isNetworkError: false };
      return null;
    }
    console.log(`[API] Conversation: ${conversationId.substring(0, 30)}...`);

    // Step 4: Fetch all messages
    const config: TeamsApiConfig = { chatServiceUrl, userRegion, ic3Token: messagingAuthToken };
    const messages = await fetchAllMessages(config, conversationId, onProgress, signal, startAtISO);
    console.log(`[API] Fetched ${messages.length} messages`);

    // Step 5: Resolve unresolved MRIs (forwarded senders, reactors) via Graph API
    const graphToken = await getGraphToken();
    if (graphToken) {
      const unresolved = collectUnresolvedMris(messages);
      if (unresolved.length > 0) {
        console.log(`[API] Resolving ${unresolved.length} unknown user MRIs via Graph API...`);
        const resolved = await resolveUserNames(unresolved, graphToken);
        if (resolved.size > 0) {
          // Inject resolved names back into messages
          for (const msg of messages) {
            // Fill in missing imdisplayname for forwarded messages
            if (!msg.imdisplayname && msg.from) {
              const name = resolved.get(msg.from);
              if (name) msg.imdisplayname = name;
            }
          }
          // Store resolved map on messages for the converter to use
          (messages as unknown as { __resolvedMris: Map<string, string> }).__resolvedMris = resolved;
          console.log(`[API] Resolved ${resolved.size} user names`);
        }
      }
    }

    // Teams Free: layer the local profiles cache on top of any Graph-
    // resolved names. Skype consumer IDs (`8:live:<id>`, `8:<id>`)
    // can't be resolved via Graph (different identity system), and their
    // names live only in this IDB store. Empty map on Work/School (no DB
    // matches the pattern there) — costs ~one IDB roundtrip in the worst
    // case. Merge order: existing __resolvedMris wins on conflict (Graph
    // is the more authoritative source when it has a record at all).
    if (isTeamsFree) {
      const localProfiles = await loadTeamsFreeProfiles();
      if (localProfiles.size > 0) {
        const existing = (messages as unknown as { __resolvedMris?: Map<string, string> }).__resolvedMris
          || new Map<string, string>();
        for (const [k, v] of localProfiles) {
          if (!existing.has(k)) existing.set(k, v);
        }
        (messages as unknown as { __resolvedMris: Map<string, string> }).__resolvedMris = existing;
        console.log(`[API:tfl] layered ${localProfiles.size} profile names`);
      }
    }

    const userId = getCurrentUserUuid();
    return { messages, conversationId, userId, userRegion, ic3Token: messagingAuthToken };
  } catch (err: any) {
    // Re-throw an abort so the caller's try/catch can branch on cancellation
    // instead of treating it as a normal API failure (and triggering DOM
    // fallback or noisy logs).
    if (signal?.aborted || err?.name === 'AbortError') throw err;
    lastApiScrapeFailure = { reason: 'thrown', error: err, isNetworkError: isNetworkError(err) };
    console.log('[API] API scrape failed:', err, lastApiScrapeFailure.isNetworkError ? '(NETWORK ERROR)' : '');
    return null;
  }
}
