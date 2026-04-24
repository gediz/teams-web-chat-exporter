/**
 * Teams API Client
 *
 * Fetches messages directly from the Teams chat service API.
 * Uses MSAL tokens from localStorage and service discovery via authz endpoint.
 * Falls back gracefully — any failure should trigger DOM scroll fallback.
 */

/* eslint-disable no-console */

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
  return await findValidToken('ic3.teams.office.com')
    || await findValidToken('ic3.teams.office365.us')
    || await findValidToken('chatsvcagg');
}

/** Get a valid Skype API token for the authz discovery endpoint. */
export async function getSkypeToken(): Promise<string | null> {
  return findValidToken('api.spaces.skype');
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
  const teamsScopes = ['ic3.teams.office.com', 'api.spaces.skype', 'graph.microsoft.com'];
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

  // Resolve each UUID via Graph (sequentially to avoid rate limits)
  let firstError: string | null = null;
  for (const [uuid, originalMris] of uuidToMris) {
    try {
      const resp = await fetch(
        `https://graph.microsoft.com/v1.0/users/${uuid}?$select=displayName`,
        {
          headers: { 'Authorization': `Bearer ${graphToken}` },
          signal: AbortSignal.timeout(5_000),
        },
      );
      if (!resp.ok) {
        if (!firstError) firstError = `HTTP ${resp.status} ${resp.statusText}`;
        continue;
      }
      const data = await resp.json();
      const name = data.displayName;
      if (name) {
        for (const mri of originalMris) result.set(mri, name);
        // Also store the short MRI forms
        result.set(`8:orgid:${uuid}`, name);
        result.set(`gid:${uuid}`, name);
      }
    } catch (e) {
      if (!firstError) firstError = String(e);
    }
  }
  if (firstError && result.size === 0) {
    console.warn(`[API] Graph user resolution failed for all ${uuidToMris.size} users — first error: ${firstError}`);
  }

  return result;
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

/**
 * Detect the authz service base URL from the current Teams domain.
 * GCC High uses teams.microsoft.us, all others use teams.microsoft.com.
 */
function getAuthzUrl(): string {
  const host = location.hostname.toLowerCase();
  if (host.includes('teams.microsoft.us')) {
    return 'https://authsvc.teams.microsoft.us/v1.0/authz';
  }
  return 'https://authsvc.teams.microsoft.com/v1.0/authz';
}

/**
 * Discover the chat service URL and region via the authz endpoint.
 * This is the official Teams service discovery mechanism.
 */
export async function discover(skypeToken: string): Promise<{ chatServiceUrl: string; userRegion: string }> {
  const resp = await fetch(getAuthzUrl(), {
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
  const chatServiceUrl = data.regionGtms?.chatService;
  const userRegion = data.userRegion || data.region || '';

  if (!chatServiceUrl || typeof chatServiceUrl !== 'string') {
    throw new Error('authz response missing chatService URL');
  }

  return { chatServiceUrl, userRegion };
}

// ── Conversation ID Extraction ─────────────────────────────────────────

/**
 * Extract the current conversation ID by looking up a visible message's
 * data-mid in IndexedDB's replychain-manager. This is the only reliable
 * method in Teams v2 — the URL doesn't contain the conversation ID and
 * sidebar focus state doesn't always match the visible conversation.
 */
export async function extractConversationId(): Promise<string | null> {
  // Method 1 (primary): Look up visible data-mid values in IndexedDB
  let midsCount = 0;
  try {
    const mids = Array.from(document.querySelectorAll('[data-mid]'))
      .map(el => el.getAttribute('data-mid'))
      .filter(Boolean) as string[];
    midsCount = mids.length;

    if (mids.length > 0) {
      const convId = await lookupConversationInIdb(mids);
      if (convId) return convId;
    }
  } catch (e) {
    console.log('[API] IDB conversation lookup failed:', e);
  }

  // Method 2: URL query parameter or hash
  const CONV_ID_RE = /19:[^|"}\s&?/]+@(?:unq\.gbl\.spaces|thread\.[a-z0-9]+)/i;
  try {
    const url = new URL(window.location.href);
    for (const [, value] of url.searchParams) {
      if (CONV_ID_RE.test(value)) return value.match(CONV_ID_RE)![0];
    }
    const hashMatch = window.location.hash.match(CONV_ID_RE);
    if (hashMatch) return decodeURIComponent(hashMatch[0]);
  } catch { /* ignore */ }

  // Method 3: DOM data attributes on chat/message pane
  const chatPane = document.querySelector('[data-tid="message-pane"]');
  const threadId = chatPane?.getAttribute('data-convid') ||
                   chatPane?.getAttribute('data-tid-convid') ||
                   chatPane?.getAttribute('data-conversation-id');
  if (threadId) return threadId;

  // All three methods exhausted. Log the state so the next failure is
  // diagnosable without the user having to reproduce under devtools.
  console.log('[API] extractConversationId exhausted all methods.', {
    visibleDataMids: midsCount,
    hasMessagePane: !!chatPane,
    url: window.location.href,
    hash: window.location.hash,
  });
  return null;
}

/** Look up which conversation owns the given message IDs via IndexedDB. */
async function lookupConversationInIdb(mids: string[]): Promise<string | null> {
  // indexedDB.databases() is not available in Firefox < 126
  if (typeof indexedDB.databases !== 'function') {
    console.log('[API] indexedDB.databases() not available');
    return null;
  }
  const databases = await indexedDB.databases();
  // The DB name carries tenant/user/locale suffixes; match the manager prefix and
  // exclude streams-replychain-manager which is a different store.
  const rcDb = databases.find(d =>
    d.name?.includes('replychain-manager:react-web-client') && !d.name.includes('streams-')
  );
  if (!rcDb?.name) {
    console.log('[API] replychain-manager DB not found; candidates:',
      databases.map(d => d.name).filter(Boolean));
    return null;
  }

  return new Promise((resolve) => {
    const req = indexedDB.open(rcDb.name!);
    req.onerror = () => { console.log('[API] IDB open onerror for', rcDb.name); resolve(null); };
    req.onblocked = () => { console.log('[API] IDB open onblocked for', rcDb.name); resolve(null); };
    req.onsuccess = (e) => {
      const db = (e.target as IDBOpenDBRequest).result;
      try {
        const tx = db.transaction('replychains', 'readonly');
        const store = tx.objectStore('replychains');
        const getAll = store.getAll();
        getAll.onsuccess = () => {
          const midSet = new Set(mids);
          type MessageEntry = { id?: string };
          const records: Array<{
            conversationId?: string;
            replyChainId?: string;
            messageMap?: Record<string, MessageEntry>;
          }> = getAll.result;

          // Pass 1: top-level replyChainId matches a visible mid (parent message).
          for (const rec of records) {
            if (rec.conversationId && rec.replyChainId && midSet.has(rec.replyChainId)) {
              db.close();
              resolve(rec.conversationId);
              return;
            }
          }

          // Pass 2: messageMap entry's `id` matches (any message in the chain).
          // In the new Teams, messageMap is keyed by "8:orgid:..._<clientMessageId>"
          // which is unrelated to the visible data-mid; the actual mid is in v.id.
          for (const rec of records) {
            if (!rec.conversationId || !rec.messageMap) continue;
            for (const v of Object.values(rec.messageMap)) {
              if (v?.id && midSet.has(v.id)) {
                db.close();
                resolve(rec.conversationId);
                return;
              }
            }
          }

          // Pass 3 (legacy fallback): older Teams stored mids directly inside
          // the messageMap key. Keep this so classic Teams keeps working.
          for (const rec of records) {
            if (!rec.conversationId || !rec.messageMap) continue;
            for (const key of Object.keys(rec.messageMap)) {
              for (const mid of mids) {
                if (key.includes(mid)) {
                  db.close();
                  resolve(rec.conversationId);
                  return;
                }
              }
            }
          }

          console.log('[API] IDB lookup: no match across',
            records.length, 'records for',
            mids.length, 'visible data-mid(s); sample mid:', mids[0]);
          db.close();
          resolve(null);
        };
        getAll.onerror = () => {
          console.log('[API] IDB getAll error for replychains store');
          db.close();
          resolve(null);
        };
      } catch {
        db.close();
        resolve(null);
      }
    };
  });
}

// ── Paginated Message Fetcher ──────────────────────────────────────────

const INITIAL_DELAY_MS = 150;
const MAX_PAGES = 500; // safety limit (~100k messages)
const MAX_RETRIES = 5;

const sleep = (ms: number) => new Promise(resolve => setTimeout(resolve, ms));

/**
 * Fetch a single page with retry logic for 429 rate limiting.
 * Returns the parsed response or throws on non-retryable errors.
 */
async function fetchPageWithRetry(
  url: string,
  token: string,
  signal?: AbortSignal,
): Promise<MessagesResponse> {
  for (let attempt = 0; attempt <= MAX_RETRIES; attempt++) {
    const resp = await fetch(url, {
      headers: { 'Authorization': `Bearer ${token}` },
      signal: signal || AbortSignal.timeout(30_000),
    });

    if (resp.status === 429) {
      if (attempt >= MAX_RETRIES) {
        throw new Error('Rate limited after max retries');
      }
      // Use Retry-After header if available, otherwise exponential backoff
      const retryAfter = resp.headers.get('Retry-After');
      const waitMs = retryAfter
        ? parseInt(retryAfter, 10) * 1000
        : Math.min(2000 * Math.pow(2, attempt), 30_000); // 2s, 4s, 8s, 16s, 30s
      console.log(`[API] Rate limited (429), retrying in ${waitMs}ms (attempt ${attempt + 1}/${MAX_RETRIES})`);
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

    // Re-read token before each page (MSAL may have refreshed it)
    const token = (await getIc3Token()) || config.ic3Token;

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
export async function apiScrape(
  onProgress?: (p: FetchProgress) => void,
  options?: { signal?: AbortSignal; startAtISO?: string | null },
): Promise<{ messages: TeamsApiMessage[]; conversationId: string; userId: string | null; userRegion: string; ic3Token: string } | null> {
  const signal = options?.signal;
  const startAtISO = options?.startAtISO ?? null;
  try {
    // Step 1: Get tokens
    const skypeToken = await getSkypeToken();
    if (!skypeToken) {
      console.log('[API] No valid Skype token found, falling back to DOM');
      return null;
    }
    const ic3Token = await getIc3Token();
    if (!ic3Token) {
      console.log('[API] No valid IC3 token found, falling back to DOM');
      return null;
    }

    // Step 2: Discover chat service
    onProgress?.({ phase: 'discover', page: 0, messagesSoFar: 0 });
    const { chatServiceUrl, userRegion } = await discover(skypeToken);
    console.log(`[API] Discovered chat service: ${chatServiceUrl} (region: ${userRegion})`);

    // Step 3: Extract conversation ID
    const conversationId = await extractConversationId();
    if (!conversationId) {
      console.log('[API] Could not extract conversation ID, falling back to DOM');
      return null;
    }
    console.log(`[API] Conversation: ${conversationId.substring(0, 30)}...`);

    // Step 4: Fetch all messages
    const config: TeamsApiConfig = { chatServiceUrl, userRegion, ic3Token };
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

    const userId = getCurrentUserUuid();
    return { messages, conversationId, userId, userRegion, ic3Token };
  } catch (err: any) {
    // Re-throw an abort so the caller's try/catch can branch on cancellation
    // instead of treating it as a normal API failure (and triggering DOM
    // fallback or noisy logs).
    if (signal?.aborted || err?.name === 'AbortError') throw err;
    console.log('[API] API scrape failed, falling back to DOM:', err);
    return null;
  }
}
