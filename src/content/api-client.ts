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

/** Find a valid (non-expired) MSAL access token matching the scope pattern. */
function findValidToken(scopePattern: string): string | null {
  const now = Math.floor(Date.now() / 1000);
  for (let i = 0; i < localStorage.length; i++) {
    const key = localStorage.key(i);
    if (!key || !key.includes('accesstoken') || !key.includes(scopePattern)) continue;
    try {
      const entry = JSON.parse(localStorage.getItem(key) || '');
      if (Number(entry.expiresOn) > now && entry.secret) {
        return entry.secret;
      }
    } catch { /* skip malformed entries */ }
  }
  return null;
}

/** Get a valid ic3 token for the message API. Re-reads localStorage each call. */
export function getIc3Token(): string | null {
  // Standard commercial: ic3.teams.office.com
  // GCC High: ic3.teams.office365.us or chatsvcagg
  return findValidToken('ic3.teams.office.com')
    || findValidToken('ic3.teams.office365.us')
    || findValidToken('chatsvcagg');
}

/** Get a valid Skype API token for the authz discovery endpoint. */
export function getSkypeToken(): string | null {
  return findValidToken('api.spaces.skype');
}

/** Get a valid Graph API token for user resolution. */
export function getGraphToken(): string | null {
  // Standard: graph.microsoft.com, GCC High: graph.microsoft.us
  return findValidToken('graph.microsoft.com')
    || findValidToken('graph.microsoft.us');
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
  for (const [uuid, originalMris] of uuidToMris) {
    try {
      const resp = await fetch(
        `https://graph.microsoft.com/v1.0/users/${uuid}?$select=displayName`,
        {
          headers: { 'Authorization': `Bearer ${graphToken}` },
          signal: AbortSignal.timeout(5_000),
        },
      );
      if (!resp.ok) continue;
      const data = await resp.json();
      const name = data.displayName;
      if (name) {
        for (const mri of originalMris) result.set(mri, name);
        // Also store the short MRI forms
        result.set(`8:orgid:${uuid}`, name);
        result.set(`gid:${uuid}`, name);
      }
    } catch { /* skip unresolvable */ }
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
  try {
    const mids = Array.from(document.querySelectorAll('[data-mid]'))
      .map(el => el.getAttribute('data-mid'))
      .filter(Boolean) as string[];

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

  return null;
}

/** Look up which conversation owns the given message IDs via IndexedDB. */
async function lookupConversationInIdb(mids: string[]): Promise<string | null> {
  // indexedDB.databases() is not available in Firefox < 126
  if (typeof indexedDB.databases !== 'function') return null;
  const databases = await indexedDB.databases();
  const rcDb = databases.find(d => d.name?.includes('replychain-manager:react-web-client'));
  if (!rcDb?.name) return null;

  return new Promise((resolve) => {
    const req = indexedDB.open(rcDb.name!);
    req.onerror = () => resolve(null);
    req.onsuccess = (e) => {
      const db = (e.target as IDBOpenDBRequest).result;
      try {
        const tx = db.transaction('replychains', 'readonly');
        const store = tx.objectStore('replychains');
        const getAll = store.getAll();
        getAll.onsuccess = () => {
          const midSet = new Set(mids);
          const records: Array<{ conversationId: string; replyChainId: string; messageMap: Record<string, unknown> }> = getAll.result;

          // Find conversations containing any of our visible message IDs
          for (const rec of records) {
            if (midSet.has(rec.replyChainId)) {
              db.close();
              resolve(rec.conversationId);
              return;
            }
          }

          // Fallback: check inside messageMap keys (they contain the mid)
          for (const rec of records) {
            for (const key of Object.keys(rec.messageMap || {})) {
              for (const mid of mids) {
                if (key.includes(mid)) {
                  db.close();
                  resolve(rec.conversationId);
                  return;
                }
              }
            }
          }

          db.close();
          resolve(null);
        };
        getAll.onerror = () => { db.close(); resolve(null); };
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
 * Paginates via backwardLink until history is exhausted.
 * Re-reads the ic3 token from localStorage before each page to handle refresh.
 * Retries on 429 with exponential backoff.
 */
export async function fetchAllMessages(
  config: TeamsApiConfig,
  conversationId: string,
  onProgress?: (p: FetchProgress) => void,
  signal?: AbortSignal,
): Promise<TeamsApiMessage[]> {
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
    const token = getIc3Token() || config.ic3Token;

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
  signal?: AbortSignal,
): Promise<{ messages: TeamsApiMessage[]; conversationId: string } | null> {
  try {
    // Step 1: Get tokens
    const skypeToken = getSkypeToken();
    if (!skypeToken) {
      console.log('[API] No valid Skype token found, falling back to DOM');
      return null;
    }
    const ic3Token = getIc3Token();
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
    const messages = await fetchAllMessages(config, conversationId, onProgress, signal);
    console.log(`[API] Fetched ${messages.length} messages`);

    // Step 5: Resolve unresolved MRIs (forwarded senders, reactors) via Graph API
    const graphToken = getGraphToken();
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

    return { messages, conversationId };
  } catch (err) {
    console.log('[API] API scrape failed, falling back to DOM:', err);
    return null;
  }
}
