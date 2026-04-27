// Direct readers for Teams web's client-side state — IndexedDB stores
// and sessionStorage navigation entries. Replaces our previous
// network-fetch + DOM-heuristic stack with deterministic local reads.
//
// Background and full schema in docs/TEAMS_INTERNALS.md.
//
// All readers are best-effort and return null/[] on failure rather
// than throwing. The picker is the only consumer; falling back to
// previous heuristics when something here returns empty keeps the
// UX functional during Teams-internal reshapes.

// ── Public types ─────────────────────────────────────────────────────

/** One row from the `conversations` store of conversation-manager. The
 *  shape mirrors what Teams writes; we've typed only the fields we use. */
export type TeamsConversationRecord = {
  id: string;
  type?: string;
  lastMessageTimeUtc?: number;
  /** Teams' own pre-computed display title — the exact string that
   *  appears in the sidebar. Populated for unnamed groups and 1:1
   *  chats (where threadProperties.topic is empty/null). When set,
   *  this should be the name source — no further enrichment needed.
   *  shortTitle is the compact "Alice, Bob, Carol, +N" form; longer
   *  variants may also exist depending on Teams version. */
  chatTitle?: {
    shortTitle?: string;
    [k: string]: unknown;
  };
  threadProperties?: {
    topic?: string;
    threadType?: string;
    productContext?: string;
    hidden?: boolean;
    isRead?: boolean;
    isStickyThread?: boolean;
    creator?: string;
    groupId?: string;
    [k: string]: unknown;
  };
  lastMessage?: {
    imdisplayname?: string;
    fromDisplayNameInToken?: string;
    composetime?: string;
    originalarrivaltime?: string;
    from?: string;
    [k: string]: unknown;
  };
  [k: string]: unknown;
};

/** One row from the `folders` store of conversation-folder-manager. */
export type TeamsFolderRecord = {
  id: string;
  name?: string;
  folderType?: string;
  conversations?: Array<{
    id: string;
    threadType?: string;
    itemType?: string;
    [k: string]: unknown;
  }>;
  isHidden?: boolean;
  isExpanded?: boolean;
  sortType?: string;
  version?: number;
  [k: string]: unknown;
};

// ── Active conversation (sessionStorage navHistory) ──────────────────

/** Read the active conversation id from Teams' in-tab navigation history.
 *  Returns null when the user is on a non-chat view (Activity, Calendar,
 *  etc.) or when the navigation state isn't readable yet (early in tab
 *  load). Locale-independent — never touches DOM, never parses titles. */
export function readActiveConversationId(selfUuid: string): string | null {
  if (!selfUuid) return null;
  try {
    const histKey  = `tmp.session.${selfUuid}-mainWindowNavHistory`;
    const indexKey = `tmp.session.${selfUuid}-mainWindowNavHistoryIndex`;
    const histRaw  = sessionStorage.getItem(histKey);
    const indexRaw = sessionStorage.getItem(indexKey);
    if (!histRaw || !indexRaw) return null;
    const history = JSON.parse(histRaw) as Array<{ activeEntities?: { mainEntity?: { id?: string; type?: string; action?: string } } }>;
    const index   = JSON.parse(indexRaw) as { windowHistoryIndex?: number };
    const i = typeof index.windowHistoryIndex === 'number' ? index.windowHistoryIndex : 0;
    const entry = history[i]?.activeEntities?.mainEntity;
    if (!entry || entry.action !== 'view' || typeof entry.id !== 'string') return null;
    // Accept any entity type that looks chat-like. The discriminator is
    // 'chats' for normal chats and meeting-derived chats; channels use
    // 'channels'. We let any of the chat-shaped types through and rely
    // on the conversation list for downstream validation.
    if (entry.type !== 'chats' && entry.type !== 'channels' && entry.type !== 'meetings') return null;
    return entry.id;
  } catch { return null; }
}

// ── IndexedDB infrastructure ─────────────────────────────────────────

/** Find Teams databases whose name starts with `Teams:<prefix>:`. The
 *  full name embeds tenant, user, and locale; multiple may exist for
 *  the same prefix (e.g. en-gb + en-us in parallel). */
async function findTeamsDbs(prefix: string): Promise<{ name: string; version: number }[]> {
  if (typeof indexedDB.databases !== 'function') return [];
  try {
    const all = await indexedDB.databases();
    return all
      .filter((d): d is { name: string; version: number } =>
        !!d.name && d.name.startsWith(`Teams:${prefix}:`) && typeof d.version === 'number')
      .sort((a, b) => a.name.localeCompare(b.name));
  } catch { return []; }
}

/** Open a database read-only. Returns null on any failure. */
function openDbRO(name: string): Promise<IDBDatabase | null> {
  return new Promise((resolve) => {
    try {
      const req = indexedDB.open(name);
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => resolve(null);
      req.onblocked = () => resolve(null);
    } catch { resolve(null); }
  });
}

/** Drain an entire object store. Returns [] on any failure. */
function readAll<T>(db: IDBDatabase, storeName: string): Promise<T[]> {
  return new Promise((resolve) => {
    try {
      if (!db.objectStoreNames.contains(storeName)) { resolve([]); return; }
      const tx = db.transaction(storeName, 'readonly');
      const req = tx.objectStore(storeName).getAll();
      req.onsuccess = () => resolve(req.result as T[]);
      req.onerror = () => resolve([]);
    } catch { resolve([]); }
  });
}

// ── Conversation list ────────────────────────────────────────────────

/** Read the user's full conversation list from
 *  Teams:conversation-manager > conversations. Merges all locale DBs
 *  by id (same locale-parallelism rationale as readReplychainSenders). */
export async function readConversationList(): Promise<TeamsConversationRecord[]> {
  const dbs = await findTeamsDbs('conversation-manager');
  if (dbs.length === 0) return [];
  const merged = new Map<string, TeamsConversationRecord>();
  for (const meta of dbs) {
    const db = await openDbRO(meta.name);
    if (!db) continue;
    try {
      const rows = await readAll<TeamsConversationRecord>(db, 'conversations');
      for (const r of rows) {
        if (!r.id) continue;
        const prev = merged.get(r.id);
        // Pick the freshest record per id — different locale DBs may
        // be stale relative to each other.
        if (!prev || (r.lastMessageTimeUtc || 0) > (prev.lastMessageTimeUtc || 0)) {
          merged.set(r.id, r);
        }
      }
    } finally {
      try { db.close(); } catch { /* ignore */ }
    }
  }
  return [...merged.values()];
}

// ── Contacts (capiv3) ────────────────────────────────────────────────

/** One row from the capiv3 contacts cache. We only type the fields we
 *  read; Teams writes a much larger record. */
export type TeamsContactRecord = {
  mri: string;
  defaultEmail?: string;
  emails?: Array<{ address?: string; displayName?: string; type?: string }>;
  name?: { displayName?: string };
  [k: string]: unknown;
};

/** Read locally-cached contacts. Teams stores users you've explicitly
 *  contacted (favourites, recent senders) here with their display
 *  names. Useful as a fallback when Graph can't resolve a federated
 *  user — capiv3 caches the name Teams was told about cross-tenant.
 *  Merges all locale DBs by mri. */
export async function readContacts(): Promise<TeamsContactRecord[]> {
  const dbs = await findTeamsDbs('capiv3-contacts-manager');
  if (dbs.length === 0) return [];
  const merged = new Map<string, TeamsContactRecord>();
  for (const meta of dbs) {
    const db = await openDbRO(meta.name);
    if (!db) continue;
    try {
      const rows = await readAll<TeamsContactRecord>(db, 'capiv3-contacts');
      for (const r of rows) {
        if (r.mri && !merged.has(r.mri)) merged.set(r.mri, r);
      }
    } finally {
      try { db.close(); } catch { /* ignore */ }
    }
  }
  return [...merged.values()];
}

/** Build a map of mri → best-available display name from the contacts
 *  cache. Falls back through name fields in priority order. */
export function buildContactNameMap(contacts: TeamsContactRecord[]): Map<string, string> {
  const out = new Map<string, string>();
  for (const c of contacts) {
    if (!c.mri) continue;
    const name =
      c.name?.displayName?.trim()
      || c.emails?.find(e => e.displayName?.trim())?.displayName?.trim()
      || (c.defaultEmail ? c.defaultEmail.split('@')[0] : undefined);
    if (name) out.set(c.mri, name);
  }
  return out;
}

// ── Message-sender names (replychain-manager) ────────────────────────

/** One sender entry surfaced from replychain-manager messages: who
 *  sent something in a conversation, what their cached display name
 *  was, and when (so we can rank by recency). */
export type SenderEntry = {
  mri: string;
  name: string;
  lastSeen: number;          // arrival-time ms
};

/** Read message senders from Teams' replychain-manager IDB store and
 *  group them per-conversation, deduped on MRI with last-seen-wins
 *  recency. The chain stores cached display names from every message
 *  Teams has received — including federated contacts whose names
 *  aren't in Graph or capiv3. This is the only durable local source
 *  for external user names.
 *
 *  Reads from EVERY locale-suffixed replychain-manager DB and merges
 *  the results. Teams keeps :en-gb / :en-us / etc. databases in
 *  parallel and writes to whichever the active session is using;
 *  reading from one locale alone misses messages stored under others.
 *
 *  Returns: Map<conversationId, SenderEntry[]> where each list is
 *  sorted recency-first.
 */
export async function readReplychainSenders(): Promise<Map<string, SenderEntry[]>> {
  const dbs = await findTeamsDbs('replychain-manager');
  if (dbs.length === 0) return new Map();

  const perConv = new Map<string, Map<string, SenderEntry>>();
  for (const meta of dbs) {
    await scanOneReplychainDb(meta.name, perConv);
  }

  const out = new Map<string, SenderEntry[]>();
  for (const [convId, bucket] of perConv) {
    out.set(convId, [...bucket.values()].sort((a, b) => b.lastSeen - a.lastSeen));
  }
  return out;
}

async function scanOneReplychainDb(name: string, perConv: Map<string, Map<string, SenderEntry>>): Promise<void> {
  const db = await openDbRO(name);
  if (!db) return;
  try {
    const storeName = db.objectStoreNames.contains('replychains')
      ? 'replychains'
      : db.objectStoreNames[0];
    if (!storeName) return;
    await new Promise<void>((resolve) => {
      const tx = db.transaction(storeName, 'readonly');
      const store = tx.objectStore(storeName);
      const req = store.openCursor();
      req.onsuccess = () => {
        const cursor = req.result;
        if (!cursor) { resolve(); return; }
        const rec = cursor.value as { messageMap?: Record<string, any> };
        if (rec?.messageMap) {
          for (const m of Object.values(rec.messageMap)) {
            const convId: string | undefined = m?.conversationId;
            const mri: string | undefined = m?.from || m?.creator;
            const display: string | undefined =
              m?.imDisplayName?.trim() || m?.fromDisplayNameInToken?.trim();
            const ts: number = Number(m?.originalArrivalTime || m?.clientArrivalTime || 0);
            if (!convId || !mri || !display) continue;
            // 'creator' may be the conversation thread itself (e.g.
            // "19:...@thread.v2") for system messages — skip those.
            if (mri.startsWith('19:')) continue;
            let bucket = perConv.get(convId);
            if (!bucket) { bucket = new Map(); perConv.set(convId, bucket); }
            const prev = bucket.get(mri);
            if (!prev || prev.lastSeen < ts) bucket.set(mri, { mri, name: display, lastSeen: ts });
          }
        }
        cursor.continue();
      };
      req.onerror = () => resolve();
    });
  } finally {
    try { db.close(); } catch { /* ignore */ }
  }
}

// ── Folders ──────────────────────────────────────────────────────────

/** Read folder definitions from
 *  Teams:conversation-folder-manager > folders. Merges all locale DBs
 *  by id (newest version wins) — same locale-parallelism rationale
 *  as readReplychainSenders. */
export async function readFolders(): Promise<TeamsFolderRecord[]> {
  const dbs = await findTeamsDbs('conversation-folder-manager');
  if (dbs.length === 0) return [];
  const merged = new Map<string, TeamsFolderRecord>();
  for (const meta of dbs) {
    const db = await openDbRO(meta.name);
    if (!db) continue;
    try {
      const rows = await readAll<TeamsFolderRecord>(db, 'folders');
      for (const r of rows) {
        if (!r.id) continue;
        const prev = merged.get(r.id);
        if (!prev || (r.version || 0) > (prev.version || 0)) {
          merged.set(r.id, r);
        }
      }
    } finally {
      try { db.close(); } catch { /* ignore */ }
    }
  }
  return [...merged.values()];
}
