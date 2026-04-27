# Teams web — internal data sources

Where Teams' web client stores its conversation state and which sources
this extension reads. Recorded so a future session doesn't redo the
research.

## TL;DR

Teams web is a heavily client-cached PWA. Conversation list, folders,
active-chat state, and pre-formatted display titles **never travel
over the wire** in normal use — they live in:

- **IndexedDB** (`Teams:*` databases on `teams.cloud.microsoft` origin)
- **sessionStorage** (`tmp.session.*` keys for the active chat)
- **Service Worker** caches (opaque to MAIN-world fetch interception)

Network capture (fetch / XHR / WebSocket) reveals tokens + infrastructure
traffic but not chat-state data. We read directly from IDB +
sessionStorage in the page's origin via our content script.

## Don't go down these rabbit holes again

- **MAIN-world fetch / XHR / WebSocket interception** — page-context
  reads served by the Service Worker never reach the network, so they
  never reach our patches. Useful only for confirming Trouter shape.
- **Trouter** (`wss://*.trouter.teams.microsoft.com/v4/c`) — Socket.IO
  frame format, well documented in `EionRobb/purple-teams`. Carries
  delta sync (new messages from others, presence) **only**. Active-chat
  is not pushed; nor is folder mutation.
- **OneGQL / Relay store** — Teams' internal GraphQL layer keeps its
  cache in-memory (Relay store). Not persisted to IndexedDB. Only
  reachable via a MAIN-world Relay tap, which we declined.
- **Microsoft Graph** for federated users — returns 404; their tenant
  is opaque. Cache the negative result (we do).
- **`/v1/threads/{id}/members`** for federated members — returns the
  MRI but **not** `friendlyName`. Useful for member counts only.

## Reader sources, in priority order

### 1. Active conversation ID — sessionStorage

Reliable, locale-independent, in-tab navigation state.

```
key:  tmp.session.<selfUuid>-mainWindowNavHistory       (JSON array)
key:  tmp.session.<selfUuid>-mainWindowNavHistoryIndex  (JSON, { windowHistoryIndex })
path: history[i].activeEntities.mainEntity
        .type === "chats" | "meetings" | "channels" | "simplecollab"
        .action === "view"
        .id    === the conversation id
```

Updates synchronously when the user navigates. Returns null on non-chat
views (Activity, Calendar). No DOM, no race.

### 2. Conversation list — IndexedDB (the source of truth)

```
DB:    Teams:conversation-manager:react-web-client:<tenant>:<userUuid>:<locale>
Store: conversations
Key:   id  (e.g. "19:xxxx@thread.v2", "48:notes")
```

**Critical fields:**

| Field | Notes |
|---|---|
| `id` | Conversation id |
| `lastMessageTimeUtc` | Sort key (recency-first) |
| **`chatTitle.shortTitle`** | **Teams' own pre-computed sidebar title.** Use first; covers unnamed groups + federated 1:1s in one shot |
| `threadProperties.topic` | Display name when explicitly set (named groups, meetings, channels) |
| `threadProperties.threadType` | `"chat"` / `"meeting"` / `"channel"` / `"streamofnotes"` |
| `threadProperties.groupId` | Channel parent-team id |
| `threadProperties.hidden` | **DO NOT FILTER** — Teams sets it true on active chats too |
| `threadProperties.isDeleted` | True if user deleted; **do** filter these |
| `lastMessage.imdisplayname` | Sender display name (last-resort fallback) |

**Ignore these `id`-prefix variants:**

- `48:notifications`, `48:mentions`, `48:calllogs`, `48:saved`,
  `48:starred`, `48:drafts` — system pseudo-chats; not exportable
- `48:notes` — keep this; it's the self-chat
- `streamofnotes` threadType — drop unless `id === 48:notes`

Counts on a long-running account: ~3× the `/v1/users/ME/conversations`
API list. The API filters server-side for recency / visibility; IDB has
the full local set.

### 3. Folders — IndexedDB

```
DB:    Teams:conversation-folder-manager:react-web-client:<tenant>:8:orgid:<userUuid>:<locale>
Store: folders
Key:   id  ("<tenant>~<user>~Favorites" for user folders, etc.)
```

| Field | Notes |
|---|---|
| `name` | `Favorites`, `MeetingChats`, `MutedChats`, `QuickViews`, `RecentChats`, `TeamsAndChannels`, custom names |
| `folderType` | `Favorites` / `UserDefined` / etc. |
| `conversations[]` | `[{id, threadType, ...}]` — member chat ids |
| `isHidden`, `isExpanded`, `sortType` | UX hints |

`QuickViews` contains slice ids (`slice-activities-mentions`, …) that
aren't real conversations — exclude from picker.

### 4. External / federated 1:1 names — fallback chain

For chats where `chatTitle.shortTitle` is missing, the resolution chain
is:

1. **Microsoft Graph `/users/{uuid}`** — works for same-tenant users.
   Cache positive AND negative results module-scope to avoid 404 spam
   and re-querying on every popup open.
2. **`Teams:capiv3-contacts-manager > capiv3-contacts`** — Teams' own
   local contact cache. Limited to users you've explicitly contacted.
3. **`Teams:replychain-manager > replychains`** (cursor scan) — the
   only durable local source for federated user names. Each message
   carries the sender's `imDisplayName`. **Sender field is `creator`
   (not `from`)**; for system events `creator` may be the conversation
   id, so skip values starting with `19:`.
4. **`/v1/threads/{id}/members`** — returns `friendlyName` for
   in-tenant members only; useless for federated.
5. **`lastMessage.imdisplayname`** when sender wasn't self.
6. Placeholder + `composetime` date.

### 5. The locale-DB trap

**Teams keeps every IDB DB in parallel locale variants** (`:en-gb`,
`:en-us`, …). Different sessions / login flows write to different
ones. Reading a single locale DB will silently miss data the other
holds.

Every reader merges across all matching DBs:

```ts
const dbs = await findTeamsDbs('conversation-manager');
for (const meta of dbs) {
  const db = await openDbRO(meta.name);
  // …merge into a single Map<id, record>, newest version wins
}
```

This bit us hard during the federated-name work — `:en-gb` had the
messages, our reader picked `:en-us`, replychain came back empty.

## Filter rules summary

For the picker:

| Drop when | Keep when |
|---|---|
| `id !== '48:notes'` AND `id.startsWith('48:')` | otherwise |
| `threadProperties.threadType === 'streamofnotes'` AND `id !== '48:notes'` | `id === '48:notes'` (self-chat) |
| `threadProperties.isDeleted === true` | otherwise |
| Folder member id starts with `slice-` | real conversations |

**Never** filter on `threadProperties.hidden` — it's set on active 1:1
chats, archive chats, and pinned chats alike. Teams uses it to mean
something other than "user removed".

## Cache layout (popup-side)

Picker stores enriched results in `chrome.storage.local` for instant
reload UX:

```
key:   convListCache
value: { version, at, conversations: ConversationSummary[], extras?: ConversationSummary[] }
```

Bump `CONV_LIST_CACHE_VERSION` whenever the *content* of a stored
entry could change semantically. Old caches are dropped, not migrated.

## What stays out of IDB

- **Message bodies for chats you haven't opened** — Teams only caches
  recent messages of recently-touched chats. Older history paginates
  over the wire on scroll.
- **Profile photos** — Graph fetches; SW-cached.
- **Live presence** — pushed via Trouter; held in memory only.
- **OneGQL / Relay store** — runtime memory, not persisted.

## Things we tried that didn't work

- **DOM title parsing** — locale-dependent, fragile across Teams UI
  reshuffles.
- **DOM sidebar `data-tabster` matching** — works ~95% but breaks for
  ambiguous chat names and cross-locale title variants like `(External)`.
- **`Teams:profiles`** for federated names — only holds app/consumer
  profiles, not actual user records.
- **`Teams:replychain-metadata-manager`** — per-conversation message-
  thread metadata, no display names.
- **Avatar URL `usersInfo` parameter** — does carry display names
  (Teams encodes them for the merged-photo service), but reading them
  is DOM scraping. We confirmed `chatTitle.shortTitle` exists *first*
  and avoided the avatar route entirely.
- **MAIN-world Trouter / fetch interception** — captured tokens and
  infrastructure pings; nothing chat-state-relevant flowed.

## Useful diagnostic snippets

Find which IDB store holds a literal string seen in the sidebar:

```js
(async () => {
  const target = '<paste sidebar text here>';
  const dbs = (await indexedDB.databases())
    .filter(d => d.name?.startsWith('Teams:'));
  for (const meta of dbs) {
    const db = await new Promise(r => { const x = indexedDB.open(meta.name); x.onsuccess = () => r(x.result); x.onerror = () => r(null); });
    if (!db) continue;
    for (const storeName of db.objectStoreNames) {
      let hit = null;
      await new Promise(res => {
        const cur = db.transaction(storeName, 'readonly').objectStore(storeName).openCursor();
        cur.onsuccess = () => {
          const c = cur.result;
          if (!c) { res(); return; }
          if (JSON.stringify(c.value).includes(target)) { hit = c.value; res(); return; }
          c.continue();
        };
        cur.onerror = () => res();
      });
      if (hit) { console.log('FOUND in', meta.name, '→', storeName); console.log(JSON.stringify(hit, null, 2)); db.close(); return; }
    }
    db.close();
  }
  console.log('not found');
})();
```

Recursively find the field path for a specific value within a record:

```js
function findFields(obj, target, path = '') {
  const hits = [];
  for (const [k, v] of Object.entries(obj || {})) {
    const here = path ? `${path}.${k}` : k;
    if (typeof v === 'string' && v.includes(target)) hits.push({ path: here, value: v.slice(0, 200) });
    else if (v && typeof v === 'object') hits.push(...findFields(v, target, here));
  }
  return hits;
}
```

## Open questions / followups

- **Multi-account**: Teams supports account switching. `selfUuid` in DB names changes; reader must look up the active account.
- **`tmp.session.*` keys missing in early tab load** — sessionStorage isn't populated until Teams' own boot completes. A brief fall-through to a DOM-based reader is still useful as a safety net.
- **Hidden-chats toggle UX** — currently we show all entries. A "Recent only" filter using `lastMessageTimeUtc` could shorten the list.

## Reference

- `EionRobb/purple-teams` — Trouter protocol implementation in C
- Microsoft Q&A confirms `48:notes` is the stable self-chat thread id
