# v1.4 release plan

Roadmap for the remaining work on the `experiment/conversation-list`
branch. Ordered by recommended sequence; each section is sized so
context-compaction between phases is safe.

## Where we are

The IDB-backed picker is feature-complete for single-chat selection:
sidebar parity for naming (Teams' `chatTitle.shortTitle`), folders
detected, multi-locale DBs merged, fast cold load via two-pass render,
sessionStorage active-chat detection. See `docs/TEAMS_INTERNALS.md`
for the data-source design.

What's left for v1.4 is wiring + cleanup:

1. Multi-chat bundle export (the original v1.4 promise)
2. Folder rail filter in the picker
3. Cleanup: rip `src/dev/` + redundant fallback paths
4. i18n catch-up (22 locales)
5. Version bump + release docs

---

## 1. Multi-chat bundle export

Largest item. The picker already supports multi-select state
(`selectedConversationIds: string[]`); the export pipeline needs to
loop, dedupe filenames, and emit a single bundle.

### Goal

Selecting N chats in the picker and clicking Export produces:

```
bundles.zip
├── <chat-1-name>/
│   ├── messages.json
│   ├── messages.html        (or whichever format(s) the user picked)
│   └── images/              (when applicable)
├── <chat-2-name>/
│   └── ...
└── FAILURES.txt              (only present when any chat failed)
```

### Key decisions (already settled with the user)

- File layout: `bundles.zip/<chatname>/{format files}` (per-chat folder)
- Filename collisions: append `(2)`, `(3)`, etc. (already handled by
  `sanitizeBase` + a small per-bundle dedupe map)
- Progress reporting: per-chat AND within-chat granularity
- Rate limiting: serial scrape per chat with a small delay; parallel
  was deemed too risky for throttling
- Partial failure: don't abort the whole run; retry failed chats; on
  persistent failure, write a `FAILURES.txt` line (not a JSON file)
- Multi-chat suffix in main `bundles.zip` filename

### Implementation sketch

Files to touch:

- `src/entrypoints/popup/components/ConversationPicker.svelte` — switch
  `mode` to `'multi'`, surface "N selected" state in Export button label.
- `src/entrypoints/popup/App.svelte` — when `selectedConversationIds.length > 1`,
  send a different request type (`START_BUNDLE_EXPORT`) carrying the
  full id array, plus the resolved name per id (for folder naming).
- `src/types/messaging.ts` — add `StartBundleExportRequest` /
  `StartBundleExportResponse` / per-chat progress events.
- `src/entrypoints/background.ts` — new `handleBundleExportMessage`:
  - serial loop over the id array
  - for each: existing scrape pipeline + per-chat zip building
  - aggregate into one outer zip via JSZip
  - small inter-chat delay (e.g. 500ms)
  - track failures in a list, write `FAILURES.txt` if non-empty
- `src/background/download.ts` — add `buildAndDownloadBundlesZip()` that
  takes an array of `{convName, scrapeRes, buildOptions}` and emits
  the nested zip
- Filename collision: per-bundle Set of seen folder names; if conflict,
  append `(2)`, `(3)`, etc. (sanitizeBase already handles individual
  names; collision logic is new)
- Status payload extensions: `{ phase: 'bundle', currentChat: N, totalChats: M, ...standardPhase }`

### Tricky bits

- **Cancellation**: STOP_EXPORT must abort the current chat AND skip
  remaining ones. The existing `currentAbortController` only covers
  one scrape — extend to a bundle-level controller.
- **Memory pressure**: serial keeps memory bounded to one chat at a
  time. Each chat's blobs must be released before the next starts.
  Streaming directly into the outer zip (vs holding all per-chat zips
  in memory) is the safer pattern.
- **Progress UI**: the picker's segmented progress bar shows phases
  for one chat. For bundle, add a "chat N of M" prefix on the phase
  label without redesigning the bar.
- **Filename**: when `selectedConversationIds.length > 1`, output is
  `bundles.zip` (or `bundles_<timestamp>.zip`); use sanitized
  individual chat names only as folder names inside.

### Estimated scope

3-4 commits, ~500-700 lines. Doable in one focused session.

---

## 2. Folder rail filter

Folders are already read into `ConversationSummary[]` indirectly via
`readFolders()` in `src/content/teams-state.ts`, but the picker doesn't
expose them as a filter axis.

### Goal

Add a second filter dimension to the picker's left rail:

- Existing kind filter stays (All / Chats / Groups / Meetings / Channels)
- New axis: by folder (Favorites / MeetingChats / MutedChats / custom names)
- User picks one or both

### Decisions to make

- UX: add as a separate dropdown above the rail? Append to the rail
  after kinds? Use a tabs-strip at the top instead of vertical icons?
- System folders to expose: probably just `Favorites` + user-created.
  Hide `MeetingChats` / `MutedChats` / `QuickViews` / `RecentChats` /
  `TeamsAndChannels` (they're system computations, often empty in our
  IDB anyway).
- "All folders" default state: shows all chats regardless of folder.

### Implementation sketch

- Extend `ConversationSummary` with `folderIds?: string[]` (a chat can
  be in multiple folders — we already see `Favorites` overlapping with
  user-created).
- `listConversationsFromIdb`: cross-reference each conversation against
  the folder list. Currently we read folders but discard them.
- `ConversationPicker.svelte`: add a folder filter UI. Smallest change
  is a horizontal pill list above the rail, showing only user-created
  folders + Favorites. Click toggles.
- Filter order: kind first (cheap), then folder (Set lookup).
- Persist last-chosen folder in extension storage so it survives popup
  reopens.

### Estimated scope

1-2 commits, ~150 lines.

---

## 3. Cleanup

The branch carries a lot of scaffolding from the research phase. Once
we're happy with the picker behavior:

### What to remove

- `src/dev/` — entire directory (probe types, ring buffer, README)
- `src/entrypoints/dev-probe-main.content.ts`
- `src/entrypoints/dev-probe-bridge.content.ts`
- `src/entrypoints/dev-probe-page/` — directory + entries
- Background message handlers `DEV_PROBE_*` in `src/entrypoints/background.ts`
- Console diagnostic log: `[API] replychain senders scanned: …`
- Console diagnostic log: `[API] productThreadType seen: …`
- Console diagnostic log: `[API] extractConversationId sidebar miss / via sidebar title-match`
- The `extractConversationId` DOM-heuristic codepath in
  `src/content/api-client.ts` — replaced by sessionStorage-based reader
- `fetchSingleConversation` + the `FETCH_CONVERSATION` message — IDB
  has the full set, fetch-on-miss is dead code
- The `discoveredExtras` machinery in `App.svelte` (cache extras for
  fetched-on-miss results) — redundant once fetch-on-miss is gone
- The chat-service-API `listConversations` path in `api-client.ts`
  (the API source) — IDB version covers it; fallback only fires when
  IDB is empty, which itself is degenerate
- The `LIST_CONVERSATIONS_QUICK` separate message could fold back into
  `LIST_CONVERSATIONS` if the popup is restructured to handle two-pass
  natively (low priority, leave as-is for now)

### Order of operations

1. One commit per logical removal — easier to revert if anything is
   load-bearing.
2. Run both Chrome + Firefox builds after each removal to make sure
   nothing was secretly imported.
3. `pnpm exec svelte-check` after each step.

### Estimated scope

3-5 commits, mostly deletes (~800-1000 lines removed).

---

## 4. i18n catch-up — 22 locales

Throughout the picker work we added new keys to `en.json` + `tr.json`
only. The other 22 locales are stale.

### Keys added during this phase

(Cross-reference `src/i18n/locales/en.json` for current set; anything
under `picker.*` added since the v1.3 baseline.)

- `picker.rail.label`, `picker.rail.all/chats/groups/meetings/channels`
- `picker.noSelection`, `picker.nSelected`
- `picker.refreshing`, `picker.refreshTitle`
- `picker.selfChatSuffix`, `picker.selfChatLabel`
- `picker.selfChatFallback`, `picker.chatFallback`,
  `picker.groupFallback`, `picker.meetingFallback`,
  `picker.channelFallback`
- `picker.plusMoreMembers`
- `picker.firstLoad`
- (Any new keys for multi-chat export progress + folder rail get added
  during their sections; sweep all together at the end)

### Approach

- Mechanical pass through `src/i18n/locales/*.json`
- Translations: my own (per the existing TODO note that all non-
  English/Turkish translations are AI-assisted, not native)
- Verify no JSON syntax errors per-file via `pnpm exec svelte-check`

### Estimated scope

1 commit, mechanical edit across 22 files.

---

## 5. Version bump + release docs

### Files to update

- `package.json` → `version: "1.4.0"`
- `wxt.config.ts` → `manifest.version: "1.4.0"` (must match)
- `docs/TODO.md` → move v1.4 items from "Planned" to "Done"
- `docs/ARCHITECTURE.md` → update if any architectural change
  (the IDB-source pivot warrants a paragraph)
- `README.md` → update feature list if multi-chat / folder filter
  add user-visible features
- `CHANGELOG.md` (if exists) → add v1.4 entry
- Issue templates (referencing version) → bump if needed

### Pre-release checks

- Both builds clean: `pnpm build && pnpm build:firefox`
- Type check clean: `pnpm exec svelte-check`
- Manual smoke test: cold load + warm load + chat switch + export
  on a real Teams account

### Estimated scope

1 commit.

---

## Recommended order

1. **Multi-chat bundle export** — biggest, most user-visible. Most
   value if we get it done.
2. **Folder rail filter** — clean, discrete, easier than #1; saves
   for a second working session.
3. **Cleanup** — once #1 + #2 are stable. No-op for users; reduces
   maintenance burden going forward.
4. **i18n catch-up** — last because all new keys must be in place.
5. **Version bump** — final commit; tagged for release.

Compaction-safe boundaries: between any of the five, after the cleanup
within #3, and especially after #1 (which is the one likely to expand
in scope as edge cases surface).
