# TODO

## Planned

- [ ] Add lint + formatter setup (ESLint/Prettier).
- [ ] Add unit tests for utility and builder modules.
- [ ] Add popup UI tests.
- [ ] Add content-script tests with mocked DOM.
- [ ] Improve handling when long exports stall (the existing `fetchPageWithRetry` covers transient 403/429/5xx; this would be a separate watchdog for the case where the API keeps returning 200s but no progress is made).
- [ ] Add participant filtering (export only messages from selected authors).
- [ ] Add user-configurable image-fetch domain allowlist (currently hardcoded in `src/content/attachments.ts`).
- [ ] Expose tunable scrape constants in Settings. Currently hardcoded:
   - `MAX_IMAGE_BYTES = 5 MB` per-image cap (drops oversize embeds; `attachments.ts` and `content.ts`).
   - `4096 × 4096` per-image pixel cap (`attachments.ts`).
   - `MAX_PAGES = 500` API pagination ceiling (`api-client.ts`); not tight in practice but could be useful for users with very large channel histories.
   - DOM-scroll stagnation thresholds (12/20 default, 30/35 for team channels in `content.ts`); slow networks may need higher.
   These are all defensive defaults today. Surfacing them as advanced settings would let power users tune for very large or media-heavy chats without a code change.
- [ ] Localize PDF timestamps per `options.lang` (currently fixed `YYYY-MM-DD HH:MM`). Use `Intl.DateTimeFormat` with the user's lang.
- [ ] Shrink CJK coverage further. Runtime HarfBuzz subsetting already trims the bundled `NotoSansSC` per-export; open question is whether to ship a smaller CJK base font or fetch glyphs on demand for users who never export CJK content.
- [ ] Native-speaker review pass over the 22 non-English, non-Turkish locales (they're at full parity but translations are AI-assisted, not native).
- [ ] TXT format silently drops attachment markers — image-only messages render as blank lines. Add a placeholder so users know the message had content.
- [ ] Optional advanced setting to expose `MutedChats` system folder in the picker rail (the only system folder that adds new filtering not covered by the kind tabs).
- [ ] Layer B chunking inside `toHTML` / `toCSV` / `toPlainText` if heavy single-format exports show popup-stall symptoms after the v1.4 between-format yields.

## Done

### Post-v1.4 fixes (on `main`, unreleased at time of writing)

- [x] **Outer-zip OOM fix.** `buildZipAsync` now resolves a `Blob` constructed directly from the chunk array; downloads consume the Blob via `URL.createObjectURL`. The previous contiguous-`Uint8Array` finalize plus `new Blob([uint8])` in `binaryToDownloadUrl` doubled peak memory at the worst moment and blew Firefox's allocation cap on a 2.5 GB / 289-chat bundle. Verified end-to-end on the same workload.
- [x] **Smart per-extension deflate level.** Already-compressed inputs (`.zip`, `.pdf`, `.jpg`, `.png`, `.gif`, `.webp`, `.mp3`, `.mp4`, `.webm`, `.gz`, `.7z`, `.rar`, `.bz2`, `.xz`) get level 0 (store); raw text gets level 6. Same final size, ~3× faster outer-zip step on real bundles.
- [x] **Bounded retry on transient 403 / 5xx in `fetchPageWithRetry`.** 3 attempts at 1s / 2s / 4s. The 5-attempt 429 policy is unchanged.
- [x] **`ScrapeOptions.noDomFallback` for multi-chat bundles.** Bundle iteration sets it per-chat. The content script reports a clean error instead of falling back to DOM scrolling, which would scrape whichever chat is currently visible in the user's tab — not the target chat. Failed chats land in `FAILURES.txt`.
- [x] **All-failed bundle UX.** When 0 chats succeed, skip the .zip wrapper and save `TeamsExport_bundle_<stamp>_FAILURES.txt` directly. New `HistoryEntry.kind = 'failed'` renders with an amber badge and "all failed" pill in the popup history; Open / Show in folder still work because the file is real on disk.
- [x] Picker now shows a clear "Open Teams web first" message when the popup opens on a non-Teams tab (instead of the misleading "No matches" empty-state).
- [x] Auto-default selection (the active chat) seeds at most once per popup open. Once the user touches the selection, later refresh phases never reinstate the original auto-pick.
- [x] Dead-code cleanup pass: tsc `--noUnusedLocals --noUnusedParameters` + un-export of internal-only symbols + an exports-never-imported scan. ~190 lines net removed across utils, types, and content modules.

### v1.4

- [x] **v1.4** Conversation picker — IDB-source list, kind + folder rail filter (P1), persistence for both axes, smart IDB-diff refresh, bulk-select via head pill dropdown + icon action bar.
- [x] **v1.4** Multi-chat bundle export — pick N chats, get one outer zip with per-chat folders + `FAILURES.txt` for any chat that errored. Hard-abort on cancel; CPU yield between formats.
- [x] **v1.4** System-event leak fixed — `ThreadActivity/*` (incl. `JoiningEnabledUpdate`) are now classified as system rather than leaking XML inner-text into TXT/CSV/HTML rows.
- [x] **v1.4** HTML code-block long-line wrap — `pre.code-block` switched from `white-space: pre` to `pre-wrap` so long commands don't push the page into horizontal scroll.
- [x] **v1.4** Cleanup — removed dev-probe scaffolding (`src/dev/`, `dev-probe-*` entrypoints, `DEV_PROBE_*` handlers), `extractConversationId` DOM heuristics (replaced by sessionStorage navHistory read), `fetchSingleConversation` + `FETCH_CONVERSATION` + `discoveredExtras` machinery, the chat-service API `listConversations` fallback, and tracing console logs. ~1,400 lines deleted.
- [x] **v1.4** Popup white-screen instrumentation — timestamped `[POPUP]` mount-boundary traces + EXPORT_STATUS broadcast-rate counter on both popup and SW sides + try/catch around `init()` so synchronous throws surface in the error banner.
- [x] Deduplicated background export handlers.
- [x] API-based message fetching with DOM scroll fallback.
- [x] MCAS proxy URL support (`.mcas.ms` suffix).
- [x] Forwarded message detection and rendering.
- [x] Inline images, GIFs, audio, and video thumbnails in HTML export.
- [x] Avatar embedding in HTML and JSON exports.
- [x] Link preview extraction and rendering.
- [x] Large export streaming to avoid 64 MiB message limit.
- [x] Auto-upgrade to ZIP for large HTML exports.
- [x] Extended ExportMessage with contentHtml, messageType, forwarded, importance, subject, mentions.
- [x] PDF export (pdf-lib + HarfBuzz runtime font subsetting + Twemoji rasterization + clickable link annotations).
- [x] Multi-format export packaged as `bundle.zip`.
- [x] Stop-in-progress export control with phase tracker.
- [x] Export history page inside the popup (replaces the post-export tile).
- [x] About card in Settings with name/version/source/issue/author links.
- [x] Last-open popup page persistence (main / settings / history).
- [x] Locale-independent chat/channel context detection (`data-tid` only; fixes Japanese and French UIs).
- [x] Own-message styling (`isOwn` + blue left rail) in HTML and PDF.
- [x] Nested reply chains render correctly in HTML (flatten to top-level parent).