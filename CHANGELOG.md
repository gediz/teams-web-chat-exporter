# Changelog

All notable changes to this project will be documented in this file.

## [1.4.0] — 2026-04-28

The major theme of v1.4 is **scale**: pick many chats and export them all at once, see them in a proper picker instead of "whatever's open", and read the result in any of five formats (now including PDF). Plus full support for Teams Free (consumer accounts).

### Highlights

- **Multi-chat bundle export.** Pick N conversations from the new picker; get one outer `.zip` with per-chat folders, plus `FAILURES.txt` and `NO_HISTORY.txt` summaries at the root.
- **Conversation picker overhaul.** Sidebar-style list backed by Teams' own IndexedDB (the full set, including meeting-derived chats and niche product types the chat-service API omits). Multi-select with bulk-select shortcuts (M for head pill dropdown, N for icon action bar), folder rail filter, kind filter, both selections persist across popup opens.
- **Teams Free (consumer accounts) support.** Auth, conversation list, message history, system event author resolution, and inline image embedding all work on `teams.live.com` accounts. Modern 1:1s, group chats, and channels export normally; legacy Skype-imported 1:1s (where Microsoft never migrated history into the consumer chat backend) are listed in `NO_HISTORY.txt`.
- **PDF export.** Built with `pdf-lib` + runtime HarfBuzz font subsetting (so a 10 MB CJK font becomes <1 MB per export) + Twemoji rasterisation + clickable link annotations. Packaged with JSON / CSV / HTML / TXT for "all formats" exports.
- **Stop in-progress exports** with a phase tracker, an export history page inside the popup (replaces the post-export tile), and an About card in Settings.

### Added

- Multi-chat bundle export with per-chat folders, `FAILURES.txt`, `NO_HISTORY.txt`, and CPU yields between formats so the popup stays responsive.
- IDB-backed conversation picker — sidebar list, kind tabs, folder rail (P1), multi-select, bulk-select via head pill dropdown (M) and icon action bar (N), persistent kind + folder selection.
- Teams Free chat-service support — AES-CBC decrypt of cached `Discover.SKYPE-TOKEN`, profile-name resolution from local IDB, direct-AMS image fetch path, `*.teams.live.com` host permission.
- PDF export — `pdf-lib`, HarfBuzz runtime subsetting, Twemoji color emoji, clickable link annotations.
- Multi-format export packaged as `bundle.zip` (any combination of JSON / CSV / HTML / TXT / PDF).
- Stop-in-progress export with phase tracker.
- Export history page in popup (replaces the post-export tile).
- About card in Settings (name, version, source, issue tracker, author links).
- Last-open popup page persistence (main / settings / history).
- Inline images, GIFs, audio, video thumbnails in HTML export.
- Avatar embedding in HTML and JSON exports.
- Link preview extraction and rendering.
- Forwarded message detection and rendering with original author + timestamp.
- Team channel export (in addition to direct message / chat).
- MCAS proxy URL support (`.mcas.ms`).
- Non-invasive review prompt.
- TXT/CSV attachment summary lines (`[image: name.png]`, `[file: name.bmp]`, `[GIF]`, `[link: title]`) so image-only and file-only messages don't render as silent blank rows.
- 18 new picker translation keys across all 24 locales.

### Changed

- Conversation list source: Teams' own IndexedDB instead of the chat-service `/v1/users/ME/conversations` endpoint. IDB has the full local set including meeting-derived chats and niche product types the API omits.
- HTML image-detection regex extended with `bmp`, `svg`, `tiff`, `heic` so non-jpeg image formats render as images instead of generic links.
- Empty PDFs drop from ~7 MB to ~20 KB (HarfBuzz subsetting now produces a 1-glyph fallback for empty documents instead of bailing and embedding the full font).
- Smart per-extension deflate level: already-compressed inputs (`.zip`, `.pdf`, `.jpg`, `.png`, `.gif`, `.webp`, `.mp3`, `.mp4`, `.webm`, `.gz`, `.7z`, `.rar`, `.bz2`, `.xz`) get level 0 (store); raw text gets level 6. ~3× faster outer-zip step on real bundles, same final size.
- Bounded retry on transient 403 / 5xx in `fetchPageWithRetry` (3 attempts at 1s / 2s / 4s).
- Multi-chat bundle iteration sets `noDomFallback: true` so per-chat failures land in `FAILURES.txt` cleanly instead of falling back to DOM-scrolling whichever chat happens to be visible in the user's tab.

### Fixed

- (#22) Web-paste alt-text leak: URLs containing closing-quote + HTML attribute syntax in `alt=` no longer corrupt downstream filename derivation or display labels.
- Conversation-ID cache invalidates on chat switch.
- HTML export now correctly renders deeply nested reply chains (flattened to top-level parent).
- HTML export no longer shows broken-image icons when "Inline images" is off — switches to a quiet `(not included)` placeholder card.
- Image fetch reliability improved (better Graph photo logging, retries on transient errors).
- Forwarded messages: duplicate-message dedup, original sender/timestamp rendered.
- System-event leak fixed — `ThreadActivity/*` (incl. `JoiningEnabledUpdate`) are now classified as system rather than leaking XML inner-text into TXT/CSV/HTML rows.
- HTML code-block long-line wrap: `pre.code-block` switched from `white-space: pre` to `pre-wrap` so long commands don't push the page into horizontal scroll.
- Outer-zip OOM on 289-chat / 2.5 GB Firefox bundle (resolves a `Blob` directly from chunks instead of materialising a contiguous `Uint8Array` first).
- Outer-zip CSP hang in Firefox MV2.
- Popup unresponsive during outer-bundle zip step.
- Popup white-screen instrumentation: timestamped `[POPUP]` mount-boundary traces + EXPORT_STATUS broadcast-rate counter on both popup and SW sides.
- 1:1 chat name derivation (resolves duplicate-name regression).
- `(You) (You)` duplicate suffix regression on Teams Free self-chats (Teams' UI already includes a localised suffix; we now strip before re-adding).
- All-failed bundle UX: when 0 chats succeed, save `TeamsExport_bundle_<stamp>_FAILURES.txt` directly without a zip wrapper. New `HistoryEntry.kind = 'failed'` renders with an amber badge in the popup history.
- Picker shows a clear "Open Teams web first" message on non-Teams tabs (was: misleading "No matches" empty-state).
- Auto-default selection (the active chat) seeds at most once per popup open. Once the user touches the selection, later refresh phases never reinstate the original auto-pick.
- Locale-independent chat/channel context detection (`data-tid` only; fixes Japanese and French UIs).
- Own-message styling (`isOwn` + blue left rail) in HTML and PDF.
- Calendar state syncs with app state on theme + remount (#8, #15.3).
- Reply connector circle aligns to avatar center.

### Internal / Dev

- ~1,400 lines deleted in v1.4 cleanup pass: `src/dev/`, dev-probe entrypoints, `extractConversationId` DOM heuristics (replaced by sessionStorage navHistory read), `fetchSingleConversation` + `FETCH_CONVERSATION` + `discoveredExtras` machinery, the chat-service API `listConversations` fallback, tracing console logs.
- ~190 lines of dead code removed in follow-up audit (`tsc --noUnusedLocals --noUnusedParameters` + exports-never-imported scan).
- 12 symbols un-exported (clarifies intent as module-private; 2 turned out completely unused once the export keyword stopped masking them).
- HTML-fragment parsing in the converter switched to `DOMParser`-based extraction, producing an inert document so script tags don't execute, image/video/iframe sources don't trigger network requests, and inline event handlers don't fire.
- All 24 locales now at 177 keys each (full picker.* parity).
- npm → pnpm migration; CI also runs on push to main.
- Documentation accuracy pass: every concrete claim in every `.md` file verified against current source; `V14_PLAN.md` removed.

### Known limitations

- Teams Free legacy Skype-imported 1:1 chats (id ends in `@oneToOne.skype` with `threadProperties.isMigrated`): Microsoft never migrated those histories into the consumer chat backend, so the server returns `messages: []`. Listed in `NO_HISTORY.txt`. See `docs/TODO.md` for a possible recovery path via the legacy Skype API.
- Teams Free SharePoint Personal Content paperclip uploads (consumer OneDrive at `my.microsoftpersonalcontent.com`): cannot fetch programmatically. The host returns a 302 to `login.live.com` for any cross-origin or non-interactive caller. The HTML export still shows the file as a clickable link the user can open in OneDrive manually.
