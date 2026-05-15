# Changelog

All notable changes to this project will be documented in this file.

## [1.4.12] — 2026-05-15

Released to Microsoft Edge Add-ons. PDF emoji rendering on Chromium MV3 is fixed (color emoji instead of tofu boxes), and emoji in exported PDFs are now real selectable text: Ctrl+F finds them, drag-select copies the codepoints, screen readers announce them.

### Added

- **Microsoft Edge Add-ons listing.** Live at https://microsoftedge.microsoft.com/addons/detail/teams-chat-exporter/phlomfiieaggnbfpacmjmidcjdlaiplp. README install section now shows three badges: Chrome Web Store, Microsoft Edge, Firefox. Build pipeline ships `pnpm build:edge` and `pnpm zip:edge` for the Partner Center submission artifact.
- **Safari build target.** `pnpm build:safari` produces `.output/safari-mv2/`, a web-extension folder that can be wrapped via Xcode's `safari-web-extension-converter` for macOS or iOS App Store distribution. Not on a public Safari store yet.
- **PDF emoji are now selectable text.** Each unique emoji used in a document becomes one glyph in a per-export PDF Type 3 font whose ToUnicode CMap maps glyph codes back to the source Unicode codepoints, including multi-codepoint sequences (family emoji, flag emoji with skin-tone modifiers, ZWJ joins). Ctrl+F finds emoji. Drag-select across mixed text and emoji preserves the codepoints on copy. PDF text-extraction tools (pdftotext, Acrobat export) include emoji in the output.

### Fixed

- **PDF emoji no longer render as tofu boxes on Chrome and Edge.** Chromium MV3 service workers throw `InvalidStateError` when decoding SVG blobs via `createImageBitmap`, which broke the Twemoji rasterization pipeline for the entire Chromium audience (96% of users). The fix routes SVG decoding through a `chrome.offscreen` document, which has DOM access and runs the same HTMLImageElement + OffscreenCanvas pipeline that already worked in Firefox. Adds the silent `offscreen` permission (no user-facing install prompt). Requires Chromium 109+, which covers effectively every current user.
- **In-app "Rate this extension" link routes Edge users to the Edge Add-ons listing.** The user-agent detection that previously sent Edge users to the Chrome Web Store reviews page now resolves Edge to its own listing URL.

### Changed

- **PDF emoji pipeline reworked internally.** The old "rasterize SVG, embed as inline image, draw at font-size dimensions" path is replaced by a Type 3 font assembled per export from the same Twemoji rasters. PDF size is comparable; the font wraps each unique PNG with thin metadata and a ToUnicode CMap. Layout, positioning, and visible rendering match the previous output.

### Known limitations introduced by Type 3 emoji

- **Triple-click line-selection in Chrome's PDF viewer**: lines containing both text and emoji split into two selection units at the emoji boundary. Drag-select across the full line works and produces the correct text. PDFium quirk specific to Type 3 fonts; no structural change to the PDF makes triple-click cross the boundary.
- **Edge on Linux**: copying supplementary-plane codepoints (most emoji, U+10000+) from any PDF produces U+FFFD replacement characters. Reproduces with PDFs from Chrome and Firefox under the same conditions; the bug is in Edge's clipboard write path on Linux, not in any PDF. Edge on Windows and macOS is unaffected. Visual rendering is fine on Edge Linux; only the clipboard copy is broken.

### i18n

- Translator-hint fields for `extName` and `extDescription` updated across all 23 locales to mention both Chrome Web Store and Microsoft Edge Add-ons surfaces. User-facing strings unchanged.

## [1.4.7] — 2026-05-01

Picker labels now resolve real names for chats Teams scrubbed (issue #22 follow-up). New opt-in image-recovery toggle for users who want to embed external thumbnails Teams' proxy fails to deliver. Two Chrome-only bug fixes (download path, filename sanitization).

### Fixed

- **Picker no longer accepts "Unknown User" / "Just me" as final names.** Teams stamps these placeholder labels onto chats whose counterparty it can no longer resolve (left-org users, deleted accounts). The picker was using them as the chat name and skipping past its own resolution chain — which usually still has the actual name cached in the Teams replychain-manager (the most-recent non-self sender's display name from cached message authors). Now we treat those two strings as "no name set" and run the resolution chain. For private channels in the same degenerate state, added a senders-based fallback so they get the same recovery treatment as 1:1 chats. Cache version bumped (10 → 11) so existing users with cached "Unknown User" rows actually see the fix on their first popup open after upgrade — one cold refresh, then instant loads resume. Verified against the JSON exports user `simbamford` emailed: the names were sitting in cached message authors all along.
- **Chrome MV3 download path no longer fails with `URL.createObjectURL is not a function`.** Some Chrome versions don't expose `URL.createObjectURL` in extension service workers (Google has flipped this on/off across releases for security reasons). The three zip-output download paths (`per-chat-zip`, `html-zip`, `bundle-outer`) called it directly and broke silently — red error badge, no history entry, no actionable error. Added `blobToDownloadUrl()` that tries Object URL first and falls back to a base64 `data:` URL (read via `Blob.arrayBuffer()`). Existing `binaryToDownloadUrl` / `textToDownloadUrl` already had this fallback; the zip paths didn't. Same fallback shape extended to them. ~33% memory overhead during the conversion (68 MB blob → 90 MB data URL), but works wherever blob URLs fail.
- **Filenames with invisible Unicode characters no longer fail with "filename must not contain illegal characters" on Chrome (issue #21).** Some chat names contain zero-width spaces (U+200B–U+200F), directional formatting overrides (U+202A–U+202E), word joiners (U+2060–U+2064), or BOMs (U+FEFF) — typically pasted in from rich text editors. `chrome.downloads.download` rejects these characters in filenames. `sanitizeBase` now strips them entirely (not replaced with a dash) so a name that visually reads "Jane Doe" but has a hidden ZWNJ doesn't end up as "Jane-Doe". Reproduction credit + diagnosis credit to user `GeorgeDuckman`.

### Added

- **"Image fetch fallback" toggle in Settings.** Opt-in. Off by default. When Teams' image proxy returns a permanent-shaped failure (HTTP 410, 429, 403, 404) on an external link-preview thumbnail, the extension can fall back to fetching the image directly from the original source. Recovers thumbnails for sites like asciinema, GitHub OG cards, news article images, etc. — anywhere the proxy's negative cache outlives upstream availability. Requires the user to grant `<all_urls>` permission on toggle-on (declared as `optional_host_permissions` on Chrome MV3 / `optional_permissions` on Firefox MV2 — invisible at install time). Permission flow handles popup-died-during-prompt cleanly: sync-on-mount reconciliation reads the live permission state on every popup open and aligns the option flag with reality, so the user doesn't have to click the toggle twice if the prompt killed the popup. Permission-revocation listener also flips the option off if the user revokes the permission from `chrome://extensions` / `about:addons`. Background-mediated direct fetch (`FETCH_BLOB_DIRECT`) re-checks permission as a safety guard before each call. New post-export log line: `Image fetch fallback: N recovered, M still failed via direct upstream fetch`.

### Changed

- **External-state log lines demoted from `console.warn` to `console.log`.** Chrome's `chrome://extensions` Errors panel surfaces both warn and error and treats anything there as "the extension misbehaved." Network blips, AMS Bearer 401 → cookie auth recovery, per-host fetch breakdowns, SharePoint fetch failures, per-avatar HTTP failures, image fetch failure summaries: all reports of external state, not extension code errors. Now scoped to `console.log` so the Errors panel reflects actual extension issues. Still in DevTools console at log level for support / debugging. Stayed at `warn`: missing IC3 token, urlp helper script load failure, scrape errors, partial-export markers, DOM parsing surprises — all real code-side warnings worth surfacing.
- **Diagnostic logging on download paths.** `per-chat-zip`, `bundle-outer`, `pdf-single`, `text-single`, `html-zip` all now log `downloads.download calling`, `resolved` (with id), or `rejected` (with error message). Outer error catch in `handleExportWithScrape` now logs the actual error text — without this, red-badge failures were silent in the SW console and there was no way to debug them.

### i18n

- 5 new keys (`settings.imageFallback`, `settings.imageFallback.hint`, `settings.imageFallback.tooltip`, `settings.imageFallback.enable`, `settings.imageFallback.permissionDenied`) translated across all 24 locales.

## [1.4.6] — 2026-04-30

Fixes the "pasted screenshots and attached images don't appear in the export" regression that affected some tenants in v1.4.0–v1.4.5 (see issue #22). Plus a small PDF banner geometry fix.

### Fixed

- **AMS-direct image fetches now fall back to cookie auth on Bearer 401.** On some tenants, the `*.asyncgw.teams.microsoft.com/v1/{userId}/objects/...` proxy rejects the IC3 Bearer token even though the same token works for the chat-service API on the same host. Result: every real chat image (pasted screenshots, paperclip uploads, etc.) silently came up missing in the export. The fix adds a per-export auth-mode state machine: the first AMS-direct image is the canary. Bearer success keeps everyone on the existing fast path with no behavior change. Bearer 401 trips the state to `failed-401`, the canary is retried via the page-world cookie helper already in place for `/urlp/` thumbnails, and every subsequent AMS-direct fetch in that export uses cookies directly. Confirmed end-to-end fix on the affected tenant: image embed rate went from 9/214 to 214/214.
- **PDF partial-export banner no longer overlaps the title block.** The amber banner under the title on page 1 had two geometry bugs: the text was being drawn at the message-body column (past the avatar gutter), and the surrounding rectangle's top edge was placed *above* the cursor instead of below it, causing the banner to bleed up into the header. Fixed by switching to `drawMixed` directly with correct `x` offsets and anchoring the rectangle's bottom at `top - blockH`.

## [1.4.5] — 2026-04-29

Detect partial exports (network drop mid-scrape) and signal them across every artifact users will see, so a truncated export never silently masquerades as a complete one.

### Added

- **Pre-flight offline check.** Popup refuses to start an export when `navigator.onLine === false`, with a clear "No network. Reconnect and try again." message. Cheap insurance for the "wifi already off when user clicks Export" case. The inverse case (online but actually offline, e.g. captive portal) is caught in-flight by the network-error detection below.
- **In-flight partial-export detection.** When `apiScrape` fails with a `NetworkError` / `Failed to fetch` (browser-canonical fetch failure shapes), the content script flags the export as `partial` with `reason: 'network'`. DOM-scroll fallback still runs (it can recover messages already rendered before the network dropped) and any messages that come back are saved, but the partial flag travels through to every downstream consumer.
- **Partial-export signals at every layer:**
  - **Filename suffix** — single-chat outputs become `*-PARTIAL.html` / `*-PARTIAL.zip` / etc. Visible in the OS file browser without opening the file.
  - **In-file warning banners** — HTML gets an amber alert block at the top of the body. PDF gets an amber-tinted rectangle under the title block on page 1. TXT gets an ASCII-bordered notice at the top of the file. CSV gets `#`-prefixed comment lines at the top. Each carries the cause tag (`[network]`) for bug-report triage.
  - **History entry kind `partial`** — the popup history page renders these rows with an amber `⚠` badge and a `partial` status pill, distinct from `success` / `cancelled` / `failed`.
  - **Bundle root `PARTIAL.txt`** — multi-chat bundles get a top-level summary listing each affected chat's folder, conversation id, and reason. Sits alongside the existing `FAILURES.txt` and `NO_HISTORY.txt`.
  - **Bundle outer-zip `-PARTIAL` suffix** — the multi-chat zip itself becomes `TeamsExport_bundle_<date>-PARTIAL.zip` if any chat in it is partial.

### Changed

- `apiScrape` exposes a `getLastApiScrapeFailure()` getter so callers can distinguish "no token" / "no conv-id" / "network error" / "other" failure modes without changing the function's existing null-on-failure return shape.

### i18n

- New keys `errors.offline` and `history.partialMeta` translated across all 24 locales.

## [1.4.4] — 2026-04-29

Big win for HTML / PDF exports that contain link previews (Giphy, YouTube, news article OG images, etc.). Plus a richer per-host diagnostic when image fetches fail.

### Fixed

- **URL-image-proxy thumbnails now embed correctly.** Teams' `/urlp/.../url/image/Thumbnail?url=<external>` endpoint authenticates via cookies set by Teams' login flow on the asyncgw domain (`authtoken_asm_urlp`, `skypetoken_asm`), not via the IC3 Bearer the rest of the image-fetch path uses. Background and content-script fetches both sit in different cookie partitions than Teams' top-level origin (Firefox Total Cookie Protection, Chrome 3rd-party cookie phaseout) so neither could see those cookies and every external link-preview thumbnail returned 401. Fix: a tiny page-world helper script (`src/public/page-helpers/urlp-fetcher.js`) injected into the Teams page does the urlp fetches in the page's cookie partition; result is shipped back via `window.postMessage` with ArrayBuffer transfer. AMS direct (`/v1/{userId}/objects/...`) is unchanged. On the test tenant, image embed rate jumped from 564/772 (73%) to 771/772 (99.87%).

### Added

- **Per-host failure breakdown** in the image-fetch log (always on). Replaces the previous single-line "First http error" with a host-by-host summary: which upstream hosts succeeded, which failed, by status code, with the first failed URL per host.
- **Auth-state log** at the start of inline-image fetching: token presence and length, userId presence, region. No token contents.
- **Verbose DEBUG mode** (opt-in via `WXT_DEBUG_IMAGE_FETCH=1` build flag or `localStorage.setItem('__teams_exporter_debug_image_fetch', '1')`): full IC3 JWT claims (decoded payload only, never the token), first 5 sample URLs (raw + transformed), first call + first response of each fetch path, first 30 failed URLs.

## [1.4.3] — 2026-04-28

Hotfix for the popup-after-install case. If you installed the extension and opened the popup before refreshing your Teams tab, the auto-inject path was broken and you saw a generic "Could not load chats" with no actionable detail.

### Fixed

- **Auto-inject path corrected.** `ensureContentScript` was passing `'content.js'` to `chrome.scripting.executeScript`, but WXT bundles the file at `'content-scripts/content.js'`. Every fresh install where the user hadn't refreshed Teams hit `"Could not load file: 'content.js'."` and silently fell into the picker error state. Now the path matches the bundle and auto-inject works the first time.

### Added

- **Visible inline error in the picker.** The error block now shows the actual error message under "Could not load chats" (used to be tooltip-only on the retry button). This is what made the auto-inject bug findable in the first place.
- **Frame-level error surfacing.** `ensureContentScript` now inspects the per-frame `InjectionResult` and reports top-frame errors, so a sandboxed iframe / page-CSP block / unmatched URL no longer fails silently.
- **Retry on post-injection PING.** Up to 5 attempts at 50 ms intervals to absorb any race between `executeScript` resolving and the listener registering.

## [1.4.2] — 2026-04-28

Onboarding-tour polish + a small picker UX fix.

### Added

- **Pre-tour prompt.** First-time popup open now asks "Want a quick 30-second tour?" with [No thanks] / [Show me] buttons before highlighting any element. Users who already know the extension can opt out without sitting through 7 steps.
- **"Teams is still loading" empty-state.** When the picker comes back empty on a Teams tab (IDB hasn't populated yet), we now show a clear "Teams is still loading your chats — try refreshing in a few seconds" message with a Refresh button, instead of the misleading "No chats found".

### Changed

- Onboarding dismiss is now visually distinct: the corner X is red, and a labeled red "Skip" button sits in the action row alongside Back/Next during the tour. Dismissal is one click away whether the user looks at the corner or the action row.
- Settings → Replay tour bypasses the new pre-tour prompt — clicking it is already an explicit opt-in, no need to ask again.

### Fixed

- Added top margin to the picker's "Open the Teams web app tab first." / "Could not load chats" error blocks so they no longer sit flush against the "Conversations" card header.

## [1.4.1] — 2026-04-28

Polish release on top of v1.4. Reworks the welcome experience and makes the conversation picker more discoverable.

### Added

- **Conversation picker collapse toggle.** New chevron in the picker header collapses / expands the list. Default is collapsed for fresh installs, persisted across sessions. Active filter auto-scrolls into view on popup open with a brief flash so the current selection is never hidden offscreen.
- **7-step interactive onboarding tour.** Replaces the old centered modal. Each step highlights a real popup element (format, picker, folder rail, date range, history, settings, export/stop) with a feathered SVG-mask spotlight and a card that auto-positions opposite the target. The folder step temporarily expands the picker and restores it on dismiss.
- **Replay tour entry in Settings.** A "Replay tour" card lets users walk through the onboarding again at any time.

### Changed

- Onboarding spotlight now uses an SVG `<mask>` with `feGaussianBlur` for soft-edged dim, drawn on a `position: fixed` scrim so it covers the popup correctly even when the highlighted element lives inside a scrollable / `overflow: hidden` ancestor.
- Scroll/resize handler for the spotlight is rAF-throttled — fast scrolls inside the picker rail no longer stutter under repeated `getBoundingClientRect` calls.
- Onboarding copy across all 24 locales updated for the new step structure; deprecated `onboarding.step1.*` / `onboarding.step2.*` keys removed.

### Fixed

- Onboarding card no longer overlaps the highlighted target when the popup is scrolled — card position is recomputed per step from the target's actual rect.

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
