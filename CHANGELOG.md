# Changelog

All notable changes to this project will be documented in this file.

## [1.6.0] - 2026-07-04

Recovers file attachments that used to fail to download, from your own uploaded files to files behind a stale session or a migrated tenant, and makes large attachment downloads far more reliable. Adds controls for the download (a size cap, a file-type skip list, an adjustable number of simultaneous downloads, and an optional last-modified date in each file's name), a Diagnostics tool to rescue files that still failed, and opt-in date stamps on saved inline images. The Settings page is reorganised into grouped sections.

### Added

- **Your own uploaded files now download.** A file you uploaded yourself often has no sharing link, so the download step had nothing to redeem and gave up on it. Those files now download the way Teams itself opens them, addressed by the file's own id, so your own shared files are saved like any other. This recovered the bulk of the files that previously failed on a large real-world chat.
- **Attachments that fail on a stale session now download.** A file whose sharing link could not be fetched with the site's session cookie (for example on a migrated tenant) is now resolved through SharePoint to a short-lived pre-authorized link and downloaded from there. If that path finds nothing (no sharing link, no access, or the file is unreachable), it falls back to the old direct link, so nothing that used to work regresses.
- **Export and its attachments now save in one folder.** With the "Files" toggle on, an export and its `attachments` tree now save together inside a single folder in Downloads, so one export is one folder. With the toggle off, nothing changes: same filenames, same Save As dialog.
- **Salvage failed files (Diagnostics).** A new tool on the Diagnostics page takes a text file of file links (a `FAILURES.txt` from an earlier export, or a plain list of URLs), resolves each one, and downloads it one at a time into a `TCE-probe` folder in Downloads. It runs in the background, so you can close the popup while it works, and it lets you recover the files that failed transiently on a big export without re-running the whole export.
- **File-access probe (Diagnostics).** A new probe lets you paste a file link and see, step by step, whether the extension can resolve and download it, to diagnose why a specific attachment fails.
- **Size cap.** A new "Skip files larger than (MB)" setting leaves big downloads out of an export. Empty or 0 downloads everything. The size is read before the transfer starts, so a skipped file never downloads and is reported as skipped rather than failed.
- **Skip by file type.** A new "Skip these file types" setting takes a comma-separated list of extensions (for example: exe, zip) and leaves those out. The type is known before downloading, so a skipped file never transfers.
- **Date in attachment filenames (opt-in, off by default).** A new setting prefixes each downloaded file's name with its real last-modified date in UTC, so files sort by date in your file manager. Downloaded files are saved by the browser, which cannot set a file's modified date, so the date is carried in the name.
- **Simultaneous downloads (advanced).** A new setting sets how many attachments download at once, from 1 to 8, default 6. Lower is gentler on a slow or metered connection; higher can be faster but may make the server throttle the burst.
- **Date in inline-image filenames (opt-in, off by default).** A new setting prefixes each saved inline image's name with the message's share time in UTC (for example `20260628T104615Z__photo.png`), so images pulled out of an export sort by date. Off by default, so existing image names are unchanged unless you turn it on.
- **Modified date on saved inline images (opt-in, off by default).** A separate new setting sets each saved image's date-modified in the zip to the time it was shared, so the extracted file's "Date modified" reflects when it was sent rather than when it was exported. Whether the date survives depends on the tool used to unzip: 7-Zip, WinRAR, macOS, and Linux keep it; Windows Explorer's "Extract All" and some cloud or Android tools may not.

### Fixed

- **Fewer failed downloads on large exports.** A burst of attachment downloads on a big export could fail in clusters from network contention, and a dropped connection was a permanent failure because nothing retried it. Downloads now stay within a fixed number of transfers at once, retry transient network failures with spread-out backoff and a fresh pre-authorized link each attempt, stagger their starts, and hold very large files (over 100 MB) to a couple at a time, so more come down on the first pass. Interrupted partial files are cleaned up instead of being left behind (a failed run could otherwise orphan over a gigabyte), and the failure list no longer repeats an entry once per attempt.
- **"N files saved" now reflects real completions.** The Files phase now waits for each attachment to reach a confirmed final state before counting it, and stopping mid-download keeps the saved export, cancels the in-flight transfers, and records them in a separate cancelled section instead of the failed list. The export file itself is also held until the browser confirms it finished, so a genuine failure is reported as an error rather than being mislabelled as complete.
- **Own uploads keep their size and date even without a download link.** A file resolved by its id that returns its size and last-modified date but no ready-to-fetch link now keeps that information, so the size cap and the filename date still apply to it.
- **Settings fields no longer show stale text.** Typing a value into the size-cap or skip-types field that reduced to the value already stored (for example a negative number, or a trailing space) could leave the typed text on screen while the saved value differed. The field now always shows what is actually saved.

### Changed

- **Settings page redesigned.** Options are now grouped into titled sections (Appearance, Export, Images, Files, PDF, Help & feedback, About) instead of one card per option. The theme choice is a Light/Dark toggle, the language picker moved to its own searchable sub-page instead of a long grid, the four image options are gathered under Images, and each row carries a short hint with an info button that reveals the full detail on hover, focus, or screen reader. A muted version footer was added.
- **"Inline images (HTML only)" renamed to "Images".** The toggle also controls inline images in the PDF, so the old label was inaccurate. It is now just "Images", matching the "Files" toggle.
- **Diagnostics page styling fixed in dark mode.** The Diagnostics cards, buttons, and code blocks previously stayed light even with the dark theme on; they now follow the theme. The header buttons on Settings, History, and Diagnostics also match the flat style of the main page.

## [1.5.2] - 2026-06-29

Adds document download, deleted-message placeholders, and an opt-in full-resolution image setting. Also fixes several image and attachment export-fidelity issues found while auditing a large real-world chat.

### Added

- **Document download.** A new "Files" toggle in the Include section saves file attachments (PDF, Word, Excel, ZIP, video, and similar) to your Downloads folder, in a per-chat `attachments` folder beside the export. It is off by default, and the readable export still keeps every file's link. Inline images stay with the existing image setting; this covers documents that were otherwise only linked. Any file that cannot be downloaded (you do not have access, or it was moved or deleted) is listed in a `FAILURES.txt` in the same folder with its link, so nothing goes missing without a record; files you stop mid-download are listed separately.
- **Deleted messages are kept as placeholders.** A message deleted for everyone used to be dropped, leaving an unexplained gap in the conversation. It now appears as "[message deleted]" with the original sender and timestamp, in its correct position.
- **Full-resolution images (opt-in, off by default).** A new setting saves inline images at their original resolution instead of Teams' downscaled display view, which caps around 1280 px. It is off by default; turning it on produces much larger exports, especially the PDF, for the same on-screen size. Any image that is too large or cannot be fetched at full resolution falls back to the downscaled view, so no image is dropped. Fewer images are downloaded at once while it is on, to keep memory in check.

### Fixed

- **Image-heavy chats are no longer dropped at full resolution.** With full-resolution images on, a chat with many large images packed close together could exceed a fixed limit on how much data the extension moves internally at once, and the whole conversation was lost. That data is now moved in safely sized pieces, and a single message holding more image data than fits at once has its images sent on their own and reassembled, so no chat or image is dropped.
- **A second person forwarding the same message is now kept.** Teams returns two rows for a single forward, which the exporter collapses to one. The rule that did this was too broad, so when two different people forwarded the same original message, the second was treated as a duplicate and dropped. It now matches only the true duplicate, so distinct forwards are all kept.
- **Attachments now show in TXT even when the message also has text.** A message that carried both body text and an attachment previously showed only the text in the TXT export. The attachment is now listed on the same line. Link previews are left out, since their URL already appears in the body.
- **Unfetched inline images render a clean placeholder in HTML.** An inline image that could not be downloaded (for example an old one the server no longer authorizes) rendered as a dead link that fails to load from a saved file. It now shows the same quiet "(not included)" placeholder as other unavailable images, and reads as "[image]" instead of "[file: image]" in TXT and CSV.
- **More images recovered on stricter tenants.** In tenants that reject the bearer-token image path, the first batch of images fetched before the extension switched to the cookie path could be lost. They are now retried on the cookie path and recovered.
- **Cleaner image labels.** Teams' generic placeholder alt text ("image", "undefined", and localized variants) no longer leaks into attachment labels, summaries, or saved filenames.

## [1.5.1] - 2026-06-23

Maintenance release. Nothing changes in how the extension behaves or what it exports. This release hardens the project's own supply chain and licensing. Each GitHub release now ships a signed build-provenance attestation (SLSA) and a SHA-256 checksum file, the build toolchain and GitHub Actions are version-pinned, and the generated font and emoji assets are hash-verified against the source. The repository now carries an MIT LICENSE file, matching what the README already stated. A dormant, log-only safeguard was added to the message fetcher that records, without blocking, any attempt to send a request off Microsoft's hosts; it does not affect exports.

## [1.5.0] - 2026-06-13

Large multi-chat exports now run faster and use far less memory. In a same-machine test a few-hundred-chat export finished about 1.7 times faster than 1.4.14, roughly an hour down to under 40 minutes while exporting more chats. The release also exports pasted tables with their row and column structure intact instead of flattening them to text, adds reactor avatars, a participant list and saved chat presets, and fixes a range of export-fidelity issues.

### Performance

- **Faster PDF building.** All text now draws through one path that reuses each embedded font instead of re-registering it on every line. The text-drawing work, the part that changed, runs nearly 4 times faster: a 1,000-page chat's text builds in about 1.5 seconds instead of 5.8 (image embedding and scraping are unchanged).
- **Lower memory on large bundles.** A multi-chat bundle used to keep every built chat plus all of the zip's compression state in memory until the very end, so memory climbed with each chat and a few-hundred-chat run could approach the browser's allocation limit. The bundle is now compressed one file at a time and streamed straight into the archive, so peak memory tracks the size of the archive being written rather than the sum of every chat. On a 292-chat run, peak working memory at the zip stage dropped from about 3 GB to under 600 MB.
- **Scraping overlaps building.** While one chat's files are being built, the next chat's messages are already being fetched. The two run in different parts of the extension, so they no longer wait on each other.
- **Batched directory lookups.** Names and profile photos are resolved through Microsoft Graph in batches of up to 20 per request instead of one request each. This is faster and removes the burst of per-photo console errors that large exports used to print.

### Added

- **Tables.** Tables pasted into Teams messages now export as real tables in every format (HTML, PDF, TXT, JSON, CSV) instead of being flattened to "cell | cell" text. Row and column structure, including merged cells, is preserved. Raised in [discussion #30](https://github.com/gediz/teams-web-chat-exporter/discussions/30).
- **Reactor avatars.** The reaction line in HTML and PDF now shows each reactor's profile photo, fetched through Microsoft Graph, falling back to initials when no photo is available. Your own reactions reuse the avatar already captured from Teams, with no extra network call.
- **Participant list.** A chat's PDF and HTML carry a "Participants" header line. Added and removed member names in system messages are resolved through Graph and the chat roster.
- **Chat presets.** Save the current chat selection as a named preset and re-apply it later from the picker. If a saved chat no longer exists, applying the preset re-selects the ones that remain and notes how many were unavailable.
- **Edited marker and reactions in TXT.** The plain-text export now appends "(edited)" to edited messages and a single-line reaction summary with reactor names, matching what the other formats already carried.
- **Forwarded column in CSV.** Forwarded messages kept their body out of the CSV text column; a dedicated forwarded column now carries the original author and text. Fields with line breaks follow RFC 4180 quoting instead of escaping to a literal "\n".
- **Clear button for the chat filter.** A Clear control empties the conversation search box in one click.
- **Verbose export stats (Diagnostics page).** An opt-in toggle, off by default. When on, the export-timing diagnostics written to the console include per-chat detail (chat names and per-stage timing) for performance debugging; otherwise those logs stay limited to ID-free totals.

### Fixed

- **CSV formula injection.** A cell beginning with `=`, `+`, `-`, `@`, tab or carriage return is prefixed with an apostrophe so spreadsheets treat it as text rather than running it as a formula (OWASP CSV Injection).
- **XML entities in system text.** System messages, recording titles and link-preview descriptions arrived entity-escaped (for example `&amp;`) and showed raw in TXT/JSON/CSV or double-escaped in HTML. They are decoded once now, and a single out-of-range numeric reference no longer discards the rest of the message.
- **Wide tables in PDF.** A very wide table used to scale its columns down until every cell wrapped to fragments. The cell font size is reduced to fit first, keeping values readable.
- **Mentions.** A mention split as "@First @Last" collapses to "@First Last", sender names win over mention names when both map to one person, and the mention list is coalesced and hardened.
- **Reaction line in PDF.** Reactor names are bounded with a dot drawn per reactor, and popover initials are centered.
- **Autolinked URLs.** A trailing CJK character or bracket next to a URL is no longer pulled into the link, and autolinked URLs are no longer double-escaped in HTML.
- **SharePoint sign-in pages saved as images.** An image whose bytes are actually a SharePoint sign-in page is rejected and shown as a labeled placeholder that links out, instead of a broken image.
- **One failed format no longer loses the others.** When a single format fails to build inside a bundle, the chat keeps its other formats instead of being dropped.
- **HTML layout.** Reply-connector dots are centered for own and compact replies, the header uses a middot separator, the reactor popover appears only for four or more reactors and is height-capped and scrollable, and compact-view thread replies are tidied.
- **Popup.** The popup is capped at 600 px so Chrome and Firefox show a single scrollbar, the progress bar no longer replays its fill animation when reopened mid-export, and the running state is written before reopen so the live status shows immediately.
- **PDF font subsets.** Subsets are padded so fontkit's empty-glyph bounding-box read stays in bounds.
- **Firefox diagnostics.** The diagnostics log size is shown even where `getBytesInUse` is unsupported.

### Changed

- **Shorter export helper text.** The line under the export button shows the source, format and date range only. The include-toggle list and the per-chat name were dropped because they overflowed and repeated information shown elsewhere.

### i18n

- New `presets.*` strings (button, save, apply, delete, and the partial-apply notice), translated across all 24 locales.
- New `diagnostics.verboseStats.*` strings for the Diagnostics toggle, English in all 24 locales for now (translations to follow).

## [1.4.14] - 2026-06-09

PDF font subsetting now runs in packaged builds. The service worker content security policy did not allow WebAssembly, so the HarfBuzz subsetter never started and every PDF embedded the full Noto fonts, about 9.5 MB of font data per file. With subsetting working a PDF keeps only the glyphs it uses, which drops the fonts to roughly 70 KB. A 1300-page chat went from 39 MB to 29 MB, the difference being fonts.

Korean text now renders in PDF instead of empty boxes, and very large exports download reliably on Chrome.

### Fixed

- **PDF font subsetting works in packaged builds.** The extension-pages CSP now grants `wasm-unsafe-eval` so `hb-subset.wasm` can compile in the service worker. A memory leak that would have exhausted the WASM heap after a few documents is fixed too (the font buffer is freed on every path).
- **Korean (Hangul) showed as tofu in PDF.** Korean routes to the bundled Noto Sans KR face; Chinese and other CJK ranges were corrected in the same change. (#28)
- **Large exports failed to download on Chrome.** MV3 service workers cannot create object URLs for large blobs, so multi-chat bundles and big single chats produced no file. The download URL is now minted in an offscreen document. (#27)
- **PDF links and reactions.** Links export as complete URLs with clickable annotations across wraps and page breaks. Reaction shortcodes resolve to emoji instead of leaking `:name:` text.
- **A failed retry leaked a blob URL.** When a single-file download failed and its retry also failed, the retry's blob URL is now revoked instead of leaking.

### Added

- **Reactor names in PDF**, following the HTML chip rule: one name, a short list, or "First and N others", with your own reaction as "You".
- **@mentions** render as `@name` in text and PDF.
- **Image placeholders.** An inline-image-only message reads `[inline image]`; an image that cannot be fetched shows a labeled placeholder card in HTML instead of a broken image.
- **Export button** reads "Export selected chat" when the single selected chat is not the one open in Teams.

### i18n

- New `actions.export.selected` string, translated across all 24 locales.

## [1.4.13] — 2026-05-21

GitHub-only release. Not pushed to the Chrome Web Store, Microsoft Edge Add-ons, or Firefox AMO. Existing installs stay on 1.4.12. Users who hit a problem and are asked to share diagnostics can install the unpacked build from this release; everyone else sees no change.

The headline addition is a Diagnostics page that produces a privacy-redacted JSON report users can attach to a bug report or share directly with the maintainer. Persistence is off by default; the feature is dormant until the user opens it.

### Added

- **Diagnostics page.** A stethoscope icon on the Settings page header opens a new Diagnostics page. The page builds a JSON report covering the extension version, browser, OS, Teams host, locale, declared and optional permissions, options snapshot, Teams data on disk (database names and store row counts; no record content), recent exports summary, and a console-log tail from the background service worker and content scripts. Save writes the JSON to disk; Copy puts the same JSON on the clipboard; Preview shows it inline.
- **Active probes.** Opt-in checklist that verifies Teams origin recognition, chat surface detection, IndexedDB access, Skype and IC3 token extraction, `asyncgw.teams.microsoft.com` and `asm.skype.com` reachability, page-world helper load, and a canary image fetch via the helper. Each probe reports pass / fail / skipped with status code or error string and elapsed milliseconds.
- **Log persistence (off by default).** When the user enables the toggle on the Diagnostics page, log lines from the SW and content scripts are flushed to `chrome.storage.local`. 8 MB byte cap; oldest entries dropped when full; `QuotaExceededError` triggers self-recovery (drops oldest half). Single-writer guard prevents concurrent flushes. Clear logs button wipes both in-memory and on-disk state. Default off means no disk footprint for users who never engage with the feature.

### Privacy and redaction

- Per-report random salt. Identifier-shaped substrings in the report are replaced with opaque tokens of the form `<kind a1b2c3d4>`. The salt lives only in memory during report build and is discarded afterwards. Two reports never share placeholders.
- Redacted shapes: UUIDs, Skype MRIs (`8:orgid:`, `8:live:`, `8:teamsvisitor:`, `gid:`, `28:`), email addresses, JWTs, Teams thread IDs (with and without the `@thread.v2` / `@unq.gbl.spaces` suffix, including truncated forms), AMS object identifiers (region + server + hex), SharePoint tenant subdomains, SharePoint `/personal/` user slugs (the `email_domain_tld` form that survives a naive email regex), and `asyncgw` / `asm.skype` regional hostnames. Verified by exhaustive scan: zero un-redacted identifiers across 1700+ string fields in a real export's diagnostic.
- "Include raw IDs" toggle (default off) for users who are sharing privately with the maintainer and want the original values.

### Fixed

- **Firefox MV2 manifest warning on load.** The `offscreen` permission was unconditionally declared, but it is MV3-only; Firefox MV2 emitted a permission warning on every load. Now gated to MV3 builds.
- **Popup page restoration was unreliable.** A reactive write to `LAST_PAGE_STORAGE_KEY` fired on initial render before the async restore could read the previously-saved page, overwriting `settings` / `history` / `diagnostics` with `main` and defeating the "resume where you left off" behaviour. Gated behind a hydrated flag so the persist write only runs after the restore has applied.

### i18n

- All 24 UI locale files updated with the new Diagnostics page strings and the probe-result labels. Languages with no translation for a new key fall back to English via the existing fallback chain.

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
