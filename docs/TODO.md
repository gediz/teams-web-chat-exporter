# TODO

## Planned

- [ ] Add lint + formatter setup (ESLint/Prettier).
- [ ] Add unit tests for utility and builder modules.
- [ ] Add popup UI tests.
- [ ] Add content-script tests with mocked DOM.
- [ ] Improve handling when long exports stall (auto-retry / diagnostics hint).
- [ ] Add participant filtering (export only messages from selected authors).
- [ ] Investigate exporting multiple chats in one run.
- [ ] Add README media (screenshots or GIF).
- [ ] Add user-configurable image-fetch domain allowlist (currently hardcoded in `src/content/attachments.ts`).
- [ ] Add user-configurable per-image pixel cap for embedded images (currently the filter drops anything over `4096 * 4096` total pixels in `src/content/attachments.ts`).
- [ ] Localize PDF timestamps per `options.lang` (currently fixed `YYYY-MM-DD HH:MM`). Small — use `Intl.DateTimeFormat` with the user's lang.
- [ ] Shrink CJK coverage further. Runtime HarfBuzz subsetting already trims the bundled `NotoSansSC` per-export; open question is whether to ship a smaller CJK base font or fetch glyphs on demand for users who never export CJK content.
- [ ] Native-speaker review pass over the 22 non-English, non-Turkish locales (they're at full parity but translations are my own, not native).

## Done

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