# TODO

## Planned

- [ ] Add lint + formatter setup (ESLint/Prettier).
- [ ] Add unit tests for utility and builder modules.
- [ ] Add popup UI tests.
- [ ] Add content-script tests with mocked DOM.
- [ ] Improve handling when long exports stall.
- [ ] Add pause/resume/stop export controls.
- [ ] Add participant filtering.
- [ ] Investigate PDF export options (library choice, layout fidelity, file size/performance, browser compatibility).
- [ ] Investigate exporting multiple chats in one run.
- [ ] Add README media (screenshots or GIF).
- [ ] Add user-configurable image-fetch domain allowlist (currently hardcoded in `src/content/attachments.ts`).
- [ ] Add user-configurable canvas size cap for embedded images (currently hardcoded at 4096x4096 in `src/content/attachments.ts`).
- [ ] Optional dismiss × on the persisted post-export outcome tile. Currently the tile clears on the next export start; a manual dismiss is only needed if users report wanting to hide it sooner.

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