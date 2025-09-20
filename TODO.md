# Teams Exporter TODO

## Critical Issues
- [x] Remove all dry-run related code (popup, background, content script)
- [x] Preserve author carry-over across scroll passes
- [x] Prevent premature pagination exit when heights stabilize

## Scroll Loading Enhancements
- [x] Trigger Teams' "load older messages" UI elements during scroll
- [x] Persist `orderCtx.lastAuthor` between collection passes
- [x] Observe header sentinel with `IntersectionObserver` to detect true top
- [x] Add hard timeout/backoff when oldest message id stops changing
- [x] Ensure post-scroll hydration so empty bubbles/reactions are re-read before export

## HTML Export Improvements
- [x] Render locale-aware absolute timestamps with readable relative labels
- [x] Display replies as emphasized blockquotes with metadata
- [x] Surface attachment metadata (type, size, owner when available)
- [x] Add compact mode toggle (show/hide reactions & attachments)

## Other Export Types
- [x] Remove NDJSON export format (minimize maintenance)
- [x] Expand CSV reactions/attachments into dedicated columns
- [x] Remove Markdown export option and supporting code

## Scope Cleanup
- [x] Remove unused `util.js` helpers
- [x] Wire "Include threaded replies" toggle to export pipeline
- [x] Make HUD overlay optional via toggle
- [x] Deduplicate popup `buildAndDownload` logic
- [x] Move export orchestration from popup to service worker so downloads complete even if popup closes

## UX Polish
- [x] Disable Export button during runs and show spinner indicator
- [x] Display elapsed time in status area
- [x] Prefill `stopAt` input with last-used value
- [x] Provide inline validation/error banner pattern
- [x] Implement stricter option validation (e.g., ensure stop date is valid & Teams tab available)
- [x] Update extension badge with parsed-message counter during exports
- [x] Allow date range selection (start + end) for exports
- [x] Skip file download when no messages were collected (show banner instead)
- [x] Reset extension badge when tab is refreshed or background restarts
- [x] Humanize badge counts above 999 (e.g., 1.2k)

## Possible Optimizations
- [x] Declare an explicit `action.default_icon` in the manifest so Chrome always renders badge text
- [ ] Rate-limit badge updates so rapid scroll passes don’t spam `setBadge`
- [ ] Clear badge immediately when chat-context check fails before an export starts

## Future Enhancements
- [ ] Plan incremental exports (persist last timestamp, diff collection)
- [ ] Add participant filtering workflow
- [ ] Design summary sheet/report for export analytics
- [ ] Investigate PDF export option

## Telemetry (Future)
- [ ] Specify schema for session stats & environment diagnostics

## Browser Support
- [ ] Document Firefox port blockers & needed refactors
