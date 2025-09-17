# Teams Exporter TODO

## Critical Issues
- [ ] Remove all dry-run related code (popup, background, content script)
- [ ] Preserve author carry-over across scroll passes
- [ ] Prevent premature pagination exit when heights stabilize

## Scroll Loading Enhancements
- [ ] Trigger Teams' "load older messages" UI elements during scroll
- [ ] Persist `orderCtx.lastAuthor` between collection passes
- [ ] Observe header sentinel with `IntersectionObserver` to detect true top
- [ ] Add hard timeout/backoff when oldest message id stops changing

## HTML Export Improvements
- [ ] Render locale-aware absolute timestamps with readable relative labels
- [ ] Display replies as emphasized blockquotes with metadata
- [ ] Surface attachment metadata (type, size, owner when available)
- [ ] Add compact mode toggle (show/hide reactions & attachments)

## Other Export Types
- [ ] Prepend NDJSON header record containing export metadata
- [ ] Expand CSV reactions/attachments into dedicated columns
- [ ] Enrich JSON export with computed fields (e.g., `timestampMs`, `dayBucket`)
- [ ] Remove Markdown export option and supporting code

## Scope Cleanup
- [ ] Remove unused `util.js` helpers
- [ ] Wire "Include threaded replies" toggle to export pipeline
- [ ] Make HUD overlay optional via toggle
- [ ] Deduplicate popup `buildAndDownload` logic

## Roadmap Clarifications
- [ ] Specify approach for Teams channel export support (selectors, metadata fields, UI toggle)

## UX Polish
- [ ] Disable Export button during runs and show spinner indicator
- [ ] Display elapsed time in status area (skip remaining estimate)
- [ ] Prefill `stopAt` input with last-used value
- [ ] Provide inline validation/error banner pattern

## Future Enhancements
- [ ] Plan incremental exports (persist last timestamp, diff collection)
- [ ] Add participant filtering workflow
- [ ] Design summary sheet/report for export analytics
- [ ] Explore in-place editing/marking of parsed messages prior to export

## Telemetry (Future)
- [ ] Specify schema for session stats & environment diagnostics

## Browser Support
- [ ] Document Firefox port blockers & needed refactors
