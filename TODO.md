# Project Roadmap

## Remaining / Follow-ups
- [ ] Tooling: Set up linting and formatting (ESLint/Prettier).
- [ ] Tests: Add unit tests for helpers and E2E tests for the export flow.
- [ ] Refactoring: Further module extraction in content script (attachments, reactions, parse helpers).
- [ ] Investigation: Investigate occasional stalls in HUD/elapsed timer.
- [ ] Add an option to pause/resume/stop export operation.
- [ ] Add participant filtering workflow
- [ ] Design summary sheet/report for export analytics
- [ ] Investigate PDF export option
- [ ] Add an option to export all available chats
- [ ] Add screenshot/GIF, version badge to README.md

## Telemetry (Future)
- [ ] Specify schema for session stats & environment diagnostics

## Known Issues
1. **Content Script Injection**: The manual `chrome.scripting.executeScript` fallback in `background.js` relies on WXT outputting a specific filename (`content.js`). This may need adjustment if build configuration changes.
2. **Data URL Size**: Large HTML exports may exceed browser Data URL limits (approx. 50MB in Chrome). Firefox uses Blob URLs which are safer.
