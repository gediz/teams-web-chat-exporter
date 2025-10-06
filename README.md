# Teams Chat Exporter

![Teams Chat Exporter icon](icons/action-128.png)

A Chrome browser extension that exports chat conversations from Microsoft Teams web application.

## Features

- **Export Formats**: JSON, CSV, HTML
- **Comprehensive Data**: Messages, reactions, threaded replies, system messages
- **Date Filtering**: Set oldest date to limit chat history
- **Avatar Embedding**: Option to embed avatars as base64 (HTML format)
- **Auto-scroll**: Automatically loads chat history by scrolling

## Installation

### Chrome Web Store

1. Open the [Teams Chat Exporter listing](https://chromewebstore.google.com/detail/teams-chat-exporter/jmghclbfbbapimhbgnpffbimphlpolnm) in Chrome.
2. Click `Add to Chrome`, then confirm by selecting `Add extension`.
3. Pin the extension for easy access from the toolbar.

### Manual (Unpacked) Installation

1. Clone or download this repository.
2. Open Chrome and navigate to `chrome://extensions/`.
3. Enable `Developer mode` (toggle in top-right).
4. Click `Load unpacked` and select the extension directory.
5. Pin the extension for easy access.

## Usage

1. Navigate to Microsoft Teams web app (`teams.microsoft.com`).
2. Open the chat conversation you want to export.
3. Click the extension icon in Chrome toolbar.
4. Configure export options:
   - Set stop date (optional) to limit how far back to export
   - Choose export format
   - Select what to include (replies, reactions, system messages)
   - Enable avatar embedding (HTML only)
5. Click `Export current chat`.
6. Wait for the extension to scroll and collect messages.
7. File will be automatically downloaded.

## Export Options

- **Stop at date**: Limits export to messages newer than specified date
- **Include threaded replies**: Exports reply context information
- **Include reactions**: Exports emoji reactions and participant lists
- **Include system messages**: Exports date dividers and system notifications
- **Embed avatars**: Downloads and embeds profile pictures (HTML format only)

## Permissions

The extension requires:
- `activeTab`: Access current Teams tab
- `downloads`: Save exported files
- `storage`: Remember user preferences
- `scripting`: Inject content scripts
- Host access to Microsoft Teams domains

## File Structure

- `manifest.json` - Extension configuration
- `popup.html/js` - User interface
- `content.js` - Teams page scraping logic
- `service-worker.js` - Background processing and file generation

## Notes

- Works only with Microsoft Teams web application
- Export time depends on chat history length
- Large exports may take several minutes to complete
- Extension preserves message order and timestamps

