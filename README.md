# Teams Chat Exporter

A browser extension that exports chat conversations from the Microsoft Teams web application. Supports Chrome, Edge, and Firefox.

## Features

- **Export Formats**: JSON, CSV, HTML, Text.
- **Comprehensive Data**: Messages, reactions, threaded replies, system messages.
- **Date Filtering**: Filter messages by date range.
- **Avatar Embedding**: Embed avatars as base64 images in HTML exports.
- **Auto-scroll**: Automatically scrolls to load chat history.
- **Cross-Browser**: Compatible with Chrome, Edge, and Firefox.

## Installation

### Chrome / Edge
1. Open the [Teams Chat Exporter listing](https://chromewebstore.google.com/detail/teams-chat-exporter/jmghclbfbbapimhbgnpffbimphlpolnm) in the Chrome Web Store.
2. Click "Add to Chrome".

### Firefox
1. Open the [Teams Chat Exporter listing](https://addons.mozilla.org/en-US/firefox/addon/teams-chat-exporter/) in Firefox Add-ons.
2. Click "Add to Firefox".

### Manual Installation
To install from a release ZIP or source code, see the [Manual Installation Guide](docs/MANUAL_INSTALL.md).

## Usage

1. Navigate to the Microsoft Teams web app.
2. Open the chat conversation you want to export.
3. Click the extension icon in the toolbar.
4. Configure export options:
   - **Date Range**: Select a preset or custom range.
   - **Format**: JSON, CSV, HTML, or Text.
   - **Include**: Toggle replies, reactions, or system messages.
5. Click "Export current chat".
6. Wait for the extension to scroll and collect messages.
7. The file will download automatically when complete.

> [!IMPORTANT]
> You are responsible for complying with your organization’s and Microsoft’s terms and policies when exporting conversations.

## Development

This project uses the WXT Framework for cross-browser support.

For development instructions, architecture, and testing guides, see [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md).

- **Source Code**: Located in the `src` directory.
- **Contribution Guide**: See [docs/CONTRIBUTING.md](docs/CONTRIBUTING.md).
