# Teams Chat Exporter

Browser extension for exporting Microsoft Teams web chat data.

Supports Chrome, Edge, and Firefox.

## What it exports

- Formats: JSON, CSV, HTML, TXT
- Sources: chat conversations and team channels
- Optional data: replies, reactions, system messages, avatars, inline images (HTML flow)
- Date range filtering

## Practical notes

- Open a chat or team channel before exporting.
- Leave date fields empty to include all loaded history.
- If format is HTML and inline images are enabled, output is a zip package.
- Export runs by scrolling and scraping the Teams web UI, then downloads locally.

## Install

[![Available in the Chrome Web Store](docs/badges/available-in-the-chrome-web-store.png)](https://chromewebstore.google.com/detail/teams-chat-exporter/jmghclbfbbapimhbgnpffbimphlpolnm)

[![Get the Add-on for Firefox](docs/badges/mozilla-firefox-get-the-addon.png)](https://addons.mozilla.org/en-US/firefox/addon/teams-chat-exporter/)

Manual install: [docs/MANUAL_INSTALL.md](docs/MANUAL_INSTALL.md)

## Basic use

1. Open Teams on the web and open a chat or channel.
2. Click the extension icon.
3. Pick date range, format, and include options.
4. Click export.
5. Wait until scraping finishes. The download starts automatically.

## Development docs

- [docs/DEVELOPMENT.md](docs/DEVELOPMENT.md)
- [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md)
- [docs/CONTRIBUTING.md](docs/CONTRIBUTING.md)
- [docs/DEPLOYMENT.md](docs/DEPLOYMENT.md)
- [docs/TODO.md](docs/TODO.md)


> [!IMPORTANT]
> You are responsible for following your organization’s and Microsoft’s policies when exporting conversations.

