# Manual Installation Guide

This guide explains how to install the Teams Chat Exporter extension manually.

## Option 1: Install from Release ZIP

If you downloaded a `.zip` file from the [Releases page](https://github.com/gediz/teams-web-chat-exporter/releases):

### Chrome / Edge
1. Extract the ZIP file.
2. Navigate to `chrome://extensions/`.
3. Enable **Developer mode**.
4. Click **Load unpacked**.
5. Select the extracted folder.

### Firefox
1. Navigate to `about:debugging#/runtime/this-firefox`.
2. Click **Load Temporary Add-on...**.
3. Select the ZIP file (or `manifest.json` inside the folder).

## Option 2: Load Unpacked (From Source)

1. Build the project:
   ```bash
   npm install
   npm run build
   ```
   This creates an `.output/` directory.

### Chrome / Edge
1. Navigate to `chrome://extensions/`.
2. Enable **Developer mode**.
3. Click **Load unpacked**.
4. Select the `.output/chrome-mv3` directory.

### Firefox
1. Navigate to `about:debugging#/runtime/this-firefox`.
2. Click **Load Temporary Add-on...**.
3. Select the `manifest.json` file inside `.output/firefox-mv2`.
