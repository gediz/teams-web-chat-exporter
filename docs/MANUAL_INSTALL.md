# Manual Installation

Use this when testing from source or from release files.

## From source code

### 1) Install dependencies

```bash
npm install
```

### 2) Build the target you need

```bash
# Chrome/Edge
npm run build

# Firefox
npm run build:firefox
```

## Load in Chrome / Edge

1. Open `chrome://extensions/` (or `edge://extensions/`).
2. Turn on **Developer mode**.
3. Click **Load unpacked**.
4. Select `.output/chrome-mv3/`.

## Load in Firefox (temporary)

1. Open `about:debugging#/runtime/this-firefox`.
2. Click **Load Temporary Add-on...**.
3. Select `.output/firefox-mv2/manifest.json`.

Firefox temporary add-ons are removed when Firefox restarts.

## From release downloads

- Chrome/Edge: extract zip, then load unpacked folder.
- Firefox: if the release contains Firefox build files, extract and load `manifest.json` as above.
