# Getting diagnostic logs

If I asked you to share diagnostic logs, here is how.

The diagnostic tool builds a JSON snapshot of the extension's state and a few active probe results. Real identifiers (UUIDs, tokens, email addresses, SharePoint tenant names) are replaced with random placeholders before save, so the output is safe to post in a public bug report.

## 1. Install the latest GitHub release

The diagnostic page is in the latest [GitHub release](https://github.com/gediz/teams-web-chat-exporter/releases/latest), which is not on the Chrome, Edge, or Firefox stores. Download the zip for your browser and follow the [manual install steps](MANUAL_INSTALL.md).

After installing you may see two extension icons: your store install and the new unpacked build. Use the unpacked one. It is the one with the Diagnostics page. If the two icons are confusing, disable the store install from `chrome://extensions` or `about:addons` until you are done.

## 2. Reproduce the issue

Run whatever triggers the problem.

## 3. Open the Diagnostics page

Open the extension popup, click the gear for Settings, then click the stethoscope icon in the top right of Settings.

![Diagnostics entry point on the Settings page](../screenshots/settings-light.png)

## 4. Run probes

Click "Run probes" and wait a few seconds. The rows populate with status for Teams origin, host reachability, token extraction, and the page-world helper.

![Diagnostics page with probes run](../screenshots/diagnostics-light.png)

## 5. Save the report

Click "Save to disk". A `.json` file lands in your downloads folder.

## 6. Share the file

Drag the file into the GitHub issue comment box, or use the paperclip / "Attach files" button below it.

---

## When to enable log persistence

If I asked you to capture more than one attempt of the same action (for example, "run the export twice and share"), enable persistence first. Otherwise the in-memory buffer is short-lived and the second attempt overwrites the first.

On the Diagnostics page, toggle "Save logs to disk" on. Then run the action the number of times I asked, then save and share.

## When to capture two reports

Some issues only happen under a specific condition (a corporate proxy, a feature toggle, a network environment). I may ask you to share two reports: one with the condition on, one with it off. Repeat steps 2 through 6 for each, then drop both `.json` files into the same reply.

## If I also ask for raw console output

The JSON filters logs to lines tagged with `[Teams Exporter]`, `[API]`, and similar prefixes, plus any `console.error`. Some failures land in the browser's own console first: CSP blocks, certificate warnings, network errors from a corporate proxy, exceptions from the popup. If I ask, capture both surfaces.

**Teams tab console**

1. Open the Teams web app in the tab where the issue happens.
2. Press F12 (or right-click and pick "Inspect").
3. Open the Console tab, then reproduce the issue.
4. Right-click in the console pane and pick "Save as…", or select the lines and copy.

**Service worker console**

Chrome / Edge:

1. Open `chrome://extensions` (or `edge://extensions`).
2. Find Teams Chat Exporter and click the **service worker** link.
3. In the DevTools window that opens, the Console tab is already selected.
4. Reproduce the issue, then save or copy the output.

Firefox:

1. Open `about:debugging` → **This Firefox**.
2. Find Teams Chat Exporter and click **Inspect**.
3. Open the Console tab, reproduce the issue, save or copy.

Raw console output is not redacted by the extension. It can contain real user IDs in URLs, authentication tokens, error stack traces, and noise from other extensions or Teams. Review before pasting publicly, or email it to <hello@teamschatexporter.com> instead.

---

## What ends up in the file

- Browser, OS, extension version, locale.
- Your extension settings.
- Permissions granted to the extension.
- Teams databases on your disk: names and row counts only. No record content.
- A summary of your last few exports: counts, timestamps, formats. No chat content.
- Console log lines from the extension's background and content scripts.
- Probe results.

## What does not end up in the file

- Your chat messages.
- Other people's names, email addresses, or contact details (placeholdered if they appear in log lines).
- Authentication tokens (hashed placeholders).
- The contents of Teams' local databases (only the database and store names).

Identifier-shaped substrings in log lines (UUIDs, Skype MRIs, email addresses, Teams thread IDs, AMS object IDs, SharePoint tenant subdomains, user slugs, regional hostnames, JWTs) are replaced with random placeholders of the form `<kind a1b2c3d4>`. Each report uses a fresh random salt, so two reports never share the same placeholders.

If anything still looks too sensitive for a public issue, open the file in a text editor and remove that part before sharing. The JSON is plain text. Or email it to me at <hello@teamschatexporter.com> instead.
