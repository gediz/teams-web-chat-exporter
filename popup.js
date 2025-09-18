// ===== popup.js (replace) =====
const $ = (s) => document.querySelector(s);
const statusEl = $("#status");
const setStatus = (t) => (statusEl.textContent = t);
const runBtn = $("#run");

function isTeamsUrl(u) {
    return /^https:\/\/(.*\.)?(teams\.microsoft\.com|cloud\.microsoft)\//.test(u || "");
}

chrome.runtime.onMessage.addListener((msg) => {
    if (msg.type === "SCRAPE_PROGRESS") {
        const p = msg.payload || {};
        if (p.phase === "scroll") {
            setStatus(`Scrolling… pass ${p.passes} • height ${p.newHeight} • visible msgs ${p.messagesVisible}`);
        } else if (p.phase === "extract") {
            setStatus(`Extracting… found ${p.messagesExtracted} messages`);
        } else if (p.phase === "hud") {
            // no-op; HUD updates are in-page
        }
    } else if (msg.type === "EXPORT_STATUS") {
        const phase = msg.phase;
        if (phase === "starting") {
            setStatus("Starting export…");
        } else if (phase === "scrape:start") {
            setStatus("Running auto-scroll + scrape…");
        } else if (phase === "scrape:complete") {
            setStatus(`Collected ${msg.messages ?? 0} messages. Building…`);
        } else if (phase === "complete" && msg.filename) {
            setStatus(`Exported ${msg.filename}`);
        } else if (phase === "error") {
            setStatus(msg.error || "Export failed.");
        }
    }
});

async function getActiveTeamsTab() {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    if (!tab || !isTeamsUrl(tab.url)) throw new Error("Open the Teams web app tab first.");
    return tab;
}

runBtn.addEventListener("click", async () => {
    if (runBtn.disabled) return;
    const originalText = runBtn.querySelector(".label")?.textContent || "Export current chat";
    try {
        setBusy(true, "Exporting…");
        setStatus("Preparing…");
        const tab = await getActiveTeamsTab();
        await pingSW(); // ensure SW is alive

        const stopAt = $("#stopAt").value ? new Date($("#stopAt").value).toISOString() : null;
        const format = $("#format").value;
        const includeReplies = $("#includeReplies").checked;
        const includeReactions = $("#includeReactions").checked;
        const includeSystem = $("#includeSystem").checked;
        const embedAvatars = $("#embedAvatars").checked;
        const showHud = $("#showHud").checked;

        setStatus("Export running… you can close this popup.");
        const response = await chrome.runtime.sendMessage({
            type: "START_EXPORT",
            data: {
                tabId: tab.id,
                scrapeOptions: { stopAt, includeReplies, includeReactions, includeSystem, showHud },
                buildOptions: { format, saveAs: true, embedAvatars }
            }
        });

        if (!response || response.error) {
            throw new Error(response?.error || "Export failed.");
        }

        setStatus(`Exported ${response.filename}`);
    } catch (e) {
        setStatus(e.message);
    } finally {
        setBusy(false, originalText);
    }
});

async function pingSW(timeoutMs = 4000) {
    return Promise.race([
        chrome.runtime.sendMessage({ type: "PING_SW" }),
        new Promise((_, rej) => setTimeout(() => rej(new Error("No response from background (PING_SW timeout)")), timeoutMs))
    ]);
}

function setBusy(state, labelText) {
    const labelEl = runBtn.querySelector(".label");
    if (state) {
        runBtn.classList.add("busy");
        runBtn.disabled = true;
        if (labelEl) labelEl.textContent = labelText;
    } else {
        runBtn.classList.remove("busy");
        runBtn.disabled = false;
        if (labelEl) labelEl.textContent = labelText;
    }
}
