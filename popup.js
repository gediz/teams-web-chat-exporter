// ===== popup.js (replace) =====
const $ = (s) => document.querySelector(s);
const statusEl = $("#status");
const setStatus = (t) => (statusEl.textContent = t);

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
    }
});

async function getActiveTeamsTab() {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    if (!tab || !isTeamsUrl(tab.url)) throw new Error("Open the Teams web app tab first.");
    return tab;
}

// Try to ping the content script. If it doesn't respond, inject it, then try again.
async function ensureContentScript(tabId) {
    try {
        const pong = await chrome.tabs.sendMessage(tabId, { type: "PING" });
        if (pong && pong.ok) return;
    } catch (_) {
        // no listener yet
    }
    await chrome.scripting.executeScript({
        target: { tabId, allFrames: true },
        files: ["content.js"],
    });
    // ping again
    const pong2 = await chrome.tabs.sendMessage(tabId, { type: "PING" });
    if (!pong2 || !pong2.ok) {
        throw new Error("Content script did not load in this tab/frame.");
    }
}

$("#run").addEventListener("click", async () => {
    try {
        setStatus("Preparing…");
        const tab = await getActiveTeamsTab();
        await ensureContentScript(tab.id);
        await pingSW(); // ensure SW is alive

        const stopAt = $("#stopAt").value ? new Date($("#stopAt").value).toISOString() : null;
        const format = $("#format").value;
        const includeReplies = $("#includeReplies").checked;
        const includeReactions = $("#includeReactions").checked;
        const includeSystem = $("#includeSystem").checked;
        const embedAvatars = $("#embedAvatars").checked;

        setStatus("Running auto-scroll + scrape…");
        const res = await chrome.tabs.sendMessage(tab.id, {
            type: "SCRAPE_TEAMS",
            options: { stopAt, includeReplies, includeReactions, includeSystem }
        });

        setStatus(`Collected ${res.messages.length} messages. Building ${format.toUpperCase()}…`);
        const payload = await Promise.race([
            chrome.runtime.sendMessage({
                type: "BUILD_AND_DOWNLOAD",
                data: { messages: res.messages, meta: res.meta, format, saveAs: true, embedAvatars }

            }),
            new Promise((_, rej) => setTimeout(() => rej(new Error("Background did not respond (BUILD_AND_DOWNLOAD timeout).")), 15000))
        ]);

        if (!payload || payload.error) throw new Error(payload?.error || "Failed to build & download.");
        setStatus(`Exported ${payload.filename}`);
    } catch (e) {
        setStatus(e.message);
    }
});

async function pingSW(timeoutMs = 4000) {
    return Promise.race([
        chrome.runtime.sendMessage({ type: "PING_SW" }),
        new Promise((_, rej) => setTimeout(() => rej(new Error("No response from background (PING_SW timeout)")), timeoutMs))
    ]);
}
