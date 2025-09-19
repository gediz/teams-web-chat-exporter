// ===== popup.js (replace) =====
const $ = (s) => document.querySelector(s);
const statusEl = $("#status");
const runBtn = $("#run");

const DEFAULT_RUN_LABEL = runBtn.querySelector(".label")?.textContent || "Export current chat";
const BUSY_LABEL_EXPORTING = "Exporting…";
const BUSY_LABEL_BUILDING = "Building…";
const STORAGE_KEY = "teamsExporterOptions";

const controls = {
    stopAt: $("#stopAt"),
    format: $("#format"),
    includeReplies: $("#includeReplies"),
    includeReactions: $("#includeReactions"),
    includeSystem: $("#includeSystem"),
    embedAvatars: $("#embedAvatars"),
    showHud: $("#showHud")
};

const DEFAULT_OPTIONS = {
    stopAt: "",
    format: "json",
    includeReplies: true,
    includeReactions: true,
    includeSystem: false,
    embedAvatars: false,
    showHud: true
};

let currentTabId = null;
let startedAtMs = null;
let elapsedTimerId = null;
let statusBaseText = "";

function isTeamsUrl(u) {
    return /^https:\/\/(.*\.)?(teams\.microsoft\.com|cloud\.microsoft)\//.test(u || "");
}

chrome.runtime.onMessage.addListener((msg) => {
    if (msg.type === "SCRAPE_PROGRESS") {
        const p = msg.payload || {};
        if (p.phase === "scroll") {
            const seen = p.seen ?? p.aggregated ?? p.messagesVisible ?? 0;
            setStatus(`Scrolling… pass ${p.passes} • seen ${seen}`);
        } else if (p.phase === "extract") {
            setStatus(`Extracting… found ${p.messagesExtracted} messages`);
        } else if (p.phase === "hud") {
            // no-op; HUD updates are in-page
        }
    } else if (msg.type === "EXPORT_STATUS") {
        handleExportStatus(msg);
    }
});

async function getActiveTeamsTab() {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    if (!tab || !isTeamsUrl(tab.url)) throw new Error("Open the Teams web app tab first.");
    return tab;
}

runBtn.addEventListener("click", async () => {
    if (runBtn.disabled) return;
    try {
        setBusy(true, BUSY_LABEL_EXPORTING);
        setStatus("Preparing…");
        const tab = await getActiveTeamsTab();
        currentTabId = tab.id;
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
        setBusy(false);
    }
});

async function pingSW(timeoutMs = 4000) {
    return Promise.race([
        chrome.runtime.sendMessage({ type: "PING_SW" }),
        new Promise((_, rej) => setTimeout(() => rej(new Error("No response from background (PING_SW timeout)")), timeoutMs))
    ]);
}

function setBusy(state, labelText = BUSY_LABEL_EXPORTING) {
    const labelEl = runBtn.querySelector(".label");
    if (state) {
        runBtn.classList.add("busy");
        runBtn.disabled = true;
        if (labelEl) labelEl.textContent = labelText || BUSY_LABEL_EXPORTING;
    } else {
        runBtn.classList.remove("busy");
        runBtn.disabled = false;
        if (labelEl) labelEl.textContent = DEFAULT_RUN_LABEL;
    }
}

async function loadStoredOptions() {
    try {
        const stored = await chrome.storage.local.get(STORAGE_KEY);
        return { ...DEFAULT_OPTIONS, ...(stored?.[STORAGE_KEY] || {}) };
    } catch (_) {
        return { ...DEFAULT_OPTIONS };
    }
}

function applyOptions(opts) {
    controls.stopAt.value = opts.stopAt || "";
    controls.format.value = opts.format || DEFAULT_OPTIONS.format;
    controls.includeReplies.checked = Boolean(opts.includeReplies);
    controls.includeReactions.checked = Boolean(opts.includeReactions);
    controls.includeSystem.checked = Boolean(opts.includeSystem);
    controls.embedAvatars.checked = Boolean(opts.embedAvatars);
    controls.showHud.checked = Boolean(opts.showHud);
}

function collectOptions() {
    return {
        stopAt: controls.stopAt.value || "",
        format: controls.format.value || DEFAULT_OPTIONS.format,
        includeReplies: controls.includeReplies.checked,
        includeReactions: controls.includeReactions.checked,
        includeSystem: controls.includeSystem.checked,
        embedAvatars: controls.embedAvatars.checked,
        showHud: controls.showHud.checked
    };
}

async function persistOptions() {
    try {
        await chrome.storage.local.set({ [STORAGE_KEY]: collectOptions() });
    } catch (_) {
        // no-op if storage fails
    }
}

function wireOptionPersistence() {
    const inputs = Object.values(controls).filter(Boolean);
    for (const el of inputs) {
        el.addEventListener("change", () => {
            persistOptions();
        });
    }
}

function handleExportStatus(msg) {
    const tabId = msg?.tabId;
    if (typeof tabId === "number") {
        if (currentTabId && tabId !== currentTabId) return;
        if (!currentTabId) currentTabId = tabId;
    }

    const phase = msg?.phase;
    if (phase === "starting") {
        const startedAt = normalizeStart(msg.startedAt);
        setBusy(true, BUSY_LABEL_EXPORTING);
        setStatus("Starting export…", { startElapsedAt: startedAt });
    } else if (phase === "scrape:start") {
        setBusy(true, BUSY_LABEL_EXPORTING);
        setStatus("Running auto-scroll + scrape…");
    } else if (phase === "scrape:complete") {
        setBusy(true, BUSY_LABEL_BUILDING);
        setStatus(`Collected ${msg.messages ?? 0} messages. Building…`);
    } else if (phase === "complete") {
        setBusy(false);
        if (msg.filename) {
            setStatus(`Exported ${msg.filename}`, { stopElapsed: true });
        } else {
            setStatus("Export complete.", { stopElapsed: true });
        }
    } else if (phase === "error") {
        setBusy(false);
        setStatus(msg.error || "Export failed.", { stopElapsed: true });
    }
}

init();

async function init() {
    setBusy(false);
    const opts = await loadStoredOptions();
    applyOptions(opts);
    wireOptionPersistence();

    try {
        const tab = await getActiveTeamsTab();
        currentTabId = tab.id;
        const status = await chrome.runtime.sendMessage({ type: "GET_EXPORT_STATUS", tabId: currentTabId });
        if (status?.active) {
            const last = status.info?.lastStatus;
            const startedAt = normalizeStart(status.info?.startedAt);
            if (startedAt) {
                startedAtMs = startedAt;
                ensureElapsedTimer();
            }
            if (last) {
                handleExportStatus(last);
            } else {
                setBusy(true, BUSY_LABEL_EXPORTING);
                setStatus("Export running…");
            }
        }
    } catch (err) {
        // Not on Teams tab; ignore until user clicks Export
    }
}

function setStatus(text, { startElapsedAt, stopElapsed } = {}) {
    statusBaseText = text || "";

    if (typeof startElapsedAt === "number" && !Number.isNaN(startElapsedAt)) {
        startedAtMs = startElapsedAt;
        ensureElapsedTimer();
        updateStatusText();
        return;
    }

    if (stopElapsed) {
        if (startedAtMs) {
            const finalText = `${statusBaseText}${formatElapsedSuffix(Date.now() - startedAtMs)}`;
            statusBaseText = finalText;
        }
        startedAtMs = null;
        clearElapsedTimer();
        statusEl.textContent = statusBaseText;
        return;
    }

    updateStatusText();
}

function updateStatusText() {
    let text = statusBaseText;
    if (startedAtMs) {
        text += formatElapsedSuffix(Date.now() - startedAtMs);
    }
    statusEl.textContent = text;
}

function ensureElapsedTimer() {
    if (elapsedTimerId != null) return;
    elapsedTimerId = setInterval(() => {
        if (!startedAtMs) {
            clearElapsedTimer();
            return;
        }
        updateStatusText();
    }, 1000);
    updateStatusText();
}

function clearElapsedTimer() {
    if (elapsedTimerId != null) {
        clearInterval(elapsedTimerId);
        elapsedTimerId = null;
    }
}

function formatElapsedSuffix(ms) {
    return ` — Elapsed: ${formatElapsed(ms)}`;
}

function formatElapsed(ms) {
    const totalSeconds = Math.max(0, Math.floor(ms / 1000));
    const hours = Math.floor(totalSeconds / 3600);
    const minutes = Math.floor((totalSeconds % 3600) / 60);
    const seconds = totalSeconds % 60;
    if (hours > 0) {
        return `${hours}:${String(minutes).padStart(2, "0")}:${String(seconds).padStart(2, "0")}`;
    }
    return `${minutes}:${String(seconds).padStart(2, "0")}`;
}

function normalizeStart(value) {
    if (typeof value === "number" && !Number.isNaN(value)) return value;
    if (typeof value === "string") {
        const parsed = Number(value);
        if (!Number.isNaN(parsed)) return parsed;
        const date = Date.parse(value);
        if (!Number.isNaN(date)) return date;
    }
    return null;
}
