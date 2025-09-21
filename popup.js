// ===== popup.js (replace) =====
const $ = (s) => document.querySelector(s);
const statusEl = $("#status");
const runBtn = $("#run");
const bannerEl = $("#banner");
const bannerMessageEl = bannerEl?.querySelector(".alert-message");
const quickRangeButtons = Array.from(document.querySelectorAll('#quickRanges .chip'));
const DAY_MS = 24 * 60 * 60 * 1000;

const DEFAULT_RUN_LABEL = runBtn.querySelector(".label")?.textContent || "Export current chat";
const BUSY_LABEL_EXPORTING = "Exporting…";
const BUSY_LABEL_BUILDING = "Building…";
const STORAGE_KEY = "teamsExporterOptions";
const ERROR_STORAGE_KEY = "teamsExporterLastError";
const EMPTY_RESULT_MESSAGE = "No messages found for the selected range.";

const controls = {
    startAt: $("#startAt"),
    endAt: $("#endAt"),
    format: $("#format"),
    includeReplies: $("#includeReplies"),
    includeReactions: $("#includeReactions"),
    includeSystem: $("#includeSystem"),
    embedAvatars: $("#embedAvatars"),
    showHud: $("#showHud"),
    themeToggle: $("#themeToggle")
};

const advancedToggleEl = document.getElementById("advancedToggle");
const advancedBodyEl = document.getElementById("advancedBody");

const DEFAULT_OPTIONS = {
    startAt: "",
    startAtISO: "",
    endAt: "",
    endAtISO: "",
    format: "json",
    includeReplies: true,
    includeReactions: true,
    includeSystem: false,
    embedAvatars: false,
    showHud: true,
    theme: "light"
};

let currentTabId = null;
let startedAtMs = null;
let elapsedTimerId = null;
let statusBaseText = "";

function isTeamsUrl(u) {
    return /^https:\/\/(.*\.)?(teams\.microsoft\.com|cloud\.microsoft)\//.test(u || "");
}

function applyTheme(theme) {
    const next = theme === "dark" ? "dark" : "light";
    document.body.dataset.theme = next;
    if (controls.themeToggle) {
        controls.themeToggle.checked = next === "dark";
    }
}

function currentTheme() {
    return controls.themeToggle?.checked ? "dark" : "light";
}

function setAdvancedExpanded(state) {
    if (!advancedToggleEl || !advancedBodyEl) return;
    advancedToggleEl.setAttribute("aria-expanded", state ? "true" : "false");
    if (state) {
        advancedBodyEl.hidden = false;
        advancedBodyEl.style.display = "flex";
    } else {
        advancedBodyEl.hidden = true;
        advancedBodyEl.style.display = "none";
    }
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
        hideErrorBanner({ clearStorage: true });
        setBusy(true, BUSY_LABEL_EXPORTING);
        setStatus("Preparing…");
        const tab = await getActiveTeamsTab();
        currentTabId = tab.id;
        await pingSW(); // ensure SW is alive

        const range = getValidatedRangeISO();
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
                scrapeOptions: { startAt: range.startISO, endAt: range.endISO, includeReplies, includeReactions, includeSystem, showHud },
                buildOptions: { format, saveAs: true, embedAvatars }
            }
        });

        if (response?.code === "EMPTY_RESULTS") {
            const message = response.error || EMPTY_RESULT_MESSAGE;
            setStatus(message, { stopElapsed: true });
            showErrorBanner(message, { persist: false });
            await clearPersistedError();
            return;
        }

        if (!response || response.error) {
            throw new Error(response?.error || "Export failed.");
        }

        setStatus(`Exported ${response.filename}`);
        hideErrorBanner({ clearStorage: true });
    } catch (e) {
        const msg = e?.message || "Export failed.";
        setStatus(msg);
        showErrorBanner(msg);
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
    applyTheme(opts.theme || DEFAULT_OPTIONS.theme);

    let startLocal = opts.startAt || isoToLocalInput(opts.startAtISO) || "";
    if (startLocal.includes('T') && !startLocal.includes(' ')) startLocal = startLocal.replace('T', ' ');
    controls.startAt.value = startLocal;

    let endLocal = opts.endAt || isoToLocalInput(opts.endAtISO) || "";
    if (endLocal.includes('T') && !endLocal.includes(' ')) endLocal = endLocal.replace('T', ' ');
    controls.endAt.value = endLocal;

    controls.format.value = opts.format || DEFAULT_OPTIONS.format;
    controls.includeReplies.checked = Boolean(opts.includeReplies);
    controls.includeReactions.checked = Boolean(opts.includeReactions);
    controls.includeSystem.checked = Boolean(opts.includeSystem);
    controls.embedAvatars.checked = Boolean(opts.embedAvatars);
    controls.showHud.checked = Boolean(opts.showHud);
    updateQuickRangeActive();
}

function collectOptions() {
    const startLocal = (controls.startAt.value || "").trim();
    const endLocal = (controls.endAt.value || "").trim();
    const startIso = startLocal ? localInputToISO(startLocal) : "";
    const endIso = endLocal ? localInputToISO(endLocal) : "";
    return {
        startAt: startLocal,
        startAtISO: startIso,
        endAt: endLocal,
        endAtISO: endIso,
        format: controls.format.value || DEFAULT_OPTIONS.format,
        includeReplies: controls.includeReplies.checked,
        includeReactions: controls.includeReactions.checked,
        includeSystem: controls.includeSystem.checked,
        embedAvatars: controls.embedAvatars.checked,
        showHud: controls.showHud.checked,
        theme: currentTheme()
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
            if (el === controls.themeToggle) {
                applyTheme(currentTheme());
            }
            persistOptions();
            if (el === controls.startAt || el === controls.endAt) updateQuickRangeActive();
        });
    }

    for (const btn of quickRangeButtons) {
        btn.addEventListener("click", () => {
            handleQuickRange(btn.dataset.range || "none");
        });
    }

    if (advancedToggleEl && advancedBodyEl) {
        advancedToggleEl.addEventListener("click", () => {
            const expanded = advancedToggleEl.getAttribute("aria-expanded") === "true";
            setAdvancedExpanded(!expanded);
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
        hideErrorBanner({ clearStorage: true });
        const startedAt = normalizeStart(msg.startedAt);
        setBusy(true, BUSY_LABEL_EXPORTING);
        setStatus("Starting export…", { startElapsedAt: startedAt });
    } else if (phase === "scrape:start") {
        setBusy(true, BUSY_LABEL_EXPORTING);
        setStatus("Running auto-scroll + scrape…");
    } else if (phase === "scrape:complete") {
        setBusy(true, BUSY_LABEL_BUILDING);
        setStatus(`Collected ${msg.messages ?? 0} messages. Building…`);
    } else if (phase === "empty") {
        const message = msg.message || EMPTY_RESULT_MESSAGE;
        setBusy(false);
        setStatus(message, { stopElapsed: true });
        showErrorBanner(message, { persist: false });
        clearPersistedError();
    } else if (phase === "complete") {
        setBusy(false);
        if (msg.filename) {
            setStatus(`Exported ${msg.filename}`, { stopElapsed: true });
        } else {
            setStatus("Export complete.", { stopElapsed: true });
        }
        hideErrorBanner({ clearStorage: true });
    } else if (phase === "error") {
        setBusy(false);
        setStatus(msg.error || "Export failed.", { stopElapsed: true });
        showErrorBanner(msg.error || "Export failed.");
    }
}

init();

async function init() {
    setBusy(false);
    const opts = await loadStoredOptions();
    applyOptions(opts);
    wireOptionPersistence();

    if (advancedToggleEl && advancedBodyEl) {
        setAdvancedExpanded(false);
    }

    const persistedError = await loadPersistedError();
    if (persistedError?.message) {
        showErrorBanner(persistedError.message, { persist: false });
        if (!statusBaseText) {
            setStatus(persistedError.message);
        }
    }

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

function getValidatedRangeISO() {
    const rawStart = (controls.startAt.value || "").trim();
    const rawEnd = (controls.endAt.value || "").trim();

    const startISO = rawStart ? localInputToISO(rawStart) : null;
    if (rawStart && !startISO) {
        const message = "Enter a valid start date/time.";
        showErrorBanner(message);
        throw new Error(message);
    }

    const endISO = rawEnd ? localInputToISO(rawEnd) : null;
    if (rawEnd && !endISO) {
        const message = "Enter a valid end date/time.";
        showErrorBanner(message);
        throw new Error(message);
    }

    if (startISO && endISO) {
        const startMs = Date.parse(startISO);
        const endMs = Date.parse(endISO);
        if (!Number.isNaN(startMs) && !Number.isNaN(endMs) && startMs > endMs) {
            const message = "Start date must be before end date.";
            showErrorBanner(message);
            throw new Error(message);
        }
    }

    return { startISO, endISO };
}

function localInputToISO(localValue) {
    if (!localValue) return "";
    let normalized = localValue.trim();
    if (!normalized) return "";
    normalized = normalized.replace(/\//g, '-');
    normalized = normalized.replace(/\s+/g, ' ');
    if (/^\d{4}-\d{2}-\d{2}$/.test(normalized)) {
        normalized += ' 00:00';
    }
    if (normalized.includes(' ')) {
        normalized = normalized.replace(' ', 'T');
    }
    const date = new Date(normalized);
    if (Number.isNaN(date.getTime())) return "";
    return date.toISOString();
}

function isoToLocalInput(isoValue) {
    if (!isoValue) return "";
    const date = new Date(isoValue);
    if (Number.isNaN(date.getTime())) return "";
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
    const hours = String(date.getHours()).padStart(2, "0");
    const minutes = String(date.getMinutes()).padStart(2, "0");
    return `${year}-${month}-${day} ${hours}:${minutes}`;
}

function handleQuickRange(range) {
    const normalized = range || "none";
    if (normalized === "none") {
        controls.startAt.value = "";
        controls.endAt.value = "";
        persistOptions();
        updateQuickRangeActive();
        return;
    }

    const now = new Date();
    let offsetMs = 0;
    if (normalized.endsWith("d")) {
        const days = Number(normalized.replace("d", ""));
        if (!Number.isNaN(days)) offsetMs = days * DAY_MS;
    }

    if (offsetMs > 0) {
        const startDate = new Date(now.getTime() - offsetMs);
        controls.startAt.value = isoToLocalInput(startDate.toISOString());
        controls.endAt.value = isoToLocalInput(now.toISOString());
    } else {
        controls.startAt.value = "";
        controls.endAt.value = "";
    }

    persistOptions();
    updateQuickRangeActive();
}

function updateQuickRangeActive() {
    if (!quickRangeButtons.length) return;
    const startISO = localInputToISO((controls.startAt.value || "").trim()) || null;
    const endISO = localInputToISO((controls.endAt.value || "").trim()) || null;
    const now = Date.now();
    const tolerance = 5 * 60 * 1000; // five minutes
    let active = "none";

    if (!startISO && !endISO) {
        active = "none";
    } else {
        const ranges = [
            { key: "1d", ms: DAY_MS },
            { key: "7d", ms: 7 * DAY_MS },
            { key: "30d", ms: 30 * DAY_MS }
        ];
        const endMs = endISO ? Date.parse(endISO) : now;
        const startMs = startISO ? Date.parse(startISO) : null;
        if (!Number.isNaN(endMs)) {
            for (const r of ranges) {
                const expectedStart = endMs - r.ms;
                const startOk = startMs != null && Math.abs(startMs - expectedStart) <= tolerance;
                const endOk = Math.abs(endMs - now) <= tolerance || (startISO && !endISO);
                if (startOk && endOk) {
                    active = r.key;
                    break;
                }
            }
        }
    }

    for (const btn of quickRangeButtons) {
        if ((btn.dataset.range || "none") === active) {
            btn.classList.add("active");
        } else {
            btn.classList.remove("active");
        }
    }
}

function showErrorBanner(message, { persist = true } = {}) {
    if (!bannerEl) return;
    const finalMessage = message || "Something went wrong.";
    const currentMessage = bannerMessageEl?.textContent || "";
    if (bannerEl.classList.contains("show") && currentMessage === finalMessage) {
        if (persist) persistError(finalMessage);
        return;
    }
    if (bannerMessageEl) bannerMessageEl.textContent = finalMessage;
    bannerEl.classList.add("show");
    if (persist) persistError(finalMessage);
}

function hideErrorBanner({ clearStorage = false } = {}) {
    if (!bannerEl) return;
    bannerEl.classList.remove("show");
    if (bannerMessageEl) bannerMessageEl.textContent = "";
    if (clearStorage) clearPersistedError();
}

async function persistError(message) {
    try {
        await chrome.storage.local.set({ [ERROR_STORAGE_KEY]: { message, timestamp: Date.now() } });
    } catch (_) {
        // ignore storage failures
    }
}

async function clearPersistedError() {
    try {
        await chrome.storage.local.remove(ERROR_STORAGE_KEY);
    } catch (_) {
        // ignore storage failures
    }
}

async function loadPersistedError() {
    try {
        const res = await chrome.storage.local.get(ERROR_STORAGE_KEY);
        return res?.[ERROR_STORAGE_KEY] || null;
    } catch (_) {
        return null;
    }
}
