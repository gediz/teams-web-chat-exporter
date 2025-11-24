import { defineBackground } from 'wxt/sandbox';
import { createBadgeManager } from '../utils/badge';
import { formatRangeLabel, sanitizeBase } from '../utils/messages';
import { formatDayLabelForExport, parseTimeStamp } from '../utils/time';
import type {
  ActiveExportInfo,
  BuildOptions,
  ExportMessage,
  ExportMeta,
  ExportStatusPayload,
  Reaction,
  ScrapeOptions,
  ScrapeResult,
} from '../types/shared';
import { embedAvatarsInRows, textToBlobUrl, textToDataUrl, toCSV, toHTML } from '../background/builders';

/* eslint-disable @typescript-eslint/no-explicit-any */
// Typed globals for Firefox builds
declare const browser: typeof chrome | undefined;

// ===== service-worker.js (WXT version) =====
export default defineBackground(() => {
// Browser API compatibility for Firefox
const runtime = typeof browser !== 'undefined' ? browser.runtime : chrome.runtime;
const tabs = typeof browser !== 'undefined' ? browser.tabs : chrome.tabs;
// Firefox MV2 uses browserAction, Chrome MV3 uses action
const action = typeof browser !== 'undefined'
    ? (browser.action || browser.browserAction)
    : chrome.action;
const downloads = typeof browser !== 'undefined' ? browser.downloads : chrome.downloads;
const scripting = typeof browser !== 'undefined' ? browser.scripting : chrome.scripting;
const badge = createBadgeManager(action);
const { set: setBadge, reset: resetBadge, clearSoon: clearBadgeSoon, updateForStatus: updateBadgeForStatus, updateForProgress: updateBadgeForProgress } = badge;

function log(...a: unknown[]) { try { console.log("[Teams Exporter SW]", ...a) } catch { } }
log("boot");

runtime.onInstalled.addListener(() => {
    log("onInstalled");
    resetBadge();
});
runtime.onStartup?.addListener(() => {
    log("onStartup");
    resetBadge();
});

tabs.onUpdated.addListener((tabId: number, changeInfo: chrome.tabs.TabChangeInfo, tab?: chrome.tabs.Tab) => {
    const nextUrl = changeInfo.url ?? tab?.url;
    if (changeInfo.status === 'loading' && isTeamsUrl(nextUrl)) {
        activeExports.delete(tabId);
        resetBadge();
    }
});

const activeExports = new Map<number, ActiveExportInfo>(); // tabId -> { startedAt, lastStatus }
// TERMINAL_PHASES: 'complete' = success, 'error' = failure, 'empty' = no data found (not a failure)
const TERMINAL_PHASES = new Set(['complete', 'error', 'empty']);
const TEAMS_URL_PATTERN = /^https:\/\/(.*\.)?(teams\.microsoft\.com|cloud\.microsoft)\//i;

function isTeamsUrl(url: string | null | undefined): boolean {
    return TEAMS_URL_PATTERN.test(url || '');
}

function updateActiveExport(tabId: number, patch: Partial<ActiveExportInfo> = {}) {
    if (tabId == null) return;
    const prev = activeExports.get(tabId) || {};
    const next: ActiveExportInfo = { ...prev, ...patch };
    activeExports.set(tabId, next);
    return next;
}

// Detect if we're in Firefox
const isFirefox = typeof browser !== 'undefined' && navigator.userAgent.includes('Firefox');

const sendMessageToTab = (tabId: number, msg: unknown) => new Promise<any>((resolve, reject) => {
    tabs.sendMessage(tabId, msg, (resp) => {
        const err = runtime.lastError;
        if (err) {
            reject(new Error(err.message || 'Failed to reach tab context'));
            return;
        }
        resolve(resp);
    });
});

async function ensureContentScript(tabId: number) {
    try {
        const pong = await sendMessageToTab(tabId, { type: 'PING' });
        if (pong?.ok) return;
    } catch (_) {
        // fallback to injection
    }
    await scripting.executeScript({ target: { tabId, allFrames: true }, files: ['content.js'] });
    const pong2 = await sendMessageToTab(tabId, { type: 'PING' });
    if (!pong2?.ok) throw new Error('Content script did not respond after injection');
}

async function requestScrape(tabId: number, options: ScrapeOptions): Promise<ScrapeResult> {
    const res = await sendMessageToTab(tabId, { type: 'SCRAPE_TEAMS', options });
    if (!res) throw new Error('No response from content script');
    if (res.error) throw new Error(res.error);
    return res;
}

async function buildAndDownload({ messages = [], meta = {}, format = 'json', saveAs = true, embedAvatars = false }: { messages?: ExportMessage[]; meta?: ExportMeta; format?: 'json' | 'csv' | 'html'; saveAs?: boolean; embedAvatars?: boolean }) {
    let rows = messages;
    if (format === 'html' && embedAvatars) {
        try {
            rows = await embedAvatarsInRows(messages);
        } catch (e: any) {
            log('embed avatars failed', e?.message || e);
            rows = messages;
        }
    }

    const rangeLabel = formatRangeLabel(meta.startAt, meta.endAt);
    const enrichedMeta = { ...meta };
    if (rangeLabel) enrichedMeta.timeRange = rangeLabel;

    const baseTitle = sanitizeBase(enrichedMeta.title || 'teams-chat');
    const stamp = new Date().toISOString().replace(/[:.]/g, '-');
    const base = `${baseTitle}-${stamp}`;
    let filename, mime, content;

    if (format === 'json') {
        filename = `${base}.json`;
        mime = 'application/json';
        const payload = { meta: { ...enrichedMeta, count: messages.length }, messages };
        content = JSON.stringify(payload, null, 2);
    } else if (format === 'csv') {
        filename = `${base}.csv`;
        mime = 'text/csv';
        content = toCSV(messages);
    } else if (format === 'html') {
        filename = `${base}.html`;
        mime = 'text/html';
        content = toHTML(rows, { ...enrichedMeta, count: messages.length });
    } else {
        throw new Error('Unknown format: ' + format);
    }

    // Firefox requires blob URLs, Chrome works with both
    const url = isFirefox ? textToBlobUrl(content, mime) : textToDataUrl(content, mime);
    log('download build', { format, bytes: content.length, filename, browser: isFirefox ? 'firefox' : 'chrome' });

    try {
        const id = await downloads.download({ url, filename, saveAs });
        log('download started', { id, filename });

        // Cleanup blob URL after download starts (Firefox only)
        if (isFirefox) {
            setTimeout(() => URL.revokeObjectURL(url), 60000); // Revoke after 1 minute
        }

        return { ok: true, filename, id };
    } catch (e: any) {
        log('download primary failed', e?.message || e);
        const safe = `${sanitizeBase('teams-chat')}-${Date.now()}.${format === 'html' ? 'html' : format === 'csv' ? 'csv' : 'json'}`;

        // Cleanup failed blob URL
        if (isFirefox) {
            URL.revokeObjectURL(url);
        }

        try {
            const url2 = isFirefox ? textToBlobUrl(content, mime) : textToDataUrl(content, mime);
            const id2 = await downloads.download({ url: url2, filename: safe, saveAs });

            // Cleanup blob URL after download starts (Firefox only)
            if (isFirefox) {
                setTimeout(() => URL.revokeObjectURL(url2), 60000);
            }

            return { ok: true, filename: safe, id: id2 };
        } catch (e2: any) {
            log('download fallback failed', e2?.message || e2);
            throw new Error(e2?.message || String(e2));
        }
    }
}

function broadcastStatus(payload: ExportStatusPayload) {
    let enriched = { ...payload };
    const tabId = payload?.tabId;
    if (tabId != null) {
        const phase = payload?.phase;
        let info;
        if (phase) {
            const record = { ...payload };
            if (TERMINAL_PHASES.has(phase)) {
                info = updateActiveExport(tabId, { lastStatus: record, phase, completedAt: Date.now() });
            } else {
                info = updateActiveExport(tabId, { lastStatus: record, phase });
            }
        } else {
            info = updateActiveExport(tabId, { lastStatus: { ...payload } });
        }
        const startedAt = info?.startedAt;
        if (startedAt && enriched.startedAt == null) {
            enriched = { ...enriched, startedAt };
        }
    }
    // Firefox compatibility: wrap in try-catch since sendMessage may not return a Promise
    try {
        const msgPromise = runtime.sendMessage({ type: 'EXPORT_STATUS', ...enriched });
        if (msgPromise && msgPromise.catch) msgPromise.catch(() => { });
    } catch (e) {
        // Ignore errors when popup is closed
    }
    updateBadgeForStatus(payload);
}

function handleBuildAndDownloadMessage(msg: any, sendResponse: (res: any) => void) {
    (async () => {
        try {
            const result = await buildAndDownload(msg.data || {});
            sendResponse(result);
        } catch (err: any) {
            sendResponse({ error: err?.message || String(err) });
        }
    })();
}

function handleStartExportMessage(msg: any, sendResponse: (res: any) => void) {
    const data = msg.data || {};
    const tabId = data.tabId;
    if (typeof tabId !== 'number') {
        sendResponse({ error: 'Missing tabId for export request' });
        return;
    }
    if (activeExports.has(tabId)) {
        sendResponse({ error: 'An export is already running for this tab' });
        return;
    }

    const scrapeOptions = data.scrapeOptions || {};
    const buildOptions = data.buildOptions || {};

    (async () => {
        let startedAt;
        try {
            await ensureContentScript(tabId);
            const ctx = await sendMessageToTab(tabId, { type: 'CHECK_CHAT_CONTEXT' });
            if (!ctx?.ok) {
                const message = ctx?.reason || 'Open a chat conversation before exporting.';
                sendResponse({ error: message });
                return;
            }

            startedAt = Date.now();
            updateActiveExport(tabId, { startedAt, phase: 'starting', lastStatus: undefined });
            broadcastStatus({ tabId, phase: 'starting', startedAt });

            broadcastStatus({ tabId, phase: 'scrape:start' });
            const scrapeRes = await requestScrape(tabId, scrapeOptions);
            const totalMessages = Array.isArray(scrapeRes.messages) ? scrapeRes.messages.length : 0;
            broadcastStatus({ tabId, phase: 'scrape:complete', messages: totalMessages });

            if (totalMessages === 0) {
                const message = 'No messages found for the selected range.';
                broadcastStatus({ tabId, phase: 'empty', message });
                sendResponse({ error: message, code: 'EMPTY_RESULTS' });
                return;
            }

            const buildRes = await buildAndDownload({
                messages: scrapeRes.messages || [],
                meta: scrapeRes.meta || {},
                format: buildOptions.format || 'json',
                saveAs: buildOptions.saveAs !== false,
                embedAvatars: Boolean(buildOptions.embedAvatars)
            });

            broadcastStatus({ tabId, phase: 'complete', filename: buildRes.filename });
            sendResponse({ ok: true, filename: buildRes.filename, downloadId: buildRes.id });
        } catch (err: any) {
            const message = err?.message || String(err);
            broadcastStatus({ tabId, phase: 'error', error: message });
            sendResponse({ error: message });
        } finally {
            if (startedAt) {
                activeExports.delete(tabId);
            }
        }
    })();
}

resetBadge();

runtime.onMessage.addListener((msg, sender, sendResponse) => {
    if (!msg || !msg.type) return;

    if (msg.type === 'PING_SW') {
        sendResponse({ ok: true, now: Date.now() });
        return;
    }

    if (msg.type === 'BUILD_AND_DOWNLOAD') {
        handleBuildAndDownloadMessage(msg, sendResponse);
        return true;
    }

    if (msg.type === 'START_EXPORT') {
        handleStartExportMessage(msg, sendResponse);
        return true;
    }

    if (msg.type === 'SCRAPE_PROGRESS') {
        updateBadgeForProgress(msg.payload || msg);
        return;
    }

    if (msg.type === 'GET_EXPORT_STATUS') {
        const tabId = typeof msg.tabId === 'number' ? msg.tabId : sender?.tab?.id;
        if (typeof tabId !== 'number') {
            sendResponse({ active: false });
            return;
        }
        const info = activeExports.get(tabId) || null;
        sendResponse({ active: Boolean(info), info });
        return;
    }
});

}); // End of defineBackground
