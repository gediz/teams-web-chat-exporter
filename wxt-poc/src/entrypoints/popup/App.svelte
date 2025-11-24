<script lang="ts" context="module">
  // Firefox polyfill global (typed loosely to avoid pulling extra deps)
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  declare const browser: any;
</script>

<script lang="ts">
  import './popup.css';
  import { onDestroy, onMount } from 'svelte';
  import HeaderSection from './components/HeaderSection.svelte';
  import QuickRangeSection from './components/QuickRangeSection.svelte';
  import OptionsSection from './components/OptionsSection.svelte';
  import AdvancedSection from './components/AdvancedSection.svelte';

  const runtime = typeof browser !== 'undefined' ? browser.runtime : chrome.runtime;
  const tabs = typeof browser !== 'undefined' ? browser.tabs : chrome.tabs;
  const storage = typeof browser !== 'undefined' ? browser.storage : chrome.storage;

  type OptionFormat = 'json' | 'csv' | 'html';
  type Theme = 'light' | 'dark';

  type Options = {
    startAt: string;
    startAtISO: string;
    endAt: string;
    endAtISO: string;
    format: OptionFormat;
    includeReplies: boolean;
    includeReactions: boolean;
    includeSystem: boolean;
    embedAvatars: boolean;
    showHud: boolean;
    theme: Theme;
  };

  type ExportStatusMsg = {
    tabId?: number;
    phase?: string;
    startedAt?: number | string;
    messages?: number;
    filename?: string;
    message?: string;
    error?: string;
  };

  type ExportStatusResponse = {
    active: boolean;
    info?: { startedAt?: number | string; lastStatus?: ExportStatusMsg };
  };

  const STORAGE_KEY = 'teamsExporterOptions';
  const ERROR_STORAGE_KEY = 'teamsExporterLastError';
  const DEFAULT_RUN_LABEL = 'Export current chat';
  const BUSY_LABEL_EXPORTING = 'Exporting…';
  const BUSY_LABEL_BUILDING = 'Building…';
  const EMPTY_RESULT_MESSAGE = 'No messages found for the selected range.';
  const DAY_MS = 24 * 60 * 60 * 1000;

  const DEFAULT_OPTIONS: Options = {
    startAt: '',
    startAtISO: '',
    endAt: '',
    endAtISO: '',
    format: 'json',
    includeReplies: true,
    includeReactions: true,
    includeSystem: false,
    embedAvatars: false,
    showHud: true,
    theme: 'light',
  };

  const quickRanges = [
    { key: 'none', label: 'No limit', icon: '∞' },
    { key: '1d', label: 'Last 24h', icon: '24h' },
    { key: '7d', label: 'Last 7d', icon: '7d' },
    { key: '30d', label: 'Last 30d', icon: '30d' },
  ] as const;

  let options: Options = { ...DEFAULT_OPTIONS };
  let bannerMessage: string | null = null;
  let advancedOpen = false;
  let quickActive = 'none';
  let statusText = '';
  let statusBaseText = '';
  let alive = true;
  let busy = false;
  let busyLabel = DEFAULT_RUN_LABEL;
  let currentTabId: number | null = null;
  let startedAtMs: number | null = null;
  let elapsedTimer: ReturnType<typeof setInterval> | null = null;

  const isTeamsUrl = (u?: string | null) =>
    /^https:\/\/(.*\.)?(teams\.microsoft\.com|cloud\.microsoft)\//.test(u || '');

  const applyTheme = (theme: Theme) => {
    const next = theme === 'dark' ? 'dark' : 'light';
    document.body.dataset.theme = next;
    options = { ...options, theme: next };
  };

  const localInputToISO = (localValue: string) => {
    if (!localValue) return '';
    let normalized = localValue.trim();
    if (!normalized) return '';
    normalized = normalized.replace(/\//g, '-').replace(/\s+/g, ' ');
    if (/^\d{4}-\d{2}-\d{2}$/.test(normalized)) {
      normalized += ' 00:00';
    }
    if (normalized.includes(' ')) {
      normalized = normalized.replace(' ', 'T');
    }
    const date = new Date(normalized);
    if (Number.isNaN(date.getTime())) return '';
    return date.toISOString();
  };

  const isoToLocalInput = (isoValue: string) => {
    if (!isoValue) return '';
    const date = new Date(isoValue);
    if (Number.isNaN(date.getTime())) return '';
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    const hours = String(date.getHours()).padStart(2, '0');
    const minutes = String(date.getMinutes()).padStart(2, '0');
    return `${year}-${month}-${day} ${hours}:${minutes}`;
  };

  const normalizeStart = (value: unknown) => {
    if (typeof value === 'number' && !Number.isNaN(value)) return value;
    if (typeof value === 'string') {
      const parsed = Number(value);
      if (!Number.isNaN(parsed)) return parsed;
      const date = Date.parse(value);
      if (!Number.isNaN(date)) return date;
    }
    return null;
  };

  const formatElapsed = (ms: number) => {
    const totalSeconds = Math.max(0, Math.floor(ms / 1000));
    const hours = Math.floor(totalSeconds / 3600);
    const minutes = Math.floor((totalSeconds % 3600) / 60);
    const seconds = totalSeconds % 60;
    if (hours > 0) {
      return `${hours}:${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;
    }
    return `${minutes}:${String(seconds).padStart(2, '0')}`;
  };

  const formatElapsedSuffix = (ms: number) => ` — Elapsed: ${formatElapsed(ms)}`;

  const updateQuickRangeActive = () => {
    const startISO = localInputToISO(options.startAt) || null;
    const endISO = localInputToISO(options.endAt) || null;
    const now = Date.now();
    const tolerance = 5 * 60 * 1000;
    let active = 'none';
    if (startISO || endISO) {
      const ranges = [
        { key: '1d', ms: DAY_MS },
        { key: '7d', ms: 7 * DAY_MS },
        { key: '30d', ms: 30 * DAY_MS },
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
    quickActive = active;
  };

  const setBusy = (state: boolean, labelText = BUSY_LABEL_EXPORTING) => {
    if (!alive) return;
    busy = state;
    busyLabel = state ? labelText : DEFAULT_RUN_LABEL;
  };

  const updateStatusText = () => {
    if (!statusBaseText) return;
    let text = statusBaseText;
    if (startedAtMs) {
      text += formatElapsedSuffix(Date.now() - startedAtMs);
    }
    statusText = text;
  };

  const ensureElapsedTimer = () => {
    if (elapsedTimer) return;
    elapsedTimer = setInterval(() => {
      if (!startedAtMs) {
        clearElapsedTimer();
        return;
      }
      updateStatusText();
    }, 1000);
    updateStatusText();
  };

  const clearElapsedTimer = () => {
    if (elapsedTimer) {
      clearInterval(elapsedTimer);
      elapsedTimer = null;
    }
  };

  const setStatus = (text: string, opts: { startElapsedAt?: number | null; stopElapsed?: boolean } = {}) => {
    if (!alive) return;
    statusBaseText = text;
    if (typeof opts.startElapsedAt === 'number' && !Number.isNaN(opts.startElapsedAt)) {
      startedAtMs = opts.startElapsedAt;
      ensureElapsedTimer();
      return;
    }
    if (opts.stopElapsed) {
      statusText = startedAtMs ? `${statusBaseText}${formatElapsedSuffix(Date.now() - startedAtMs)}` : statusBaseText;
      startedAtMs = null;
      clearElapsedTimer();
      return;
    }
    updateStatusText();
  };

  const showErrorBanner = (message: string, persist = true) => {
    if (!alive) return;
    bannerMessage = message;
    if (persist) void persistError(message);
  };

  const hideErrorBanner = (clearStorage = false) => {
    if (!alive) return;
    bannerMessage = null;
    if (clearStorage) void clearPersistedError();
  };

  const persistError = async (message: string) => {
    try {
      await storage.local.set({ [ERROR_STORAGE_KEY]: { message, timestamp: Date.now() } });
    } catch {
      /* ignore */
    }
  };

  const clearPersistedError = async () => {
    try {
      await storage.local.remove(ERROR_STORAGE_KEY);
    } catch {
      /* ignore */
    }
  };

  const loadPersistedError = async () => {
    try {
      const res = await storage.local.get(ERROR_STORAGE_KEY);
      return res?.[ERROR_STORAGE_KEY] || null;
    } catch {
      return null;
    }
  };

  const loadStoredOptions = async (): Promise<Options> => {
    try {
      const stored = await storage.local.get(STORAGE_KEY);
      const loaded = (stored?.[STORAGE_KEY] || {}) as Partial<Options>;
      const merged: Options = { ...DEFAULT_OPTIONS, ...loaded };
      merged.startAt = merged.startAt || isoToLocalInput(merged.startAtISO);
      merged.endAt = merged.endAt || isoToLocalInput(merged.endAtISO);
      return merged;
    } catch {
      return { ...DEFAULT_OPTIONS };
    }
  };

  const persistOptions = async () => {
    const startISO = localInputToISO(options.startAt);
    const endISO = localInputToISO(options.endAt);
    const payload: Options = { ...options, startAtISO: startISO || '', endAtISO: endISO || '' };
    if (!alive) return;
    options = payload;
    try {
      await storage.local.set({ [STORAGE_KEY]: payload });
    } catch {
      /* ignore */
    }
  };

  const updateOption = <K extends keyof Options>(key: K, value: Options[K]) => {
    if (!alive) return;
    options = { ...options, [key]: value };
    if (key === 'startAt' || key === 'endAt') {
      updateQuickRangeActive();
    }
    if (key === 'theme') {
      applyTheme(value as Theme);
    }
    void persistOptions();
  };

  const handleQuickRange = (range: string) => {
    if (!alive) return;
    const normalized = range || 'none';
    if (normalized === 'none') {
      options = { ...options, startAt: '', endAt: '' };
      updateQuickRangeActive();
      void persistOptions();
      return;
    }
    const now = new Date();
    let offsetMs = 0;
    if (normalized.endsWith('d')) {
      const days = Number(normalized.replace('d', ''));
      if (!Number.isNaN(days)) offsetMs = days * DAY_MS;
    }
    if (offsetMs > 0) {
      const startDate = new Date(now.getTime() - offsetMs);
      options = {
        ...options,
        startAt: isoToLocalInput(startDate.toISOString()),
        endAt: isoToLocalInput(now.toISOString()),
      };
    } else {
      options = { ...options, startAt: '', endAt: '' };
    }
    updateQuickRangeActive();
    void persistOptions();
  };

  const getValidatedRangeISO = () => {
    const rawStart = (options.startAt || '').trim();
    const rawEnd = (options.endAt || '').trim();
    const startISO = rawStart ? localInputToISO(rawStart) : null;
    if (rawStart && !startISO) {
      const message = 'Enter a valid start date/time.';
      showErrorBanner(message);
      throw new Error(message);
    }
    const endISO = rawEnd ? localInputToISO(rawEnd) : null;
    if (rawEnd && !endISO) {
      const message = 'Enter a valid end date/time.';
      showErrorBanner(message);
      throw new Error(message);
    }
    if (startISO && endISO) {
      const startMs = Date.parse(startISO);
      const endMs = Date.parse(endISO);
      if (!Number.isNaN(startMs) && !Number.isNaN(endMs) && startMs > endMs) {
        const message = 'Start date must be before end date.';
        showErrorBanner(message);
        throw new Error(message);
      }
    }
    return { startISO, endISO };
  };

  const getActiveTeamsTab = async () => {
    const [tab] = await tabs.query({ active: true, currentWindow: true });
    if (!tab || !isTeamsUrl(tab.url)) throw new Error('Open the Teams web app tab first.');
    return tab;
  };

  const pingSW = async (timeoutMs = 4000) =>
    Promise.race([
      runtime.sendMessage({ type: 'PING_SW' }),
      new Promise((_, rej) => setTimeout(() => rej(new Error('No response from background (PING_SW timeout)')), timeoutMs)),
    ]);

  const handleExportStatus = (msg: ExportStatusMsg) => {
    const tabId = msg?.tabId;
    if (typeof tabId === 'number') {
      if (currentTabId && tabId !== currentTabId) return;
      if (!currentTabId) currentTabId = tabId;
    }
    const phase = msg?.phase;
    if (phase === 'starting') {
      hideErrorBanner(true);
      const startedAt = normalizeStart(msg.startedAt);
      setBusy(true, BUSY_LABEL_EXPORTING);
      setStatus('Starting export…', { startElapsedAt: startedAt });
    } else if (phase === 'scrape:start') {
      setBusy(true, BUSY_LABEL_EXPORTING);
      setStatus('Running auto-scroll + scrape…');
    } else if (phase === 'scrape:complete') {
      setBusy(true, BUSY_LABEL_BUILDING);
      setStatus(`Collected ${msg.messages ?? 0} messages. Building…`);
    } else if (phase === 'empty') {
      const message = msg.message || EMPTY_RESULT_MESSAGE;
      setBusy(false);
      setStatus(message, { stopElapsed: true });
      showErrorBanner(message, false);
      void clearPersistedError();
    } else if (phase === 'complete') {
      setBusy(false);
      if (msg.filename) {
        setStatus(`Exported ${msg.filename}`, { stopElapsed: true });
      } else {
        setStatus('Export complete.', { stopElapsed: true });
      }
      hideErrorBanner(true);
    } else if (phase === 'error') {
      setBusy(false);
      setStatus(msg.error || 'Export failed.', { stopElapsed: true });
      showErrorBanner(msg.error || 'Export failed.');
    }
  };

  const onRuntimeMessage = (msg: any) => {
    if (msg?.type === 'SCRAPE_PROGRESS') {
      const p = msg.payload || {};
      if (p.phase === 'scroll') {
        const seen = p.seen ?? p.aggregated ?? p.messagesVisible ?? 0;
        setStatus(`Scrolling… pass ${p.passes} • seen ${seen}`);
      } else if (p.phase === 'extract') {
        setStatus(`Extracting… found ${p.messagesExtracted} messages`);
      }
    } else if (msg?.type === 'EXPORT_STATUS') {
      handleExportStatus(msg);
    }
  };

  const startExport = async () => {
    if (busy || !alive) return;
    try {
      hideErrorBanner(true);
      setBusy(true, BUSY_LABEL_EXPORTING);
      setStatus('Preparing…');
      const tab = await getActiveTeamsTab();
      if (!alive) return;
      currentTabId = tab.id ?? null;
      await pingSW();
      const range = getValidatedRangeISO();
      const format = options.format;
      const { includeReplies, includeReactions, includeSystem, embedAvatars, showHud } = options;
      setStatus('Export running… you can close this popup.');
      const response = await runtime.sendMessage({
        type: 'START_EXPORT',
        data: {
          tabId: tab.id,
          scrapeOptions: {
            startAt: range.startISO,
            endAt: range.endISO,
            includeReplies,
            includeReactions,
            includeSystem,
            showHud,
          },
          buildOptions: { format, saveAs: true, embedAvatars },
        },
      });
      if (response?.code === 'EMPTY_RESULTS') {
        const message = response.error || EMPTY_RESULT_MESSAGE;
        setStatus(message, { stopElapsed: true });
        showErrorBanner(message, false);
        await clearPersistedError();
        return;
      }
      if (!response || response.error) {
        throw new Error(response?.error || 'Export failed.');
      }
      setStatus(`Exported ${response.filename}`);
      hideErrorBanner(true);
    } catch (e: any) {
      const msg = e?.message || 'Export failed.';
      setStatus(msg);
      showErrorBanner(msg);
    } finally {
      setBusy(false);
    }
  };

  onMount(() => {
    const init = async () => {
      setBusy(false);
      const loaded = await loadStoredOptions();
      if (!alive) return;
      options = loaded;
      applyTheme(options.theme || 'light');
      updateQuickRangeActive();
      advancedOpen = false;
      const persistedError = await loadPersistedError();
      if (!alive) return;
      if (persistedError?.message) {
        showErrorBanner(persistedError.message, false);
        if (!statusText) {
          setStatus(persistedError.message);
        }
      }
      try {
      const tab = await getActiveTeamsTab();
      currentTabId = tab.id ?? null;
      const status = (await runtime.sendMessage({
        type: 'GET_EXPORT_STATUS',
        tabId: currentTabId,
      })) as ExportStatusResponse;
        if (!alive) return;
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
            setStatus('Export running…');
          }
        }
      } catch {
        /* user not on Teams tab */
      }
    };
    void init();
    runtime.onMessage.addListener(onRuntimeMessage);
  });

  onDestroy(() => {
    alive = false;
    runtime.onMessage.removeListener(onRuntimeMessage);
    clearElapsedTimer();
  });
</script>

<div class="popup">
  <HeaderSection theme={options.theme} on:toggleTheme={(e) => updateOption('theme', e.detail)} />

  {#if bannerMessage}
    <div id="banner" class="alert error show" role="alert" aria-live="assertive">
      <span class="alert-title">Error</span>
      <span class="alert-message">{bannerMessage}</span>
    </div>
  {/if}

  <QuickRangeSection
    startAt={options.startAt}
    endAt={options.endAt}
    activeRange={quickActive}
    ranges={quickRanges}
    on:changeStart={(e) => updateOption('startAt', e.detail)}
    on:changeEnd={(e) => updateOption('endAt', e.detail)}
    on:quickSelect={(e) => handleQuickRange(e.detail)}
  />

  <OptionsSection
    format={options.format}
    includeReplies={options.includeReplies}
    includeReactions={options.includeReactions}
    includeSystem={options.includeSystem}
    embedAvatars={options.embedAvatars}
    on:formatChange={(e) => updateOption('format', e.detail)}
    on:includeRepliesChange={(e) => updateOption('includeReplies', e.detail)}
    on:includeReactionsChange={(e) => updateOption('includeReactions', e.detail)}
    on:includeSystemChange={(e) => updateOption('includeSystem', e.detail)}
    on:embedAvatarsChange={(e) => updateOption('embedAvatars', e.detail)}
  />

  <AdvancedSection
    open={advancedOpen}
    showHud={options.showHud}
    on:toggleOpen={(e) => (advancedOpen = e.detail)}
    on:showHudChange={(e) => updateOption('showHud', e.detail)}
  />

  <button id="run" class:busy={busy} on:click={startExport} disabled={busy}>
    <span class="spinner" aria-hidden="true"></span>
    <span class="button-icon" aria-hidden="true">⤓</span>
    <span class="label">{busy ? busyLabel : DEFAULT_RUN_LABEL}</span>
  </button>

  <div id="status" aria-live="polite">{statusText}</div>

  <p class="footer-note">
    Keep this popup open to monitor progress, or close it and watch the badge.
  </p>
</div>
