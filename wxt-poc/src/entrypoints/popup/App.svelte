<script lang="ts" context="module">
  // Firefox polyfill global (typed loosely to avoid pulling extra deps)
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  declare const browser: any;
</script>

<script lang="ts">
  import './popup.css';
  import { onDestroy, onMount } from 'svelte';
  import {
    clearLastError,
    DEFAULT_OPTIONS,
    loadLastError,
    loadOptions,
    persistErrorMessage,
    saveOptions,
    validateRange,
    type OptionFormat,
    type Options,
    type Theme,
  } from '../../utils/options';
  import { formatElapsedSuffix, isoToLocalInput, localInputToISO } from '../../utils/time';
  import { runtimeSend } from '../../utils/messaging';
  import type {
    GetExportStatusRequest,
    GetExportStatusResponse,
    PingSWRequest,
    StartExportRequest,
    StartExportResponse,
  } from '../../types/messaging';
  import HeaderSection from './components/HeaderSection.svelte';
  import QuickRangeSection from './components/QuickRangeSection.svelte';
  import OptionsSection from './components/OptionsSection.svelte';
  import AdvancedSection from './components/AdvancedSection.svelte';
  import ActionSection from './components/ActionSection.svelte';
  import { t, setLanguage, getLanguage } from '../../i18n/i18n';

  const runtime = typeof browser !== 'undefined' ? browser.runtime : chrome.runtime;
  const tabs = typeof browser !== 'undefined' ? browser.tabs : chrome.tabs;
  const storage = typeof browser !== 'undefined' ? browser.storage : chrome.storage;

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

  const DAY_MS = 24 * 60 * 60 * 1000;
  const languageOptions = [
    { value: 'en', label: 'English' },
    { value: 'zh-CN', label: '简体中文' },
    { value: 'pt-BR', label: 'Português (Brasil)' },
    { value: 'nl', label: 'Nederlands' },
    { value: 'fr', label: 'Français' },
    { value: 'de', label: 'Deutsch' },
    { value: 'it', label: 'Italiano' },
    { value: 'ja', label: '日本語' },
    { value: 'ko', label: '한국어' },
    { value: 'ru', label: 'Русский' },
    { value: 'es', label: 'Español' },
    { value: 'tr', label: 'Türkçe' },
    { value: 'ar', label: 'العربية' },
    { value: 'he', label: 'עברית' },
  ];

  let options: Options = { ...DEFAULT_OPTIONS };
  const currentLang = () => options.lang || 'en';
  const runLabel = () => t('actions.export', {}, currentLang());
  const busyExportLabel = () => t('actions.busy.exporting', {}, currentLang());
  const busyBuildLabel = () => t('actions.busy.building', {}, currentLang());
  const emptyLabel = () => t('status.empty', {}, currentLang());

  let quickRanges = [
    { key: 'none', label: t('quick.none', {}, currentLang()), icon: '∞' },
    { key: '1d', label: t('quick.1d', {}, currentLang()), icon: '24h' },
    { key: '7d', label: t('quick.7d', {}, currentLang()), icon: '7d' },
    { key: '30d', label: t('quick.30d', {}, currentLang()), icon: '30d' },
  ];
  let bannerMessage: string | null = null;
  let advancedOpen = false;
  let quickActive = 'none';
  let statusText = '';
  let statusBaseText = '';
  let alive = true;
  let busy = false;
  let busyLabel = runLabel();
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

  const applyLanguage = async (lang: string) => {
    await setLanguage(lang || 'en');
    options = { ...options, lang: getLanguage() };
    const langNow = currentLang();
    quickRanges = [
      { key: 'none', label: t('quick.none', {}, langNow), icon: '∞' },
      { key: '1d', label: t('quick.1d', {}, langNow), icon: '24h' },
      { key: '7d', label: t('quick.7d', {}, langNow), icon: '7d' },
      { key: '30d', label: t('quick.30d', {}, langNow), icon: '30d' },
    ];
    if (!busy) busyLabel = runLabel();
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

  const setBusy = (state: boolean, labelText?: string) => {
    if (!alive) return;
    busy = state;
    busyLabel = state ? (labelText ?? busyExportLabel()) : runLabel();
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
    if (persist) void persistErrorMessage(storage, message);
  };

  const hideErrorBanner = (clearStorage = false) => {
    if (!alive) return;
    bannerMessage = null;
    if (clearStorage) void clearLastError(storage);
  };

  const loadPersistedError = () => loadLastError(storage);

  const loadStoredOptions = () => loadOptions(storage, DEFAULT_OPTIONS);

  const persistOptions = async () => {
    if (!alive) return;
    options = await saveOptions(storage, options, DEFAULT_OPTIONS);
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
    if (key === 'lang') {
      void applyLanguage(String(value));
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
    try {
      return validateRange(options);
    } catch (e: unknown) {
      const raw = e instanceof Error ? e.message : '';
      const msg = raw.includes('Start date must be before end date.')
        ? t('errors.startAfterEnd')
        : t('errors.invalidRange');
      showErrorBanner(msg);
      throw new Error(msg);
    }
  };

  const getActiveTeamsTab = async () => {
    const [tab] = await tabs.query({ active: true, currentWindow: true });
    if (!tab || !isTeamsUrl(tab.url)) throw new Error(t('errors.needsTeams'));
    return tab;
  };

  const pingSW = async (timeoutMs = 4000) =>
    Promise.race([
      runtimeSend<PingSWRequest>(runtime, { type: 'PING_SW' }),
      new Promise((_, rej) => setTimeout(() => rej(new Error(t('errors.ping'))), timeoutMs)),
    ]);

  const handleExportStatus = (msg: ExportStatusMsg) => {
    const langNow = currentLang();
    const tabId = msg?.tabId;
    if (typeof tabId === 'number') {
      if (currentTabId && tabId !== currentTabId) return;
      if (!currentTabId) currentTabId = tabId;
    }
    const phase = msg?.phase;
    if (phase === 'starting') {
      hideErrorBanner(true);
      const startedAt = normalizeStart(msg.startedAt);
      setBusy(true, busyExportLabel());
      setStatus(t('status.preparing', {}, langNow), { startElapsedAt: startedAt });
    } else if (phase === 'scrape:start') {
      setBusy(true, busyExportLabel());
      setStatus(t('status.running', {}, langNow));
    } else if (phase === 'scrape:complete') {
      setBusy(true, busyBuildLabel());
      setStatus(t('status.building', {}, langNow));
    } else if (phase === 'empty') {
      const message = msg.message || emptyLabel();
      setBusy(false);
      setStatus(message, { stopElapsed: true });
      showErrorBanner(message, false);
      void clearLastError(storage);
    } else if (phase === 'complete') {
      setBusy(false);
      if (msg.filename) {
        setStatus(t('status.complete', {}, langNow), { stopElapsed: true });
      } else {
        setStatus(t('status.complete', {}, langNow), { stopElapsed: true });
      }
      hideErrorBanner(true);
    } else if (phase === 'error') {
      setBusy(false);
      setStatus(msg.error || t('status.error', {}, langNow), { stopElapsed: true });
      showErrorBanner(msg.error || t('status.error', {}, langNow));
    }
  };

  const onRuntimeMessage = (msg: any) => {
    if (msg?.type === 'SCRAPE_PROGRESS') {
      const langNow = currentLang();
      const p = msg.payload || {};
      if (p.phase === 'scroll') {
        const seen = p.seen ?? p.aggregated ?? p.messagesVisible ?? 0;
        setStatus(t('status.scroll', { pass: p.passes ?? 0, seen }, langNow));
      } else if (p.phase === 'extract') {
        setStatus(t('status.extract', { count: p.messagesExtracted ?? 0 }, langNow));
      }
    } else if (msg?.type === 'EXPORT_STATUS') {
      handleExportStatus(msg);
    }
  };

  const startExport = async () => {
    if (busy || !alive) return;
    try {
      hideErrorBanner(true);
      setBusy(true, busyExportLabel());
      setStatus(t('status.preparing', {}, currentLang()));
      const tab = await getActiveTeamsTab();
      if (!alive) return;
      currentTabId = tab.id ?? null;
      await pingSW();
      const range = getValidatedRangeISO();
      const format = options.format;
      const { includeReplies, includeReactions, includeSystem, embedAvatars, showHud } = options;
      setStatus(t('status.running', {}, currentLang()));
      const response = await runtimeSend<StartExportRequest>(runtime, {
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
        const message = response.error || emptyLabel();
        setStatus(message, { stopElapsed: true });
        showErrorBanner(message, false);
        await clearLastError(storage);
        return;
      }
      if (!response || response.error) {
        throw new Error(response?.error || t('status.error', {}, currentLang()));
      }
      const langNow = currentLang();
      setStatus(response.filename ? `${t('status.complete', {}, langNow)} (${response.filename})` : t('status.complete', {}, langNow));
      hideErrorBanner(true);
    } catch (e: any) {
      const raw = e?.message || '';
      const msg =
        raw.includes('Teams web app') || raw.includes('Teams tab')
          ? t('errors.needsTeams', {}, currentLang())
          : raw.includes('background')
            ? t('errors.ping', {}, currentLang())
            : raw || t('status.error', {}, currentLang());
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
      await applyLanguage(options.lang || 'en');
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
      const status = await runtimeSend<GetExportStatusRequest>(runtime, {
        type: 'GET_EXPORT_STATUS',
        tabId: currentTabId,
      });
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
            setBusy(true, busyExportLabel());
            setStatus(t('status.running'));
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
  <HeaderSection lang={options.lang || 'en'} />

{#if bannerMessage}
  <div id="banner" class="alert error show" role="alert" aria-live="assertive">
    <span class="alert-title">{t('banner.error', {}, options.lang || 'en')}</span>
    <span class="alert-message">{bannerMessage}</span>
  </div>
{/if}

  <QuickRangeSection
    lang={options.lang || 'en'}
    startAt={options.startAt}
    endAt={options.endAt}
    activeRange={quickActive}
    ranges={quickRanges as any}
    on:changeStart={(e) => updateOption('startAt', e.detail)}
    on:changeEnd={(e) => updateOption('endAt', e.detail)}
    on:quickSelect={(e) => handleQuickRange(e.detail)}
  />

  <OptionsSection
    lang={options.lang || 'en'}
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
    lang={options.lang || 'en'}
    open={advancedOpen}
    showHud={options.showHud}
    theme={options.theme}
    languages={languageOptions}
    on:toggleOpen={(e) => (advancedOpen = e.detail)}
    on:showHudChange={(e) => updateOption('showHud', e.detail)}
    on:themeChange={(e) => updateOption('theme', e.detail)}
    on:langChange={(e) => updateOption('lang', e.detail)}
  />

  <ActionSection
    lang={options.lang || 'en'}
    busy={busy}
    busyLabel={busyLabel}
    defaultRunLabel={runLabel()}
    statusText={statusText}
    on:run={startExport}
  />
</div>
