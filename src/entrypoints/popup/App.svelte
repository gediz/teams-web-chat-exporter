<script lang="ts" module>
  // Firefox polyfill global (typed loosely to avoid pulling extra deps)
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  declare const browser: any;
</script>

<script lang="ts">
  import "./popup.css";
  import { onDestroy, onMount } from "svelte";
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
  } from "../../utils/options";
  import {
    formatElapsed,
    isoToLocalInput,
    localInputToISO,
  } from "../../utils/time";
  import { runtimeSend } from "../../utils/messaging";
  import type {
    GetExportStatusRequest,
    GetExportStatusResponse,
    PingSWRequest,
    StartExportRequest,
    StartExportResponse,
    StartExportZipRequest,
    StartExportZipResponse,
  } from "../../types/messaging";
  import ExportButton from "./components/ExportButton.svelte";
  import FormatSection from "./components/FormatSection.svelte";
  import DateRangeSection, {
    type QuickRange,
  } from "./components/DateRangeSection.svelte";
  import IncludeSection from "./components/IncludeSection.svelte";
  import StatusBar from "./components/StatusBar.svelte";
  import HeaderActions from "./components/HeaderActions.svelte";
  import SettingsPage from "./components/SettingsPage.svelte";
  import { t, setLanguage, getLanguage } from "../../i18n/i18n";

  const runtime =
    typeof browser !== "undefined" ? browser.runtime : chrome.runtime;
  const tabs = typeof browser !== "undefined" ? browser.tabs : chrome.tabs;
  const storage =
    typeof browser !== "undefined" ? browser.storage : chrome.storage;

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
    { value: "en", code: "EN", native: "English", label: "English" },
    { value: "ar", code: "AR", native: "العربية", label: "Arabic (العربية)" },
    { value: "az", code: "AZ", native: "Azərbaycan", label: "Azerbaijani (Azərbaycan)" },
    { value: "bn", code: "BN", native: "বাংলা", label: "Bengali (বাংলা)" },
    { value: "bg", code: "BG", native: "Български", label: "Bulgarian (Български)" },
    { value: "zh-CN", code: "ZH", native: "简体中文", label: "Chinese Simplified (简体中文)" },
    { value: "zh-TW", code: "TW", native: "繁體中文", label: "Chinese Traditional (繁體中文)" },
    { value: "cs", code: "CS", native: "Čeština", label: "Czech (Čeština)" },
    { value: "nl", code: "NL", native: "Nederlands", label: "Dutch (Nederlands)" },
    { value: "fr", code: "FR", native: "Français", label: "French (Français)" },
    { value: "de", code: "DE", native: "Deutsch", label: "German (Deutsch)" },
    { value: "he", code: "HE", native: "עברית", label: "Hebrew (עברית)" },
    { value: "hi", code: "HI", native: "हिन्दी", label: "Hindi (हिन्दी)" },
    { value: "hu", code: "HU", native: "Magyar", label: "Hungarian (Magyar)" },
    { value: "it", code: "IT", native: "Italiano", label: "Italian (Italiano)" },
    { value: "ja", code: "JA", native: "日本語", label: "Japanese (日本語)" },
    { value: "ko", code: "KO", native: "한국어", label: "Korean (한국어)" },
    { value: "ms", code: "MS", native: "Melayu", label: "Malay (Bahasa Melayu)" },
    { value: "pl", code: "PL", native: "Polski", label: "Polish (Polski)" },
    { value: "pt-BR", code: "PT", native: "Português", label: "Portuguese (Português)" },
    { value: "ru", code: "RU", native: "Русский", label: "Russian (Русский)" },
    { value: "es", code: "ES", native: "Español", label: "Spanish (Español)" },
    { value: "th", code: "TH", native: "ไทย", label: "Thai (ไทย)" },
    { value: "tr", code: "TR", native: "Türkçe", label: "Turkish (Türkçe)" },
  ];

  let showSettings = false;

  let options: Options = { ...DEFAULT_OPTIONS };
  const currentLang = () => options.lang || "en";
  const runLabel = () => t("actions.export", {}, currentLang());
  const busyExportLabel = () => t("actions.busy.exporting", {}, currentLang());
  const busyBuildLabel = () => t("actions.busy.building", {}, currentLang());
  const emptyLabel = () => t("status.empty", {}, currentLang());

  let quickRanges: QuickRange[] = [
    { key: "none", label: t("quick.none", {}, currentLang()), icon: "∞" },
    { key: "1d", label: t("quick.1d", {}, currentLang()), icon: "24h" },
    { key: "7d", label: t("quick.7d", {}, currentLang()), icon: "7d" },
    { key: "30d", label: t("quick.30d", {}, currentLang()), icon: "30d" },
  ];
  let bannerMessage: string | null = null;
  let quickActive = "none";
  let statusText = t("status.ready", {}, currentLang());
  let statusBaseText = "";
  let statusCount = 0;
  let alive = true;
  let busy = false;
  let busyLabel = runLabel();
  let currentTabId: number | null = null;
  let startedAtMs: number | null = null;
  let elapsedTimer: ReturnType<typeof setInterval> | null = null;
  let exportSummary = "";

  const isTeamsUrl = (u?: string | null) =>
    /^https:\/\/(.*\.)?(teams\.microsoft\.com|teams\.microsoft\.us|cloud\.microsoft|teams\.live\.com)\//.test(
      u || "",
    );

  const formatElapsedSuffix = (ms: number) =>
    ` — ${t("status.elapsed", {}, currentLang())}: ${formatElapsed(ms)}`;

  const applyTheme = (theme: Theme) => {
    const next = theme === "dark" ? "dark" : "light";
    document.body.dataset.theme = next;
    options = { ...options, theme: next };
  };

  const applyLanguage = async (lang: string) => {
    await setLanguage(lang || "en");
    options = { ...options, lang: getLanguage() };
    const langNow = currentLang();
    quickRanges = [
      { key: "none", label: t("quick.none", {}, langNow), icon: "∞" },
      { key: "1d", label: t("quick.1d", {}, langNow), icon: "24h" },
      { key: "7d", label: t("quick.7d", {}, langNow), icon: "7d" },
      { key: "30d", label: t("quick.30d", {}, langNow), icon: "30d" },
    ];
    if (!busy) busyLabel = runLabel();
    // Update status text if it's still at "Ready"
    if (!busy && !startedAtMs) {
      statusText = t("status.ready", {}, langNow);
      statusBaseText = "";
    }
  };

  const normalizeStart = (value: unknown) => {
    if (typeof value === "number" && !Number.isNaN(value)) return value;
    if (typeof value === "string") {
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
    let active = "none";
    if (startISO || endISO) {
      const ranges = [
        { key: "1d", ms: DAY_MS },
        { key: "7d", ms: 7 * DAY_MS },
        { key: "30d", ms: 30 * DAY_MS },
      ];
      const endMs = endISO ? Date.parse(endISO) : now;
      const startMs = startISO ? Date.parse(startISO) : null;
      if (!Number.isNaN(endMs)) {
        for (const r of ranges) {
          const expectedStart = endMs - r.ms;
          const startOk =
            startMs != null && Math.abs(startMs - expectedStart) <= tolerance;
          const endOk =
            Math.abs(endMs - now) <= tolerance || (startISO && !endISO);
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

  const computeSummary = () => {
    const parts: string[] = [];
    const lang = currentLang();

    // Format
    const formatLabel = t(`format.${options.format}`, {}, lang);
    parts.push(formatLabel);

    // Date range
    if (quickActive && quickActive !== "none") {
      const rangeLabel = quickRanges.find((r) => r.key === quickActive)?.label;
      if (rangeLabel) parts.push(rangeLabel);
    }

    // Include options (only for non-txt formats)
    if (options.format !== "txt") {
      const includes: string[] = [];
      if (options.includeReplies) includes.push(t("summary.replies", {}, lang));
      if (options.includeReactions)
        includes.push(t("summary.reactions", {}, lang));
      if (options.includeSystem) includes.push(t("summary.system", {}, lang));
      if (options.embedAvatars) includes.push(t("summary.avatars", {}, lang));
      if (options.format === "html" && options.downloadImages) {
        includes.push(t("summary.images", {}, lang));
        includes.push(t("summary.zip", {}, lang));
      }
      if (includes.length > 0) parts.push(includes.join(", "));
    }

    return parts.join(" • ");
  };

  // Update summary when options change
  $: {
    options.format;
    options.includeReplies;
    options.includeReactions;
    options.includeSystem;
    options.embedAvatars;
    options.downloadImages;
    quickActive;
    exportSummary = computeSummary();
  }

  // Compute highlight mode for date range section
  $: highlightMode = (() => {
    const hasCustomDates = !!(options.startAt || options.endAt);

    // Manual mode: user has specified custom dates without a matching quick range
    if (hasCustomDates && quickActive === "none") {
      return "manual" as const;
    }
    // Quick range mode: a quick range is active
    if (quickActive && quickActive !== "none") {
      return "quick-range" as const;
    }
    // None mode: No limit is active (no dates AND activeRange is 'none')
    return "none" as const;
  })();

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

  const setStatus = (
    text: string,
    opts: {
      startElapsedAt?: number | null;
      stopElapsed?: boolean;
      count?: number;
    } = {},
  ) => {
    if (!alive) return;
    statusBaseText = text;
    if (typeof opts.count === "number") statusCount = opts.count;
    if (
      typeof opts.startElapsedAt === "number" &&
      !Number.isNaN(opts.startElapsedAt)
    ) {
      startedAtMs = opts.startElapsedAt;
      ensureElapsedTimer();
      return;
    }
    if (opts.stopElapsed) {
      statusText = startedAtMs
        ? `${statusBaseText}${formatElapsedSuffix(Date.now() - startedAtMs)}`
        : statusBaseText;
      startedAtMs = null;
      clearElapsedTimer();
      return;
    }
    updateStatusText();
  };

  const translateError = (message: string): string => {
    const lang = currentLang();
    // Map common error messages to translation keys
    if (message.includes("already running")) {
      return t("errors.alreadyRunning", {}, lang);
    }
    if (
      message.includes("Could not load file") ||
      message.includes("content.js")
    ) {
      return t("errors.contentScript", {}, lang);
    }
    if (message.includes("No messages found")) {
      return t("errors.noMessages", {}, lang);
    }
    if (message.includes("Missing tabId")) {
      return t("errors.missingTabId", {}, lang);
    }
    if (message.includes("Switch to the Chat app")) {
      return t("errors.switchToChat", {}, lang);
    }
    if (message.includes("Open a chat conversation")) {
      return t("errors.chatNotOpen", {}, lang);
    }
    // Return original message if no translation found
    return message;
  };

  const showErrorBanner = (message: string, persist = true) => {
    if (!alive) return;
    const translated = translateError(message);
    bannerMessage = translated;
    if (persist) void persistErrorMessage(storage, translated);
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
    await saveOptions(storage, options, DEFAULT_OPTIONS);
  };

  const updateOption = <K extends keyof Options>(key: K, value: Options[K]) => {
    if (!alive) return;
    options = { ...options, [key]: value };
    if (key === "startAt" || key === "endAt") {
      updateQuickRangeActive();
    }
    if (key === "theme") {
      applyTheme(value as Theme);
    }
    if (key === "lang") {
      void applyLanguage(String(value));
    }
    void persistOptions();
  };

  const handleQuickRange = (range: string) => {
    if (!alive) return;
    const normalized = range || "none";
    if (normalized === "none") {
      options = { ...options, startAt: "", endAt: "" };
      quickActive = "none";
      // Immediately save to storage to prevent restoration of old values
      void saveOptions(storage, options, DEFAULT_OPTIONS);
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
      options = {
        ...options,
        startAt: isoToLocalInput(startDate.toISOString()),
        endAt: isoToLocalInput(now.toISOString()),
      };
    } else {
      options = { ...options, startAt: "", endAt: "" };
    }
    updateQuickRangeActive();
    void persistOptions();
  };

  const getValidatedRangeISO = () => {
    try {
      return validateRange(options);
    } catch (e: unknown) {
      const raw = e instanceof Error ? e.message : "";
      const msg = raw.includes("Start date must be before end date.")
        ? t("errors.startAfterEnd")
        : t("errors.invalidRange");
      showErrorBanner(msg);
      throw new Error(msg);
    }
  };

  const getActiveTeamsTab = async () => {
    const [tab] = await tabs.query({ active: true, currentWindow: true });
    if (!tab || !isTeamsUrl(tab.url)) throw new Error(t("errors.needsTeams"));
    return tab;
  };

  const pingSW = async (timeoutMs = 4000) =>
    Promise.race([
      runtimeSend<PingSWRequest>(runtime, { type: "PING_SW" }),
      new Promise((_, rej) =>
        setTimeout(() => rej(new Error(t("errors.ping"))), timeoutMs),
      ),
    ]);

  const handleExportStatus = (msg: ExportStatusMsg) => {
    const langNow = currentLang();
    const tabId = msg?.tabId;
    if (typeof tabId === "number") {
      if (currentTabId && tabId !== currentTabId) return;
      if (!currentTabId) currentTabId = tabId;
    }
    const phase = msg?.phase;
    if (phase === "starting") {
      hideErrorBanner(true);
      const startedAt = normalizeStart(msg.startedAt);
      setBusy(true, busyExportLabel());
      setStatus(t("status.preparing", {}, langNow), {
        startElapsedAt: startedAt,
      });
    } else if (phase === "scrape:start") {
      setBusy(true, busyExportLabel());
      setStatus(t("status.running", {}, langNow));
    } else if (phase === "scrape:complete") {
      setBusy(true, busyBuildLabel());
      setStatus(t("status.building", {}, langNow));
    } else if (phase === "empty") {
      const message = msg.message || emptyLabel();
      setBusy(false);
      setStatus(message, { stopElapsed: true });
      showErrorBanner(message, false);
      void clearLastError(storage);
    } else if (phase === "complete") {
      setBusy(false);
      if (msg.filename) {
        setStatus(t("status.complete", {}, langNow), { stopElapsed: true });
      } else {
        setStatus(t("status.complete", {}, langNow), { stopElapsed: true });
      }
      hideErrorBanner(true);
    } else if (phase === "error") {
      setBusy(false);
      setStatus(msg.error || t("status.error", {}, langNow), {
        stopElapsed: true,
      });
      showErrorBanner(msg.error || t("status.error", {}, langNow));
    }
  };

  const onRuntimeMessage = (msg: any) => {
    if (msg?.type === "SCRAPE_PROGRESS") {
      const langNow = currentLang();
      const p = msg.payload || {};
      if (p.phase === "scroll") {
        const seen = p.seen ?? p.aggregated ?? p.messagesVisible ?? 0;
        setStatus(t("status.scroll", { pass: p.passes ?? 0, seen }, langNow), {
          count: seen,
        });
      } else if (p.phase === "extract") {
        setStatus(
          t("status.extract", { count: p.messagesExtracted ?? 0 }, langNow),
          { count: p.messagesExtracted ?? 0 },
        );
      }
    } else if (msg?.type === "EXPORT_STATUS") {
      handleExportStatus(msg);
    }
  };

  const startExport = async () => {
    if (busy || !alive) return;
    try {
      hideErrorBanner(true);
      setBusy(true, busyExportLabel());
      setStatus(t("status.preparing", {}, currentLang()));
      const tab = await getActiveTeamsTab();
      if (!alive) return;
      currentTabId = tab.id ?? null;
      await pingSW();
      const range = getValidatedRangeISO();
      const format = options.format;
      const {
        includeReplies,
        includeReactions,
        includeSystem,
        embedAvatars,
        downloadImages,
        showHud,
      } = options;
      setStatus(t("status.running", {}, currentLang()));
      const requestData = {
        tabId: tab.id,
        scrapeOptions: {
          startAt: range.startISO,
          endAt: range.endISO,
          includeReplies,
          includeReactions,
          includeSystem,
          showHud,
        },
        buildOptions: { format, saveAs: true, embedAvatars, downloadImages },
      };
      let response: StartExportResponse | StartExportZipResponse;
      if (format === "html" && downloadImages) {
        response = await runtimeSend<StartExportZipRequest>(runtime, {
          type: "START_EXPORT_ZIP",
          data: requestData,
        });
      } else {
        response = await runtimeSend<StartExportRequest>(runtime, {
          type: "START_EXPORT",
          data: requestData,
        });
      }
      if (response?.code === "EMPTY_RESULTS") {
        const message = response.error || emptyLabel();
        setStatus(message, { stopElapsed: true });
        showErrorBanner(message, false);
        await clearLastError(storage);
        return;
      }
      if (!response || response.error) {
        throw new Error(
          response?.error || t("status.error", {}, currentLang()),
        );
      }
      const langNow = currentLang();
      setStatus(
        response.filename
          ? `${t("status.complete", {}, langNow)} (${response.filename})`
          : t("status.complete", {}, langNow),
      );
      hideErrorBanner(true);
    } catch (e: any) {
      const raw = e?.message || "";
      const msg =
        raw.includes("Teams web app") || raw.includes("Teams tab")
          ? t("errors.needsTeams", {}, currentLang())
          : raw.includes("background")
            ? t("errors.ping", {}, currentLang())
            : raw || t("status.error", {}, currentLang());
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
      await applyLanguage(options.lang || "en");
      applyTheme(options.theme || "light");
      updateQuickRangeActive();
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
          type: "GET_EXPORT_STATUS",
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
            setStatus(t("status.running"));
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
  {#if showSettings}
    <div class="popup-content">
      <SettingsPage
        theme={options.theme}
        lang={options.lang || "en"}
        languages={languageOptions}
        on:back={() => (showSettings = false)}
        on:themeChange={(e) => updateOption("theme", e.detail)}
        on:langChange={(e) => updateOption("lang", e.detail)}
      />
    </div>
  {:else}
    <div class="popup-content">
      <!-- Header -->
      <header class="header">
        <h1>
          {t("title.app", {}, options.lang || "en") || "Teams Chat Exporter"}
        </h1>
        <HeaderActions
          on:openSettings={() => (showSettings = true)}
        />
      </header>

      <!-- Alert Banner -->
      {#if bannerMessage}
        <div class="alert error show" role="alert" aria-live="assertive">
          <span class="alert-title"
            >{t("banner.error", {}, options.lang || "en")}</span
          >
          <span>{bannerMessage}</span>
        </div>
      {/if}

      <!-- Export Button -->
      <ExportButton
        disabled={false}
        {busy}
        {busyLabel}
        summary={exportSummary}
        lang={options.lang || "en"}
        on:run={startExport}
      />

      <!-- Format Section (Full Width) -->
      <FormatSection
        format={options.format}
        lang={options.lang || "en"}
        on:formatChange={(e) => updateOption("format", e.detail)}
      />

      <!-- Two Column Grid: Date Range + Include -->
      <div class="settings-grid">
        <DateRangeSection
          startAt={options.startAt}
          endAt={options.endAt}
          activeRange={quickActive}
          ranges={quickRanges}
          lang={options.lang || "en"}
          {highlightMode}
          on:changeStart={(e) => updateOption("startAt", e.detail)}
          on:changeEnd={(e) => updateOption("endAt", e.detail)}
          on:quickSelect={(e) => handleQuickRange(e.detail)}
        />

        <IncludeSection
          includeReplies={options.includeReplies}
          includeReactions={options.includeReactions}
          includeSystem={options.includeSystem}
          embedAvatars={options.embedAvatars}
          downloadImages={options.downloadImages}
          lang={options.lang || "en"}
          disableReplies={options.format === "txt"}
          disableReactions={options.format === "txt"}
          disableAvatars={options.format === "txt" || options.format === "csv"}
          disableImages={options.format !== "html"}
          on:includeRepliesChange={(e) =>
            updateOption("includeReplies", e.detail)}
          on:includeReactionsChange={(e) =>
            updateOption("includeReactions", e.detail)}
          on:includeSystemChange={(e) =>
            updateOption("includeSystem", e.detail)}
          on:embedAvatarsChange={(e) => updateOption("embedAvatars", e.detail)}
          on:includeImagesChange={(e) =>
            updateOption("downloadImages", e.detail)}
        />
      </div>
    </div>

    <!-- Status Bar (Sticky Bottom) -->
    <StatusBar status={statusText} count={statusCount} isBusy={busy} />
  {/if}
</div>
