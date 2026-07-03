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
    loadHistory,
    loadHistoryViewedAt,
    loadLastError,
    loadOptions,
    markHistorySeen,
    persistErrorMessage,
    removeHistoryEntry as removeHistoryEntryFromStorage,
    clearHistory as clearHistoryStorage,
    loadSavedPresets,
    saveSavedPreset,
    removeSavedPreset,
    saveOptions,
    updateHistoryEntry,
    validateRange,
    LAST_PAGE_STORAGE_KEY,
    ACTIVE_EXPORTS_STORAGE_KEY,
    FIRST_INSTALL_STORAGE_KEY,
    REVIEW_PROMPT_STORAGE_KEY,
    PICKER_FOLDER_STORAGE_KEY,
    PICKER_KIND_STORAGE_KEY,
    PICKER_COLLAPSED_STORAGE_KEY,
    type OptionFormat,
    type Options,
    type PopupPage,
    type ReviewPromptResponse,
    type ReviewPromptState,
    type Theme,
  } from "../../utils/options";
  import type { ActiveExportInfo } from "../../types/shared";
  import { getIssueNewUrl, getReviewStoreUrl } from "../../utils/store-urls";
  import {
    formatElapsed,
    isoToLocalInput,
    localInputToISO,
  } from "../../utils/time";
  import { runtimeSend } from "../../utils/messaging";
  import { isTeamsUrl } from "../../utils/teams-urls";
  import { conversationDisplayName } from "../../utils/conversation-labels";
  import type {
    GetExportStatusRequest,
    GetExportStatusResponse,
    PingSWRequest,
    StartExportRequest,
    StartExportResponse,
    StartBundleExportRequest,
    StartBundleExportResponse,
    StopExportRequest,
  } from "../../types/messaging";
  import type { ConversationKind, ConversationSummary, ExportStatusPayload, FolderSummary, HistoryEntry, SavedPreset } from "../../types/shared";
  import type {
    ListConversationsRequest, ListConversationsResponse,
    ListConversationsQuickRequest, ListConversationsQuickResponse,
  } from "../../types/messaging";
  import ExportButton from "./components/ExportButton.svelte";
  import FormatSection from "./components/FormatSection.svelte";
  import TargetSection from "./components/TargetSection.svelte";
  import DateRangeSection, {
    type QuickRange,
  } from "./components/DateRangeSection.svelte";
  import IncludeSection from "./components/IncludeSection.svelte";
  import HeaderActions from "./components/HeaderActions.svelte";
  import SettingsPage from "./components/SettingsPage.svelte";
  import HistoryPage from "./components/HistoryPage.svelte";
  import DiagnosticsPage from "./components/DiagnosticsPage.svelte";
  import OnboardingOverlay from "./components/OnboardingOverlay.svelte";
  import ReviewPrompt from "./components/ReviewPrompt.svelte";
  import ConversationPicker from "./components/ConversationPicker.svelte";
  import { t, setLanguage, getLanguage } from "../../i18n/i18n";

  const runtime =
    typeof browser !== "undefined" ? browser.runtime : chrome.runtime;
  const tabs = typeof browser !== "undefined" ? browser.tabs : chrome.tabs;
  const storage =
    typeof browser !== "undefined" ? browser.storage : chrome.storage;

  type ExportStatusMsg = ExportStatusPayload;

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

  // Page routing — popup views: main, settings, history, diagnostics.
  // Only one can be visible at a time. The active view is persisted
  // under LAST_PAGE_STORAGE_KEY so reopening the popup resumes where
  // the user left off (fire-and-forget — page state is a UX nicety, a
  // dropped write just falls back to 'main' next time).
  let showSettings = false;
  let showHistory = false;
  let showDiagnostics = false;
  // Gate persistence until the page-restore in onMount has had a
  // chance to run. Without this gate, the reactive write below fires
  // on initial render with currentPage='main' and overwrites the
  // persisted value before the restore code can read it, defeating
  // the "resume where you left off" behaviour for anything other
  // than 'main'.
  let pageRestored = false;
  $: currentPage = showDiagnostics
    ? 'diagnostics'
    : showSettings
      ? 'settings'
      : showHistory
        ? 'history'
        : 'main';
  $: if (pageRestored) void persistCurrentPage(currentPage as PopupPage);

  async function persistCurrentPage(page: PopupPage) {
    try {
      await chrome.storage.local.set({ [LAST_PAGE_STORAGE_KEY]: page });
    } catch { /* best-effort; nothing hard-depends on this */ }
  }
  // Welcome overlay visibility. Becomes true once options load and the
  // persisted flag is still false; the user dismisses it once and it
  // stays dismissed across sessions.
  let showOnboarding = false;
  // Whether to skip the pre-tour prompt and start the walkthrough
  // immediately. true only when the user explicitly clicks Replay
  // tour in Settings — they've already opted in. First-time auto-
  // opens go through the prompt so users who already know the
  // extension can dismiss without sitting through 7 steps.
  let tourAutoStart = false;

  const dismissOnboarding = () => {
    showOnboarding = false;
    void updateOption("onboardingDismissed", true);
  };

  // Replay tour — invoked from Settings. Closes the settings page so
  // the popup's main view (the tour's actual targets) is mounted, then
  // shows the overlay. We don't flip onboardingDismissed back to false:
  // the persisted flag is "user has seen it once"; the replay is a
  // user-initiated reopen, not a re-onboarding. The overlay will write
  // onboardingDismissed=true again on its own dismiss.
  const replayTour = () => {
    showSettings = false;
    tourAutoStart = true;
    showOnboarding = true;
  };

  // Review-prompt gate state — inline one-liner rendered under the
  // export button when ALL of:
  //   - not yet shown (one-shot flag)
  //   - onboarding already dismissed (don't stack two overlays)
  //   - first install ≥ 7 days ago
  //   - ≥ 2 successful history entries
  //   - not currently busy (never compete with an in-progress export)
  //   - on the main page (not settings / history)
  // Rendered permanently once eligible until the user clicks rate,
  // feedback, or dismiss; persisted response in chrome.storage.local
  // under REVIEW_PROMPT_STORAGE_KEY.
  let reviewPromptState: ReviewPromptState = { shown: false };
  let firstInstalledAt: number | null = null;
  const REVIEW_MIN_AGE_MS = 7 * 24 * 60 * 60 * 1000;
  const REVIEW_MIN_EXPORTS = 2;
  $: successfulExportCount = historyEntries.filter(e => e.kind === 'success').length;
  $: reviewPromptEligible =
    !reviewPromptState.shown &&
    options.onboardingDismissed === true &&
    !busy &&
    currentPage === 'main' &&
    firstInstalledAt != null &&
    Date.now() - firstInstalledAt >= REVIEW_MIN_AGE_MS &&
    successfulExportCount >= REVIEW_MIN_EXPORTS;

  const reviewStoreUrl = getReviewStoreUrl();
  const reviewIssueUrl = getIssueNewUrl();

  async function onReviewPromptRespond(response: ReviewPromptResponse) {
    const next: ReviewPromptState = { shown: true, response, at: Date.now() };
    reviewPromptState = next;
    try {
      await chrome.storage.local.set({ [REVIEW_PROMPT_STORAGE_KEY]: next });
    } catch { /* best-effort; one-shot, worst case user is asked once more */ }
  }

  // ConversationPicker state. The popup fetches the user's chat list
  // from Teams' API once per open and caches it while the popup stays
  // mounted. If the API call fails (e.g. no valid token, Teams tab
  // logged out), we surface a retry button in the picker rather than
  // silently falling back to auto-detect — explicit failure modes beat
  // confusing "it exported the wrong chat" behaviour.
  let conversations: ConversationSummary[] = [];
  // Folder rail: list of Favorites + user-defined folders the picker
  // shows. Source of truth comes from the content script's IDB read
  // (see listConversationsFromIdb). System-computed folders are stripped
  // upstream — the picker never sees MeetingChats / MutedChats / etc.
  let folders: FolderSummary[] = [];
  // Currently selected folder id. 'all' means no folder filter (the
  // picker shows every conversation regardless of folder membership).
  // Persisted under PICKER_FOLDER_STORAGE_KEY so reopening the popup
  // keeps the user's filter; reset to 'all' when the saved id no
  // longer exists in the next read.
  let selectedFolderId: string = 'all';
  // Currently selected kind tab. Persisted under PICKER_KIND_STORAGE_KEY
  // so reopening the popup keeps the user's filter for symmetry with
  // the folder. Restored eagerly from storage during init() so the
  // very first paint reflects the saved choice — no flash of "all".
  let selectedKind: 'all' | ConversationKind = 'all';
  // Picker collapsed/expanded state. Persisted alongside the kind +
  // folder choices. True (collapsed) by default — popup opens minimal,
  // matching pre-v1.4.0 single-export habits. Even when collapsed, the
  // active chat is auto-picked into selectedConversationIds at line
  // ~354, so a single Export click works exactly like pre-v1.4.0.
  // Users who want bulk export expand once and the choice persists.
  let pickerCollapsed = true;
  // Multi-select: the picker binds this and mutates it on each toggle.
  // A single-chat export is simply `selectedConversationIds.length === 1`.
  // The first id is the "primary" selection for backwards-compatible
  // single-chat paths (export button label, scraper conversationId).
  let selectedConversationIds: string[] = [];
  $: selectedConversationId = selectedConversationIds[0] ?? null;
  // Id of the chat currently open in Teams (the active-chat hint), captured
  // when we resolve the default selection. Used only to label the export
  // button: a single selection that equals this is the "current chat"; a
  // single selection that differs is a "selected chat". Null when unknown
  // (e.g. not on a Teams tab) — then the button keeps its historical label.
  let activeChatId: string | null = null;
  $: singleSelectionIsOther =
    selectedConversationIds.length === 1 &&
    activeChatId != null &&
    selectedConversationId !== activeChatId;
  let pickerState: 'idle' | 'loading' | 'ok' | 'error' = 'idle';
  // True when the popup opened on a tab that isn't Teams web — we never
  // attempt to load conversations in that case, so without this flag the
  // picker would otherwise fall through to its "No matches" branch (the
  // default empty-list rendering when state is 'idle' rather than
  // 'loading' or 'error'). This tells the picker to show a dedicated
  // "open Teams web first" message instead.
  let notOnTeamsTab = false;
  let pickerError = '';

  // Folder persistence — fire-and-forget on every change. The read on
  // mount happens inside init() once we have the conversation list.
  async function persistFolderChoice(id: string) {
    try { await storage.local.set({ [PICKER_FOLDER_STORAGE_KEY]: id }); }
    catch { /* best-effort */ }
  }
  async function readSavedFolderChoice(): Promise<string | null> {
    try {
      const obj: any = await storage.local.get(PICKER_FOLDER_STORAGE_KEY);
      const v = obj?.[PICKER_FOLDER_STORAGE_KEY];
      return typeof v === 'string' ? v : null;
    } catch { return null; }
  }
  // Same pattern for the kind. Validates against the known set so a
  // corrupt storage value doesn't propagate through the rest of the
  // picker (which assumes selectedKind is one of the five literal
  // strings).
  const VALID_KINDS: ReadonlyArray<'all' | ConversationKind> = ['all', 'chat', 'group', 'meeting', 'channel'];
  async function persistKindChoice(k: 'all' | ConversationKind) {
    try { await storage.local.set({ [PICKER_KIND_STORAGE_KEY]: k }); }
    catch { /* best-effort */ }
  }
  async function readSavedKindChoice(): Promise<'all' | ConversationKind | null> {
    try {
      const obj: any = await storage.local.get(PICKER_KIND_STORAGE_KEY);
      const v = obj?.[PICKER_KIND_STORAGE_KEY];
      return (typeof v === 'string' && (VALID_KINDS as readonly string[]).includes(v))
        ? v as ('all' | ConversationKind)
        : null;
    } catch { return null; }
  }
  // Picker collapse persistence — same fire-and-forget pattern. Stored
  // strictly as a boolean; anything else is treated as missing and the
  // default (expanded) wins.
  async function persistPickerCollapsed(v: boolean) {
    try { await storage.local.set({ [PICKER_COLLAPSED_STORAGE_KEY]: v }); }
    catch { /* best-effort */ }
  }
  async function readSavedPickerCollapsed(): Promise<boolean | null> {
    try {
      const obj: any = await storage.local.get(PICKER_COLLAPSED_STORAGE_KEY);
      const v = obj?.[PICKER_COLLAPSED_STORAGE_KEY];
      return typeof v === 'boolean' ? v : null;
    } catch { return null; }
  }
  // Reconcile the persisted folder against the freshly-loaded folder
  // list. If the saved id no longer exists (deleted in Teams, or this
  // is a new install with no saved value), fall back to 'all'.
  function reconcileFolderChoice(saved: string | null | undefined) {
    if (!saved || saved === 'all') { selectedFolderId = 'all'; return; }
    if (folders.some(f => f.id === saved)) { selectedFolderId = saved; return; }
    selectedFolderId = 'all';
    void persistFolderChoice('all');
  }

  // Stale-while-revalidate cache for the conversation list. First cold
  // fetch on a ~100-conversation account takes ~5–8s (Graph joinedTeams +
  // per-unnamed roster calls). Persisting the result keyed on the
  // MSAL-resolved self-UUID lets every subsequent popup open render
  // instantly; we then fire a refresh in the background and swap in the
  // updated list when it arrives.
  const CONV_LIST_CACHE_KEY = 'convListCache';
  const CONV_LIST_CACHE_TTL_MS = 24 * 60 * 60_000; // 1 day — full discard
  // Bump when the ConversationSummary shape changes in a way that makes
  // older cached rows useless (e.g. adding isSelfChat, folders). Old
  // caches are dropped rather than served stale; users see one cold
  // fetch after the extension updates, then instant loads again.
  const CONV_LIST_CACHE_VERSION = 11;
  let isRefreshingList = false;

  type CachedConvList = {
    version?: number;
    at: number;
    conversations: ConversationSummary[];
    // Folder rail data — cached alongside conversations so the picker
    // shows folders on instant cache hit instead of waiting for the
    // background full refresh (~5–8s).
    folders?: FolderSummary[];
  };

  async function readCachedList(): Promise<CachedConvList | null> {
    try {
      const obj: any = await storage.local.get(CONV_LIST_CACHE_KEY);
      const c = obj?.[CONV_LIST_CACHE_KEY];
      if (!c || typeof c.at !== 'number' || !Array.isArray(c.conversations)) return null;
      if (c.version !== CONV_LIST_CACHE_VERSION) return null;
      if (Date.now() - c.at > CONV_LIST_CACHE_TTL_MS) return null;
      return c as CachedConvList;
    } catch { return null; }
  }
  async function writeCachedList(list: ConversationSummary[], folderList: FolderSummary[]) {
    try {
      await storage.local.set({
        [CONV_LIST_CACHE_KEY]: {
          version: CONV_LIST_CACHE_VERSION,
          at: Date.now(),
          conversations: list,
          folders: folderList,
        } satisfies CachedConvList,
      });
    } catch { /* best-effort */ }
  }

  // Ask the content script for the active-chat hint, with a short
  // retry budget for cases where the Teams sidebar hadn't finished
  // rendering its rows on first call (extractConversationId returns
  // null until the sidebar's data-tabster attributes are populated).
  // Three quick attempts is enough to cover the post-tab-switch
  // animation window without making the user wait when the hint
  // genuinely isn't available.
  async function fetchActiveHint(tab: number): Promise<string | null> {
    const delays = [0, 250, 600];
    for (const delay of delays) {
      if (delay > 0) await new Promise(r => setTimeout(r, delay));
      try {
        const idResp = await tabs.sendMessage(tab, { type: 'GET_CONV_ID' });
        const hint: string | null = idResp?.convId || null;
        if (hint) return hint;
      } catch { return null; }
    }
    return null;
  }

  // Resolves the default selection using the Teams-tab active-chat
  // hint (extractConversationId in the content script). Safe to call
  // on both cache-hit and after-refresh — if the hint matches a row
  // in the latest list, we align the selection; otherwise we don't
  // touch it. This matters because the hint might initially miss a
  // stale cache (e.g. self-chat absent from v1 cached rows) and only
  // match once the fresh list arrives.
  //
  // Fires at most ONCE per popup open: the cache, IDB-quick, and full-Graph
  // phases each call this, but only the first one that finds the active
  // chat actually seeds the selection. After that the user owns the
  // selection — even if they unselect the auto-pick, the next refresh
  // phase MUST NOT silently reinstate it.
  let didAutoPickDefault = false;
  async function pickDefaultSelection(list: ConversationSummary[]) {
    if (didAutoPickDefault) return;
    if (selectedConversationIds.length > 0) {
      // User already picked something (or unpicked the prior auto-default).
      // Lock in: never auto-modify their selection from here on.
      didAutoPickDefault = true;
      return;
    }
    try {
      const tab = currentTabId;
      if (typeof tab !== 'number') return;
      const hint = await fetchActiveHint(tab);
      if (!hint) return;
      // Remember the open chat so the export button can tell a "current
      // chat" selection apart from a different single chat the user picks.
      activeChatId = hint;
      const hit = list.find(c => c.id === hint);
      if (!hit) return;
      // Re-check selection length — a refresh phase could resolve between
      // the await above and this assignment, and the user might have
      // touched the picker in the meantime.
      if (selectedConversationIds.length === 0) {
        selectedConversationIds = [hit.id];
      }
      didAutoPickDefault = true;
    } catch { /* best-effort — leave whatever selection exists */ }
  }

  async function loadConversations(opts: { forceRefresh?: boolean } = {}) {
    if (currentTabId == null) return;

    // Hydrate from cache immediately when available. The user sees a
    // populated picker within a few ms.
    const cached = opts.forceRefresh ? null : await readCachedList();
    if (cached && cached.conversations.length > 0) {
      conversations = cached.conversations;
      folders = Array.isArray(cached.folders) ? cached.folders : [];
      const savedCached = await readSavedFolderChoice();
      reconcileFolderChoice(savedCached);
      pickerState = 'ok';
      await pickDefaultSelection(conversations);
    } else {
      pickerState = 'loading';
    }
    pickerError = '';

    // ALWAYS run the IDB-quick read on every open. It's ~50 ms and tells
    // us exactly which conversations exist + their last-activity stamps
    // — that's what the user perceives as "is the picker fresh?". The
    // expensive part (Graph + roster name resolution) is the FULL refresh
    // below; we only fire that when the IDB diff says we need it.
    let quickResult: { conversations: ConversationSummary[]; folders: FolderSummary[] } | null = null;
    try {
      const quick = await runtimeSend<ListConversationsQuickRequest>(runtime, {
        type: 'LIST_CONVERSATIONS_QUICK',
        tabId: currentTabId,
      }) as ListConversationsQuickResponse;
      if (quick && quick.ok === true) {
        quickResult = {
          conversations: quick.conversations,
          folders: Array.isArray(quick.folders) ? quick.folders : [],
        };
      }
    } catch { /* fall through — full refresh below covers cold-IDB cases */ }

    // Apply the quick result on top of cache: take IDB's truth for the
    // conversation set + activity timestamps + folder list, but PATCH
    // names from cache (cache holds Graph-resolved names, quick does not
    // unless the chat has chatTitle.shortTitle in IDB). This is what
    // makes "open popup" feel instant for repeat opens — no Graph spin
    // on a chat we already know the name of.
    let needsFullRefresh = !!opts.forceRefresh;
    if (quickResult && quickResult.conversations.length > 0) {
      const cachedById = new Map<string, ConversationSummary>(
        (cached?.conversations || []).map(c => [c.id, c]),
      );
      const merged: ConversationSummary[] = quickResult.conversations.map(qc => {
        const cc = cachedById.get(qc.id);
        if (!cc) return qc; // genuinely new — keep IDB's (possibly empty) name
        // Prefer cached enriched fields, but always take IDB's
        // lastActivity (it's the freshest) and folderIds (folder
        // membership lives in IDB and may have changed).
        return {
          ...cc,
          lastActivity: qc.lastActivity || cc.lastActivity,
          folderIds: qc.folderIds,
        };
      });
      conversations = merged;
      folders = quickResult.folders;
      const savedQuick = await readSavedFolderChoice();
      reconcileFolderChoice(savedQuick);
      pickerState = 'ok';
      await pickDefaultSelection(conversations);

      // Decide whether the slow Graph refresh is worth running. If
      // every IDB chat already has a cached name AND the folder set
      // hasn't changed, the cache covers everything — skip the
      // expensive call. Self-chat is allowed to have an empty name
      // (the popup renders a localised "(You)" label).
      const anyMissingName = merged.some(c => !c.name && !c.isSelfChat);
      const cachedFolderIds = new Set((cached?.folders || []).map(f => f.id));
      const newFolderIds = quickResult.folders.filter(f => !cachedFolderIds.has(f.id));
      const removedFolders = (cached?.folders || []).some(f => !quickResult!.folders.some(qf => qf.id === f.id));
      if (anyMissingName || newFolderIds.length > 0 || removedFolders) needsFullRefresh = true;

      // Even when we skip the full refresh, persist the merged result
      // so the cache stays current with IDB activity timestamps and
      // folder-set changes that don't require Graph.
      if (!needsFullRefresh) {
        await writeCachedList(conversations, folders);
      }
    } else if (!cached) {
      // Quick failed AND no cache — must run the full refresh to have
      // anything to show.
      needsFullRefresh = true;
    }

    if (!needsFullRefresh) {
      isRefreshingList = false;
      return;
    }

    isRefreshingList = true;
    try {
      const resp = await runtimeSend<ListConversationsRequest>(runtime, {
        type: 'LIST_CONVERSATIONS',
        tabId: currentTabId,
      }) as ListConversationsResponse;
      if (!resp || resp.ok !== true) {
        // If we already showed cached/quick data, keep it visible —
        // don't flip the user back to an error screen just because a
        // background refresh failed. Only surface error when the
        // picker was empty.
        if (conversations.length === 0) {
          pickerState = 'error';
          pickerError = (resp && 'error' in resp ? resp.error : '') || 'load-failed';
        }
        return;
      }
      conversations = resp.conversations;
      folders = Array.isArray(resp.folders) ? resp.folders : [];
      // Re-reconcile after the full enrichment lands — the folder set
      // can change (e.g. quick path saw stale IDB before refresh).
      const savedFull = await readSavedFolderChoice();
      reconcileFolderChoice(savedFull);
      pickerState = 'ok';
      await writeCachedList(conversations, folders);
      await pickDefaultSelection(conversations);
    } catch (e: any) {
      if (conversations.length === 0) {
        pickerState = 'error';
        pickerError = String(e?.message || e);
      }
    } finally {
      isRefreshingList = false;
    }
  }

  // History state. Loaded once on mount, refreshed when an entry is appended
  // (popup hears phase=complete or phase=cancelled), and after the user opens
  // the History page (which marks them all as seen).
  let historyEntries: HistoryEntry[] = [];
  let savedPresets: SavedPreset[] = [];

  // Saved selection presets: loaded on mount; mutated via the picker's Presets
  // menu. Internal identifiers and the storage key keep the legacy "group"
  // name so existing saved data survives; only the UI label changed.
  // Apply is handled inside the picker (re-selects); save/remove come here so
  // persistence stays in App alongside the other storage writes.
  const refreshSavedPresets = async () => {
    if (!alive) return;
    savedPresets = await loadSavedPresets(storage);
  };
  async function onSavePreset(name: string) {
    const now = Date.now();
    const preset: SavedPreset = {
      id: crypto.randomUUID(),
      name,
      convIds: [...selectedConversationIds],
      createdAt: now,
      updatedAt: now,
    };
    await saveSavedPreset(storage, preset);
    await refreshSavedPresets();
  }
  async function onRemovePreset(id: string) {
    await removeSavedPreset(storage, id);
    await refreshSavedPresets();
  }
  let lastHistoryViewedAt = 0;
  // Number of entries added since the last visit to the History page.
  // Drives the count badge on the history icon (capped to "9+" in
  // HeaderActions for display).
  $: newHistoryCount = historyEntries.filter(e => e.savedAt > lastHistoryViewedAt).length;

  // Counters that increment on success / cancel — passed to ExportButton and
  // HeaderActions to trigger one-shot animations (button flash, icon pulse).
  // Using a counter (not a boolean) lets the same animation re-fire on the
  // next export without us having to manually toggle the trigger off.
  let successFlashTrigger = 0;
  let pulseHistoryIcon = 0;

  let options: Options = { ...DEFAULT_OPTIONS };
  const currentLang = () => options.lang || "en";
  const runLabel = () => t(`actions.export.${options.exportTarget}`, {}, currentLang());
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
  let statusCountLabel = '';
  let alive = true;
  let busy = false;
  let busyLabel = runLabel();
  // True from popup mount until GET_EXPORT_STATUS has replied. While
  // this is true the export button renders a neutral "checking…" label
  // instead of its idle default, so when the real status arrives (busy
  // or idle) the user doesn't see a brief wrong-state flash.
  let statusKnown = false;
  let currentTabId: number | null = null;
  let startedAtMs: number | null = null;
  let elapsedTimer: ReturnType<typeof setInterval> | null = null;
  let exportSummary = "";

  // Phase-tracker state for the export button. The 4 segments map to
  // (messages · images · people · file). A segment value of `null` means
  // "not started yet" (dim), `-1` means "active, indeterminate" (animated
  // stripe), and 0..100 means a determinate fill percentage.
  type SegState = number | null;
  let phaseLabel = '';
  let phaseBaseLabel = '';
  let counterValue = '—';
  let counterLabel = '';
  let segments: SegState[] = [null, null, null, null];

  // Bundle context — present while a multi-chat export is running. Drives
  // the "Chat 3 of 12" prefix on the phase tracker label so the user sees
  // overall position in the bundle (the per-chat name was dropped to keep
  // the one-line detail from overflowing).
  type BundleContext = {
    current: number;
    total: number;
  };
  let bundleContext: BundleContext | null = null;

  const resetPhaseTracker = () => {
    phaseLabel = '';
    phaseBaseLabel = '';
    counterValue = '—';
    counterLabel = '';
    segments = [null, null, null, null];
    bundleContext = null;
  };

  // Re-read history from storage (called after a phase=complete / cancelled
  // arrives, after the user removes a row, etc.). Cheap (~ms).
  const refreshHistory = async () => {
    if (!alive) return;
    const [list, viewedAt] = await Promise.all([
      loadHistory(storage),
      loadHistoryViewedAt(storage),
    ]);
    if (!alive) return;
    historyEntries = list;
    lastHistoryViewedAt = viewedAt;
  };

  const elapsedNow = () =>
    startedAtMs != null ? formatElapsed(Date.now() - startedAtMs) : '';

  const refreshPhaseLabel = () => {
    if (!phaseBaseLabel) return;
    const elapsed = elapsedNow();
    const parts: string[] = [];
    if (bundleContext) {
      const lang = currentLang();
      // Bundle prefix is just "Chat X of Y" now — the per-chat name was
      // dropped from the phase label (it kept overflowing the one-line
      // detail and the name is already visible in the picker).
      const prefix = t(
        'status.bundleProgress',
        { current: bundleContext.current, total: bundleContext.total },
        lang,
      ) || `Chat ${bundleContext.current} of ${bundleContext.total}`;
      parts.push(prefix);
    }
    parts.push(phaseBaseLabel);
    if (elapsed) parts.push(elapsed);
    phaseLabel = parts.join(' · ');
  };

  const setPhase = (
    idx: 0 | 1 | 2 | 3,
    activeProgress: SegState,
    label: string,
    value: string,
    valueLabel: string,
  ) => {
    const next: SegState[] = [...segments];
    for (let i = 0; i < 4; i++) {
      if (i < idx) next[i] = 100;
      else if (i === idx) next[i] = activeProgress;
      else next[i] = null;
    }
    // Skip the reassignment when nothing changed. Without this guard,
    // every progress tick fires a new array reference and Svelte
    // re-applies inline styles on .seg-fill, which restarts the CSS
    // stripe animation on indeterminate segments — the user perceives
    // this as the bar repeatedly "resetting and starting over".
    const changed = next.some((v, i) => v !== segments[i]);
    if (changed) segments = next;
    phaseBaseLabel = label;
    refreshPhaseLabel();
    counterValue = value;
    counterLabel = valueLabel;
  };

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
    // Also refresh the export-button's phase label so its " · 0:12" suffix
    // ticks every second (the same elapsedTimer drives both surfaces).
    refreshPhaseLabel();
  };

  const computeSummary = () => {
    const parts: string[] = [];
    const lang = currentLang();

    const targetLabel = t(`target.${options.exportTarget}`, {}, lang);
    parts.push(targetLabel);

    // Format(s). For multi-format we list all selected, then suffix "(zip)"
    // since the output is a bundle.
    const labels = options.formats.map(f => t(`format.${f}`, {}, lang));
    const formatLabel = labels.join(", ");
    if (options.formats.length >= 2) {
      parts.push(`${formatLabel} (zip)`);
    } else {
      parts.push(formatLabel);
    }

    // Date range
    if (quickActive && quickActive !== "none") {
      const rangeLabel = quickRanges.find((r) => r.key === quickActive)?.label;
      if (rangeLabel) parts.push(rangeLabel);
    }

    // The include-options toggle list (replies/reactions/system/avatars/
    // images) used to be appended here. It was dropped: the line kept
    // overflowing the one-line button detail, and the toggles are already
    // visible right above the button. Summary is now just target • format •
    // range.
    return parts.join(" • ");
  };

  // Update summary when options change. Only target, formats and the active
  // quick-range feed the summary now (the include toggles were dropped), so
  // those are the only reactive dependencies we touch here.
  $: {
    options.exportTarget;
    options.formats;
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
      countLabel?: string;
    } = {},
  ) => {
    if (!alive) return;
    statusBaseText = text;
    if (typeof opts.count === "number") statusCount = opts.count;
    statusCountLabel = opts.countLabel ?? '';
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
    // "Switch to the Chat/Teams app" error paths were removed in the
    // locale-independent checkChatContext refactor (issues #10/#19) —
    // the content script now only emits "Open a chat/team" messages.
    if (message.includes("Open a chat conversation")) {
      return t("errors.chatNotOpen", {}, lang);
    }
    if (message.includes("Open a team channel")) {
      return t("errors.teamNotOpen", {}, lang);
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
    // Capture bundle context from any status payload that carries it.
    // The SW broadcasts these fields on every per-chat status during a
    // bundle run; capturing them here lets refreshPhaseLabel render
    // "Chat N of M" without a separate code path per phase.
    if (typeof msg?.bundleCurrentChat === 'number' && typeof msg?.bundleTotalChats === 'number') {
      bundleContext = {
        current: msg.bundleCurrentChat,
        total: msg.bundleTotalChats,
      };
    }
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
      const count = msg.messages ?? 0;
      setStatus(t("status.building", {}, langNow), { count });
    } else if (phase === "build") {
      setBusy(true, busyBuildLabel());
      const count = msg.messages ?? 0;
      const fname = msg.filename || '';
      const isZip = fname.endsWith('.zip');
      const key = isZip ? "status.compressing" : "status.building";
      setStatus(t(key, {}, langNow), { count });
    } else if (phase === "downloading-files") {
      // Attachment download (the "Files" toggle). Count is only known after the
      // scrape, so this phase fires after the main export is saved.
      setBusy(true, busyBuildLabel());
      const done = msg.filesDone ?? 0;
      const total = msg.filesTotal ?? 0;
      setStatus(
        t("status.downloadingFiles", { done, total }, langNow),
        total > 0 ? { countLabel: `${done}/${total}` } : {},
      );
    } else if (phase === "empty") {
      const message = msg.message || emptyLabel();
      setBusy(false);
      setStatus(message, { stopElapsed: true });
      showErrorBanner(message, false);
      void clearLastError(storage);
    } else if (phase === "complete") {
      setBusy(false);
      segments = [100, 100, 100, 100];
      // Drop the bundle prefix on success so a future "show last phase
      // even when idle" path doesn't render stale "Chat 12 of 12" text.
      // The user-facing status is now the bundle-summary line set above.
      bundleContext = null;
      // Append the attachment-download summary when the Files toggle ran.
      // Only non-zero buckets are shown so a clean run reads "… · N files saved".
      const fs = msg.filesSummary;
      let completeText = t("status.complete", {}, langNow);
      if (fs && fs.total > 0) {
        // Counts are settled outcomes: the background now waits for every
        // attachment download to finish and verify before broadcasting
        // 'complete', so `saved` means on disk and `failed` already includes
        // files verified as inaccessible (they're also listed in FAILURES.txt).
        // fs.cancelled (files stopped mid-download) is not itemized here —
        // stopping was the user's own action and FAILURES.txt lists them.
        const parts = [t("status.filesSaved", { n: fs.saved }, langNow)];
        // Host-gate links kept as links rather than downloaded.
        if (fs.links > 0) parts.push(t("status.filesLinks", { n: fs.links }, langNow));
        if (fs.failed > 0) parts.push(t("status.filesFailed", { n: fs.failed }, langNow));
        completeText += " · " + parts.join(" · ");
      }
      setStatus(completeText, { stopElapsed: true });
      hideErrorBanner(true);
      // The success path no longer renders an inline outcome tile. The
      // background appends a row to the persisted history; we re-read it
      // here so the dot on the history icon updates immediately. Then we
      // bump the animation triggers — ExportButton runs a green flash on
      // the bar and HeaderActions pulses the history icon.
      void refreshHistory();
      successFlashTrigger += 1;
      pulseHistoryIcon += 1;
    } else if (phase === "error") {
      setBusy(false);
      bundleContext = null;
      setStatus(msg.error || t("status.error", {}, langNow), {
        stopElapsed: true,
      });
      showErrorBanner(msg.error || t("status.error", {}, langNow));
    } else if (phase === "cancelling") {
      // Background acknowledged the stop; show feedback while teardown runs.
      setStatus(
        t("status.cancelling", {}, langNow) || "Cancelling…",
        { stopElapsed: true },
      );
    } else if (phase === "cancelled") {
      setBusy(false);
      // Reset the phase tracker rather than freezing it. The button is a
      // pure Export button again now that outcomes live in the History page.
      resetPhaseTracker();
      setStatus(
        t("status.cancelled", {}, langNow) || "Cancelled",
        { stopElapsed: true },
      );
      hideErrorBanner(true);
      // Background appended a cancelled row; refresh the dot + nudge the
      // history icon so the user knows something changed.
      void refreshHistory();
      pulseHistoryIcon += 1;
    }
  };

  const onRuntimeMessage = (msg: any) => {
    // Count EXPORT_STATUS / EXPORT_PROGRESS / SCRAPE_PROGRESS arrivals so
    // __reportExportStatusRate can log the per-second rate. A high rate
    // (> ~50/s) during a bundle export points at a broadcast storm — the
    // suspected cause of the popup white-pixel symptom.
    if (msg?.type === "EXPORT_STATUS" || msg?.type === "EXPORT_PROGRESS" || msg?.type === "SCRAPE_PROGRESS") {
      __exportStatusCount += 1;
    }
    if (msg?.type === "SCRAPE_PROGRESS" || msg?.type === "EXPORT_PROGRESS") {
      const langNow = currentLang();
      const p = msg?.type === "EXPORT_PROGRESS" ? msg : (msg.payload || {});
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
      } else if (p.phase === "api-fetch") {
        const count = p.messagesVisible ?? p.messagesSoFar ?? 0;
        const label = t("phase.messages", {}, langNow);
        setStatus(`${label}…`, { count });
        // Pagination has no known total — drive segment 0 with the
        // indeterminate stripe (-1) and show the running count on the right.
        setPhase(0, -1, label, count.toLocaleString(), t("phase.label.messages", {}, langNow));
      } else if (p.phase === "images") {
        const done = p.imagesDone ?? 0;
        const total = p.imagesTotal ?? 0;
        const label = t("phase.images", {}, langNow);
        const lbl = t("phase.label.images", {}, langNow);
        if (total > 0) {
          setStatus(`${label}…`, { countLabel: `${done}/${total}` });
          setPhase(1, Math.round((done / total) * 100), label, `${done} / ${total}`, lbl);
        } else {
          // No total yet — show 0% fill rather than an indeterminate
          // full-width stripe. Otherwise the bar appears "full", then
          // snaps back to the real percentage on the next event,
          // which reads like a regression instead of early progress.
          setStatus(`${label}…`);
          setPhase(1, 0, label, '—', lbl);
        }
      } else if (p.phase === "avatars") {
        const done = p.avatarsDone ?? 0;
        const total = p.avatarsTotal ?? 0;
        const label = t("phase.avatars", {}, langNow);
        const lbl = t("phase.label.people", {}, langNow);
        if (total > 0) {
          setStatus(`${label}…`, { countLabel: `${done}/${total}` });
          setPhase(2, Math.round((done / total) * 100), label, `${done} / ${total}`, lbl);
        } else {
          // Same reasoning as 'images' above — start at 0%, not 100%.
          setStatus(`${label}…`);
          setPhase(2, 0, label, '—', lbl);
        }
      } else if (p.phase === "build") {
        const done = p.messagesBuilt ?? 0;
        const total = p.messagesTotal ?? 0;
        const label = t("phase.build", {}, langNow);
        const lbl = t("phase.label.written", {}, langNow);
        if (total > 0) {
          setStatus(`${label}…`, { countLabel: `${done}/${total}` });
          setPhase(3, Math.round((done / total) * 100), label, `${done} / ${total}`, lbl);
        } else {
          setStatus(`${label}…`);
          setPhase(3, -1, label, '—', lbl);
        }
      }
    } else if (msg?.type === "EXPORT_STATUS") {
      handleExportStatus(msg);
    }
  };

  const stopExport = async () => {
    if (!busy || !alive) return;
    const tabId = currentTabId;
    if (typeof tabId !== "number") return;
    setStatus(
      t("status.cancelling", {}, currentLang()) || "Cancelling…",
      { stopElapsed: true },
    );
    try {
      await runtimeSend<StopExportRequest>(runtime, {
        type: "STOP_EXPORT",
        tabId,
      });
    } catch {
      // Even if the message round-trip fails, the export's own teardown
      // path will still run when the content script catches the abort.
    }
  };

  const startExport = async () => {
    if (busy || !alive) return;
    // Pre-flight: refuse to start when the browser tells us we're
    // offline. navigator.onLine === false is high-confidence — browsers
    // report it conservatively (rarely a false positive). The inverse
    // (online === true while actually offline, e.g. captive portal) is
    // the common false reading; that case is caught later by mid-export
    // NetworkError detection. So we only refuse on the clear signal.
    if (typeof navigator !== 'undefined' && navigator.onLine === false) {
      showErrorBanner(t('errors.offline', {}, currentLang()) || 'No network. Reconnect and try again.');
      return;
    }
    try {
      hideErrorBanner(true);
      // Reset the phase tracker so the new export starts at a clean state.
      resetPhaseTracker();
      setBusy(true, busyExportLabel());
      setStatus(t("status.preparing", {}, currentLang()));
      const tab = await getActiveTeamsTab();
      if (!alive) return;
      currentTabId = tab.id ?? null;
      await pingSW();
      const range = getValidatedRangeISO();
      const formats = options.formats;
      const {
        includeReplies,
        includeReactions,
        includeSystem,
        embedAvatars,
        downloadImages,
        downloadFiles,
        fullResImages,
        imageFilenameDate,
        imageModifiedDate,
        exportTarget,
      } = options;
      setStatus(t("status.running", {}, currentLang()));

      // Multi-chat bundle path. When the picker has 2+ selections we
      // route to START_BUNDLE_EXPORT instead — same per-chat scrape
      // pipeline runs serially in the SW, output packed into one outer
      // zip with FAILURES.txt for any chat that errored.
      if (selectedConversationIds.length > 1) {
        const langNow = options.lang || 'en';
        const conversationsPayload = selectedConversationIds.map((id) => {
          const sel = conversations.find(c => c.id === id);
          const title = sel ? conversationDisplayName(sel, langNow, t) : id;
          return { id, title };
        });
        const bundleData = {
          tabId: tab.id,
          conversations: conversationsPayload,
          scrapeOptions: {
            startAt: range.startISO,
            endAt: range.endISO,
            includeReplies,
            includeReactions,
            includeSystem,
            exportTarget,
            formats,
            embedAvatars,
            downloadImages,
            fullResImages,
            // No conversationId / conversationTitle here — the SW injects
            // them per-iteration from `conversations`.
          },
          buildOptions: {
            formats,
            saveAs: true,
            embedAvatars,
            downloadImages,
            downloadFiles,
            imageFilenameDate,
            imageModifiedDate,
            afterExport: options.afterExport,
            avatarMode: options.avatarMode,
            pdfPageSize: options.pdfPageSize,
            pdfBodyFontSize: options.pdfBodyFontSize,
            pdfShowPageNumbers: options.pdfShowPageNumbers,
            pdfIncludeAvatars: options.pdfIncludeAvatars,
          },
        };
        const bundleResp: StartBundleExportResponse =
          await runtimeSend<StartBundleExportRequest>(runtime, {
            type: "START_BUNDLE_EXPORT",
            data: bundleData,
          });
        if (bundleResp?.code === "EMPTY_RESULTS") {
          const message = bundleResp.error || emptyLabel();
          setStatus(message, { stopElapsed: true });
          showErrorBanner(message, false);
          await clearLastError(storage);
          return;
        }
        if (bundleResp?.cancelled) {
          setStatus(
            t("status.cancelled", {}, currentLang()) || "Cancelled",
            { stopElapsed: true },
          );
          return;
        }
        if (!bundleResp || bundleResp.error) {
          throw new Error(
            bundleResp?.error || t("status.error", {}, currentLang()),
          );
        }
        const langNow2 = currentLang();
        const success = bundleResp.successChats ?? 0;
        const total = bundleResp.totalChats ?? success;
        const failed = bundleResp.failedChats ?? 0;
        const summaryKey = failed > 0 ? 'status.bundleCompleteWithFailures' : 'status.bundleComplete';
        const summaryText = t(summaryKey, { success, total, failed }, langNow2)
          || `Bundle export complete: ${success} of ${total} chats${failed ? ` (${failed} failed)` : ''}.`;
        setStatus(
          bundleResp.filename
            // Leaf only: in package mode (Files on) the filename is a
            // Downloads-relative path ('folder/file.zip').
            ? `${summaryText} (${bundleResp.filename.split("/").pop()})`
            : summaryText,
        );
        hideErrorBanner(true);
        return;
      }

      const requestData = {
        tabId: tab.id,
        scrapeOptions: {
          startAt: range.startISO,
          endAt: range.endISO,
          includeReplies,
          includeReactions,
          includeSystem,
          exportTarget,
          // Let the scraper skip image/avatar fetches that none of the
          // selected formats would render. The SW reads the same array
          // when deciding single-file vs HTML.zip vs bundle.zip.
          formats,
          embedAvatars,
          downloadImages,
          fullResImages,
          // Explicit conversation chosen by the picker. When set, the
          // scraper skips its DOM/IDB auto-detection and targets this
          // conversation directly.
          conversationId: selectedConversationId,
          // Picker-resolved display name, so the scraper stamps the
          // chosen chat's name on meta.title (and thus the filename)
          // instead of pulling whichever chat is currently visible in
          // the Teams tab DOM. Mirrors the picker's locale-aware
          // placeholder logic: resolves "(You)" suffixes, unnamed-
          // group member lists, and kind-specific fallbacks in the
          // active UI language so filenames don't collide and stay
          // consistent with what the user selected in the picker.
          conversationTitle: (() => {
            const sel = conversations.find(c => c.id === selectedConversationId);
            if (!sel) return null;
            return conversationDisplayName(sel, options.lang || 'en', t);
          })(),
        },
        buildOptions: {
          formats,
          saveAs: true,
          embedAvatars,
          downloadImages,
          downloadFiles,
          imageFilenameDate,
          imageModifiedDate,
          afterExport: options.afterExport,
          avatarMode: options.avatarMode,
          pdfPageSize: options.pdfPageSize,
          pdfBodyFontSize: options.pdfBodyFontSize,
          pdfShowPageNumbers: options.pdfShowPageNumbers,
          pdfIncludeAvatars: options.pdfIncludeAvatars,
        },
      };
      const response: StartExportResponse = await runtimeSend<StartExportRequest>(runtime, {
        type: "START_EXPORT",
        data: requestData,
      });
      if (response?.code === "EMPTY_RESULTS") {
        const message = response.error || emptyLabel();
        setStatus(message, { stopElapsed: true });
        showErrorBanner(message, false);
        await clearLastError(storage);
        return;
      }
      // User pressed Stop while the export was running. The background already
      // broadcast a 'cancelled' status and tore down everything; here we just
      // show a clean message and skip the error banner.
      if (response?.cancelled) {
        setStatus(
          t("status.cancelled", {}, currentLang()) || "Cancelled",
          { stopElapsed: true },
        );
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
          // Leaf only: in package mode (Files on) the filename is a
          // Downloads-relative path ('folder/file.zip').
          ? `${t("status.complete", {}, langNow)} (${response.filename.split("/").pop()})`
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

  // Handlers passed down to HistoryPage. They mutate storage and re-read it
  // so the list and dot stay in sync.
  const onHistoryRemove = async (id: string) => {
    await removeHistoryEntryFromStorage(storage, id);
    await refreshHistory();
  };
  const onHistoryClear = async () => {
    await clearHistoryStorage(storage);
    await refreshHistory();
  };
  // Persist the "file is missing on disk" observation. Without this, the
  // grayed-out state would only survive within one popup session — on
  // Firefox in particular, downloads.search() on next mount returns stale
  // 'exists: true' and the row reappears as if the file is still there.
  const onHistoryMarkMissing = async (id: string) => {
    await updateHistoryEntry(storage, id, { fileExists: false });
    await refreshHistory();
  };
  // Fired when the user navigates into the History page — record the visit
  // so the dot clears and stays cleared until a new entry is added.
  const onHistoryOpened = async () => {
    await markHistorySeen(storage);
    lastHistoryViewedAt = Date.now();
  };

  // Debug: timestamped popup-mount tracer + EXPORT_STATUS broadcast rate
  // counter. Helps diagnose the "popup opens to a white pixel" symptom
  // observed during heavy bundle exports — the symptom is consistent with
  // a render/init exception OR a broadcast-storm starving the main thread.
  // Both are now visible:
  //   1. mount boundaries print "[POPUP] <stage> +<ms-since-mount>"
  //   2. EXPORT_STATUS message rate logs every 2s when > 0
  //   3. The init() body is wrapped in try/catch — any synchronous throw
  //      that would otherwise leave the popup blank now prints to console
  //      AND surfaces a visible error banner.
  const __popupMountStart = performance.now();
  const __popupTrace = (stage: string) => {
    try {
      const dt = (performance.now() - __popupMountStart).toFixed(1);
      console.log(`[POPUP] ${stage} +${dt}ms`);
    } catch { /* console may be unavailable in odd contexts */ }
  };
  let __exportStatusCount = 0;
  let __exportStatusWindowStart = performance.now();
  const __reportExportStatusRate = setInterval(() => {
    if (__exportStatusCount === 0) return;
    const elapsedSec = (performance.now() - __exportStatusWindowStart) / 1000;
    const rate = (__exportStatusCount / elapsedSec).toFixed(1);
    console.log(`[POPUP] EXPORT_STATUS rate ${rate}/s (${__exportStatusCount} in ${elapsedSec.toFixed(1)}s)`);
    __exportStatusCount = 0;
    __exportStatusWindowStart = performance.now();
  }, 2000);
  __popupTrace('script-eval');

  onMount(() => {
    __popupTrace('onMount-enter');
    const init = async () => {
      __popupTrace('init-enter');
      try {
      setBusy(false);
      const loaded = await loadStoredOptions();
      if (!alive) return;
      options = loaded;
      __popupTrace('options-loaded');
      void refreshSavedPresets();
      // Reconcile imageFetchFallback with the live <all_urls>
      // permission state on every popup open. Both Firefox and Chrome
      // can close the popup mid-await on focus-loss when the permission
      // prompt appears, killing our dispatch chain. The user accepts in
      // the prompt, the browser grants the permission, but the popup
      // never persists the option flag. Next open, Settings shows the
      // toggle as off even though permission is granted. Trust the
      // browser's permission state as the source of truth and update
      // the option to match. Also handles the inverse: user revoked
      // <all_urls> from the browser's settings page (chrome://extensions
      // / about:addons) while the popup was closed; the live
      // permission removal listener only catches in-popup-session
      // changes. Failure to call the API leaves the option alone.
      try {
        // @ts-ignore - browser global on Firefox; chrome polyfill on Chrome
        const reconcilePerms = typeof browser !== 'undefined' ? browser.permissions : chrome.permissions;
        const granted = await reconcilePerms.contains({ origins: ['<all_urls>'] });
        if (alive && granted !== options.imageFetchFallback) {
          void updateOption('imageFetchFallback', !!granted);
        }
      } catch { /* permissions API unavailable in this context — leave option alone */ }
      // Restore the picker's collapsed/expanded preference. Read in the
      // same early phase as the language so the very first paint already
      // reflects the user's choice — no flash of "expanded then collapse".
      const savedCollapsed = await readSavedPickerCollapsed();
      if (alive && savedCollapsed !== null) pickerCollapsed = savedCollapsed;
      await applyLanguage(options.lang || "en");
      applyTheme(options.theme || "light");
      updateQuickRangeActive();
      // Restore the last-open page BEFORE onboarding check — both branches
      // inspect showSettings/showHistory. Guard against corrupt storage
      // values by ignoring anything that isn't one of the three pages.
      try {
        const stored = await chrome.storage.local.get(LAST_PAGE_STORAGE_KEY);
        const last = stored?.[LAST_PAGE_STORAGE_KEY];
        if (alive) {
          if (last === 'settings') showSettings = true;
          else if (last === 'history') showHistory = true;
          else if (last === 'diagnostics') showDiagnostics = true;
        }
      } catch { /* fall through to 'main' default */ }
      // Open the persistence gate now that any restored page state
      // is applied. From here on, any user navigation reactively
      // writes the active page to storage.
      if (alive) pageRestored = true;

      // Review-prompt gate: pull the one-shot flag and first-install
      // timestamp. Both reads are best-effort; any failure leaves the
      // prompt permanently hidden for this session.
      try {
        const reviewStored = await chrome.storage.local.get([
          REVIEW_PROMPT_STORAGE_KEY,
          FIRST_INSTALL_STORAGE_KEY,
        ]);
        if (alive) {
          const s = reviewStored?.[REVIEW_PROMPT_STORAGE_KEY];
          if (s && typeof s === 'object' && typeof s.shown === 'boolean') {
            reviewPromptState = s as ReviewPromptState;
          }
          const t = reviewStored?.[FIRST_INSTALL_STORAGE_KEY];
          if (typeof t === 'number' && Number.isFinite(t)) {
            firstInstalledAt = t;
          }
        }
      } catch { /* keep defaults */ }
      // Show the welcome overlay on the first popup open. We deliberately
      // DON'T show it when Settings or History routes are active on open
      // (those routes imply returning users), and we skip it while an
      // export is already running to avoid interrupting the task.
      if (!options.onboardingDismissed && !showSettings && !showHistory && !showDiagnostics) {
        showOnboarding = true;
      }
      const persistedError = await loadPersistedError();
      if (!alive) return;
      if (persistedError?.message) {
        // Only show errors less than 60 seconds old — stale errors from prior sessions are noise
        const age = Date.now() - (persistedError.timestamp || 0);
        if (age < 60_000) {
          showErrorBanner(persistedError.message, false);
          if (!statusText) {
            setStatus(persistedError.message);
          }
        } else {
          void clearLastError(storage);
        }
      }
      try {
        const tab = await getActiveTeamsTab();
        currentTabId = tab.id ?? null;

        // Restore the saved picker kind BEFORE loadConversations fires —
        // the picker reads selectedKind during its first render, and we
        // want that render to reflect the saved choice rather than flash
        // "all" for one frame and then snap to e.g. "Chats". The folder
        // gets reconciled inside loadConversations after the folder list
        // is known; the kind set is fixed so we can restore it eagerly.
        try {
          const savedKind = await readSavedKindChoice();
          if (savedKind) selectedKind = savedKind;
        } catch { /* keep default 'all' */ }

        // Kick off conversation list load in parallel — the picker needs
        // it before an export can be started, but it's independent of the
        // status/history hydration below, so firing here gets the list
        // ready as early as possible. Errors render as a retry button in
        // the picker, not a blocking banner.
        if (currentTabId != null) void loadConversations();

        // Pre-hydrate from chrome.storage.local BEFORE the async message
        // round-trip to the background. The background mirrors its in-
        // memory activeExports Map here on every update, so we can paint
        // the correct busy/idle state almost instantly. GET_EXPORT_STATUS
        // below still runs and overwrites with the authoritative answer
        // (covers the "service worker got killed and lost state" case).
        if (currentTabId != null) {
          try {
            const stored = await chrome.storage.local.get(ACTIVE_EXPORTS_STORAGE_KEY);
            const snapshot = (stored?.[ACTIVE_EXPORTS_STORAGE_KEY] || {}) as Record<string, ActiveExportInfo>;
            const preInfo = snapshot[String(currentTabId)];
            if (alive && preInfo) {
              const startedAt = normalizeStart(preInfo.startedAt);
              if (startedAt) { startedAtMs = startedAt; ensureElapsedTimer(); }
              if (preInfo.lastStatus) {
                handleExportStatus(preInfo.lastStatus);
              } else {
                setBusy(true, busyExportLabel());
                setStatus(t("status.running"));
              }
              // We have a confident pre-state — clear the checking flag
              // so the reveal is instant rather than waiting for the
              // message reply below.
              statusKnown = true;
            }
          } catch { /* storage unavailable — fall through to async */ }
        }

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
        } else if (busy) {
          // Pre-hydration said busy, but authoritative says idle — the
          // export finished while the popup was closed. Drop back to
          // idle and let the history page carry the outcome signal.
          setBusy(false);
          setStatus(t("status.ready"));
          resetPhaseTracker();
          startedAtMs = null;
        }
      } catch {
        /* user not on Teams tab — flag the picker so it renders a
           dedicated "open Teams web first" message instead of falling
           through to the misleading "No matches" empty-state. */
        if (alive) notOnTeamsTab = true;
      } finally {
        // Flip the "checking…" neutral state regardless of outcome so the
        // export button renders its real label (idle or busy) from here on.
        if (alive) statusKnown = true;
      }
      // Always load history so the dot reflects any entries added while
      // the popup was closed.
      await refreshHistory();
      __popupTrace('init-done');
      } catch (e: any) {
        // Surface ANY synchronous throw during init so the popup never
        // ends up blank-and-silent. The error banner is intentionally
        // separate from the structured per-step traces — the trace tells
        // us where init stopped, the banner tells the user something
        // went wrong without making them open DevTools.
        const msg = e?.message || String(e);
        console.error('[POPUP] init failed:', e);
        try { showErrorBanner(`Popup init failed: ${msg}`); } catch { /* noop */ }
        __popupTrace(`init-error: ${msg.slice(0, 80)}`);
      }
    };
    void init();
    runtime.onMessage.addListener(onRuntimeMessage);
    // Keep imageFetchFallback in sync with the actual permission state.
    // Firing on browser-side revocation (chrome://extensions or
    // about:addons) means the toggle reflects reality even if the user
    // revokes outside our UI. We only react to <all_urls> removal —
    // other permission events are unrelated. Use browser.* on Firefox
    // (chrome.* there is callback-only) and chrome.* on Chrome MV3.
    // @ts-ignore
    const permsEvents = typeof browser !== 'undefined' ? browser.permissions : chrome.permissions;
    permsEvents.onRemoved.addListener(onPermissionsRemoved);
    __popupTrace('listener-registered');
  });

  onDestroy(() => {
    alive = false;
    try { clearInterval(__reportExportStatusRate); } catch { /* noop */ }
    runtime.onMessage.removeListener(onRuntimeMessage);
    try {
      // @ts-ignore
      const permsEvents = typeof browser !== 'undefined' ? browser.permissions : chrome.permissions;
      (permsEvents.onRemoved as unknown as { removeListener: (cb: typeof onPermissionsRemoved) => void })
        .removeListener(onPermissionsRemoved);
    } catch { /* best-effort */ }
    clearElapsedTimer();
  });

  function onPermissionsRemoved(perms: chrome.permissions.Permissions) {
    if (!alive) return;
    if (!perms.origins || !perms.origins.includes('<all_urls>')) return;
    if (!options.imageFetchFallback) return;
    void updateOption('imageFetchFallback', false);
  }
</script>

<div class="popup">
  {#if showSettings}
    <div class="popup-content">
      <SettingsPage
        theme={options.theme}
        lang={options.lang || "en"}
        languages={languageOptions}
        afterExport={options.afterExport}
        avatarMode={options.avatarMode}
        embedAvatars={options.embedAvatars}
        pdfPageSize={options.pdfPageSize}
        pdfBodyFontSize={options.pdfBodyFontSize}
        pdfShowPageNumbers={options.pdfShowPageNumbers}
        pdfIncludeAvatars={options.pdfIncludeAvatars}
        imageFetchFallback={options.imageFetchFallback}
        fullResImages={options.fullResImages}
        imageFilenameDate={options.imageFilenameDate}
        imageModifiedDate={options.imageModifiedDate}
        on:back={() => (showSettings = false)}
        on:themeChange={(e) => updateOption("theme", e.detail)}
        on:langChange={(e) => updateOption("lang", e.detail)}
        on:afterExportChange={(e) => updateOption("afterExport", e.detail)}
        on:avatarModeChange={(e) => updateOption("avatarMode", e.detail)}
        on:pdfPageSizeChange={(e) => updateOption("pdfPageSize", e.detail)}
        on:pdfBodyFontSizeChange={(e) => updateOption("pdfBodyFontSize", e.detail)}
        on:pdfShowPageNumbersChange={(e) => updateOption("pdfShowPageNumbers", e.detail)}
        on:pdfIncludeAvatarsChange={(e) => updateOption("pdfIncludeAvatars", e.detail)}
        on:imageFetchFallbackChange={(e) => updateOption("imageFetchFallback", e.detail)}
        on:fullResImagesChange={(e) => updateOption("fullResImages", e.detail)}
        on:imageFilenameDateChange={(e) => updateOption("imageFilenameDate", e.detail)}
        on:imageModifiedDateChange={(e) => updateOption("imageModifiedDate", e.detail)}
        on:replayTour={replayTour}
        on:openDiagnostics={() => { showSettings = false; showDiagnostics = true; }}
      />
    </div>
  {:else if showDiagnostics}
    <div class="popup-content">
      <DiagnosticsPage
        lang={options.lang || "en"}
        on:back={() => { showDiagnostics = false; showSettings = true; }}
      />
    </div>
  {:else if showHistory}
    <div class="popup-content">
      <HistoryPage
        entries={historyEntries}
        lang={options.lang || "en"}
        on:back={() => (showHistory = false)}
        on:remove={(e) => onHistoryRemove(e.detail)}
        on:clearAll={onHistoryClear}
        on:markMissing={(e) => onHistoryMarkMissing(e.detail)}
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
          {newHistoryCount}
          {pulseHistoryIcon}
          lang={options.lang || "en"}
          on:openSettings={() => (showSettings = true)}
          on:openHistory={() => { showHistory = true; void onHistoryOpened(); }}
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

      <!-- Conversation picker. User selects which chat/channel to export.
           Replaces the old "guess from DOM" flow end-to-end — the ID the
           user picks here is passed straight through to the scraper. -->
      <ConversationPicker
        lang={options.lang || "en"}
        {conversations}
        {folders}
        bind:selectedIds={selectedConversationIds}
        bind:selectedFolderId
        bind:selectedKind
        bind:collapsed={pickerCollapsed}
        mode="multi"
        state={pickerState}
        errorMessage={pickerError}
        refreshing={isRefreshingList}
        notOnTeams={notOnTeamsTab}
        on:retry={() => loadConversations({ forceRefresh: true })}
        on:folderChange={(e) => persistFolderChoice(e.detail)}
        on:kindChange={(e) => persistKindChoice(e.detail)}
        on:collapseChange={(e) => persistPickerCollapsed(e.detail)}
        {savedPresets}
        on:savePreset={(e) => onSavePreset(e.detail)}
        on:removePreset={(e) => onRemovePreset(e.detail)}
      />

      <!-- Export button — plain in every state. Outcomes live in History page. -->
      <ExportButton
        disabled={selectedConversationIds.length === 0}
        selectionCount={selectedConversationIds.length}
        selectionIsOther={singleSelectionIsOther}
        {busy}
        {statusKnown}
        summary={exportSummary}
        {phaseLabel}
        {counterValue}
        {counterLabel}
        {segments}
        flashTrigger={successFlashTrigger}
        lang={options.lang || "en"}
        on:run={startExport}
        on:stop={stopExport}
      />

      {#if reviewPromptEligible}
        <ReviewPrompt
          lang={options.lang || "en"}
          storeUrl={reviewStoreUrl}
          issueUrl={reviewIssueUrl}
          on:respond={(e) => onReviewPromptRespond(e.detail)}
        />
      {/if}

      <TargetSection
        target={options.exportTarget}
        lang={options.lang || "en"}
        on:targetChange={(e) => updateOption("exportTarget", e.detail)}
      />

      <!-- Format Section (Full Width) -->
      <FormatSection
        formats={options.formats}
        downloadImages={options.downloadImages}
        embedAvatars={options.embedAvatars}
        avatarMode={options.avatarMode}
        lang={options.lang || "en"}
        on:formatsChange={(e) => updateOption("formats", e.detail)}
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
          downloadFiles={options.downloadFiles}
          lang={options.lang || "en"}
          disableReplies={options.formats.every((f) => f === "txt")}
          disableReactions={options.formats.every((f) => f === "txt")}
          disableAvatars={options.formats.every((f) => f === "txt" || f === "csv")}
          disableImages={!options.formats.includes("html") && !options.formats.includes("pdf")}
          on:includeRepliesChange={(e) =>
            updateOption("includeReplies", e.detail)}
          on:includeReactionsChange={(e) =>
            updateOption("includeReactions", e.detail)}
          on:includeSystemChange={(e) =>
            updateOption("includeSystem", e.detail)}
          on:embedAvatarsChange={(e) => updateOption("embedAvatars", e.detail)}
          on:includeImagesChange={(e) =>
            updateOption("downloadImages", e.detail)}
          on:downloadFilesChange={(e) =>
            updateOption("downloadFiles", e.detail)}
        />
      </div>
    </div>
  {/if}

  {#if showOnboarding}
    <!-- Tour temporarily expands the picker for the 'folder' step;
         two-way bind keeps App's pickerCollapsed in sync, and the
         tour restores the original value on dismiss. Tour-driven
         changes deliberately don't hit chrome.storage — only the
         picker's own toggle click calls persistPickerCollapsed, so
         an open + restore cycle leaves the persisted default alone. -->
    <OnboardingOverlay
      lang={options.lang || "en"}
      autoStart={tourAutoStart}
      bind:pickerCollapsed
      on:dismiss={dismissOnboarding}
    />
  {/if}
</div>
