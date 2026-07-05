<script lang="ts" module>
  // Firefox polyfill global
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  declare const browser: any;
</script>

<script lang="ts">
  import { createEventDispatcher, onMount } from 'svelte';
  import { ArrowLeft, Copy, Download, Eye, EyeOff, RefreshCw, Trash2 } from 'lucide-svelte';
  import { t } from '../../../i18n/i18n';
  import { DEFAULT_OPTIONS, loadHistory, loadOptions, saveOptions } from '../../../utils/options';
  import { isTeamsUrl } from '../../../utils/teams-urls';
  import { runtimeSend } from '../../../utils/messaging';
  import {
    buildDiagnosticJson,
    type DiagnosticReportInput,
    type EnvInfo,
    type PermissionsInfo,
    type IdbShape,
    type ExportsSection,
    type OptionsSection,
    type LogTail,
    type DiagLogEntry,
    type ProbeResult,
    type ProbesSection,
  } from '../../../utils/diagnostics';

  export let lang = 'en';

  const dispatch = createEventDispatcher<{ back: void }>();

  const runtime = typeof browser !== 'undefined' ? browser.runtime : chrome.runtime;
  const tabs = typeof browser !== 'undefined' ? browser.tabs : chrome.tabs;
  const permissions = typeof browser !== 'undefined' ? browser.permissions : chrome.permissions;
  const storage = typeof browser !== 'undefined' ? browser.storage : chrome.storage;

  type LoadState = 'idle' | 'loading' | 'ready' | 'error';
  let state: LoadState = 'idle';
  let errorMessage = '';
  let toggleError = '';
  let saveError = '';
  let probesError = '';
  let reportJson = '';
  let includeRawIds = false;
  let showPreview = false;
  let copyConfirmTimer: ReturnType<typeof setTimeout> | null = null;
  let saveConfirmTimer: ReturnType<typeof setTimeout> | null = null;
  let copyConfirmed = false;
  let saveConfirmed = false;
  let lastInput: DiagnosticReportInput | null = null;
  // Layer 2 probes: state moves from not-run -> done (or failed).
  // The probe section renders rows once a run completes.
  let probesRunning = false;
  let probes: ProbesSection = { state: 'not-run' };
  // File-access probe: paste a SharePoint file link (e.g. from FAILURES.txt),
  // resolve it via the shares API in the Teams tab's content script, then hand
  // the pre-authenticated downloadUrl to the background for a real download.
  // Row labels/values are raw English like the probe details above.
  let fileProbeUrl = '';
  let fileProbeRunning = false;
  let fileProbeRows: { label: string; value: string; warn?: boolean }[] = [];
  let fileProbeError = '';
  // Salvage: import a list of file links (e.g. a FAILURES.txt or the URL list),
  // resolve + download each one serially into Downloads/TCE-salvage/.
  let salvageRunning = false;
  let salvageStatus = '';
  let salvageRows: { label: string; value: string; warn?: boolean }[] = [];
  let salvageError = '';
  // Raw file-field dump: shows which fields the open chat's file attachments
  // carry (to check for a sharing-link field). Field NAMES only.
  let fieldDumpRunning = false;
  let fieldDumpRows: { label: string; value: string; warn?: boolean }[] = [];
  let fieldDumpError = '';
  // Persistence state mirrored from the loaded options. Toggling
  // updates both the storage option AND tells BG to flip its runtime
  // flag without waiting for an SW restart.
  let persistEnabled = false;
  let persistBytesUsed: number | null = null;
  let persistFlushError: { ts: number; reason: string } | null = null;
  let persistBusy = false;
  let clearBusy = false;
  // Verbose export-stats state, mirrored from the loaded option. Off keeps the
  // [export-stats] console line ID-free; on adds per-chat detail for local
  // perf debugging.
  let verboseStatsEnabled = false;
  let verboseStatsBusy = false;
  let summary: {
    env: EnvInfo;
    idb: IdbShape;
    exports: ExportsSection;
    logsMissing: string | null;
  } | null = null;

  onMount(() => {
    void refresh();
    // Verbose-stats reflects an option only (no BG round-trip needed).
    void loadOptions(storage, DEFAULT_OPTIONS).then((o) => { verboseStatsEnabled = !!o.diagVerboseStats; });
  });

  async function refresh() {
    state = 'loading';
    errorMessage = '';
    toggleError = '';
    saveError = '';
    try {
      const input = await collectDiagnosticInput();
      lastInput = input;
      summary = {
        env: input.env,
        idb: input.idb,
        exports: input.exports,
        logsMissing: input.logs.missing ?? null,
      };
      reportJson = await buildDiagnosticJson(input, { includeRawIds });
      state = 'ready';
    } catch (e) {
      errorMessage = e instanceof Error ? e.message : String(e);
      state = 'error';
    }
  }

  async function collectDiagnosticInput(): Promise<DiagnosticReportInput> {
    const activeTab = await getActiveTab();
    const env = await collectEnv(activeTab);
    const perms = await collectPermissions();
    const options = await collectOptionsSnapshot();
    const exports = await collectExportsSummary();
    const bgResp = await fetchBackgroundLogs();
    const contentResp = await fetchContentDiagnostics(activeTab);

    // Mirror BG's flags into local state so the UI matches what BG
    // actually thinks (in case the option storage and BG runtime
    // drift; the SET_DIAG_LOG_PERSIST handler keeps them in sync).
    if (bgResp.ok) {
      persistEnabled = bgResp.persistEnabled;
      persistBytesUsed = bgResp.persistEnabled ? bgResp.bytesUsed : null;
      persistFlushError = bgResp.persistEnabled ? bgResp.lastFlushError : null;
    }

    const logs: LogTail = {
      entries: bgResp.ok ? bgResp.entries : null,
      missing: bgResp.ok ? undefined : bgResp.reason,
      bytesUsed: bgResp.ok ? bgResp.bytesUsed : null,
      persistEnabled: bgResp.ok ? bgResp.persistEnabled : false,
      lastFlushError: bgResp.ok ? bgResp.lastFlushError : null,
      forwarding: contentResp.ok ? (contentResp.forwardingStats ?? undefined) : undefined,
    };
    const idb: IdbShape = contentResp.ok
      ? contentResp.idbShape
      : { available: false, reason: contentResp.reason };

    return { env, permissions: perms, options, idb, exports, logs, probes };
  }

  // Single source of truth for "what tab is the user on". Avoids two
  // tabs.query calls per refresh and keeps env + content paths in sync
  // on what they think the active tab is.
  type ActiveTabInfo = { id?: number; url?: string };
  async function getActiveTab(): Promise<ActiveTabInfo | null> {
    try {
      const list = await tabs.query({ active: true, currentWindow: true });
      const t = Array.isArray(list) ? list[0] : null;
      if (!t) return null;
      return { id: t.id ?? undefined, url: t.url ?? undefined };
    } catch {
      return null;
    }
  }

  async function collectEnv(activeTab: ActiveTabInfo | null): Promise<EnvInfo> {
    const manifest = runtime.getManifest?.() ?? { version: '', manifest_version: 3 };
    const ua = (navigator as { userAgentData?: { brands?: { brand: string; version: string }[]; platform?: string } }).userAgentData;
    const brandFromUACH = ua?.brands?.find(b => !/Brand|Not\W?A.Brand/i.test(b.brand));
    const browserBrand = brandFromUACH?.brand
      ?? (/Firefox\//.test(navigator.userAgent) ? 'Firefox' : 'Unknown');
    const browserVersion = brandFromUACH?.version
      ?? (navigator.userAgent.match(/Firefox\/([0-9.]+)/)?.[1] || '');
    const os = ua?.platform || (() => {
      const m = navigator.userAgent.match(/\(([^)]+)\)/);
      return m?.[1] ?? navigator.platform ?? 'unknown';
    })();
    const colorScheme = matchMedia('(prefers-color-scheme: dark)').matches ? 'dark'
      : matchMedia('(prefers-color-scheme: light)').matches ? 'light' : 'unknown';

    // Only capture the tab URL when it's clearly a Teams web tab. A
    // user running Diagnostics on a private internal app should not
    // ship that URL into the report even with redaction on; the
    // redactor only catches well-known identifier shapes inside the
    // string, not arbitrary hostnames or paths.
    let teamsTabUrl: string | null = null;
    let teamsHost: string | null = null;
    if (activeTab?.url && isTeamsUrl(activeTab.url)) {
      teamsTabUrl = activeTab.url;
      try { teamsHost = new URL(teamsTabUrl).host; } catch { teamsHost = null; }
    }

    return {
      extensionVersion: String(manifest.version || ''),
      buildStamp: __BUILD_STAMP__,
      manifestVersion: Number(manifest.manifest_version ?? 3),
      browserBrand,
      browserVersion,
      os,
      teamsTabUrl,
      teamsHost,
      locale: navigator.language || '',
      documentLang: document.documentElement?.lang || null,
      prefersColorScheme: colorScheme,
    };
  }

  async function collectPermissions(): Promise<PermissionsInfo> {
    const manifest: any = runtime.getManifest?.() ?? {};
    const declared: string[] = Array.isArray(manifest.host_permissions)
      ? manifest.host_permissions
      : Array.isArray(manifest.permissions)
        ? manifest.permissions.filter((p: string) => /:\/\//.test(p))
        : [];
    // Three states: true (granted), false (declined), null (couldn't
    // determine). Conflating null with false made the report tell the
    // analyst "tell the user to grant the permission" when the user
    // already had.
    let allUrlsGranted: boolean | null;
    try {
      allUrlsGranted = await permissions.contains({ origins: ['<all_urls>'] });
    } catch {
      allUrlsGranted = null;
    }
    return { declaredHostPermissions: declared, optionalAllUrlsGranted: allUrlsGranted };
  }

  async function collectOptionsSnapshot(): Promise<OptionsSection> {
    try {
      const opts = await loadOptions(storage);
      return { available: true, values: opts as unknown as Record<string, unknown> };
    } catch (e) {
      return { available: false, reason: e instanceof Error ? e.message : String(e) };
    }
  }

  async function collectExportsSummary(): Promise<ExportsSection> {
    try {
      const all = await loadHistory(storage);
      const items = all.slice(0, 5).map(e => ({
        savedAt: e.savedAt,
        kind: e.kind,
        partialReason: e.partialReason,
        formats: e.formats,
        messageCount: e.messageCount,
        elapsedMs: e.elapsedMs,
        isZip: e.isZip,
      }));
      return { available: true, items };
    } catch (e) {
      return { available: false, reason: e instanceof Error ? e.message : String(e) };
    }
  }

  async function fetchBackgroundLogs(): Promise<
    | {
        ok: true;
        entries: DiagLogEntry[];
        bytesUsed: number | null;
        persistEnabled: boolean;
        lastFlushError: { ts: number; reason: string } | null;
      }
    | { ok: false; reason: string }
  > {
    try {
      const resp = await runtimeSend(runtime, { type: 'GET_DIAGNOSTICS_BG' });
      if (!resp || !Array.isArray(resp.entries)) {
        return { ok: false, reason: 'background returned malformed payload' };
      }
      return {
        ok: true,
        entries: resp.entries,
        bytesUsed: typeof resp.bytesUsed === 'number' || resp.bytesUsed === null ? resp.bytesUsed : null,
        persistEnabled: !!resp.persistEnabled,
        lastFlushError: resp.lastFlushError ?? null,
      };
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { ok: false, reason: msg };
    }
  }

  async function fetchContentDiagnostics(activeTab: ActiveTabInfo | null): Promise<
    | { ok: true; idbShape: IdbShape; forwardingStats: { lostBatches: number; lostEntries: number; lastError: string | null } | null }
    | { ok: false; reason: string }
  > {
    if (!activeTab?.id) return { ok: false, reason: 'no active tab' };
    if (!isTeamsUrl(activeTab.url ?? '')) return { ok: false, reason: 'active tab is not a Teams URL' };

    try {
      // Both Chrome MV3 and Firefox return a Promise from tabs.sendMessage
      // (since Chrome 99); errors come through as rejections, including
      // "Could not establish connection. Receiving end does not exist."
      // when no content script is injected.
      const resp = await tabs.sendMessage(activeTab.id, { type: 'GET_DIAGNOSTICS_CONTENT' });
      if (!resp || typeof resp !== 'object') return { ok: false, reason: 'content script returned no payload' };
      const r = resp as { idbShape?: unknown; forwardingStats?: unknown };
      const idbShape = (r.idbShape as IdbShape | undefined) ?? { available: false, reason: 'missing idbShape in response' };
      const fs = r.forwardingStats as { lostBatches?: number; lostEntries?: number; lastError?: string | null } | undefined;
      const forwardingStats = fs && typeof fs.lostBatches === 'number'
        ? { lostBatches: fs.lostBatches, lostEntries: fs.lostEntries ?? 0, lastError: fs.lastError ?? null }
        : null;
      return { ok: true, idbShape, forwardingStats };
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e);
      return { ok: false, reason: msg };
    }
  }

  async function fetchBgProbes(): Promise<
    | { ok: true; results: ProbeResult[]; totalMs: number }
    | { ok: false; reason: string }
  > {
    try {
      const resp = await runtimeSend(runtime, { type: 'RUN_PROBES_BG' });
      if (!resp || typeof resp !== 'object' || !('ok' in resp)) {
        return { ok: false, reason: 'malformed RUN_PROBES_BG response' };
      }
      if (resp.ok) {
        return { ok: true, results: resp.results, totalMs: resp.totalMs };
      }
      return { ok: false, reason: resp.reason };
    } catch (e) {
      return { ok: false, reason: e instanceof Error ? e.message : String(e) };
    }
  }

  async function fetchContentProbes(activeTab: ActiveTabInfo | null): Promise<
    | { ok: true; results: ProbeResult[]; totalMs: number }
    | { ok: false; reason: string }
  > {
    if (!activeTab?.id) return { ok: false, reason: 'no active tab' };
    if (!isTeamsUrl(activeTab.url ?? '')) return { ok: false, reason: 'active tab is not a Teams URL' };
    try {
      const resp = await tabs.sendMessage(activeTab.id, { type: 'RUN_PROBES_CONTENT' });
      if (!resp || typeof resp !== 'object') {
        return { ok: false, reason: 'content script returned no payload' };
      }
      const r = resp as { ok?: boolean; results?: ProbeResult[]; totalMs?: number; reason?: string };
      if (r.ok && Array.isArray(r.results) && typeof r.totalMs === 'number') {
        return { ok: true, results: r.results, totalMs: r.totalMs };
      }
      return { ok: false, reason: r.reason || 'probe runner returned no results' };
    } catch (e) {
      return { ok: false, reason: e instanceof Error ? e.message : String(e) };
    }
  }

  // Stable order for synthesizing skipped rows when the content side
  // is unreachable. Keeps the checklist visually complete instead of
  // silently shrinking — analysts seeing fewer rows in a bad-state
  // run would otherwise have no way to tell which probes ran.
  const KNOWN_CONTENT_PROBES: string[] = [
    'teams_origin_recognized',
    'chat_surface_detected',
    'idb_accessible',
    'skype_token_extractable',
    'ic3_token_extractable',
    'asyncgw_reachable',
    'asm_skype_reachable',
    'page_world_helper',
    'canary_image_fetch',
  ];

  async function runFieldDump() {
    if (fieldDumpRunning) return;
    fieldDumpRunning = true;
    fieldDumpError = '';
    fieldDumpRows = [];
    try {
      const activeTab = await getActiveTab();
      if (!activeTab?.id || !isTeamsUrl(activeTab.url ?? '')) {
        fieldDumpError = t('diagnostics.fileProbeNoTab', {}, lang);
        return;
      }
      const r = (await tabs.sendMessage(activeTab.id, { type: 'DUMP_FILE_FIELDS' })) as
        | { ok?: boolean; messages?: number; fileRecords?: number; keys?: string[]; linkFields?: string[]; error?: string }
        | undefined;
      if (!r?.ok) {
        fieldDumpError = r?.error || 'no response';
        return;
      }
      fieldDumpRows = [
        { label: 'messages scanned', value: String(r.messages ?? 0) },
        { label: 'file records', value: String(r.fileRecords ?? 0), warn: (r.fileRecords ?? 0) === 0 },
        { label: 'link-ish fields', value: (r.linkFields && r.linkFields.length ? r.linkFields.join(', ') : 'none'), warn: !(r.linkFields && r.linkFields.length) },
        { label: 'all fields', value: (r.keys && r.keys.length ? r.keys.join(', ') : '-') },
      ];
    } catch (e) {
      fieldDumpError = e instanceof Error ? e.message : String(e);
    } finally {
      fieldDumpRunning = false;
    }
  }

  async function runFileProbe() {
    const href = fileProbeUrl.trim();
    if (!href || fileProbeRunning) return;
    fileProbeRunning = true;
    fileProbeError = '';
    fileProbeRows = [];
    try {
      const activeTab = await getActiveTab();
      if (!activeTab?.id || !isTeamsUrl(activeTab.url ?? '')) {
        fileProbeError = t('diagnostics.fileProbeNoTab', {}, lang);
        return;
      }
      const resolve = (await tabs.sendMessage(activeTab.id, { type: 'RESOLVE_SHARE_FILE', href })) as
        | { ok?: boolean; status?: number; name?: string; mimeType?: string; downloadUrl?: string; blocksDownload?: boolean; error?: string; via?: string; matchedName?: string; mode?: string }
        | undefined;
      const r = resolve ?? {};
      const rows: { label: string; value: string; warn?: boolean }[] = [
        { label: 'via', value: r.via === 'shareUrl' ? `sharing link (${r.matchedName || 'matched'})` : 'raw URL (no chat match)', warn: r.via !== 'shareUrl' },
        { label: 'resolve', value: `${r.ok ? `ok (HTTP ${r.status})` : `failed (HTTP ${r.status || 0}${r.error ? `: ${r.error}` : ''})`}${r.mode ? ` [creds: ${r.mode}]` : ''}`, warn: !r.ok },
        { label: 'token', value: r.error && r.error.includes('no SharePoint token') ? 'not found in MSAL cache' : 'found', warn: !!(r.error && r.error.includes('no SharePoint token')) },
      ];
      if (r.ok) {
        rows.push({ label: 'name', value: r.name || '-' });
        rows.push({ label: 'mimeType', value: r.mimeType || '-' });
        rows.push({ label: 'blocksDownload', value: r.blocksDownload == null ? 'unknown' : String(r.blocksDownload), warn: r.blocksDownload === true });
        // Presence only — the URL embeds a short-lived token, never shown.
        rows.push({ label: 'downloadUrl', value: r.downloadUrl ? 'present (short-lived)' : 'absent', warn: !r.downloadUrl });
      }
      fileProbeRows = rows;
      if (r.ok && r.downloadUrl) {
        const dl = await runtimeSend(runtime, { type: 'PROBE_FILE_DOWNLOAD', url: r.downloadUrl, name: r.name });
        fileProbeRows = [
          ...rows,
          dl && 'ok' in dl && dl.ok
            ? { label: 'download', value: `${dl.outcome} (${dl.mime || '?'}, ${dl.bytes ?? '?'} bytes -> Downloads/TCE-probe/${dl.filename || ''})`, warn: dl.outcome !== 'complete' }
            : { label: 'download', value: `failed: ${dl && 'error' in dl ? dl.error : 'no response'}`, warn: true },
        ];
      }
    } catch (e) {
      fileProbeError = e instanceof Error ? e.message : String(e);
    } finally {
      fileProbeRunning = false;
    }
  }

  // Salvage: import a file of links, resolve + download each serially.
  async function runSalvage(ev: Event) {
    const input = ev.target as HTMLInputElement;
    const file = input.files?.[0];
    if (!file || salvageRunning) return;
    salvageRunning = true;
    salvageError = '';
    salvageRows = [];
    salvageStatus = '';
    try {
      // Read the file, then hand the links to the BACKGROUND to do the whole
      // resolve+download job. A file picker closes the popup (killing this
      // context), so the download loop must not live here — the background
      // survives and logs progress to the service worker console.
      const text = await file.text();
      const hrefs = Array.from(new Set(text.match(/https?:\/\/\S+/g) || []));
      if (!hrefs.length) { salvageError = t('diagnostics.salvageNoLinks', {}, lang); return; }
      const activeTab = await getActiveTab();
      if (!activeTab?.id || !isTeamsUrl(activeTab.url ?? '')) { salvageError = t('diagnostics.fileProbeNoTab', {}, lang); return; }
      await runtimeSend(runtime, { type: 'SALVAGE_LINKS', hrefs, tabId: activeTab.id });
      salvageRows = [{ label: 'links', value: String(hrefs.length) }];
      salvageStatus = t('diagnostics.salvageStarted', { n: hrefs.length }, lang);
    } catch (e) {
      salvageError = e instanceof Error ? e.message : String(e);
    } finally {
      salvageRunning = false;
      input.value = ''; // allow re-importing the same file
    }
  }

  async function runProbes() {
    if (probesRunning) return;
    probesRunning = true;
    probesError = '';
    try {
      const runAt = Date.now();
      const activeTab = await getActiveTab();
      const [bg, content] = await Promise.all([
        fetchBgProbes(),
        fetchContentProbes(activeTab),
      ]);
      const results: ProbeResult[] = [];
      let totalMs = 0;
      if (bg.ok) {
        results.push(...bg.results);
        totalMs = Math.max(totalMs, bg.totalMs);
      } else {
        // Surface the BG failure as a probe row so it shows up in the
        // checklist. Better than disappearing silently.
        results.push({ name: 'service_worker_alive', status: 'fail', detail: bg.reason, ms: 0 });
      }
      if (content.ok) {
        results.push(...content.results);
        totalMs = Math.max(totalMs, content.totalMs);
      } else {
        results.push({ name: 'content_script_reachable', status: 'fail', detail: content.reason, ms: 0 });
        // Emit a skipped row per content probe so the analyst sees
        // explicitly that those checks did not run, rather than
        // wondering whether they were omitted by design.
        for (const name of KNOWN_CONTENT_PROBES) {
          results.push({ name, status: 'skipped', detail: `not run: ${content.reason}`, ms: 0 });
        }
      }
      probes = { state: 'done', results, runAt, totalMs };
      // Fold probes into the cached input so subsequent toggles /
      // copies pick up the new data.
      if (lastInput) {
        lastInput = { ...lastInput, probes };
        reportJson = await buildDiagnosticJson(lastInput, { includeRawIds });
      }
    } catch (e) {
      probesError = e instanceof Error ? e.message : String(e);
    } finally {
      probesRunning = false;
    }
  }

  async function onCopy() {
    try {
      await navigator.clipboard.writeText(reportJson);
      copyConfirmed = true;
      if (copyConfirmTimer) clearTimeout(copyConfirmTimer);
      copyConfirmTimer = setTimeout(() => { copyConfirmed = false; }, 1800);
    } catch (e) {
      errorMessage = `Copy failed: ${e instanceof Error ? e.message : String(e)}`;
    }
  }

  async function onTogglePersist(event: Event) {
    if (persistBusy) return;
    const desired = (event.target as HTMLInputElement).checked;
    persistBusy = true;
    try {
      // Stash the previous state so we can fully revert (checkbox AND
      // option storage) if BG rejects or the message fails.
      const previousEnabled = persistEnabled;
      const opts = await loadOptions(storage, DEFAULT_OPTIONS);
      await saveOptions(storage, { ...opts, diagLogPersist: desired }, DEFAULT_OPTIONS);
      try {
        const resp = await runtimeSend(runtime, { type: 'SET_DIAG_LOG_PERSIST', enabled: desired });
        if (resp?.ok) {
          persistEnabled = desired;
          if (!desired) persistBytesUsed = null;
          else {
            const bg = await fetchBackgroundLogs();
            if (bg.ok) persistBytesUsed = bg.bytesUsed;
          }
          if (lastInput) {
            lastInput = {
              ...lastInput,
              logs: { ...lastInput.logs, persistEnabled: desired, bytesUsed: desired ? persistBytesUsed : null },
            };
            reportJson = await buildDiagnosticJson(lastInput, { includeRawIds });
          }
        } else {
          // BG handler returned !ok. Roll both the option storage and
          // the visible checkbox back to the previous state so they
          // don't drift apart at the next SW boot.
          await saveOptions(storage, { ...opts, diagLogPersist: previousEnabled }, DEFAULT_OPTIONS);
          (event.target as HTMLInputElement).checked = previousEnabled;
        }
      } catch (innerErr) {
        // Message-level failure (SW unreachable, channel dropped).
        // Same revert as the !ok branch.
        await saveOptions(storage, { ...opts, diagLogPersist: previousEnabled }, DEFAULT_OPTIONS);
        (event.target as HTMLInputElement).checked = previousEnabled;
        throw innerErr;
      }
    } catch (e) {
      errorMessage = e instanceof Error ? e.message : String(e);
    } finally {
      persistBusy = false;
    }
  }

  async function onToggleVerbose(event: Event) {
    if (verboseStatsBusy) return;
    const desired = (event.target as HTMLInputElement).checked;
    verboseStatsBusy = true;
    try {
      const previousEnabled = verboseStatsEnabled;
      const opts = await loadOptions(storage, DEFAULT_OPTIONS);
      await saveOptions(storage, { ...opts, diagVerboseStats: desired }, DEFAULT_OPTIONS);
      try {
        const resp = await runtimeSend(runtime, { type: 'SET_DIAG_VERBOSE_STATS', enabled: desired });
        if (resp?.ok) {
          verboseStatsEnabled = desired;
        } else {
          // BG rejected: revert both option storage and the checkbox so they
          // don't drift apart at the next SW boot.
          await saveOptions(storage, { ...opts, diagVerboseStats: previousEnabled }, DEFAULT_OPTIONS);
          (event.target as HTMLInputElement).checked = previousEnabled;
        }
      } catch (innerErr) {
        await saveOptions(storage, { ...opts, diagVerboseStats: previousEnabled }, DEFAULT_OPTIONS);
        (event.target as HTMLInputElement).checked = previousEnabled;
        throw innerErr;
      }
    } catch (e) {
      errorMessage = e instanceof Error ? e.message : String(e);
    } finally {
      verboseStatsBusy = false;
    }
  }

  async function onClearLogs() {
    if (clearBusy) return;
    clearBusy = true;
    try {
      await runtimeSend(runtime, { type: 'CLEAR_DIAGNOSTICS_LOGS' });
      // Reset the in-page probe results too. Otherwise the "Clear logs"
      // button looks dishonest: logs are wiped but the probe section
      // still shows results from before. The Copy / Save text built
      // afterwards would also still carry those probe results.
      probes = { state: 'not-run' };
      // Refresh to reflect cleared state.
      await refresh();
    } catch (e) {
      errorMessage = e instanceof Error ? e.message : String(e);
    } finally {
      clearBusy = false;
    }
  }

  function formatBytes(b: number | null): string {
    if (b === null || !Number.isFinite(b)) return '?';
    if (b < 1024) return `${b} B`;
    if (b < 1024 * 1024) return `${(b / 1024).toFixed(1)} KB`;
    return `${(b / (1024 * 1024)).toFixed(2)} MB`;
  }

  // Shared helper for both Save-text and Save-JSON. The latter
  // builds the JSON on demand from lastInput because we don't
  // keep a parallel "json text" cache.
  function downloadAs(filename: string, content: string, mime: string) {
    const blob = new Blob([content], { type: mime });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.style.display = 'none';
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      try { document.body.removeChild(a); } catch { /* noop */ }
      URL.revokeObjectURL(url);
    }, 250);
  }

  function stamp(): string {
    return new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
  }

  // Save writes the same JSON the user sees in the Preview and the
  // Copy button puts on the clipboard. One serialisation, three
  // delivery paths (preview / clipboard / file). Avoids the format
  // drift that would happen if Save and Copy used different shapes.
  function onSave() {
    saveError = '';
    if (!reportJson) {
      saveError = 'No report available yet';
      return;
    }
    try {
      downloadAs(`teams-exporter-diagnostic-${stamp()}.json`, reportJson, 'application/json;charset=utf-8');
      saveConfirmed = true;
      if (saveConfirmTimer) clearTimeout(saveConfirmTimer);
      saveConfirmTimer = setTimeout(() => { saveConfirmed = false; }, 1800);
    } catch (e) {
      saveError = `Save failed: ${e instanceof Error ? e.message : String(e)}`;
    }
  }

  async function onRawIdsToggle(event: Event) {
    const requested = (event.target as HTMLInputElement).checked;
    if (state !== 'ready' || !lastInput) {
      // Page isn't ready yet. Accept the new toggle state so the
      // next refresh uses it, but no rebuild.
      includeRawIds = requested;
      return;
    }
    // Rebuild only the text from the cached input. Re-running the
    // full data collection on every toggle would re-probe IDB, which
    // is the slow part. Nothing in DiagnosticReportInput depends on
    // includeRawIds, so the cache is safe to reuse.
    toggleError = '';
    try {
      const text = await buildDiagnosticJson(lastInput, { includeRawIds: requested });
      // Commit both the toggle state AND the new text together, so
      // the checkbox can never drift out of sync with the report
      // body being displayed.
      reportJson = text;
      includeRawIds = requested;
    } catch (e) {
      // Rebuild failed. Revert the checkbox so it matches the
      // report still on screen. Surface the error inline; keep
      // the prior snapshot intact.
      toggleError = e instanceof Error ? e.message : String(e);
      (event.target as HTMLInputElement).checked = includeRawIds;
    }
  }
</script>

<div class="diag-page">
  <div class="settings-header">
    <button class="icon-btn" title={t('common.back', {}, lang)} on:click={() => dispatch('back')}>
      <ArrowLeft size={18} />
    </button>
    <h1>{t('diagnostics.title', {}, lang)}</h1>
    <button class="icon-btn settings-header-right" title={t('common.refresh', {}, lang)} on:click={refresh} disabled={state === 'loading' || probesRunning}>
      <RefreshCw size={18} class={state === 'loading' ? 'spin' : ''} />
    </button>
  </div>

  <div class="subtitle">
    {t('diagnostics.subtitle', {}, lang)}
  </div>

  {#if state === 'loading'}
    <div class="status">{t('diagnostics.collecting', {}, lang)}</div>
  {:else if state === 'error'}
    <div class="status err">{t('diagnostics.buildFailed', {}, lang)} {errorMessage}</div>
  {:else if state === 'ready' && summary}
    <div class="section-title">{t('diagnostics.snapshot', {}, lang)}</div>
    <div class="card stack">
      <div class="row">
        <div class="label">{t('diagnostics.field.extension', {}, lang)}</div>
        <div class="val">{summary.env.extensionVersion} (manifest v{summary.env.manifestVersion})</div>
      </div>
      <div class="row">
        <div class="label">Build</div>
        <div class="val">{summary.env.buildStamp}</div>
      </div>
      <div class="row">
        <div class="label">{t('diagnostics.field.browser', {}, lang)}</div>
        <div class="val">{summary.env.browserBrand} {summary.env.browserVersion}</div>
      </div>
      <div class="row">
        <div class="label">{t('diagnostics.field.os', {}, lang)}</div>
        <div class="val">{summary.env.os}</div>
      </div>
      <div class="row">
        <div class="label">{t('diagnostics.field.teamsHost', {}, lang)}</div>
        <div class="val">{summary.env.teamsHost ?? t('diagnostics.env.noTab', {}, lang)}</div>
      </div>
      <div class="row">
        <div class="label">{t('diagnostics.field.locale', {}, lang)}</div>
        <div class="val">{summary.env.locale}{summary.env.documentLang ? ` · ${t('diagnostics.env.docLang', { lang: summary.env.documentLang }, lang)}` : ''}</div>
      </div>
      <div class="row">
        <div class="label">{t('diagnostics.field.teamsData', {}, lang)}</div>
        <div class="val">
          {#if !summary.idb.available}
            {t('diagnostics.idb.unavailable', { reason: summary.idb.reason }, lang)}
          {:else if summary.idb.databases.length === 0}
            {t('diagnostics.idb.empty', {}, lang)}
          {:else if summary.idb.databases.length === 1}
            {t('diagnostics.idb.one', {}, lang)}
          {:else}
            {t('diagnostics.idb.many', { n: summary.idb.databases.length }, lang)}
          {/if}
        </div>
      </div>
      <div class="row">
        <div class="label">{t('diagnostics.field.recentExports', {}, lang)}</div>
        <div class="val">
          {#if !summary.exports.available}
            {t('diagnostics.exports.unavailable', { reason: summary.exports.reason }, lang)}
          {:else if summary.exports.items.length === 0}
            {t('diagnostics.exports.none', {}, lang)}
          {:else}
            {t('diagnostics.exports.last', { n: summary.exports.items.length, kind: summary.exports.items[0].kind }, lang)}
          {/if}
        </div>
      </div>
      {#if summary.logsMissing}
        <div class="row warn">
          <div class="label">{t('diagnostics.field.note', {}, lang)}</div>
          <div class="val">{t('diagnostics.logsMissing', { reason: summary.logsMissing }, lang)}</div>
        </div>
      {/if}
    </div>

    <div class="section-title section-title-with-action">
      <span>{t('diagnostics.probes', {}, lang)}</span>
      <button
        class="section-action"
        on:click={runProbes}
        disabled={probesRunning}
        title={t('diagnostics.runProbes', {}, lang)}
      >
        {probesRunning ? t('diagnostics.probesRunning', {}, lang) : t('diagnostics.runProbes', {}, lang)}
      </button>
    </div>
    <div class="card stack">
      {#if probes.state === 'not-run'}
        <div class="row">
          <div class="probes-empty">{t('diagnostics.probesEmpty', {}, lang)}</div>
        </div>
      {:else if probes.state === 'failed'}
        <div class="row warn">
          <div class="label">{t('diagnostics.field.note', {}, lang)}</div>
          <div class="val">{probes.reason}</div>
        </div>
      {:else}
        {#each probes.results as r (r.name)}
          <div class="row probe-row">
            <span class="probe-badge probe-{r.status}" aria-label={r.status}>
              {r.status === 'pass' ? '✓' : r.status === 'fail' ? '✕' : '–'}
            </span>
            <span class="probe-name">{t(`diagnostics.probe.${r.name}`, {}, lang)}</span>
            {#if r.detail}
              <span class="probe-detail">{r.detail}</span>
            {/if}
            <span class="probe-ms">{r.ms} ms</span>
          </div>
        {/each}
      {/if}
    </div>
    {#if probesError}
      <div class="inline-error">{probesError}</div>
    {/if}

    <div class="section-title section-title-with-action">
      <span>{t('diagnostics.fieldDump', {}, lang)}</span>
      <button
        class="section-action"
        on:click={runFieldDump}
        disabled={fieldDumpRunning}
        title={t('diagnostics.fieldDumpRun', {}, lang)}
      >
        {fieldDumpRunning ? t('diagnostics.probesRunning', {}, lang) : t('diagnostics.fieldDumpRun', {}, lang)}
      </button>
    </div>
    <div class="card stack">
      <div class="row">
        <div class="probes-empty">{t('diagnostics.fieldDumpHint', {}, lang)}</div>
      </div>
      {#each fieldDumpRows as row (row.label)}
        <div class="row" class:warn={row.warn}>
          <div class="label">{row.label}</div>
          <div class="val">{row.value}</div>
        </div>
      {/each}
      {#if fieldDumpError}
        <div class="row warn">
          <div class="label">error</div>
          <div class="val">{fieldDumpError}</div>
        </div>
      {/if}
    </div>

    <div class="section-title section-title-with-action">
      <span>{t('diagnostics.fileProbe', {}, lang)}</span>
      <button
        class="section-action"
        on:click={runFileProbe}
        disabled={fileProbeRunning || !fileProbeUrl.trim()}
        title={t('diagnostics.fileProbeRun', {}, lang)}
      >
        {fileProbeRunning ? t('diagnostics.probesRunning', {}, lang) : t('diagnostics.fileProbeRun', {}, lang)}
      </button>
    </div>
    <div class="card stack">
      <div class="row">
        <div class="probes-empty">{t('diagnostics.fileProbeHint', {}, lang)}</div>
      </div>
      <div class="row">
        <input
          class="file-probe-input"
          type="text"
          placeholder="https://…sharepoint.com/…"
          bind:value={fileProbeUrl}
          disabled={fileProbeRunning}
          spellcheck="false"
        />
      </div>
      {#each fileProbeRows as row (row.label)}
        <div class="row" class:warn={row.warn}>
          <div class="label">{row.label}</div>
          <div class="val">{row.value}</div>
        </div>
      {/each}
      {#if fileProbeError}
        <div class="row warn">
          <div class="label">error</div>
          <div class="val">{fileProbeError}</div>
        </div>
      {/if}
    </div>

    <div class="section-title section-title-with-action">
      <span>{t('diagnostics.salvage', {}, lang)}</span>
      <label class="section-action" class:disabled={salvageRunning}>
        {salvageRunning ? t('diagnostics.probesRunning', {}, lang) : t('diagnostics.salvageImport', {}, lang)}
        <input type="file" accept=".txt,text/plain" on:change={runSalvage} disabled={salvageRunning} style="display:none" />
      </label>
    </div>
    <div class="card stack">
      <div class="row">
        <div class="probes-empty">{t('diagnostics.salvageHint', {}, lang)}</div>
      </div>
      {#if salvageStatus}
        <div class="row"><div class="label">status</div><div class="val">{salvageStatus}</div></div>
      {/if}
      {#each salvageRows as row (row.label)}
        <div class="row" class:warn={row.warn}>
          <div class="label">{row.label}</div>
          <div class="val">{row.value}</div>
        </div>
      {/each}
      {#if salvageError}
        <div class="row warn"><div class="label">error</div><div class="val">{salvageError}</div></div>
      {/if}
    </div>

    <div class="actions">
      <label class="raw-toggle" title={t('diagnostics.includeRawIds.tooltip', {}, lang)}>
        <input type="checkbox" checked={includeRawIds} on:change={onRawIdsToggle} />
        {t('diagnostics.includeRawIds', {}, lang)}
      </label>
      <div class="spacer"></div>
      <button class="btn" on:click={() => (showPreview = !showPreview)} title={showPreview ? t('diagnostics.hidePreview', {}, lang) : t('diagnostics.showPreview', {}, lang)}>
        {#if showPreview}<EyeOff size={14} />{:else}<Eye size={14} />{/if}
        {showPreview ? t('diagnostics.hidePreview', {}, lang) : t('diagnostics.showPreview', {}, lang)}
      </button>
    </div>
    {#if toggleError}
      <div class="inline-error">{t('diagnostics.toggleFailed', {}, lang)} {toggleError}</div>
    {/if}

    {#if showPreview}
      <pre class="preview" aria-label="Diagnostic report preview">{reportJson}</pre>
    {/if}

    <div class="actions actions-foot">
      <button class="btn" on:click={onSave} title={t('diagnostics.saveDisk', {}, lang)}>
        <Download size={14} /> {saveConfirmed ? t('diagnostics.saved', {}, lang) : t('diagnostics.saveDisk', {}, lang)}
      </button>
      <div class="spacer"></div>
      <button class="btn primary" on:click={onCopy} title={t('diagnostics.copyReport', {}, lang)}>
        <Copy size={14} /> {copyConfirmed ? t('diagnostics.copied', {}, lang) : t('diagnostics.copyReport', {}, lang)}
      </button>
    </div>
    {#if saveError}
      <div class="inline-error">{saveError}</div>
    {/if}

    <div class="section-title">{t('diagnostics.persistence.title', {}, lang)}</div>
    <div class="card stack">
      <div class="row persist-row">
        <label class="persist-toggle">
          <input
            type="checkbox"
            checked={persistEnabled}
            disabled={persistBusy}
            on:change={onTogglePersist}
          />
          <span class="persist-label">{t('diagnostics.persistence.toggle', {}, lang)}</span>
        </label>
      </div>
      <div class="row persist-hint">
        {persistEnabled
          ? t('diagnostics.persistence.on', {}, lang)
          : t('diagnostics.persistence.off', {}, lang)}
      </div>
      {#if persistEnabled}
        <div class="row persist-storage">
          <span class="storage-label">{t('diagnostics.persistence.size', { bytes: formatBytes(persistBytesUsed) }, lang)}</span>
          <button
            class="btn btn-small"
            on:click={onClearLogs}
            disabled={clearBusy}
            title={t('diagnostics.persistence.clear', {}, lang)}
          >
            <Trash2 size={12} />
            {t('diagnostics.persistence.clear', {}, lang)}
          </button>
        </div>
        {#if persistFlushError}
          <div class="row warn">
            <div class="label">{t('diagnostics.field.note', {}, lang)}</div>
            <div class="val">{t('diagnostics.persistence.writeFailed', { reason: persistFlushError.reason }, lang)}</div>
          </div>
        {/if}
      {/if}
    </div>

    <div class="section-title">{t('diagnostics.verboseStats.title', {}, lang)}</div>
    <div class="card stack">
      <div class="row persist-row">
        <label class="persist-toggle">
          <input
            type="checkbox"
            checked={verboseStatsEnabled}
            disabled={verboseStatsBusy}
            on:change={onToggleVerbose}
          />
          <span class="persist-label">{t('diagnostics.verboseStats.toggle', {}, lang)}</span>
        </label>
      </div>
      <div class="row persist-hint">
        {verboseStatsEnabled
          ? t('diagnostics.verboseStats.on', {}, lang)
          : t('diagnostics.verboseStats.off', {}, lang)}
      </div>
    </div>

    <div class="footer-note">
      {t('diagnostics.footer', {}, lang)}
    </div>
  {/if}
</div>

<style>
  .diag-page {
    padding: 0;
  }
  /* Header and icon buttons reuse the shared .settings-header / .icon-btn
     rules in popup.css, so this page matches Settings and History. Only the
     refresh-spin animation is page-specific. */
  :global(.spin) {
    animation: tce-diag-spin 0.8s linear infinite;
  }
  @keyframes tce-diag-spin {
    to { transform: rotate(360deg); }
  }
  .subtitle {
    color: var(--color-subtle);
    font-size: 12px;
    margin: 0 4px 12px;
    line-height: 1.4;
  }
  .status {
    padding: 14px;
    text-align: center;
    color: var(--color-subtle);
    font-size: 13px;
  }
  .status.err { color: var(--color-danger); }
  .section-title {
    font-size: 11px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    color: var(--color-subtle);
    margin: 14px 6px 6px;
  }
  .section-title-with-action {
    display: flex;
    align-items: center;
    justify-content: space-between;
  }
  .section-action {
    font-size: 12px;
    color: var(--color-accent);
    background: transparent;
    border: 0;
    padding: 0;
    cursor: pointer;
    text-transform: none;
    letter-spacing: 0;
  }
  .section-action[disabled],
  .section-action.disabled {
    color: var(--color-subtle);
    cursor: default;
  }
  .probes-empty {
    font-size: 12px;
    color: var(--color-subtle);
    padding: 4px 6px;
  }
  .file-probe-input {
    width: 100%;
    box-sizing: border-box;
    font-size: 12px;
    font-family: ui-monospace, monospace;
    color: var(--color-text);
    background: var(--color-bg);
    border: 1px solid var(--color-border);
    border-radius: 6px;
    padding: 6px 8px;
  }
  .probe-row {
    align-items: center;
    gap: 10px;
  }
  .probe-badge {
    flex: 0 0 18px;
    width: 18px;
    height: 18px;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    border-radius: 50%;
    color: white;
    font-size: 11px;
    font-weight: 700;
  }
  .probe-pass { background: #16a34a; }
  .probe-fail { background: #dc2626; }
  .probe-skipped { background: #9ca3af; }
  .probe-name {
    flex: 1;
    font-size: 13px;
    color: inherit;
  }
  .probe-detail {
    color: var(--color-subtle);
    font-family: ui-monospace, "SF Mono", Menlo, Consolas, monospace;
    font-size: 11px;
    text-align: right;
  }
  .probe-ms {
    color: var(--color-subtle);
    font-size: 11px;
    flex: 0 0 auto;
    min-width: 50px;
    text-align: right;
  }
  .card {
    background: var(--color-surface);
    border: 1px solid var(--color-border);
    border-radius: 10px;
    box-shadow: 0 1px 2px rgba(0,0,0,0.05);
    margin: 0 4px 12px;
    overflow: hidden;
  }
  .stack .row {
    display: flex; gap: 10px;
    padding: 8px 12px;
    border-bottom: 1px solid var(--color-border);
    font-size: 13px;
  }
  .stack .row:last-child { border-bottom: 0; }
  .stack .label {
    color: var(--color-subtle);
    font-size: 12px;
    min-width: 100px;
  }
  .stack .val {
    flex: 1;
    font-family: ui-monospace, "SF Mono", Menlo, Consolas, monospace;
    font-size: 12px;
    word-break: break-all;
  }
  .stack .row.warn .val { color: var(--color-warn); }
  .actions {
    display: flex; align-items: center; gap: 8px;
    padding: 0 4px;
    margin-bottom: 10px;
  }
  .actions-foot { margin-top: 8px; }
  .actions .spacer { flex: 1; }
  .raw-toggle {
    display: inline-flex; align-items: center; gap: 6px;
    font-size: 12px;
    color: var(--color-subtle);
    cursor: pointer;
  }
  .btn {
    background: var(--color-surface);
    color: inherit;
    border: 1px solid var(--color-border);
    border-radius: 8px;
    padding: 6px 12px;
    cursor: pointer;
    font-size: 13px;
    display: inline-flex; align-items: center; gap: 6px;
  }
  .btn:hover { filter: brightness(1.05); }
  .btn.primary {
    background: var(--color-accent);
    color: white;
    border-color: var(--color-accent);
  }
  .preview {
    background: var(--color-bg);
    border: 1px solid var(--color-border);
    border-radius: 8px;
    margin: 6px 4px 10px;
    padding: 10px;
    max-height: 280px;
    overflow: auto;
    font-family: ui-monospace, "SF Mono", Menlo, Consolas, monospace;
    font-size: 11.5px;
    line-height: 1.5;
    white-space: pre;
    color: inherit;
  }
  .footer-note {
    font-size: 11px;
    color: var(--color-subtle);
    margin: 10px 4px 0;
    line-height: 1.4;
  }
  .inline-error {
    font-size: 12px;
    color: var(--color-danger);
    margin: 4px 4px 8px;
    line-height: 1.4;
  }
  .persist-toggle {
    display: inline-flex;
    align-items: center;
    gap: 8px;
    cursor: pointer;
    font-size: 13px;
  }
  .persist-label { font-weight: 500; }
  .persist-hint {
    font-size: 12px;
    color: var(--color-subtle);
    line-height: 1.4;
  }
  .persist-storage {
    display: flex;
    align-items: center;
    gap: 8px;
  }
  .storage-label {
    flex: 1;
    font-size: 12px;
    color: var(--color-subtle);
  }
  .btn-small {
    padding: 4px 8px;
    font-size: 11.5px;
  }
</style>
