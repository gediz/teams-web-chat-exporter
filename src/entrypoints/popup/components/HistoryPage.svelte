<script lang="ts" module>
  // Firefox polyfill global
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  declare const browser: any;
</script>

<script lang="ts">
  import { createEventDispatcher, onDestroy, onMount } from 'svelte';
  import { ArrowLeft, Trash2, ExternalLink, FolderOpen, X, MessageSquareText } from 'lucide-svelte';
  import { t } from '../../../i18n/i18n';
  import type { HistoryEntry } from '../../../types/shared';

  // Props from App.svelte. The list comes already sorted newest-first
  // (background appends to the front of the array).
  export let entries: HistoryEntry[] = [];
  export let lang = 'en';

  const dispatch = createEventDispatcher<{
    back: void;
    remove: string;
    clearAll: void;
    // Tell the parent to persist "this file is missing on disk" so the
    // grayed state survives popup close/reopen. Especially important on
    // Firefox, where downloads.search() returns stale 'exists: true' for
    // files deleted outside the browser.
    markMissing: string;
  }>();

  // chrome.downloads alias — same browser/chrome polyfill pattern as App.svelte.
  const downloads =
    typeof browser !== 'undefined' ? browser.downloads : chrome.downloads;

  // Per-entry "does this file still exist on disk?" state. Three sources
  // feed this map:
  //   - Persisted entry.fileExists (loaded from storage; survives reopens)
  //   - Live downloads.search() / downloads.onChanged (this session)
  //   - markMissing on click failure (this session AND persisted)
  // Once we know a file is missing, that fact is durable — Firefox's
  // search() may report stale 'exists: true' on the next mount, but
  // entry.fileExists=false outranks it.
  let existsById: Record<string, boolean> = {};

  // Hydrate existsById from persisted entry.fileExists for any entries
  // we haven't already tracked locally. Reactive on `entries` so newly
  // appended rows pick up their persisted state immediately. Local-only
  // updates (live verify, markMissing) take precedence — the merge
  // doesn't overwrite an id that's already in the map.
  $: {
    let changed = false;
    const next = { ...existsById };
    for (const e of entries) {
      if (next[e.id] == null && typeof e.fileExists === 'boolean') {
        next[e.id] = e.fileExists;
        changed = true;
      }
    }
    if (changed) existsById = next;
  }

  // ============================================================
  // Firefox-specific limitation (don't try to "fix" with a Refresh button)
  // ------------------------------------------------------------
  // Firefox's downloads API does NOT actively detect files deleted
  // outside the browser. downloads.search() doesn't trigger a real stat
  // (Bugzilla 1381031, NEW since 2017), and downloads.onChanged doesn't
  // fire for external deletions either.
  //
  // This matches Firefox's own about:downloads UI, which also only
  // greys out actions when the user interacts with a row. We can't do
  // better than the host browser does internally.
  //
  // Detection paths that DO work for us:
  //   1. User clicks Open  → API rejects → markMissing → persisted
  //   2. Browser fires onChanged exists=false (Chrome reliably; Firefox rarely)
  //   3. Hover/scroll re-verify → works on Chrome only
  //
  // Once an entry is marked missing, we don't re-verify it (Firefox would
  // flip it back to "exists: true" from cached state). The persisted
  // fileExists=false stays until the user manually removes the row.
  // A "Refresh" button would do nothing on Firefox — calling search()
  // returns the same stale data.
  // ============================================================

  // Verify a single entry. On Chrome, .search() actually re-stats the
  // file and returns fresh state (and downloads.onChanged fires shortly
  // after with any change). On Firefox, the returned value is stale for
  // externally-deleted files; we rely on the click-failure path instead.
  const verifyEntry = async (entry: HistoryEntry): Promise<boolean> => {
    if (entry.kind !== 'success' || entry.downloadId == null) return true;
    try {
      const items = await downloads.search({ id: entry.downloadId });
      const item = Array.isArray(items) ? items[0] : undefined;
      return !!item?.exists;
    } catch {
      return false;
    }
  };

  // Live re-verify all entries. Skips entries already known missing so
  // Firefox's stale 'exists: true' doesn't accidentally un-mark them.
  const verifyAll = async () => {
    try {
      const toCheck = entries.filter(e => existsById[e.id] !== false);
      const checks = await Promise.all(
        toCheck.map(async (e) => [e.id, await verifyEntry(e)] as const),
      );
      const next = { ...existsById };
      for (const [id, ok] of checks) next[id] = ok;
      existsById = next;
    } catch {
      // best-effort
    }
  };

  // Engagement-triggered re-verification. Each entry has a per-id cooldown
  // so a row that was just checked won't be re-checked when the user wiggles
  // their mouse over it. Cheap, silent, and only runs when the user is
  // actually interacting with the list.
  const VERIFY_COOLDOWN_MS = 2000;
  const lastVerifiedAt = new Map<string, number>();
  const verifyOne = (entry: HistoryEntry) => {
    if (entry.kind !== 'success' || entry.downloadId == null) return;
    // Already known missing — don't re-ask Firefox for stale "exists: true".
    if (existsById[entry.id] === false) return;
    const now = Date.now();
    const last = lastVerifiedAt.get(entry.id) ?? 0;
    if (now - last < VERIFY_COOLDOWN_MS) return;
    lastVerifiedAt.set(entry.id, now);
    void verifyEntry(entry).then((exists) => {
      existsById = { ...existsById, [entry.id]: exists };
    });
  };
  // Scroll handler: verify entries whose last check is older than the
  // cooldown. Throttled at the call-site by the per-id lastVerifiedAt
  // check inside verifyOne, so even a 60fps scroll only fires real
  // verifies for entries that genuinely need a refresh.
  const onListScroll = () => {
    for (const e of entries) verifyOne(e);
  };

  // downloads.onChanged fires when Chrome/Firefox finishes the deferred
  // existence check that .search() kicked off. This is the one and only
  // reliable path on Firefox — the search() return value is cached/stale
  // there. We map the changed download id back to its history entry id
  // and update the row immediately.
  const onDownloadChanged = (delta: any) => {
    if (!delta || delta.exists == null) return;
    const matching = entries.find(e => e.downloadId === delta.id);
    if (!matching) return;
    const exists = !!delta.exists.current;
    console.log('[History] downloads.onChanged exists=', exists, 'for id', delta.id);
    existsById = { ...existsById, [matching.id]: exists };
    // Persist when the browser tells us the file is gone — so the next
    // popup open reflects this without us having to detect it again.
    if (!exists) dispatch('markMissing', matching.id);
  };

  // Re-run verification whenever `entries` changes (parent removes/clears).
  $: void verifyAll(), entries;

  onMount(() => {
    void verifyAll();
    // Re-verify after a short delay. Firefox's first .search() returns
    // cached state and triggers an async existence check; the second
    // .search() (and/or the onChanged event) returns the updated state.
    // 600ms is enough for the OS-level filesystem stat to complete in
    // practice but short enough that the user won't notice the catch-up.
    setTimeout(() => { void verifyAll(); }, 600);
    // When the user comes back to the popup after working in their OS file
    // manager (where they may have deleted exported files), re-verify so
    // missing-file rows update without requiring a popup close/reopen.
    window.addEventListener('focus', verifyAll);
    // Cross-browser: listen for deferred existence-check results.
    try { downloads.onChanged.addListener(onDownloadChanged); } catch { /* noop */ }
  });
  onDestroy(() => {
    window.removeEventListener('focus', verifyAll);
    try { downloads.onChanged.removeListener(onDownloadChanged); } catch { /* noop */ }
  });

  // Smart hybrid date format. Recent → relative ("just now", "today HH:mm",
  // "yesterday HH:mm"). Older → absolute month-day. Tooltip always carries
  // the full ISO timestamp for precision.
  const formatDate = (savedAt: number): string => {
    const now = Date.now();
    const diffMs = now - savedAt;
    const diffSec = Math.floor(diffMs / 1000);
    if (diffSec < 60) return t('history.justNow', {}, lang) || 'just now';

    const date = new Date(savedAt);
    const today = new Date();
    const isToday = date.toDateString() === today.toDateString();
    const yesterday = new Date(today);
    yesterday.setDate(today.getDate() - 1);
    const isYesterday = date.toDateString() === yesterday.toDateString();

    const hh = String(date.getHours()).padStart(2, '0');
    const mm = String(date.getMinutes()).padStart(2, '0');
    const time = `${hh}:${mm}`;

    if (isToday) return `${t('history.today', {}, lang) || 'today'} ${time}`;
    if (isYesterday) return `${t('history.yesterday', {}, lang) || 'yesterday'} ${time}`;

    const days = Math.floor(diffMs / (24 * 60 * 60 * 1000));
    if (days < 7) return t('history.daysAgo', { n: days }, lang) || `${days} days ago`;

    const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
    return `${months[date.getMonth()]} ${date.getDate()}`;
  };

  const formatTooltip = (savedAt: number): string => {
    const d = new Date(savedAt);
    return d.toISOString().replace('T', ' ').slice(0, 19);
  };

  // Open / Show actions. Important: NO `await` between the click and the
  // downloads.* call — Firefox enforces a strict user-activation window
  // for downloads.open() and any await (even a fast cache lookup) breaks
  // it ("downloads.open may only be called from a user input handler").
  // For the missing-file guard we use the synchronous existsById cache.
  //
  // On a failed Open, we mark the entry missing immediately. The failure
  // itself is the strongest signal we have — Firefox's downloads.search()
  // returns stale 'exists: true' for files deleted outside the browser,
  // and onChanged doesn't reliably fire either. The successful click
  // attempt is what proves the file is gone.
  const markMissing = (entry: HistoryEntry) => {
    existsById = { ...existsById, [entry.id]: false };
    // Persist via parent so the missing state survives popup close/reopen.
    dispatch('markMissing', entry.id);
  };
  // Translate the browser's raw error into something a developer reading
  // the console can actually act on. Firefox in particular returns a
  // useless "An unexpected error occurred" for every downloads.open()
  // failure — but we just synchronously checked that we believed the
  // file existed, so the overwhelming cause is "user deleted it from
  // disk after the last successful check."
  const explainOpenError = (raw: string): string => {
    const generic = /unexpected error|file does not exist|not.*ready|state.*in_progress/i.test(raw);
    if (generic) {
      return `the file is no longer on disk (raw: "${raw}")`;
    }
    return raw;
  };

  const onOpen = (e: MouseEvent, entry: HistoryEntry) => {
    e.stopPropagation();
    if (entry.downloadId == null) return;
    // Sync gate against local cache — never await before the API call.
    if (existsById[entry.id] === false) return;
    const id = entry.downloadId;
    Promise.resolve(downloads.open(id)).catch((err: any) => {
      const raw = err?.message || String(err);
      console.warn(
        `[History] Open failed for download id ${id} — ${explainOpenError(raw)}. ` +
        `Marking row as missing.`,
      );
      markMissing(entry);
    });
  };
  // For Show, we can't use the failure as a missing-file signal: Firefox's
  // downloads.show() returns successfully even when the file isn't on disk
  // (it just opens the folder). So we keep the cache-based gate; the only
  // way to detect post-deletion is the eventual onChanged event.
  const onShow = (e: MouseEvent, entry: HistoryEntry) => {
    e.stopPropagation();
    if (entry.downloadId == null) return;
    if (existsById[entry.id] === false) return;
    const id = entry.downloadId;
    Promise.resolve(downloads.show(id)).catch((err: any) => {
      const raw = err?.message || String(err);
      console.warn(
        `[History] Show failed for download id ${id} — ${explainOpenError(raw)}. ` +
        `Marking row as missing.`,
      );
      markMissing(entry);
    });
  };
  const onRemove = (e: MouseEvent, id: string) => {
    e.stopPropagation();
    dispatch('remove', id);
  };

  // Headline rendered in the row title slot. Prefer the chat title (more
  // useful for browsing — "the Project X export"), fall back to the
  // filename if title is missing (older entries before we captured it),
  // and use a fixed "(cancelled)" string for cancelled entries.
  const headlineFor = (entry: HistoryEntry): string => {
    if (entry.kind === 'cancelled') {
      return t('history.cancelledTitle', {}, lang) || '(cancelled — no file saved)';
    }
    return entry.title || entry.filename || '(untitled export)';
  };

  // Badge text. Bundle exports (2+ formats packaged into one .zip) get a
  // dedicated "BUNDLE" badge instead of any single format name. For older
  // entries that pre-date the `formats` array, fall back to the singular
  // `format` field, then to the file extension.
  const formatLabel = (entry: HistoryEntry): string => {
    if (entry.formats && entry.formats.length >= 2) return 'BUNDLE';
    const single = entry.formats?.[0] || entry.format;
    if (single) return single.toUpperCase();
    if (entry.filename) {
      const ext = entry.filename.match(/\.([a-z0-9]+)$/i)?.[1];
      if (ext) return ext.toUpperCase();
    }
    return '?';
  };
  const formatClass = (entry: HistoryEntry): string => {
    const label = formatLabel(entry).toLowerCase();
    if (label === 'html' || label === 'htm') return 'html';
    if (label === 'json') return 'json';
    if (label === 'csv') return 'csv';
    if (label === 'txt') return 'txt';
    if (label === 'pdf') return 'pdf';
    if (label === 'zip') return 'zip';
    if (label === 'bundle') return 'bundle';
    return 'html';
  };
  // Tooltip showing the bundle's contents (e.g. "HTML, JSON, CSV").
  // Empty for non-bundle entries — let the badge speak for itself.
  const formatBadgeTooltip = (entry: HistoryEntry): string => {
    if (!entry.formats || entry.formats.length < 2) return '';
    return entry.formats.map(f => f.toUpperCase()).join(', ');
  };
</script>

<div class="settings-page">
  <div class="settings-header">
    <button class="icon-btn" title={t('actions.back', {}, lang) || 'Back'} on:click={() => dispatch('back')}>
      <ArrowLeft size={18} />
    </button>
    <h1>{t('history.title', {}, lang) || 'Export history'}</h1>
    {#if entries.length > 0}
      <span class="header-count">{entries.length}</span>
      <button class="header-clear" on:click={() => dispatch('clearAll')}>
        <Trash2 size={12} />
        {t('history.clearAll', {}, lang) || 'Clear all'}
      </button>
    {/if}
  </div>

  {#if entries.length === 0}
    <div class="card history-empty">
      <div class="empty-icon">📭</div>
      <div class="empty-title">{t('history.empty.title', {}, lang) || 'No exports yet'}</div>
      <div class="empty-msg">{t('history.empty.msg', {}, lang) || "Run your first export and it'll appear here."}</div>
    </div>
  {:else}
    <div class="card history-card" on:scroll={onListScroll}>
      {#each entries as entry (entry.id)}
        {@const fileExists = entry.kind !== 'success' || entry.downloadId == null
          ? true
          : (existsById[entry.id] ?? true)}
        {@const isMissing = entry.kind === 'success' && !fileExists}
        {@const hasMsgCount = typeof entry.messageCount === 'number' && entry.messageCount > 0}
        <!-- The row itself isn't interactive — the buttons inside are.
             role="presentation" tells assistive tech to skip the row
             container and focus on the button children. The mouseenter
             handler is incidental UX (file-existence freshness). -->
        <div
          class="row"
          role="presentation"
          class:cancelled={entry.kind === 'cancelled'}
          class:missing={isMissing}
          on:mouseenter={() => verifyOne(entry)}
        >
          <div
            class="badge badge-{formatClass(entry)}"
            class:badge-cancelled={entry.kind === 'cancelled'}
            class:badge-missing={isMissing}
            title={formatBadgeTooltip(entry)}
          >
            {entry.kind === 'cancelled' ? '✕' : formatLabel(entry)}
          </div>
          <div class="body">
            <!-- Headline = chat title (more useful than auto filename).
                 Filename lives in the tooltip for users who need it. -->
            <div class="title-line">
              <span class="title" title={entry.filename || ''}>{headlineFor(entry)}</span>
              {#if isMissing}
                <span class="status-pill status-missing">{t('history.fileMissing', {}, lang) || 'file missing'}</span>
              {:else if entry.kind === 'cancelled'}
                <span class="status-pill status-cancelled">{t('history.cancelledMeta', {}, lang) || 'cancelled'}</span>
              {/if}
            </div>
            <div class="meta">
              {#if hasMsgCount}
                <span class="meta-item" title={t('phase.label.messages', {}, lang) || 'messages'}>
                  <MessageSquareText size={11} />
                  {entry.messageCount?.toLocaleString()}
                </span>
                <span class="meta-sep" aria-hidden="true">·</span>
              {/if}
              <span class="meta-item when" title={formatTooltip(entry.savedAt)}>
                {formatDate(entry.savedAt)}
              </span>
            </div>
          </div>
          <div class="actions">
            {#if entry.kind === 'success' && !isMissing}
              <button class="ico-btn" title={t('actions.open', {}, lang) || 'Open'} on:click={(e) => onOpen(e, entry)}>
                <ExternalLink size={14} />
              </button>
              <button class="ico-btn" title={t('actions.showFolder', {}, lang) || 'Show in folder'} on:click={(e) => onShow(e, entry)}>
                <FolderOpen size={14} />
              </button>
            {/if}
            <button class="ico-btn dismiss" title={t('history.remove', {}, lang) || 'Remove from history'} on:click={(e) => onRemove(e, entry.id)}>
              <X size={14} />
            </button>
          </div>
        </div>
      {/each}
    </div>
  {/if}
</div>

<style>
  /* The page wrapper and header reuse the global .settings-page,
     .settings-header, .icon-btn and h1 rules from popup.css so this view
     visually matches the Settings page. Per-page extras live below. */

  .header-count {
    font-size: 11px; font-weight: 600;
    color: var(--color-text-muted);
    background: rgba(0, 0, 0, 0.05);
    padding: 2px 8px;
    border-radius: 999px;
    margin-left: 4px;
  }
  :global(body[data-theme='dark']) .header-count {
    background: rgba(255, 255, 255, 0.08);
  }
  .header-clear {
    margin-left: auto;
    background: none; border: none;
    color: var(--color-text-muted);
    font-size: 12px; font-weight: 500;
    padding: 4px 8px; border-radius: 4px;
    cursor: pointer;
    display: inline-flex; align-items: center; gap: 4px;
  }
  .header-clear:hover { background: rgba(220, 38, 38, 0.10); color: #dc2626; }

  /* Empty-state card matches the rest of the page: same surface card. */
  .history-empty {
    text-align: center;
    padding: 40px 24px;
  }
  .empty-icon { font-size: 32px; margin-bottom: 8px; }
  .empty-title {
    font-size: 14px; font-weight: 600;
    color: var(--color-text);
    margin-bottom: 4px;
  }
  .empty-msg {
    font-size: 12px;
    color: var(--color-text-muted);
  }

  /* History list lives inside a .card so it shares the surface with
     Settings cards. Internal padding is 0 so rows can fill edge-to-edge. */
  .history-card {
    padding: 4px;
    max-height: 460px;
    overflow-y: auto;
  }
  .row {
    display: grid;
    grid-template-columns: 36px 1fr auto;
    gap: 10px;
    padding: 10px;
    border-radius: 6px;
    align-items: center;
    transition: background 0.15s ease;
  }
  .row:hover { background: rgba(0, 0, 0, 0.025); }
  :global(body[data-theme='dark']) .row:hover { background: rgba(255, 255, 255, 0.04); }
  .row + .row { border-top: 1px solid var(--color-border); }

  .row.missing { opacity: 0.7; }
  .row.missing .title { text-decoration: line-through; color: var(--color-text-muted); }
  .row.cancelled .title { color: var(--color-text-muted); }

  .badge {
    width: 36px; height: 36px;
    border-radius: 8px;
    display: flex; align-items: center; justify-content: center;
    font-size: 11px; font-weight: 700;
    flex-shrink: 0;
    letter-spacing: 0.04em;
  }
  .badge-html { background: rgba(37, 99, 235, 0.10); color: #2563eb; }
  .badge-json { background: rgba(217, 119, 6, 0.10); color: #d97706; }
  .badge-csv  { background: rgba(22, 163, 74, 0.10); color: #16a34a; }
  .badge-txt  { background: rgba(124, 58, 237, 0.10); color: #7c3aed; }
  /* PDF gets a red family to read as "document/print" — distinct enough
     from HTML's blue and BUNDLE's purple that the badge identifies format
     at a glance. */
  .badge-pdf  { background: rgba(220, 38, 38, 0.10); color: #dc2626; }
  .badge-zip  { background: rgba(124, 58, 237, 0.10); color: #7c3aed; }
  /* Bundle = 2+ formats packed into one .zip. Same purple family as
     .zip but darker/saturated so it reads as "this is a multi-file
     archive" at a glance. Kept narrower-feeling via slightly tighter
     letter-spacing to fit "BUNDLE" without growing the badge. */
  .badge-bundle { background: rgba(109, 40, 217, 0.14); color: #6d28d9; letter-spacing: 0.02em; }
  .badge-cancelled { background: rgba(220, 38, 38, 0.10); color: #dc2626; font-size: 14px; }
  .badge-missing { background: rgba(0, 0, 0, 0.05); color: var(--color-text-muted); }

  .body {
    min-width: 0;
  }
  /* Title-line is the headline + a status pill side-by-side. The pill
   * is flex-shrink: 0 so it never gets clipped — it's a critical state
   * indicator (file missing / cancelled). The title text takes the
   * remaining space and clips silently. */
  .title-line {
    display: flex;
    align-items: center;
    gap: 6px;
    min-width: 0;
  }
  .title {
    font-size: 13px; font-weight: 600;
    color: var(--color-text);
    overflow: hidden;
    white-space: nowrap;
    text-overflow: clip;
    min-width: 0;
    flex: 1 1 auto;
    /* Soft fade-out at the right edge so a long chat title doesn't
     * crash into the status pill or the action buttons. The mask
     * is applied here (not on the parent .body) so the pill — which
     * is a sibling of .title — stays fully opaque. */
    -webkit-mask-image: linear-gradient(to right, black calc(100% - 18px), transparent);
    mask-image:         linear-gradient(to right, black calc(100% - 18px), transparent);
  }
  .status-pill {
    flex-shrink: 0;
    font-size: 10px; font-weight: 600;
    padding: 1px 6px;
    border-radius: 999px;
    line-height: 1.4;
    letter-spacing: 0.01em;
  }
  .status-missing {
    background: rgba(217, 119, 6, 0.12);
    color: #b45309;
  }
  .status-cancelled {
    background: rgba(220, 38, 38, 0.10);
    color: #dc2626;
  }
  /* Meta line uses inline-flex with a fixed gap so spacing around
   * separators is structural — no fragile reliance on whitespace inside
   * Svelte template literals (which Svelte was collapsing). Each item
   * (count, separator, date) is its own .meta-item span. */
  .meta {
    display: flex;
    align-items: center;
    flex-wrap: nowrap;
    gap: 6px;
    font-size: 11px;
    color: var(--color-text-muted);
    overflow: hidden;
    white-space: nowrap;
    margin-top: 2px;
    min-width: 0;
  }
  .meta .meta-item {
    display: inline-flex;
    align-items: center;
    gap: 4px;
    flex-shrink: 0;
  }
  .meta .meta-item.when { color: var(--color-text-dim, #94a3b8); }
  .meta .meta-sep {
    color: var(--color-text-dim, #94a3b8);
    flex-shrink: 0;
  }

  .actions {
    display: flex; gap: 2px; flex-shrink: 0;
  }
  .ico-btn {
    width: 26px; height: 26px;
    border: none; background: transparent;
    color: var(--color-text-dim, #94a3b8);
    border-radius: 5px;
    cursor: pointer;
    display: flex; align-items: center; justify-content: center;
    opacity: 0.55;
    transition: opacity 0.15s ease, background 0.15s ease, color 0.15s ease;
  }
  .row:hover .ico-btn { opacity: 1; }
  .ico-btn:hover { background: rgba(37, 99, 235, 0.10); color: #2563eb; opacity: 1; }
  .ico-btn.dismiss:hover { background: rgba(220, 38, 38, 0.10); color: #dc2626; }
  /* For missing / cancelled rows, the × is the only meaningful action.
   * Bump its baseline visibility so users can find it without hovering. */
  .row.missing .ico-btn.dismiss,
  .row.cancelled .ico-btn.dismiss { opacity: 0.85; }
</style>
