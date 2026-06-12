<script lang="ts">
  import { createEventDispatcher, tick, onDestroy } from 'svelte';
  import {
    MessageSquare, Users, Calendar, Hash,
    Search, RefreshCw, AlertCircle, List, Check, Folder, Star,
    ChevronDown, CheckSquare, Square, Repeat, Trash2,
  } from 'lucide-svelte';
  import { t } from '../../../i18n/i18n';
  import type { ConversationSummary, ConversationKind, FolderSummary, SavedGroup } from '../../../types/shared';
  import { conversationDisplayName, conversationDisplaySubtitle } from '../../../utils/conversation-labels';
  import SavedGroups from './SavedGroups.svelte';

  export let lang = 'en';
  export let conversations: ConversationSummary[] = [];
  // Folder rail data — Favorites + user-created folders that have at
  // least one conversation. System-computed folders are stripped at
  // the source (see listConversationsFromIdb), so anything here is
  // safe to render as a filterable folder.
  export let folders: FolderSummary[] = [];
  // Caller owns the array and binds it two-way. In 'single' mode clicking
  // a row replaces the selection; in 'multi' mode clicking toggles so the
  // user can build up a batch. The picker always emits an array so the
  // caller's shape is stable across modes.
  export let selectedIds: string[] = [];
  // Selected folder id; 'all' means no folder filter. Two-way bound so
  // the caller can persist it across popup reopens. Per design (Q2 in
  // the round-5 mockup) there is no "All folders" row — clicking the
  // active folder again clears it back to 'all'.
  export let selectedFolderId: string = 'all';
  // Selected kind ("all" / "chat" / "group" / "meeting" / "channel").
  // Two-way bound so the caller can persist it the same way it persists
  // the folder. Defaulting to "all" keeps the picker open on every
  // conversation when there's no saved preference.
  export let selectedKind: 'all' | ConversationKind = 'all';
  export let mode: 'single' | 'multi' = 'single';
  // Load state: 'idle' (not started) → 'loading' → 'ok' | 'error'.
  // 'error' shows a retry button; 'loading' dims the list.
  export let state: 'idle' | 'loading' | 'ok' | 'error' = 'idle';
  // Background refresh flag — fires when a cached list is visible and a
  // fresh fetch is running alongside. Shows a subtle "Refreshing…" hint
  // in the header without blocking interaction.
  export let refreshing = false;
  export let errorMessage = '';
  // True when the popup opened on a tab that isn't Teams web. Bypasses
  // the entire load/error/empty rendering: there's literally nothing
  // for the picker to do, so we render a single "open Teams" hint
  // instead of the misleading default empty-state.
  export let notOnTeams = false;
  // Collapsed state — when true, only the card header (title + count
  // + chevron) renders. Two-way bindable so the parent can persist it
  // across popup opens. Backwards-compat: the active chat is still
  // auto-picked into selectedIds on mount, so a user who never
  // expands the picker can still single-click Export to export the
  // current chat (matches the pre-v1.4.0 habit).
  export let collapsed = false;

  // Saved selection presets (owned + persisted by App). The picker hosts the
  // Presets menu in its head, forwards save/remove to App, and handles apply
  // itself. Internal identifiers keep the "group" name (and the storage key is
  // unchanged) so existing saved data isn't orphaned; only the UI label is
  // "Presets" to avoid colliding with the rail's "Groups" chat-type filter.
  export let savedGroups: SavedGroup[] = [];

  const dispatch = createEventDispatcher<{
    change: string[];           // fired whenever the selection set changes
    folderChange: string;       // fired when the selected folder changes
    kindChange: 'all' | ConversationKind; // fired when the kind tab changes
    collapseChange: boolean;    // fired when the user toggles collapse
    retry: void;
    saveGroup: string;          // new group name — App persists it from the live selection
    removeGroup: string;        // saved-group id to delete
  }>();

  function toggleCollapse() {
    collapsed = !collapsed;
    dispatch('collapseChange', collapsed);
  }

  // ── Auto-scroll the active folder into view (+ brief flash) ──────────
  // When the picker opens (or expands from collapsed), the user might
  // have a folder filter active that lives below the rail's visible
  // viewport. Without this, they'd see a short conversation list with
  // no obvious explanation. Solution: scroll the active folder into
  // view, then flash its background briefly so the user sees "here is
  // your active filter".
  //
  // Only triggers when:
  //   - picker is currently expanded (collapsed === false)
  //   - a folder filter is active (selectedFolderId !== 'all')
  //   - the active folder isn't already fully visible
  //
  // Pattern: a `bind:this` on the rail + querySelector on the active
  // folder row at trigger time. Reactive `$:` watches the (mounted,
  // !collapsed) gate so it fires on initial mount AND any subsequent
  // expand from the collapsed state.
  let railEl: HTMLDivElement | undefined;
  let railMounted = false;
  // Svelte action that flips the railMounted flag once the rail node is
  // attached. We can't simply use `bind:this` to gate the reactive
  // statement because bind:this fires during the same microtask as the
  // first render — we need the next tick so the active rail item's
  // class has been applied.
  function railMountedAction(_node: HTMLElement) {
    railMounted = true;
    return { destroy() { railMounted = false; } };
  }
  function scrollActiveFolderIntoView() {
    if (!railEl || collapsed) return;
    if (selectedFolderId === 'all') return;
    // Last `.rail-item.active` is the folder (folders render after
    // type tabs in the rail). If only a kind tab is active, no scroll.
    const actives = railEl.querySelectorAll('.rail-item.active');
    const target = actives.length > 0 ? actives[actives.length - 1] as HTMLElement : null;
    if (!target) return;
    // Don't scroll if already fully visible — avoids unnecessary motion
    // when the user reopens the popup with the active folder still in
    // view from last session (and prevents a double-scroll if the rail
    // happens to be at the right position already).
    const rRect = railEl.getBoundingClientRect();
    const tRect = target.getBoundingClientRect();
    const fullyVisible = tRect.top >= rRect.top && tRect.bottom <= rRect.bottom;
    if (!fullyVisible) {
      target.scrollIntoView({ block: 'center', behavior: 'smooth' });
    }
    // Flash regardless — even when already visible, the brief flash
    // makes the active filter pop on popup open. ~350ms delay so the
    // smooth-scroll has time to land before the colour change starts.
    setTimeout(() => {
      target.classList.remove('rail-item-flash');
      // Force a reflow so re-adding the class restarts the animation.
      void target.offsetWidth;
      target.classList.add('rail-item-flash');
      setTimeout(() => target.classList.remove('rail-item-flash'), 1500);
    }, 350);
  }
  $: if (railMounted && !collapsed) {
    // Ensure DOM has settled (the active row needs to be rendered with
    // its current `active` class) before we try to find + scroll to it.
    tick().then(scrollActiveFolderIntoView);
  }

  type RailKind = 'all' | ConversationKind;
  // currentKind is internal state derived from the bound selectedKind
  // prop. We keep an internal alias for terser conditionals and to mark
  // the rail-seeding block as having "owned" the value.
  $: currentKind = selectedKind;
  let query = '';
  let searchInputEl: HTMLInputElement | undefined;
  function clearQuery() {
    query = '';
    tick().then(() => searchInputEl?.focus());
  }

  // After the picker has fully hydrated, slide the currently-selected
  // row into view so the user doesn't have to scroll to find what's
  // already picked. We deliberately do NOT change the kind tab here —
  // the parent persists the user's last kind choice across popup
  // opens, and overriding it with "kind of currently-selected chat"
  // has been a source of bugs (the kind would silently flip from
  // "All" to "Chats" a few seconds after open as the selection
  // hydrated). Keep scroll-into-view, drop the auto-seed.
  let scrolledToSelection = false;
  let listEl: HTMLDivElement | undefined;
  $: if (!scrolledToSelection && selectedIds.length > 0 && conversations.length > 0) {
    scrolledToSelection = true;
    tick().then(() => {
      const el = listEl?.querySelector('.picker-row.selected') as HTMLElement | null;
      el?.scrollIntoView({ block: 'nearest' });
    });
  }

  // Kind order in the rail. Keeping it stable so the tooltip/layout
  // matches the mockup the user picked (L).
  const RAIL_KINDS: { kind: RailKind; labelKey: string; fallback: string }[] = [
    { kind: 'all',     labelKey: 'picker.rail.all',      fallback: 'All' },
    { kind: 'chat',    labelKey: 'picker.rail.chats',    fallback: 'Chats' },
    { kind: 'group',   labelKey: 'picker.rail.groups',   fallback: 'Groups' },
    { kind: 'meeting', labelKey: 'picker.rail.meetings', fallback: 'Meetings' },
    { kind: 'channel', labelKey: 'picker.rail.channels', fallback: 'Channels' },
  ];

  const iconFor = (kind: ConversationKind) => {
    switch (kind) {
      case 'chat':    return MessageSquare;
      case 'group':   return Users;
      case 'meeting': return Calendar;
      case 'channel': return Hash;
    }
  };

  // Normalise name for substring filtering. Lowercasing is good enough
  // for the v1 UX; diacritics folding can follow if Turkish users
  // complain about "Güçlü" not matching "guclu".
  const norm = (s: string) => (s || '').toLowerCase();

  const rowName = (c: ConversationSummary) => conversationDisplayName(c, lang, t);
  const rowSubtitle = (c: ConversationSummary) => conversationDisplaySubtitle(c, lang, t);

  $: selectedSet = new Set(selectedIds);

  // Per-kind selection marker for the rail dots. Only populated for
  // kinds that have at least one selected item so the dot doesn't
  // render on empty kinds.
  $: kindHasSelection = (() => {
    const s: Record<ConversationKind, boolean> = {
      chat: false, group: false, meeting: false, channel: false,
    };
    for (const id of selectedSet) {
      const c = conversations.find(x => x.id === id);
      if (c) s[c.kind] = true;
    }
    return s;
  })();

  // Filter pipeline: kind first (cheap), then folder (Set lookup),
  // then query substring (most expensive — runs locale-aware label
  // resolution per row).
  $: byKind = currentKind === 'all'
    ? conversations
    : conversations.filter(c => c.kind === currentKind);

  $: byFolder = selectedFolderId === 'all'
    ? byKind
    : byKind.filter(c => Array.isArray(c.folderIds) && c.folderIds.includes(selectedFolderId));

  $: filtered = !query.trim()
    ? byFolder
    : byFolder.filter(c => {
        const q = norm(query.trim());
        // Filter against what the user actually sees (rowName/rowSubtitle
        // already applied locale-aware placeholders + list formatting),
        // so a search for e.g. the Turkish "Kendi sohbetiniz" string
        // matches the self-chat row.
        return norm(rowName(c)).includes(q)
          || norm(rowSubtitle(c) || '').includes(q);
      });

  // Count of selections visible under the current rail kind. Drives
  // the "N selected" pill in the header — we want it to reflect what
  // the user can actually see right now, not a global total.
  $: visibleSelectedCount = filtered.reduce(
    (n, c) => n + (selectedSet.has(c.id) ? 1 : 0),
    0,
  );

  function toggle(id: string) {
    let next: string[];
    if (mode === 'single') {
      // Radio-style: clicking a row makes it the sole selection. Clicking
      // the already-selected row does nothing (we don't want a no-
      // selection state to be reachable via a misclick).
      next = selectedSet.has(id) ? selectedIds : [id];
    } else {
      next = selectedSet.has(id)
        ? selectedIds.filter(x => x !== id)
        : [...selectedIds, id];
    }
    if (next === selectedIds) return;
    selectedIds = next;
    dispatch('change', next);
  }

  // ── Bulk-select helpers ────────────────────────────────────────────
  // Two affordances per the M+N design from the bulk-select mockup:
  //   M  the existing "N selected" head pill becomes a dropdown that
  //      opens a context-aware menu (bulk actions for the current view).
  //   N  a small icon-only action bar below the search with three fixed-
  //      width buttons: select-all-visible, invert-visible, clear-all.
  // Both routes manipulate `selectedIds` and dispatch 'change' so the
  // parent stays in sync.
  let bulkMenuOpen = false;
  // Anchor element for the dropdown popover — used to compute click-
  // outside dismissal. Bound to the head pill button.
  let bulkAnchorEl: HTMLButtonElement | undefined;

  function setSelection(next: string[]) {
    if (next.length === selectedIds.length
        && next.every((id, i) => id === selectedIds[i])) return;
    selectedIds = next;
    dispatch('change', next);
  }

  // Add every conversation in `list` to the current selection. Order is
  // preserved (existing first, then new IDs in their original order).
  function addAll(list: ConversationSummary[]) {
    if (mode === 'single') return; // bulk only meaningful in multi mode
    const have = new Set(selectedIds);
    const additions: string[] = [];
    for (const c of list) {
      if (!have.has(c.id)) additions.push(c.id);
    }
    if (additions.length === 0) return;
    setSelection([...selectedIds, ...additions]);
  }
  // Remove every conversation in `list` from the current selection.
  function removeAll(list: ConversationSummary[]) {
    if (mode === 'single') return;
    if (!list.length) return;
    const drop = new Set(list.map(c => c.id));
    const next = selectedIds.filter(id => !drop.has(id));
    setSelection(next);
  }
  function toggleAll(list: ConversationSummary[]) {
    if (mode === 'single') return;
    const allSelected = list.length > 0 && list.every(c => selectedSet.has(c.id));
    if (allSelected) removeAll(list);
    else addAll(list);
  }
  function invertSelection(list: ConversationSummary[]) {
    if (mode === 'single') return;
    const drop = new Set<string>();
    const add: string[] = [];
    for (const c of list) {
      if (selectedSet.has(c.id)) drop.add(c.id);
      else add.push(c.id);
    }
    const next = selectedIds.filter(id => !drop.has(id)).concat(add);
    setSelection(next);
  }
  function clearAll() {
    if (mode === 'single') return;
    if (selectedIds.length === 0) return;
    setSelection([]);
  }

  // Saved-group apply: re-select the group's convIds that still exist in this
  // account (account-agnostic — a foreign group resolves to 0 matches). Shows
  // a transient banner only when some ids couldn't be matched.
  let groupBanner = '';
  let groupBannerTimer: ReturnType<typeof setTimeout> | undefined;
  function applyGroup(g: SavedGroup) {
    const available = new Set(conversations.map(c => c.id));
    const matched = g.convIds.filter(id => available.has(id));
    const missing = g.convIds.length - matched.length;
    setSelection(matched);
    if (groupBannerTimer) clearTimeout(groupBannerTimer);
    if (missing > 0) {
      groupBanner = t('groups.appliedPartial', { found: matched.length, total: g.convIds.length, missing }, lang)
        || `Applied ${matched.length} of ${g.convIds.length} (${missing} unavailable)`;
      groupBannerTimer = setTimeout(() => { groupBanner = ''; }, 4000);
    } else {
      groupBanner = '';
    }
  }
  onDestroy(() => { if (groupBannerTimer) clearTimeout(groupBannerTimer); });

  // Bulk targets: action-bar buttons act on what the user can actually
  // see right now (`filtered`), so the kind/folder/search composes
  // naturally. The dropdown menu (M) additionally exposes per-kind /
  // per-folder shortcuts when those filters are active.
  $: allVisibleSelected = filtered.length > 0
    && filtered.every(c => selectedSet.has(c.id));
  $: anyVisibleSelected = filtered.some(c => selectedSet.has(c.id));

  function chatsOfKind(k: RailKind): ConversationSummary[] {
    if (k === 'all') return conversations;
    return conversations.filter(c => c.kind === k);
  }
  function chatsInFolder(folderId: string): ConversationSummary[] {
    if (folderId === 'all') return conversations;
    return conversations.filter(c => Array.isArray(c.folderIds) && c.folderIds.includes(folderId));
  }

  function handleClickOutside(e: MouseEvent) {
    if (!bulkMenuOpen) return;
    const target = e.target as Node;
    if (bulkAnchorEl && (bulkAnchorEl === target || bulkAnchorEl.contains(target))) return;
    const menu = bulkAnchorEl?.parentElement?.querySelector('.bulk-menu');
    if (menu && menu.contains(target)) return;
    bulkMenuOpen = false;
  }

  function relTime(iso?: string): string {
    if (!iso) return '';
    const then = Date.parse(iso);
    if (!Number.isFinite(then)) return '';
    const diffMs = Date.now() - then;
    if (diffMs < 0) return '';
    const mins = Math.floor(diffMs / 60_000);
    if (mins < 1) return t('picker.justNow', {}, lang) || 'now';
    if (mins < 60) return `${mins}m`;
    const hrs = Math.floor(mins / 60);
    if (hrs < 24) return `${hrs}h`;
    const days = Math.floor(hrs / 24);
    if (days < 7) return `${days}d`;
    return iso.slice(0, 10);
  }

  // Label for the list header. 'all' uses a generic string; the others
  // reuse the rail labels so translators only localise them once.
  $: currentKindLabel = (() => {
    const entry = RAIL_KINDS.find(k => k.kind === currentKind);
    return entry ? (t(entry.labelKey, {}, lang) || entry.fallback) : '';
  })();

  // When the user switches kinds, reset the filter — otherwise a
  // left-over query from "Chats" leaks into "Groups" and the list
  // looks empty for no obvious reason. Writes through the bound
  // selectedKind prop and dispatches kindChange so the caller can
  // persist the choice.
  function switchKind(k: RailKind) {
    if (k === currentKind) return;
    selectedKind = k;
    query = '';
    dispatch('kindChange', k);
  }

  // Folder toggle. Per design (no "All folders" affordance), clicking
  // the active folder again clears it back to 'all'. Clicking a
  // different folder switches to it. Either way we emit
  // `folderChange` so the caller can persist the choice.
  function switchFolder(id: string) {
    const next = (selectedFolderId === id && id !== 'all') ? 'all' : id;
    if (next === selectedFolderId) return;
    selectedFolderId = next;
    // Clear the query like switchKind does — left-over filters across
    // axes look like an empty list for no obvious reason.
    query = '';
    dispatch('folderChange', next);
  }

  // Composite label for the picker header. Shows "Chats · Work" when
  // both axes are constrained, just the kind label otherwise. Folder
  // name is the visible affordance for "you have a folder filter on".
  $: currentFolderLabel = (() => {
    if (selectedFolderId === 'all') return '';
    const f = folders.find(x => x.id === selectedFolderId);
    return f ? f.name : '';
  })();
  $: composedHeadLabel = currentFolderLabel
    ? `${currentKindLabel} · ${currentFolderLabel}`
    : currentKindLabel;
</script>

<svelte:window on:click={handleClickOutside} />

<section class="picker-section" data-lang={lang}>
  <div class="card picker-card" class:collapsed data-tour="picker">
    <div class="card-header">
      <div class="card-icon"><MessageSquare size={16} /></div>
      <h2 class="card-title">{t('picker.title', {}, lang) || 'Chat'}</h2>
      <!-- Right-side actions cluster — wrapped so the chevron always
           anchors to the same X position regardless of which optional
           middle element (loading hint / refresh button / count text)
           is currently rendered. The wrapper takes margin-left:auto;
           the chevron always sits as the wrapper's last child. -->
      <div class="picker-header-actions">
        {#if state === 'loading'}
          <span class="picker-header-hint">
            <RefreshCw size={12} class="spin" />
            {t('picker.loading', {}, lang) || 'Loading chats…'}
          </span>
        {:else if state === 'error'}
          <button type="button" class="picker-retry" on:click={() => dispatch('retry')} title={errorMessage}>
            <RefreshCw size={12} />
            {t('picker.retry', {}, lang) || 'Retry'}
          </button>
        {:else if state === 'ok' && refreshing}
          <span class="picker-header-hint">
            <RefreshCw size={12} class="spin" />
            {t('picker.refreshing', {}, lang) || 'Refreshing…'}
          </span>
        {:else if state === 'ok' && collapsed}
          <span class="picker-header-hint picker-header-hint--count">
            {selectedIds.length === 0
              ? t('picker.noSelection', {}, lang)
              : t('picker.nSelected', { n: selectedIds.length }, lang)}
          </span>
        {:else if state === 'ok'}
          <button
            type="button"
            class="picker-retry"
            on:click={() => dispatch('retry')}
            title={t('picker.refreshTitle', {}, lang) || 'Refresh chat list'}
          >
            <RefreshCw size={12} />
          </button>
        {/if}
        {#if !notOnTeams && state !== 'error'}
          <button
            type="button"
            class="picker-collapse-toggle"
            class:collapsed
            on:click={toggleCollapse}
            aria-expanded={!collapsed}
            aria-label={collapsed
              ? (t('picker.expand', {}, lang) || 'Expand picker')
              : (t('picker.collapse', {}, lang) || 'Collapse picker')}
            title={collapsed
              ? (t('picker.expand', {}, lang) || 'Expand picker')
              : (t('picker.collapse', {}, lang) || 'Collapse picker')}
          >
            <ChevronDown size={14} />
          </button>
        {/if}
      </div>
    </div>

    {#if notOnTeams}
      <div class="picker-error">
        <AlertCircle size={14} />
        <span>{t('errors.needsTeams', {}, lang) || 'Open the Teams web app tab first.'}</span>
      </div>
    {:else if state === 'error'}
      <!-- Surface the actual error inline (used to be tooltip-only on
           the retry button). When auto-inject misbehaves or the
           content script can't reach Teams' tokens, the generic
           "Could not load chats" copy hid the diagnostic detail
           users needed to report back. -->
      <div class="picker-error">
        <AlertCircle size={14} />
        <div class="picker-error-text">
          <div class="picker-error-headline">{t('picker.errorShort', {}, lang) || 'Could not load chats'}</div>
          {#if errorMessage}
            <div class="picker-error-detail">{errorMessage}</div>
          {/if}
        </div>
      </div>
    {:else if !collapsed}
      <div class="picker-body" class:dim={state === 'loading'} class:has-folders={folders.length > 0}>
        <div
          class="picker-rail"
          role="tablist"
          aria-label={t('picker.rail.label', {}, lang) || 'Filter by kind and folder'}
          data-tour="folder"
          bind:this={railEl}
          use:railMountedAction
        >
          <div class="rail-section-head">{t('picker.rail.sectionType', {}, lang) || 'Type'}</div>
          {#each RAIL_KINDS as rk (rk.kind)}
            {@const label = t(rk.labelKey, {}, lang) || rk.fallback}
            {@const active = rk.kind === currentKind}
            {@const dot = rk.kind !== 'all' && kindHasSelection[rk.kind as ConversationKind]}
            <button
              type="button"
              role="tab"
              aria-selected={active}
              class="rail-item"
              class:active
              on:click={() => switchKind(rk.kind)}
            >
              <span class="rail-ic">
                {#if rk.kind === 'all'}<List size={13} />{/if}
                {#if rk.kind === 'chat'}<MessageSquare size={13} />{/if}
                {#if rk.kind === 'group'}<Users size={13} />{/if}
                {#if rk.kind === 'meeting'}<Calendar size={13} />{/if}
                {#if rk.kind === 'channel'}<Hash size={13} />{/if}
              </span>
              <span class="rail-lbl">{label}</span>
              {#if dot}<span class="rail-dot" aria-hidden="true"></span>{/if}
            </button>
          {/each}

          {#if folders.length > 0}
            <div class="rail-divider"></div>
            <div class="rail-section-head">{t('picker.rail.sectionFolder', {}, lang) || 'Folder'}</div>
            {#each folders as f (f.id)}
              {@const fActive = selectedFolderId === f.id}
              <button
                type="button"
                role="tab"
                aria-selected={fActive}
                class="rail-item"
                class:active={fActive}
                on:click={() => switchFolder(f.id)}
                title={fActive ? (t('picker.rail.folderClearHint', {}, lang) || 'Click again to clear') : f.name}
              >
                <span class="rail-ic">
                  {#if f.kind === 'favorites'}<Star size={13} class="rail-star" />{:else}<Folder size={13} />{/if}
                </span>
                <span class="rail-lbl">{f.name}</span>
                <span class="rail-count">{f.count}</span>
              </button>
            {/each}
          {/if}
        </div>

        <div class="picker-main">
          <div class="picker-head">
            <span class="picker-head-label">{composedHeadLabel}</span>
            {#if mode === 'multi'}
              <!-- M: head pill is now a dropdown trigger. Click opens a
                   context-aware bulk-actions menu (Select/Deselect all
                   visible, Invert, per-kind / per-folder toggles when
                   those filters are active, Clear all). The total count
                   stays informative; the caret hints clickability. -->
              <button
                type="button"
                class="picker-head-pill bulk-trigger"
                class:zero={selectedIds.length === 0}
                class:open={bulkMenuOpen}
                bind:this={bulkAnchorEl}
                aria-haspopup="menu"
                aria-expanded={bulkMenuOpen}
                title={t('picker.bulk.menuTitle', {}, lang) || 'Bulk selection actions'}
                on:click|stopPropagation={() => (bulkMenuOpen = !bulkMenuOpen)}
              >
                <span>
                  {selectedIds.length === 0
                    ? (t('picker.noSelection', {}, lang) || 'None selected')
                    : (t('picker.nSelected', { n: selectedIds.length }, lang) || `${selectedIds.length} selected`)}
                </span>
                <ChevronDown size={11} strokeWidth={2.5} />
              </button>
              {#if bulkMenuOpen}
                <div class="bulk-menu" role="menu">
                  <button
                    type="button"
                    class="bulk-menu-item"
                    role="menuitem"
                    disabled={filtered.length === 0}
                    on:click={() => { toggleAll(filtered); bulkMenuOpen = false; }}
                  >
                    <span>
                      {allVisibleSelected
                        ? (t('picker.bulk.deselectAllVisible', {}, lang) || 'Deselect all visible')
                        : (t('picker.bulk.selectAllVisible', {}, lang) || 'Select all visible')}
                    </span>
                    <span class="bulk-menu-n">{filtered.length}</span>
                  </button>
                  <button
                    type="button"
                    class="bulk-menu-item"
                    role="menuitem"
                    disabled={filtered.length === 0}
                    on:click={() => { invertSelection(filtered); bulkMenuOpen = false; }}
                  >
                    <span>{t('picker.bulk.invertVisible', {}, lang) || 'Invert visible'}</span>
                    <span class="bulk-menu-n">{filtered.length}</span>
                  </button>
                  {#if currentKind !== 'all'}
                    {@const kindList = chatsOfKind(currentKind)}
                    <button
                      type="button"
                      class="bulk-menu-item bulk-menu-divider"
                      role="menuitem"
                      on:click={() => { toggleAll(kindList); bulkMenuOpen = false; }}
                    >
                      <span>
                        {t('picker.bulk.toggleAllKind', { kind: currentKindLabel }, lang)
                          || `Toggle all ${currentKindLabel}`}
                      </span>
                      <span class="bulk-menu-n">{kindList.length}</span>
                    </button>
                  {/if}
                  {#if selectedFolderId !== 'all' && currentFolderLabel}
                    {@const folderList = chatsInFolder(selectedFolderId)}
                    <button
                      type="button"
                      class="bulk-menu-item"
                      class:bulk-menu-divider={currentKind === 'all'}
                      role="menuitem"
                      on:click={() => { toggleAll(folderList); bulkMenuOpen = false; }}
                    >
                      <span>
                        {t('picker.bulk.toggleAllFolder', { name: currentFolderLabel }, lang)
                          || `Toggle all in ${currentFolderLabel}`}
                      </span>
                      <span class="bulk-menu-n">{folderList.length}</span>
                    </button>
                  {/if}
                  <button
                    type="button"
                    class="bulk-menu-item bulk-menu-divider bulk-menu-danger"
                    role="menuitem"
                    disabled={selectedIds.length === 0}
                    on:click={() => { clearAll(); bulkMenuOpen = false; }}
                  >
                    <span>{t('picker.bulk.clearAll', {}, lang) || 'Clear all'}</span>
                    <span class="bulk-menu-n">{selectedIds.length}</span>
                  </button>
                </div>
              {/if}
              <SavedGroups
                groups={savedGroups}
                selectionCount={selectedIds.length}
                {lang}
                on:save={(e) => dispatch('saveGroup', e.detail)}
                on:apply={(e) => applyGroup(e.detail)}
                on:remove={(e) => dispatch('removeGroup', e.detail)}
              />
            {:else}
              <span
                class="picker-head-pill"
                class:zero={visibleSelectedCount === 0}
              >
                {visibleSelectedCount === 0
                  ? (t('picker.noSelection', {}, lang) || 'None selected')
                  : (t('picker.nSelected', { n: visibleSelectedCount }, lang) || `${visibleSelectedCount} selected`)}
              </span>
            {/if}
          </div>

          {#if groupBanner}
            <div class="picker-group-banner" role="status">{groupBanner}</div>
          {/if}

          <div class="picker-search">
            <Search size={12} />
            <input
              type="text"
              placeholder={t('picker.filter', {}, lang) || 'Filter…'}
              bind:value={query}
              bind:this={searchInputEl}
            />
            <button
              type="button"
              class="picker-search-clear"
              disabled={!query}
              title={t('picker.clearFilter', {}, lang) || 'Clear'}
              on:click={clearQuery}
            >
              {t('picker.clearFilter', {}, lang) || 'Clear'}
            </button>
          </div>

          {#if mode === 'multi' && filtered.length > 0}
            <!-- N: persistent icon-only action bar. Fixed-width buttons
                 swap state without changing width. Acts on the current
                 view (kind + folder + search) — combines with filters
                 to do the bulk of bulk operations. -->
            <div class="picker-bulk-bar">
              <span class="picker-bulk-summary">
                {t('picker.bulk.summary', { selected: visibleSelectedCount, total: filtered.length }, lang)
                  || `${visibleSelectedCount} / ${filtered.length} visible selected`}
              </span>
              <button
                type="button"
                class="picker-bulk-btn"
                class:active={allVisibleSelected}
                title={allVisibleSelected
                  ? (t('picker.bulk.deselectAllVisible', {}, lang) || 'Deselect all visible')
                  : (t('picker.bulk.selectAllVisible', {}, lang) || 'Select all visible')}
                on:click={() => toggleAll(filtered)}
              >
                {#if allVisibleSelected}
                  <CheckSquare size={14} />
                {:else}
                  <Square size={14} />
                {/if}
              </button>
              <button
                type="button"
                class="picker-bulk-btn"
                disabled={!anyVisibleSelected && !allVisibleSelected}
                title={t('picker.bulk.invertVisible', {}, lang) || 'Invert visible'}
                on:click={() => invertSelection(filtered)}
              >
                <Repeat size={14} />
              </button>
              <button
                type="button"
                class="picker-bulk-btn picker-bulk-danger"
                disabled={selectedIds.length === 0}
                title={t('picker.bulk.clearAll', {}, lang) || 'Clear all'}
                on:click={() => clearAll()}
              >
                <Trash2 size={14} />
              </button>
            </div>
          {/if}

          <div class="picker-list" role="listbox" aria-multiselectable={mode === 'multi'} bind:this={listEl}>
            {#if conversations.length === 0 && state === 'loading'}
              <div class="picker-loading-row">
                <RefreshCw size={16} class="spin" />
                <span>{t('picker.firstLoad', {}, lang) || 'Loading conversations from Teams…'}</span>
              </div>
            {:else if conversations.length === 0 && state === 'ok'}
              <!-- An empty list on a Teams tab almost never means "no
                   chats" — every Teams account has at least the self-
                   chat (48:notes) Microsoft auto-creates. The empty
                   state therefore points at IDB-not-yet-populated:
                   Teams' SPA is loading but hasn't synced the chat
                   list yet. We surface that explicitly + offer a one-
                   click retry instead of the misleading 'No chats'. -->
              <div class="picker-still-loading">
                <RefreshCw size={16} />
                <span class="picker-still-loading-msg">
                  {t('picker.stillLoading', {}, lang) || 'Teams is still loading your chats — try refreshing in a few seconds.'}
                </span>
                <button
                  type="button"
                  class="picker-retry"
                  on:click={() => dispatch('retry')}
                >
                  {t('picker.stillLoading.retry', {}, lang) || 'Refresh'}
                </button>
              </div>
            {:else if filtered.length === 0}
              <div class="picker-empty-row">{t('picker.noMatch', {}, lang) || 'No matches'}</div>
            {:else}
              {#each filtered as c (c.id)}
                {@const Icon = iconFor(c.kind)}
                {@const sel = selectedSet.has(c.id)}
                {@const nameText = rowName(c)}
                {@const subText = rowSubtitle(c)}
                <button
                  type="button"
                  class="picker-row"
                  class:selected={sel}
                  role="option"
                  aria-selected={sel}
                  on:click={() => toggle(c.id)}
                >
                  <span class="picker-row-check" aria-hidden="true">
                    {#if sel}<Check size={10} strokeWidth={3} />{/if}
                  </span>
                  <span class="picker-row-icon"><Icon size={12} /></span>
                  <span class="picker-row-body">
                    <span class="picker-row-name">{nameText}</span>
                    {#if subText}
                      <span class="picker-row-sub">{subText}</span>
                    {/if}
                  </span>
                  <span class="picker-row-rel">{relTime(c.lastActivity)}</span>
                </button>
              {/each}
            {/if}
          </div>
        </div>
      </div>
    {/if}
  </div>
</section>

<style>
  .picker-card { position: relative; }
  .picker-header-hint {
    font-size: 10px;
    color: var(--color-subtle);
    font-weight: 500;
    margin-left: auto;
    display: inline-flex;
    align-items: center;
    gap: 4px;
  }
  .picker-retry {
    margin-left: auto;
    display: inline-flex;
    align-items: center;
    gap: 4px;
    border: 1px solid var(--color-border);
    background: var(--color-surface);
    color: var(--color-text);
    font: inherit;
    font-size: 11px;
    padding: 3px 8px;
    border-radius: 6px;
    cursor: pointer;
  }
  .picker-retry:hover { border-color: var(--color-accent); color: var(--color-accent); }

  /* Right-side actions cluster — anchored to the right edge of the
     card-header via margin-left:auto. Containing the chevron in this
     wrapper means the chevron's X position stays constant no matter
     what the middle slot contains (loading hint, refresh button,
     count text), so the header doesn't visibly shift when the user
     toggles collapse / expand. */
  .picker-header-actions {
    margin-left: auto;
    display: inline-flex;
    align-items: center;
    gap: 6px;
    flex-shrink: 0;
  }
  /* Inside the actions wrapper, the existing .picker-retry and
     .picker-header-hint don't need their own margin-left:auto — the
     wrapper's auto margin already pushes the whole cluster right.
     Override the global margin-left:auto so all middle elements sit
     immediately next to each other inside the wrapper. */
  .picker-header-actions :global(.picker-retry),
  .picker-header-actions .picker-header-hint {
    margin-left: 0;
  }

  /* Collapse / expand chevron — always last child of the actions
     cluster, so its X position is constant. When expanded the caret
     points up (rotated 180°); collapsed it points down. */
  .picker-collapse-toggle {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 22px;
    height: 22px;
    border: 1px solid var(--color-border);
    border-radius: 6px;
    background: var(--color-surface);
    color: var(--color-subtle);
    cursor: pointer;
    padding: 0;
    flex-shrink: 0;
  }
  .picker-collapse-toggle:hover {
    border-color: var(--color-accent);
    color: var(--color-accent);
    background: var(--color-accent-light);
  }
  .picker-collapse-toggle :global(svg) {
    transition: transform 0.18s ease;
    transform: rotate(180deg);  /* expanded → caret up */
  }
  .picker-collapse-toggle.collapsed :global(svg) {
    transform: rotate(0deg);    /* collapsed → caret down */
  }
  /* Selection-count hint shown only while collapsed, so the user can
     read "1 selected" / "3 selected" without expanding. */
  .picker-header-hint--count {
    color: var(--color-subtle);
    font-size: 11px;
    font-weight: 500;
  }
  /* Collapsed card keeps the standard .card padding (12px on all
     sides) so the title and chevron stay at the SAME pixel position
     they occupy when expanded — the only thing that changes between
     the two states is whether the body renders. The card-header's
     11px bottom margin still gets zeroed when the body is hidden so
     the card collapses to header height with no wasted gap. */
  .picker-card.collapsed :global(.card-header) {
    margin-bottom: 0;
  }
  .picker-error {
    display: flex;
    align-items: flex-start;
    gap: 8px;
    /* margin-top separates this from the card-header above. Without
       it the "Open the Teams web app tab first." line sat flush
       against the "Conversations" title — visually crowded. */
    margin-top: 10px;
    padding: 14px 12px;
    background: var(--color-bg);
    border: 1px solid var(--color-border);
    border-radius: var(--radius-md);
    color: var(--color-subtle);
    font-size: 12px;
  }
  .picker-error-text {
    display: flex;
    flex-direction: column;
    gap: 3px;
    min-width: 0;
    flex: 1;
  }
  .picker-error-headline { font-weight: 500; color: var(--color-text); }
  .picker-error-detail {
    font-family: ui-monospace, SFMono-Regular, Menlo, monospace;
    font-size: 10.5px;
    color: var(--color-subtle);
    word-break: break-word;
    line-height: 1.4;
  }

  .picker-body {
    display: grid;
    grid-template-columns: 100px 1fr;
    height: 300px;
    background: var(--color-surface);
    border: 1px solid var(--color-border);
    border-radius: var(--radius-md);
    overflow: hidden;
  }
  .picker-body.dim { opacity: 0.55; pointer-events: none; }

  /* Vertical rail with single-line items (icon + label + count). One
     section for kinds, optional second section for folders, separated
     by a thin divider with a tiny uppercase header on each. The rail
     scrolls vertically when content overflows; horizontal scroll is
     explicitly suppressed because content widths can flutter and a
     spurious horizontal scrollbar is the most reliable way to make
     the popup feel broken. */
  .picker-rail {
    background: var(--color-bg);
    border-right: 1px solid var(--color-border);
    display: flex;
    flex-direction: column;
    padding: 4px 0;
    overflow-y: auto;
    overflow-x: hidden;
    scrollbar-width: thin;
    scrollbar-color: var(--color-border-hover) transparent;
  }
  .picker-rail::-webkit-scrollbar { width: 5px; }
  .picker-rail::-webkit-scrollbar-track { background: transparent; }
  .picker-rail::-webkit-scrollbar-thumb { background: var(--color-border-hover); border-radius: 3px; }

  .rail-section-head {
    font-size: 9px;
    font-weight: 700;
    color: var(--color-subtle);
    text-transform: uppercase;
    letter-spacing: 0.06em;
    padding: 8px 11px 3px;
    opacity: 0.7;
    user-select: none;
  }
  .rail-divider {
    height: 1px;
    margin: 6px 8px;
    background: var(--color-border);
  }

  .rail-item {
    display: flex;
    align-items: center;
    gap: 8px;
    width: calc(100% - 6px);
    margin: 1px 3px;
    padding: 6px 8px;
    border-radius: 5px;
    cursor: pointer;
    color: var(--color-subtle);
    background: transparent;
    border: 0;
    font: inherit;
    font-size: 11px;
    font-weight: 500;
    text-align: left;
    position: relative;
    transition: background 0.15s, color 0.15s;
  }
  .rail-item:hover { background: var(--color-accent-light); color: var(--color-accent); }
  .rail-item.active { background: var(--color-accent-light); color: var(--color-accent); font-weight: 600; }
  .rail-item.active::before {
    content: '';
    position: absolute;
    left: -3px;
    top: 5px;
    bottom: 5px;
    width: 3px;
    background: var(--color-accent);
    border-radius: 0 2px 2px 0;
  }
  /* One-shot "look here" flash applied right after the picker auto-
     scrolls the active folder into view on open / re-expand. The active
     row briefly saturates to the full accent colour (white-on-blue),
     then animates back to its normal active state. The animation
     class is removed by JS after the keyframe finishes so it can
     replay on the next open.
     :global() because the class is added imperatively (not in the
     template), so Svelte's component-scoped CSS would otherwise prune
     the selector as unused. */
  .rail-item:global(.rail-item-flash) {
    animation: rail-item-flash 1.4s ease-out 1;
  }
  @keyframes rail-item-flash {
    0%   { background: var(--color-accent); color: #fff; }
    100% { background: var(--color-accent-light); color: var(--color-accent); }
  }
  /* The white-text variant needs the count colour to flip too, so the
     "5" or "3" next to the folder name doesn't sit invisibly on its
     accent-coloured background during the first frame. */
  .rail-item:global(.rail-item-flash) .rail-count {
    animation: rail-item-flash-count 1.4s ease-out 1;
  }
  @keyframes rail-item-flash-count {
    0%   { color: rgba(255, 255, 255, 0.9); }
    100% { color: var(--color-accent); }
  }
  .rail-ic {
    flex: 0 0 14px;
    width: 14px;
    height: 14px;
    display: inline-flex;
    align-items: center;
    justify-content: center;
  }
  .rail-lbl {
    flex: 1;
    overflow: hidden;
    white-space: nowrap;
    text-overflow: ellipsis;
  }
  .rail-count {
    font-size: 9px;
    opacity: 0.6;
    font-variant-numeric: tabular-nums;
    flex-shrink: 0;
  }
  .rail-item :global(.rail-star) { color: #f59e0b; }
  .rail-item.active :global(.rail-star) { color: #f59e0b; }
  .rail-dot {
    position: absolute;
    top: 6px; right: 6px;
    width: 6px; height: 6px;
    border-radius: 50%;
    background: var(--color-accent);
    border: 1.5px solid var(--color-bg);
  }

  .picker-main {
    display: flex;
    flex-direction: column;
    overflow: hidden;
    min-width: 0;
  }
  .picker-head {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 8px;
    padding: 9px 12px;
    border-bottom: 1px solid var(--color-border);
  }
  .picker-head-label {
    font-size: 13px;
    font-weight: 600;
    color: var(--color-text);
    flex: 1;
    min-width: 0;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }
  .picker-head-pill {
    font-size: 10px;
    padding: 2px 8px;
    border-radius: 999px;
    background: var(--color-accent-light);
    color: var(--color-accent);
    font-weight: 500;
    white-space: nowrap;
  }
  .picker-head-pill.zero {
    color: var(--color-subtle);
    background: var(--color-border);
  }

  /* M: head pill becomes a dropdown trigger when in multi mode. Looks
     identical to the static pill in normal state; gets a subtle accent
     ring on hover and an inverted look while open so the menu's anchor
     stays obvious. */
  .picker-head-pill.bulk-trigger {
    display: inline-flex;
    align-items: center;
    gap: 3px;
    border: 0;
    cursor: pointer;
    font: inherit;
    font-size: 10px;
    font-weight: 500;
    transition: background 0.15s, color 0.15s, box-shadow 0.15s;
  }
  .picker-head-pill.bulk-trigger:hover {
    box-shadow: 0 0 0 1px var(--color-accent) inset;
  }
  .picker-head-pill.bulk-trigger.open {
    background: var(--color-accent);
    color: white;
  }
  .picker-head-pill.bulk-trigger.zero.open {
    background: var(--color-text);
    color: var(--color-bg);
  }

  /* Dropdown popover positioned below the head-pill. Picker-head is
     already position:relative implicitly (flex parent); the menu is
     position:absolute and sits in the same row. We anchor to the
     right because the pill sits at the right end of the head row. */
  .picker-head { position: relative; }
  .bulk-menu {
    position: absolute;
    top: calc(100% + 4px);
    right: 8px;
    z-index: 20;
    min-width: 200px;
    background: var(--color-surface);
    border: 1px solid var(--color-border);
    border-radius: 8px;
    padding: 4px;
    box-shadow: 0 4px 14px rgba(0, 0, 0, 0.12);
  }
  .bulk-menu-item {
    width: 100%;
    text-align: left;
    border: 0;
    background: transparent;
    color: var(--color-text);
    font: inherit;
    font-size: 12px;
    padding: 6px 10px;
    border-radius: 5px;
    cursor: pointer;
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 12px;
  }
  .bulk-menu-item:hover { background: var(--color-accent-light); color: var(--color-accent); }
  .bulk-menu-item:disabled { opacity: 0.4; cursor: not-allowed; }
  .bulk-menu-item:disabled:hover { background: transparent; color: var(--color-text); }
  .bulk-menu-item.bulk-menu-divider {
    border-top: 1px solid var(--color-border);
    margin-top: 4px;
    padding-top: 8px;
  }
  .bulk-menu-item.bulk-menu-danger:hover {
    background: rgba(220, 38, 38, 0.1);
    color: #dc2626;
  }
  .bulk-menu-n {
    font-size: 10px;
    opacity: 0.7;
    font-variant-numeric: tabular-nums;
    flex-shrink: 0;
  }

  /* N: persistent icon-only action bar between search and list. Buttons
     are 24×24 fixed — only the icon swaps when state flips, so the row
     never reflows. Tooltips name each action so the icons don't have
     to carry the full meaning. */
  .picker-bulk-bar {
    display: flex;
    align-items: center;
    gap: 4px;
    padding: 5px 10px;
    border-bottom: 1px solid var(--color-border);
    background: var(--color-bg);
  }
  .picker-bulk-summary {
    flex: 1;
    font-size: 10px;
    color: var(--color-subtle);
    font-weight: 500;
    overflow: hidden;
    white-space: nowrap;
    text-overflow: ellipsis;
  }
  .picker-bulk-btn {
    width: 24px;
    height: 24px;
    border: 1px solid var(--color-border);
    background: var(--color-surface);
    color: var(--color-text);
    border-radius: 5px;
    cursor: pointer;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    flex-shrink: 0;
    padding: 0;
    transition: border-color 0.15s, color 0.15s, background 0.15s;
  }
  .picker-bulk-btn:hover {
    border-color: var(--color-accent);
    color: var(--color-accent);
    background: var(--color-accent-light);
  }
  .picker-bulk-btn.active {
    background: var(--color-accent);
    border-color: var(--color-accent);
    color: white;
  }
  .picker-bulk-btn.active:hover {
    background: var(--color-accent-hover);
    border-color: var(--color-accent-hover);
    color: white;
  }
  .picker-bulk-btn:disabled {
    opacity: 0.4;
    cursor: not-allowed;
  }
  .picker-bulk-btn:disabled:hover {
    border-color: var(--color-border);
    color: var(--color-text);
    background: var(--color-surface);
  }
  .picker-bulk-btn.picker-bulk-danger:hover {
    border-color: #dc2626;
    color: #dc2626;
    background: rgba(220, 38, 38, 0.1);
  }
  .picker-bulk-btn.picker-bulk-danger:disabled:hover {
    border-color: var(--color-border);
    color: var(--color-text);
    background: var(--color-surface);
  }

  .picker-search {
    display: flex;
    align-items: center;
    gap: 6px;
    padding: 6px 10px;
    border-bottom: 1px solid var(--color-border);
    color: var(--color-subtle);
  }
  .picker-search input {
    flex: 1;
    border: 0;
    background: transparent;
    color: var(--color-text);
    font: inherit;
    font-size: 12px;
    outline: none;
    min-width: 0;
  }

  .picker-list {
    flex: 1;
    overflow-y: auto;
    padding: 2px 0;
  }
  .picker-row {
    width: 100%;
    display: grid;
    grid-template-columns: 18px 14px 1fr auto;
    gap: 8px;
    align-items: center;
    padding: 6px 10px;
    border: 0;
    background: transparent;
    color: var(--color-text);
    font: inherit;
    font-size: 12px;
    text-align: left;
    cursor: pointer;
    min-width: 0;
  }
  .picker-row:hover { background: var(--color-accent-light); }
  .picker-row.selected { background: var(--color-accent-light); }
  .picker-row-check {
    width: 14px;
    height: 14px;
    border: 1.5px solid var(--color-border-hover);
    border-radius: 4px;
    display: flex;
    align-items: center;
    justify-content: center;
    background: var(--color-surface);
    color: white;
    flex-shrink: 0;
  }
  .picker-row.selected .picker-row-check {
    background: var(--color-accent);
    border-color: var(--color-accent);
  }
  .picker-row-icon {
    color: var(--color-subtle);
    display: inline-flex;
    justify-content: center;
  }
  .picker-row.selected .picker-row-icon { color: var(--color-accent); }
  .picker-row-body {
    display: flex;
    flex-direction: column;
    min-width: 0;
    gap: 1px;
  }
  .picker-row-name {
    overflow: hidden;
    white-space: nowrap;
    text-overflow: ellipsis;
  }
  .picker-row.selected .picker-row-name { font-weight: 500; color: var(--color-accent); }
  .picker-row-sub {
    font-size: 10px;
    color: var(--color-subtle);
    overflow: hidden;
    white-space: nowrap;
    text-overflow: ellipsis;
  }
  .picker-row-rel {
    font-size: 10px;
    color: var(--color-subtle);
    font-variant-numeric: tabular-nums;
    flex-shrink: 0;
  }
  .picker-empty-row {
    padding: 16px 14px;
    color: var(--color-subtle);
    font-size: 12px;
    text-align: center;
  }
  .picker-loading-row {
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 8px;
    padding: 24px 14px;
    color: var(--color-subtle);
    font-size: 12px;
  }
  /* "Teams is still loading" empty-state. Stacked layout (icon + msg
     + action) instead of the inline "No chats found" we used to
     show, because the message is longer and we want the Refresh
     button as a clearly clickable affordance. */
  .picker-still-loading {
    display: flex;
    flex-direction: column;
    align-items: center;
    gap: 10px;
    padding: 22px 16px;
    color: var(--color-subtle);
    font-size: 12px;
    text-align: center;
  }
  .picker-still-loading-msg { line-height: 1.45; }
  .picker-still-loading .picker-retry {
    /* Match the existing .picker-retry visual but slightly larger so
       it stands alone, since this empty state isn't squeezed next to
       the "Loaded N" label like the inline retry is. */
    padding: 6px 14px;
    font-size: 12px;
  }

  :global(.spin) { animation: picker-spin 1s linear infinite; }
  @keyframes picker-spin { to { transform: rotate(360deg); } }
  /* Transient banner shown when applying a saved group that referenced chats
     not present in this account. */
  .picker-group-banner {
    padding: 6px 12px;
    font-size: 11px;
    background: #fef3c7;
    color: #92400e;
    border-bottom: 1px solid var(--color-border);
  }
  /* Search filter clear: always visible, grayed (disabled) until there's text. */
  .picker-search-clear {
    border: 0;
    background: transparent;
    color: var(--color-accent);
    font: inherit;
    font-size: 11px;
    font-weight: 500;
    cursor: pointer;
    padding: 2px 4px;
    flex: 0 0 auto;
    transition: opacity 0.15s ease;
  }
  .picker-search-clear:disabled {
    opacity: 0.4;
    cursor: default;
    color: var(--color-subtle);
  }
</style>
