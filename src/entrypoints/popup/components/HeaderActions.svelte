<script lang="ts">
  import { createEventDispatcher, tick } from 'svelte';
  import { Settings, History } from 'lucide-svelte';
  import { t } from '../../../i18n/i18n';

  // Number of history entries added since the user last opened the
  // History page. Rendered as a small notification badge on the history
  // icon (capped to "9+" for readability). Clears the moment the user
  // opens the page (parent updates lastHistoryViewedAt).
  export let newHistoryCount = 0;
  $: badgeText = newHistoryCount > 9 ? '9+' : String(newHistoryCount);

  // One-shot pulse trigger. When the parent increments this counter (after
  // a phase=complete or phase=cancelled), the icon does a brief scale +
  // tint pulse to direct the eye. Same counter pattern as ExportButton's
  // flashTrigger.
  export let pulseHistoryIcon = 0;
  export let lang = 'en';

  const dispatch = createEventDispatcher<{
    openSettings: void;
    openHistory: void;
  }>();

  let pulseActive = false;
  let lastPulseSeen = 0;
  $: if (pulseHistoryIcon !== lastPulseSeen) {
    lastPulseSeen = pulseHistoryIcon;
    void replayPulse();
  }
  async function replayPulse() {
    pulseActive = false;
    await tick();
    pulseActive = true;
    setTimeout(() => { pulseActive = false; }, 700);
  }
</script>

<div class="header-actions">
  <button
    class="icon-btn"
    class:pulse={pulseActive}
    title={t('history.title', {}, lang) || 'Export history'}
    on:click={() => dispatch('openHistory')}
  >
    <History size={18} />
    {#if newHistoryCount > 0}
      <span class="new-badge" aria-label="{newHistoryCount} new">{badgeText}</span>
    {/if}
  </button>
  <button
    class="icon-btn"
    title={t('settings.title', {}, lang) || 'Settings'}
    on:click={() => dispatch('openSettings')}
  >
    <Settings size={18} />
  </button>
</div>

<style>
  .header-actions {
    display: flex;
    align-items: center;
    gap: 4px;
  }
  .icon-btn {
    position: relative;
    width: 28px; height: 28px;
    display: flex; align-items: center; justify-content: center;
    border: none;
    background: transparent;
    color: var(--color-text-muted, #64748b);
    border-radius: 6px;
    cursor: pointer;
    transition: background 0.15s ease, color 0.15s ease;
  }
  .icon-btn:hover {
    background: var(--color-accent-soft, #eff6ff);
    color: var(--color-accent, #2563eb);
  }
  /* One-shot pulse — replays each time the parent bumps `pulseHistoryIcon`. */
  .icon-btn.pulse {
    animation: history-pulse 0.7s ease-out 1;
  }
  @keyframes history-pulse {
    0%   { transform: scale(1); }
    40%  { transform: scale(1.18); background: rgba(22, 163, 74, 0.12); color: #16a34a; }
    100% { transform: scale(1); }
  }

  /* Notification-badge style — small pill with the unread count. Sits in
     the icon's top-right corner. Capped to "9+" by the parent computation
     so it never overflows. */
  .new-badge {
    position: absolute;
    top: -3px; right: -3px;
    min-width: 16px; height: 16px;
    padding: 0 4px;
    border-radius: 999px;
    background: var(--color-accent, #2563eb);
    color: #fff;
    font-size: 10px; font-weight: 700;
    line-height: 16px;
    text-align: center;
    box-shadow: 0 0 0 2px var(--color-surface, #fff);
    pointer-events: none;
  }
</style>
