<script lang="ts">
  import { createEventDispatcher } from 'svelte';
  import { Download, Square } from 'lucide-svelte';
  import { t } from '../../../i18n/i18n';

  // The button is the popup's single status surface. Idle: shows the export
  // action. Busy: turns red, becomes a Stop button, and the bottom of the
  // button hosts a 4-segment phase tracker (messages · images · people · file).
  // After complete or cancel, the right zone keeps the result visible (sticky)
  // until the user clicks the button again.

  export let disabled = false;
  export let busy = false;
  export let lang = 'en';
  // Idle subtitle (e.g. options summary like "Chat · JSON · with replies")
  export let summary = '';

  // Phase + progress (passed from App.svelte while busy).
  export let phaseLabel = '';            // e.g. "Fetching messages · 0:12"
  export let counterValue = '—';         // e.g. "1,400" or "5 / 98"
  export let counterLabel = '';          // e.g. "messages" / "images" / "people"
  // Segments correspond to: messages, images, people, file.
  // -1 = indeterminate (animated stripe). 0..100 = filled %.
  // null = not yet started (dim).
  export let segments: Array<number | null> = [null, null, null, null];

  // Sticky outcome: shown in the right zone after complete or cancel until
  // the user starts a new export.
  export let outcome: null | {
    kind: 'success' | 'cancelled';
    primary: string;          // big text e.g. "✓ 12,591" or "✕ 0:18"
    secondary: string;        // small text e.g. "saved · 0:48" / "cancelled"
  } = null;

  const dispatch = createEventDispatcher<{
    run: void;
    stop: void;
  }>();

  function handleClick() {
    if (busy) {
      dispatch('stop');
      return;
    }
    if (!disabled) {
      dispatch('run');
    }
  }

  $: idleLabel = t('actions.export', {}, lang);
  $: stopLabel = t('actions.stop', {}, lang) || 'Stop export';

  // Subtitle shown in the left "detail" line. While busy, prefer the phase
  // label; otherwise (idle / outcome state) fall back to the options summary.
  $: detailText = busy ? phaseLabel : summary;

  // Map a segment value to a styling state.
  function segClass(v: number | null) {
    if (v == null) return '';
    if (v < 0) return 'active';   // indeterminate stripe
    return '';
  }
  function segWidth(v: number | null) {
    if (v == null || v < 0) return '0%';
    return `${Math.max(0, Math.min(100, v))}%`;
  }
</script>

<button
  class="export-primary"
  class:busy
  class:has-outcome={!busy && outcome}
  disabled={!busy && disabled}
  title={busy ? stopLabel : ''}
  on:click={handleClick}
>
  <div class="left">
    <div class="icon">
      {#if busy}
        <Square size={20} />
      {:else}
        <Download size={22} />
      {/if}
    </div>
    <div class="text">
      <span class="title">{busy ? stopLabel : idleLabel}</span>
      {#if detailText}
        <span class="detail">{detailText}</span>
      {/if}
    </div>
  </div>

  {#if busy || outcome}
    <div class="right" class:sticky={outcome}>
      {#if outcome}
        <span class="v">{outcome.primary}</span>
        <span class="l">{outcome.secondary}</span>
      {:else}
        <span class="v">{counterValue}</span>
        {#if counterLabel}
          <span class="l">{counterLabel}</span>
        {/if}
      {/if}
    </div>
  {/if}

  {#if busy}
    <div class="phase-track" aria-hidden="true">
      {#each segments as seg}
        <div class="seg {segClass(seg)}">
          <div class="seg-fill" style:width={segWidth(seg)}></div>
        </div>
      {/each}
    </div>
  {/if}
</button>

<style>
  .export-primary {
    width: 100%;
    box-sizing: border-box;
    border: none;
    border-radius: 10px;
    cursor: pointer;
    color: #fff;
    font: inherit;
    display: flex;
    align-items: stretch;
    overflow: hidden;
    position: relative;
    transition: background 0.2s ease;
    margin-bottom: 12px;
  }
  .left {
    flex: 1; min-width: 0;
    padding: 14px;
    display: flex; align-items: center; gap: 10px;
  }
  .icon {
    flex-shrink: 0;
    width: 22px; height: 22px;
    display: flex; align-items: center; justify-content: center;
  }
  .text { flex: 1; min-width: 0; }
  .title { font-weight: 600; font-size: 14px; display: block; }
  .detail {
    display: block;
    font-size: 11px;
    opacity: 0.92;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
  }
  .right {
    width: 96px;
    border-left: 1px solid rgba(255, 255, 255, 0.22);
    padding: 10px 8px;
    display: flex; flex-direction: column; align-items: center; justify-content: center;
    font-size: 11px;
    text-align: center;
    gap: 2px;
    flex-shrink: 0;
  }
  .right.sticky { background: rgba(0, 0, 0, 0.20); }
  .right .v { font-size: 14px; font-weight: 600; line-height: 1.1; }
  .right .l { opacity: 0.85; font-size: 9px; text-transform: uppercase; letter-spacing: 0.05em; }

  .phase-track {
    position: absolute;
    left: 0; right: 0; bottom: 0;
    height: 3px;
    display: flex;
    gap: 2px;
    padding: 0 2px;
  }
  .seg {
    flex: 1;
    height: 100%;
    background: rgba(255, 255, 255, 0.22);
    border-radius: 2px;
    overflow: hidden;
    position: relative;
  }
  .seg > .seg-fill {
    height: 100%;
    background: rgba(255, 255, 255, 0.85);
    transition: width 0.3s ease;
  }
  /* Indeterminate: animated highlight scrolling across the segment. */
  .seg.active > .seg-fill {
    width: 100%;
    background: linear-gradient(90deg,
      rgba(255, 255, 255, 0.4) 0%,
      rgba(255, 255, 255, 0.95) 50%,
      rgba(255, 255, 255, 0.4) 100%);
    background-size: 200% 100%;
    animation: vc-stripe 1.4s linear infinite;
  }
  @keyframes vc-stripe {
    0%   { background-position: 100% 0; }
    100% { background-position: -100% 0; }
  }

  .export-primary:disabled {
    opacity: 0.6;
    cursor: not-allowed;
  }
</style>
