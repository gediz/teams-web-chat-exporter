<script lang="ts">
  import { createEventDispatcher, tick } from 'svelte';
  import { Download, Square } from 'lucide-svelte';
  import { t } from '../../../i18n/i18n';

  // The button has exactly two visible states now:
  //   Idle  — Download icon, "Export current chat", summary on detail line.
  //   Busy  — red, Square icon, "Stop export", phase tracker animating along
  //           the bottom edge, counter on the right.
  //
  // Post-export outcomes live in the History page (see HistoryPage.svelte).
  // Success and cancellation are signalled to the user by:
  //   - one-shot green flash on this button (via `flashTrigger` prop bump)
  //   - one-shot pulse on the history icon (handled by HeaderActions)
  //   - persistent "new" dot on the history icon
  // After cancel/complete, the button immediately returns to its idle state.

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

  // Increment-counter trigger for the success flash. When the parent
  // increments this value, we briefly toggle the `success-flash` class so
  // the CSS animation re-fires. The counter pattern (vs a boolean) means
  // the animation can fire repeatedly without the parent having to reset
  // anything.
  export let flashTrigger = 0;

  const dispatch = createEventDispatcher<{ run: void; stop: void; }>();

  function handleClick() {
    if (busy) { dispatch('stop'); return; }
    if (!disabled) { dispatch('run'); }
  }

  $: idleLabel = t('actions.export', {}, lang);
  $: stopLabel = t('actions.stop', {}, lang) || 'Stop export';

  // Subtitle: phase label while busy, options summary while idle.
  $: detailText = busy ? phaseLabel : summary;

  // One-shot flash. When `flashTrigger` changes (parent bumps the counter),
  // we momentarily clear then set the class so the CSS keyframes re-run.
  let flashActive = false;
  let lastFlashSeen = 0;
  $: if (flashTrigger !== lastFlashSeen) {
    lastFlashSeen = flashTrigger;
    void replayFlash();
  }
  async function replayFlash() {
    flashActive = false;
    await tick();
    flashActive = true;
    setTimeout(() => { flashActive = false; }, 700);
  }

  function segClass(v: number | null) {
    if (v == null) return '';
    if (v < 0) return 'active';
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
  class:success-flash={flashActive}
  type="button"
  disabled={!busy && disabled}
  title={busy ? stopLabel : ''}
  on:click={handleClick}
>
  <div class="left">
    <div class="icon">
      {#if busy}<Square size={20} />{:else}<Download size={22} />{/if}
    </div>
    <div class="text">
      <span class="title">
        {#if busy}{stopLabel}{:else}{idleLabel}{/if}
      </span>
      {#if detailText}<span class="detail">{detailText}</span>{/if}
    </div>
  </div>

  {#if busy}
    <div class="right">
      <span class="v">{counterValue}</span>
      {#if counterLabel}<span class="l">{counterLabel}</span>{/if}
    </div>

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
    color: #fff;
    font: inherit;
    display: flex;
    align-items: stretch;
    overflow: hidden;
    position: relative;
    margin-bottom: 12px;
    cursor: pointer;
    transition: background 0.2s ease;
  }
  .export-primary:disabled {
    opacity: 0.6;
    cursor: not-allowed;
  }

  /* Success flash — one-shot green pulse on the whole bar. The animation
     class is briefly removed/re-added by the script so it can replay on
     subsequent successful exports. ~700ms total. */
  .export-primary.success-flash {
    animation: btn-success-flash 0.7s ease-out 1;
  }
  @keyframes btn-success-flash {
    0%   { background: var(--color-accent); }
    35%  { background: #16a34a; box-shadow: 0 0 0 4px rgba(22, 163, 74, 0.18); }
    100% { background: var(--color-accent); box-shadow: 0 0 0 0 rgba(22, 163, 74, 0); }
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
    font-size: 18px;
    font-weight: 700;
  }
  .text { flex: 1; min-width: 0; }
  .title { font-weight: 600; font-size: 14px; display: block; }
  .detail {
    display: block;
    font-size: 11px;
    opacity: 0.92;
    white-space: nowrap;
    overflow: hidden;
    /* Project convention: clip silently, no ellipsis. */
    text-overflow: clip;
  }

  /* Right zone: busy counter only (cancelled tile is gone). */
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
</style>
