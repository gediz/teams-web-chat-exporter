<script lang="ts">
  import { createEventDispatcher } from 'svelte';
  import { X } from 'lucide-svelte';
  import { t } from '../../../i18n/i18n';

  export let lang = 'en';

  const dispatch = createEventDispatcher<{ dismiss: void }>();

  // Two steps — features the user is most likely to miss at a glance:
  //   1. Multi-format + bundle.zip (brand new in this release)
  //   2. Stop button (runtime-only affordance; easy to miss pre-export)
  // Keep copy short; the overlay scrims the UI behind it, so long prose
  // would feel intrusive.
  let step: 1 | 2 = 1;

  const next = () => { step = 2; };
  const back = () => { step = 1; };
  const finish = () => dispatch('dismiss');
</script>

<!-- Full-popup scrim + centered card. No spotlight / cutout in v1 —
     the popup is only ~420px tall, so a centered modal already makes
     the right elements visible "behind" the dimmed scrim. Upgrading to
     a real SVG spotlight is deferred until there's a clear UX need. -->
<div class="onb-scrim" role="dialog" aria-modal="true" aria-labelledby="onb-title" data-lang={lang}>
  <div class="onb-card">
    <button
      type="button"
      class="onb-skip"
      on:click={finish}
      title={t('onboarding.skip', {}, lang) || 'Skip'}
      aria-label={t('onboarding.skip', {}, lang) || 'Skip'}
    >
      <X size={14} />
    </button>

    {#if step === 1}
      <div class="onb-step">1 / 2</div>
      <h2 id="onb-title" class="onb-title">
        {t('onboarding.step1.title', {}, lang) || 'Pick one or more formats'}
      </h2>
      <p class="onb-body">
        {t('onboarding.step1.body', {}, lang) || 'Tap any of the format cards to toggle it. When two or more are selected, the exports are packaged together as a single .zip.'}
      </p>
    {:else}
      <div class="onb-step">2 / 2</div>
      <h2 id="onb-title" class="onb-title">
        {t('onboarding.step2.title', {}, lang) || 'Export is reversible'}
      </h2>
      <p class="onb-body">
        {t('onboarding.step2.body', {}, lang) || 'Click Start to begin. During an export the button turns red — click it again any time to stop. The thin bar at the bottom shows progress.'}
      </p>
    {/if}

    <div class="onb-actions">
      {#if step === 2}
        <button type="button" class="onb-btn secondary" on:click={back}>
          {t('onboarding.back', {}, lang) || 'Back'}
        </button>
      {/if}
      {#if step === 1}
        <button type="button" class="onb-btn primary" on:click={next}>
          {t('onboarding.next', {}, lang) || 'Next'}
        </button>
      {:else}
        <button type="button" class="onb-btn primary" on:click={finish}>
          {t('onboarding.gotIt', {}, lang) || 'Got it'}
        </button>
      {/if}
    </div>
  </div>
</div>

<style>
  .onb-scrim {
    position: fixed;
    inset: 0;
    background: rgba(15, 23, 42, 0.55);
    backdrop-filter: blur(2px);
    -webkit-backdrop-filter: blur(2px);
    display: flex;
    align-items: center;
    justify-content: center;
    z-index: 1000;
    /* Popup animates itself in; skip any entry animation here so the
       first render doesn't look jittery. */
  }

  .onb-card {
    width: calc(100% - 40px);
    max-width: 320px;
    background: var(--color-bg);
    color: var(--color-text);
    border-radius: 12px;
    padding: 20px 18px 16px;
    box-shadow: 0 20px 40px -12px rgba(0, 0, 0, 0.45), 0 0 0 1px var(--color-border);
    position: relative;
  }

  .onb-skip {
    position: absolute;
    top: 8px;
    right: 8px;
    background: transparent;
    border: none;
    color: var(--color-muted);
    cursor: pointer;
    padding: 4px;
    border-radius: 6px;
    display: flex;
    align-items: center;
    justify-content: center;
  }
  .onb-skip:hover { background: var(--color-accent-light); color: var(--color-text); }

  .onb-step {
    font-size: 10px;
    font-weight: 600;
    color: var(--color-accent);
    letter-spacing: 0.04em;
    text-transform: uppercase;
    margin-bottom: 6px;
  }

  .onb-title {
    font-size: 15px;
    font-weight: 600;
    margin: 0 0 8px;
    line-height: 1.3;
  }

  .onb-body {
    font-size: 12px;
    line-height: 1.5;
    color: var(--color-muted);
    margin: 0 0 14px;
  }

  .onb-actions {
    display: flex;
    gap: 8px;
    justify-content: flex-end;
  }

  .onb-btn {
    font: inherit;
    font-size: 12px;
    font-weight: 600;
    padding: 6px 12px;
    border-radius: 6px;
    cursor: pointer;
    border: 1px solid transparent;
    transition: background 0.15s ease, border-color 0.15s ease;
  }
  .onb-btn.primary {
    background: var(--color-accent);
    color: #fff;
    border-color: var(--color-accent);
  }
  .onb-btn.primary:hover { filter: brightness(1.05); }
  .onb-btn.secondary {
    background: transparent;
    color: var(--color-text);
    border-color: var(--color-border);
  }
  .onb-btn.secondary:hover { background: var(--color-accent-light); }
</style>
