<script lang="ts">
  import { createEventDispatcher } from 'svelte';
  import { t } from '../../../i18n/i18n';

  export let open = false;
  export let showHud = true;
  export let lang = 'en';

  const dispatch = createEventDispatcher<{
    toggleOpen: boolean;
    showHudChange: boolean;
  }>();
</script>

<section class="card" aria-labelledby="advanced-section" data-lang={lang}>
  <button
    type="button"
    class="card-toggle"
    id="advancedToggle"
    aria-expanded={open}
    aria-controls="advancedBody"
    on:click={() => dispatch('toggleOpen', !open)}
  >
    <div class="section-head">
      <h2 class="section-title" id="advanced-section">{t('advanced.title')}</h2>
      <p class="section-sub">{t('advanced.subtitle')}</p>
    </div>
    <span class="card-toggle-icon" aria-hidden="true">▼</span>
  </button>
  {#if open}
    <div class="card-body" id="advancedBody">
      <label class="toggle">
        <span class="toggle-label">
          <span class="toggle-icon">👁</span>
          <span>{t('advanced.hud')}</span>
        </span>
        <input
          id="showHud"
          type="checkbox"
          checked={showHud}
          on:change={(e) => dispatch('showHudChange', (e.currentTarget as HTMLInputElement).checked)}
        />
      </label>
    </div>
  {/if}
</section>
