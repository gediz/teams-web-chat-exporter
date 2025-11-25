<script lang="ts" context="module">
  export type QuickRange = { key: string; label: string; icon: string };
</script>

<script lang="ts">
  import { createEventDispatcher } from 'svelte';
  import { t } from '../../../i18n/i18n';

  export let startAt = '';
  export let endAt = '';
  export let activeRange = 'none';
  export let ranges: QuickRange[] = [];
  export let lang = 'en';

  const dispatch = createEventDispatcher<{
    changeStart: string;
    changeEnd: string;
    quickSelect: string;
  }>();
</script>

<section class="card" aria-labelledby="range-section" data-lang={lang}>
  <div class="section-head">
    <h2 class="section-title" id="range-section">{t('range.title')}</h2>
    <p class="section-sub">{t('range.subtitle')}</p>
  </div>
  <div class="grid-two">
    <div class="field">
      <label for="startAt">{t('range.from')}</label>
      <input
        id="startAt"
        class="input-text"
        type="text"
        placeholder={t('range.placeholder')}
        autocomplete="off"
        value={startAt}
        on:input={(e) => dispatch('changeStart', (e.currentTarget as HTMLInputElement).value)}
      />
    </div>
    <div class="field">
      <label for="endAt">{t('range.to')}</label>
      <input
        id="endAt"
        class="input-text"
        type="text"
        placeholder={t('range.placeholder')}
        autocomplete="off"
        value={endAt}
        on:input={(e) => dispatch('changeEnd', (e.currentTarget as HTMLInputElement).value)}
      />
    </div>
  </div>
  <div class="field">
    <label class="section-sub" for="quickRanges">{t('range.quick')}</label>
    <div id="quickRanges" aria-label={t('range.quick')}>
      {#each ranges as qr}
        <button
          type="button"
          class={`chip ${activeRange === qr.key ? 'active' : ''}`}
          data-range={qr.key}
          data-icon={qr.icon}
          on:click={() => dispatch('quickSelect', qr.key)}
        >
          {qr.label}
        </button>
      {/each}
    </div>
  </div>
</section>
