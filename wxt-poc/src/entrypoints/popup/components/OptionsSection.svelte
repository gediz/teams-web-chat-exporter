<script lang="ts" context="module">
  export type OptionFormat = 'json' | 'csv' | 'html' | 'txt';
</script>

<script lang="ts">
  import { createEventDispatcher } from 'svelte';
  import { t } from '../../../i18n/i18n';

  export let format: OptionFormat = 'json';
  export let lang = 'en';
  export let includeReplies = true;
  export let includeReactions = true;
  export let includeSystem = false;
  export let embedAvatars = false;

  const dispatch = createEventDispatcher<{
    formatChange: OptionFormat;
    includeRepliesChange: boolean;
    includeReactionsChange: boolean;
    includeSystemChange: boolean;
    embedAvatarsChange: boolean;
  }>();
</script>

<section class="card" aria-labelledby="options-section" data-lang={lang}>
  <div class="section-head">
    <h2 class="section-title" id="options-section">{t('options.title', {}, lang)}</h2>
    <p class="section-sub">{t('options.subtitle', {}, lang)}</p>
  </div>
  <div class="field">
    <label for="format">{t('options.format', {}, lang)}</label>
    <select
      id="format"
      value={format}
      on:change={(e) => dispatch('formatChange', (e.currentTarget as HTMLSelectElement).value as OptionFormat)}
    >
      <option value="json">{t('format.json', {}, lang)}</option>
      <option value="csv">{t('format.csv', {}, lang)}</option>
      <option value="html">{t('format.html', {}, lang)}</option>
      <option value="txt">{t('format.txt', {}, lang)}</option>
    </select>
  </div>

  <div class="toggle-list" role="group" aria-label="Include options">
    <label class="toggle">
      <span class="toggle-label">
        <span class="toggle-icon">↩</span>
        <span>{t('options.replies', {}, lang)}</span>
      </span>
      <input
        id="includeReplies"
        type="checkbox"
        checked={includeReplies}
        on:change={(e) => dispatch('includeRepliesChange', (e.currentTarget as HTMLInputElement).checked)}
      />
    </label>
    <label class="toggle">
      <span class="toggle-label">
        <span class="toggle-icon">😊</span>
        <span>{t('options.reactions', {}, lang)}</span>
      </span>
      <input
        id="includeReactions"
        type="checkbox"
        checked={includeReactions}
        on:change={(e) => dispatch('includeReactionsChange', (e.currentTarget as HTMLInputElement).checked)}
      />
    </label>
    <label class="toggle">
      <span class="toggle-label">
        <span class="toggle-icon">⚙</span>
        <span>{t('options.system', {}, lang)}</span>
      </span>
      <input
        id="includeSystem"
        type="checkbox"
        checked={includeSystem}
        on:change={(e) => dispatch('includeSystemChange', (e.currentTarget as HTMLInputElement).checked)}
      />
    </label>
    <label class="toggle">
      <span class="toggle-label">
        <span class="toggle-icon">👤</span>
        <span>{t('options.avatars', {}, lang)}</span>
      </span>
      <input
        id="embedAvatars"
        type="checkbox"
        checked={embedAvatars}
        on:change={(e) => dispatch('embedAvatarsChange', (e.currentTarget as HTMLInputElement).checked)}
      />
    </label>
  </div>
</section>
