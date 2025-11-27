<script lang="ts" context="module">
  export type OptionFormat = 'json' | 'csv' | 'html' | 'txt';
</script>

<script lang="ts">
  import { createEventDispatcher } from 'svelte';
  import { FileJson2, FileSpreadsheet, FileCode, FileText, FileType } from 'lucide-svelte';
  import { t } from '../../../i18n/i18n';

  export let format: OptionFormat = 'html';
  export let lang = 'en';

  const dispatch = createEventDispatcher<{
    formatChange: OptionFormat;
  }>();

  const formats = [
    { id: 'json' as const, icon: FileJson2 },
    { id: 'csv' as const, icon: FileSpreadsheet },
    { id: 'html' as const, icon: FileCode },
    { id: 'txt' as const, icon: FileText }
  ];

  function handleChange(newFormat: OptionFormat) {
    dispatch('formatChange', newFormat);
  }
</script>

<section class="format-section" data-lang={lang}>
  <div class="card">
    <div class="card-header">
      <div class="card-icon">
        <FileType size={16} />
      </div>
      <h2 class="card-title">{t('options.format', {}, lang)}</h2>
    </div>
    <div class="format-grid">
      {#each formats as fmt}
        {@const Icon = fmt.icon}
        <label
          class="format-radio"
          class:active={format === fmt.id}
        >
          <input
            type="radio"
            name="format"
            value={fmt.id}
            checked={format === fmt.id}
            on:change={() => handleChange(fmt.id)}
          />
          <div class="format-icon">
            <Icon size={22} />
          </div>
          <span class="format-label">{t(`format.${fmt.id}`, {}, lang)}</span>
        </label>
      {/each}
    </div>
  </div>
</section>
