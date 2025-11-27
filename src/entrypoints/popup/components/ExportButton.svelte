<script lang="ts">
  import { createEventDispatcher } from 'svelte';
  import { Download } from 'lucide-svelte';
  import { t } from '../../../i18n/i18n';

  export let disabled = false;
  export let busy = false;
  export let busyLabel = '';
  export let summary = '';
  export let lang = 'en';

  const dispatch = createEventDispatcher<{
    run: void;
  }>();

  function handleClick() {
    if (!disabled && !busy) {
      dispatch('run');
    }
  }

  $: label = busy ? busyLabel : t('actions.export', {}, lang);
</script>

<button
  class="export-primary"
  disabled={disabled || busy}
  on:click={handleClick}
>
  <div class="export-btn-content">
    <div class="export-icon">
      <Download size={22} />
    </div>
    <div class="export-text">
      <span class="export-title">{label}</span>
      {#if summary}
        <span class="export-summary">{summary}</span>
      {/if}
    </div>
  </div>
</button>
