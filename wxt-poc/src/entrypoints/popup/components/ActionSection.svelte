<script lang="ts">
  import { createEventDispatcher } from 'svelte';
  import { t } from '../../../i18n/i18n';

  export let busy = false;
  export let busyLabel = '';
  export let defaultRunLabel = '';
  export let statusText = '';
  export let lang = 'en';

  const dispatch = createEventDispatcher<{ run: void }>();
</script>

<button id="run" class:busy={busy} on:click={() => dispatch('run')} disabled={busy} data-lang={lang}>
  <span class="spinner" aria-hidden="true"></span>
  <span class="button-icon" aria-hidden="true">⤓</span>
  <span class="label">{busy ? busyLabel : defaultRunLabel}</span>
</button>

<div id="status" aria-live="polite">{statusText}</div>

<p class="footer-note">
  {t('footer.tip1')}
  <br />
  {t('footer.tip2')}
</p>
