<script lang="ts">
  import { createEventDispatcher } from 'svelte';

  export let busy = false;
  export let busyLabel = '';
  export let defaultRunLabel = '';
  export let statusText = '';

  const dispatch = createEventDispatcher<{ run: void }>();
</script>

<button id="run" class:busy={busy} on:click={() => dispatch('run')} disabled={busy}>
  <span class="spinner" aria-hidden="true"></span>
  <span class="button-icon" aria-hidden="true">⤓</span>
  <span class="label">{busy ? busyLabel : defaultRunLabel}</span>
</button>

<div id="status" aria-live="polite">{statusText}</div>

<p class="footer-note">
  Keep this popup open to monitor progress, or close it and watch the badge.
  <br />
  Need help? Reopen the popup to reconnect to an active export.
</p>
