<script lang="ts">
  import { createEventDispatcher } from 'svelte';

  export let open = false;
  export let showHud = true;

  const dispatch = createEventDispatcher<{
    toggleOpen: boolean;
    showHudChange: boolean;
  }>();
</script>

<section class="card" aria-labelledby="advanced-section">
  <button
    type="button"
    class="card-toggle"
    id="advancedToggle"
    aria-expanded={open}
    aria-controls="advancedBody"
    on:click={() => dispatch('toggleOpen', !open)}
  >
    <div class="section-head">
      <h2 class="section-title" id="advanced-section">Advanced</h2>
      <p class="section-sub">Developer / diagnostic features.</p>
    </div>
    <span class="card-toggle-icon" aria-hidden="true">▼</span>
  </button>
  {#if open}
    <div class="card-body" id="advancedBody">
      <label class="toggle">
        <span class="toggle-label">
          <span class="toggle-icon">👁</span>
          <span>Show in-page progress HUD</span>
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
