<script lang="ts">
  import { createEventDispatcher } from 'svelte';

  export type QuickRange = { key: string; label: string; icon: string };

  export let startAt = '';
  export let endAt = '';
  export let activeRange = 'none';
  export let ranges: QuickRange[] = [];

  const dispatch = createEventDispatcher<{
    changeStart: string;
    changeEnd: string;
    quickSelect: string;
  }>();
</script>

<section class="card" aria-labelledby="range-section">
  <div class="section-head">
    <h2 class="section-title" id="range-section">Date range</h2>
    <p class="section-sub">Leave blank to include everything.</p>
  </div>
  <div class="grid-two">
    <div class="field">
      <label for="startAt">From (inclusive)</label>
      <input
        id="startAt"
        class="input-text"
        type="text"
        placeholder="YYYY-MM-DD HH:MM"
        autocomplete="off"
        value={startAt}
        on:input={(e) => dispatch('changeStart', (e.currentTarget as HTMLInputElement).value)}
      />
    </div>
    <div class="field">
      <label for="endAt">To (exclusive)</label>
      <input
        id="endAt"
        class="input-text"
        type="text"
        placeholder="YYYY-MM-DD HH:MM"
        autocomplete="off"
        value={endAt}
        on:input={(e) => dispatch('changeEnd', (e.currentTarget as HTMLInputElement).value)}
      />
    </div>
  </div>
  <div class="field">
    <label class="section-sub" for="quickRanges">Quick ranges</label>
    <div id="quickRanges" aria-label="Quick ranges">
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
