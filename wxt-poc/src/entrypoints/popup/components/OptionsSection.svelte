<script lang="ts">
  import { createEventDispatcher } from 'svelte';

  export type OptionFormat = 'json' | 'csv' | 'html';

  export let format: OptionFormat = 'json';
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

<section class="card" aria-labelledby="options-section">
  <div class="section-head">
    <h2 class="section-title" id="options-section">Options</h2>
    <p class="section-sub">Choose export format and message details.</p>
  </div>
  <div class="field">
    <label for="format">Export format</label>
    <select
      id="format"
      value={format}
      on:change={(e) => dispatch('formatChange', (e.currentTarget as HTMLSelectElement).value as OptionFormat)}
    >
      <option value="json">JSON</option>
      <option value="csv">CSV</option>
      <option value="html">HTML</option>
    </select>
  </div>

  <div class="toggle-list" role="group" aria-label="Include options">
    <label class="toggle">
      <span class="toggle-label">
        <span class="toggle-icon">↩</span>
        <span>Include threaded replies</span>
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
        <span>Include reactions</span>
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
        <span>Include system messages</span>
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
        <span>Embed avatars (base64)</span>
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
