<script lang="ts">
  import { createEventDispatcher } from 'svelte';

  export let theme: Theme = 'light';

  const dispatch = createEventDispatcher<{ toggleTheme: Theme }>();

  const handleToggle = (evt: Event) => {
    const checked = (evt.currentTarget as HTMLInputElement)?.checked;
    dispatch('toggleTheme', checked ? 'dark' : 'light');
  };
</script>

<header class="header">
  <div class="header-top">
    <h1>Teams Chat Exporter</h1>
    <label class="theme-toggle" for="themeToggle">
      <span class="icon" aria-hidden="true">☀</span>
      <input
        id="themeToggle"
        type="checkbox"
        aria-label="Toggle dark mode"
        checked={theme === 'dark'}
        on:change={handleToggle}
      />
      <span class="icon" aria-hidden="true">🌙</span>
    </label>
  </div>
  <p>Export the active Teams chat with your preferred format and filters.</p>
</header>
<script lang="ts" context="module">
  export type Theme = 'light' | 'dark';
</script>
