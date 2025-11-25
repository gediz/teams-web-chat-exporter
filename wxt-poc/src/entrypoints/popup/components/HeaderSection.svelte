<script lang="ts">
  import { createEventDispatcher } from 'svelte';
  import { t } from '../../../i18n/i18n';

  export let theme: Theme = 'light';
  export let lang = 'en';

  const dispatch = createEventDispatcher<{ toggleTheme: Theme }>();

  const handleToggle = (evt: Event) => {
    const checked = (evt.currentTarget as HTMLInputElement)?.checked;
    dispatch('toggleTheme', checked ? 'dark' : 'light');
  };
</script>

<header class="header" data-lang={lang}>
  <div class="header-top">
    <h1>{t('title.app') || 'Teams Chat Exporter'}</h1>
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
  <p>{t('title.tagline') || ''}</p>
</header>
<script lang="ts" context="module">
  export type Theme = 'light' | 'dark';
</script>
