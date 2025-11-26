<script lang="ts">
  import { createEventDispatcher } from 'svelte';
  import { Moon, Sun, Globe } from 'lucide-svelte';

  type Theme = 'light' | 'dark';
  type LanguageOption = { value: string; label: string };

  export let theme: Theme = 'light';
  export let lang = 'en';
  export let languages: LanguageOption[] = [];

  const dispatch = createEventDispatcher<{
    themeChange: Theme;
    langChange: string;
  }>();

  let showLangMenu = false;

  function toggleLangMenu() {
    showLangMenu = !showLangMenu;
  }

  function selectLanguage(code: string) {
    dispatch('langChange', code);
    showLangMenu = false;
  }

  function handleThemeToggle() {
    dispatch('themeChange', theme === 'dark' ? 'light' : 'dark');
  }

  let currentLangLabel = '';
  $: currentLangLabel = languages.find((l) => l.value === lang)?.label || (lang || 'en').toUpperCase();
</script>

<div class="header-actions">
  <button
    class="icon-btn"
    title="Toggle theme"
    on:click={handleThemeToggle}
  >
    {#if theme === 'dark'}
      <Sun size={18} />
    {:else}
      <Moon size={18} />
    {/if}
  </button>

  <div class="lang-dropdown">
    <button
      class="icon-btn"
      title="Select language"
      on:click={toggleLangMenu}
    >
      <Globe size={18} />
    </button>
    {#if showLangMenu}
      <div class="lang-menu show">
        {#each languages as language}
          <div
            class="lang-item"
            class:active={lang === language.value}
            on:click={() => selectLanguage(language.value)}
            role="button"
            tabindex="0"
            on:keydown={(e) => e.key === 'Enter' && selectLanguage(language.value)}
          >
            <span class="lang-name">{language.label}</span>
          </div>
        {/each}
      </div>
    {/if}
  </div>
</div>
