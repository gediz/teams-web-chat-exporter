<script lang="ts">
  import { createEventDispatcher } from 'svelte';
  import { ArrowLeft, Sun, Moon, Globe } from 'lucide-svelte';
  import { t } from '../../../i18n/i18n';

  type Theme = 'light' | 'dark';
  type LanguageOption = { value: string; label: string; native: string; code: string };

  export let theme: Theme = 'light';
  export let lang = 'en';
  export let languages: LanguageOption[] = [];

  const dispatch = createEventDispatcher<{
    back: void;
    themeChange: Theme;
    langChange: string;
  }>();
</script>

<div class="settings-page">
  <div class="settings-header">
    <button class="icon-btn" title="Back" on:click={() => dispatch('back')}>
      <ArrowLeft size={18} />
    </button>
    <h1>{t('settings.title', {}, lang)}</h1>
  </div>

  <!-- Theme Card -->
  <div class="card settings-card">
    <div class="card-header">
      <span class="card-icon">
        {#if theme === 'dark'}<Moon size={16} />{:else}<Sun size={16} />{/if}
      </span>
      <span class="card-title">{t('settings.theme', {}, lang)}</span>
    </div>
    <div class="theme-toggle">
      <button
        class="theme-option"
        class:active={theme === 'light'}
        on:click={() => dispatch('themeChange', 'light')}
      >
        <Sun size={16} /> {t('settings.theme.light', {}, lang)}
      </button>
      <button
        class="theme-option"
        class:active={theme === 'dark'}
        on:click={() => dispatch('themeChange', 'dark')}
      >
        <Moon size={16} /> {t('settings.theme.dark', {}, lang)}
      </button>
    </div>
  </div>

  <!-- Language Card -->
  <div class="card settings-card">
    <div class="card-header">
      <span class="card-icon"><Globe size={16} /></span>
      <span class="card-title">{t('lang.label', {}, lang)}</span>
    </div>
    <div class="settings-subtitle">{t('lang.subtitle', {}, lang)}</div>
    <div class="lang-grid">
      {#each languages as language}
        <button
          class="lang-pill"
          class:active={lang === language.value}
          on:click={() => dispatch('langChange', language.value)}
          title={language.label}
        >
          <span class="lang-code">{language.code}</span>
          <span class="lang-native">{language.native}</span>
        </button>
      {/each}
    </div>
  </div>
</div>
