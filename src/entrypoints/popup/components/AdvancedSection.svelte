<script lang="ts">
  import { createEventDispatcher } from 'svelte';
  import { t } from '../../../i18n/i18n';

  type Theme = 'light' | 'dark';
  type LanguageOption = { value: string; label: string };

  export let open = false;
  export let showHud = true;
  export let lang = 'en';
  export let theme: Theme = 'light';
  export let languages: LanguageOption[] = [];

  const dispatch = createEventDispatcher<{
    toggleOpen: boolean;
    showHudChange: boolean;
    themeChange: Theme;
    langChange: string;
  }>();

  const handleThemeToggle = (evt: Event) => {
    const checked = (evt.currentTarget as HTMLInputElement)?.checked;
    dispatch('themeChange', checked ? 'dark' : 'light');
  };

  const cycleLang = (dir: 1 | -1) => {
    if (!languages.length) return;
    const current = languages.findIndex((l) => l.value === lang);
    const nextIndex = (current + dir + languages.length) % languages.length;
    dispatch('langChange', languages[nextIndex].value);
  };

  let currentLangLabel = '';
  $: currentLangLabel = languages.find((l) => l.value === lang)?.label || (lang || 'en').toUpperCase();
</script>

<section class="card" aria-labelledby="advanced-section" data-lang={lang}>
  <button
    type="button"
    class="card-toggle"
    id="advancedToggle"
    aria-expanded={open}
    aria-controls="advancedBody"
    on:click={() => dispatch('toggleOpen', !open)}
  >
    <div class="section-head">
      <h2 class="section-title" id="advanced-section">{t('advanced.title', {}, lang)}</h2>
      <p class="section-sub">{t('advanced.subtitle', {}, lang)}</p>
    </div>
    <span class="card-toggle-icon" aria-hidden="true">▼</span>
  </button>
  {#if open}
    <div class="card-body" id="advancedBody">
      <div class="pref-row">
        <div class="pref-meta">
          <span class="pref-title">Theme</span>
        </div>
        <label class="theme-toggle inline" for="themeToggle">
          <span class="icon" aria-hidden="true">☀</span>
          <input
            id="themeToggle"
            type="checkbox"
            aria-label="Toggle dark mode"
            checked={theme === 'dark'}
            on:change={handleThemeToggle}
          />
          <span class="icon" aria-hidden="true">🌙</span>
        </label>
      </div>

      <div class="pref-row">
        <div class="pref-meta">
          <span class="pref-title">{t('lang.label', {}, lang)}</span>
          <span class="pref-sub">{t('lang.subtitle', {}, lang)}</span>
        </div>
        <div class="lang-cycle">
          <button type="button" class="lang-arrow" aria-label="Previous language" on:click={() => cycleLang(-1)}>↑</button>
          <span class="lang-name">{currentLangLabel}</span>
          <button type="button" class="lang-arrow" aria-label="Next language" on:click={() => cycleLang(1)}>↓</button>
        </div>
      </div>

      <label class="toggle">
        <span class="toggle-label">
          <span class="toggle-icon">👁</span>
          <span>{t('advanced.hud', {}, lang)}</span>
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
