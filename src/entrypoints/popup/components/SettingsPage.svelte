<script lang="ts">
  import { createEventDispatcher } from 'svelte';
  import { ArrowLeft, Sun, Moon, Globe, FolderOpen, CircleUserRound, Printer, Info, ExternalLink, Bug, Star } from 'lucide-svelte';
  import { t } from '../../../i18n/i18n';
  import { getReviewStoreUrl } from '../../../utils/store-urls';

  // Repo + issues + author URLs used by the About card. Hard-coded
  // strings (vs. reading from the manifest) because they never change
  // and a runtime lookup for a static string is needless overhead.
  const REPO_URL = 'https://github.com/gediz/teams-web-chat-exporter';
  const ISSUES_URL = `${REPO_URL}/issues`;
  const AUTHOR_NAME = 'Gediz';
  const AUTHOR_LINKEDIN_URL = 'https://www.linkedin.com/in/nazimgedizaydindogmus/';
  // Store review URL picked from UA. Shown alongside Source / Report
  // so engaged users have a permanent entry point regardless of
  // whether the inline one-liner ever fires for them.
  const REVIEW_STORE_URL = getReviewStoreUrl();

  // Version shown in the About card. Falls back to an empty string if
  // chrome.runtime isn't available (e.g. unit-test harness), so the
  // header still renders cleanly.
  const extensionVersion = (() => {
    try { return chrome?.runtime?.getManifest?.()?.version || ''; }
    catch { return ''; }
  })();

  type Theme = 'light' | 'dark';
  type AfterExport = 'manual' | 'show';
  type AvatarMode = 'inline' | 'files';
  type PdfPageSize = 'a4' | 'letter';
  type LanguageOption = { value: string; label: string; native: string; code: string };

  // Keep in sync with PDF_FONT_MIN/MAX in utils/options.ts. Duplicated
  // here as plain literals so this component stays a pure view.
  const PDF_FONT_MIN = 8;
  const PDF_FONT_MAX = 16;

  export let theme: Theme = 'light';
  export let lang = 'en';
  export let languages: LanguageOption[] = [];
  export let afterExport: AfterExport = 'manual';
  export let avatarMode: AvatarMode = 'inline';
  // embedAvatars is read-only here; the toggle lives in IncludeSection.
  // We surface it so the avatarMode select can be disabled when avatars
  // aren't being embedded — 'inline vs files' is meaningless otherwise,
  // and the disabled state makes the dependency obvious.
  export let embedAvatars = false;
  export let pdfPageSize: PdfPageSize = 'a4';
  export let pdfBodyFontSize = 10;
  export let pdfShowPageNumbers = true;
  export let pdfIncludeAvatars = true;

  const dispatch = createEventDispatcher<{
    back: void;
    themeChange: Theme;
    langChange: string;
    afterExportChange: AfterExport;
    avatarModeChange: AvatarMode;
    pdfPageSizeChange: PdfPageSize;
    pdfBodyFontSizeChange: number;
    pdfShowPageNumbersChange: boolean;
    pdfIncludeAvatarsChange: boolean;
  }>();

  function onAfterExportChange(e: Event) {
    const value = (e.target as HTMLSelectElement).value as AfterExport;
    dispatch('afterExportChange', value);
  }

  function onAvatarModeChange(e: Event) {
    const value = (e.target as HTMLSelectElement).value as AvatarMode;
    dispatch('avatarModeChange', value);
  }

  function onPdfPageSizeChange(e: Event) {
    const value = (e.target as HTMLSelectElement).value as PdfPageSize;
    dispatch('pdfPageSizeChange', value);
  }

  function onPdfBodyFontSizeChange(e: Event) {
    const raw = Number((e.target as HTMLInputElement).value);
    // Clamp on the way out so a user typing "999" doesn't propagate a
    // crazy value to the builder. The input's min/max attrs handle
    // most cases, but typing past them or pasting doesn't trigger them.
    if (!Number.isFinite(raw)) return;
    const clamped = Math.max(PDF_FONT_MIN, Math.min(PDF_FONT_MAX, Math.round(raw)));
    dispatch('pdfBodyFontSizeChange', clamped);
  }

  function onPdfShowPageNumbersChange(e: Event) {
    dispatch('pdfShowPageNumbersChange', (e.target as HTMLInputElement).checked);
  }

  function onPdfIncludeAvatarsChange(e: Event) {
    dispatch('pdfIncludeAvatarsChange', (e.target as HTMLInputElement).checked);
  }
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

  <!-- After export card -->
  <div class="card settings-card">
    <div class="card-header">
      <span class="card-icon"><FolderOpen size={16} /></span>
      <span class="card-title">{t('settings.afterExport', {}, lang) || 'After export'}</span>
    </div>
    <div class="settings-subtitle">{t('settings.afterExport.hint', {}, lang) || 'What happens once the file is saved.'}</div>
    <select
      class="after-export-select"
      value={afterExport}
      on:change={onAfterExportChange}
    >
      <option value="manual">{t('settings.afterExport.manual', {}, lang) || 'Let me decide'}</option>
      <option value="show">{t('settings.afterExport.show', {}, lang) || 'Show in folder automatically'}</option>
    </select>
  </div>

  <!-- Avatar mode card. The select is disabled when embedAvatars is
       off — pick 'inline' or 'files' only means something when avatars
       are actually being saved. The subtitle also swaps to a hint
       pointing at the Include section so the user knows how to enable it. -->
  <div class="card settings-card" class:disabled-card={!embedAvatars}>
    <div class="card-header">
      <span class="card-icon"><CircleUserRound size={16} /></span>
      <span class="card-title">{t('settings.avatarMode', {}, lang) || 'Avatars in HTML'}</span>
    </div>
    <div class="settings-subtitle">
      {#if embedAvatars}
        {t('settings.avatarMode.hint', {}, lang) || 'How avatar images are packaged in HTML exports.'}
      {:else}
        {t('settings.avatarMode.disabledHint', {}, lang) || 'Enable "Embed avatars" in the main page to choose how they are packaged.'}
      {/if}
    </div>
    <select
      class="after-export-select"
      value={avatarMode}
      disabled={!embedAvatars}
      on:change={onAvatarModeChange}
    >
      <option value="inline">{t('settings.avatarMode.inline', {}, lang) || 'Embed in HTML file (larger single file)'}</option>
      <option value="files">{t('settings.avatarMode.files', {}, lang) || 'Save as separate files (requires .zip)'}</option>
    </select>
  </div>

  <!-- PDF export card -->
  <div class="card settings-card">
    <div class="card-header">
      <span class="card-icon"><Printer size={16} /></span>
      <span class="card-title">{t('settings.pdf', {}, lang) || 'PDF export'}</span>
    </div>
    <div class="settings-subtitle">{t('settings.pdf.hint', {}, lang) || 'Layout preferences applied when exporting as PDF.'}</div>

    <!-- Page size -->
    <label class="pdf-row">
      <span class="pdf-row-label">{t('settings.pdf.pageSize', {}, lang) || 'Page size'}</span>
      <select
        class="after-export-select"
        value={pdfPageSize}
        on:change={onPdfPageSizeChange}
      >
        <option value="a4">{t('settings.pdf.pageSize.a4', {}, lang) || 'A4'}</option>
        <option value="letter">{t('settings.pdf.pageSize.letter', {}, lang) || 'US Letter'}</option>
      </select>
    </label>

    <!-- Body font size — numeric input 8–16 pt -->
    <label class="pdf-row">
      <span class="pdf-row-label">{t('settings.pdf.fontSize', {}, lang) || 'Body font size (pt)'}</span>
      <input
        class="pdf-num"
        type="number"
        min={PDF_FONT_MIN}
        max={PDF_FONT_MAX}
        step="1"
        value={pdfBodyFontSize}
        on:change={onPdfBodyFontSizeChange}
      />
    </label>
    <div class="pdf-hint">{t('settings.pdf.fontSize.hint', {}, lang) || 'Typical 10. Compact 9. Large 12.'}</div>

    <!-- Page numbers -->
    <label class="pdf-toggle">
      <input type="checkbox" checked={pdfShowPageNumbers} on:change={onPdfShowPageNumbersChange} />
      <span>{t('settings.pdf.pageNumbers', {}, lang) || 'Show page numbers'}</span>
    </label>

    <!-- Avatars in PDF -->
    <label class="pdf-toggle">
      <input type="checkbox" checked={pdfIncludeAvatars} on:change={onPdfIncludeAvatarsChange} />
      <span>{t('settings.pdf.avatars', {}, lang) || 'Include avatars in PDF'}</span>
    </label>
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

  <!-- About Card. Links open in a new tab via target=_blank so the
       popup doesn't close under the user. rel=noopener keeps the new
       tab's window.opener null — standard hygiene for external links. -->
  <div class="card settings-card about-card">
    <div class="card-header">
      <span class="card-icon"><Info size={16} /></span>
      <span class="card-title">{t('settings.about', {}, lang) || 'About'}</span>
    </div>
    <div class="about-body">
      <div class="about-name">
        {t('appName', {}, lang) || 'Teams Chat Exporter'}
        {#if extensionVersion}<span class="about-version">v{extensionVersion}</span>{/if}
      </div>
      <div class="about-links">
        <a class="about-link" href={REPO_URL} target="_blank" rel="noopener">
          <ExternalLink size={12} />
          <span>{t('settings.about.source', {}, lang) || 'Source code'}</span>
        </a>
        <a class="about-link" href={ISSUES_URL} target="_blank" rel="noopener">
          <Bug size={12} />
          <span>{t('settings.about.feedback', {}, lang) || 'Report an issue'}</span>
        </a>
        <a class="about-link" href={REVIEW_STORE_URL} target="_blank" rel="noopener">
          <Star size={12} />
          <span>{t('review.rate', {}, lang) || 'Rate on store'}</span>
        </a>
      </div>
      <!-- Author line: short name + external link to LinkedIn profile.
           No icon — the attribution sits below the two action links
           and reads cleanest as plain "Author: Gediz →". -->
      <div class="about-author">
        <span class="about-author-label">{t('settings.about.author', {}, lang) || 'Author'}:</span>
        <a class="about-author-link" href={AUTHOR_LINKEDIN_URL} target="_blank" rel="noopener">
          {AUTHOR_NAME}<ExternalLink size={10} />
        </a>
      </div>
    </div>
  </div>
</div>
