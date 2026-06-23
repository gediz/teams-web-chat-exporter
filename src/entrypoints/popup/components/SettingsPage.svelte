<script lang="ts">
  import { createEventDispatcher } from 'svelte';
  import { ArrowLeft, Sun, Moon, Globe, FolderOpen, CircleUserRound, Printer, Info, ExternalLink, Bug, Star, GraduationCap, Image, Stethoscope } from 'lucide-svelte';
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
  export let imageFetchFallback = false;
  export let fullResImages = false;

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
    imageFetchFallbackChange: boolean;
    fullResImagesChange: boolean;
    replayTour: void;
    openDiagnostics: void;
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

  function onFullResImagesChange(e: Event) {
    dispatch('fullResImagesChange', (e.target as HTMLInputElement).checked);
  }

  function onPdfIncludeAvatarsChange(e: Event) {
    dispatch('pdfIncludeAvatarsChange', (e.target as HTMLInputElement).checked);
  }

  // Permissions API selector. Firefox MV2 ships a callback-based
  // `chrome.permissions.*` polyfill that returns void from
  // request()/remove(); only `browser.permissions.*` is promise-based.
  // Chrome MV3 makes `chrome.permissions.*` itself promise-based and
  // doesn't expose `browser` (unless polyfilled). Picking
  // `browser ?? chrome` selects the promise-returning surface on
  // both. Same pattern App.svelte uses for runtime.
  // @ts-ignore - `browser` is the WebExtension global on Firefox
  const permsApi: typeof browser.permissions =
    // @ts-ignore
    typeof browser !== 'undefined' ? browser.permissions : chrome.permissions;

  // Image-fetch-fallback toggle. Two-phase because flipping ON has to
  // succeed at requesting <all_urls> before we can persist the option.
  // The browser's permission prompt requires a live user gesture (this
  // click handler runs inside one).
  //
  // We intercept `click` (not `change`) and preventDefault to stop the
  // browser's default toggle. That keeps the checkbox visually pinned
  // to whatever the prop says until *we* dispatch a state change. The
  // checkbox only flips visually after the parent's option update
  // flows back through the prop — i.e., after permission was actually
  // granted. While the prompt is visible the checkbox stays unchecked,
  // matching the user's mental model ("not enabled yet — pending").
  // On disable we also remove the host permission as cleanup; users
  // who re-enable will see the prompt again, which feels more honest
  // than silently re-arming a previously-granted permission.
  let imageFetchFallbackBusy = false;
  async function onImageFetchFallbackClick(e: MouseEvent) {
    e.preventDefault();
    if (imageFetchFallbackBusy) return;
    const desired = !imageFetchFallback;
    imageFetchFallbackBusy = true;
    try {
      if (desired) {
        let granted = false;
        try {
          granted = await permsApi.request({ origins: ['<all_urls>'] });
        } catch { /* user gesture lost / API unavailable — treat as denied */ }
        if (!granted) return;
        dispatch('imageFetchFallbackChange', true);
      } else {
        try {
          await permsApi.remove({ origins: ['<all_urls>'] });
        } catch { /* removal isn't critical — user can revoke from browser settings */ }
        dispatch('imageFetchFallbackChange', false);
      }
    } finally {
      imageFetchFallbackBusy = false;
    }
  }
</script>

<div class="settings-page">
  <div class="settings-header">
    <button class="icon-btn" title={t('common.back', {}, lang)} on:click={() => dispatch('back')}>
      <ArrowLeft size={18} />
    </button>
    <h1>{t('settings.title', {}, lang)}</h1>
    <!-- Stethoscope at the right edge mirrors the main popup's
         right-side icon idiom (history + gear). Diagnostics is a
         tool, not a setting, so it sits in the chrome of the page
         rather than as a card or a link in the body. -->
    <button
      class="icon-btn settings-header-right"
      title={t('settings.diagnostics.link', {}, lang)}
      aria-label={t('settings.diagnostics.link', {}, lang)}
      on:click={() => dispatch('openDiagnostics')}
    >
      <Stethoscope size={18} />
    </button>
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

  <!-- Image fetch fallback card. Off by default. Flipping ON triggers
       a runtime permission prompt for <all_urls>; copy is intentionally
       restrained about the broad permission so users opting in are
       making an informed choice. Tooltip on the ⓘ explains the
       use case + the permission cost. -->
  <div class="card settings-card">
    <div class="card-header">
      <span class="card-icon"><Image size={16} /></span>
      <span class="card-title">{t('settings.imageFallback', {}, lang) || 'Image fetch fallback'}</span>
      <span
        class="card-info"
        title={t('settings.imageFallback.tooltip', {}, lang) || 'Sometimes Teams\' image proxy fails to load external images (link previews, repo cards, article thumbnails). When enabled, the extension falls back to fetching them directly from the source. Requires permission to access all websites.'}
        aria-label={t('settings.imageFallback.tooltip', {}, lang) || 'Tooltip'}
        role="note"
      >ⓘ</span>
    </div>
    <div class="settings-subtitle">{t('settings.imageFallback.hint', {}, lang) || 'Try to get images directly when Teams\' proxy fails.'}</div>
    <label class="pdf-toggle">
      <input type="checkbox" checked={imageFetchFallback} on:click={onImageFetchFallbackClick} />
      <span>{t('settings.imageFallback.enable', {}, lang) || 'Enable'}</span>
    </label>
  </div>

  <div class="card settings-card">
    <div class="card-header">
      <span class="card-icon"><Image size={16} /></span>
      <span class="card-title">{t('settings.fullRes', {}, lang) || 'Full-resolution images'}</span>
      <span
        class="card-info"
        title={t('settings.fullRes.tooltip', {}, lang) || 'Save the original image instead of Teams\' downscaled view. Files get much larger, especially PDF, for the same on-screen size. Images that are too large or unavailable fall back to the downscaled view, so none are dropped.'}
        aria-label={t('settings.fullRes.tooltip', {}, lang) || 'Tooltip'}
        role="note"
      >ⓘ</span>
    </div>
    <div class="settings-subtitle">{t('settings.fullRes.hint', {}, lang) || 'Save original-quality images instead of the downscaled view. Much larger files.'}</div>
    <label class="pdf-toggle">
      <input type="checkbox" checked={fullResImages} on:change={onFullResImagesChange} />
      <span>{t('settings.fullRes.enable', {}, lang) || 'Enable'}</span>
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

  <!-- Replay tour card. Lives just before About so it's near the
       version / help affordances; users tend to look for "show help
       again" near "About". The button just dispatches; App.svelte
       closes the settings page and shows the overlay (it deliberately
       leaves onboardingDismissed=true since the user already saw the
       tour once — replay is a manual reopen, not a re-onboarding). -->
  <div class="card settings-card">
    <div class="card-header">
      <span class="card-icon"><GraduationCap size={16} /></span>
      <span class="card-title">{t('settings.replayTour', {}, lang) || 'Replay tour'}</span>
    </div>
    <div class="settings-subtitle">{t('settings.replayTour.hint', {}, lang) || 'Walk through the introduction tour again.'}</div>
    <button
      type="button"
      class="replay-tour-btn"
      on:click={() => dispatch('replayTour')}
    >
      <GraduationCap size={14} />
      {t('settings.replayTour.btn', {}, lang) || 'Replay tour'}
    </button>
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

<style>
  .replay-tour-btn {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    padding: 6px 12px;
    margin-top: 6px;
    border: 1px solid var(--color-border);
    background: var(--color-surface);
    color: var(--color-text);
    font: inherit;
    font-size: 12px;
    font-weight: 500;
    border-radius: 6px;
    cursor: pointer;
    transition: background 0.15s ease, border-color 0.15s ease, color 0.15s ease;
  }
  .replay-tour-btn:hover {
    background: var(--color-accent-light);
    border-color: var(--color-accent);
    color: var(--color-accent);
  }
  .replay-tour-btn :global(svg) {
    stroke: currentColor;
    fill: none;
    stroke-width: 2;
  }

  /* Inline ⓘ next to a card title — same affordance as the
     bundle.zip pill in FormatSection. Native title= tooltip on
     hover; no custom popover. */
  .card-info {
    margin-left: 4px;
    font-size: 12px;
    color: var(--color-muted);
    cursor: help;
    user-select: none;
  }
</style>
