<script lang="ts">
  import { createEventDispatcher } from 'svelte';
  import {
    ArrowLeft, Palette, Download, Image, FileText,
    LifeBuoy, Info, ChevronRight, ExternalLink, Stethoscope, Search,
  } from 'lucide-svelte';
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

  // Version shown in the version footer. Falls back to an empty string if
  // chrome.runtime isn't available (e.g. unit-test harness), so the
  // footer still renders cleanly.
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
  export let imageFilenameDate = false;
  export let imageModifiedDate = false;
  export let attachmentSizeCapMb = 0;
  export let attachmentFilenameDate = false;
  export let attachmentSkipTypes = '';
  export let attachmentMaxConcurrent = 6;
  // Keep in sync with ATTACH_CONCURRENCY_MIN/MAX in src/utils/options.ts.
  const ATTACH_CONCURRENCY_MIN = 1;
  const ATTACH_CONCURRENCY_MAX = 8;

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
    imageFilenameDateChange: boolean;
    imageModifiedDateChange: boolean;
    attachmentSizeCapMbChange: number;
    attachmentFilenameDateChange: boolean;
    attachmentSkipTypesChange: string;
    attachmentMaxConcurrentChange: number;
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

  // Body font size is a −/+ stepper. Step by 1 pt and clamp to
  // [MIN, MAX]; the buttons also disable at the bounds so the value can
  // never leave range. Kept as a stepper (not a free-type field) so
  // there's no invalid intermediate state to sanitise.
  function stepFontSize(delta: number) {
    const next = Math.max(
      PDF_FONT_MIN,
      Math.min(PDF_FONT_MAX, Math.round(pdfBodyFontSize) + delta),
    );
    if (next !== pdfBodyFontSize) dispatch('pdfBodyFontSizeChange', next);
  }

  function onPdfShowPageNumbersChange(e: Event) {
    dispatch('pdfShowPageNumbersChange', (e.target as HTMLInputElement).checked);
  }

  function onFullResImagesChange(e: Event) {
    dispatch('fullResImagesChange', (e.target as HTMLInputElement).checked);
  }

  function onImageFilenameDateChange(e: Event) {
    dispatch('imageFilenameDateChange', (e.target as HTMLInputElement).checked);
  }

  function onImageModifiedDateChange(e: Event) {
    dispatch('imageModifiedDateChange', (e.target as HTMLInputElement).checked);
  }

  // Size cap in MB. Empty or non-positive means "no cap" (0). Clamp negatives
  // and junk to 0 so a bad value can never skip every attachment.
  function onAttachmentSizeCapChange(e: Event) {
    const raw = Number((e.target as HTMLInputElement).value);
    const mb = Number.isFinite(raw) && raw > 0 ? Math.floor(raw) : 0;
    dispatch('attachmentSizeCapMbChange', mb);
  }

  function onAttachmentFilenameDateChange(e: Event) {
    dispatch('attachmentFilenameDateChange', (e.target as HTMLInputElement).checked);
  }

  function onAttachmentSkipTypesChange(e: Event) {
    dispatch('attachmentSkipTypesChange', (e.target as HTMLInputElement).value.trim());
  }

  // Concurrency is a −/+ stepper (like body font size): step by 1 and clamp to
  // [MIN, MAX]; the buttons disable at the bounds so the value can't leave range.
  function stepConcurrency(delta: number) {
    const next = Math.max(
      ATTACH_CONCURRENCY_MIN,
      Math.min(ATTACH_CONCURRENCY_MAX, Math.round(attachmentMaxConcurrent) + delta),
    );
    if (next !== attachmentMaxConcurrent) dispatch('attachmentMaxConcurrentChange', next);
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

  // Language picker is a drill-in sub-page rather than an inline grid:
  // 24 languages as pills doubled the settings height and pushed every
  // other card down. The drill-in keeps the main list stable and gives
  // the language list room + a search box. 'view' swaps the whole page
  // body; the header morphs to a Language header with its own back arrow.
  let view: 'main' | 'lang' = 'main';
  let langQuery = '';

  // Current language's native name, shown as the value on the Language
  // row. Falls back to the label then the raw code so the row always
  // shows something even if `languages` is empty (unit-test harness).
  $: currentLangNative =
    languages.find((l) => l.value === lang)?.native ??
    languages.find((l) => l.value === lang)?.label ??
    lang;

  $: filteredLanguages = (() => {
    const q = langQuery.trim().toLowerCase();
    if (!q) return languages;
    return languages.filter(
      (l) =>
        l.native.toLowerCase().includes(q) ||
        l.label.toLowerCase().includes(q) ||
        l.code.toLowerCase().includes(q) ||
        l.value.toLowerCase().includes(q),
    );
  })();

  function openLang() {
    langQuery = '';
    view = 'lang';
  }
  function closeLang() {
    view = 'main';
    langQuery = '';
  }
  function pickLanguage(value: string) {
    dispatch('langChange', value);
    closeLang();
  }
</script>

<div class="settings-page">
  {#if view === 'lang'}
    <!-- Language drill-in sub-page -->
    <div class="settings-header">
      <button class="icon-btn" title={t('common.back', {}, lang)} on:click={closeLang}>
        <ArrowLeft size={18} />
      </button>
      <h1>{t('lang.label', {}, lang)}</h1>
    </div>

    <div class="lang-search">
      <Search size={15} />
      <!-- svelte-ignore a11y-autofocus -->
      <input
        type="text"
        bind:value={langQuery}
        placeholder={t('common.search', {}, lang)}
        aria-label={t('common.search', {}, lang)}
      />
    </div>

    <section class="ac">
      <div class="group">
        {#each filteredLanguages as language (language.value)}
          <button
            class="srow rowlink lang-item"
            class:cur={lang === language.value}
            aria-current={lang === language.value ? 'true' : undefined}
            on:click={() => pickLanguage(language.value)}
          >
            <span class="txt">
              <span class="label">{language.native}</span>
              <span class="sub">{language.label}</span>
            </span>
            {#if lang === language.value}<span class="chk" aria-hidden="true">✓</span>{/if}
          </button>
        {:else}
          <div class="lang-empty">—</div>
        {/each}
      </div>
    </section>
  {:else}
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

    <!-- Appearance: theme + language -->
    <section class="ac">
      <div class="ah"><Palette size={14} /> {t('settings.grp.appearance', {}, lang)}</div>
      <div class="group">
        <div class="srow">
          <span class="txt"><span class="label">{t('settings.theme', {}, lang)}</span></span>
          <span class="seg">
            <button
              class="seg-btn"
              class:sel={theme === 'light'}
              on:click={() => dispatch('themeChange', 'light')}
            >
              {t('settings.theme.light', {}, lang)}
            </button>
            <button
              class="seg-btn"
              class:sel={theme === 'dark'}
              on:click={() => dispatch('themeChange', 'dark')}
            >
              {t('settings.theme.dark', {}, lang)}
            </button>
          </span>
        </div>
        <button class="srow rowlink" on:click={openLang}>
          <span class="txt"><span class="label">{t('lang.label', {}, lang)}</span></span>
          <span class="val">{currentLangNative}</span>
          <span class="chev"><ChevronRight size={16} /></span>
        </button>
      </div>
    </section>

    <!-- Export: after-export behaviour + HTML avatar packaging.
         Both use long option labels, so they stack (label above a
         full-width select) rather than sitting flush-right. -->
    <section class="ac">
      <div class="ah"><Download size={14} /> {t('settings.grp.export', {}, lang)}</div>
      <div class="group">
        <div class="srow stacked">
          <span class="txt">
            <span class="label">{t('settings.afterExport', {}, lang)}</span>
            <span class="sub">{t('settings.afterExport.hint', {}, lang)}</span>
          </span>
          <select class="field" value={afterExport} on:change={onAfterExportChange}>
            <option value="manual">{t('settings.afterExport.manual', {}, lang)}</option>
            <option value="show">{t('settings.afterExport.show', {}, lang)}</option>
          </select>
        </div>

        <!-- Avatar mode dims + disables when embedAvatars is off — the
             inline/files choice is meaningless when no avatars are saved.
             The sub-line swaps to a hint pointing back to the main page. -->
        <div class="srow stacked" class:is-disabled={!embedAvatars}>
          <span class="txt">
            <span class="label">{t('settings.avatarMode', {}, lang)}</span>
            <span class="sub">
              {#if embedAvatars}
                {t('settings.avatarMode.hint', {}, lang)}
              {:else}
                {t('settings.avatarMode.disabledHint', {}, lang)}
              {/if}
            </span>
          </span>
          <select
            class="field"
            value={avatarMode}
            disabled={!embedAvatars}
            on:change={onAvatarModeChange}
          >
            <option value="inline">{t('settings.avatarMode.inline', {}, lang)}</option>
            <option value="files">{t('settings.avatarMode.files', {}, lang)}</option>
          </select>
        </div>
      </div>
    </section>

    <!-- Images: four toggles consolidated into one section. Each row shows
         a short hint as the sub-line and a visible ⓘ next to the label; the
         ⓘ carries the full detail via title (hover) and aria-label (focus /
         screen reader). The ⓘ is a <button> so tabbing reaches it and its
         accessible name is read.

         Each label carries an explicit for=/id= to its checkbox: <button>
         is itself a labelable element, so without for= the wrapping label
         would associate with the ⓘ (first labelable descendant) instead of
         the checkbox and the row would stop toggling. for= pins the label
         to the checkbox; clicking the ⓘ (interactive content) is still
         suppressed, and stopPropagation is belt-and-suspenders. -->
    <section class="ac">
      <div class="ah"><Image size={14} /> {t('settings.grp.images', {}, lang)}</div>
      <div class="group">
        <label class="srow toggle" for="opt-full-res">
          <span class="txt">
            <span class="label-row">
              <span class="label">{t('settings.fullRes', {}, lang)}</span>
              <button
                type="button"
                class="info-btn"
                title={t('settings.fullRes.tooltip', {}, lang)}
                aria-label={t('settings.fullRes.tooltip', {}, lang)}
                on:click|stopPropagation|preventDefault={() => {}}
              >ⓘ</button>
            </span>
            <span class="sub">{t('settings.fullRes.hint', {}, lang)}</span>
          </span>
          <span class="switch">
            <input id="opt-full-res" type="checkbox" checked={fullResImages} on:change={onFullResImagesChange} />
            <span class="track"></span>
          </span>
        </label>

        <!-- Image fetch fallback: on:click (not on:change) drives the
             permission-gated two-phase flow. Do not switch to on:change. -->
        <label class="srow toggle" for="opt-img-fallback">
          <span class="txt">
            <span class="label-row">
              <span class="label">{t('settings.imageFallback', {}, lang)}</span>
              <button
                type="button"
                class="info-btn"
                title={t('settings.imageFallback.tooltip', {}, lang)}
                aria-label={t('settings.imageFallback.tooltip', {}, lang)}
                on:click|stopPropagation|preventDefault={() => {}}
              >ⓘ</button>
            </span>
            <span class="sub">{t('settings.imageFallback.hint', {}, lang)}</span>
          </span>
          <span class="switch">
            <input id="opt-img-fallback" type="checkbox" checked={imageFetchFallback} on:click={onImageFetchFallbackClick} />
            <span class="track"></span>
          </span>
        </label>

        <label class="srow toggle" for="opt-img-date">
          <span class="txt">
            <span class="label-row">
              <span class="label">{t('settings.imageDate', {}, lang)}</span>
              <button
                type="button"
                class="info-btn"
                title={t('settings.imageDate.tooltip', {}, lang)}
                aria-label={t('settings.imageDate.tooltip', {}, lang)}
                on:click|stopPropagation|preventDefault={() => {}}
              >ⓘ</button>
            </span>
            <span class="sub">{t('settings.imageDate.hint', {}, lang)}</span>
          </span>
          <span class="switch">
            <input id="opt-img-date" type="checkbox" checked={imageFilenameDate} on:change={onImageFilenameDateChange} />
            <span class="track"></span>
          </span>
        </label>

        <label class="srow toggle" for="opt-img-mtime">
          <span class="txt">
            <span class="label-row">
              <span class="label">{t('settings.imageMtime', {}, lang)}</span>
              <button
                type="button"
                class="info-btn"
                title={t('settings.imageMtime.tooltip', {}, lang)}
                aria-label={t('settings.imageMtime.tooltip', {}, lang)}
                on:click|stopPropagation|preventDefault={() => {}}
              >ⓘ</button>
            </span>
            <span class="sub">{t('settings.imageMtime.hint', {}, lang)}</span>
          </span>
          <span class="switch">
            <input id="opt-img-mtime" type="checkbox" checked={imageModifiedDate} on:change={onImageModifiedDateChange} />
            <span class="track"></span>
          </span>
        </label>
      </div>
    </section>

    <!-- Attachments: only take effect when the "Files" toggle is on. -->
    <section class="ac">
      <div class="ah"><FileText size={14} /> {t('settings.grp.attachments', {}, lang)}</div>
      <div class="group">
        <div class="srow stacked">
          <span class="txt">
            <span class="label-row">
              <span class="label">{t('settings.attachSizeCap', {}, lang)}</span>
              <button
                type="button"
                class="info-btn"
                title={t('settings.attachSizeCap.tooltip', {}, lang)}
                aria-label={t('settings.attachSizeCap.tooltip', {}, lang)}
                on:click|stopPropagation|preventDefault={() => {}}
              >ⓘ</button>
            </span>
            <span class="sub">{t('settings.attachSizeCap.hint', {}, lang)}</span>
          </span>
          <input
            class="field num"
            type="number"
            min="0"
            step="1"
            inputmode="numeric"
            value={attachmentSizeCapMb || ''}
            placeholder="0"
            aria-label={t('settings.attachSizeCap', {}, lang)}
            on:change={onAttachmentSizeCapChange}
          />
        </div>

        <div class="srow stacked">
          <span class="txt">
            <span class="label-row">
              <span class="label">{t('settings.attachSkipTypes', {}, lang)}</span>
              <button
                type="button"
                class="info-btn"
                title={t('settings.attachSkipTypes.tooltip', {}, lang)}
                aria-label={t('settings.attachSkipTypes.tooltip', {}, lang)}
                on:click|stopPropagation|preventDefault={() => {}}
              >ⓘ</button>
            </span>
            <span class="sub">{t('settings.attachSkipTypes.hint', {}, lang)}</span>
          </span>
          <input
            class="field"
            type="text"
            value={attachmentSkipTypes}
            placeholder="exe, zip"
            aria-label={t('settings.attachSkipTypes', {}, lang)}
            on:change={onAttachmentSkipTypesChange}
          />
        </div>

        <label class="srow toggle" for="opt-attach-date">
          <span class="txt">
            <span class="label-row">
              <span class="label">{t('settings.attachDate', {}, lang)}</span>
              <button
                type="button"
                class="info-btn"
                title={t('settings.attachDate.tooltip', {}, lang)}
                aria-label={t('settings.attachDate.tooltip', {}, lang)}
                on:click|stopPropagation|preventDefault={() => {}}
              >ⓘ</button>
            </span>
            <span class="sub">{t('settings.attachDate.hint', {}, lang)}</span>
          </span>
          <span class="switch">
            <input id="opt-attach-date" type="checkbox" checked={attachmentFilenameDate} on:change={onAttachmentFilenameDateChange} />
            <span class="track"></span>
          </span>
        </label>

        <div class="srow">
          <span class="txt">
            <span class="label-row">
              <span class="label">{t('settings.attachConcurrency', {}, lang)}</span>
              <button
                type="button"
                class="info-btn"
                title={t('settings.attachConcurrency.tooltip', {}, lang)}
                aria-label={t('settings.attachConcurrency.tooltip', {}, lang)}
                on:click|stopPropagation|preventDefault={() => {}}
              >ⓘ</button>
            </span>
            <span class="sub">{t('settings.attachConcurrency.hint', {}, lang)}</span>
          </span>
          <span class="stepper" role="group" aria-label={t('settings.attachConcurrency', {}, lang)}>
            <button
              type="button"
              class="step-btn"
              on:click={() => stepConcurrency(-1)}
              disabled={attachmentMaxConcurrent <= ATTACH_CONCURRENCY_MIN}
            >−</button>
            <span class="step-val">{attachmentMaxConcurrent}</span>
            <button
              type="button"
              class="step-btn"
              on:click={() => stepConcurrency(1)}
              disabled={attachmentMaxConcurrent >= ATTACH_CONCURRENCY_MAX}
            >+</button>
          </span>
        </div>
      </div>
    </section>

    <!-- PDF: compact controls flush-right (short values). -->
    <section class="ac">
      <div class="ah"><FileText size={14} /> {t('settings.pdf', {}, lang)}</div>
      <div class="group">
        <div class="srow">
          <span class="txt"><span class="label">{t('settings.pdf.pageSize', {}, lang)}</span></span>
          <select class="field mini" value={pdfPageSize} on:change={onPdfPageSizeChange}>
            <option value="a4">{t('settings.pdf.pageSize.a4', {}, lang)}</option>
            <option value="letter">{t('settings.pdf.pageSize.letter', {}, lang)}</option>
          </select>
        </div>

        <div class="srow">
          <span class="txt">
            <span class="label">{t('settings.pdf.fontSize', {}, lang)}</span>
            <span class="sub">{t('settings.pdf.fontSize.hint', {}, lang)}</span>
          </span>
          <!-- role=group + the field name as the group's accessible label
               so assistive tech announces what the ± buttons control. Each
               button's name comes from its visible − / + glyph. -->
          <span class="stepper" role="group" aria-label={t('settings.pdf.fontSize', {}, lang)}>
            <button
              type="button"
              class="step-btn"
              on:click={() => stepFontSize(-1)}
              disabled={pdfBodyFontSize <= PDF_FONT_MIN}
            >−</button>
            <span class="step-val">{pdfBodyFontSize}</span>
            <button
              type="button"
              class="step-btn"
              on:click={() => stepFontSize(1)}
              disabled={pdfBodyFontSize >= PDF_FONT_MAX}
            >+</button>
          </span>
        </div>

        <label class="srow toggle">
          <span class="txt"><span class="label">{t('settings.pdf.pageNumbers', {}, lang)}</span></span>
          <span class="switch">
            <input type="checkbox" checked={pdfShowPageNumbers} on:change={onPdfShowPageNumbersChange} />
            <span class="track"></span>
          </span>
        </label>

        <label class="srow toggle">
          <span class="txt"><span class="label">{t('settings.pdf.avatars', {}, lang)}</span></span>
          <span class="switch">
            <input type="checkbox" checked={pdfIncludeAvatars} on:change={onPdfIncludeAvatarsChange} />
            <span class="track"></span>
          </span>
        </label>
      </div>
    </section>

    <!-- Help & feedback: replay (in-app, chevron) + external links (↗).
         External links open in a new tab so the popup doesn't close under
         the user; rel=noopener keeps window.opener null. -->
    <section class="ac">
      <div class="ah"><LifeBuoy size={14} /> {t('settings.grp.help', {}, lang)}</div>
      <div class="group">
        <button class="srow rowlink" on:click={() => dispatch('replayTour')}>
          <span class="txt"><span class="label">{t('settings.replayTour.btn', {}, lang)}</span></span>
          <span class="chev"><ChevronRight size={16} /></span>
        </button>
        <a class="srow rowlink" href={ISSUES_URL} target="_blank" rel="noopener">
          <span class="txt"><span class="label">{t('settings.about.feedback', {}, lang)}</span></span>
          <span class="chev"><ExternalLink size={14} /></span>
        </a>
        <a class="srow rowlink" href={REVIEW_STORE_URL} target="_blank" rel="noopener">
          <span class="txt"><span class="label">{t('review.rate', {}, lang)}</span></span>
          <span class="chev"><ExternalLink size={14} /></span>
        </a>
      </div>
    </section>

    <!-- About: source code + author attribution. -->
    <section class="ac">
      <div class="ah"><Info size={14} /> {t('settings.about', {}, lang)}</div>
      <div class="group">
        <a class="srow rowlink" href={REPO_URL} target="_blank" rel="noopener">
          <span class="txt"><span class="label">{t('settings.about.source', {}, lang)}</span></span>
          <span class="chev"><ExternalLink size={14} /></span>
        </a>
        <div class="srow author">
          <span class="txt"><span class="label">{t('settings.about.author', {}, lang)}</span></span>
          <a class="author-link" href={AUTHOR_LINKEDIN_URL} target="_blank" rel="noopener">
            {AUTHOR_NAME}<ExternalLink size={11} />
          </a>
        </div>
      </div>
    </section>

    <!-- Muted app/version footer — conventional bottom-of-settings. The
         separator + version live in their own spans so the gap renders
         reliably (Svelte collapses whitespace at an {#if} boundary). -->
    <div class="verfoot">
      <span class="nm">{t('appName', {}, lang)}</span>{#if extensionVersion}<span class="sep">·</span><span class="ver">v{extensionVersion}</span>{/if}
    </div>
  {/if}
</div>

<style>
  /* Accent-titled section card. Tinted uppercase header band + a
     bordered body of rows. All colours resolve from the real popup
     tokens so light/dark track automatically. */
  .ac {
    background: var(--color-surface);
    border: 1px solid var(--color-border);
    border-radius: var(--radius-md);
    overflow: hidden;
    margin-bottom: var(--gap);
  }
  .ah {
    display: flex;
    align-items: center;
    gap: 8px;
    padding: 9px 12px;
    font-size: 12px;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.03em;
    color: var(--color-accent);
    background: var(--color-accent-light);
  }
  .ah :global(svg) {
    width: 14px;
    height: 14px;
    stroke: currentColor;
    fill: none;
    stroke-width: 2;
    flex: 0 0 auto;
  }

  .group {
    display: flex;
    flex-direction: column;
  }

  /* One setting row. Also used for <button>/<a> rows (link/drill-in),
     hence the explicit resets. */
  .srow {
    display: flex;
    align-items: center;
    gap: 12px;
    width: 100%;
    padding: 11px 12px;
    border: 0;
    border-bottom: 1px solid var(--color-border);
    background: transparent;
    color: var(--color-text);
    font: inherit;
    /* Logical `start` so button/link row text mirrors correctly when the
       app switches to dir=rtl for Arabic / Hebrew. */
    text-align: start;
    text-decoration: none;
    box-sizing: border-box;
  }
  .srow:last-child {
    border-bottom: 0;
  }
  .rowlink {
    cursor: pointer;
    transition: background 0.15s ease;
  }
  .rowlink:hover {
    background: var(--color-accent-light);
  }
  .toggle {
    cursor: pointer;
  }
  /* Stacked variant: label block above a full-width control. */
  .srow.stacked {
    flex-direction: column;
    align-items: stretch;
    gap: 8px;
  }

  .txt {
    flex: 1;
    min-width: 0;
    display: flex;
    flex-direction: column;
    gap: 2px;
  }
  .label {
    font-size: 13px;
    color: var(--color-text);
    line-height: 1.3;
  }
  /* Label + inline ⓘ on one line, hint below. */
  .label-row {
    display: flex;
    align-items: center;
    gap: 5px;
  }
  /* Visible info affordance. Muted glyph, accent on hover/focus; carries
     the full explanation via title (hover) + aria-label (focus / AT). */
  .info-btn {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    flex: 0 0 auto;
    width: 16px;
    height: 16px;
    padding: 0;
    border: 0;
    border-radius: 50%;
    background: transparent;
    color: var(--color-subtle);
    font-size: 12px;
    line-height: 1;
    cursor: help;
    transition: color 0.15s ease;
  }
  .info-btn:hover {
    color: var(--color-accent);
  }
  .info-btn:focus-visible {
    outline: 2px solid var(--color-accent);
    outline-offset: 1px;
    color: var(--color-accent);
  }
  .sub {
    font-size: 11px;
    color: var(--color-subtle);
    line-height: 1.35;
  }
  .val {
    flex: 0 0 auto;
    font-size: 12px;
    color: var(--color-subtle);
    max-width: 45%;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }
  .chev {
    flex: 0 0 auto;
    display: inline-flex;
    color: var(--color-subtle);
  }
  .chev :global(svg) {
    stroke: currentColor;
    fill: none;
    stroke-width: 2;
  }

  /* Dimmed row (avatar mode while 'Embed avatars' is off). */
  .srow.is-disabled {
    opacity: 0.6;
  }
  .srow.is-disabled .sub {
    font-style: italic;
  }

  /* Segmented theme control. */
  .seg {
    display: inline-flex;
    flex: 0 0 auto;
    border: 1px solid var(--color-border);
    border-radius: 8px;
    overflow: hidden;
  }
  .seg-btn {
    padding: 6px 14px;
    border: 0;
    background: var(--color-bg);
    color: var(--color-subtle);
    font: inherit;
    font-size: 12px;
    font-weight: 600;
    cursor: pointer;
    transition: background 0.15s ease, color 0.15s ease;
  }
  .seg-btn + .seg-btn {
    border-left: 1px solid var(--color-border);
  }
  .seg-btn:hover:not(.sel) {
    background: var(--color-accent-light);
    color: var(--color-accent);
  }
  .seg-btn.sel {
    background: var(--color-accent);
    color: #fff;
  }

  /* Toggle switch (styled native checkbox). */
  .switch {
    position: relative;
    width: 38px;
    height: 22px;
    flex: 0 0 auto;
    display: inline-block;
  }
  .switch input {
    position: absolute;
    inset: 0;
    width: 100%;
    height: 100%;
    margin: 0;
    opacity: 0;
    cursor: pointer;
  }
  .track {
    position: absolute;
    inset: 0;
    background: var(--color-border);
    border-radius: 999px;
    transition: background 0.15s ease;
  }
  .track::after {
    content: '';
    position: absolute;
    top: 2px;
    left: 2px;
    width: 18px;
    height: 18px;
    background: #fff;
    border-radius: 50%;
    box-shadow: 0 1px 2px rgba(0, 0, 0, 0.25);
    transition: transform 0.15s ease;
  }
  .switch input:checked + .track {
    background: var(--color-accent);
  }
  .switch input:checked + .track::after {
    transform: translateX(16px);
  }

  /* Selects + number input. */
  .field {
    width: 100%;
    font: inherit;
    font-size: 13px;
    padding: 8px 10px;
    border: 1px solid var(--color-border);
    border-radius: var(--radius-md);
    background: var(--color-bg);
    color: var(--color-text);
    cursor: pointer;
    transition: border-color 0.15s ease;
  }
  .field:focus {
    outline: none;
    border-color: var(--color-accent);
  }
  .field:disabled {
    opacity: 0.55;
    cursor: not-allowed;
  }
  .field.mini {
    width: auto;
    flex: 0 0 auto;
    min-width: 96px;
  }
  /* Small number field in a stacked row: compact and left-aligned, not the
     full-width box a free-text field wants. */
  .field.num {
    align-self: flex-start;
    width: auto;
    min-width: 90px;
    max-width: 140px;
  }
  /* −/+ font-size stepper, pinned right. */
  .stepper {
    display: inline-flex;
    align-items: center;
    flex: 0 0 auto;
    border: 1px solid var(--color-border);
    border-radius: var(--radius-md);
    overflow: hidden;
    background: var(--color-bg);
  }
  .step-btn {
    width: 30px;
    height: 32px;
    border: 0;
    background: var(--color-bg);
    color: var(--color-text);
    font: inherit;
    font-size: 16px;
    line-height: 1;
    cursor: pointer;
    transition: background 0.15s ease, color 0.15s ease;
  }
  .step-btn:hover:not(:disabled) {
    background: var(--color-accent-light);
    color: var(--color-accent);
  }
  .step-btn:disabled {
    opacity: 0.4;
    cursor: not-allowed;
  }
  .step-val {
    min-width: 34px;
    text-align: center;
    font-size: 13px;
    font-variant-numeric: tabular-nums;
    color: var(--color-text);
    border-left: 1px solid var(--color-border);
    border-right: 1px solid var(--color-border);
    padding: 6px 4px;
  }

  /* Author attribution row — plain underlined link, no pill. */
  .srow.author .txt {
    flex: 1;
  }
  .author-link {
    flex: 0 0 auto;
    display: inline-flex;
    align-items: center;
    gap: 3px;
    font-size: 13px;
    color: var(--color-text);
    text-decoration: underline;
    text-decoration-color: var(--color-border);
    text-underline-offset: 2px;
    transition: color 0.15s ease, text-decoration-color 0.15s ease;
  }
  .author-link:hover {
    color: var(--color-accent);
    text-decoration-color: var(--color-accent);
  }
  .author-link :global(svg) {
    stroke: currentColor;
    fill: none;
    stroke-width: 2;
  }

  /* Version footer. */
  .verfoot {
    text-align: center;
    color: var(--color-subtle);
    font-size: 12px;
    padding: 16px 12px 6px;
    line-height: 1.6;
  }
  .verfoot .nm {
    font-weight: 600;
    color: var(--color-text);
  }
  .verfoot .sep {
    margin: 0 6px;
    opacity: 0.6;
  }

  /* Language drill-in. */
  .lang-search {
    display: flex;
    align-items: center;
    gap: 8px;
    padding: 8px 10px;
    margin-bottom: var(--gap);
    border: 1px solid var(--color-border);
    border-radius: var(--radius-md);
    background: var(--color-bg);
    transition: border-color 0.15s ease;
  }
  .lang-search:focus-within {
    border-color: var(--color-accent);
  }
  .lang-search :global(svg) {
    width: 15px;
    height: 15px;
    stroke: var(--color-subtle);
    fill: none;
    stroke-width: 2;
    flex: 0 0 auto;
  }
  .lang-search input {
    flex: 1;
    min-width: 0;
    border: 0;
    background: transparent;
    color: var(--color-text);
    font: inherit;
    font-size: 13px;
    outline: none;
  }
  .lang-item.cur .label {
    color: var(--color-accent);
    font-weight: 600;
  }
  .chk {
    flex: 0 0 auto;
    margin-inline-start: auto;
    color: var(--color-accent);
    font-weight: 700;
  }
  .lang-empty {
    padding: 14px 12px;
    font-size: 12px;
    color: var(--color-subtle);
    text-align: center;
  }
</style>
