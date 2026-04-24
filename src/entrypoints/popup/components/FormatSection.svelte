<script lang="ts" context="module">
  export type OptionFormat = 'json' | 'csv' | 'html' | 'txt' | 'pdf';
</script>

<script lang="ts">
  import { createEventDispatcher } from 'svelte';
  import { FileJson2, FileSpreadsheet, FileCode, FileText, FileType, Printer } from 'lucide-svelte';
  import { t } from '../../../i18n/i18n';

  export let formats: OptionFormat[] = ['html'];
  // Downstream signals that determine whether the HTML output has to be
  // packaged as a .zip:
  //   - downloadImages=true attaches an images/ folder
  //   - embedAvatars=true + avatarMode='files' attaches an avatars/ folder
  // Either trigger means "HTML" becomes "HTML.zip". Both controls live
  // in other components (IncludeSection, SettingsPage); we read-only.
  // Defaults keep pre-wired call sites working.
  export let downloadImages = false;
  export let embedAvatars = false;
  export let avatarMode: 'inline' | 'files' = 'inline';
  export let lang = 'en';

  const dispatch = createEventDispatcher<{
    formatsChange: OptionFormat[];
  }>();

  const allFormats = [
    { id: 'json' as const, icon: FileJson2 },
    { id: 'csv' as const, icon: FileSpreadsheet },
    { id: 'html' as const, icon: FileCode },
    { id: 'txt' as const, icon: FileText },
    // Printer is a stand-in for PDF — lucide has no dedicated PDF icon;
    // the printer reads as "print-ready document" which maps cleanly to PDF.
    { id: 'pdf' as const, icon: Printer },
  ];

  // Toggle a single format. Refuses to remove the last selection so we
  // never end up with an empty array — there's no meaningful "no format"
  // export, and the schema invariant is "always non-empty".
  function toggle(id: OptionFormat) {
    const isOn = formats.includes(id);
    if (isOn && formats.length === 1) return;
    const next = isOn ? formats.filter(f => f !== id) : [...formats, id];
    dispatch('formatsChange', next);
  }

  // Pill + "Will save:" metadata.
  //
  //   0 formats                  — can't happen (enforced elsewhere)
  //   1 format, no zip           — plain pill matching the format ("HTML")
  //   1 HTML + images/avatars    — purple "HTML.zip" pill with ⓘ tooltip
  //   2+ formats                 — purple "bundle.zip" pill + ⓘ + contents reveal
  //
  // We don't try to predict the auto-zip-large-HTML branch in
  // download.ts — that's runtime data-size dependent and would lie when
  // the conversation turns out small. Only show the zip variant when
  // the user's selections deterministically produce a zip.
  $: avatarsAsFiles = embedAvatars && avatarMode === 'files';
  $: isHtmlZip = formats.length === 1 && formats[0] === 'html' && (downloadImages || avatarsAsFiles);
  $: isBundle = formats.length >= 2;
  $: isZip = isHtmlZip || isBundle;
  $: pillLabel = isBundle
      ? 'bundle.zip'
      : isHtmlZip
        ? 'HTML.zip'
        : (formats[0] ?? '').toUpperCase();
  // Contents reveal — the ": HTML, JSON, CSV" line that appears after
  // the pill. Bundles list the formats; HTML.zip lists HTML + the
  // asset folders that forced the .zip so the user sees exactly what
  // lives inside the archive.
  $: contentsList = isBundle
      ? formats.map(f => f.toUpperCase()).join(', ')
      : isHtmlZip
        ? ['HTML',
            downloadImages ? (t('format.zipContent.images', {}, lang) || 'images') : null,
            avatarsAsFiles ? (t('format.zipContent.avatars', {}, lang) || 'avatars') : null,
          ].filter(Boolean).join(', ')
        : '';
  // The HTML.zip tooltip is rebuilt from the same content list so any
  // toggle in IncludeSection/SettingsPage flows through to it — e.g.
  // turning 'Embed avatars' on with avatarMode=files extends the
  // tooltip to mention avatars without needing a separate translation.
  $: htmlZipTooltipBase = t('format.htmlZipTooltipBase', {}, lang) || 'HTML plus';
  $: htmlZipTooltipSuffix = t('format.htmlZipTooltipSuffix', {}, lang) || ', packaged as a .zip.';
  $: htmlZipExtras = [
      downloadImages ? (t('format.zipContent.imageFiles', {}, lang) || 'image files') : null,
      avatarsAsFiles ? (t('format.zipContent.avatarFiles', {}, lang) || 'avatar files') : null,
    ].filter(Boolean);
  $: pillTooltip = isBundle
      ? t('format.bundleTooltip', {}, lang) || 'Multiple formats packaged as a .zip.'
      : isHtmlZip
        ? `${htmlZipTooltipBase} ${htmlZipExtras.join(' and ')}${htmlZipTooltipSuffix}`
        : '';
</script>

<section class="format-section" data-lang={lang}>
  <div class="card">
    <div class="card-header">
      <div class="card-icon">
        <FileType size={16} />
      </div>
      <h2 class="card-title">{t('options.format', {}, lang)}</h2>
    </div>
    <div class="format-grid">
      {#each allFormats as fmt}
        {@const Icon = fmt.icon}
        {@const active = formats.includes(fmt.id)}
        {@const onlyOne = active && formats.length === 1}
        <button
          type="button"
          class="format-radio"
          class:active
          aria-pressed={active}
          aria-disabled={onlyOne}
          on:click={() => toggle(fmt.id)}
        >
          <div class="format-icon">
            <Icon size={22} />
          </div>
          <span class="format-label">{t(`format.${fmt.id}`, {}, lang)}</span>
        </button>
      {/each}
    </div>

    <!-- "Will save:" line. Lives on a single wrapped row (white-space:
         nowrap on the outer flex) so it never breaks across lines; the
         contents span has max-width animation so appearing/disappearing
         contents feel continuous rather than popping in. -->
    <div class="will-save">
      <span class="will-save-label">{t('format.willSave', {}, lang) || 'Will save:'}</span>
      <span
        class="will-save-pill"
        class:zip={isZip}
        title={pillTooltip}
      >
        {pillLabel}
        {#if isZip}
          <span class="will-save-info" aria-hidden="true">ⓘ</span>
        {/if}
      </span>
      <span class="will-save-contents" class:visible={isZip} aria-hidden={!isZip}>
        {#if isZip}: {contentsList}{/if}
      </span>
    </div>
  </div>
</section>

<style>
  .will-save {
    display: flex;
    align-items: center;
    gap: 6px;
    margin-top: 10px;
    font-size: 11px;
    white-space: nowrap;
    overflow: hidden;
  }
  .will-save-label {
    color: var(--color-muted);
    font-weight: 500;
  }
  .will-save-pill {
    display: inline-flex;
    align-items: center;
    gap: 4px;
    padding: 2px 7px;
    border-radius: 999px;
    font-size: 10px;
    font-weight: 600;
    letter-spacing: 0.01em;
    background: var(--color-accent-light, rgba(37, 99, 235, 0.1));
    color: var(--color-accent, #2563eb);
  }
  .will-save-pill.zip {
    /* Same purple family as the History page BUNDLE badge so users learn
       one visual cue across the whole extension. */
    background: rgba(109, 40, 217, 0.14);
    color: #6d28d9;
  }
  .will-save-info {
    font-size: 10px;
    opacity: 0.7;
    cursor: help;
  }
  .will-save-contents {
    color: var(--color-muted);
    /* max-width reveal: 0 → 340 over 0.35s. Keeps layout stable when
       toggling between 1-format and bundle, since the space that would
       hold the list is already committed by flex. */
    max-width: 0;
    opacity: 0;
    overflow: hidden;
    transition: max-width 0.35s ease, opacity 0.25s ease;
  }
  .will-save-contents.visible {
    max-width: 340px;
    opacity: 1;
  }
</style>
