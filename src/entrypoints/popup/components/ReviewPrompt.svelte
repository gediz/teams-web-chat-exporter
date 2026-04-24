<script lang="ts">
  import { createEventDispatcher } from 'svelte';
  import { t } from '../../../i18n/i18n';
  import type { ReviewPromptResponse } from '../../../utils/options';

  export let lang = 'en';
  export let storeUrl = '';
  export let issueUrl = '';

  const dispatch = createEventDispatcher<{ respond: ReviewPromptResponse }>();

  // All three actions open a new tab (or simply dismiss) and fire one
  // event. The parent writes the one-shot flag so subsequent popup
  // opens skip rendering this component entirely.
  function onRate() {
    try { window.open(storeUrl, '_blank', 'noopener,noreferrer'); } catch { /* popup blocker */ }
    dispatch('respond', 'rated');
  }
  function onFeedback() {
    try { window.open(issueUrl, '_blank', 'noopener,noreferrer'); } catch { /* popup blocker */ }
    dispatch('respond', 'feedback');
  }
  function onDismiss() {
    dispatch('respond', 'dismissed');
  }
</script>

<div class="review-prompt" role="region" aria-label="{t('review.prompt', {}, lang) || 'Rate'}">
  <span class="review-prompt-label">{t('review.prompt', {}, lang) || 'Rate Teams Chat Exporter'}</span>
  <div class="review-prompt-actions">
    <button type="button" class="review-prompt-link" on:click={onRate}>{t('review.rate', {}, lang) || 'Rate on store'}</button>
    <span class="review-prompt-sep" aria-hidden="true">·</span>
    <button type="button" class="review-prompt-link" on:click={onFeedback}>{t('review.feedback', {}, lang) || 'Send feedback'}</button>
    <span class="review-prompt-sep" aria-hidden="true">·</span>
    <button type="button" class="review-prompt-link muted" on:click={onDismiss}>{t('review.dismiss', {}, lang) || 'Dismiss'}</button>
  </div>
</div>

<style>
  /* Inline one-liner that slots under the export button. Dashed border
     and muted baseline color deliberately read as "informational", not
     a CTA — it sits passively in the popup and a user who ignores it
     pays no visual tax beyond one row. */
  .review-prompt {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 8px;
    padding: 8px 12px;
    border: 1px dashed var(--color-border);
    border-radius: var(--radius-md, 10px);
    font-size: 11px;
    color: var(--color-subtle);
    margin-bottom: 12px;
  }
  .review-prompt-label {
    white-space: nowrap;
    overflow: hidden;
    text-overflow: clip;
  }
  .review-prompt-actions {
    display: flex;
    align-items: center;
    gap: 4px;
    flex-shrink: 0;
  }
  .review-prompt-link {
    background: transparent;
    border: 0;
    font: inherit;
    font-size: 11px;
    color: var(--color-accent);
    text-decoration: underline;
    text-decoration-color: var(--color-border);
    text-underline-offset: 2px;
    padding: 2px 4px;
    cursor: pointer;
    transition: text-decoration-color 0.15s ease, color 0.15s ease;
  }
  .review-prompt-link:hover {
    text-decoration-color: currentColor;
  }
  .review-prompt-link.muted {
    color: var(--color-subtle);
  }
  .review-prompt-link.muted:hover {
    color: var(--color-text);
  }
  .review-prompt-sep {
    color: var(--color-subtle);
    opacity: 0.6;
  }
</style>
