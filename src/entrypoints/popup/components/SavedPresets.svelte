<script lang="ts">
  import { createEventDispatcher, tick } from 'svelte';
  import { X } from 'lucide-svelte';
  import { t } from '../../../i18n/i18n';
  import type { SavedPreset } from '../../../types/shared';

  export let presets: SavedPreset[] = [];
  export let selectionCount = 0;
  export let lang = 'en';

  const dispatch = createEventDispatcher<{
    save: string;          // new preset name (parent builds the SavedPreset from the live selection)
    apply: SavedPreset;     // parent re-selects this preset's convIds
    remove: string;        // preset id
  }>();

  let open = false;
  let naming = false;
  let name = '';
  let nameInputEl: HTMLInputElement | undefined;
  let rootEl: HTMLDivElement | undefined;

  function toggle() { open = !open; naming = false; name = ''; }
  function close() { open = false; naming = false; name = ''; }

  async function startNaming() {
    if (selectionCount === 0) return;
    naming = true;
    await tick();
    nameInputEl?.focus();
  }
  function confirmSave() {
    const trimmed = name.trim();
    if (!trimmed) return;
    dispatch('save', trimmed);
    name = '';
    naming = false;
  }

  // Outside-dismiss on mousedown, NOT click: clicking the in-menu "Save"
  // button flips to the name input, which reactively removes that button from
  // the DOM in the same event; a click-based check then sees a detached target
  // as "outside" and wrongly closes the menu. mousedown fires before that
  // removal, so rootEl.contains() is still accurate.
  function onWindowPointerDown(e: MouseEvent) {
    if (open && rootEl && !rootEl.contains(e.target as Node)) close();
  }
  function onKeydown(e: KeyboardEvent) {
    if (e.key === 'Escape' && open) close();
  }
</script>

<svelte:window on:mousedown={onWindowPointerDown} on:keydown={onKeydown} />

<div class="presets" bind:this={rootEl}>
  <button
    type="button"
    class="presets-btn"
    class:active={open}
    aria-haspopup="menu"
    aria-expanded={open}
    title={t('presets.button', {}, lang) || 'Presets'}
    on:click|stopPropagation={toggle}
  >
    <span>{t('presets.button', {}, lang) || 'Presets'}</span>
    {#if presets.length > 0}<span class="presets-badge">{presets.length}</span>{/if}
  </button>

  {#if open}
    <div class="presets-menu">
      <div class="presets-save" class:naming>
        {#if !naming}
          <button
            type="button"
            class="presets-save-trigger"
            disabled={selectionCount === 0}
            on:click={startNaming}
          >
            {t('presets.save', { n: selectionCount }, lang) || `Save selection (${selectionCount}) as preset…`}
          </button>
        {:else}
          <div class="presets-name-row">
            <input
              type="text"
              bind:value={name}
              bind:this={nameInputEl}
              placeholder={t('presets.namePlaceholder', {}, lang) || 'Preset name…'}
              on:keydown={(e) => { if (e.key === 'Enter') confirmSave(); }}
            />
            <button type="button" class="presets-confirm" on:click={confirmSave}>
              {t('presets.confirm', {}, lang) || 'Save'}
            </button>
          </div>
        {/if}
      </div>

      {#if presets.length === 0}
        <div class="presets-empty">{t('presets.empty', {}, lang) || 'No saved presets yet'}</div>
      {:else}
        {#each presets as g (g.id)}
          <div class="presets-row">
            <span class="presets-name" title={g.name}>{g.name}</span>
            <span class="presets-count">
              {t('presets.chats', { n: g.convIds.length }, lang) || `${g.convIds.length} chats`}
            </span>
            <button type="button" class="presets-apply" on:click={() => { dispatch('apply', g); close(); }}>
              {t('presets.apply', {}, lang) || 'Apply'}
            </button>
            <button
              type="button"
              class="presets-del"
              title={t('presets.delete', {}, lang) || 'Delete preset'}
              on:click={() => dispatch('remove', g.id)}
            >
              <X size={13} />
            </button>
          </div>
        {/each}
      {/if}
    </div>
  {/if}
</div>

<style>
  .presets { position: relative; display: inline-flex; }
  .presets-btn {
    border: 1px solid var(--color-border);
    background: var(--color-bg);
    color: var(--color-text);
    border-radius: 999px;
    padding: 2px 9px;
    font: inherit;
    font-size: 10px;
    font-weight: 500;
    cursor: pointer;
    display: inline-flex;
    align-items: center;
    gap: 4px;
    white-space: nowrap;
  }
  .presets-btn:hover,
  .presets-btn.active {
    border-color: var(--color-accent);
    color: var(--color-accent);
    background: var(--color-accent-light);
  }
  .presets-badge {
    background: var(--color-accent);
    color: #fff;
    border-radius: 999px;
    font-size: 8px;
    padding: 0 4px;
    line-height: 14px;
  }
  /* Capped to fit inside the picker-body (overflow:hidden), so the menu is
     never clipped at the bottom. */
  .presets-menu {
    position: absolute;
    right: 0;
    top: calc(100% + 6px);
    width: 232px;
    max-height: 240px;
    overflow-y: auto;
    background: var(--color-surface);
    border: 1px solid var(--color-border);
    border-radius: 10px;
    box-shadow: 0 16px 40px -10px rgba(0, 0, 0, 0.28);
    z-index: 40;
  }
  .presets-save { padding: 8px 10px; border-bottom: 1px solid var(--color-border); }
  .presets-save-trigger {
    width: 100%;
    border: 1px solid var(--color-accent);
    background: var(--color-accent);
    color: #fff;
    border-radius: 7px;
    padding: 6px 9px;
    font: inherit;
    font-size: 11px;
    cursor: pointer;
  }
  .presets-save-trigger:disabled { opacity: 0.45; cursor: not-allowed; }
  .presets-name-row { display: flex; gap: 6px; }
  .presets-name-row input {
    flex: 1;
    min-width: 0;
    border: 1px solid var(--color-border-hover);
    border-radius: 7px;
    padding: 5px 8px;
    font: inherit;
    font-size: 12px;
    background: var(--color-surface);
    color: var(--color-text);
    outline: none;
  }
  .presets-confirm {
    border: 1px solid var(--color-accent);
    background: var(--color-accent);
    color: #fff;
    border-radius: 7px;
    padding: 5px 9px;
    font: inherit;
    font-size: 11px;
    cursor: pointer;
  }
  .presets-empty { padding: 10px; text-align: center; color: var(--color-subtle); font-size: 11px; }
  .presets-row {
    display: flex;
    align-items: center;
    gap: 6px;
    padding: 7px 10px;
    border-bottom: 1px solid var(--color-border);
    font-size: 12px;
  }
  .presets-row:last-child { border-bottom: 0; }
  .presets-name { flex: 1; min-width: 0; font-weight: 500; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
  .presets-count { color: var(--color-subtle); font-size: 11px; white-space: nowrap; }
  .presets-apply {
    border: 1px solid var(--color-accent);
    background: transparent;
    color: var(--color-accent);
    border-radius: 7px;
    padding: 3px 9px;
    font: inherit;
    font-size: 11px;
    cursor: pointer;
  }
  .presets-apply:hover { background: var(--color-accent-light); }
  .presets-del {
    border: 0;
    background: transparent;
    color: var(--color-subtle);
    cursor: pointer;
    padding: 1px 4px;
    display: inline-flex;
    align-items: center;
  }
  .presets-del:hover { color: var(--color-danger); }
</style>
