<script lang="ts" context="module">
  export type OptionFormat = 'json' | 'csv' | 'html' | 'txt';
</script>

<script lang="ts">
  import { createEventDispatcher } from 'svelte';
  import { t } from '../../../i18n/i18n';

  export let format: OptionFormat = 'json';
  export let lang = 'en';
  export let includeReplies = true;
  export let includeReactions = true;
  export let includeSystem = false;
  export let embedAvatars = false;

  const dispatch = createEventDispatcher<{
    langChange: string;
    formatChange: OptionFormat;
    includeRepliesChange: boolean;
    includeReactionsChange: boolean;
    includeSystemChange: boolean;
    embedAvatarsChange: boolean;
  }>();
</script>

<section class="card" aria-labelledby="options-section" data-lang={lang}>
  <div class="section-head">
    <h2 class="section-title" id="options-section">{t('options.title')}</h2>
    <p class="section-sub">{t('options.subtitle')}</p>
  </div>
  <div class="field">
    <label for="lang">{t('lang.label')}</label>
    <select id="lang" value={lang} on:change={(e) => dispatch('langChange', (e.currentTarget as HTMLSelectElement).value)}>
      <option value="en">English</option>
      <option value="zh-CN">简体中文</option>
      <option value="pt-BR">Português (Brasil)</option>
      <option value="nl">Nederlands</option>
      <option value="fr">Français</option>
      <option value="de">Deutsch</option>
      <option value="it">Italiano</option>
      <option value="ja">日本語</option>
      <option value="ko">한국어</option>
      <option value="ru">Русский</option>
      <option value="es">Español</option>
      <option value="tr">Türkçe</option>
      <option value="ar">العربية</option>
      <option value="he">עברית</option>
    </select>
  </div>
  <div class="field">
    <label for="format">{t('options.format')}</label>
    <select
      id="format"
      value={format}
      on:change={(e) => dispatch('formatChange', (e.currentTarget as HTMLSelectElement).value as OptionFormat)}
    >
      <option value="json">{t('format.json')}</option>
      <option value="csv">{t('format.csv')}</option>
      <option value="html">{t('format.html')}</option>
      <option value="txt">{t('format.txt')}</option>
    </select>
  </div>

  <div class="toggle-list" role="group" aria-label="Include options">
    <label class="toggle">
      <span class="toggle-label">
        <span class="toggle-icon">↩</span>
        <span>{t('options.replies')}</span>
      </span>
      <input
        id="includeReplies"
        type="checkbox"
        checked={includeReplies}
        on:change={(e) => dispatch('includeRepliesChange', (e.currentTarget as HTMLInputElement).checked)}
      />
    </label>
    <label class="toggle">
      <span class="toggle-label">
        <span class="toggle-icon">😊</span>
        <span>{t('options.reactions')}</span>
      </span>
      <input
        id="includeReactions"
        type="checkbox"
        checked={includeReactions}
        on:change={(e) => dispatch('includeReactionsChange', (e.currentTarget as HTMLInputElement).checked)}
      />
    </label>
    <label class="toggle">
      <span class="toggle-label">
        <span class="toggle-icon">⚙</span>
        <span>{t('options.system')}</span>
      </span>
      <input
        id="includeSystem"
        type="checkbox"
        checked={includeSystem}
        on:change={(e) => dispatch('includeSystemChange', (e.currentTarget as HTMLInputElement).checked)}
      />
    </label>
    <label class="toggle">
      <span class="toggle-label">
        <span class="toggle-icon">👤</span>
        <span>{t('options.avatars')}</span>
      </span>
      <input
        id="embedAvatars"
        type="checkbox"
        checked={embedAvatars}
        on:change={(e) => dispatch('embedAvatarsChange', (e.currentTarget as HTMLInputElement).checked)}
      />
    </label>
  </div>
</section>
