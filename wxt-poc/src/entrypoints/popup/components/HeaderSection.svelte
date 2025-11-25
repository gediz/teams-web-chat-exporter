<script lang="ts">
  import { t } from '../../../i18n/i18n';

  export let lang = 'en';
</script>

<header class="header" data-lang={lang}>
  <div class="header-top">
    <h1>{t('title.app', {}, lang) || 'Teams Chat Exporter'}</h1>
  </div>
  <p>{t('title.tagline', {}, lang) || ''}</p>
</header>
