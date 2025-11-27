<script lang="ts">
  import { createEventDispatcher } from 'svelte';
  import { MessageSquareQuote, Smile, Bell, CircleUserRound, ListChecks } from 'lucide-svelte';
  import { t } from '../../../i18n/i18n';

  export let includeReplies = true;
  export let includeReactions = true;
  export let includeSystem = false;
  export let embedAvatars = false;
  export let lang = 'en';
  export let disableReplies = false;
  export let disableReactions = false;
  export let disableAvatars = false;

  const dispatch = createEventDispatcher<{
    includeRepliesChange: boolean;
    includeReactionsChange: boolean;
    includeSystemChange: boolean;
    embedAvatarsChange: boolean;
  }>();
</script>

<div class="card" data-lang={lang}>
  <div class="card-header">
    <div class="card-icon">
      <ListChecks size={16} />
    </div>
    <h2 class="card-title">{t('include.title', {}, lang)}</h2>
  </div>

  <div class="checkbox-list">
    <label class="checkbox-item">
      <div class="checkbox-content">
        <div class="checkbox-icon">
          <MessageSquareQuote size={18} />
        </div>
        <span class="checkbox-label">{t('include.replies', {}, lang)}</span>
      </div>
      <input
        type="checkbox"
        checked={includeReplies}
        disabled={disableReplies}
        on:change={(e) => dispatch('includeRepliesChange', e.currentTarget.checked)}
      />
    </label>

    <label class="checkbox-item">
      <div class="checkbox-content">
        <div class="checkbox-icon">
          <Smile size={18} />
        </div>
        <span class="checkbox-label">{t('include.reactions', {}, lang)}</span>
      </div>
      <input
        type="checkbox"
        checked={includeReactions}
        disabled={disableReactions}
        on:change={(e) => dispatch('includeReactionsChange', e.currentTarget.checked)}
      />
    </label>

    <label class="checkbox-item">
      <div class="checkbox-content">
        <div class="checkbox-icon">
          <Bell size={18} />
        </div>
        <span class="checkbox-label">{t('include.system', {}, lang)}</span>
      </div>
      <input
        type="checkbox"
        checked={includeSystem}
        on:change={(e) => dispatch('includeSystemChange', e.currentTarget.checked)}
      />
    </label>

    <label class="checkbox-item">
      <div class="checkbox-content">
        <div class="checkbox-icon">
          <CircleUserRound size={18} />
        </div>
        <span class="checkbox-label">{t('include.avatars', {}, lang)}</span>
      </div>
      <input
        type="checkbox"
        checked={embedAvatars}
        disabled={disableAvatars}
        on:change={(e) => dispatch('embedAvatarsChange', e.currentTarget.checked)}
      />
    </label>
  </div>
</div>
