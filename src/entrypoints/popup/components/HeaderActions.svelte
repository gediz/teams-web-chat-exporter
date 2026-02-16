<script lang="ts">
  import { createEventDispatcher } from 'svelte';
  import { Settings } from 'lucide-svelte';

  const dispatch = createEventDispatcher<{
    openSettings: void;
  }>();
</script>

<div class="header-actions">
  <button
    class="icon-btn"
    title="Settings"
    on:click={() => dispatch('openSettings')}
  >
    <Settings size={18} />
  </button>
</div>
