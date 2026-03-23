<script lang="ts">
  interface Props {
    status: string;
    count: number;
    countLabel?: string;
    isBusy: boolean;
  }

  let { status = 'Ready', count = 0, countLabel = '', isBusy = false }: Props = $props();

  const displayCount = $derived(countLabel || (count > 0 ? count.toLocaleString() : ''));
</script>

<section class="status-section">
  <div class="status-bar">
    <div class="status-icon" class:busy={isBusy}></div>
    <div class="status-text">{status}</div>
    <div class="status-count" style:visibility={displayCount ? 'visible' : 'hidden'}>
      {displayCount}
    </div>
  </div>
</section>
