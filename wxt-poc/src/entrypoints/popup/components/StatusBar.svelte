<script lang="ts">
  interface Props {
    status: string;
    count: number;
    isBusy: boolean;
  }

  let { status = 'Ready', count = 0, isBusy = false }: Props = $props();
</script>

<section class="status-section">
  <div class="status-bar">
    <div class="status-icon" class:busy={isBusy}></div>
    <div class="status-text">{status}</div>
    <div class="status-count" style:visibility={count > 0 ? 'visible' : 'hidden'}>
      {count}
    </div>
  </div>
</section>
