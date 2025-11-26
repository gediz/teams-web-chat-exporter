<script lang="ts" context="module">
  export type QuickRange = { key: string; label: string; icon: string };
</script>

<script lang="ts">
  import { createEventDispatcher } from 'svelte';
  import { Calendar } from 'lucide-svelte';
  import { onMount } from 'svelte';
  import { t } from '../../../i18n/i18n';

  export let startAt = '';
  export let endAt = '';
  export let activeRange = 'none';
  export let ranges: QuickRange[] = [];
  export let lang = 'en';
  export let theme: 'light' | 'dark' = 'light';

  const dispatch = createEventDispatcher<{
    changeStart: string;
    changeEnd: string;
    quickSelect: string;
  }>();

  let startInputEl: HTMLInputElement;
  let endInputEl: HTMLInputElement;
  let startCalendar: any = null;
  let endCalendar: any = null;

  function formatDateForDisplay(dateStr: string): string {
    if (!dateStr) return '';
    const date = new Date(dateStr);
    // Use the current language for date formatting
    const locale = lang || 'en';
    return date.toLocaleDateString(locale, { month: 'short', day: 'numeric', year: 'numeric' });
  }

  function handleQuickSelect(key: string) {
    dispatch('quickSelect', key);

    // Flash animation
    if (startInputEl) startInputEl.classList.add('flash');
    if (endInputEl) endInputEl.classList.add('flash');
    setTimeout(() => {
      if (startInputEl) startInputEl.classList.remove('flash');
      if (endInputEl) endInputEl.classList.remove('flash');
    }, 600);
  }

  async function initCalendars() {
    const { Calendar } = await import('vanilla-calendar-pro');
    const isDark = document.body.dataset.theme === 'dark';
    const today = new Date().toISOString().split('T')[0];

    if (startInputEl) {
      startCalendar = new Calendar(startInputEl, {
        inputMode: true,
        selectedTheme: isDark ? 'dark' : 'light',
        dateMax: today as any,
        onClickDate(self: any) {
          const selectedDate = self.context.selectedDates?.[0];
          if (selectedDate) {
            // Validate: start date cannot be after end date
            if (endAt && selectedDate > endAt) {
              return; // Don't allow selection
            }
            dispatch('changeStart', selectedDate);
          }
        }
      });
      startCalendar.init();
    }

    if (endInputEl) {
      endCalendar = new Calendar(endInputEl, {
        inputMode: true,
        selectedTheme: isDark ? 'dark' : 'light',
        dateMax: today as any,
        onClickDate(self: any) {
          const selectedDate = self.context.selectedDates?.[0];
          if (selectedDate) {
            // Validate: end date cannot be before start date
            if (startAt && selectedDate < startAt) {
              return; // Don't allow selection
            }
            dispatch('changeEnd', selectedDate);
          }
        }
      });
      endCalendar.init();
    }
  }

  onMount(() => {
    initCalendars();
  });

  // Update calendar theme when app theme changes
  $: if (startCalendar && endCalendar && theme) {
    const selectedTheme = theme === 'dark' ? 'dark' : 'light';
    if (startCalendar.settings) {
      startCalendar.settings.selectedTheme = selectedTheme;
      startCalendar.update();
    }
    if (endCalendar.settings) {
      endCalendar.settings.selectedTheme = selectedTheme;
      endCalendar.update();
    }
  }

  // Update date displays when language or dates change
  $: if (startInputEl) {
    startInputEl.value = formatDateForDisplay(startAt);
  }
  $: if (endInputEl) {
    endInputEl.value = formatDateForDisplay(endAt);
  }
</script>

<div class="card" data-lang={lang}>
  <div class="card-header">
    <div class="card-icon">
      <Calendar size={16} />
    </div>
    <h2 class="card-title">{t('range.title', {}, lang)}</h2>
  </div>

  <div class="date-chips">
    {#each ranges as qr}
      <button
        type="button"
        class="date-chip"
        class:active={activeRange === qr.key}
        data-range={qr.key}
        on:click={() => handleQuickSelect(qr.key)}
      >
        {qr.label}
      </button>
    {/each}
  </div>

  <div class="date-fields">
    <div class="date-field">
      <label for="start-date">{t('range.from', {}, lang)}</label>
      <div class="date-input-wrapper">
        <div class="date-input-icon">
          <Calendar size={14} />
        </div>
        <input
          bind:this={startInputEl}
          type="text"
          id="start-date"
          class="date-input"
          value={formatDateForDisplay(startAt)}
          placeholder={t('range.placeholder', {}, lang)}
          readonly
        />
      </div>
    </div>

    <div class="date-field">
      <label for="end-date">{t('range.to', {}, lang)}</label>
      <div class="date-input-wrapper">
        <div class="date-input-icon">
          <Calendar size={14} />
        </div>
        <input
          bind:this={endInputEl}
          type="text"
          id="end-date"
          class="date-input"
          value={formatDateForDisplay(endAt)}
          placeholder={t('range.placeholder', {}, lang)}
          readonly
        />
      </div>
    </div>
  </div>
</div>
