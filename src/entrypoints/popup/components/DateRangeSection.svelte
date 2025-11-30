<script lang="ts" context="module">
  export type QuickRange = { key: string; label: string; icon: string };
</script>

<script lang="ts">
  import { createEventDispatcher } from "svelte";
  import { Calendar } from "lucide-svelte";
  import { onMount } from "svelte";
  import { t } from "../../../i18n/i18n";

  export let startAt = "";
  export let endAt = "";
  export let activeRange = "none";
  export let ranges: QuickRange[] = [];
  export let lang = "en";
  export let theme: "light" | "dark" = "light";
  export let highlightMode: "none" | "quick-range" | "manual" = "none";

  const dispatch = createEventDispatcher<{
    changeStart: string;
    changeEnd: string;
    quickSelect: string;
  }>();

  let startInputEl: HTMLInputElement;
  let endInputEl: HTMLInputElement;
  let startCalendar: any = null;
  let endCalendar: any = null;
  let validationErrorTimer: ReturnType<typeof setTimeout> | null = null;

  function showValidationError(inputEl: HTMLInputElement) {
    // Add error class for visual feedback
    const wrapper = inputEl.closest(".date-input-wrapper");
    if (wrapper) {
      wrapper.classList.add("validation-error");
      // Clear any existing timer
      if (validationErrorTimer) clearTimeout(validationErrorTimer);
      // Remove error class after 1.5 seconds
      validationErrorTimer = setTimeout(() => {
        wrapper.classList.remove("validation-error");
      }, 1500);
    }
  }

  function formatDateForDisplay(dateStr: string): string {
    if (!dateStr) return "";
    const date = new Date(dateStr);
    // Use the current language for date formatting
    const locale = lang || "en";
    return date.toLocaleDateString(locale, {
      month: "short",
      day: "numeric",
      year: "numeric",
    });
  }

  function handleQuickSelect(key: string) {
    dispatch("quickSelect", key);

    // Flash animation
    if (startInputEl) startInputEl.classList.add("flash");
    if (endInputEl) endInputEl.classList.add("flash");
    setTimeout(() => {
      if (startInputEl) startInputEl.classList.remove("flash");
      if (endInputEl) endInputEl.classList.remove("flash");
    }, 600);
  }

  async function initCalendars() {
    const { Calendar } = await import("vanilla-calendar-pro");
    const isDark = document.body.dataset.theme === "dark";
    const today = new Date().toISOString().split("T")[0];

    // Map lang to calendar locale
    const calendarLocale = lang === "zh-CN" ? "zh" : lang.split("-")[0];

    if (startInputEl) {
      startCalendar = new Calendar(startInputEl, {
        inputMode: true,
        selectedTheme: isDark ? "dark" : "light",
        dateMax: today as any,
        locale: calendarLocale,
        onClickDate(self: any) {
          const selectedDate = self.context.selectedDates?.[0];
          if (selectedDate) {
            // Validate: start date cannot be after end date
            if (endAt && selectedDate > endAt) {
              showValidationError(startInputEl);
              return; // Don't allow selection, but keep calendar open
            }
            dispatch("changeStart", selectedDate);
            // Hide calendar after valid selection
            self.hide();
          }
        },
      });
      startCalendar.init();
    }

    if (endInputEl) {
      endCalendar = new Calendar(endInputEl, {
        inputMode: true,
        selectedTheme: isDark ? "dark" : "light",
        dateMax: today as any,
        locale: calendarLocale,
        onClickDate(self: any) {
          const selectedDate = self.context.selectedDates?.[0];
          if (selectedDate) {
            // Validate: end date cannot be before start date
            if (startAt && selectedDate < startAt) {
              showValidationError(endInputEl);
              return; // Don't allow selection, but keep calendar open
            }
            dispatch("changeEnd", selectedDate);
            // Hide calendar after valid selection
            self.hide();
          }
        },
      });
      endCalendar.init();
    }
  }

  onMount(() => {
    initCalendars();
  });

  // Update date displays when dates or language change
  $: {
    // Track all dependencies explicitly
    const _startAt = startAt;
    const _endAt = endAt;
    const _lang = lang;

    // Update both inputs whenever any dependency changes
    if (startInputEl) {
      startInputEl.value = formatDateForDisplay(_startAt);
    }
    if (endInputEl) {
      endInputEl.value = formatDateForDisplay(_endAt);
    }
  }
</script>

<div class="card" data-lang={lang}>
  <div class="card-header">
    <div class="card-icon">
      <Calendar size={16} />
    </div>
    <h2 class="card-title">{t("range.title", {}, lang)}</h2>
  </div>

  <div class="date-chips">
    {#each ranges as qr}
      <button
        type="button"
        class="date-chip"
        class:active={activeRange === qr.key &&
          (qr.key === "none"
            ? highlightMode === "none"
            : highlightMode === "quick-range")}
        data-range={qr.key}
        on:click={() => handleQuickSelect(qr.key)}
      >
        {qr.label}
      </button>
    {/each}
  </div>

  <div class="date-fields">
    <div class="date-field">
      <label for="start-date">{t("range.from", {}, lang)}</label>
      <div class="date-input-wrapper">
        <div class="date-input-icon">
          <Calendar size={14} />
        </div>
        <input
          bind:this={startInputEl}
          type="text"
          id="start-date"
          class="date-input"
          class:has-value={highlightMode === "manual" && !!startAt}
          placeholder={t("range.placeholder", {}, lang)}
          readonly
        />
      </div>
    </div>

    <div class="date-field">
      <label for="end-date">{t("range.to", {}, lang)}</label>
      <div class="date-input-wrapper">
        <div class="date-input-icon">
          <Calendar size={14} />
        </div>
        <input
          bind:this={endInputEl}
          type="text"
          id="end-date"
          class="date-input"
          class:has-value={highlightMode === "manual" && !!endAt}
          placeholder={t("range.placeholder", {}, lang)}
          readonly
        />
      </div>
    </div>
  </div>
</div>
