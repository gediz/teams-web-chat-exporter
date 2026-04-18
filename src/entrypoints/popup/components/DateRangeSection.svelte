<script lang="ts" context="module">
  export type QuickRange = { key: string; label: string; icon: string };
</script>

<script lang="ts">
  import { createEventDispatcher } from "svelte";
  import { Calendar } from "lucide-svelte";
  import { onMount, onDestroy } from "svelte";
  import { t } from "../../../i18n/i18n";

  export let startAt = "";
  export let endAt = "";
  export let activeRange = "none";
  export let ranges: QuickRange[] = [];
  export let lang = "en";
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

  // Build the calendar's selectedDates/Month/Year from any date-bearing string.
  // Returns null when the value is empty or unparseable. Accepts both the raw
  // calendar format ("YYYY-MM-DD") and the stored local form ("YYYY-MM-DD HH:MM"
  // produced by isoToLocalInput) — extracts the date portion regardless.
  // Typed `any` because the library's `selectedMonth` is a strict Range<12>
  // tuple, which a runtime Date.getMonth() can't be statically narrowed to.
  function dateParts(input: string): any {
    if (!input) return null;
    const m = input.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (!m) return null;
    const iso = `${m[1]}-${m[2]}-${m[3]}`;
    return {
      selectedDates: [iso],
      selectedMonth: Number(m[2]) - 1,
      selectedYear: Number(m[1]),
    };
  }

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

    // Map lang to calendar locale
    const calendarLocale = lang === "zh-CN" ? "zh" : lang.split("-")[0];

    // themeAttrDetect runs an internal MutationObserver on the given selector,
    // so the calendar tracks app theme toggles automatically — no manual sync.
    const themeAttrDetect = "body[data-theme]";

    if (startInputEl) {
      startCalendar = new Calendar(startInputEl, {
        inputMode: true,
        themeAttrDetect,
        locale: calendarLocale,
        ...(dateParts(startAt) || {}),
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
        themeAttrDetect,
        locale: calendarLocale,
        ...(dateParts(endAt) || {}),
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

  // Tear down calendar instances when the component unmounts. The library
  // keeps a static memoizedElements map keyed by the input element; if we
  // don't destroy on unmount, remounting the section (e.g. after the user
  // visits the settings page) leaves a stale binding that prevents the
  // freshly-constructed Calendar from picking up our seeded selectedDates.
  onDestroy(() => {
    try { startCalendar?.destroy(); } catch { /* noop */ }
    try { endCalendar?.destroy(); } catch { /* noop */ }
    startCalendar = null;
    endCalendar = null;
  });

  // Reposition a calendar to match the given ISO date (or today when empty).
  // The library's option-level selectedDates is only set at init; manual
  // user picks live in context, not options. So any later update() that
  // doesn't re-push the option-level value would wipe the visible selection.
  function syncCalendar(cal: any, iso: string) {
    if (!cal) return;
    const p = dateParts(iso);
    if (p) {
      cal.selectedDates = p.selectedDates;
      cal.selectedMonth = p.selectedMonth;
      cal.selectedYear = p.selectedYear;
    } else {
      cal.selectedDates = [];
      cal.selectedMonth = new Date().getMonth();
      cal.selectedYear = new Date().getFullYear();
    }
    cal.update({ dates: true, month: true, year: true });
  }

  // Update date displays + calendar selection when dates or language change
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

    // Keep the calendars' own selectedDates in sync so reopening lands on
    // the saved date and theme switches don't clear manual picks.
    syncCalendar(startCalendar, _startAt);
    syncCalendar(endCalendar, _endAt);
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
