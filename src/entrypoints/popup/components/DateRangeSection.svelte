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
  // Inline containers the calendars render INTO (see initCalendars). CSS
  // positions each as an absolute panel above its input, but the panel stays a
  // child of .popup-content, so it scrolls with the page content natively — no
  // detach, no jitter. The library's inputMode popover attaches to <body>
  // instead, which is what caused the #33 clipping and the cross-browser
  // scroll weirdness (Edge clipped it below the capped popup; Chrome pinned it
  // to the viewport; Firefox jittered while re-placing it every frame).
  // startPopEl/endPopEl are the panels WE own (the .cal-pop wrapper we toggle).
  // startCalEl/endCalEl are throwaway inner divs the library takes over: it
  // REPLACES the element handed to `new Calendar` with its own calendar node,
  // so we must not hand it our wrapper or it would strip the wrapper's class
  // and our show/hide control with it.
  let startPopEl: HTMLDivElement;
  let endPopEl: HTMLDivElement;
  let startCalEl: HTMLDivElement;
  let endCalEl: HTMLDivElement;
  let startCalendar: any = null;
  let endCalendar: any = null;
  // Which picker is open, if any. We own visibility instead of the library.
  let openField: "start" | "end" | null = null;
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
    const wrapper = inputEl.closest(".date-input-wrapper");
    if (wrapper) {
      // Restart the shake on every invalid attempt. Re-adding a class that is
      // already present does not replay a CSS animation, so remove it, force a
      // reflow, then add it back.
      wrapper.classList.remove("validation-error");
      void (wrapper as HTMLElement).offsetWidth;
      wrapper.classList.add("validation-error");
      // Clear any existing timer, then drop the class again after 1.5s.
      if (validationErrorTimer) clearTimeout(validationErrorTimer);
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
    closeCalendar();

    // Flash animation
    if (startInputEl) startInputEl.classList.add("flash");
    if (endInputEl) endInputEl.classList.add("flash");
    setTimeout(() => {
      if (startInputEl) startInputEl.classList.remove("flash");
      if (endInputEl) endInputEl.classList.remove("flash");
    }, 600);
  }

  function closeCalendar() {
    openField = null;
  }

  function toggleCalendar(field: "start" | "end") {
    openField = openField === field ? null : field;
  }

  function onInputKeydown(e: KeyboardEvent, field: "start" | "end") {
    // The inputs are readonly, so Enter/Space are ours to open the picker.
    if (e.key === "Enter" || e.key === " ") {
      e.preventDefault();
      toggleCalendar(field);
    }
  }

  // Close the open picker on a click outside it. We listen for `click`, not
  // `mousedown`, so dragging the popup scrollbar (mousedown + move + mouseup,
  // no click) doesn't close the panel while the user scrolls. Capture phase so
  // we run before the target's own handlers. Clicks on an input are left to
  // that input's toggle; clicks inside the open panel (day cells, month nav)
  // stay open.
  function handleDocumentClick(e: MouseEvent) {
    if (openField === null) return;
    const target = e.target as Node;
    if (startInputEl?.contains(target) || endInputEl?.contains(target)) return;
    if (startPopEl?.contains(target) || endPopEl?.contains(target)) return;
    closeCalendar();
  }

  function handleKeydown(e: KeyboardEvent) {
    if (e.key === "Escape") closeCalendar();
  }

  async function initCalendars() {
    const { Calendar } = await import("vanilla-calendar-pro");

    // Map lang to calendar locale
    const calendarLocale = lang === "zh-CN" ? "zh" : lang.split("-")[0];

    // themeAttrDetect runs an internal MutationObserver on the given selector,
    // so the calendar tracks app theme toggles automatically — no manual sync.
    const themeAttrDetect = "body[data-theme]";

    if (startCalEl) {
      // Inline render (inputMode omitted): the library takes over startCalEl
      // inside our .cal-pop wrapper. We show/hide the wrapper via CSS; the
      // library does no positioning of its own in this mode.
      startCalendar = new Calendar(startCalEl, {
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
            closeCalendar();
          }
        },
      });
      startCalendar.init();
    }

    if (endCalEl) {
      endCalendar = new Calendar(endCalEl, {
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
            closeCalendar();
          }
        },
      });
      endCalendar.init();
    }
  }

  onMount(() => {
    initCalendars();
    document.addEventListener("click", handleDocumentClick, true);
    document.addEventListener("keydown", handleKeydown);
  });

  // Tear down calendar instances when the component unmounts. The library
  // keeps a static memoizedElements map keyed by the element; if we don't
  // destroy on unmount, remounting the section (e.g. after the user visits the
  // settings page) leaves a stale binding that prevents the freshly-constructed
  // Calendar from picking up our seeded selectedDates.
  onDestroy(() => {
    document.removeEventListener("click", handleDocumentClick, true);
    document.removeEventListener("keydown", handleKeydown);
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

<div class="card" data-lang={lang} data-tour="range">
  <div class="card-header">
    <div class="card-icon">
      <Calendar size={16} />
    </div>
    <h2 class="card-title">{t("range.title", {}, lang)}</h2>
  </div>

  <div class="date-chips">
    {#each ranges as qr (qr.key)}
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
          role="combobox"
          aria-haspopup="dialog"
          aria-expanded={openField === "start"}
          aria-controls="start-cal-pop"
          on:click={() => toggleCalendar("start")}
          on:keydown={(e) => onInputKeydown(e, "start")}
        />
        <div
          bind:this={startPopEl}
          id="start-cal-pop"
          class="cal-pop"
          class:open={openField === "start"}
        >
          <div bind:this={startCalEl}></div>
        </div>
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
          role="combobox"
          aria-haspopup="dialog"
          aria-expanded={openField === "end"}
          aria-controls="end-cal-pop"
          on:click={() => toggleCalendar("end")}
          on:keydown={(e) => onInputKeydown(e, "end")}
        />
        <div
          bind:this={endPopEl}
          id="end-cal-pop"
          class="cal-pop"
          class:open={openField === "end"}
        >
          <div bind:this={endCalEl}></div>
        </div>
      </div>
    </div>
  </div>
</div>
