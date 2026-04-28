<script lang="ts" context="module">
  // Step definitions live in a module-scope const so the type stays
  // narrow. Each step names its target via the `data-tour="<key>"`
  // attribute set on the corresponding element in App.svelte / sub-
  // components, so neither side has to hold component refs across the
  // boundary. The overlay queries the live DOM at step change.
  type StepKey = 'format' | 'picker' | 'folder' | 'range' | 'history' | 'settings' | 'stop';
  type StepDef = {
    target: StepKey;        // value of data-tour=
    titleKey: string;
    bodyKey: string;
    titleFallback: string;
    bodyFallback: string;
    /** Force-expand the picker for this step (only matters for 'folder'). */
    needsExpandedPicker?: boolean;
    /** Card position relative to popup. Defaults to 'bottom'. */
    cardPosition?: 'top' | 'bottom';
  };
  export const STEPS: ReadonlyArray<StepDef> = [
    { target: 'format',
      titleKey: 'onboarding.format.title',
      bodyKey:  'onboarding.format.body',
      titleFallback: 'Pick one or more formats',
      bodyFallback:  'Tap any format card to toggle it. When two or more are selected, the exports are packaged together as a single .zip.',
      cardPosition: 'top' },
    { target: 'picker',
      titleKey: 'onboarding.picker.title',
      bodyKey:  'onboarding.picker.body',
      titleFallback: 'Need to export multiple chats?',
      bodyFallback:  'Click the chevron to expand the conversation picker. Tick as many chats as you want — they\'ll all come back as one bundled .zip with a folder per chat.',
      cardPosition: 'bottom' },
    { target: 'folder',
      titleKey: 'onboarding.folder.title',
      bodyKey:  'onboarding.folder.body',
      titleFallback: 'Filter by folder',
      bodyFallback:  'The rail on the left filters by chat type or by your Teams folders. Useful when you have hundreds of chats and only want to bulk-export one team\'s.',
      needsExpandedPicker: true,
      cardPosition: 'bottom' },
    { target: 'range',
      titleKey: 'onboarding.range.title',
      bodyKey:  'onboarding.range.body',
      titleFallback: 'Pick a date range',
      bodyFallback:  'Quick presets for the last 24 h, 7 d, or 30 d, plus ∞ for no filter (everything). Use the date inputs below the chips for a specific window.',
      cardPosition: 'top' },
    { target: 'history',
      titleKey: 'onboarding.history.title',
      bodyKey:  'onboarding.history.body',
      titleFallback: 'Past exports live here',
      bodyFallback:  'The clock icon opens your export history — re-open any saved file or jump to its folder. Failed exports get an amber badge.',
      cardPosition: 'bottom' },
    { target: 'settings',
      titleKey: 'onboarding.settings.title',
      bodyKey:  'onboarding.settings.body',
      titleFallback: 'Defaults, language, and more',
      bodyFallback:  'The gear icon opens settings: change the language (24 supported), pick what happens after an export, tweak defaults.',
      cardPosition: 'bottom' },
    { target: 'stop',
      titleKey: 'onboarding.stop.title',
      bodyKey:  'onboarding.stop.body',
      titleFallback: 'Export is reversible',
      bodyFallback:  'Click Export to begin. During an export the button turns red — click it again any time to stop. The thin bar at the bottom shows progress.',
      cardPosition: 'top' },
  ] as const;
</script>

<script lang="ts">
  import { createEventDispatcher, onMount, onDestroy, tick } from 'svelte';
  import { X, Sparkles } from 'lucide-svelte';
  import { t } from '../../../i18n/i18n';

  export let lang = 'en';
  // Two-way bound from the parent so the 'folder' step can temporarily
  // expand the picker and restore on tour end.
  export let pickerCollapsed = false;
  // Skip the pre-tour prompt and start the walkthrough immediately.
  // Used when the user explicitly clicks "Replay tour" in Settings —
  // they've already opted in, no need to ask again.
  export let autoStart = false;

  const dispatch = createEventDispatcher<{ dismiss: void }>();

  // Two phases:
  //   'prompt' — initial pre-tour card asking the user whether they
  //              want the tour at all. Some users know exactly what
  //              the extension does and don't need a walkthrough; we
  //              respect that without forcing them to click X.
  //   'tour'   — the actual 7-step walkthrough.
  // The 'prompt' phase is bypassed entirely when the parent invokes
  // the tour from Settings → Replay (autoStart prop).
  let phase: 'prompt' | 'tour' = 'prompt';
  let stepIdx = 0;
  // Snapshot the picker's collapsed state at tour start so we can
  // restore it on dismiss. Otherwise users who started with the picker
  // collapsed would find it stuck open after the folder step.
  let initialPickerCollapsed = pickerCollapsed;

  $: step = STEPS[stepIdx];

  // Reactive: apply the highlight class to the current step's target
  // and toggle the picker open/closed as the step requires. Runs after
  // tick() so the picker's expanded state has time to render before we
  // try to find the target inside it.
  let highlighted: HTMLElement | null = null;
  // Card position is recomputed per step based on the target's actual
  // location after scrollIntoView lands. Falls back to the step's hint
  // if we can't measure (e.g. target not yet in DOM).
  let cardSide: 'top' | 'bottom' = 'bottom';
  // Bounding rect of the hole the scrim should leave bright around the
  // highlighted target. null ⇒ no hole (full dim). Used to feed the
  // SVG mask in the template — the mask's blurred black rect creates
  // a soft-edged transparent hole through the dim.
  let hole: { x: number; y: number; w: number; h: number; rx: number } | null = null;
  // Tracks the viewport size so the SVG's <rect width=...> stays
  // pixel-accurate (using "100%" works too, but explicit values let
  // us round and avoid sub-pixel filter artefacts on some browsers).
  let viewW = 0;
  let viewH = 0;
  // Blur radius of the hole's edge feather, in CSS px. 3 reads as a
  // gentle anti-aliasing softening rather than a glow, which keeps
  // the dim feeling crisp instead of muddy. The SVG filter region is
  // sized 2× so the blur never gets clipped at its own bounds.
  const HOLE_BLUR = 3;
  // Only apply step targeting once we're actually in the tour phase.
  // During the 'prompt' phase no element is highlighted and the scrim
  // is just a full dim — no DOM target to find.
  $: if (phase === 'tour') void applyStep(step);
  async function applyStep(s: StepDef) {
    // 1) Coordinate picker open/closed for this step. With bind: in
    // the parent the assignment alone propagates back — no extra
    // dispatch needed.
    const wantOpen = !!s.needsExpandedPicker;
    if (wantOpen && pickerCollapsed) {
      pickerCollapsed = false;
    } else if (!wantOpen && !pickerCollapsed && initialPickerCollapsed && s.target !== 'picker') {
      // If user originally had it collapsed AND this step doesn't need
      // it open, fold it back. Skip the 'picker' step since that step
      // is specifically about the picker — leaving it as-is feels
      // natural there (user can see what we're talking about).
      pickerCollapsed = true;
    }
    // 2) Wait for DOM (picker expand transition + Svelte rerender).
    await tick();
    // 3) Re-target the highlight.
    if (highlighted) highlighted.classList.remove('onb-target');
    highlighted = document.querySelector<HTMLElement>(`[data-tour="${s.target}"]`);
    if (!highlighted) {
      cardSide = s.cardPosition || 'bottom';
      hole = null;
      return;
    }
    highlighted.classList.add('onb-target');
    // 4) Scroll the target into the popup-content's viewport so users
    // never have to scroll manually to find what we're highlighting.
    // 'nearest' avoids unnecessary motion when target is already in
    // view; the browser also handles the scroll-into-view of the
    // closest scrollable ancestor automatically.
    highlighted.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
    // 5) Pick card position OPPOSITE to where the target sits, so the
    // card never covers the highlighted element. Wait one more tick so
    // the scroll has landed before we measure the target's rect.
    await tick();
    cardSide = pickCardSide(highlighted, s.cardPosition);
    updateHole();
  }

  // Compute whether the card should sit at the top or bottom of the
  // popup based on the target's position. The target is laid out in
  // the popup viewport; we want the card on the opposite side so it
  // doesn't overlap. Falls back to the step's hint if measurement
  // fails (e.g. zero-size element).
  function pickCardSide(el: HTMLElement, hint: 'top' | 'bottom' | undefined): 'top' | 'bottom' {
    const popup = el.closest('.popup') as HTMLElement | null;
    const popupRect = popup ? popup.getBoundingClientRect() : { top: 0, height: 640 } as DOMRect;
    const targetRect = el.getBoundingClientRect();
    if (!popupRect.height || !targetRect.height) return hint || 'bottom';
    const targetCenter = targetRect.top + targetRect.height / 2;
    const popupMid = popupRect.top + popupRect.height / 2;
    return targetCenter < popupMid ? 'bottom' : 'top';
  }

  // Compute the hole the SVG mask should leave bright around the
  // highlighted target. We use an SVG mask + Gaussian blur instead of
  // CSS clip-path so the hole's edges feather smoothly into the dim,
  // matching the soft glow of the target's accent ring.
  // Reads the target's computed border-radius so the rounded corners
  // of the hole follow the rounded corners of the target itself.
  function updateHole() {
    viewW = window.innerWidth;
    viewH = window.innerHeight;
    if (!highlighted) {
      hole = null;
      return;
    }
    const r = highlighted.getBoundingClientRect();
    if (!r.width || !r.height) {
      hole = null;
      return;
    }
    const pad = 4;
    const cs = window.getComputedStyle(highlighted);
    // borderRadius can be "9px" or "9px 9px 9px 9px"; parseFloat picks
    // the first value which is what we want for a uniform rounding.
    const radius = parseFloat(cs.borderRadius) || 8;
    hole = {
      x: Math.round(r.left - pad),
      y: Math.round(r.top - pad),
      w: Math.round(r.width + pad * 2),
      h: Math.round(r.height + pad * 2),
      // Add the pad to the radius so the hole's corners stay parallel
      // to the target's, rather than looking sharper at the edge.
      rx: radius + pad,
    };
  }

  function startTour() {
    phase = 'tour';
    stepIdx = 0;
  }
  function next() {
    if (stepIdx < STEPS.length - 1) stepIdx++;
    else finish();
  }
  function back() {
    if (stepIdx > 0) stepIdx--;
  }
  function finish() {
    if (highlighted) { highlighted.classList.remove('onb-target'); highlighted = null; }
    // Restore picker to the user's pre-tour state.
    if (pickerCollapsed !== initialPickerCollapsed) {
      pickerCollapsed = initialPickerCollapsed;
    }
    dispatch('dismiss');
  }

  // Keep the hole lined up while the user (or the browser) scrolls or
  // resizes the popup. Capture-phase scroll listener catches inner
  // scroll containers like .picker-rail too. Both passive — no
  // preventDefault.
  //
  // Throttled with rAF so a fast scroll inside .picker-rail (which
  // can fire 100+ events/sec) only triggers one updateHole per frame.
  // Without this each scroll event ran a full getBoundingClientRect +
  // getComputedStyle + Svelte re-render + SVG mask repaint, which
  // visibly stuttered the scroll on slower machines.
  let rafPending = 0;
  function onScrollOrResize() {
    if (rafPending) return;
    rafPending = requestAnimationFrame(() => {
      rafPending = 0;
      updateHole();
    });
  }
  onMount(() => {
    initialPickerCollapsed = pickerCollapsed;
    viewW = window.innerWidth;
    viewH = window.innerHeight;
    // Skip the prompt phase when the parent explicitly asked to start —
    // currently that's only Settings → Replay tour.
    if (autoStart) phase = 'tour';
    window.addEventListener('scroll', onScrollOrResize, { passive: true, capture: true });
    window.addEventListener('resize', onScrollOrResize, { passive: true });
  });
  onDestroy(() => {
    if (highlighted) highlighted.classList.remove('onb-target');
    if (rafPending) cancelAnimationFrame(rafPending);
    window.removeEventListener('scroll', onScrollOrResize, { capture: true } as any);
    window.removeEventListener('resize', onScrollOrResize);
  });

  // Translation helpers — fall back to the English copy embedded in
  // STEPS so a missing key never shows the raw key text to the user.
  // Cache the t() lookup so we don't call it twice per reactive run.
  $: titleT = t(step.titleKey, {}, lang);
  $: title  = titleT === step.titleKey ? step.titleFallback : titleT;
  $: bodyT  = t(step.bodyKey, {}, lang);
  $: body   = bodyT === step.bodyKey ? step.bodyFallback : bodyT;
</script>

<!-- Fixed-position scrim built as an inline SVG so we can use a
     <mask> with a Gaussian-blurred black rect to feather the edges of
     the hole over the highlighted target. The dim is drawn as a
     translucent <rect> over the full viewport; the mask makes the
     hole transparent. pointer-events:none → dim is visual only and
     never blocks interaction with whatever is underneath, including
     the highlighted target itself. -->
<svg
  class="onb-scrim"
  width={viewW}
  height={viewH}
  viewBox="0 0 {viewW} {viewH}"
  aria-hidden="true"
>
  <defs>
    <!-- The blur extends beyond the rect's bounds, so the filter
         region must be larger than its default 110%. 200% gives plenty
         of headroom for HOLE_BLUR up to ~10px without clipping. -->
    <filter id="onb-soft" x="-50%" y="-50%" width="200%" height="200%">
      <feGaussianBlur stdDeviation={HOLE_BLUR} />
    </filter>
    <mask id="onb-hole-mask">
      <!-- White = visible (dim shows), black = hidden (transparent
           hole). The blurred black rect therefore feathers the dim
           into the hole. -->
      <rect x="0" y="0" width={viewW} height={viewH} fill="white" />
      {#if hole}
        <rect
          x={hole.x}
          y={hole.y}
          width={hole.w}
          height={hole.h}
          rx={hole.rx}
          ry={hole.rx}
          fill="black"
          filter="url(#onb-soft)"
        />
      {/if}
    </mask>
  </defs>
  <rect
    x="0" y="0"
    width={viewW} height={viewH}
    fill="rgba(15, 23, 42, 0.55)"
    mask="url(#onb-hole-mask)"
  />
</svg>

<div
  class="onb-card"
  class:top={phase === 'tour' && cardSide === 'top'}
  class:centered={phase === 'prompt'}
  role="dialog"
  aria-modal="true"
  aria-labelledby="onb-title"
  data-lang={lang}
>
  <!-- Distinct red dismiss in the corner. The accent-red treatment
       reads as "exit" rather than the generic gray X used for "close
       a panel" elsewhere in the popup; users have learned the red-X
       convention from browsers and OSes. -->
  <button
    type="button"
    class="onb-skip onb-skip--danger"
    on:click={finish}
    title={t('onboarding.skip', {}, lang) || 'Skip'}
    aria-label={t('onboarding.skip', {}, lang) || 'Skip'}
  >
    <X size={14} />
  </button>

  {#if phase === 'prompt'}
    <!-- Pre-tour prompt. Centered card asking permission so users who
         already know the extension can skip the walkthrough without
         hunting for an escape hatch. -->
    <div class="onb-prompt-icon" aria-hidden="true">
      <Sparkles size={28} />
    </div>
    <h2 id="onb-title" class="onb-title onb-title--centered">
      {t('onboarding.prompt.title', {}, lang) || 'Welcome!'}
    </h2>
    <p class="onb-body onb-body--centered">
      {t('onboarding.prompt.body', {}, lang) || 'Want a quick 30-second tour of what\'s new?'}
    </p>
    <div class="onb-actions onb-actions--centered">
      <button type="button" class="onb-btn secondary" on:click={finish}>
        {t('onboarding.prompt.skip', {}, lang) || 'No thanks'}
      </button>
      <button type="button" class="onb-btn primary" on:click={startTour}>
        {t('onboarding.prompt.show', {}, lang) || 'Show me'}
      </button>
    </div>
  {:else}
    <div class="onb-step">{stepIdx + 1} / {STEPS.length}</div>
    <h2 id="onb-title" class="onb-title">{title}</h2>
    <p class="onb-body">{body}</p>

    <div class="onb-actions">
      <div class="onb-progress" aria-hidden="true">
        {#each STEPS as _, i}
          <span class="dot" class:done={i < stepIdx} class:current={i === stepIdx}></span>
        {/each}
      </div>
      <!-- Labeled Skip button in the action row. Sits next to the
           navigation buttons so users see an explicit dismiss
           affordance even if they never glance at the corner X. The
           --danger styling matches the red corner X for consistency. -->
      <button type="button" class="onb-btn onb-btn--danger" on:click={finish}>
        {t('onboarding.skip', {}, lang) || 'Skip'}
      </button>
      {#if stepIdx > 0}
        <button type="button" class="onb-btn secondary" on:click={back}>
          {t('onboarding.back', {}, lang) || 'Back'}
        </button>
      {/if}
      <button type="button" class="onb-btn primary" on:click={next}>
        {stepIdx === STEPS.length - 1
          ? (t('onboarding.gotIt', {}, lang) || 'Got it')
          : (t('onboarding.next', {}, lang) || 'Next')}
      </button>
    </div>
  {/if}
</div>

<style>
  /* Full-viewport SVG scrim. The dim itself is drawn as a translucent
     <rect> inside the SVG — see template — masked by a blurred black
     rect to feather the hole around the target. pointer-events:none
     so the dim is visual only and never blocks interaction with the
     highlighted element underneath. */
  .onb-scrim {
    position: fixed;
    inset: 0;
    z-index: 1000;
    pointer-events: none;
    /* Display:block strips the inline-svg baseline gap, which would
       otherwise push the SVG 4-ish px below the viewport top. */
    display: block;
  }

  .onb-card {
    position: fixed;
    z-index: 1001;
    left: 14px; right: 14px;
    bottom: 14px;
    max-width: 392px;
    margin: 0 auto;
    background: var(--color-bg);
    color: var(--color-text);
    border-radius: 12px;
    padding: 16px 16px 12px;
    box-shadow:
      0 20px 40px -12px rgba(0, 0, 0, 0.45),
      0 0 0 1px var(--color-border);
  }
  .onb-card.top {
    top: 50px;
    bottom: auto;
  }
  /* Centered card for the pre-tour prompt — no specific element to
     point at, so we anchor in the middle of the popup with a more
     generous padding for the welcome layout. */
  .onb-card.centered {
    top: 50%;
    bottom: auto;
    transform: translateY(-50%);
    padding: 22px 18px 16px;
    text-align: center;
  }

  .onb-skip {
    position: absolute;
    top: 8px; right: 8px;
    width: 24px; height: 24px;
    background: transparent;
    border: 1px solid transparent;
    color: var(--color-text-muted, #64748b);
    cursor: pointer;
    padding: 0;
    border-radius: 50%;
    display: inline-flex;
    align-items: center;
    justify-content: center;
    transition: background 0.15s, color 0.15s, border-color 0.15s;
  }
  .onb-skip:hover {
    background: var(--color-accent-light);
    color: var(--color-text);
  }
  /* Distinct red dismiss — clearly readable as "exit" rather than the
     ambient gray X used for low-stakes panel-close buttons elsewhere
     in the popup. Subdued by default, saturates on hover. */
  .onb-skip--danger {
    color: #dc2626;
    border-color: rgba(220, 38, 38, 0.3);
  }
  .onb-skip--danger:hover {
    background: rgba(220, 38, 38, 0.12);
    color: #b91c1c;
    border-color: rgba(220, 38, 38, 0.55);
  }

  /* Sparkle icon at the top of the prompt card — gives the welcome
     layout a focal point without needing decorative copy. */
  .onb-prompt-icon {
    display: inline-flex;
    width: 48px; height: 48px;
    margin: 0 auto 10px;
    border-radius: 50%;
    align-items: center;
    justify-content: center;
    color: var(--color-accent);
    background: var(--color-accent-light, rgba(37, 99, 235, 0.1));
  }
  .onb-title--centered {
    padding-right: 0;  /* no skip X to dodge — it's outside the centered text column */
    margin-top: 0;
  }
  .onb-body--centered {
    margin-bottom: 16px;
  }
  .onb-actions--centered {
    justify-content: center;
  }

  .onb-step {
    font-size: 10px;
    font-weight: 600;
    color: var(--color-accent);
    letter-spacing: 0.04em;
    text-transform: uppercase;
    margin-bottom: 6px;
  }

  .onb-title {
    font-size: 14px;
    font-weight: 600;
    margin: 0 0 6px;
    line-height: 1.3;
    padding-right: 24px;  /* avoid overlap with the skip × */
  }

  .onb-body {
    font-size: 12px;
    line-height: 1.5;
    color: var(--color-text-muted, #64748b);
    margin: 0 0 12px;
  }

  .onb-actions {
    display: flex;
    gap: 8px;
    justify-content: flex-end;
    align-items: center;
  }
  .onb-progress {
    margin-right: auto;
    display: flex;
    gap: 4px;
  }
  .onb-progress .dot {
    width: 6px; height: 6px;
    border-radius: 50%;
    background: var(--color-border);
    transition: background 0.15s, transform 0.15s;
  }
  .onb-progress .dot.done {
    background: var(--color-accent);
    opacity: 0.55;
  }
  .onb-progress .dot.current {
    background: var(--color-accent);
    transform: scale(1.4);
  }

  .onb-btn {
    font: inherit;
    font-size: 12px;
    font-weight: 600;
    padding: 6px 12px;
    border-radius: 6px;
    cursor: pointer;
    border: 1px solid transparent;
    transition: background 0.15s, border-color 0.15s, filter 0.15s;
  }
  .onb-btn.primary {
    background: var(--color-accent);
    color: #fff;
    border-color: var(--color-accent);
  }
  .onb-btn.primary:hover { filter: brightness(1.05); }
  .onb-btn.secondary {
    background: transparent;
    color: var(--color-text);
    border-color: var(--color-border);
  }
  .onb-btn.secondary:hover { background: var(--color-accent-light); }
  /* Red-tinted Skip in the action row. Same family as the corner X so
     users learn one visual language for "exit the tour". Sits visually
     subordinate to Back/Next so dismissal is available but not the
     loudest affordance. */
  .onb-btn--danger {
    background: transparent;
    color: #dc2626;
    border-color: rgba(220, 38, 38, 0.35);
    margin-right: 4px;
  }
  .onb-btn--danger:hover {
    background: rgba(220, 38, 38, 0.1);
    color: #b91c1c;
    border-color: rgba(220, 38, 38, 0.6);
  }

  /* The target highlight: just a pulsing accent ring. The dim is now
     the separate .onb-scrim with a clip-path hole — see above. The
     ring is purely decorative; the scrim's transparent hole is what
     visually isolates the target.
     :global() because the class is added/removed imperatively via
     classList from JS — Svelte's component-scoped CSS would otherwise
     prune the selector. */
  :global(.onb-target) {
    border-radius: 9px;
    animation: onb-target-pulse 1.6s ease-in-out infinite;
  }
  @keyframes onb-target-pulse {
    0%, 100% { box-shadow: 0 0 0 3px rgba(37, 99, 235, 0.7); }
    50%      { box-shadow: 0 0 0 6px rgba(37, 99, 235, 0.35); }
  }
</style>
