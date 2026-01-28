import { $$, $ } from '../utils/dom';
import { parseTimeStamp, startOfLocalDay } from '../utils/time';
import type { AggregatedItem, ExportMessage, OrderContext, ScrapeOptions } from '../types/shared';

type ContentAggregated<M extends ExportMessage> = AggregatedItem & { message?: M };

export type ScrollDeps<M extends ExportMessage> = {
  hud: (text: string, opts?: { includeElapsed?: boolean }) => void;
  runtime: typeof chrome.runtime;
  extractOne: (item: Element, opts: ScrapeOptions, lastAuthorRef: { value: string }, orderCtx: OrderContext & { seq?: number }) => Promise<ContentAggregated<M> | null>;
  hydrateSparseMessages: (agg: Map<string, ContentAggregated<M>>, opts: ScrapeOptions) => Promise<void>;
  getScroller: () => Element | null;
  getItems?: () => Element[];
  getItemId?: (item: Element, index: number) => string;
  isLoading?: () => boolean;
  makeDayDivider: (dayKey: number, ts: number) => ContentAggregated<M>;
  tuning?: {
    dwellMs?: number;
    maxStagnant?: number;
    maxStagnantAtTop?: number;
    loadingStallPasses?: number;
    loadingExtraDelayMs?: number;
  };
};

const sleep = (ms: number) => new Promise(resolve => setTimeout(resolve, ms));
const isScrollable = (el: Element | null): el is HTMLElement => {
  if (!el) return false;
  const target = el as HTMLElement;
  return target.scrollHeight > target.clientHeight + 1;
};


const defaultGetItems = () => $$('[data-tid="chat-pane-item"]');

const defaultGetItemId = (item: Element, index: number) => {
  const midSelector = '[data-tid="chat-pane-message"], [data-tid="channel-pane-message"]';
  return (
    $(midSelector, item)?.getAttribute('data-mid') ||
    $('[data-tid="control-message-renderer"]', item)?.getAttribute('data-mid') ||
    $('[data-mid]', item)?.getAttribute('data-mid') ||
    $('.fui-Divider__wrapper', item)?.id ||
    item.id ||
    `node-${index}`
  );
};

export async function autoScrollAggregate<M extends ExportMessage>(
  deps: ScrollDeps<M>,
  { startAtISO, endAtISO, includeSystem, includeReactions, includeReplies = true }: ScrapeOptions & { includeReplies?: boolean },
  currentRunStartedAt: number | null,
): Promise<M[]> {
  const { hud, runtime, extractOne, hydrateSparseMessages, getScroller } = deps;
  const getItems = deps.getItems || defaultGetItems;
  const getItemId = deps.getItemId || defaultGetItemId;
  const isLoading = deps.isLoading || (() => false);
  let scroller = getScroller();
  if (!scroller) throw new Error('Scroller not found');

  const agg = new Map<string, ContentAggregated<M>>();
  const orderCtx: OrderContext = {
    lastTimeMs: null,
    yearHint: null,
    seqBase: Date.now(),
    seq: 0,
    lastAuthor: '',
    lastId: null,
    systemCursor: -9e15,
  };

  let prevHeight = -1;
  let lastCount = -1;
  let passes = 0;
  let stagnantPasses = 0;
  let lastOldestId = null;
  let lastAggSize = 0;
  let loadingNoProgress = 0;
  const baseDwellMs = deps.tuning?.dwellMs ?? 700;
  const hasLoadingSignal = typeof deps.isLoading === 'function';
  const maxStagnant = deps.tuning?.maxStagnant ?? (hasLoadingSignal ? 20 : 12);
  const maxStagnantAtTop = deps.tuning?.maxStagnantAtTop ?? (hasLoadingSignal ? 6 : 3);
  const loadingStallPasses = deps.tuning?.loadingStallPasses ?? 6;
  const loadingExtraDelayMs = deps.tuning?.loadingExtraDelayMs ?? 350;

  const headerSentinel = document.querySelector('[data-tid="message-pane-header"]');
  let topReached = false;
  const createObserver = (root: Element | null) => {
    if (!headerSentinel || !root) return null;
    const obs = new IntersectionObserver(entries => {
      const entry = entries[0];
      if (entry?.isIntersecting) topReached = true;
    }, { root, threshold: 0.01 });
    obs.observe(headerSentinel);
    return obs;
  };
  let observer = createObserver(scroller);
  const ensureScroller = () => {
    const candidate = getScroller() || document.scrollingElement;
    if (!candidate) return scroller;
    const currentConnected = (scroller as Element).isConnected;
    const shouldSwap =
      candidate !== scroller &&
      (isScrollable(candidate) || !isScrollable(scroller) || !currentConnected);
    if (shouldSwap) {
      observer?.disconnect();
      scroller = candidate;
      observer = createObserver(scroller);
    }
    return scroller;
  };

  scroller = ensureScroller();
  scroller.scrollTop = scroller.scrollHeight;
  await new Promise(r => requestAnimationFrame(r));
  await sleep(300);
  await collectCurrentVisible(agg, { includeSystem, includeReactions, includeReplies }, orderCtx, extractOne, getItems, getItemId);

  const startLimit = typeof startAtISO === 'string' ? parseTimeStamp(startAtISO) : null;
  const endLimit = typeof endAtISO === 'string' ? parseTimeStamp(endAtISO) : null;

const dispatchWheel = (el: HTMLElement, deltaY: number) => {
  try {
    el.dispatchEvent(new WheelEvent('wheel', { deltaY, deltaMode: 0, bubbles: true, cancelable: true }));
  } catch {}
};

const slowScrollUp = (el: HTMLElement, stepPx = 0) => {
  const step = stepPx > 0 ? stepPx : Math.max(120, Math.floor(el.clientHeight * 0.35));
  const nextTop = Math.max(0, el.scrollTop - step);
  el.scrollTop = nextTop;
  el.dispatchEvent(new Event('scroll', { bubbles: true }));
  dispatchWheel(el, -step);
  return nextTop;
};

const forceScrollUp = (el: HTMLElement, multiplier = 2) => {
  const jump = el.clientHeight * multiplier;
  el.scrollTop = Math.max(0, el.scrollTop - jump);
  el.dispatchEvent(new Event('scroll', { bubbles: true }));
  dispatchWheel(el, -jump);
  };

  try {
    while (true) {
      passes++;
      scroller = ensureScroller();
      const scrollerEl = scroller as HTMLElement;
      const prevTop = scrollerEl.scrollTop;
      slowScrollUp(scrollerEl);
      // Adaptive dwell: Teams gets slower with deep history
      const dwellMs = baseDwellMs + Math.min(Math.floor(agg.size / 500) * 200, 2000);
      await new Promise(r => requestAnimationFrame(r));
      // Virtualized lists can snap; nudge upward if no movement.
      if (scrollerEl.scrollTop >= prevTop && scrollerEl.scrollTop > 2) {
        slowScrollUp(scrollerEl, Math.max(60, Math.floor(scrollerEl.clientHeight * 0.2)));
        await new Promise(r => requestAnimationFrame(r));
      }
      await sleep(dwellMs);

      await collectCurrentVisible(agg, { includeSystem, includeReactions, includeReplies }, orderCtx, extractOne, getItems, getItemId);

      const nodes = getItems();
      if (!nodes.length) break;
      const newCount = nodes.length;
      const newHeight = (scroller as HTMLElement).scrollHeight;
      const oldestNode = nodes[0];
      const oldestTimeAttr = $('time[datetime]', oldestNode)?.getAttribute('datetime') || null;
      const oldestTs = parseTimeStamp(oldestTimeAttr);
      const oldestId = getItemId(oldestNode, 0) || null;
      const loadingRaw = isLoading();

      const hiddenButtons = Array.from(document.querySelectorAll<HTMLButtonElement>('[data-tid="show-hidden-chat-history-btn"]')).filter(
        btn => btn && !btn.disabled && btn.offsetParent !== null,
      );
      if (hiddenButtons.length) {
        for (const btn of hiddenButtons) {
          try { btn.click(); } catch (err) { console.warn('[Teams Exporter] failed to click hidden-history button', err); }
          await sleep(400);
        }
        slowScrollUp(scrollerEl);
        await new Promise(r => requestAnimationFrame(r));
        slowScrollUp(scrollerEl);
        await sleep(300);
        stagnantPasses = 0;
        prevHeight = -1;
        lastCount = -1;
        lastOldestId = null;
        await sleep(600);
        continue;
      }

      const elapsedMs = currentRunStartedAt ? Date.now() - currentRunStartedAt : null;
      const seen = agg.size;
      let filteredSeen = 0;
      for (const entry of agg.values()) {
        const candidate = entry?.tsMs ?? (entry?.message?.timestamp ? parseTimeStamp(entry.message.timestamp) : null);
        if (candidate == null) {
          filteredSeen++;
          continue;
        }
        if (startLimit != null && candidate < startLimit) continue;
        if (endLimit != null && candidate >= endLimit) continue;
        filteredSeen++;
      }

      const heightUnchanged = newHeight === prevHeight;
      const countUnchanged = newCount === lastCount;
      const oldestUnchanged = oldestId && lastOldestId === oldestId;
      const aggUnchanged = seen === lastAggSize;
      const progressObserved = !heightUnchanged || !countUnchanged || !oldestUnchanged || !aggUnchanged;

      let loading = loadingRaw;
      if (loadingRaw) {
        if (scrollerEl.scrollTop > 2) {
          // Keep nudging upward while the loader is visible.
          forceScrollUp(scrollerEl, 3);
          await new Promise(r => requestAnimationFrame(r));
        }
        loadingNoProgress = progressObserved ? 0 : (loadingNoProgress + 1);
        if (loadingNoProgress === loadingStallPasses) {
          // Loader looks stuck; try a stronger upward nudge before giving up.
          scrollerEl.scrollTop = 0;
          await new Promise(r => requestAnimationFrame(r));
          if (scrollerEl.scrollTop > 2) {
            forceScrollUp(scrollerEl, 3);
            await new Promise(r => requestAnimationFrame(r));
          }
          await sleep(200);
        } else if (loadingNoProgress > loadingStallPasses) {
          // Loader is still stuck; treat as not loading so we can terminate.
          loading = false;
        }
      } else {
        loadingNoProgress = 0;
      }

      try {
        const msgPromise = runtime.sendMessage({
          type: 'SCRAPE_PROGRESS',
          payload: {
            phase: 'scroll',
            passes,
            newHeight,
            messagesVisible: newCount,
            aggregated: seen,
            seen: filteredSeen,
            filteredSeen,
            oldestTime: oldestTimeAttr,
            oldestId,
            elapsedMs,
            loading,
          },
        });
        if (msgPromise && msgPromise.catch) msgPromise.catch(() => {});
      } catch {}
      hud(`scroll pass ${passes} • seen ${filteredSeen}`);

      if (startLimit != null && oldestTs != null && oldestTs <= startLimit) {
        console.log('[Teams Exporter] scroll stop: startAt date reached', { oldestTimeAttr, startAtISO });
        break;
      }

      if (loading) {
        // Virtual list is still fetching; avoid counting stagnation.
        stagnantPasses = Math.max(0, stagnantPasses - 1);
        if (oldestId && lastOldestId !== oldestId) {
          lastOldestId = oldestId;
        }
        prevHeight = newHeight;
        lastCount = newCount;
        lastAggSize = seen;
        await sleep(dwellMs + loadingExtraDelayMs);
        continue;
      }
      // Teams virtualizes: DOM count/height may stay constant while content swaps.
      // Track whether we're still collecting new messages.
      const aggGrew = seen > lastAggSize;

      if (aggGrew) {
        // Still finding new messages — not stagnant regardless of DOM metrics
        stagnantPasses = 0;
      } else if (heightUnchanged && countUnchanged) {
        stagnantPasses++;
      } else if (oldestUnchanged) {
        stagnantPasses++;
      } else {
        stagnantPasses = 0;
      }

      if (oldestId && lastOldestId !== oldestId) {
        lastOldestId = oldestId;
      }

      prevHeight = newHeight;
      lastCount = newCount;
      lastAggSize = seen;

      if (hasLoadingSignal && stagnantPasses > 6) {
        // Nudge upward to retrigger loading if the list is stubborn.
        slowScrollUp(scrollerEl, Math.max(80, Math.floor(scrollerEl.clientHeight * 0.25)));
      }

      if (scrollerEl.scrollTop <= 2) topReached = true;
      if (topReached && stagnantPasses >= maxStagnantAtTop) {
        console.log('[Teams Exporter] scroll stop: header sentinel reached + stagnant', { passes, stagnantPasses, aggregated: agg.size, oldestTimeAttr });
        break;
      }
      if (!topReached && stagnantPasses >= maxStagnant) {
        console.log('[Teams Exporter] scroll stop: stagnation threshold', { passes, stagnantPasses, aggregated: agg.size, oldestTimeAttr, dwellMs });
        break;
      }
    }
  } finally {
    if (observer && headerSentinel) observer.disconnect();
  }

  // Final bottom pass to capture any newest messages that appeared during upward scrolling.
  scroller = ensureScroller();
  scroller.scrollTop = scroller.scrollHeight;
  await new Promise(r => requestAnimationFrame(r));
  await sleep(deps.tuning?.dwellMs ?? 700);
  await collectCurrentVisible(agg, { includeSystem, includeReactions, includeReplies }, orderCtx, extractOne, getItems, getItemId);

  await hydrateSparseMessages(agg, { includeSystem, includeReactions });

  const entries = Array.from(agg.values());
  entries.sort((a, b) => a.orderKey - b.orderKey);

  let nextMessageTs: number | null = null;
  for (let i = entries.length - 1; i >= 0; i--) {
    const entry = entries[i];
    if (entry.kind === 'message') {
      if (entry.tsMs != null) nextMessageTs = entry.tsMs;
      continue;
    }
    if (nextMessageTs != null) {
      if (entry.tsMs == null || entry.tsMs >= nextMessageTs) {
        entry.anchorTs = nextMessageTs;
        entry.tsMs = (entry.tsMs == null ? nextMessageTs : entry.tsMs) - 1;
        if (entry.tsMs != null) entry.orderKey = entry.tsMs - 0.1;
      }
    }
  }

  let filtered = entries.filter(entry => entry.kind !== 'day-divider');
  filtered.sort((a, b) => {
    const aTs = (a.tsMs ?? a.anchorTs ?? a.orderKey ?? 0);
    const bTs = (b.tsMs ?? b.anchorTs ?? b.orderKey ?? 0);
    if (aTs !== bTs) return aTs - bTs;
    return a.orderKey - b.orderKey;
  });

  filtered = filtered.filter(entry => {
    const ts = entry.anchorTs ?? entry.tsMs ?? (entry.message?.timestamp ? parseTimeStamp(entry.message.timestamp) : null);
    if (ts == null) return true;
    if (startLimit != null && ts < startLimit) return false;
    if (endLimit != null && ts >= endLimit) return false;
    return true;
  });

  const buckets = new Map<number, { ts: number; message: M }[]>();
  const noDate: { ts: number; message: M }[] = [];

  for (const entry of filtered) {
    const msg = entry.message;
    if (!msg) continue;
    // Filter out system messages if includeSystem is false
    if (msg.system && !includeSystem) {
      continue;
    }
    // Also filter out empty/generic system messages even if includeSystem is true
    if (msg.system && (!msg.text || msg.text.trim().toLowerCase() === 'system')) {
      continue;
    }
    const ts = entry.anchorTs ?? entry.tsMs ?? (msg.timestamp ? parseTimeStamp(msg.timestamp) : null);
    if (ts == null) {
      noDate.push({ ts: Number.MIN_SAFE_INTEGER, message: msg });
      continue;
    }
    const dayKey = startOfLocalDay(ts);
    if (!buckets.has(dayKey)) buckets.set(dayKey, []);
    const list = buckets.get(dayKey);
    if (list) list.push({ ts, message: msg });
  }

  const finalMessages: M[] = [];
  const sortedDayKeys = Array.from(buckets.keys()).sort((a, b) => a - b);
  for (const dayKey of sortedDayKeys) {
    const items = buckets.get(dayKey);
    if (!items || !items.length) continue;
    const representativeTs = items[0].ts;
    // Only add day dividers if includeSystem is true
    if (includeSystem) {
      const divider = deps.makeDayDivider(dayKey, representativeTs);
      if (divider?.message) finalMessages.push(divider.message);
    }
    items.sort((a, b) => a.ts - b.ts);
    for (const item of items) finalMessages.push(item.message);
  }

  // Deduplicate obvious clones (same author+timestamp+text in same thread)
  const deduped: M[] = [];
  const seenKeys = new Set<string>();

  for (const msg of finalMessages) {
    if (!msg) continue;

    const text = (msg.text || '').trim();
    const author = msg.author || '';
    const ts = msg.timestamp || '';
    // if you have a thread/threadRoot id in your type, include it here
    const threadRoot = (msg as any).threadRootId || (msg as any).threadId || '';

    const key = `${threadRoot}|${author}|${ts}|${text}`;

    if (seenKeys.has(key)) {
      // Skip cloned message (like your 573/574 copies)
      continue;
    }
    seenKeys.add(key);
    deduped.push(msg);
  }

  return deduped;

}

async function collectCurrentVisible<M extends ExportMessage>(
  agg: Map<string, ContentAggregated<M>>,
  opts: ScrapeOptions,
  orderCtx: OrderContext,
  extractOne: ScrollDeps<M>['extractOne'],
  getItems: () => Element[],
  getItemId: (item: Element, index: number) => string,
) {
  const nodes = getItems();
  const lastAuthorRef = { value: orderCtx.lastAuthor || '' };
  for (let i = 0; i < nodes.length; i++) {
    const item = nodes[i];
    const idCandidate = getItemId(item, i);
    if (agg.has(idCandidate)) continue;

    const extracted = await extractOne(item, opts, lastAuthorRef, orderCtx);
    if (!extracted) continue;
    if (extracted.kind === 'day-divider') {
      if (typeof extracted.tsMs === 'number' && Number.isFinite(extracted.tsMs)) {
        orderCtx.lastTimeMs = extracted.tsMs;
        orderCtx.yearHint = new Date(extracted.tsMs).getFullYear();
      }
      continue;
    }
    const { message, orderKey, tsMs, kind } = extracted;
    if (!message) continue;

    agg.set(message.id || `${orderKey}`, { message, orderKey, tsMs, kind });
    if (!message.system && message.timestamp) {
      const tms = Date.parse(message.timestamp);
      if (!Number.isNaN(tms)) {
        orderCtx.lastTimeMs = tms;
        orderCtx.yearHint = new Date(tms).getFullYear();
      }
    }
    if (!message.system && message.author) {
      orderCtx.lastAuthor = message.author;
    }
  }
}
