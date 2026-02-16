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
  makeDayDivider: (dayKey: number, ts: number) => ContentAggregated<M>;
};

const sleep = (ms: number) => new Promise(resolve => setTimeout(resolve, ms));

const findPaneItemByMessageId = (id: string | null | undefined): Element | null => {
  if (!id) return null;
  const msgNode = document.querySelector(`[data-mid="${CSS.escape(id)}"]`);
  return msgNode?.closest('[data-tid="chat-pane-item"]') || null;
};

export async function autoScrollAggregate<M extends ExportMessage>(
  deps: ScrollDeps<M>,
  { startAtISO, endAtISO, includeSystem, includeReactions, includeReplies = true }: ScrapeOptions & { includeReplies?: boolean },
  currentRunStartedAt: number | null,
): Promise<M[]> {
  const { hud, runtime, extractOne, hydrateSparseMessages, getScroller } = deps;
  const scroller = getScroller();
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

  scroller.scrollTop = scroller.scrollHeight;
  await new Promise(r => requestAnimationFrame(r));
  await sleep(300);
  await collectCurrentVisible(agg, { includeSystem, includeReactions, includeReplies }, orderCtx, extractOne);

  let prevHeight = -1;
  let lastCount = -1;
  let passes = 0;
  let stagnantPasses = 0;
  let lastOldestId = null;
  let lastAggSize = 0;
  const baseDwellMs = 700;

  const headerSentinel = document.querySelector('[data-tid="message-pane-header"]');
  let topReached = false;
  const observer = headerSentinel
    ? new IntersectionObserver(entries => {
        const entry = entries[0];
        if (entry?.isIntersecting) topReached = true;
      }, { root: scroller, threshold: 0.01 })
    : null;
  if (observer && headerSentinel) observer.observe(headerSentinel);

  const startLimit = typeof startAtISO === 'string' ? parseTimeStamp(startAtISO) : null;
  const endLimit = typeof endAtISO === 'string' ? parseTimeStamp(endAtISO) : null;

  try {
    while (true) {
      passes++;
      // Adaptive dwell: Teams gets slower with deep history
      const dwellMs = baseDwellMs + Math.min(Math.floor(agg.size / 500) * 200, 2000);
      scroller.scrollTop = 0;
      await new Promise(r => requestAnimationFrame(r));
      await sleep(dwellMs);

      await collectCurrentVisible(agg, { includeSystem, includeReactions, includeReplies }, orderCtx, extractOne);

      const nodes = $$('[data-tid="chat-pane-item"]');
      if (!nodes.length) break;
      const newCount = nodes.length;
      const newHeight = (scroller as HTMLElement).scrollHeight;
      const oldestNode = nodes[0];
      const oldestTimeAttr = $('time[datetime]', oldestNode)?.getAttribute('datetime') || null;
      const oldestTs = parseTimeStamp(oldestTimeAttr);
      const oldestId = $('[data-tid="chat-pane-message"]', oldestNode)?.getAttribute('data-mid') || oldestNode?.id || null;

      const hiddenButtons = Array.from(document.querySelectorAll<HTMLButtonElement>('[data-tid="show-hidden-chat-history-btn"]')).filter(
        btn => btn && !btn.disabled && btn.offsetParent !== null,
      );
      if (hiddenButtons.length) {
        for (const btn of hiddenButtons) {
          try { btn.click(); } catch (err) { console.warn('[Teams Exporter] failed to click hidden-history button', err); }
          await sleep(400);
        }
        scroller.scrollTop = 0;
        await new Promise(r => requestAnimationFrame(r));
        scroller.scrollTop = 0;
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
          },
        });
        if (msgPromise && msgPromise.catch) msgPromise.catch(() => {});
      } catch {}
      hud(`scroll pass ${passes} • seen ${filteredSeen}`);

      if (startLimit != null && oldestTs != null && oldestTs <= startLimit) {
        console.log('[Teams Exporter] scroll stop: startAt date reached', { oldestTimeAttr, startAtISO });
        break;
      }

      const heightUnchanged = newHeight === prevHeight;
      const countUnchanged = newCount === lastCount;
      const oldestUnchanged = oldestId && lastOldestId === oldestId;
      // Teams virtualizes: DOM count/height may stay constant while content swaps.
      // Track whether we're still collecting new messages.
      const aggGrew = agg.size > lastAggSize;
      lastAggSize = agg.size;

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

      if (topReached && stagnantPasses >= 3) {
        console.log('[Teams Exporter] scroll stop: header sentinel reached + stagnant', { passes, stagnantPasses, aggregated: agg.size, oldestTimeAttr });
        break;
      }
      if (!topReached && stagnantPasses >= 25) {
        console.log('[Teams Exporter] scroll stop: stagnation threshold', { passes, stagnantPasses, aggregated: agg.size, oldestTimeAttr, dwellMs });
        break;
      }
    }
  } finally {
    if (observer && headerSentinel) observer.disconnect();
  }

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

  for (const item of noDate) finalMessages.push(item.message);

  return finalMessages;
}

async function collectCurrentVisible<M extends ExportMessage>(
  agg: Map<string, ContentAggregated<M>>,
  opts: ScrapeOptions,
  orderCtx: OrderContext,
  extractOne: ScrollDeps<M>['extractOne'],
) {
  const nodes = $$('[data-tid="chat-pane-item"]');
  const lastAuthorRef = { value: orderCtx.lastAuthor || '' };
  for (let i = 0; i < nodes.length; i++) {
    const item = nodes[i];
    const idCandidate =
      $('[data-tid="chat-pane-message"]', item)?.getAttribute('data-mid') ||
      $('[data-tid="control-message-renderer"]', item)?.getAttribute('data-mid') ||
      $('.fui-Divider__wrapper', item)?.id ||
      item.id ||
      `node-${i}`;
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
