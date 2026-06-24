// Console-only export diagnostics.
//
// Collects, per export, the data that explains where an export spent its time
// and how big it was: per-stage durations (the four segments shown on the
// export button — messages, images, people, file), per-chat message counts
// and wall time, and per-format output sizes. At completion it logs one
// structured block to the service-worker console.
//
// Nothing here is persisted or shown in the UI. It exists purely so a slow or
// oversized export can be diagnosed from the SW console without instrumenting
// the build by hand. Every hook is passive: if no stats object is registered
// for a tab, the callers no-op.

// The four stages the export button surfaces as progress segments. Raw
// progress phases collapse onto these so the log matches what the user sees.
export type ButtonStage = 'messages' | 'images' | 'people' | 'file';

// Raw progress phase -> button stage. Phases not listed here (e.g. 'hud',
// 'scrape:start', 'complete') carry no stage time and are ignored by
// markPhase. 'scroll'/'extract' are the DOM-fallback equivalents of the
// API 'api-fetch' phase, so they fold into 'messages'.
const PHASE_TO_STAGE: Readonly<Record<string, ButtonStage>> = {
  scroll: 'messages',
  extract: 'messages',
  'api-fetch': 'messages',
  images: 'images',
  avatars: 'people',
  build: 'file',
};

// The three scrape stages that are bounded by a single chat's
// scrape:start -> scrape:complete window, so they can be attributed per chat.
// 'file' (build) happens AFTER the scrape, so it stays in the global total only.
type ScrapeStage = 'messages' | 'images' | 'people';

type FormatBytes = { format: string; bytes: number };
type ChatStat = {
  title?: string;
  convId?: string;
  messages: number;
  scrapeMs: number;
  // Per-chat scrape-stage split, present once a chat has been opened with
  // beginChat. Only logged when verbose mode is on (see log()).
  stageMs?: Record<ScrapeStage, number>;
};

export class ExportStats {
  readonly kind: 'single' | 'bundle';
  private readonly startedAt: number;
  private readonly stageMs: Record<ButtonStage, number> = {
    messages: 0,
    images: 0,
    people: 0,
    file: 0,
  };
  private curStage: ButtonStage | null = null;
  private curStageStart = 0;
  // Per-chat scrape-stage accumulator. Non-null between beginChat
  // (scrape:start) and addChat (scrape:complete); markPhase banks the
  // fetch stages into it alongside the global total.
  private curChatStages: Record<ScrapeStage, number> | null = null;
  private readonly formats: FormatBytes[] = [];
  private readonly chats: ChatStat[] = [];

  constructor(kind: 'single' | 'bundle', now: number) {
    this.kind = kind;
    this.startedAt = now;
  }

  // Fold a raw progress phase into the running stage timeline. When the
  // active stage changes, the elapsed time since the last change is banked
  // into the previous stage. Phases progress roughly linearly
  // (messages -> images -> people -> file), so the per-stage totals reflect
  // how long each export-button segment was actually lit. For a bundle the
  // sequence repeats per chat and the totals aggregate across all chats.
  markPhase(phase: string | undefined, now: number): void {
    if (!phase) return;
    const stage = PHASE_TO_STAGE[phase];
    if (!stage || stage === this.curStage) return;
    if (this.curStage) {
      const d = Math.max(0, now - this.curStageStart);
      this.stageMs[this.curStage] += d;
      // 'file' (build) runs after a chat's scrape window, when curChatStages
      // is already closed, so it never reaches the per-chat split.
      if (this.curChatStages && this.curStage !== 'file') {
        this.curChatStages[this.curStage] += d;
      }
    }
    this.curStage = stage;
    this.curStageStart = now;
  }

  // Open a fresh per-chat scrape-stage accumulator at a chat's scrape:start.
  // First banks any still-open stage (the PREVIOUS chat's trailing build
  // phase) into the global total only, since that chat is already finalized.
  beginChat(now: number): void {
    if (this.curStage) {
      this.stageMs[this.curStage] += Math.max(0, now - this.curStageStart);
      this.curStage = null;
    }
    this.curChatStages = { messages: 0, images: 0, people: 0 };
  }

  // Finalize a chat at scrape:complete. Flushes the trailing fetch stage into
  // both the global total and this chat's split, then closes the stage so the
  // gap until the build phase isn't misattributed to a fetch stage.
  addChat(chat: Omit<ChatStat, 'stageMs'>, now: number): void {
    if (this.curStage && this.curStage !== 'file') {
      const d = Math.max(0, now - this.curStageStart);
      this.stageMs[this.curStage] += d;
      if (this.curChatStages) this.curChatStages[this.curStage] += d;
      this.curStage = null;
    }
    const stageMs = this.curChatStages ? { ...this.curChatStages } : undefined;
    this.chats.push({ ...chat, stageMs });
    this.curChatStages = null;
  }

  // Per-format output size. Keys are the document formats (json/csv/html/txt/
  // pdf) holding that file's serialized, pre-compression byte length, plus an
  // 'assets' key aggregating any separately-bundled image/avatar files (so a
  // files-mode HTML zip isn't undersold by reporting only index.html). May be
  // called more than once per key in a multi-chat bundle (once per chat); the
  // log sums by key below.
  addFormat(format: string, bytes: number): void {
    if (!format || !Number.isFinite(bytes) || bytes < 0) return;
    this.formats.push({ format, bytes });
  }

  // Bank the final open stage and emit the structured log line. `outcome`
  // distinguishes a clean finish from a cancel/empty/error so a short export
  // log isn't mistaken for a complete one.
  // `verbose` gates the per-chat detail. When false (the default, every
  // shipped build) only the ID-free aggregates are logged; when true (a
  // deliberate, visible opt-in on the Diagnostics page) the block also carries
  // per-chat titles, convIds and the scrape-stage split. The aggregates are
  // safe to share; the per-chat block is not, which is why it is opt-in.
  log(
    now: number,
    outcome: string,
    extra: { messages?: number; outputBytes?: number; filename?: string; verbose?: boolean; buildStamp?: string } = {},
  ): void {
    if (this.curStage) {
      this.stageMs[this.curStage] += Math.max(0, now - this.curStageStart);
      this.curStage = null;
    }
    const round = (ms: number) => Math.round(ms);
    const stages: Record<string, number> = {};
    for (const k of Object.keys(this.stageMs) as ButtonStage[]) {
      if (this.stageMs[k] > 0) stages[k] = round(this.stageMs[k]);
    }
    // Sum byte sizes by format (a bundle reports each format once per chat).
    const formatBytes: Record<string, number> = {};
    for (const f of this.formats) formatBytes[f.format] = (formatBytes[f.format] || 0) + f.bytes;

    const totalMessages = extra.messages ?? this.chats.reduce((s, c) => s + c.messages, 0);
    const block = {
      kind: this.kind,
      outcome,
      ...(extra.buildStamp ? { build: extra.buildStamp } : {}),
      totalMs: round(now - this.startedAt),
      totalMessages,
      chats: this.chats.length,
      stageMs: stages,
      formatBytes,
      ...(extra.outputBytes != null ? { outputBytes: extra.outputBytes } : {}),
      ...(extra.filename ? { filename: extra.filename } : {}),
      // Per-chat detail only in verbose mode, and last so the aggregate summary
      // reads first in the console.
      ...(extra.verbose ? { perChat: this.chats } : {}),
    };
    // Single grouped line keyed so it's greppable in the SW console.
    console.log('[export-stats]', JSON.stringify(block));
  }
}

// Per-tab registry. One export runs per tab at a time (the SW rejects a
// second concurrent export for the same tab), so tabId is a safe key.
const statsByTab = new Map<number, ExportStats>();

export function beginExportStats(tabId: number, kind: 'single' | 'bundle', now: number): ExportStats {
  const stats = new ExportStats(kind, now);
  statsByTab.set(tabId, stats);
  return stats;
}

export function getExportStats(tabId: number | undefined): ExportStats | undefined {
  return typeof tabId === 'number' ? statsByTab.get(tabId) : undefined;
}

export function endExportStats(tabId: number): void {
  statsByTab.delete(tabId);
}
