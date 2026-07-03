// Settled-completion wait for the export artifact download (package mode,
// i.e. the Files toggle). The legacy waitForDownloadComplete in background.ts
// waits at most 30s and then PROCEEDS on timeout — acceptable when nothing
// depends on the artifact, but the attachments phase must never start against
// an export that isn't actually on disk. This wait has no proceed-on-timeout:
// it resolves only on a real terminal signal, an abort, or a genuine stall.
//
// Mechanism: a downloads.onChanged listener catches the state transition
// (complete/interrupted); a slow poll of downloads.search backs it up. The
// poll doubles as the MV3 keep-alive — each extension API call resets the
// service worker's 30s idle timer (Chrome 110+), the same pattern
// awaitTerminal uses in attachment-download.ts.
//
// Stall semantics: only a download that is actively failing to make progress
// counts as stalled. Two states are exempt from the stall clock because the
// download is waiting on the USER, not the disk:
//   - empty item.filename: the target is undecided — Chrome's global "Ask
//     where to save each file" preference forces a chooser even when the
//     extension passed saveAs:false, and while it is open the item sits
//     in_progress with no target;
//   - item.paused: paused from the downloads shelf; it may be resumed.
//
// Dep-injected (mirrors attachment-download.ts) so it is unit-testable
// without a browser.

type DownloadsSearchApi = Pick<typeof chrome.downloads, 'search'>;

export interface DownloadChangedEvent {
  addListener(cb: (delta: chrome.downloads.DownloadDelta) => void): void;
  removeListener(cb: (delta: chrome.downloads.DownloadDelta) => void): void;
}

export type SettledOutcome = 'complete' | 'interrupted' | 'stalled' | 'missing' | 'aborted';

export function waitForDownloadSettled(
  deps: { downloads: DownloadsSearchApi; onChanged: DownloadChangedEvent },
  id: number,
  opts: { pollMs?: number; stallMs?: number; signal?: AbortSignal } = {},
): Promise<SettledOutcome> {
  const { pollMs = 2_000, stallMs = 120_000, signal } = opts;
  return new Promise<SettledOutcome>(resolve => {
    let done = false;
    let timer: ReturnType<typeof setTimeout> | undefined;
    const finish = (outcome: SettledOutcome) => {
      if (done) return;
      done = true;
      try { deps.onChanged.removeListener(onChange); } catch { /* noop */ }
      signal?.removeEventListener('abort', onAbort);
      if (timer) clearTimeout(timer);
      resolve(outcome);
    };
    const onChange = (delta: chrome.downloads.DownloadDelta) => {
      if (delta.id !== id) return;
      if (delta.state?.current === 'complete') finish('complete');
      else if (delta.state?.current === 'interrupted') finish('interrupted');
    };
    const onAbort = () => finish('aborted');

    let lastBytes = -1;
    let lastProgressAt = Date.now();
    const poll = async () => {
      if (done) return;
      let item: chrome.downloads.DownloadItem | undefined;
      let searched = false;
      try {
        const r = await Promise.resolve(deps.downloads.search({ id }));
        item = Array.isArray(r) ? r[0] : undefined;
        searched = true;
      } catch {
        // Transient search failure: the onChanged listener still covers the
        // terminal transition — but the stall clock must keep running, or a
        // permanently broken search would spin this poll forever.
        if (Date.now() - lastProgressAt > stallMs) return finish('stalled');
      }
      if (done) return;
      if (searched && !item) return finish('missing'); // erased from history: no file is coming
      if (item) {
        if (item.state === 'complete') return finish('complete');
        if (item.state === 'interrupted') return finish('interrupted');
        if (!item.filename || item.paused) {
          // Waiting on the user (target chooser open, or paused): park the clock.
          lastProgressAt = Date.now();
        } else {
          const received = item.bytesReceived || 0;
          if (received > lastBytes) {
            lastBytes = received;
            lastProgressAt = Date.now();
          } else if (Date.now() - lastProgressAt > stallMs) {
            return finish('stalled');
          }
        }
      }
      timer = setTimeout(() => void poll(), pollMs);
    };

    try { deps.onChanged.addListener(onChange); } catch { /* noop */ }
    if (signal) {
      if (signal.aborted) return finish('aborted');
      signal.addEventListener('abort', onAbort, { once: true });
    }
    void poll();
  });
}
