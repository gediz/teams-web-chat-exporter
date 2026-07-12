// Behavior lock for waitForDownloadSettled (src/background/download-wait.ts) — the
// settlement oracle that decides whether the export artifact actually reached disk
// before the attachments phase starts. A wrong terminal classification means the
// Files phase runs against a missing zip, or the export hangs. Fully DI'd, so we
// drive it with a hand-rolled onChanged emitter + a stub downloads.search.
import { test, expect, vi } from 'vitest';
import { waitForDownloadSettled } from '../src/background/download-wait.ts';

function makeOnChanged() {
  const listeners = new Set<(d: unknown) => void>();
  return {
    addListener: (cb: (d: unknown) => void) => listeners.add(cb),
    removeListener: (cb: (d: unknown) => void) => listeners.delete(cb),
    emit: (delta: unknown) => listeners.forEach((cb) => cb(delta)),
  };
}
const inProgress = (bytes = 0) => [{ id: 1, state: 'in_progress', filename: '/x/out.zip', bytesReceived: bytes }];

test('resolves "complete" on an onChanged complete transition', async () => {
  const onChanged = makeOnChanged();
  const deps = { downloads: { search: async () => inProgress() }, onChanged } as never;
  const p = waitForDownloadSettled(deps, 1, { pollMs: 100, stallMs: 5000 });
  onChanged.emit({ id: 1, state: { current: 'complete' } });
  await expect(p).resolves.toBe('complete');
});

test('resolves "interrupted" on an onChanged interrupted transition', async () => {
  const onChanged = makeOnChanged();
  const deps = { downloads: { search: async () => inProgress() }, onChanged } as never;
  const p = waitForDownloadSettled(deps, 1, { pollMs: 100, stallMs: 5000 });
  onChanged.emit({ id: 1, state: { current: 'interrupted' } });
  await expect(p).resolves.toBe('interrupted');
});

test('resolves "missing" when the item is gone from downloads.search (erased history)', async () => {
  const onChanged = makeOnChanged();
  const deps = { downloads: { search: async () => [] }, onChanged } as never;
  await expect(waitForDownloadSettled(deps, 1, { pollMs: 100 })).resolves.toBe('missing');
});

test('resolves "complete" when the poll itself observes a completed item', async () => {
  const onChanged = makeOnChanged();
  const deps = { downloads: { search: async () => [{ id: 1, state: 'complete', filename: '/x/out.zip' }] }, onChanged } as never;
  await expect(waitForDownloadSettled(deps, 1, { pollMs: 100 })).resolves.toBe('complete');
});

test('resolves "aborted" when the signal aborts', async () => {
  const onChanged = makeOnChanged();
  const ac = new AbortController();
  const deps = { downloads: { search: async () => inProgress() }, onChanged } as never;
  const p = waitForDownloadSettled(deps, 1, { pollMs: 100, signal: ac.signal });
  ac.abort();
  await expect(p).resolves.toBe('aborted');
});

test('resolves "stalled" when bytes stop advancing past stallMs', async () => {
  vi.useFakeTimers();
  const onChanged = makeOnChanged();
  const deps = { downloads: { search: async () => inProgress(0) }, onChanged } as never; // bytes never grow
  try {
    const p = waitForDownloadSettled(deps, 1, { pollMs: 100, stallMs: 500 });
    await vi.advanceTimersByTimeAsync(1200); // well past the 500ms stall window with no progress
    await expect(p).resolves.toBe('stalled');
  } finally { vi.useRealTimers(); }
});
