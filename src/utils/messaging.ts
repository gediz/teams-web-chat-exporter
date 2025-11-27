import type { RuntimeRequest, RuntimeResponse } from '../types/messaging';

type RuntimeLike = typeof chrome.runtime;

/**
 * Thin wrapper around runtime.sendMessage with Promise + type support.
 * Handles both callback- and promise-based runtimes (Chrome vs Firefox).
 */
export function runtimeSend<T extends RuntimeRequest>(
  runtime: RuntimeLike,
  msg: T,
): Promise<RuntimeResponse<T>> {
  return new Promise((resolve, reject) => {
    let settled = false;
    const done = (value: unknown, error?: unknown) => {
      if (settled) return;
      settled = true;
      if (error) {
        reject(error instanceof Error ? error : new Error(String(error)));
        return;
      }
      resolve(value as RuntimeResponse<T>);
    };

    const callback = (resp: unknown) => {
      const err = runtime.lastError;
      if (err) {
        done(undefined, new Error(err.message || 'Runtime message failed'));
        return;
      }
      done(resp);
    };

    try {
      const maybePromise = runtime.sendMessage(msg, callback as never);
      // Firefox can return a Promise if callback is omitted; guard to avoid double resolve.
      if (typeof maybePromise === 'object' && typeof (maybePromise as Promise<unknown>)?.then === 'function') {
        (maybePromise as Promise<unknown>).then(
          resp => done(resp),
          err => done(undefined, err),
        );
      }
    } catch (e) {
      done(undefined, e);
    }
  });
}
