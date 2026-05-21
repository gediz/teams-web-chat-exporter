// Layer 2 active probes for the Diagnostics page. Each probe is a
// single targeted check whose outcome maps to a specific failure
// class: DNS resolution, cookie partition, auth token availability,
// Teams surface detection, etc. The page renders the results in a
// checklist next to the Snapshot.
//
// Probes here are standalone: they do not require the page-world
// urlp helper RPC. The two helper-dependent probes live in
// `entrypoints/content.ts` next to the helper they exercise.

import { getIc3Token, getSkypeToken } from './api-client';
import { isTeamsUrl } from '../utils/teams-urls';
import type { ProbeResult, ProbeStatus } from '../utils/diagnostics';

// Hard ceiling for any single network probe. Generous: a healthy host
// answers in well under a second; if the request hangs longer, we
// surface that as `fail (timeout)` rather than blocking the probe run.
const PROBE_NET_TIMEOUT_MS = 5000;

async function timed(
  name: string,
  fn: () => Promise<{ status: ProbeStatus; detail?: string }>,
): Promise<ProbeResult> {
  const t0 = performance.now();
  try {
    const out = await fn();
    return { name, status: out.status, detail: out.detail, ms: Math.round(performance.now() - t0) };
  } catch (e) {
    return {
      name,
      status: 'fail',
      detail: e instanceof Error ? e.message : String(e),
      ms: Math.round(performance.now() - t0),
    };
  }
}

function probeTeamsOrigin(): Promise<ProbeResult> {
  return timed('teams_origin_recognized', async () => {
    if (isTeamsUrl(location.href)) return { status: 'pass', detail: location.host };
    return { status: 'fail', detail: `not a Teams host: ${location.host}` };
  });
}

function probeChatSurface(): Promise<ProbeResult> {
  return timed('chat_surface_detected', async () => {
    const chat = document.querySelector(
      '[data-tid="message-pane-list-viewport"], [data-tid="chat-message-list"], [data-tid="chat-pane"]',
    );
    const channel = document.querySelector(
      '[data-tid="channel-pane-runway"], [data-tid="channel-pane-message"], [data-tid="channel-pane"]',
    );
    if (chat) return { status: 'pass', detail: 'chat pane' };
    if (channel) return { status: 'pass', detail: 'channel pane' };
    return { status: 'fail', detail: 'no chat or channel pane in DOM' };
  });
}

// A real Skype / IC3 JWT is hundreds of characters. Treat anything
// shorter than this as "we found something but it's not plausibly a
// token" rather than confidently passing.
const MIN_PLAUSIBLE_TOKEN_LEN = 32;

function probeSkypeToken(): Promise<ProbeResult> {
  return timed('skype_token_extractable', async () => {
    const tok = await getSkypeToken();
    if (!tok) return { status: 'fail', detail: 'no token in storage' };
    if (tok.length < MIN_PLAUSIBLE_TOKEN_LEN) {
      return { status: 'fail', detail: `suspiciously short: ${tok.length} chars` };
    }
    return { status: 'pass', detail: `${tok.length} chars` };
  });
}

function probeIc3Token(): Promise<ProbeResult> {
  return timed('ic3_token_extractable', async () => {
    const tok = await getIc3Token();
    if (!tok) return { status: 'fail', detail: 'no token in storage' };
    if (tok.length < MIN_PLAUSIBLE_TOKEN_LEN) {
      return { status: 'fail', detail: `suspiciously short: ${tok.length} chars` };
    }
    return { status: 'pass', detail: `${tok.length} chars` };
  });
}

// Cheap reachability check: any HTTP response from the target host
// (even 4xx / 5xx) confirms DNS + TLS + the route is alive. Only a
// network error / timeout / DNS failure counts as `fail`.
//
// Both target hosts are covered by host_permissions:
//   - asm.skype.com is explicit (`*.asm.skype.com/*`)
//   - asyncgw matches `*.teams.microsoft.com/*` (Chrome's `*` host
//     segment matches multi-label subdomains)
// and Teams' own pages do these same cross-origin GETs, so asyncgw
// responds with CORS headers for the Teams origin we run in. A real
// HTTP status comes back (typically 401 for an unauthenticated probe
// URL), which is much more informative than the opaque success of
// `mode: 'no-cors'`.
async function reachabilityProbe(name: string, url: string): Promise<ProbeResult> {
  return timed(name, async () => {
    const ctrl = new AbortController();
    const timer = setTimeout(() => ctrl.abort(), PROBE_NET_TIMEOUT_MS);
    try {
      const resp = await fetch(url, {
        method: 'GET',
        signal: ctrl.signal,
        credentials: 'omit',
        cache: 'no-store',
      });
      return { status: 'pass', detail: `HTTP ${resp.status}` };
    } catch (e: unknown) {
      const msg = e instanceof Error ? e.message : String(e);
      const isAbort = msg.toLowerCase().includes('abort') || (e instanceof DOMException && e.name === 'AbortError');
      if (isAbort) {
        return { status: 'fail', detail: `timeout ${PROBE_NET_TIMEOUT_MS}ms` };
      }
      // Browser collapses several distinct failure modes into the same
      // `TypeError: Failed to fetch`. Flag this so the analyst doesn't
      // assume it's strictly a DNS / connectivity issue.
      return { status: 'fail', detail: `${msg} (DNS / CORS / CSP / TLS — not disambiguated)` };
    } finally {
      clearTimeout(timer);
    }
  });
}

function probeAsyncgwReachable(): Promise<ProbeResult> {
  return reachabilityProbe(
    'asyncgw_reachable',
    'https://eu-prod.asyncgw.teams.microsoft.com/v1/objects/probe/views/imgo',
  );
}

function probeAsmSkypeReachable(): Promise<ProbeResult> {
  return reachabilityProbe(
    'asm_skype_reachable',
    'https://us-api.asm.skype.com/v1/skypetokenauth',
  );
}

function probeIdbAccessible(): Promise<ProbeResult> {
  return timed('idb_accessible', async () => {
    if (typeof indexedDB === 'undefined') return { status: 'fail', detail: 'indexedDB not available' };
    if (typeof indexedDB.databases !== 'function') return { status: 'fail', detail: 'databases() unsupported' };
    const list = await indexedDB.databases();
    const teamsDbs = list.filter(d => typeof d.name === 'string' && d.name.startsWith('Teams:'));
    if (teamsDbs.length === 0) return { status: 'fail', detail: 'no Teams databases' };
    return { status: 'pass', detail: `${teamsDbs.length} Teams databases visible` };
  });
}

/**
 * Run every probe that does not require the page-world urlp helper.
 * Fired in parallel so the slowest network probe sets the floor for
 * total time. Caller adds the helper-dependent probes afterwards.
 */
export async function runStandaloneProbes(): Promise<ProbeResult[]> {
  return Promise.all([
    probeTeamsOrigin(),
    probeChatSurface(),
    probeIdbAccessible(),
    probeSkypeToken(),
    probeIc3Token(),
    probeAsyncgwReachable(),
    probeAsmSkypeReachable(),
  ]);
}
