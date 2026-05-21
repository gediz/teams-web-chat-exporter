// Diagnostic report assembly and redaction.
//
// Collects environment, options, IDB shape, recent exports, the
// log buffer (BG + content-forwarded), and Layer 2 probe results.
// Output is a JSON document safe to paste into a public bug report
// once tokenization is applied.
//
// Redaction: per-report random salt plus SHA-256(salt || id), truncated
// to 4 bytes (8 hex chars). Same id within one report tokenizes
// consistently; two reports tokenize the same id differently. Salt
// lives in memory only and is discarded when the builder returns.

export type EnvInfo = {
  extensionVersion: string;
  manifestVersion: number;
  browserBrand: string;
  browserVersion: string;
  os: string;
  teamsTabUrl: string | null;
  teamsHost: string | null;
  locale: string;
  documentLang: string | null;
  prefersColorScheme: 'light' | 'dark' | 'unknown';
};

export type PermissionsInfo = {
  declaredHostPermissions: string[];
  optionalAllUrlsGranted: boolean | null;
};

export type OptionsSnapshot = Record<string, unknown>;
export type OptionsSection =
  | { available: true; values: OptionsSnapshot }
  | { available: false; reason: string };

export type IdbStoreInfo = { name: string; count: number; error?: string };
export type IdbDatabaseInfo = {
  name: string;
  version: number;
  status: 'opened' | 'blocked' | 'error';
  reason?: string;
  stores: IdbStoreInfo[];
};
export type IdbShape =
  | { available: true; databases: IdbDatabaseInfo[] }
  | { available: false; reason: string };

export type ExportSummary = {
  savedAt: number;
  kind: 'success' | 'cancelled' | 'failed' | 'partial';
  partialReason?: 'network' | 'truncation';
  formats?: string[];
  messageCount?: number;
  elapsedMs?: number;
  isZip?: boolean;
};
export type ExportsSection =
  | { available: true; items: ExportSummary[] }
  | { available: false; reason: string };

// Combined log buffer shape. Each entry carries its origin so the
// report renderer can split into LOGS_BACKGROUND / LOGS_CONTENT
// sections. BG owns the buffer; content forwards its captures to BG.
export type DiagLogEntry = { src: 'bg' | 'content'; ts: number; level: string; line: string };

export type LogTail = {
  entries: DiagLogEntry[] | null;
  missing?: string;
  // Bytes currently used by the persisted buffer in chrome.storage.local.
  // Null when persistence is off or the metric is unavailable.
  bytesUsed: number | null;
  // Whether the user has enabled persistence. Surfaced in the report so
  // the analyst can tell if older logs would have been preserved
  // through SW eviction.
  persistEnabled: boolean;
  // Most recent persistence write failure, if any. Lets the analyst
  // tell a quota / corrupt-storage failure apart from "no writes yet".
  lastFlushError?: { ts: number; reason: string } | null;
  // Content-script forwarding stats. Counts batches and entries that
  // could not be shipped to BG (SW evicted, port dropped, etc.). If
  // non-zero, the analyst knows some content-script log lines were
  // lost between capture and persistence.
  forwarding?: { lostBatches: number; lostEntries: number; lastError: string | null };
};

// Layer 2: active probes. Each result is the outcome of a single
// targeted check (DNS reachability, token extraction, helper
// injection, etc.). `pass` means the check completed successfully,
// `fail` means it returned a definite negative (the detail string
// carries the specifics: HTTP status, error message, etc.), and
// `skipped` is used for checks that don't apply on the current
// tenant (e.g. Skype JWT extraction skips on work / school).
export type ProbeStatus = 'pass' | 'fail' | 'skipped';
export type ProbeResult = {
  name: string;
  status: ProbeStatus;
  detail?: string;
  ms: number;
};
// Discriminated by `state`, not `available`. The earlier shape mixed a
// boolean and a string literal in the same field, which let careless
// code path-test with `if (probes.available)` and treat 'not-run' as
// available. A named state avoids that footgun.
export type ProbesSection =
  | { state: 'not-run' }
  | { state: 'done'; results: ProbeResult[]; runAt: number; totalMs: number }
  | { state: 'failed'; reason: string };

export type DiagnosticReportInput = {
  env: EnvInfo;
  permissions: PermissionsInfo;
  options: OptionsSection;
  idb: IdbShape;
  exports: ExportsSection;
  logs: LogTail;
  probes: ProbesSection;
};

export type ReportBuildOptions = {
  includeRawIds?: boolean;
};

// Tokenizer

type Tokenizer = {
  token(kind: string, id: string): Promise<string>;
};

async function makeTokenizer(): Promise<Tokenizer> {
  const salt = crypto.getRandomValues(new Uint8Array(16));
  const cache = new Map<string, string>();
  return {
    async token(kind: string, id: string): Promise<string> {
      const key = `${kind} ${id}`;
      const cached = cache.get(key);
      if (cached) return cached;
      const data = new TextEncoder().encode(id);
      const buf = new Uint8Array(salt.length + data.length);
      buf.set(salt);
      buf.set(data, salt.length);
      const hash = await crypto.subtle.digest('SHA-256', buf);
      const bytes = new Uint8Array(hash, 0, 4);
      const hex = Array.from(bytes, b => b.toString(16).padStart(2, '0')).join('');
      const out = `<${kind} ${hex}>`;
      cache.set(key, out);
      return out;
    },
  };
}

const UUID_RE = /\b[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\b/gi;
const MRI_RE = /\b(?:8:orgid:|8:live:|8:teamsvisitor:|gid:|28:)([A-Za-z0-9._-]+)/g;
// SharePoint tenant subdomain. Format examples:
//   amlogicglobaleur-my.sharepoint.com  (work/school OneDrive)
//   contoso.sharepoint.com
// The first label names the organisation. Tokenized; the
// `.sharepoint.<tld>` suffix is preserved as URL-shape context.
const SP_TENANT_RE = /\b([a-z0-9-]+?)((?:-my)?\.sharepoint\.(?:com|us|cn|de))\b/gi;
// SharePoint `/personal/<slug>` path segment. The slug is the user's
// email with `@` and `.` rewritten to `_` — e.g.
//   alice_smith_contoso_com  ←  alice.smith@contoso.com
// This is a direct PII leak (an analyst can reconstruct the email
// trivially). Anchored on the `sharepoint.<tld>/personal/` context
// to avoid false-matching unrelated `/personal/` paths.
const SP_PERSONAL_RE = /(sharepoint\.[a-z]+\/personal\/)([a-z0-9_-]+)/gi;
// Asyncgw / AMS regional hostnames. Format:
//   eu-prod.asyncgw.teams.microsoft.com
//   noam-prod.asyncgw.teams.microsoft.com
//   us-api.asm.skype.com
//   euno1-api-0.asm.skype.com
// The leading sublabel names the Azure region serving the user's
// tenant, which correlates with the org's data-residency choice.
// Same rationale as AMS object IDs: redact the region, keep the
// fixed suffix as URL-shape context.
const ASYNCGW_HOST_RE = /\b([a-z0-9]+)(-prod\.asyncgw\.teams\.microsoft\.com)\b/g;
const ASM_HOST_RE = /\b([a-z0-9]+(?:-[a-z0-9]+)*)(\.asm\.skype\.com)\b/g;
// Teams conversation / meeting / channel thread IDs. Shape examples:
//   19:meeting_<32hex>@thread.v2
//   19:<uuid>_<uuid>@unq.gbl.spaces
//   19:<22-char base64-ish>@thread.v2  (group chat)
//   19:<id>@thread.tacv2
//   19:cce85c22ba1f4144afca0601814...  (truncated in logs; trailing
//                                       `...` is literal, excluded by
//                                       the char class)
// The `@<suffix>` is preserved when present so the analyst can tell
// chat from meeting; when absent (truncated log line) the id is still
// tokenized. Minimum id length 16 chars to avoid false positives on
// short `19:foo`-style debug shorthand.
const THREAD_RE = /\b19:([A-Za-z0-9_-]{16,})(?:@(thread\.v2|thread\.tacv2|unq\.gbl\.spaces|skype))?/g;
// AMS object identifiers as they appear in asyncgw / asm.skype.com
// URLs. Shape: `<digit>-<region>-<server>-<hex>` e.g.
//   0-frca-d3-c8158b104c4510990350396c2d197afa
//   0-weu-d16-3e640a12098d2b421ca0420ba028411a
// The whole id is tokenized including the region+server prefix:
// `frca`, `weu`, etc. name specific Azure regions which correlate
// with the user's tenant location. Keeping just `<obj abcd1234>`
// loses no diagnostic value (the surrounding URL path already
// labels the asset as AMS) and avoids leaking tenant geo in a
// public bug report.
const AMS_OBJ_RE = /\b\d+-[a-z]+\d*-[a-z]+\d*-[a-f0-9]{20,}\b/g;
const EMAIL_RE = /\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b/g;
const JWT_RE = /\beyJ[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}\b/g;

async function redactString(t: Tokenizer, raw: string): Promise<string> {
  if (!raw) return raw;
  let out = raw;
  out = await replaceAllAsync(out, JWT_RE, async (m) => t.token('jwt', m));
  out = await replaceAllAsync(out, EMAIL_RE, async (m) => t.token('email', m.toLowerCase()));
  out = await replaceAllAsync(out, MRI_RE, async (m, p1) => {
    // Keep the prefix (8:orgid: / 8:live: / 28: ...) visible so the
    // analyst can still tell Work-School from Free from bot, then
    // wrap the user-identifier part as a token. Result reads e.g.
    // `8:orgid:<mri a1b2c3d4>`.
    const prefix = m.slice(0, m.length - p1.length);
    const tok = await t.token('mri', p1);
    return prefix + tok;
  });
  out = await replaceAllAsync(out, THREAD_RE, async (m, p1, p2) => {
    // Preserve the `19:` marker. Append the `@<suffix>` when it was
    // present in the source; truncated log lines (no @) just get
    // `19:<thread ...>`.
    const tok = await t.token('thread', p1);
    return p2 ? `19:${tok}@${p2}` : `19:${tok}`;
  });
  out = await replaceAllAsync(out, AMS_OBJ_RE, async (m) => {
    // Full id (region+server+hex) tokenizes as one. Keying the hash
    // on the entire string means two reports for the same asset
    // tokenize the same way within the report; cross-report
    // correlation is still blocked by the per-report salt.
    return t.token('obj', m);
  });
  out = await replaceAllAsync(out, SP_TENANT_RE, async (m, tenant, suffix) => {
    const tok = await t.token('tenant', tenant);
    return tok + suffix;
  });
  out = await replaceAllAsync(out, SP_PERSONAL_RE, async (m, prefix, slug) => {
    const tok = await t.token('user', slug);
    return prefix + tok;
  });
  out = await replaceAllAsync(out, ASYNCGW_HOST_RE, async (m, region, suffix) => {
    const tok = await t.token('region', region);
    return tok + suffix;
  });
  out = await replaceAllAsync(out, ASM_HOST_RE, async (m, region, suffix) => {
    const tok = await t.token('region', region);
    return tok + suffix;
  });
  out = await replaceAllAsync(out, UUID_RE, async (m) => t.token('id', m.toLowerCase()));
  return out;
}

async function replaceAllAsync(
  s: string,
  re: RegExp,
  repl: (m: string, ...groups: string[]) => Promise<string>,
): Promise<string> {
  const matches: { start: number; end: number; replacement: string }[] = [];
  // Clone so two callers cannot stomp each other's lastIndex. The
  // module-scoped patterns are reused across every redactString call.
  const local = new RegExp(re.source, re.flags);
  let m: RegExpExecArray | null;
  while ((m = local.exec(s)) !== null) {
    const start = m.index;
    const end = start + m[0].length;
    const replacement = await repl(m[0], ...m.slice(1));
    matches.push({ start, end, replacement });
    if (m[0].length === 0) local.lastIndex++;
  }
  if (!matches.length) return s;
  let out = '';
  let cursor = 0;
  for (const mm of matches) {
    out += s.slice(cursor, mm.start) + mm.replacement;
    cursor = mm.end;
  }
  out += s.slice(cursor);
  return out;
}

// Report builder

// JSON-only output. The page's Preview, Copy, and Save buttons all
// deliver the same JSON; one serialisation, three delivery paths.
// Same redaction passes apply; identifier shapes get tokenized,
// salt is per-build and discarded after return.
export async function buildDiagnosticJson(
  input: DiagnosticReportInput,
  opts: ReportBuildOptions = {},
): Promise<string> {
  const t = await makeTokenizer();
  const redact = opts.includeRawIds
    ? async (s: string) => s
    : (s: string) => redactString(t, s);

  // Deep-redact strings in any value. Recurses arrays / records.
  const walk = async (v: unknown): Promise<unknown> => {
    if (typeof v === 'string') return await redact(v);
    if (Array.isArray(v)) {
      const out: unknown[] = [];
      for (const item of v) out.push(await walk(item));
      return out;
    }
    if (v && typeof v === 'object') {
      const out: Record<string, unknown> = {};
      for (const [k, val] of Object.entries(v)) out[k] = await walk(val);
      return out;
    }
    return v;
  };

  const redactedInput = await walk(input) as DiagnosticReportInput;
  const payload = {
    generatedAt: new Date().toISOString(),
    rawIds: !!opts.includeRawIds,
    report: redactedInput,
  };
  return JSON.stringify(payload, null, 2);
}
