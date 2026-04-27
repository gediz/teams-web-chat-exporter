import type { ConversationSummary } from '../types/shared';

type TFn = (key: string, params?: Record<string, string | number>, lang?: string) => string;

// Intl.ListFormat ships in every browser we target (Chrome 72+, Firefox
// 78+) but the current tsconfig's lib doesn't declare it. A local
// feature-detected reference avoids bumping the whole project's lib
// target for one call site.
type ListFormatCtor = new (
  locales?: string | string[],
  options?: { type?: 'conjunction' | 'disjunction' | 'unit'; style?: 'long' | 'short' | 'narrow' },
) => { format(list: string[]): string };

// First word of each name, sidebar-style. "Jane Doe" → "Jane",
// "John Smith" → "John". Matches what Teams shows in its own
// sidebar for unnamed group rows; less verbose than full names in
// the narrow row.
const firstName = (full: string): string => (full || '').trim().split(/\s+/)[0] || full;

const fmtMemberList = (items: string[], lang: string): string => {
  const firsts = items.map(firstName);
  try {
    const LF = (Intl as unknown as { ListFormat?: ListFormatCtor }).ListFormat;
    if (LF) return new LF(lang, { type: 'unit', style: 'short' }).format(firsts);
  } catch { /* fall through */ }
  return firsts.join(', ');
};

/**
 * Resolve the primary display name for a conversation row, applying
 * locale-aware suffixes and fallbacks:
 *   - self-chat  → "<name> (<locale 'You'>)" or the pure fallback
 *   - unnamed group with roster → "A, B, C" joined via Intl.ListFormat
 *   - missing name for each kind → the kind's localized "(untitled …)"
 *
 * Kept as a pure function (no Svelte dependency) so the popup's picker
 * and the export-filename builder share one source of truth for how a
 * conversation is labelled in the active language.
 */
export function conversationDisplayName(
  c: ConversationSummary,
  lang: string,
  t: TFn,
): string {
  if (c.name) {
    if (c.isSelfChat) {
      const suffix = t('picker.selfChatSuffix', {}, lang) || 'You';
      return `${c.name} (${suffix})`;
    }
    return c.name;
  }
  if (c.groupMembers && c.groupMembers.length > 0) {
    // Sidebar parity: "First, First, First" for the visible names + an
    // inline "+N" suffix when there are more we didn't fit. Keeps the
    // picker row visually consistent with how Teams renders the same
    // unnamed group in its own sidebar.
    const visible = fmtMemberList(c.groupMembers, lang);
    if (c.groupExtraMembers && c.groupExtraMembers > 0) {
      return `${visible}, +${c.groupExtraMembers}`;
    }
    return visible;
  }
  if (c.isSelfChat) return t('picker.selfChatFallback', {}, lang) || '(You)';
  switch (c.kind) {
    case 'chat':    return t('picker.chatFallback',    {}, lang) || '(direct message)';
    case 'group':   return t('picker.groupFallback',   {}, lang) || '(unnamed group)';
    case 'meeting': return t('picker.meetingFallback', {}, lang) || '(untitled meeting)';
    case 'channel': return t('picker.channelFallback', {}, lang) || '(untitled channel)';
  }
}

/** Composite subtitle string (self-chat label + API-provided subtitle + "+N more"). */
export function conversationDisplaySubtitle(
  c: ConversationSummary,
  lang: string,
  t: TFn,
): string | undefined {
  const parts: string[] = [];
  if (c.isSelfChat) parts.push(t('picker.selfChatLabel', {}, lang) || 'Self-chat');
  if (c.subtitle) parts.push(c.subtitle);
  // groupExtraMembers is now baked into the primary name (sidebar
  // parity), so it isn't echoed in the subtitle.
  return parts.length > 0 ? parts.join(' · ') : undefined;
}
