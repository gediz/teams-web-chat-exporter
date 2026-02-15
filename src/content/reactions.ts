import type { Reaction } from '../types/shared';

export function extractReactions(item: Element): Reaction[] {
  const pills = Array.from(item.querySelectorAll<HTMLButtonElement>('[data-tid="diverse-reaction-pill-button"]'));
  const out: Reaction[] = [];

  const REACTION_EMOJI: Record<string, string> = {
    ok: '👌',
    like: '👍',
    thumbsup: '👍',
    thumbs_up: '👍',
    heart: '❤️',
    laugh: '😂',
    haha: '😂',
    surprised: '😮',
    wow: '😮',
    sad: '😢',
    angry: '😡',
    crossmark: '❌',
    skull: '💀',
    check: '✔️',
    checkmark: '✔️',
  };
  const normalizeKey = (val: string) =>
    val.toLowerCase().replace(/[^a-z0-9]+/g, '_').replace(/^_+|_+$/g, '');
  const parseEmojiFromItemId = (itemId: string) => {
    if (!itemId) return '';
    const parts = itemId.split('_');
    const codes: number[] = [];
    for (const part of parts) {
      if (/^[0-9a-f]{4,6}$/i.test(part)) {
        codes.push(parseInt(part, 16));
      }
    }
    if (!codes.length) return '';
    try {
      return String.fromCodePoint(...codes);
    } catch {
      return '';
    }
  };
  const getItemId = (img: HTMLImageElement | null, container: Element | null) => {
    const direct = img?.getAttribute('itemid') || container?.getAttribute('itemid') || '';
    if (direct) return direct;
    const src = img?.getAttribute('src') || '';
    const match = src.match(/\/emoticons\/([^/]+)\//i);
    if (!match) return '';
    try { return decodeURIComponent(match[1]); } catch { return match[1]; }
  };
  const keyFromLabelledBy = (labelledBy: string | null) => {
    if (!labelledBy) return '';
    const id = labelledBy.split(/\s+/).find(val => val.startsWith('message-'));
    if (!id) return '';
    const m = id.match(/^message-([^-]+)-/);
    return m ? m[1] : '';
  };
  const keyFromLabelText = (labelText: string) => {
    const m = labelText.match(/\b([A-Za-z ]+)\s+reaction\b/i);
    return m ? normalizeKey(m[1]) : '';
  };

  for (const btn of pills) {
    const self = btn.getAttribute('aria-pressed') === 'true';
    const emojiImg = btn.querySelector<HTMLImageElement>('[data-tid="emoticon-renderer"] img');
    const emojiContainer = btn.querySelector('[data-tid="emoticon-renderer"]');
    let emoji =
      emojiImg?.getAttribute('alt') ||
      emojiImg?.getAttribute('aria-label') ||
      emojiContainer?.getAttribute('aria-label') ||
      '';

    let labelText = '';
    const labelledBy = btn.getAttribute('aria-labelledby');
    if (labelledBy) {
      labelText = labelledBy
        .split(/\s+/)
        .map(id => (document.getElementById(id)?.innerText || '').trim())
        .filter(Boolean)
        .join(' ')
        .trim();
    }
    if (!labelText) {
      labelText = (btn.getAttribute('aria-label') || '').trim();
    }

    if (!emoji) {
      const labelEmoji = (labelText.match(/[\p{Extended_Pictographic}]/u) || [])[0];
      if (labelEmoji) emoji = labelEmoji;
    }
    if (!emoji) {
      const itemId = getItemId(emojiImg, emojiContainer);
      const fromItemId = parseEmojiFromItemId(itemId);
      if (fromItemId) {
        emoji = fromItemId;
      } else if (itemId) {
        const mapped = REACTION_EMOJI[normalizeKey(itemId)];
        emoji = mapped || `:${itemId}:`;
      }
    }
    if (!emoji) {
      const key = keyFromLabelledBy(labelledBy) || keyFromLabelText(labelText);
      if (key && REACTION_EMOJI[key]) emoji = REACTION_EMOJI[key];
    }

    const count = parseInt((labelText.match(/\d+/) || ['1'])[0], 10) || 1;
    const entry: Reaction = self ? { emoji, count, self: true } : { emoji, count };

    const looksLikeSummary =
      /^\d+\s+\w+\s+reactions?\b/i.test(labelText) ||
      /^\d+\s+reactions?\b/i.test(labelText);
    if (labelText && !looksLikeSummary && /react/i.test(labelText)) {
      const beforeReact = labelText.split(/react/i)[0];
      let names = beforeReact
        .split(/,\s*|\s+and\s+/)
        .map(s => s.trim())
        .filter(Boolean)
        .filter(s => !/^\d+\s+others?$/i.test(s) && !/others$/i.test(s));
      if (names.length) entry.reactors = Array.from(new Set(names)).slice(0, 100);
    }

    out.push(entry);
  }
  return out;
}
