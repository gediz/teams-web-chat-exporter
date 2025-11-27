import type { ReplyContext } from '../types/shared';
import { extractTextWithEmojis } from './text';
import { textFrom } from '../utils/text';

export function extractReplyContext(item: Element, body: Element): ReplyContext | null {
  // 1) Structured quoted-reply card (preferred)
  const card = body?.querySelector('[data-tid="quoted-reply-card"]');
  if (card) {
    const tsEl = card.querySelector('[data-tid="quoted-reply-timestamp"]');
    const authorEl = tsEl?.previousElementSibling || null; // author sits right before timestamp
    const textEl = card.querySelector('[data-tid="quoted-reply-preview-content"]');

    const author = textFrom(authorEl);
    const timestamp = textFrom(tsEl); // e.g. "12/09/2025, 11:12"
    const text = extractTextWithEmojis(textEl || card).trim(); // full preview text

    if (author || timestamp || text) return { author, timestamp, text };
  }

  // 2) Fallback: aria-label on the group with "Begin Reference, …, <author>, <timestamp>, End reference"
  const group = body?.querySelector('[role="group"][aria-label^="Begin Reference"]');
  if (group) {
    const aria = group.getAttribute('aria-label') || '';
    const m = aria.match(/^Begin Reference,\s*([\s\S]*),\s*([^,]+),\s*([^,]+),\s*End reference$/);
    if (m) {
      const [, text, author, timestamp] = m;
      return { author: (author || '').trim(), timestamp: (timestamp || '').trim(), text: (text || '').trim() };
    }
  }

  // 3) Legacy fallback: "Begin Reference, … by <author>"
  const heading = item.querySelector('div[role="heading"]');
  if (heading) {
    const raw = textFrom(heading);
    const m = raw.match(/Begin Reference,\s*(.*?)\s*by\s*(.+)$/i);
    if (m) return { author: m[2].trim(), timestamp: '', text: m[1].trim() };
  }
  return null;
}
