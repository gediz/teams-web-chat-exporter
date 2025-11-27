import type { Reaction } from '../types/shared';

export function extractReactions(item: Element): Reaction[] {
  const pills = Array.from(item.querySelectorAll<HTMLButtonElement>('[data-tid="diverse-reaction-pill-button"]'));
  const out: Reaction[] = [];

  for (const btn of pills) {
    const emoji = btn.querySelector('[data-tid="emoticon-renderer"] img')?.getAttribute('alt') || '';

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

    const count = parseInt((labelText.match(/\d+/) || ['1'])[0], 10) || 1;
    const entry: Reaction = { emoji, count };

    if (labelText) {
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
