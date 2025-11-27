import { textFrom } from '../utils/text';

// Extract text including emojis (IMG alt) and preserve basic block breaks.
export function extractTextWithEmojis(root: Element | null): string {
  if (!root) return '';
  let out = '';
  const walk = (n: ChildNode) => {
    if (n.nodeType === Node.TEXT_NODE) {
      out += n.nodeValue;
      return;
    }
    if (n.nodeType !== Node.ELEMENT_NODE) return;
    const el = n as Element;
    const tag = el.tagName;
    if (tag === 'BR') {
      out += '\n';
      return;
    }
    if (tag === 'IMG') {
      out += (el.getAttribute('alt') || el.getAttribute('aria-label') || '');
      return;
    }
    const blockish = /^(DIV|P|LI|BLOCKQUOTE)$/;
    const start = out.length;
    for (const c of el.childNodes) walk(c);
    if (blockish.test(tag) && out.length > start) out += '\n';
  };
  walk(root);
  return out.replace(/\n{3,}/g, '\n\n').trim();
}

// Normalize Teams mentions into plain text (@name).
export function normalizeMentions(root: Element) {
  if (!root || typeof root.querySelectorAll !== 'function') return;
  const wrappers = Array.from(
    root.querySelectorAll('[data-lpc-hover-target-id][aria-label^="Mentioned"], [itemtype*="schema.skype.com/Mention"]')
  );
  if (!wrappers.length) return;
  const processed = new Set<Element>();
  for (const node of wrappers) {
    const wrapper = node.closest('[data-lpc-hover-target-id][aria-label^="Mentioned"]') || node;
    if (!wrapper || processed.has(wrapper) || !root.contains(wrapper)) continue;
    const parent = wrapper.parentElement;
    const group: Element[] = [];
    if (parent) {
      for (const sibling of Array.from(parent.childNodes)) {
        if (sibling === wrapper) {
          group.push(wrapper);
          continue;
        }
        if (sibling.nodeType === Node.ELEMENT_NODE) {
          const sibEl = sibling as Element;
          const mentionWrapper = sibEl.closest?.('[data-lpc-hover-target-id][aria-label^="Mentioned"]');
          if (mentionWrapper && wrappers.includes(mentionWrapper)) {
            group.push(mentionWrapper);
            processed.add(mentionWrapper);
            continue;
          }
          if (group.length) break;
        } else if (sibling.nodeType === Node.TEXT_NODE) {
          if (!sibling.textContent?.trim() && group.length) continue;
          if (group.length) break;
        }
      }
    }
    if (!group.length) {
      group.push(wrapper);
    }
    let combined =
      group
        .map(el => (el.textContent || '').replace(/\s+/g, ' ').trim())
        .filter(Boolean)
        .join(' ') ||
      group
        .map(el => (el.getAttribute('aria-label') || '').replace(/^Mentioned\s+/i, '').trim())
        .filter(Boolean)
        .join(' ');
    combined = combined.trim();
    if (!combined) continue;
    const owner = wrapper.ownerDocument || document;
    const replacement = owner.createTextNode(`@${combined}`);
    for (const el of group) {
      processed.add(el);
      if (el !== wrapper) el.remove();
    }
    wrapper.replaceWith(replacement);
  }
}

// Fallbacks for reply headers use this to grab text.
export function textFromHeading(heading: Element | null): string {
  return textFrom(heading);
}
