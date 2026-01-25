import { textFrom } from '../utils/text';

// Extract text including emojis (IMG alt) and preserve basic block breaks.
export function extractTextWithEmojis(root: Element | null): string {
  if (!root) return '';
  let out = '';
  const extractCodeBlock = (el: Element) => {
    let code = '';
    const walkCode = (n: ChildNode) => {
      if (n.nodeType === Node.TEXT_NODE) {
        code += n.nodeValue;
        return;
      }
      if (n.nodeType !== Node.ELEMENT_NODE) return;
      const child = n as Element;
      const tagName = child.tagName;
      if (tagName === 'BR') {
        code += '\n';
        return;
      }
      if (tagName === 'IMG') {
        code += (child.getAttribute('alt') || child.getAttribute('aria-label') || '');
        return;
      }
      for (const c of child.childNodes) walkCode(c);
    };
    walkCode(el);
    return code.replace(/\u00a0/g, ' ').replace(/\n+$/, '');
  };
  const collectText = (node: Element | null): string => {
    if (!node) return '';
    let buf = '';
    const walkCollect = (n: ChildNode) => {
      if (n.nodeType === Node.TEXT_NODE) {
        buf += n.nodeValue;
        return;
      }
      if (n.nodeType !== Node.ELEMENT_NODE) return;
      const el = n as Element;
      const tag = el.tagName;
      if (tag === 'BR') {
        buf += '\n';
        return;
      }
      if (tag === 'IMG') {
        buf += (el.getAttribute('alt') || el.getAttribute('aria-label') || '');
        return;
      }
      if (tag === 'CODE') {
        buf += '`';
        for (const c of el.childNodes) walkCollect(c);
        buf += '`';
        return;
      }
      if (tag === 'PRE') {
        const code = extractCodeBlock(el);
        if (code) buf += `\n\`\`\`\n${code}\n\`\`\`\n`;
        return;
      }
      const blockish = /^(DIV|P|LI|BLOCKQUOTE)$/;
      const start = buf.length;
      for (const c of el.childNodes) walkCollect(c);
      if (blockish.test(tag) && buf.length > start) buf += '\n';
    };
    walkCollect(node);
    return buf.replace(/\n{3,}/g, '\n\n').trim();
  };
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
    if (tag === 'CODE') {
      out += '`';
      for (const c of el.childNodes) walk(c);
      out += '`';
      return;
    }
    if (tag === 'PRE') {
      const code = extractCodeBlock(el);
      if (code) out += `\n\`\`\`\n${code}\n\`\`\`\n`;
      return;
    }
    if (tag === 'BLOCKQUOTE') {
      const quoted = collectText(el);
      if (quoted) {
        const lines = quoted.split(/\n/);
        if (out && !out.endsWith('\n')) out += '\n';
        out += lines.map(line => (line ? `> ${line}` : '>')).join('\n');
        out += '\n';
      }
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

export function extractTables(root: Element | null): string[][][] {
  if (!root) return [];
  const skip = [
    '[data-tid="quoted-reply-card"]',
    '[data-tid="referencePreview"]',
    '[role="group"][aria-label^="Begin Reference"]',
    '[data-tid="adaptive-card"]',
    '.ac-adaptiveCard',
    '[aria-label*="card message"]',
  ];
  const tables = Array.from(root.querySelectorAll<HTMLTableElement>('table[itemprop="copy-paste-table"]'));
  const out: string[][][] = [];
  for (const table of tables) {
    if (skip.some(sel => table.closest(sel))) continue;
    const rows: string[][] = [];
    table.querySelectorAll('tr').forEach(tr => {
      const cells = Array.from(tr.querySelectorAll<HTMLElement>('th, td'));
      if (!cells.length) return;
      const row = cells
        .map(cell => extractTextWithEmojis(cell).trim())
        .map(cell => cell.replace(/\s+\n/g, '\n').replace(/[ \t]{2,}/g, ' '));
      rows.push(row);
    });
    if (rows.length) out.push(rows);
  }
  return out;
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
