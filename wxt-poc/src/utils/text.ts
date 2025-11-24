export const normalizeText = (s: string) => (s ?? '').replace(/\u00a0/g, ' ');

export const isPlaceholderText = (s: string) => {
  const clean = normalizeText(s).trim();
  if (!clean) return true;
  return /^loading(?:\.\.\.?|…)?$/i.test(clean);
};

export const textScore = (s: string) => normalizeText(s).trim().length;

export const preferText = (prev: string, next: string | null | undefined) => {
  if (next == null) return prev;
  const nextPlaceholder = isPlaceholderText(next);
  const prevPlaceholder = isPlaceholderText(prev);
  if (nextPlaceholder && !prevPlaceholder) return prev;
  if (!nextPlaceholder && prevPlaceholder) return next;
  if (nextPlaceholder && prevPlaceholder) return prev;
  const prevLen = textScore(prev);
  const nextLen = textScore(next);
  return nextLen >= prevLen ? next : prev;
};

export const textFrom = (el: Element | null | undefined): string => {
  if (!el) return '';
  const h = el as HTMLElement;
  return (h.innerText || h.textContent || '').trim();
};

export const cssEscape = (s: string) => {
  if (typeof CSS !== 'undefined' && CSS.escape) return CSS.escape(s);
  return (s || '').toString().replace(/([\0-\x1f\x7f-\x9f!"#$%&'()*+,./:;<=>?@[\\\]^`{|}~])/g, '\\$1');
};
