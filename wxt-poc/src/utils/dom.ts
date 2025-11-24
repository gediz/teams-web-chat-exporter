export const $$ = <T extends Element = Element>(sel: string, root: Document | Element = document): T[] =>
  Array.from(root.querySelectorAll(sel)) as T[];

export const $ = <T extends Element = Element>(sel: string, root: Document | Element = document): T | null =>
  root.querySelector(sel) as T | null;
