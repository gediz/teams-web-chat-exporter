/**
 * Table model
 *
 * Teams sends each message body as HTML. A table pasted into a chat arrives
 * as `<table itemprop="copy-paste-table">`. The plain-text flattening in
 * htmlToText() collapses that to "cell | cell" rows, which a rowspan merge
 * shifts out of alignment (the reported bug: tables stop looking like tables).
 *
 * This module parses a message body into ordered BodyBlocks — runs of text
 * and structured tables in their original positions — so every export
 * renderer (HTML, PDF, TXT, JSON) can rebuild a real table. It runs in the
 * content script (DOMParser available); the background builders consume the
 * serializable result via ExportMessage.bodyBlocks.
 *
 * Representation choices (see tce-debug/table-design/DESIGN.md and the research
 * behind it): the dense `rows` grid REPEATS a spanned cell's value into every
 * position it covers (the pandas read_html / tidy-data convention, so every
 * row is self-contained), and `merges` records each span once on its anchor
 * cell (the Excel / Google Sheets / Handsontable convention, so the merge is
 * losslessly recoverable). Together they are a superset of both conventions.
 */

import type { TableData, BodyBlock, MergeRegion } from '../types/shared';

// Every <table> in a message body is treated as real tabular content. Only
// about half carry itemprop="copy-paste-table" (that marker is added when a
// table is pasted into Teams' own compose box; tables arriving from email or
// other clients have a bare <table>), so filtering on it would miss many real
// tables, including large rowspan ones. Adaptive cards live in attachments,
// not the RichText body, so a body <table> is genuine tabular content.
const TABLE_SELECTOR = 'table';

// A top-level table is one not nested inside another table (nested tables are
// rendered as part of their outer cell's text, so we don't parse them twice).
const isTopLevelTable = (t: Element) => !t.parentElement || !t.parentElement.closest('table');

// Private-use sentinels used to splice tables out of the body, flatten the
// remaining text with the caller's htmlToText, then splice positions back.
// PUA chars don't occur in chat text, so the split is unambiguous.
const MARK_OPEN = '\uE000';
const MARK_CLOSE = '\uE001';
const markerFor = (i: number) => `${MARK_OPEN}${i}${MARK_CLOSE}`;
const MARKER_RE = /\uE000(\d+)\uE001/g;

/**
 * Parse one <table> element into a dense grid with spans resolved.
 *
 * `cellText` extracts a cell's inline text (mentions, links, <br>). It is
 * injected so the same structural walk serves both the API path (htmlToText)
 * and any DOM-side caller.
 */
export function parseTable(table: HTMLTableElement, cellText: (el: Element) => string): TableData {
  // querySelectorAll is recursive, so scope to THIS table's own rows/cells.
  // Without this, a nested table's <tr> would be walked as extra rows of the
  // outer grid (and its text is already in the parent cell via cellText).
  const own = (el: Element) => el.closest('table') === table;
  const trs = Array.from(table.querySelectorAll('tr')).filter(own);
  const rows: string[][] = [];
  const filled: boolean[][] = [];
  const merges: MergeRegion[] = [];

  // Header rows: explicit <thead> rows, else a leading row of all <th>.
  let headerRowCount = Array.from(table.querySelectorAll('thead tr')).filter(own).length;

  const isFilled = (r: number, c: number) => !!(filled[r] && filled[r][c]);
  const setCell = (r: number, c: number, v: string) => {
    (rows[r] ||= [])[c] = v;
    (filled[r] ||= [])[c] = true;
  };

  trs.forEach((tr, r) => {
    rows[r] ||= [];
    const cells = Array.from(tr.children).filter(
      el => el.tagName === 'TD' || el.tagName === 'TH',
    ) as HTMLTableCellElement[];
    if (!cells.length) return;
    if (headerRowCount === 0 && r === 0 && cells.every(c => c.tagName === 'TH')) {
      headerRowCount = 1;
    }
    let c = 0;
    for (const cell of cells) {
      while (isFilled(r, c)) c++; // skip positions a rowspan above already took
      // Clamp rowspan to the rows that actually exist. The rowSpan IDL property
      // reflects the raw attribute, not the browser's render-time clamp, so a
      // malformed rowspan="9" in a 3-row table would otherwise invent phantom
      // rows the user never saw.
      const rowspan = Math.min(Math.max(1, cell.rowSpan || 1), trs.length - r);
      const colspan = Math.max(1, cell.colSpan || 1);
      const value = cellText(cell).trim();
      // Forward-fill the value into every covered position so each row is dense.
      for (let dr = 0; dr < rowspan; dr++) {
        for (let dc = 0; dc < colspan; dc++) setCell(r + dr, c + dc, value);
      }
      if (rowspan > 1 || colspan > 1) merges.push({ row: r, col: c, rowspan, colspan });
      c += colspan;
    }
  });

  const columns = rows.reduce((m, row) => Math.max(m, row.length), 0);
  // Pad ragged rows so the grid is rectangular (covered/empty cells become '').
  for (let r = 0; r < rows.length; r++) {
    const row = (rows[r] ||= []);
    for (let c = 0; c < columns; c++) if (row[c] === undefined) row[c] = '';
  }
  return { columns, headerRowCount, rows, merges };
}

/**
 * Split a message-body HTML string into ordered text/table blocks.
 *
 * Returns [] when the body has no real table, so non-table messages are left
 * exactly as they were (the caller keeps using the flat `text`).
 *
 * @param html       the message body HTML (ExportMessage.contentHtml)
 * @param htmlToText flattens a fragment to text (mentions/links/emoji); reused
 *                   for both the surrounding text and each cell's content
 */
export function parseBodyBlocksFromHtml(
  html: string,
  htmlToText: (fragment: string) => string,
): BodyBlock[] {
  if (!html || !/<table/i.test(html)) return [];
  // Strip any pre-existing sentinel code points so pasted PUA glyphs (icon
  // fonts live at U+E000+) can't be mistaken for our table markers.
  const safeHtml = html.replace(/[\uE000\uE001]/g, '');
  const body = new DOMParser().parseFromString(safeHtml, 'text/html').body;
  const tables = Array.from(body.querySelectorAll<HTMLTableElement>(TABLE_SELECTOR)).filter(isTopLevelTable);
  if (!tables.length) return [];

  const cellText = (el: Element) => htmlToText((el as HTMLElement).innerHTML);
  // Parse every candidate first (cellText reads the live cell), then splice in
  // markers only for tables that produced a usable grid. A degenerate table
  // (no rows/cells) is left in place so htmlToText flattens it normally rather
  // than emitting an empty table block and losing any text it held.
  const parsed: TableData[] = [];
  const kept: HTMLTableElement[] = [];
  for (const table of tables) {
    const data = parseTable(table, cellText);
    if (data.columns > 0 && data.rows.length > 0) { parsed.push(data); kept.push(table); }
  }
  if (!parsed.length) return [];
  kept.forEach((table, i) => {
    table.replaceWith(body.ownerDocument.createTextNode(markerFor(i)));
  });

  // Flatten the now table-free body to text; markers survive as plain text.
  const flat = htmlToText(body.innerHTML);
  const blocks: BodyBlock[] = [];
  let last = 0;
  for (const m of flat.matchAll(MARKER_RE)) {
    const idx = m.index ?? 0;
    const pre = flat.slice(last, idx).trim();
    if (pre) blocks.push({ type: 'text', text: pre });
    const table = parsed[Number(m[1])];
    if (table) blocks.push({ type: 'table', table });
    last = idx + m[0].length;
  }
  const tail = flat.slice(last).trim();
  if (tail) blocks.push({ type: 'text', text: tail });

  // Only return blocks when at least one real table survived; otherwise the
  // caller should keep its existing flat-text behavior.
  return blocks.some(b => b.type === 'table') ? blocks : [];
}
