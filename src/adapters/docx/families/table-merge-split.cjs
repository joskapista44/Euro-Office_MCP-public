'use strict';

const { norm, textUnits } = require('../inventory.cjs');
const { UNRESOLVED } = require('./text.cjs');
const { cellSelector, CELL_KEYS } = require('./table.cjs');

/**
 * DOCX `table-merge-split` family (M4 Wave 4d): merge_table_cells, split_table_cell -- both P1,
 * PRIMITIVE_INTEGRAL (owner brief W4d, owner channel, 41 points).
 *
 * Source oracle (joskapista44/euro-mcp f8c6d6cf, read-only; audit: tracker
 * m4-w4d-table-merge-split-20260924 #ref):
 *   planner     docx-agent-planner.cjs:61-62 (merge: endSelector + terminalMergedText +
 *               terminalTableAnchor; split: splitRows/splitCols 1..20, not 1x1, + a surviving anchor)
 *   target      docx-tables.cjs:23-60 (a GLOBAL cell text -> the HTML's PHYSICAL row/col; merge rect
 *               from the two endpoints' physical coordinates)
 *   lower       docx-session-tables.cjs (frame.evaluate: internal SelectRange + editor MergeCells;
 *               the internal logic document for split -- the backdoor)
 *   retry / L2  docx-agent-idempotency.cjs:127-154, docx-agent-whole-verify.cjs:72-73 (the merged text
 *               once in the anchor table; split: the row count and the target row's physical width)
 *
 * Measured live (test/live/docx-merge-split-probe.live.cjs, merge/split probe 1, 2026-09-24):
 *   - public ApiTable.MergeCells(cells) inside callCommand merges 1x2, 2x1, 2x2, an edge and a middle
 *     region, next to a colspan or a rowspan elsewhere, and a region holding a WHOLE merged cell;
 *     the owner keeps every source paragraph IN ROW-MAJOR ORDER, each with its own runs and alignment;
 *     a region covering a merged cell PARTLY, an L-shape, two non-adjacent cells -> null, no change;
 *   - the projection renders the owner colspan/rowspan and OMITS a continuation cell; the public
 *     table read keeps ONE empty cell per covered row (the internal gridSpan/vMerge, read only as a
 *     control, agree) -- so the grid is reconciled from public reads alone (no internal dependency);
 *   - public ApiTable.Split(cell, rows, cols) with rows x cols == the cell's span restores 1x1 cells;
 *     ALL the content (every paragraph, formatting kept) stays in the top-left cell, the new cells are
 *     empty; a shape other than the span refines the grid of OTHER rows (1x3 on a colspan 2, a plain
 *     cell) -> refused here; 0x2 -> null; 1x1 -> no change.
 *
 * Contract (owner brief W4d points 4-6, 10-16, 18-21, 27, 32-34):
 *   identity   TABLE identity (a REQUIRED anchor: tableCell.tableAnchor or terminalTableAnchor, a cell
 *              outside the region) + CELL identity by full text in that table; the logical grid
 *              (row, column, colspan, rowspan) from the reconciled read, never a physical index.
 *   merge      selector (start) + endSelector (end) name the region's corners; the region is their
 *              logical bounding rectangle. Refused before dispatch: a merged cell partly inside it, a
 *              vertical merge inside it (not measured), fewer than two cells, the anchor inside it.
 *              The operation carries its terminal identity: terminalMergedText (the owner's text =
 *              the region's paragraphs in row-major order) and terminalMergedSpan { rows, cols }; both
 *              must equal what the fresh read derives, else FAIL_CLOSED (contract mismatch).
 *   split      the target is a merged owner; splitRows x splitCols MUST equal its span (else
 *              FAIL_CLOSED_UNSUPPORTED_SHAPE); a 1x1 target that already holds the split block ->
 *              SATISFIED, otherwise FAIL_CLOSED (not a merged cell).
 *   retry      merge SATISFIED: the endpoints absent, the terminal text once as an owner of the
 *              terminal span; split SATISFIED: the target 1x1 at the top-left of an empty 1x1 block.
 *              A later split of this merge's result in the same task (same anchor) is its SUCCESSOR:
 *              the merge is then SATISFIED by the successor's terminal state (W2.1).
 *   L1         merge: the owner at the region's top-left with the region's span, its non-empty
 *              paragraphs the region's in row-major order, nothing else starting inside; split: the
 *              target 1x1 in place with all its paragraphs, the rest of the block empty 1x1 cells.
 *   L3         every owner outside the region/block unchanged (place, spans, text, paragraphs,
 *              first-paragraph properties, per-paragraph formats); the merged/split cell's paragraph
 *              formats kept; every other table unchanged; the text units outside the tables in order.
 *
 * The W4b/W4c generic guard is NOT relaxed: `table.cjs tableRead` still refuses a vertical merge for
 * the cell, row and column families. This family's own reconciled read is used by this family only.
 */

const INTENTS = Object.freeze({ merge_table_cells: 'merge', split_table_cell: 'split' });

function anchorOf(step, sel, index) {
  if (step.terminalTableAnchor !== undefined && typeof step.terminalTableAnchor !== 'string') return { ok: false, detail: `step ${index}: terminalTableAnchor must be a string` };
  const fromStep = typeof step.terminalTableAnchor === 'string' ? norm(step.terminalTableAnchor) : null;
  const fromSel = sel.tableAnchor === null ? null : norm(sel.tableAnchor);
  if (fromStep !== null && fromSel !== null && fromStep !== fromSel) return { ok: false, detail: `step ${index}: tableCell.tableAnchor and terminalTableAnchor name different cells` };
  const a = fromSel !== null ? fromSel : fromStep;
  if (!a) return { ok: false, detail: `step ${index}: ${step.intent} requires a table anchor (tableCell.tableAnchor or terminalTableAnchor): the table identity` };
  return { ok: true, anchor: a };
}

function plan(step, index) {
  const kind = INTENTS[step.intent];
  if (!kind) return { ok: false, detail: `step ${index}: ${step.intent} is not a table merge/split intent` };
  const sel = cellSelector(step.selector);
  if (!sel || !norm(sel.text)) return { ok: false, detail: `step ${index}: a ${step.intent} selector is { text }, { tableCell: { text [, tableAnchor] } } or the combined form with the same non-empty text` };
  const a = anchorOf(step, sel, index);
  if (!a.ok) return a;
  const target = norm(sel.text);
  const dep = Array.isArray(step.dependsOn) ? { dependsOn: step.dependsOn } : {};
  if (kind === 'merge') {
    const end = cellSelector(step.endSelector);
    if (!end || !norm(end.text)) return { ok: false, detail: `step ${index}: merge_table_cells requires endSelector, the region's other corner cell` };
    if (end.tableAnchor !== null && norm(end.tableAnchor) !== a.anchor) return { ok: false, detail: `step ${index}: the endSelector names another table (cross-table merge)` };
    const endText = norm(end.text);
    const merged = typeof step.terminalMergedText === 'string' ? norm(step.terminalMergedText) : '';
    const span = step.terminalMergedSpan;
    if (!merged) return { ok: false, detail: `step ${index}: merge_table_cells requires terminalMergedText, the merged cell's text` };
    if (!span || !Number.isInteger(span.rows) || !Number.isInteger(span.cols) || span.rows < 1 || span.cols < 1 || span.rows * span.cols < 2) return { ok: false, detail: `step ${index}: merge_table_cells requires terminalMergedSpan { rows, cols } (the merged cell's span, at least two cells)` };
    if (endText === target) return { ok: false, detail: `step ${index}: the two corners must be different cells` };
    if ([target, endText, merged].includes(a.anchor)) return { ok: false, detail: `step ${index}: the table anchor must be a cell outside the region` };
    if (merged === target || merged === endText) return { ok: false, detail: `step ${index}: terminalMergedText must differ from both corner texts` };
    return { ok: true, operation: { id: step.id || `s${index}`, intent: step.intent, selector: step.selector, endSelector: step.endSelector, startText: target, endText, terminalMergedText: merged, terminalMergedSpan: Object.freeze({ rows: span.rows, cols: span.cols }), tableAnchor: a.anchor, preservationMode: 'PRIMITIVE_INTEGRAL', ...dep } };
  }
  const rows = step.splitRows;
  const cols = step.splitCols;
  if (!Number.isInteger(rows) || !Number.isInteger(cols) || rows < 1 || cols < 1 || rows > 20 || cols > 20) return { ok: false, detail: `step ${index}: split_table_cell requires integer splitRows and splitCols in 1..20` };
  if (rows === 1 && cols === 1) return { ok: false, detail: `step ${index}: a 1x1 split is degenerate` };
  if (a.anchor === target) return { ok: false, detail: `step ${index}: the table anchor must be a cell other than the target` };
  return { ok: true, operation: { id: step.id || `s${index}`, intent: step.intent, selector: step.selector, targetText: target, splitRows: rows, splitCols: cols, tableAnchor: a.anchor, preservationMode: 'PRIMITIVE_INTEGRAL', ...dep } };
}

/** Task level (pure): a split of THIS merge's result (same anchor, target == the merged text) later in the task is the merge's successor. */
function annotateTask(operations) {
  return operations.map((op, i) => {
    if (op.intent !== 'merge_table_cells') return op;
    const next = operations.slice(i + 1).find(b => b.intent === 'split_table_cell' && b.tableAnchor === op.tableAnchor && b.targetText === op.terminalMergedText);
    return next ? { ...op, successorSplit: Object.freeze({ id: next.id, rows: next.splitRows, cols: next.splitCols }) } : op;
  });
}

// ---- the observation: the reconciled grid ------------------------------------------------------

/**
 * One table's grid from the two public reads: the HTML projection places each rendered cell by its
 * colspan/rowspan and omits a continuation cell; the driver read has, per row, the rendered cells
 * PLUS one empty cell per vertical merge covering that row (probe 1). A driver row that is not
 * exactly that sequence, a continuation that is not empty, or uneven logical widths -> not ok.
 */
function gridOf(tableRead, html) {
  const covered = []; // covered[r] = Map(startCol -> owner) for owners from rows above
  const owners = [];
  const entries = [];
  const widths = new Set();
  (html.rows || []).forEach((hrow, r) => {
    const cover = covered[r] || new Map();
    const row = [];
    let col = 0;
    const flushCover = () => { let o; while ((o = cover.get(col))) { row.push({ kind: 'cont', owner: o, c: col }); col += o.cs; } };
    for (const h of hrow) {
      flushCover();
      const o = { r, c: col, rs: h.rowspan || 1, cs: h.span || 1 };
      for (let k = 1; k < o.rs; k += 1) { covered[r + k] = covered[r + k] || new Map(); covered[r + k].set(col, o); }
      owners.push(o);
      row.push({ kind: 'owner', owner: o, c: col });
      col += o.cs;
    }
    flushCover();
    widths.add(col);
    entries.push(row);
  });
  if (widths.size > 1) return { ok: false, reason: 'FAIL_CLOSED_UNSUPPORTED_SHAPE:uneven-logical-width' };
  const drows = (tableRead && tableRead.rows) || [];
  if (drows.length !== entries.length || entries.some((row, r) => drows[r].length !== row.length)) return { ok: false, reason: 'table-read-shape-mismatch' };
  for (let r = 0; r < entries.length; r += 1) {
    for (let k = 0; k < entries[r].length; k += 1) {
      const e = entries[r][k];
      const cell = drows[r][k];
      if (e.kind === 'cont') { if (cell.text !== '') return { ok: false, reason: 'table-read-continuation-not-empty' }; continue; }
      e.owner.cell = cell;
      e.owner.phys = [r, k];
    }
  }
  return { ok: true, owners, rows: entries.length, width: widths.size ? [...widths][0] : 0, gridCalc: (tableRead && tableRead.gridCalc) || null };
}

/**
 * M5 gate 8 (owner decision 8A after 28880): a column split is dispatched ONLY when its cut lines
 * provably fall on the EXISTING grid lines. Measured live (split-grid probe, 2026-09-25): the public
 * ApiTable.Split(cell, 1, 2) cuts the merged cell at HALF of its CALCULATED width -- equal spanned
 * columns (14.59|14.59) kept the grid; unequal ones (40.27|14.59 after a longer cell text; 63.5|21.17
 * set through the public SetWidth) cut a NEW grid line (width 4 -> 5). The model is measured for
 * TWO columns only; any other column count is refused (not generalized). splitCols 1 cuts no column.
 * The widths come from the internal calculated grid (EO_INTERNAL_DEPENDENCY READ); unread or not
 * matching the logical grid -> refused. No width is ever changed to make a split fit. An unequal cut is
 * an ADAPTER boundary (FAIL_CLOSED_UNSUPPORTED_SHAPE), NOT a proven EO_LIMIT (owner, #ref): the probe
 * proved only that Split(cell, 1, 2) cuts at half of the calculated width; whether another public call
 * sequence could cut on the existing line is NOT_MEASURED.
 */
const SPLIT_COLS_MEASURED = Object.freeze([1, 2]);
const GRID_EPS_MM = 0.01;
function splitGridGuard(g, o, cols) {
  if (cols === 1) return { ok: true, model: 'no-column-cut' };
  if (!SPLIT_COLS_MEASURED.includes(cols)) return { ok: false, reason: `FAIL_CLOSED_UNSUPPORTED_SHAPE:split-boundary-model-not-measured-for-${cols}-columns` };
  const gc = g.gridCalc;
  if (!Array.isArray(gc)) return { ok: false, reason: 'eo-internal-table-grid-unavailable', dependency: 'docx.table.effectiveGrid.read/v1' };
  if (gc.length !== g.width) return { ok: false, reason: 'eo-internal-table-grid-disagrees-with-logical-grid', gridColumns: gc.length, logicalWidth: g.width };
  const spanned = gc.slice(o.c, o.c + o.cs);
  const total = spanned.reduce((a, b) => a + b, 0);
  const cuts = []; const lines = [];
  for (let i = 1; i < cols; i += 1) { cuts.push((total * i) / cols); lines.push(spanned.slice(0, i).reduce((a, b) => a + b, 0)); }
  const aligned = cuts.every((x, i) => Math.abs(x - lines[i]) <= GRID_EPS_MM);
  return aligned ? { ok: true, model: 'equal-cut-on-existing-grid', spanned } : { ok: false, reason: 'FAIL_CLOSED_UNSUPPORTED_SHAPE:split-would-cut-a-new-grid-line', spanned, cuts, gridLines: lines };
}

function gridRead(primary) {
  const d = primary && primary.document;
  const tables = d ? d.tables : null;
  if (!Array.isArray(tables)) return { ok: false, reason: 'tables-unreadable' };
  const html = ((primary && primary.blocks) || []).filter(b => b.kind === 'table');
  if (html.length !== tables.length) return { ok: false, reason: 'table-read-shape-mismatch' };
  const grids = [];
  // W-B1920-C1: the M8 build-gate check this used to run per table is gone -- the
  // calculated grid now comes from the Word bridge (docx.table.effectiveGrid.read/v1), fail-closed
  // at the SOURCE (a bridge failure already leaves gridCalc null on the driver's own read).
  for (let i = 0; i < tables.length; i += 1) {
    const g = gridOf(tables[i], html[i]);
    if (!g.ok) return g;
    grids.push(g);
  }
  return { ok: true, grids };
}

/** The anchored table: the ONE table with an owner whose text is the anchor. */
function locateTable(primary, anchor) {
  const read = gridRead(primary);
  if (!read.ok) return read;
  const tis = read.grids.map((g, i) => (g.owners.some(o => o.cell.text === anchor) ? i : -1)).filter(i => i >= 0);
  if (tis.length !== 1) return { ok: false, reason: tis.length ? 'table-anchor-ambiguous' : 'table-anchor-not-found', tables: tis.length };
  return { ok: true, read, ti: tis[0], g: read.grids[tis[0]] };
}

const byText = (g, text) => g.owners.filter(o => o.cell.text === text);
const rect = o => ({ r1: o.r, c1: o.c, r2: o.r + o.rs - 1, c2: o.c + o.cs - 1 });
const inside = (o, R) => o.r >= R.r1 && o.r + o.rs - 1 <= R.r2 && o.c >= R.c1 && o.c + o.cs - 1 <= R.c2;
const touches = (o, R) => !(o.r > R.r2 || o.r + o.rs - 1 < R.r1 || o.c > R.c2 || o.c + o.cs - 1 < R.c1);
const nonEmpty = pf => (pf || []).filter(p => p.text !== '');
const pfKey = pf => JSON.stringify(nonEmpty(pf).map(p => [p.text, CELL_KEYS.filter(k => p.format[k] !== undefined).map(k => [k, p.format[k]])]));
const ownerKey = o => JSON.stringify([o.r, o.c, o.rs, o.cs, o.cell.text, o.cell.paragraphs, o.cell.simple, CELL_KEYS.filter(k => o.cell.format[k] !== undefined).map(k => [k, o.cell.format[k]]), o.cell.paraFormats && pfKey(o.cell.paraFormats)]);

/** Merge: the region from the two corners, its owners in row-major order, the guards, the derivation. */
function mergePlanOf(loc, op) {
  const s = byText(loc.g, op.startText);
  const e = byText(loc.g, op.endText);
  if (s.length !== 1 || e.length !== 1) return { ok: false, reason: 'merge-corner-not-unique', start: s.length, end: e.length };
  const a = rect(s[0]);
  const b = rect(e[0]);
  const R = { r1: Math.min(a.r1, b.r1), c1: Math.min(a.c1, b.c1), r2: Math.max(a.r2, b.r2), c2: Math.max(a.c2, b.c2) };
  const region = loc.g.owners.filter(o => touches(o, R)).sort((x, y) => x.r - y.r || x.c - y.c);
  if (region.some(o => !inside(o, R))) return { ok: false, reason: 'FAIL_CLOSED_UNSUPPORTED_SHAPE:merge-region-partially-covers-a-merged-cell' };
  if (region.some(o => o.rs > 1)) return { ok: false, reason: 'FAIL_CLOSED_UNSUPPORTED_SHAPE:merge-region-contains-a-vertical-merge' };
  if (region.length < 2) return { ok: false, reason: 'merge-region-single-cell' };
  if (region.some(o => o.cell.text === op.tableAnchor)) return { ok: false, reason: 'table-anchor-in-merge-region' };
  if (region.some(o => !Array.isArray(o.cell.paraFormats))) return { ok: false, reason: 'merge-region-paragraphs-unreadable' };
  const paras = region.flatMap(o => o.cell.paraFormats);
  const derivedText = norm(paras.map(p => p.text).join(' '));
  return { ok: true, R, span: { rows: R.r2 - R.r1 + 1, cols: R.c2 - R.c1 + 1 }, region, paras, derivedText, start: s[0], end: e[0] };
}

/** Split terminal claim: the target 1x1 at the top-left of a rows x cols block of 1x1 owners, the rest empty. */
function splitClaim(g, targetText, rows, cols) {
  const t = byText(g, targetText);
  if (t.length !== 1 || t[0].rs !== 1 || t[0].cs !== 1) return false;
  const { r, c } = t[0];
  if (r + rows > g.rows || c + cols > g.width) return false;
  for (let i = 0; i < rows; i += 1) for (let j = 0; j < cols; j += 1) {
    if (i === 0 && j === 0) continue;
    const o = g.owners.find(x => x.r === r + i && x.c === c + j);
    if (!o || o.rs !== 1 || o.cs !== 1 || o.cell.text !== '') return false;
  }
  return true;
}

// ---- classify / resolve -------------------------------------------------------------------------

function classifyMerge(primary, op) {
  const loc = locateTable(primary, op.tableAnchor);
  if (!loc.ok) return { state: 'FAIL_CLOSED', evidence: { reason: loc.reason } };
  const s = byText(loc.g, op.startText);
  const e = byText(loc.g, op.endText);
  const t = byText(loc.g, op.terminalMergedText);
  const counts = { start: s.length, end: e.length, merged: t.length };
  if (s.length > 1 || e.length > 1 || t.length > 1) return { state: 'FAIL_CLOSED', evidence: { reason: 'merge-cell-ambiguous', ...counts } };
  if (s.length === 1 && e.length === 1) {
    if (t.length) return { state: 'FAIL_CLOSED', evidence: { reason: 'merge-conflicting-old-and-new-state', ...counts } };
    const m = mergePlanOf(loc, op);
    if (!m.ok) return { state: 'FAIL_CLOSED', evidence: { reason: m.reason } };
    const want = op.terminalMergedSpan;
    if (m.derivedText !== op.terminalMergedText || m.span.rows !== want.rows || m.span.cols !== want.cols) {
      return { state: 'FAIL_CLOSED', evidence: { reason: 'merge-terminal-contract-mismatch', derivedText: m.derivedText, derivedSpan: m.span, declaredText: op.terminalMergedText, declaredSpan: want } };
    }
    return { state: 'APPLY', evidence: { region: m.R, span: m.span, cells: m.region.length } };
  }
  const satisfied = proof => (op.recovery === UNRESOLVED ? { state: 'FAIL_CLOSED', evidence: { reason: 'preservation-baseline-unavailable' } } : { state: 'SATISFIED', evidence: { proof, ...counts } });
  if (s.length === 0 && e.length === 0 && t.length === 1) {
    const o = t[0];
    if (o.rs === op.terminalMergedSpan.rows && o.cs === op.terminalMergedSpan.cols) return satisfied('fresh-live-corners-absent-merged-owner-of-the-terminal-span');
    if (op.successorSplit && splitClaim(loc.g, op.terminalMergedText, op.successorSplit.rows, op.successorSplit.cols)) return satisfied('fresh-live-successor-split-terminal-state');
  }
  return { state: 'FAIL_CLOSED', evidence: { reason: 'merge-terminal-unprovable', ...counts } };
}

function classifySplit(primary, op) {
  const loc = locateTable(primary, op.tableAnchor);
  if (!loc.ok) return { state: 'FAIL_CLOSED', evidence: { reason: loc.reason } };
  const t = byText(loc.g, op.targetText);
  if (t.length !== 1) return { state: 'FAIL_CLOSED', evidence: { reason: t.length ? 'split-target-ambiguous' : 'split-target-not-found', candidates: t.length } };
  const o = t[0];
  if (o.rs * o.cs > 1) {
    if (o.rs !== op.splitRows || o.cs !== op.splitCols) return { state: 'FAIL_CLOSED', evidence: { reason: 'FAIL_CLOSED_UNSUPPORTED_SHAPE:split-shape-is-not-the-cell-span', span: { rows: o.rs, cols: o.cs } } };
    if (!Array.isArray(o.cell.paraFormats)) return { state: 'FAIL_CLOSED', evidence: { reason: 'split-target-paragraphs-unreadable' } };
    const grid = splitGridGuard(loc.g, o, op.splitCols);
    if (!grid.ok) return { state: 'FAIL_CLOSED', evidence: { ...grid, at: [o.r, o.c], span: { rows: o.rs, cols: o.cs } } };
    return { state: 'APPLY', evidence: { at: [o.r, o.c], span: { rows: o.rs, cols: o.cs }, grid: grid.model } };
  }
  if (splitClaim(loc.g, op.targetText, op.splitRows, op.splitCols)) {
    return op.recovery === UNRESOLVED ? { state: 'FAIL_CLOSED', evidence: { reason: 'preservation-baseline-unavailable' } } : { state: 'SATISFIED', evidence: { proof: 'fresh-live-target-1x1-top-left-of-an-empty-1x1-block', at: [o.r, o.c] } };
  }
  return { state: 'FAIL_CLOSED', evidence: { reason: 'split-target-not-merged' } };
}

function classify(primary, op) {
  return INTENTS[op.intent] === 'merge' ? classifyMerge(primary, op) : classifySplit(primary, op);
}

function resolve(primary, selector, ctx = {}) {
  const op = ctx.operation;
  const loc = op ? locateTable(primary, op.tableAnchor) : { ok: false, reason: 'no-operation' };
  const up = r => r.toUpperCase().replace(/[-:]/g, '_');
  if (!loc.ok) return { ok: false, target: null, selector, ambiguity: up(loc.reason) };
  if (INTENTS[op.intent] === 'merge') {
    const m = mergePlanOf(loc, op);
    if (!m.ok) return { ok: false, target: null, selector, ambiguity: up(m.reason) };
    return { ok: true, target: { kind: 'table_merge', tableAnchor: op.tableAnchor, startText: op.startText, endText: op.endText, start: m.start.phys, end: m.end.phys, cells: m.region.map(o => o.phys), region: m.R }, selector, ambiguity: null };
  }
  const t = byText(loc.g, op.targetText);
  if (t.length !== 1) return { ok: false, target: null, selector, ambiguity: t.length ? 'AMBIGUOUS' : 'NOT_FOUND' };
  return { ok: true, target: { kind: 'table_split', tableAnchor: op.tableAnchor, targetText: op.targetText, cell: t[0].phys, rows: op.splitRows, cols: op.splitCols }, selector, ambiguity: null };
}

function preDispatchCheck(op, target) {
  const merge = INTENTS[op.intent] === 'merge';
  const ok = target && target.tableAnchor === op.tableAnchor && (merge
    ? target.kind === 'table_merge' && target.startText === op.startText && target.endText === op.endText && target.cells.length >= 2
    : target.kind === 'table_split' && target.targetText === op.targetText && target.rows === op.splitRows && target.cols === op.splitCols);
  return ok ? { ok: true } : { ok: false, dispatched: false, detail: 'table-merge-split-target-unresolved' };
}

// ---- lower --------------------------------------------------------------------------------------

/** Lossless: the merge carries both corners and every region cell; the split its cell, rows and cols. */
function lower(op, target) {
  const meta = INTENTS[op.intent] === 'merge'
    ? Object.freeze({ family: 'table', action: 'merge', tableAnchor: target.tableAnchor, startText: target.startText, endText: target.endText, start: target.start, end: target.end, cells: target.cells, region: target.region })
    : Object.freeze({ family: 'table', action: 'split', tableAnchor: target.tableAnchor, targetText: target.targetText, cell: target.cell, rows: target.rows, cols: target.cols });
  return { kind: 'callCommand', body: commandBody(meta), meta };
}

/**
 * One callCommand (probe 1): the table by its anchor cell; the physical cells by the fresh read's
 * coordinates, re-checked by their full text (a stale coordinate -> no change); then the public
 * ApiTable.MergeCells(cells) / Split(cell, rows, cols). The return value is not the authority.
 */
function commandBody(meta) {
  return `
try {
  const m = ${JSON.stringify(meta)};
  const textOpts = { Numbering: false };
  const norm = s => String(s || '').replace(/\\s+/g, ' ').trim();
  const cellsOf = t => { const out = []; for (let r = 0; r < t.GetRowsCount(); r += 1) { const row = t.GetRow(r); for (let c = 0; c < row.GetCellsCount(); c += 1) out.push(row.GetCell(c)); } return out; };
  const textOf = cell => norm(cell.GetContent().GetText(textOpts));
  const tables = (Api.GetDocument().GetAllTables() || []).filter(t => cellsOf(t).some(cell => textOf(cell) === m.tableAnchor));
  if (tables.length !== 1) return;
  const t = tables[0];
  const at = ([r, c]) => { const row = r < t.GetRowsCount() ? t.GetRow(r) : null; return row && c < row.GetCellsCount() ? row.GetCell(c) : null; };
  if (m.action === 'merge') {
    const s = at(m.start);
    const e = at(m.end);
    if (!s || !e || textOf(s) !== m.startText || textOf(e) !== m.endText) return;
    const cells = m.cells.map(at);
    if (cells.some(x => !x)) return;
    t.MergeCells(cells);
  } else {
    const cell = at(m.cell);
    if (!cell || textOf(cell) !== m.targetText) return;
    t.Split(cell, m.rows, m.cols);
  }
} catch (_) {}
`;
}

// ---- verify -------------------------------------------------------------------------------------

function otherTablesUnchanged(b, a) {
  const view = (read, ti) => read.grids.map((g, i) => (i === ti ? null : JSON.stringify(g.owners.map(ownerKey)))).filter(x => x !== null);
  return JSON.stringify(view(a.read, a.ti)) === JSON.stringify(view(b.read, b.ti));
}

/** L1 merge: the owner at the region's top-left, the region's span, its non-empty paragraphs in row-major order; nothing else starts inside. */
function verifyMerge(op, before, after) {
  const b = locateTable(before && before.primary, op.tableAnchor);
  if (!b.ok) return { ok: false, evidence: { reason: `before-${b.reason}` } };
  const m = mergePlanOf(b, op);
  if (!m.ok) return { ok: false, evidence: { reason: `before-${m.reason}` } };
  const a = locateTable(after.primary, op.tableAnchor);
  if (!a.ok) return { ok: false, evidence: { reason: `after-${a.reason}` } };
  const t = byText(a.g, op.terminalMergedText);
  const o = t.length === 1 ? t[0] : null;
  const placed = !!o && o.r === m.R.r1 && o.c === m.R.c1 && o.rs === m.span.rows && o.cs === m.span.cols;
  const content = !!o && Array.isArray(o.cell.paraFormats) && JSON.stringify(nonEmpty(o.cell.paraFormats).map(p => p.text)) === JSON.stringify(nonEmpty(m.paras).map(p => p.text));
  const alone = a.g.owners.filter(x => x !== o && touches(x, m.R)).length === 0;
  const cornersGone = byText(a.g, op.startText).length === 0 && byText(a.g, op.endText).length === 0;
  const shape = a.g.rows === b.g.rows && a.g.width === b.g.width;
  const ok = placed && content && alone && cornersGone && shape;
  return { ok, evidence: { region: m.R, span: m.span, owner: o ? { at: [o.r, o.c], span: [o.rs, o.cs] } : null, mergedFound: t.length, placed, content, alone, cornersGone, shape } };
}

/** L1 split: the target 1x1 in place with all its paragraphs; the rest of the block empty 1x1 cells. */
function verifySplit(op, before, after) {
  const b = locateTable(before && before.primary, op.tableAnchor);
  if (!b.ok) return { ok: false, evidence: { reason: `before-${b.reason}` } };
  const bt = byText(b.g, op.targetText);
  if (bt.length !== 1 || bt[0].rs * bt[0].cs < 2) return { ok: false, evidence: { reason: 'before-target-not-a-merged-cell' } };
  const a = locateTable(after.primary, op.tableAnchor);
  if (!a.ok) return { ok: false, evidence: { reason: `after-${a.reason}` } };
  const at = byText(a.g, op.targetText);
  const o = at.length === 1 ? at[0] : null;
  const inPlace = !!o && o.r === bt[0].r && o.c === bt[0].c;
  const block = splitClaim(a.g, op.targetText, op.splitRows, op.splitCols);
  const content = !!o && JSON.stringify((o.cell.paraFormats || []).map(p => p.text)) === JSON.stringify((bt[0].cell.paraFormats || []).map(p => p.text));
  const shape = a.g.rows === b.g.rows && a.g.width === b.g.width;
  const ok = inPlace && block && content && shape;
  return { ok, evidence: { at: [bt[0].r, bt[0].c], inPlace, block, content, shape } };
}

function verifyPrimitive(op, before, after) {
  return INTENTS[op.intent] === 'merge' ? verifyMerge(op, before, after) : verifySplit(op, before, after);
}

/**
 * L3, PRIMITIVE_INTEGRAL: every owner outside the region/block unchanged at its place; the merged
 * (split) cell's paragraph formats kept; other tables and the text outside the tables unchanged.
 */
function verifyPreservation(op, before, after) {
  const merge = INTENTS[op.intent] === 'merge';
  const b = locateTable(before && before.primary, op.tableAnchor);
  const a = locateTable(after.primary, op.tableAnchor);
  if (!b.ok || !a.ok) return { ok: false, evidence: { reason: !b.ok ? `before-${b.reason}` : `after-${a.reason}` } };
  let R;
  let formatsKept;
  if (merge) {
    const m = mergePlanOf(b, op);
    if (!m.ok) return { ok: false, evidence: { reason: `before-${m.reason}` } };
    R = m.R;
    const o = byText(a.g, op.terminalMergedText)[0];
    formatsKept = !!o && Array.isArray(o.cell.paraFormats) && pfKey(o.cell.paraFormats) === pfKey(m.paras);
  } else {
    const bt = byText(b.g, op.targetText)[0];
    if (!bt) return { ok: false, evidence: { reason: 'before-target-not-found' } };
    R = { r1: bt.r, c1: bt.c, r2: bt.r + op.splitRows - 1, c2: bt.c + op.splitCols - 1 };
    const o = byText(a.g, op.targetText)[0];
    formatsKept = !!o && Array.isArray(o.cell.paraFormats) && pfKey(o.cell.paraFormats) === pfKey(bt.cell.paraFormats);
  }
  const outside = g => g.owners.filter(o => !touches(o, R)).map(ownerKey);
  const cellsOutside = JSON.stringify(outside(a.g)) === JSON.stringify(outside(b.g));
  const otherTables = otherTablesUnchanged(b, a);
  const units = p => textUnits(p).map(u => norm(u.text));
  const textOutside = JSON.stringify(units(after.primary)) === JSON.stringify(units(before.primary));
  const ok = cellsOutside && formatsKept && otherTables && textOutside;
  return { ok, evidence: { cellsOutsideUnchanged: cellsOutside, paragraphFormatsKept: formatsKept, otherTablesUnchanged: otherTables, textUnitsOutsideTablesUnchanged: textOutside, expectationSource: 'derived-from-fresh-pre-mutation-observation' } };
}

/** L2: merge -- the merged owner of the terminal span, or its successor split's terminal state; split -- the block claim. */
function wholeTaskChecks(op, finalObservation) {
  const loc = locateTable(finalObservation.primary, op.tableAnchor);
  if (INTENTS[op.intent] === 'merge') {
    const t = loc.ok ? byText(loc.g, op.terminalMergedText) : [];
    const own = t.length === 1 && t[0].rs === op.terminalMergedSpan.rows && t[0].cs === op.terminalMergedSpan.cols;
    const succ = !!op.successorSplit && loc.ok && splitClaim(loc.g, op.terminalMergedText, op.successorSplit.rows, op.successorSplit.cols);
    return [{ name: `${op.id}: merge "${op.startText}".."${op.endText}" -> one owner "${op.terminalMergedText}" ${op.terminalMergedSpan.rows}x${op.terminalMergedSpan.cols}${op.successorSplit ? ' (or its successor split\'s terminal state)' : ''}`, ok: loc.ok && (own || succ) && byText(loc.g, op.startText).length === 0 && byText(loc.g, op.endText).length === 0 }];
  }
  return [{ name: `${op.id}: split "${op.targetText}" ${op.splitRows}x${op.splitCols} -> 1x1 cells, the content top-left`, ok: loc.ok && splitClaim(loc.g, op.targetText, op.splitRows, op.splitCols) }];
}

// refs: none -- a merged cell is not a text unit a later step can name
function producedText() {
  return null;
}

module.exports = { plan, annotateTask, classify, resolve, preDispatchCheck, lower, commandBody, verifyPrimitive, verifyPreservation, wholeTaskChecks, producedText, gridOf, gridRead, locateTable, mergePlanOf, splitClaim, splitGridGuard, INTENTS };
