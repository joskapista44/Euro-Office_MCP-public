'use strict';

const { norm, textUnits } = require('../inventory.cjs');
const { UNRESOLVED } = require('./text.cjs');
const { cellSelector, CELL_KEYS } = require('./table.cjs');
const { columnRead } = require('./table-column.cjs');

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
 *   split      RESTORE: the target is a merged owner; splitRows x splitCols MUST equal its span
 *              (else FAIL_CLOSED_UNSUPPORTED_SHAPE); a 1x1 target that already holds the split
 *              block -> SATISFIED. NEW SPLIT (W05, see below): the target is a never-merged 1x1
 *              cell; splitRows x splitCols is a MEASURED-membership geometry (SPLIT_COLS_MEASURED
 *              / SPLIT_ROWS_MEASURED), else FAIL_CLOSED_UNSUPPORTED_SHAPE.
 *   retry      merge SATISFIED: the endpoints absent, the terminal text once as an owner of the
 *              terminal span; split SATISFIED: the target 1x1 at the top-left of an empty 1x1 block.
 *              A later split of this merge's result in the same task (same anchor) is its SUCCESSOR:
 *              the merge is then SATISFIED by the successor's terminal state (W2.1).
 *   L1         merge: the owner at the region's top-left with the region's span, its non-empty
 *              paragraphs the region's in row-major order, nothing else starting inside; split: the
 *              target 1x1 in place with all its paragraphs, the rest of the block empty 1x1 cells
 *              (RESTORE and NEW SPLIT alike -- the same post-split placement, W05 point 3).
 *   L3         merge and a RESTORE split: every owner outside the region/block unchanged, by exact
 *              position+span (place, spans, text, paragraphs, first-paragraph properties,
 *              per-paragraph formats). A NEW split (W05-5): every owner outside the grown block
 *              mapped to its MEASURED expected after-position/span/rowspan (design doc point 6),
 *              then compared there by position, span, rowspan AND text -- the measured gridSpan/
 *              rowspan reflow is tolerated (it is the expectation), but two outside cells trading
 *              their text is now caught (W05-4's 'outside-swap' canary flipped to a permanent
 *              regression guard in docx-w4d-merge-split.test.cjs).
 *              In both cases: the merged/split cell's paragraph formats kept; every other table
 *              unchanged; the text units outside the tables in order.
 *
 * The W4b/W4c generic guard is NOT relaxed: `table.cjs tableRead` still refuses a vertical merge for
 * the cell, row and column families. This family's own reconciled read is used by this family only.
 *
 * W05 (design: docs/W05-SPLIT-GEOMETRY-DESIGN-20261007.md v2): `split_table_cell`
 * WIDENED, measured live (W05-1a/1b/2 probes, all accepted):
 *   - a RESTORE (rows x cols == the target's current span) now covers 1..3 columns and 1..2
 *     rows (SPLIT_COLS_MEASURED/SPLIT_ROWS_MEASURED below), not just 2 columns -- the measured
 *     equal-cut math already generalized, the guard was simply over-narrow.
 *   - a NEW split (rows x cols on a PLAIN, never-merged 1x1 cell) is now a supported APPLY path,
 *     not only SATISFIED/FAIL_CLOSED -- measured to grow the table's grid to fit (not refined
 *     within the existing width), content staying in the top-left new cell, every other new cell
 *     empty -- the SAME content rule as a restore (point 3 of the design doc).
 *   - WIDTH is explicitly OUT of L1/this family's contract: MEASURED (W05-2) that an explicit
 *     SetWidth on the new cells persists as `tcW` but does NOT reconcile with `tblGrid` --
 *     two conflicting sources in the saved file, worse than the engine's own auto (0) width.
 *     `lower()`/`commandBody()` therefore NEVER call SetWidth after Split; the table may come
 *     out WIDER than before a new split, and that is accepted, not corrected.
 *   - the vMerge='restart' leftover after a row-restore is an accepted residue (point 4): no
 *     public method exists to clear it (reflected, zero candidates); L3 does not assert its
 *     absence.
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
 * GH #291 H5m (LIVE, docs/evidence/gh291-gap-remaining-probe-20261023.json): a short row -- a grid gap (`gaps` / `leads`, the column
 * family's reconciled read: the projection's placeholders removed) or a row short WITHOUT a marker (an implicit trailing gap, read
 * as the column family does) -- counts its gap toward the width; `gapped` names such a table.
 */
function gridOf(tableRead, html, gaps = [], leads = []) {
  const covered = []; // covered[r] = Map(startCol -> owner) for owners from rows above
  const owners = [];
  const entries = [];
  const ends = [];
  const widths = new Set();
  (html.rows || []).forEach((hrow, r) => {
    const cover = covered[r] || new Map();
    const row = [];
    let col = leads[r] || 0;
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
    ends.push(col);
    entries.push(row);
  });
  let eff = ends.map((e, r) => gaps[r] || 0);
  let totals = ends.map((e, r) => e + eff[r]);
  if (new Set(totals).size > 1) {
    const full = Math.max(...totals);
    const implicit = totals.map(x => x < full);
    if (implicit.some((x, r) => x && ((gaps[r] || 0) > 0 || (leads[r] || 0) > 0))) return { ok: false, reason: 'FAIL_CLOSED_UNSUPPORTED_SHAPE:uneven-logical-width' };
    eff = eff.map((g, r) => (implicit[r] ? full - ends[r] : g));
    totals = totals.map(() => full);
  }
  for (const w of totals) widths.add(w);
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
  const lead = entries.map((_, r) => leads[r] || 0);
  return { ok: true, owners, rows: entries.length, width: widths.size ? [...widths][0] : 0, gridCalc: (tableRead && tableRead.gridCalc) || null,
    gaps: eff, leads: lead, gapped: eff.some(x => x > 0) || lead.some(x => x > 0) };
}

/**
 * M5 gate 8 (owner decision 8A after 28880): a column split is dispatched ONLY when its cut lines
 * provably fall on the EXISTING grid lines. Measured live (split-grid probe, 2026-09-25): the public
 * ApiTable.Split(cell, 1, 2) cuts the merged cell at HALF of its CALCULATED width -- equal spanned
 * columns (14.59|14.59) kept the grid; unequal ones (40.27|14.59 after a longer cell text; 63.5|21.17
 * set through the public SetWidth) cut a NEW grid line (width 4 -> 5). splitCols 1 cuts no column.
 * The widths come from the internal calculated grid (EO_INTERNAL_DEPENDENCY READ); unread or not
 * matching the logical grid -> refused. No width is ever changed to make a split fit. An unequal cut is
 * an ADAPTER boundary (FAIL_CLOSED_UNSUPPORTED_SHAPE), NOT a proven EO_LIMIT (owner, #ref): the probe
 * proved only that Split(cell, 1, 2) cuts at half of the calculated width; whether another public call
 * sequence could cut on the existing line is NOT_MEASURED.
 *
 * W05:
 * W05-1a measured a 1x3 restore (cell.Split(1,3)) landing cleanly on the public API -- the SAME
 * equal-cut math this guard already runs generalizes to N cuts (the `for (let i = 1; i < cols;
 * i += 1)` loop below was never N=2-specific). SPLIT_COLS_MEASURED therefore WIDENS to [1,2,3],
 * per the reviewer's own correction: "a kaput NE TOROLD, hanem a MERT halmazra TAGITSD... egy
 * torolt kapu minden nem mert alakot atenged." Any cols NOT in this set is still refused with the
 * SAME named reason -- the discipline is unchanged, only the measured set is larger.
 */
// GH #291 S2 (LIVE, docs/evidence/gh291-s-split-probe-20261022.json): a 1x4 restore and a 1x4 split of a plain cell land the
// same way (equal cut on the existing grid / the W05 grid growth) -- widened to 4. GH #291 S2x (LIVE,
// docs/evidence/gh291-split-remaining-probe-20261023.json): 1x5 / 1x6 / 2x5 restores and a 1x5 split of a plain cell the same way --
// widened to 6, the published maximum; a wider split is not offered.
const SPLIT_COLS_MEASURED = Object.freeze([1, 2, 3, 4, 5, 6]);
/**
 * W05-1a/1b measured BOTH a row restore (Split(2,1)) and a combined row+col restore (Split(2,2))
 * landing cleanly (consistent gridSpan/vMerge, no boundary ambiguity) -- rows have NO analogue to
 * the column guard's internal-gridCalc alignment check, because that check exists specifically
 * for UNEQUAL-WIDTH ambiguity; nothing measured suggests an equivalent row-height
 * ambiguity exists. This guard is therefore a SIMPLE measured-value-membership check, not a
 * gridCalc cross-check -- deliberately simpler than splitGridGuard, because nothing measured
 * justifies the extra complexity (the reviewer's own point: "javaslat: igen, a mert {1,2}-vel,
 * vagy miert nem kell" -- the "why not" answer is that rows were measured clean without it).
 */
// GH #291 S2 (LIVE, same evidence): a 3x1 restore and a 3x1 split of a plain cell land clean -- widened to 3. GH #291 S2x: a 4x1
// restore and a 4x1 split of a plain cell too -- widened to 4, the published maximum.
const SPLIT_ROWS_MEASURED = Object.freeze([1, 2, 3, 4]);
/**
 * GH #291 S3 (LIVE, same evidence): a split of a merged cell into FEWER pieces than its span -- each piece keeps an equal share of
 * the span, the grid unchanged, the content in the top-left piece: a 2x2 cell -> Split(1, 2) gives two 2x1 cells, Split(2, 1) two
 * 1x2 cells. Measured for exactly these: [span rows, span cols, split rows, split cols]. A cut between two grid lines (a 1x3 cell
 * -> Split(1, 2): a new grid line, every other row's cell in that column gets a span) stays refused (split-shape-is-not-the-cell-span).
 * GH #291 S3x (LIVE, docs/evidence/gh291-split-remaining-probe-20261023.json): the same for a 2x3 cell -> Split(1, 3) / Split(2, 1), a
 * 3x2 cell -> Split(3, 1) / Split(1, 2). A column cut is dispatched only when it lands on the existing grid lines (splitGridGuard, the
 * same M5 gate 8 check as a restore). A 1x4 cell -> Split(1, 2) lands clean too (measured), but each piece then has the request's
 * own shape (1x2): a retry cannot tell "done" from "split this piece again" -- refused (split-shape-is-not-the-cell-span).
 */
const SPLIT_PARTIAL_MEASURED = Object.freeze([[2, 2, 1, 2], [2, 2, 2, 1], [2, 3, 1, 3], [2, 3, 2, 1], [3, 2, 3, 1], [3, 2, 1, 2]]);
const partialMeasured = (rs, cs, rows, cols) => SPLIT_PARTIAL_MEASURED.some(([a, b, c, d]) => a === rs && b === cs && c === rows && d === cols);
function splitRowsGuard(rows) {
  if (rows === 1) return { ok: true, model: 'no-row-cut' };
  if (!SPLIT_ROWS_MEASURED.includes(rows)) return { ok: false, reason: `FAIL_CLOSED_UNSUPPORTED_SHAPE:split-boundary-model-not-measured-for-${rows}-rows` };
  return { ok: true, model: 'row-split-measured-clean' };
}
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
  // a piece spans o.cs / cols grid columns: a restore 1 (the cut after every column), a partial split (GH #291 S3) more
  const step = o.cs / cols;
  for (let i = 1; i < cols; i += 1) { cuts.push((total * i) / cols); lines.push(spanned.slice(0, i * step).reduce((a, b) => a + b, 0)); }
  const aligned = cuts.every((x, i) => Math.abs(x - lines[i]) <= GRID_EPS_MM);
  return aligned ? { ok: true, model: 'equal-cut-on-existing-grid', spanned } : { ok: false, reason: 'FAIL_CLOSED_UNSUPPORTED_SHAPE:split-would-cut-a-new-grid-line', spanned, cuts, gridLines: lines };
}

function gridRead(primary) {
  const d = primary && primary.document;
  const tables = d ? d.tables : null;
  if (!Array.isArray(tables)) return { ok: false, reason: 'tables-unreadable' };
  let html = ((primary && primary.blocks) || []).filter(b => b.kind === 'table');
  if (html.length !== tables.length) return { ok: false, reason: 'table-read-shape-mismatch' };
  // GH #291 H5m: a projection with grid-gap placeholders -- the column family's reconciled read (real cells, gaps, leads)
  let gaps = html.map(() => []);
  let leads = html.map(() => []);
  if (html.some(t => (t.rows || []).some(r => r.some(c => c.placeholder)))) {
    const cr = columnRead(primary);
    if (!cr.ok) return cr;
    ({ html, gaps, leads } = cr);
  }
  const grids = [];
  // W-B1920-C1: the M8 build-gate check this used to run per table is gone -- the
  // calculated grid now comes from the Word bridge (docx.table.effectiveGrid.read/v1), fail-closed
  // at the SOURCE (a bridge failure already leaves gridCalc null on the driver's own read).
  for (let i = 0; i < tables.length; i += 1) {
    const g = gridOf(tables[i], html[i], gaps[i], leads[i]);
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
  // GH #291 H5m (LIVE, docs/evidence/gh291-gap-remaining-probe-20261023.json + -r2-): on a table with a short row (a trailing gap,
  // explicit or implicit, or a gridBefore gap) a region of real cells merges as on a full table -- inside the short row, in the full
  // row over the column the short row lacks, vertically next to the gap. A region reaching into a gap is not one rectangle of
  // cells: the editor itself refuses it (MergeCells returns null, nothing changes, measured G1MX) -- refused before dispatch.
  if (region.reduce((n, o) => n + o.rs * o.cs, 0) !== (R.r2 - R.r1 + 1) * (R.c2 - R.c1 + 1)) return { ok: false, reason: 'FAIL_CLOSED_UNSUPPORTED_SHAPE:merge-region-covers-a-grid-gap' };
  // GH #291 T-07 (LIVE, docs/evidence/gh291-vmerge-structure-probe-20261021.json): a region that WHOLLY contains a vertical merge merges
  // cleanly (one owner of the region's span, the paragraphs in row-major owner order); a partial cover stays refused above
  if (region.length < 2) return { ok: false, reason: 'merge-region-single-cell' };
  if (region.some(o => o.cell.text === op.tableAnchor)) return { ok: false, reason: 'table-anchor-in-merge-region' };
  if (region.some(o => !Array.isArray(o.cell.paraFormats))) return { ok: false, reason: 'merge-region-paragraphs-unreadable' };
  const paras = region.flatMap(o => o.cell.paraFormats);
  const derivedText = norm(paras.map(p => p.text).join(' '));
  return { ok: true, R, span: { rows: R.r2 - R.r1 + 1, cols: R.c2 - R.c1 + 1 }, region, paras, derivedText, start: s[0], end: e[0] };
}

/**
 * Split terminal claim: the target at the top-left of a rows x cols block of equal pieces, the rest empty. A piece is 1x1, or --
 * GH #291 S3, a partial split of a merged cell -- the target's own span, when that is a measured partial shape.
 */
function splitClaim(g, targetText, rows, cols) {
  const t = byText(g, targetText);
  if (t.length !== 1) return false;
  const { r, c, rs: pr, cs: pc } = t[0];
  if ((pr !== 1 || pc !== 1) && !partialMeasured(pr * rows, pc * cols, rows, cols)) return false;
  if (r + rows * pr > g.rows || c + cols * pc > g.width) return false;
  for (let i = 0; i < rows; i += 1) for (let j = 0; j < cols; j += 1) {
    if (i === 0 && j === 0) continue;
    const o = g.owners.find(x => x.r === r + i * pr && x.c === c + j * pc);
    if (!o || o.rs !== pr || o.cs !== pc || o.cell.text !== '') return false;
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
  // GH #291 H5m (LIVE, gh291-gap-remaining-probe-r2-20261023.json): on a table with a short row a NEW split of a 1x1 cell grows the
  // grid the W05 way, aligned in every row, the gap kept (implicit, gridAfter, gridBefore); the split of a MERGED cell there is not
  // measured: refused
  if (loc.g.gapped && o.rs * o.cs > 1) return { state: 'FAIL_CLOSED', evidence: { reason: 'FAIL_CLOSED_UNSUPPORTED_SHAPE:split-of-a-merged-cell-on-a-table-with-a-grid-gap' } };
  if (o.rs * o.cs > 1) {
    if (o.rs !== op.splitRows || o.cs !== op.splitCols) {
      // GH #291 S3: a partial split -- its retry finds the target as the top-left piece of the measured block
      if (splitClaim(loc.g, op.targetText, op.splitRows, op.splitCols)) {
        return op.recovery === UNRESOLVED ? { state: 'FAIL_CLOSED', evidence: { reason: 'preservation-baseline-unavailable' } } : { state: 'SATISFIED', evidence: { proof: 'fresh-live-target-top-left-piece-of-an-empty-partial-split-block', at: [o.r, o.c] } };
      }
      if (!partialMeasured(o.rs, o.cs, op.splitRows, op.splitCols)) return { state: 'FAIL_CLOSED', evidence: { reason: 'FAIL_CLOSED_UNSUPPORTED_SHAPE:split-shape-is-not-the-cell-span', span: { rows: o.rs, cols: o.cs } } };
      if (!Array.isArray(o.cell.paraFormats)) return { state: 'FAIL_CLOSED', evidence: { reason: 'split-target-paragraphs-unreadable' } };
      // the column cut of a partial split is the same public Split as a restore's (half of the calculated width): on unequal
      // columns it would cut a NEW grid line -- the M5 gate 8 check, which the partial path lacked until release 27
      const pg = splitGridGuard(loc.g, o, op.splitCols);
      if (!pg.ok) return { state: 'FAIL_CLOSED', evidence: { ...pg, at: [o.r, o.c], span: { rows: o.rs, cols: o.cs } } };
      return { state: 'APPLY', evidence: { at: [o.r, o.c], span: { rows: o.rs, cols: o.cs }, partial: { rows: op.splitRows, cols: op.splitCols }, grid: pg.model } };
    }
    if (!Array.isArray(o.cell.paraFormats)) return { state: 'FAIL_CLOSED', evidence: { reason: 'split-target-paragraphs-unreadable' } };
    const rowsGuard = splitRowsGuard(op.splitRows);
    if (!rowsGuard.ok) return { state: 'FAIL_CLOSED', evidence: { ...rowsGuard, at: [o.r, o.c], span: { rows: o.rs, cols: o.cs } } };
    const grid = splitGridGuard(loc.g, o, op.splitCols);
    if (!grid.ok) return { state: 'FAIL_CLOSED', evidence: { ...grid, at: [o.r, o.c], span: { rows: o.rs, cols: o.cs } } };
    return { state: 'APPLY', evidence: { at: [o.r, o.c], span: { rows: o.rs, cols: o.cs }, grid: grid.model, rows: rowsGuard.model } };
  }
  if (splitClaim(loc.g, op.targetText, op.splitRows, op.splitCols)) {
    return op.recovery === UNRESOLVED ? { state: 'FAIL_CLOSED', evidence: { reason: 'preservation-baseline-unavailable' } } : { state: 'SATISFIED', evidence: { proof: 'fresh-live-target-1x1-top-left-of-an-empty-1x1-block', at: [o.r, o.c] } };
  }
  // W05-3: a NEW split on a never-merged (1x1) cell -- MEASURED (W05-1a/1b): the engine
  // grows the table's grid to fit the requested geometry, content stays in the top-left (new)
  // cell, every other new cell is empty (table-merge-split.cjs's own doc comment, point 3 of
  // docs/W05-SPLIT-GEOMETRY-DESIGN-20261007.md v2). Gated by the SAME measured-value sets as the
  // restore case above -- never the gridCalc alignment check, whose premise (an existing
  // sub-grid to land on) does not apply when there is no prior split/merge to restore.
  if (!Array.isArray(o.cell.paraFormats)) return { state: 'FAIL_CLOSED', evidence: { reason: 'split-target-paragraphs-unreadable' } };
  const newColsGuard = op.splitCols === 1 ? { ok: true } : (SPLIT_COLS_MEASURED.includes(op.splitCols) ? { ok: true } : { ok: false, reason: `FAIL_CLOSED_UNSUPPORTED_SHAPE:split-boundary-model-not-measured-for-${op.splitCols}-columns` });
  if (!newColsGuard.ok) return { state: 'FAIL_CLOSED', evidence: { reason: newColsGuard.reason, at: [o.r, o.c] } };
  const newRowsGuard = splitRowsGuard(op.splitRows);
  if (!newRowsGuard.ok) return { state: 'FAIL_CLOSED', evidence: { ...newRowsGuard, at: [o.r, o.c] } };
  return { state: 'APPLY', evidence: { at: [o.r, o.c], span: { rows: op.splitRows, cols: op.splitCols }, newGeometry: true } };
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

/**
 * L1 split: the target 1x1 in place with all its paragraphs; the rest of the block empty 1x1
 * cells. W05-3: `before` is no longer required to already be a merged cell -- a NEW
 * split (never merged) is a valid before-state too (`table-merge-split.cjs`'s own new APPLY
 * branch in `classifySplit`). The two cases differ in exactly one respect, MEASURED: a RESTORE
 * (undoing a prior merge) leaves the table's overall row/column count unchanged (every restore
 * measured so far did); a NEW split GROWS it (W05-1a/1b, every new-split case measured so far
 * did) -- `shape` is asserted only for a restore, since asserting it for a new split would
 * reject the split's own, expected, measured result.
 */
function verifySplit(op, before, after) {
  const b = locateTable(before && before.primary, op.tableAnchor);
  if (!b.ok) return { ok: false, evidence: { reason: `before-${b.reason}` } };
  const bt = byText(b.g, op.targetText);
  if (bt.length !== 1) return { ok: false, evidence: { reason: 'before-target-not-found' } };
  const wasMerged = bt[0].rs * bt[0].cs > 1;
  const a = locateTable(after.primary, op.tableAnchor);
  if (!a.ok) return { ok: false, evidence: { reason: `after-${a.reason}` } };
  const at = byText(a.g, op.targetText);
  const o = at.length === 1 ? at[0] : null;
  const inPlace = !!o && o.r === bt[0].r && o.c === bt[0].c;
  const block = splitClaim(a.g, op.targetText, op.splitRows, op.splitCols);
  // the target's own piece: 1x1 for a restore / a new split, the span's equal share for a partial split (GH #291 S3) -- without it a
  // block of 1x1 cells would pass for a partial split whose pieces came out the wrong size (release 27)
  const piece = !!o && (wasMerged ? o.rs * op.splitRows === bt[0].rs && o.cs * op.splitCols === bt[0].cs : o.rs === 1 && o.cs === 1);
  const content = !!o && JSON.stringify((o.cell.paraFormats || []).map(p => p.text)) === JSON.stringify((bt[0].cell.paraFormats || []).map(p => p.text));
  const shape = wasMerged ? (a.g.rows === b.g.rows && a.g.width === b.g.width) : true;
  const ok = inPlace && block && piece && content && shape;
  return { ok, evidence: { at: [bt[0].r, bt[0].c], inPlace, block, piece, content, shape, wasMerged } };
}

function verifyPrimitive(op, before, after) {
  return INTENTS[op.intent] === 'merge' ? verifyMerge(op, before, after) : verifySplit(op, before, after);
}

/**
 * L3, PRIMITIVE_INTEGRAL: every owner outside the region/block unchanged at its place; the merged
 * (split) cell's paragraph formats kept; other tables and the text outside the tables unchanged.
 *
 * W05-3: a NEW split (never merged before) is MEASURED to reflow OTHER rows' gridSpan to
 * accommodate the grown grid width (W05-1a: row 1's untouched cell gained gridSpan=2 when row 0's
 * cell split into two), AND to shift SAME-ROW neighbours to the right of the target into new
 * physical columns (the grid grows in place of the 1x1 target). A POSITIONAL "outside" check
 * (comparing owners that fall outside the SAME rectangle R in both the before- and after-read)
 * is wrong for this case on TWO counts: (1) `ownerKey` embeds span/position, which legitimately
 * shifts; (2) before the split, R's full rectangle does not yet exist -- only the 1x1 target sits
 * there -- so a neighbour that happens to sit inside R's pre-split column range (e.g. the very
 * next cell in the row) is wrongly excluded from "outside" in the BEFORE read but wrongly INCLUDED
 * after it physically shifts out of that range post-split, producing a spurious mismatch even
 * though that neighbour's own text never changed. The fix is REGION-SCOPED, not just key-scoped:
 * for a NEW split, "inside" (excluded from the outside-comparison) is the single target owner in
 * the before-read, and the full rows x cols block in the after-read; "outside" is then compared as
 * a TEXT-ONLY, ORDER-INDEPENDENT multiset (sorted), which tolerates both the span reflow and the
 * column shift while still catching genuine corruption (a cell's text changing or disappearing).
 * For a RESTORE (merge/un-merge), the exact positional `ownerKey` comparison against the SAME
 * rectangle R is unchanged and correct in both reads, since the cell already occupied that exact
 * span before the split (no such reflow was ever measured there).
 */
function verifyPreservation(op, before, after) {
  const merge = INTENTS[op.intent] === 'merge';
  const b = locateTable(before && before.primary, op.tableAnchor);
  const a = locateTable(after.primary, op.tableAnchor);
  if (!b.ok || !a.ok) return { ok: false, evidence: { reason: !b.ok ? `before-${b.reason}` : `after-${a.reason}` } };
  let formatsKept;
  let wasMerged = true;
  let cellsOutside;
  if (merge) {
    const m = mergePlanOf(b, op);
    if (!m.ok) return { ok: false, evidence: { reason: `before-${m.reason}` } };
    const R = m.R;
    const o = byText(a.g, op.terminalMergedText)[0];
    formatsKept = !!o && Array.isArray(o.cell.paraFormats) && pfKey(o.cell.paraFormats) === pfKey(m.paras);
    const outside = g => g.owners.filter(x => !touches(x, R)).map(ownerKey);
    cellsOutside = JSON.stringify(outside(a.g)) === JSON.stringify(outside(b.g));
  } else {
    const bt = byText(b.g, op.targetText)[0];
    if (!bt) return { ok: false, evidence: { reason: 'before-target-not-found' } };
    wasMerged = bt.rs * bt.cs > 1;
    const o = byText(a.g, op.targetText)[0];
    formatsKept = !!o && Array.isArray(o.cell.paraFormats) && pfKey(o.cell.paraFormats) === pfKey(bt.cell.paraFormats);
    if (wasMerged) {
      // the merged cell's own span (a restore: rows x cols; GH #291 S3, a partial split: the span the pieces share)
      const R = { r1: bt.r, c1: bt.c, r2: bt.r + bt.rs - 1, c2: bt.c + bt.cs - 1 };
      const outside = g => g.owners.filter(x => !touches(x, R)).map(ownerKey);
      cellsOutside = JSON.stringify(outside(a.g)) === JSON.stringify(outside(b.g));
    } else {
      // W05-5 (the reviewer's dispatch): a POSITION-RESPECTING comparison using the MEASURED
      // offset, replacing the W05-3/W05-4 text-only multiset (which could not tell two outside
      // cells apart if they swapped text -- the named weakness, kept as a canary test). Every
      // BEFORE owner outside the target maps to an EXPECTED after-position/span/rowspan, derived
      // from the measured reflow (point 0/6 of the design doc):
      //   same row as the target, LEFT of it   -> unchanged (r, c, cs, rs)
      //   same row as the target, RIGHT of it  -> c shifts by (splitCols-1); if splitRows>1 its
      //                                            rs also grows by (splitRows-1) (same-row
      //                                            neighbours get the rowspan, either side)
      //   a DIFFERENT row, aligned to the target's OWN column (c===bt.c) -> r,c unchanged, cs
      //                                            grows by (splitCols-1) (the measured gridSpan
      //                                            reflow); rs unchanged
      //   a DIFFERENT row, PAST that column (c>bt.c) -> c shifts by (splitCols-1); r shifts by
      //                                            (splitRows-1) if that row is BELOW the
      //                                            target's own row (new rows inserted there)
      //   a DIFFERENT row, BEFORE that column (c<bt.c) -> c unchanged; r shifts by (splitRows-1)
      //                                            if below the target's own row
      // The mapped position/span/rowspan AND the text must both match -- a swap at two
      // "expected identical" positions now fails (the canary flips), and the measured gridSpan/
      // rowspan growth is still tolerated (it is PART of the expectation, not ignored).
      const dRow = o => (o.r > bt.r ? op.splitRows - 1 : 0);
      const expected = o => {
        if (o.r === bt.r) {
          const c = o.c > bt.c ? o.c + (op.splitCols - 1) : o.c;
          const rs = o.rs + (op.splitRows > 1 ? op.splitRows - 1 : 0);
          return { r: o.r, c, rs, cs: o.cs };
        }
        if (o.c === bt.c) return { r: o.r + dRow(o), c: o.c, rs: o.rs, cs: o.cs + (op.splitCols > 1 ? op.splitCols - 1 : 0) };
        const c = o.c > bt.c ? o.c + (op.splitCols - 1) : o.c;
        return { r: o.r + dRow(o), c, rs: o.rs, cs: o.cs };
      };
      const before = b.g.owners.filter(x => x !== bt);
      const after = new Map(a.g.owners.map(x => [`${x.r},${x.c}`, x]));
      // the block itself replaces the 1x1 target with splitRows*splitCols new 1x1 owners -- the
      // cardinality check accounts for that, else an extra/missing OUTSIDE owner would be masked.
      cellsOutside = before.length === a.g.owners.length - op.splitRows * op.splitCols && before.every(o => {
        const exp = expected(o);
        const found = after.get(`${exp.r},${exp.c}`);
        return !!found && found.rs === exp.rs && found.cs === exp.cs && found.cell.text === o.cell.text;
      });
    }
  }
  const otherTables = otherTablesUnchanged(b, a);
  const units = p => textUnits(p).map(u => norm(u.text));
  const textOutside = JSON.stringify(units(after.primary)) === JSON.stringify(units(before.primary));
  const ok = cellsOutside && formatsKept && otherTables && textOutside;
  return { ok, evidence: { cellsOutsideUnchanged: cellsOutside, paragraphFormatsKept: formatsKept, otherTablesUnchanged: otherTables, textUnitsOutsideTablesUnchanged: textOutside, expectationSource: 'derived-from-fresh-pre-mutation-observation', wasMerged } };
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

module.exports = { plan, annotateTask, classify, resolve, preDispatchCheck, lower, commandBody, verifyPrimitive, verifyPreservation, wholeTaskChecks, producedText, gridOf, gridRead, locateTable, mergePlanOf, splitClaim, splitGridGuard, splitRowsGuard, SPLIT_COLS_MEASURED, SPLIT_ROWS_MEASURED, INTENTS };
