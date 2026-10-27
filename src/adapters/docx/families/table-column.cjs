'use strict';

const { norm, textUnits } = require('../inventory.cjs');
const { UNRESOLVED } = require('./text.cjs');
const { tableReadWithVerticalMerges, cellSelector, scopeOf, CELL_KEYS } = require('./table.cjs');

/**
 * DOCX `table-column` family (M4 Wave 4c): insert_table_column_before, insert_table_column_after,
 * delete_table_column -- all P1, PRIMITIVE_INTEGRAL (owner brief W4c, owner channel, 31 points).
 *
 * Source oracle (joskapista44/euro-mcp f8c6d6cf, read-only; audit: tracker
 * m4-w4c-table-column-20260924 #ref):
 *   planner     docx-agent-planner.cjs:60 (terminalColumnCount required; delete: terminalTableAnchor)
 *   target      docx-tables.cjs:23-31 (a GLOBAL cell text; the HTML projection's PHYSICAL index)
 *   lower       docx-session-tables.cjs:37-39 (frame.evaluate addColumnLeft / addColumnRight /
 *               remColumn on the outer editor -- the backdoor)
 *   L1          docx-tables.cjs:531-760 (every row's physical cell count +-1, original texts shifted)
 *   retry / L2  docx-agent-idempotency.cjs:155-162, docx-agent-whole-verify.cjs:74 (the physical
 *               width; a table whose rows differ in cell count -> fail closed)
 *
 * Measured live (test/live/docx-column-probe.live.cjs, column probe 1, 2026-09-24):
 *   - public ApiTable.AddColumn(cell, isBefore) inside callCommand puts ONE empty column directly
 *     before/after the target's column -- first, middle and last alike; RemoveColumn(cell) removes
 *     exactly it and RETURNS false even then (the fresh read decides, never the callback);
 *   - colspan (a 3-wide merge): an insert on the span's EDGE or outside it keeps the span and its
 *     content; an insert INSIDE the span put the merged row's new cell before the span, so the merge
 *     then covered other columns (semantic drift); a delete of a column the span covers LOST the
 *     merged content; a merged cell as the target of a delete removed all its columns;
 *   - the public Split does not make a ragged grid: the other rows' cells get colspan (the same rule);
 *   - a vertical merge in the TARGET table: refused (unchanged reason string); a merge in ANOTHER
 *     table no longer refuses this one (W06-3 point 2 -- per-table, not document-wide).
 *
 * Contract (owner brief W4c points 3-4, 6-9, 11-12, 19, 23-26):
 *   identity     TABLE identity + CELL identity (the W4a `cellSelector`); the target's LOGICAL column
 *                comes from the fresh HTML grid (colspan), never from its physical index. A target
 *                cell spanning several columns has no single logical column -> FAIL_CLOSED.
 *   shape        before dispatch: rows of unequal logical width, an insert boundary strictly inside
 *                any colspan, a delete of a column any colspan covers -> FAIL_CLOSED_UNSUPPORTED_SHAPE.
 *   terminal     `terminalColumnCount` (the target table's LOGICAL width after THIS operation) is
 *                carried by the operation; delete REQUIRES a table anchor outside the deleted column.
 *   task         two or more column operations in one task each carry a table anchor; the group's
 *                LAST terminal width is the final projected state (successor-aware L2, W2.1).
 *   insert       SATISFIED only when the target is present and the declared-side logical column is
 *                empty (every row's cell there empty, span 1) at the terminal width.
 *   delete       SATISFIED only when the target is absent from the anchor table at the terminal width.
 *   L1           every row: its cells in order, one new EMPTY cell at the declared boundary (insert)
 *                or the target column's cell gone (delete); rows unchanged; the logical width terminal.
 *   L3           every preserved cell (text, paragraphs, content shape, first-paragraph properties,
 *                colspan) at its mapped place; every other table unchanged; text outside unchanged.
 *                (A new cell wider than span 1 already fails the L1 terminal width.)
 */

const INTENTS = Object.freeze({
  insert_table_column_before: { action: 'add', before: true },
  insert_table_column_after: { action: 'add', before: false },
  delete_table_column: { action: 'remove', before: null },
});

function plan(step, index) {
  const kind = INTENTS[step.intent];
  if (!kind) return { ok: false, detail: `step ${index}: ${step.intent} is not a table-column intent` };
  const sel = cellSelector(step.selector);
  if (!sel || !norm(sel.text)) return { ok: false, detail: `step ${index}: a ${step.intent} selector is { text }, { tableCell: { text [, tableAnchor] } } or the combined form with the same non-empty text` };
  if (step.terminalTableAnchor !== undefined && typeof step.terminalTableAnchor !== 'string') return { ok: false, detail: `step ${index}: terminalTableAnchor must be a string` };
  const fromStep = typeof step.terminalTableAnchor === 'string' ? norm(step.terminalTableAnchor) : null;
  const fromSel = sel.tableAnchor === null ? null : norm(sel.tableAnchor);
  if (fromStep !== null && fromSel !== null && fromStep !== fromSel) return { ok: false, detail: `step ${index}: tableCell.tableAnchor and terminalTableAnchor name different cells` };
  const tableAnchor = fromSel !== null ? fromSel : fromStep;
  if (tableAnchor !== null && (!tableAnchor || tableAnchor === norm(sel.text))) return { ok: false, detail: `step ${index}: a table anchor must be a non-empty cell text other than the target cell's` };
  if (kind.action === 'remove' && tableAnchor === null) return { ok: false, detail: `step ${index}: delete_table_column requires a surviving table anchor (tableCell.tableAnchor or terminalTableAnchor)` };
  if (!Number.isInteger(step.terminalColumnCount) || step.terminalColumnCount < (kind.action === 'add' ? 2 : 1)) return { ok: false, detail: `step ${index}: ${step.intent} requires terminalColumnCount, the target table's column count after this operation` };
  return { ok: true, operation: { id: step.id || `s${index}`, intent: step.intent, selector: step.selector, targetText: norm(sel.text), tableAnchor, terminalColumnCount: step.terminalColumnCount, preservationMode: 'PRIMITIVE_INTEGRAL', ...(Array.isArray(step.dependsOn) ? { dependsOn: step.dependsOn } : {}) } };
}

/** Task level (pure), the table-row rule for columns: grouped by table anchor, the group's final width. */
function annotateTask(operations) {
  const cols = operations.filter(op => INTENTS[op.intent]);
  if (cols.length < 2) return operations;
  const ungroupable = cols.some(op => op.tableAnchor === null);
  const finalOf = new Map();
  for (const op of cols) if (op.tableAnchor !== null) finalOf.set(op.tableAnchor, op.terminalColumnCount);
  return operations.map((op, i) => {
    if (!INTENTS[op.intent]) return op;
    if (ungroupable) return { ...op, columnGroup: Object.freeze({ ungroupable: true }) };
    const later = operations.slice(i + 1).filter(b => INTENTS[b.intent] && b.tableAnchor === op.tableAnchor);
    return { ...op, columnGroup: Object.freeze({ finalColumnCount: finalOf.get(op.tableAnchor), later: Object.freeze(later.map(b => b.id)),
      laterOps: Object.freeze(later.map(b => Object.freeze({ intent: b.intent, targetText: b.targetText, tableAnchor: b.tableAnchor }))) }) };
  });
}

// ---- the observation ----------------------------------------------------------------------------

const cellKey = c => JSON.stringify([c.text, c.paragraphs, c.simple, CELL_KEYS.filter(k => c.format[k] !== undefined).map(k => [k, c.format[k]])]);

/**
 * M5 (owner decision A after owner channel, known-limits row 14): the column family's table read
 * on tables with a grid GAP (w:gridAfter). The projection adds a trailing EMPTY cell marked
 * `placeholder` that the driver read does not have (gridAfter probe, run ...fcd2df). Reconciled per
 * row from the fresh read: the placeholders are the row's trailing gap, each span 1 and empty, and the
 * projected cells before them match the driver row one to one. Anything else -> not reconciled
 * (refused). A table set without any placeholder takes the PER-TABLE reconciled path (W06-3)
 * -- `tableReadWithVerticalMerges`, the same read table-row.cjs already uses; `locate()`
 * below still refuses a vertical merge in the TARGET table specifically (unchanged outcome for
 * that table, decision point 2), but no longer for every table merely because ANOTHER one
 * in the document has a merge (LIVE-measured READ_DIVERGENCE: a merge-free table's two
 * reads agree row-for-row regardless of what another table in the same document carries). The
 * grid-gap branch below is UNCHANGED -- its own, narrower blanket is a separate condition,
 * reviewed on its own (design doc point 2).
 */
function columnRead(primary) {
  const html0 = ((primary && primary.blocks) || []).filter(b => b.kind === 'table');
  if (!html0.some(t => (t.rows || []).some(r => r.some(c => c.placeholder)))) {
    const read = tableReadWithVerticalMerges(primary);
    return read.ok ? { ...read, gaps: read.html.map(t => (t.rows || []).map(() => 0)), leads: read.html.map(t => (t.rows || []).map(() => 0)) } : read;
  }
  const d = primary && primary.document;
  const tables = d ? d.tables : null;
  if (!Array.isArray(tables)) return { ok: false, reason: 'tables-unreadable' };
  if (tables.length !== html0.length) return { ok: false, reason: 'table-read-shape-mismatch' };
  // the refusal names the failed check and what it saw (the first LIVE run refused every case with no
  // way to tell which condition fired, run ...ab094e)
  const unreconciled = (table, row, check) => ({ ok: false, reason: 'FAIL_CLOSED_UNSUPPORTED_SHAPE:grid-gap-unreconciled',
    detail: { table, row, check, projected: (html0[table].rows[row] || []).map(c => ({ placeholder: c.placeholder === true, span: c.span, text: c.placeholder ? JSON.stringify(c.text).slice(0, 24) : undefined })), driverCells: ((tables[table].rows || [])[row] || []).length } });
  const html = [];
  const gaps = [];
  const leads = [];
  for (let ti = 0; ti < html0.length; ti += 1) {
    const hrows = html0[ti].rows || [];
    if (hrows.length !== tables[ti].rows.length) return unreconciled(ti, 0, 'row-count');
    const real = [];
    const gap = [];
    const lead = [];
    // GH #291 V13 (LIVE, docs/evidence/gh291-v13-h5-probe-20261022.json): a vertical merge next to a grid gap -- the projection
    // omits a continuation cell, the driver row keeps it (as tableReadWithVerticalMerges counts it)
    const vspans = [];
    hrows.forEach((row, r) => row.forEach(c => { if (!c.placeholder && c.rowspan > 1) vspans.push([r, r + c.rowspan - 1]); }));
    if (vspans.some(([, bottom]) => bottom >= hrows.length)) return unreconciled(ti, 0, 'vertical-merge-past-the-table');
    for (let r = 0; r < hrows.length; r += 1) {
      // GH #291 H4 (LIVE, docs/evidence/gh291-h4-grid-gap-probe-20261022.json): a gap of n columns is ONE placeholder of span n
      // (gridAfter 2 -> span 2), and a gridBefore gap is a LEADING placeholder; the real cells sit between them
      const row = hrows[r];
      let a = 0;
      while (a < row.length && row[a].placeholder) a += 1;
      let z = row.length;
      while (z > a && row[z - 1].placeholder) z -= 1;
      const cells = row.slice(a, z);
      const pad = [...row.slice(0, a), ...row.slice(z)];
      // a placeholder may carry whitespace only (the projection's &nbsp; paragraph) -- a gap holds no text
      if (cells.some(c => c.placeholder) || pad.some(c => !(c.span >= 1) || c.text.replace(/\s/g, '') !== '')) return unreconciled(ti, r, 'gap-shape');
      const covered = vspans.filter(([top, bottom]) => top < r && r <= bottom).length;
      if (cells.length + covered !== tables[ti].rows[r].length) return unreconciled(ti, r, 'cell-count');
      const width = cs => cs.reduce((n, c) => n + c.span, 0);
      real.push(cells);
      lead.push(width(row.slice(0, a)));
      gap.push(width(row.slice(z)));
    }
    html.push({ ...html0[ti], rows: real });
    gaps.push(gap);
    leads.push(lead);
  }
  const cells = [];
  tables.forEach((t, ti) => t.rows.forEach((row, ri) => row.forEach((cell, ci) => cells.push({ at: `${ti}:${ri}:${ci}`, cell }))));
  return { ok: true, tables, html, cells, gaps, leads };
}

/**
 * The logical grid of one table from its HTML projection: per row, each physical cell's start and
 * colspan; a row's grid gap counts toward its width.
 *
 * W-06/2: a vertical merge's continuation row OMITS the merged column's cell from the HTML
 * projection entirely (the owner cell above carries `rowspan`, table.cjs's own read comment) --
 * `active` tracks each open merge's column position (recorded when its owning cell, `rowspan > 1`,
 * is walked) so a continuation row's REAL cells still land at their correct column. The driver's
 * physical read (table.cjs) keeps the continuation as an actual, empty cell, so the two reads
 * diverge in per-row LENGTH exactly there -- `driverIndex[r][k]` names which driver-row cell each
 * `rows[r][k]` grid cell came from (identity when no merge is open, same as before this change).
 * A merge whose owner also spans columns (GH #291 V11, measured): the driver contributes exactly ONE
 * continuation cell regardless of the colspan width, so its gap advances the grid by the span and the
 * driver index by one. A boundary inside such a block / a delete of one of its columns is refused by
 * shapeGuard (measured: the grid misaligns / the whole block is lost).
 */
function gridOf(html, gaps = [], leads = []) {
  const rows = html.rows || [];
  const active = []; // open vertical merges still owed a continuation gap: { start, until }
  const gridRows = [];
  const driverIndex = [];
  const rowWidths = [];
  for (let r = 0; r < rows.length; r += 1) {
    // GH #291 V11 (LIVE, docs/evidence/gh291-v11-vmerge-colspan-probe-20261022.json): a merged owner that also spans columns leaves ONE
    // driver continuation cell that covers its whole span -- the gap advances the grid by that span and the driver index by one
    const contGaps = active.filter(a => a.until >= r).sort((x, y) => x.start - y.start);
    let pos = leads[r] || 0; // GH #291 H4: a gridBefore gap shifts the row's real cells right
    let gi = 0;
    let di = 0;
    const rowGrid = [];
    const rowDriverIndex = [];
    for (const c of rows[r]) {
      while (gi < contGaps.length && contGaps[gi].start === pos) { pos += contGaps[gi].span; di += 1; gi += 1; }
      const span = c.span || 1;
      const rowspan = c.rowspan || 1;
      rowGrid.push({ start: pos, span });
      rowDriverIndex.push(di);
      pos += span;
      di += 1;
      if (rowspan > 1) active.push({ start: pos - span, span, until: r + rowspan - 1 });
    }
    while (gi < contGaps.length) { pos += contGaps[gi].span; di += 1; gi += 1; }
    gridRows.push(rowGrid);
    driverIndex.push(rowDriverIndex);
    rowWidths.push(pos);
  }
  const widths = new Set(rowWidths.map((w, i) => w + (gaps[i] || 0)));
  // `ends`: per row, where its real cells (and continuation gaps) end -- a short row's gap starts there
  return widths.size === 1 ? { ok: true, rows: gridRows, width: [...widths][0], driverIndex, ends: rowWidths } : { ok: false, reason: 'FAIL_CLOSED_UNSUPPORTED_SHAPE:uneven-logical-width', ends: rowWidths };
}

/** The target table, its grid and the target's logical column, from ONE fresh read. */
function locate(primary, op) {
  const read = columnRead(primary);
  if (!read.ok) return read;
  const scope = scopeOf(read, op.tableAnchor);
  if (!scope.ok) return scope;
  const hits = scope.cells.filter(c => c.cell.text === op.targetText);
  if (hits.length > 1) return { ok: false, reason: 'table-column-target-ambiguous', candidates: hits.length };
  let ti = scope.prefix;
  let ri = null;
  let ci = null;
  if (hits.length === 1) [ti, ri, ci] = hits[0].at.split(':').map(Number);
  if (ti === null) return { ok: false, reason: 'table-column-target-not-found' };
  ti = Number(ti);
  // W06-3 (point 2): the TARGET table's own vertical-merge state decides, not the document's.
  // W-06/2 (the vMerge delete follow-on to W-06/1's LIVE measurement): the REMOVE action no longer refuses it --
  // its measured post-delete behaviour (gridOf above) is now read instead; ADD stays refused
  // UNCHANGED (never measured -- column probe 1 only exercised the public delete on a vMerge
  // table). A merge in ANOTHER table still never reaches here (columnRead's per-table read).
  // GH #291 T-04 (LIVE, docs/evidence/gh291-vmerge-structure-probe-20261021.json): ADD no longer refuses either -- an insert before / after
  // the merge column or any other column keeps the vMerge on its own column and gives the new column no vMerge; L1/L3 below map
  // the continuation rows through gridOf's driverIndex. GH #291 V11: a merge whose owner also spans columns is gridded too (gridOf);
  // its inside is refused by shapeGuard.
  const table = read.tables[ti];
  let gaps = read.gaps[ti];
  let grid = gridOf(read.html[ti], gaps, read.leads[ti]);
  // GH #291 H5 (LIVE, docs/evidence/gh291-v13-h5-probe-20261022.json): a row one or more columns short WITHOUT a gap marker -- the
  // editor treats the missing columns as a trailing gap (and writes gridAfter on the next column write). Read as such when every
  // short row carries no marker of its own; `implicit` names those rows.
  let implicit = gaps.map(() => false);
  if (!grid.ok && grid.ends) {
    const full = Math.max(...grid.ends.map((e, r) => e + gaps[r]));
    implicit = grid.ends.map((e, r) => e + gaps[r] < full);
    if (implicit.some((x, r) => x && (gaps[r] > 0 || read.leads[ti][r] > 0))) return grid;
    gaps = grid.ends.map((e, r) => (implicit[r] ? full - e : gaps[r]));
    grid = gridOf(read.html[ti], gaps, read.leads[ti]);
    if (!grid.ok) return grid;
  }
  let col = null;
  if (ri !== null) {
    // `ci` is table.cjs's driver-row index (read.cells was built off read.tables, the physical
    // read); gridOf's driverIndex maps it back to the grid's own cell order -- identity when this
    // row holds no continuation gap, the same lookup a plain `grid.rows[ri][ci]` used to be.
    const k = grid.driverIndex[ri].indexOf(ci);
    if (k < 0) return { ok: false, reason: 'table-column-target-column-unresolved' };
    const g = grid.rows[ri][k];
    // GH #291 H3 / H3v (LIVE, docs/evidence/gh291-h-colspan-probe-20261022.json, gh291-h3v-colspan-rowspan-probe-20261022.json): an
    // INSERT next to a target that spans columns -- also when it spans rows too -- lands at the span's own edge in every row
    // (after: right of the span, before: left of it) -- the target's column is that edge. A DELETE on it removes EVERY column the
    // span covers (measured): refused.
    const spanned = g.span !== 1;
    if (spanned && INTENTS[op.intent].action !== 'add') return { ok: false, reason: 'FAIL_CLOSED_UNSUPPORTED_SHAPE:target-spans-columns' };
    col = spanned && !INTENTS[op.intent].before ? g.start + g.span - 1 : g.start;
    if (op.tableAnchor !== null && INTENTS[op.intent].action === 'remove' && table.rows.some((row, r) => row.some((c, dk) => {
      if (c.text !== op.tableAnchor) return false;
      const gk = grid.driverIndex[r].indexOf(dk);
      return gk >= 0 && grid.rows[r][gk].start <= col && col < grid.rows[r][gk].start + grid.rows[r][gk].span;
    }))) {
      return { ok: false, reason: 'table-anchor-in-target-column' };
    }
  }
  return { ok: true, read, ti, ri, ci, col, table, html: read.html[ti], grid, width: grid.width, gaps, leads: read.leads[ti], implicit };
}

const boundaryOf = (op, col) => (INTENTS[op.intent].before ? col : col + 1);

/** The shape guard (column probe 1): measured unsafe colspan cases, refused before dispatch. */
function shapeGuard(loc, op) {
  const cells = loc.grid.rows.flat();
  // M5 row 14 + GH #291 H4: a table with a grid gap -- only the measured primitives (gridAfter probe ...fcd2df; LIVE
  // docs/evidence/gh291-h4-grid-gap-probe-20261022.json): an insert at a boundary no further right than a short row's real cells
  // (the new cell lands at the boundary in every row, the gap stays trailing -- also a two-column gap and next to a colspan); a
  // boundary PAST a short row's real cells is refused (measured H4a / H4c: the editor puts the new cell before the gap, the grid
  // misaligns). A delete of any column: a real cell in a short row goes, a gap column shrinks that row's gap by one (H4d / H4e).
  // A gridBefore (leading) gap: an insert at or past the gap's end and a delete of a real column (H4f / H4g) -- also next to a
  // vertical merge (V13l), next to a colspan and in a row gapped on both sides (H4q; LIVE docs/evidence/gh291-gap-remaining-probe-
  // 20261023.json: aligned in every row, the leading gap kept, a trailing gap column delete consumes that gap). Measured unsafe, so
  // refused: a boundary inside or before the leading gap (the editor puts the new cell AFTER the gap, so in that row it lands in
  // another column) and a delete of a leading gap column (the gap stays, the other rows shift: the grid misaligns, a vertical merge
  // moves). GH #291 V13 / H5: a vertical merge next to a trailing gap is gridded (gridOf) -- the same rules; a row short WITHOUT a
  // marker is a trailing gap too, and deleting one of ITS missing columns removes the column from the full rows only (H5d, measured).
  const realEnd = r => loc.grid.ends[r];
  if (loc.leads.some(n => n > 0)) {
    if (INTENTS[op.intent].action === 'add') {
      if (loc.leads.some(n => boundaryOf(op, loc.col) < n)) return 'FAIL_CLOSED_UNSUPPORTED_SHAPE:insert-boundary-inside-a-leading-grid-gap';
    } else if (loc.leads.some(n => loc.col < n)) return 'FAIL_CLOSED_UNSUPPORTED_SHAPE:leading-grid-gap-column-delete';
  }
  if (loc.gaps.some(n => n > 0) && INTENTS[op.intent].action === 'add') {
    const b = boundaryOf(op, loc.col);
    if (loc.grid.rows.some((row, r) => loc.gaps[r] > 0 && b > realEnd(r))) return 'FAIL_CLOSED_UNSUPPORTED_SHAPE:insert-boundary-inside-a-grid-gap';
  }
  if (INTENTS[op.intent].action === 'add') {
    const b = boundaryOf(op, loc.col);
    if (cells.some(g => g.start < b && b < g.start + g.span)) return 'FAIL_CLOSED_UNSUPPORTED_SHAPE:colspan-straddles-insert-boundary';
  } else if (cells.some(g => g.span > 1 && g.start <= loc.col && loc.col < g.start + g.span)) {
    return 'FAIL_CLOSED_UNSUPPORTED_SHAPE:colspan-covers-deleted-column';
  }
  return null;
}

/** The local terminal claim of THIS operation, in a located state. */
function localClaim(loc, op) {
  const kind = INTENTS[op.intent];
  if (kind.action === 'remove') return loc.ri === null;
  if (loc.ri === null) return false;
  const n = kind.before ? loc.col - 1 : loc.col + 1;
  if (n < 0 || n >= loc.width) return false;
  return loc.grid.rows.every((row, r) => {
    const k = row.findIndex(g => g.start <= n && n < g.start + g.span);
    // GH #291: a continuation row's driver cells are offset from its grid cells -- map through driverIndex (identity without a merge)
    return k >= 0 && row[k].span === 1 && loc.table.rows[r][loc.grid.driverIndex[r][k]].text === '';
  });
}

// ---- classify / resolve -------------------------------------------------------------------------

function classify(primary, op) {
  if (op.columnGroup && op.columnGroup.ungroupable) return { state: 'FAIL_CLOSED', evidence: { reason: 'column-operations-without-table-anchor-cannot-be-grouped' } };
  const loc = locate(primary, op);
  if (!loc.ok) return { state: 'FAIL_CLOSED', evidence: { reason: loc.reason, candidates: loc.candidates, detail: loc.detail } };
  const kind = INTENTS[op.intent];
  const later = op.columnGroup && op.columnGroup.later.length > 0;
  const groupFinal = op.columnGroup ? op.columnGroup.finalColumnCount : op.terminalColumnCount;
  const before = kind.action === 'add' ? op.terminalColumnCount - 1 : op.terminalColumnCount + 1;
  const claim = localClaim(loc, op);
  const terminalCount = loc.width === op.terminalColumnCount || (later && loc.width === groupFinal);
  const evidence = { columns: loc.width, terminalColumnCount: op.terminalColumnCount, groupFinalColumnCount: groupFinal, targetColumn: loc.col };
  const satisfied = () => (op.recovery === UNRESOLVED
    ? { state: 'FAIL_CLOSED', evidence: { reason: 'preservation-baseline-unavailable', ...evidence } }
    : { state: 'SATISFIED', evidence: { ...evidence, proof: kind.action === 'add' ? 'fresh-live-empty-column-on-the-declared-side-and-terminal-column-count' : 'fresh-live-target-absent-in-the-anchor-table-and-terminal-column-count' } });
  if (loc.width === before && loc.ri !== null) {
    if (later && loc.width === groupFinal) {
      if (claim) return satisfied();
      const laterAll = op.columnGroup.laterOps.every(o => { const l = locate(primary, o); return l.ok && localClaim(l, o); });
      if (laterAll) return { state: 'FAIL_CLOSED', evidence: { reason: 'group-final-state-without-this-operation-claim', ...evidence } };
    }
    const unsafe = shapeGuard(loc, op);
    if (unsafe) return { state: 'FAIL_CLOSED', evidence: { reason: unsafe, ...evidence } };
    // W-06/2 (requirement 3): the deleted column's own text, per row -- including a
    // vMerge owner's -- is mechanically reported, not silently dropped. The EXECUTOR keeps
    // verifyPrimitive's evidence (not this one) for an APPLIED op, so `deletedColumnTexts` lives
    // there (same computation, same driverIndex reasoning); `targetColumnTexts` here is this
    // function's own, pre-dispatch capture of the same thing, kept for a decisive-preflight reader.
    return { state: 'APPLY', evidence: { ...evidence, targetColumnTexts: loc.table.rows.map((row, r) => {
      const k = loc.grid.rows[r].findIndex(g => g.start <= loc.col && loc.col < g.start + g.span);
      return k < 0 ? null : row[loc.grid.driverIndex[r][k]].text;
    }) } };
  }
  if (terminalCount && claim) return satisfied();
  if (kind.action === 'add' && loc.ri === null) return { state: 'FAIL_CLOSED', evidence: { reason: 'table-column-target-not-found', ...evidence } };
  if (kind.action === 'remove' && loc.ri !== null) return { state: 'FAIL_CLOSED', evidence: { reason: 'delete-target-present-but-column-count-not-terminal-plus-one', ...evidence } };
  return { state: 'FAIL_CLOSED', evidence: { reason: 'table-column-terminal-unprovable', ...evidence } };
}

function resolve(primary, selector, ctx = {}) {
  const op = ctx.operation;
  const loc = op ? locate(primary, op) : { ok: false, reason: 'no-operation' };
  if (!loc.ok || loc.ri === null) return { ok: false, target: null, selector, ambiguity: loc.ok ? 'NOT_FOUND' : loc.reason.toUpperCase().replace(/[-:]/g, '_') };
  return { ok: true, target: { kind: 'table_column', text: op.targetText, tableAnchor: op.tableAnchor, at: `${loc.ti}:${loc.ri}:${loc.ci}`, column: loc.col }, selector, ambiguity: null };
}

function preDispatchCheck(op, target) {
  return target && target.kind === 'table_column' && target.text === op.targetText && target.tableAnchor === op.tableAnchor ? { ok: true } : { ok: false, dispatched: false, detail: 'table-column-target-unresolved' };
}

// ---- lower --------------------------------------------------------------------------------------

function lower(op, target) {
  const kind = INTENTS[op.intent];
  const meta = Object.freeze({ family: 'table', action: 'column', columnAction: kind.action, before: kind.before, targetText: target.text, tableAnchor: target.tableAnchor });
  return { kind: 'callCommand', body: commandBody(meta), meta };
}

/**
 * One callCommand (column probe 1): the table by its anchor cell (or every table), the target cell
 * by its FULL normalized text, unique in scope -- never Search; then the public ApiTable.AddColumn /
 * RemoveColumn. The shape guard ran on the fresh read before dispatch. The return value is not
 * trusted (RemoveColumn answers false on success): the fresh read decides.
 */
function commandBody(meta) {
  return `
try {
  const m = ${JSON.stringify(meta)};
  const textOpts = { Numbering: false };
  const norm = s => String(s || '').replace(/\\s+/g, ' ').trim();
  const doc = Api.GetDocument();
  const cellsOf = t => { const out = []; for (let r = 0; r < t.GetRowsCount(); r += 1) { const row = t.GetRow(r); for (let c = 0; c < row.GetCellsCount(); c += 1) out.push(row.GetCell(c)); } return out; };
  const textOf = cell => norm(cell.GetContent().GetText(textOpts));
  let tables = doc.GetAllTables() || [];
  if (m.tableAnchor !== null) {
    tables = tables.filter(t => cellsOf(t).some(cell => textOf(cell) === m.tableAnchor));
    if (tables.length !== 1) return;
  }
  const hits = [];
  for (const t of tables) for (const cell of cellsOf(t)) if (textOf(cell) === m.targetText) hits.push({ t, cell });
  if (hits.length !== 1) return;
  if (m.columnAction === 'add') hits[0].t.AddColumn(hits[0].cell, m.before === true);
  else hits[0].t.RemoveColumn(hits[0].cell);
} catch (_) {}
`;
}

// ---- verify -------------------------------------------------------------------------------------

/** The expected rows after the operation: per row, indices into the BEFORE row's cells (null = the new cell). */
function expectedRows(op, loc) {
  return loc.grid.rows.map(row => {
    const idx = row.map((_, k) => k);
    if (INTENTS[op.intent].action === 'remove') return idx.filter(k => row[k].start !== loc.col);
    const b = boundaryOf(op, loc.col);
    const at = row.findIndex(g => g.start >= b);
    const p = at < 0 ? row.length : at;
    return [...idx.slice(0, p), null, ...idx.slice(p)];
  });
}

/**
 * The gaps after the operation: unchanged, except a delete of a gap column shrinks that row's gap by one (a one-column gap is
 * consumed; GH #291 H4: a two-column gap becomes one). A leading gap is kept (measured H4f / H4g; inside it is refused).
 */
function expectedGaps(op, loc) {
  const leads = loc.leads;
  if (INTENTS[op.intent].action !== 'remove') return { gaps: loc.gaps, leads };
  const realEnd = r => loc.grid.ends[r];
  return { gaps: loc.gaps.map((g, r) => (g > 0 && loc.col >= realEnd(r) ? g - 1 : g)), leads };
}

/**
 * L1: rows unchanged; per row the original texts in order with one new EMPTY cell at the boundary
 * (or the target's gone); width terminal.
 *
 * W-06/2: `seq`/`want` index BEFORE's own grid (real cells only, no vMerge continuation gap); the
 * driver's physical row can carry MORE cells than that when a continuation survives, so both sides
 * of the text comparison are built from the GRID (via `driverIndex`), never the raw driver row --
 * the same reasoning `gridOf`'s own comment gives for why a plain index is not safe past a merge.
 */
function verifyPrimitive(op, before, after) {
  const b = locate(before && before.primary, op);
  if (!b.ok || b.ri === null) return { ok: false, evidence: { reason: b.ok ? 'before-target-not-found' : `before-${b.reason}` } };
  const a = locate(after.primary, op);
  if (!a.ok) return { ok: false, evidence: { reason: `after-${a.reason}` } };
  const want = expectedRows(op, b);
  const rowsOk = a.table.rows.length === b.table.rows.length;
  const widthOk = a.width === op.terminalColumnCount;
  const gapsOk = JSON.stringify({ gaps: a.gaps, leads: a.leads }) === JSON.stringify(expectedGaps(op, b));
  const gridText = (loc, r, k) => loc.table.rows[r][loc.grid.driverIndex[r][k]].text;
  const mismatches = [];
  if (rowsOk) want.forEach((seq, r) => {
    const got = a.grid.rows[r].map((_, k) => gridText(a, r, k));
    const exp = seq.map(k => (k === null ? '' : gridText(b, r, k)));
    if (JSON.stringify(got) !== JSON.stringify(exp)) mismatches.push({ row: r, expected: exp, actual: got });
  });
  const ok = rowsOk && widthOk && gapsOk && mismatches.length === 0;
  // requirement 3 (W-06/2): a delete loses whatever text sat in the target column -- for a
  // vMerge's OWN column that includes the merged cell's content -- and the EXECUTOR keeps this
  // function's evidence, not classify()'s, for an APPLIED op (executor.cjs's `semantic` unwrap),
  // so the mechanical report of what was lost belongs here, not only in classify()'s APPLY state.
  // A continuation row that never had a real cell there (the gap) reports null, not ''.
  const deletedColumnTexts = INTENTS[op.intent].action === 'remove' ? b.table.rows.map((row, r) => {
    const k = b.grid.rows[r].findIndex(g => g.start <= b.col && b.col < g.start + g.span);
    return k < 0 ? null : row[b.grid.driverIndex[r][k]].text;
  }) : null;
  return { ok, evidence: { rows: [b.table.rows.length, a.table.rows.length], gapsBefore: b.gaps, gapsAfter: a.gaps, leadsBefore: b.leads, leadsAfter: a.leads, columnsBefore: b.width, columnsAfter: a.width, terminalColumnCount: op.terminalColumnCount, targetColumnBefore: b.col, targetColumnAfter: a.col, mismatches: mismatches.slice(0, 10), ...(deletedColumnTexts !== null ? { deletedColumnTexts } : {}) } };
}

/**
 * L3, PRIMITIVE_INTEGRAL: every preserved cell and its colspan at the mapped place; other tables
 * and outside text unchanged. W-06/2: the shape check and the per-cell driver lookups use the
 * GRID (real-cell) index via `driverIndex`, same reasoning as verifyPrimitive above -- `a.html`/
 * `b.html` are already grid-indexed (the HTML projection IS the grid's own source), so `span()`
 * needs no translation.
 */
function verifyPreservation(op, before, after) {
  const b = locate(before && before.primary, op);
  if (!b.ok || b.ri === null) return { ok: false, evidence: { reason: b.ok ? 'before-target-not-found' : `before-${b.reason}` } };
  const a = locate(after.primary, op);
  if (!a.ok) return { ok: false, evidence: { reason: `after-${a.reason}` } };
  const want = expectedRows(op, b);
  if (a.table.rows.length !== want.length || want.some((seq, r) => a.grid.rows[r].length !== seq.length)) return { ok: false, evidence: { reason: 'table-shape-not-as-expected' } };
  const span = c => [c.span, c.rowspan];
  const changed = [];
  want.forEach((seq, r) => seq.forEach((k, i) => {
    const got = a.table.rows[r][a.grid.driverIndex[r][i]];
    if (k === null) return; // the new cell: empty (L1), span 1 (the L1 terminal width)
    const cells = cellKey(got) === cellKey(b.table.rows[r][b.grid.driverIndex[r][k]]);
    const spans = JSON.stringify(span(a.html.rows[r][i])) === JSON.stringify(span(b.html.rows[r][k]));
    if (!cells || !spans) changed.push({ row: r, cell: i, from: k, cells, spans });
  }));
  const others = (r, ti) => r.tables.map((t, i) => (i === ti ? null : JSON.stringify([t.rows.map(row => row.map(cellKey)), r.html[i].rows.map(row => row.map(span)), r.gaps[i], r.leads[i]]))).filter(x => x !== null);
  const otherTables = JSON.stringify(others(a.read, a.ti)) === JSON.stringify(others(b.read, b.ti));
  const units = p => textUnits(p).map(u => norm(u.text));
  const outside = JSON.stringify(units(after.primary)) === JSON.stringify(units(before.primary));
  const gapsKept = JSON.stringify({ gaps: a.gaps, leads: a.leads }) === JSON.stringify(expectedGaps(op, b));
  const ok = changed.length === 0 && otherTables && outside && gapsKept;
  return { ok, evidence: { changedCells: changed.slice(0, 10), gridGapsKept: gapsKept, otherTablesUnchanged: otherTables, textUnitsOutsideTablesUnchanged: outside, expectationSource: 'derived-from-fresh-pre-mutation-observation' } };
}

/** L2: the group's final projected width, and this operation's local terminal claim. */
function wholeTaskChecks(op, finalObservation) {
  const loc = locate(finalObservation.primary, op);
  const groupFinal = op.columnGroup && !op.columnGroup.ungroupable ? op.columnGroup.finalColumnCount : op.terminalColumnCount;
  return [{ name: `${op.id}: ${op.intent} "${op.targetText}" -- table columns ${groupFinal} (the task's final projection) and the operation's terminal claim`, ok: loc.ok && loc.width === groupFinal && localClaim(loc, op) }];
}

// refs: none -- a column is not a text unit a later step can name
function producedText() {
  return null;
}

module.exports = { plan, annotateTask, classify, resolve, preDispatchCheck, lower, commandBody, verifyPrimitive, verifyPreservation, wholeTaskChecks, producedText, locate, gridOf, shapeGuard, columnRead, INTENTS };
