'use strict';

const { norm, textUnits } = require('../inventory.cjs');
const { UNRESOLVED } = require('./text.cjs');
const { tableReadWithVerticalMerges, cellSelector, scopeOf, CELL_KEYS } = require('./table.cjs');

/**
 * DOCX `table-row` family (M4 Wave 4b): insert_table_row_before, insert_table_row_after,
 * delete_table_row -- all P1, PRIMITIVE_INTEGRAL (owner brief W4b, owner channel, 29 points).
 *
 * Source oracle (joskapista44/euro-mcp f8c6d6cf, read-only; audit: tracker f95228ca #ref):
 *   planner     docx-agent-planner.cjs:59 (terminalRowCount required; delete: terminalTableAnchor)
 *   target      docx-agent-preflight.cjs:66-73 (a GLOBAL cell text across every table)
 *   lower       docx-session-tables.cjs nativeMutation (callCommand Search + Select, then
 *               frame.evaluate selectCell + addRowAbove / addRowBelow / remRow -- the backdoor)
 *   L1          docx-tables.cjs:214-525 (rows +-1, the inserted row's cell count, every original row
 *               at its shifted place, the deleted target's text absent)
 *   retry       docx-agent-idempotency.cjs:118-125 (row count == terminalRowCount; delete: the
 *               surviving terminalTableAnchor table)
 *   L2          docx-agent-whole-verify.cjs:71
 *
 * Measured live (test/live/docx-row-probe*.live.cjs, 2026-09-24):
 *   - public ApiTable.AddRow(cell, isBefore) inside callCommand puts ONE row directly before/after
 *     the target row -- first, middle and last row alike; the new row has the target row's cell
 *     count, empty cells (one empty paragraph). RemoveRow(cell) removes exactly the target row, and
 *     RETURNS false even then: the callback value is never evidence, the fresh read is (probe 3);
 *   - a ragged table (public Split) and a colspan row keep their shape; the new row takes the
 *     target row's shape (probe 3);
 *   - a VERTICAL merge: deleting the row that carries its top removed the merged content (probe 3);
 *     an insert inside the span EXTENDED the merge; the projection renders rowspan and omits the
 *     continuation cell (probe 4) -> a table with a vertical merge was refused before dispatch.
 *     M5 (owner decision after Talk 28963, #ref; the SOURCE inserts a row outside a vertical merge
 *     correctly, run m5-srcdiff ...dce796): on such a table a row INSERT is supported when the fresh
 *     reconciled grid (table.cjs tableReadWithVerticalMerges) shows that the new row's boundary cuts NO
 *     vertical merge in ANY column AND the target row is outside every merge (a row copied from inside
 *     or at the edge of a merge could extend it -- NOT_MEASURED, so refused); an unreconciled grid ->
 *     refused. A DELETE on such a table stays refused (FAIL_CLOSED_UNSUPPORTED_SHAPE:vertical-merge);
 *   - an inserted row does NOT inherit the target row's text properties (bold 14 pt target ->
 *     plain 11 pt new row, probe 4): the contract claims the new row's SHAPE and emptiness only.
 *
 * Contract (owner brief W4b points 2, 4-9, 17-19):
 *   identity     TABLE identity + CELL identity (the W4a `cellSelector`: { tableCell: { text,
 *                tableAnchor } }, { text }, or the source's combined form). The source's
 *                `terminalTableAnchor` is accepted as the table anchor.
 *   terminal     `terminalRowCount` (the target table's row count after THIS operation) is carried
 *                by the operation, as the source requires; delete REQUIRES a table anchor (the
 *                surviving identity: it may not sit in the deleted row).
 *   task         two or more row operations in one task must each carry a table anchor; the
 *                operations with the same anchor are one GROUP, and the group's LAST terminal row
 *                count is the final projected state (successor-aware L2, W2.1 principle).
 *   insert       SATISFIED only when the target is present, the declared-side neighbour is an
 *                empty row of the target row's shape, and the row count is the operation's (or,
 *                with later group operations, the group's final) terminal count.
 *   delete       SATISFIED only when the target is absent from the anchor table and its row count
 *                is terminal -- never from a receipt. A count that is both this operation's start
 *                and the group's final is decided by the claims: this operation's holds ->
 *                SATISFIED; every LATER group operation's holds but not this one's -> FAIL_CLOSED;
 *                else it is the start -> APPLY.
 *   L1           rows +-1; the new row directly on the declared side, empty, of the target shape;
 *                every original row's texts in ORDER at its shifted place; delete: the target gone.
 *   L3           every preserved row's cells (text, paragraphs, content shape, first-paragraph
 *                properties) and HTML spans; every other table unchanged; the text units outside
 *                the tables unchanged in order.
 */

const INTENTS = Object.freeze({
  insert_table_row_before: { action: 'add', before: true },
  insert_table_row_after: { action: 'add', before: false },
  delete_table_row: { action: 'remove', before: null },
});

function plan(step, index) {
  const kind = INTENTS[step.intent];
  if (!kind) return { ok: false, detail: `step ${index}: ${step.intent} is not a table-row intent` };
  const sel = cellSelector(step.selector);
  if (!sel || !norm(sel.text)) return { ok: false, detail: `step ${index}: a ${step.intent} selector is { text }, { tableCell: { text [, tableAnchor] } } or the combined form with the same non-empty text` };
  if (step.terminalTableAnchor !== undefined && typeof step.terminalTableAnchor !== 'string') return { ok: false, detail: `step ${index}: terminalTableAnchor must be a string` };
  const fromStep = typeof step.terminalTableAnchor === 'string' ? norm(step.terminalTableAnchor) : null;
  const fromSel = sel.tableAnchor === null ? null : norm(sel.tableAnchor);
  if (fromStep !== null && fromSel !== null && fromStep !== fromSel) return { ok: false, detail: `step ${index}: tableCell.tableAnchor and terminalTableAnchor name different cells` };
  const tableAnchor = fromSel !== null ? fromSel : fromStep;
  if (tableAnchor !== null && (!tableAnchor || tableAnchor === norm(sel.text))) return { ok: false, detail: `step ${index}: a table anchor must be a non-empty cell text other than the target cell's` };
  if (kind.action === 'remove' && tableAnchor === null) return { ok: false, detail: `step ${index}: delete_table_row requires a surviving table anchor (tableCell.tableAnchor or terminalTableAnchor)` };
  if (!Number.isInteger(step.terminalRowCount) || step.terminalRowCount < (kind.action === 'add' ? 2 : 1)) return { ok: false, detail: `step ${index}: ${step.intent} requires terminalRowCount, the target table's row count after this operation` };
  return { ok: true, operation: { id: step.id || `s${index}`, intent: step.intent, selector: step.selector, targetText: norm(sel.text), tableAnchor, terminalRowCount: step.terminalRowCount, preservationMode: 'PRIMITIVE_INTEGRAL', ...(Array.isArray(step.dependsOn) ? { dependsOn: step.dependsOn } : {}) } };
}

/**
 * Task level (pure): row operations of one task are grouped by their table anchor; each operation
 * learns the group's final terminal row count and whether later group operations follow it. Two or
 * more row operations with an anchor-less one among them cannot be grouped provably -> the task is
 * marked, and every row operation of it fails closed at classification.
 */
function annotateTask(operations) {
  const rows = operations.filter(op => INTENTS[op.intent]);
  if (rows.length < 2) return operations;
  const ungroupable = rows.some(op => op.tableAnchor === null);
  const finalOf = new Map();
  for (const op of rows) if (op.tableAnchor !== null) finalOf.set(op.tableAnchor, op.terminalRowCount);
  return operations.map((op, i) => {
    if (!INTENTS[op.intent]) return op;
    if (ungroupable) return { ...op, rowGroup: Object.freeze({ ungroupable: true }) };
    const later = operations.slice(i + 1).filter(b => INTENTS[b.intent] && b.tableAnchor === op.tableAnchor);
    return { ...op, rowGroup: Object.freeze({ finalRowCount: finalOf.get(op.tableAnchor), later: Object.freeze(later.map(b => b.id)),
      laterOps: Object.freeze(later.map(b => Object.freeze({ intent: b.intent, targetText: b.targetText, tableAnchor: b.tableAnchor }))) }) };
  });
}

// ---- the observation ----------------------------------------------------------------------------

const cellKey = c => JSON.stringify([c.text, c.paragraphs, c.simple, CELL_KEYS.filter(k => c.format[k] !== undefined).map(k => [k, c.format[k]])]);
const rowTexts = row => row.map(c => c.text);
const isEmptyRow = row => row.every(c => c.text === '');

/**
 * The target table and row, from ONE fresh read: the scope (anchor table, or every table), the
 * target cell unique in it. With an anchor, the table is found even when the target is gone.
 */
function locate(primary, op) {
  const read = tableReadWithVerticalMerges(primary);
  if (!read.ok) return read;
  const scope = scopeOf(read, op.tableAnchor);
  if (!scope.ok) return scope;
  const hits = scope.cells.filter(c => c.cell.text === op.targetText);
  if (hits.length > 1) return { ok: false, reason: 'table-row-target-ambiguous', candidates: hits.length };
  let ti = scope.prefix;
  let ri = null;
  if (hits.length === 1) { const [t, r] = hits[0].at.split(':'); ti = t; ri = Number(r); }
  if (ti === null) return { ok: false, reason: 'table-row-target-not-found' };
  const table = read.tables[Number(ti)];
  const html = read.html[Number(ti)];
  // a DELETE of the anchor's own row would remove the anchor it is proven by; an insert only adds an
  // empty row beside it, the anchor stays unique (M5, owner decision on Talk 28972: the source does it)
  if (INTENTS[op.intent].action === 'remove' && op.tableAnchor !== null && ri !== null && table.rows[ri].some(c => c.text === op.tableAnchor)) return { ok: false, reason: 'table-anchor-in-target-row' };
  const vguard = verticalMergeGuard(read.vspans[Number(ti)], op, ri);
  if (vguard) return { ok: false, reason: vguard };
  return { ok: true, read, ti: Number(ti), ri, table, html, rows: table.rows.length };
}

/**
 * M5: a row operation on a table with vertical merges ([top, bottom] rows). Delete: refused. Insert
 * before row R cuts a merge covering R-1 and R; after R one covering R and R+1 -- every merge, i.e.
 * every column, is checked. A target row inside any merge is refused too (the copied row could extend
 * it). The target row may be absent (a delete's terminal state): then only the delete rule applies.
 */
function verticalMergeGuard(spans, op, ri) {
  if (!spans || spans.length === 0) return null;
  const kind = INTENTS[op.intent];
  if (kind.action === 'remove') return 'FAIL_CLOSED_UNSUPPORTED_SHAPE:vertical-merge';
  if (ri === null) return null;
  const [a, b] = kind.before ? [ri - 1, ri] : [ri, ri + 1];
  if (spans.some(([top, bottom]) => top <= a && bottom >= b)) return 'FAIL_CLOSED_UNSUPPORTED_SHAPE:row-insert-cuts-a-vertical-merge';
  if (spans.some(([top, bottom]) => top <= ri && ri <= bottom)) return 'FAIL_CLOSED_UNSUPPORTED_SHAPE:row-insert-target-row-inside-a-vertical-merge';
  return null;
}

/** The local terminal claim of THIS operation, in a located state. */
function localClaim(loc, op) {
  const kind = INTENTS[op.intent];
  if (kind.action === 'remove') return loc.ri === null;
  if (loc.ri === null) return false;
  const n = loc.table.rows[loc.ri].length;
  const nb = loc.table.rows[kind.before ? loc.ri - 1 : loc.ri + 1];
  return !!nb && nb.length === n && isEmptyRow(nb);
}

// ---- classify / resolve -------------------------------------------------------------------------

function classify(primary, op) {
  if (op.rowGroup && op.rowGroup.ungroupable) return { state: 'FAIL_CLOSED', evidence: { reason: 'row-operations-without-table-anchor-cannot-be-grouped' } };
  const loc = locate(primary, op);
  if (!loc.ok) return { state: 'FAIL_CLOSED', evidence: { reason: loc.reason, candidates: loc.candidates } };
  const kind = INTENTS[op.intent];
  const later = op.rowGroup && op.rowGroup.later.length > 0;
  const groupFinal = op.rowGroup ? op.rowGroup.finalRowCount : op.terminalRowCount;
  const before = kind.action === 'add' ? op.terminalRowCount - 1 : op.terminalRowCount + 1;
  const claim = localClaim(loc, op);
  const terminalCount = loc.rows === op.terminalRowCount || (later && loc.rows === groupFinal);
  const evidence = { rows: loc.rows, terminalRowCount: op.terminalRowCount, groupFinalRowCount: groupFinal, targetRow: loc.ri };
  const satisfied = () => (op.recovery === UNRESOLVED
    ? { state: 'FAIL_CLOSED', evidence: { reason: 'preservation-baseline-unavailable', ...evidence } }
    : { state: 'SATISFIED', evidence: { ...evidence, proof: kind.action === 'add' ? 'fresh-live-empty-row-on-the-declared-side-and-terminal-row-count' : 'fresh-live-target-absent-in-the-anchor-table-and-terminal-row-count' } });
  const applicable = kind.action === 'add' ? loc.ri !== null : loc.ri !== null;
  if (loc.rows === before && applicable) {
    // the count alone cannot tell this operation's START from the group's FINAL state: the later
    // group operations' own claims decide -- all hold -> the task is at its final state (then this
    // operation's claim must hold too); any fails -> this is the start, APPLY
    if (later && loc.rows === groupFinal) {
      if (claim) return satisfied();
      const laterAll = op.rowGroup.laterOps.every(o => { const l = locate(primary, o); return l.ok && localClaim(l, o); });
      if (laterAll) return { state: 'FAIL_CLOSED', evidence: { reason: 'group-final-state-without-this-operation-claim', ...evidence } };
    }
    return { state: 'APPLY', evidence: { ...evidence, targetRowTexts: rowTexts(loc.table.rows[loc.ri]) } };
  }
  if (terminalCount && claim) return satisfied();
  if (kind.action === 'add' && loc.ri === null) return { state: 'FAIL_CLOSED', evidence: { reason: 'table-row-target-not-found', ...evidence } };
  if (kind.action === 'remove' && loc.ri !== null) return { state: 'FAIL_CLOSED', evidence: { reason: 'delete-target-present-but-row-count-not-terminal-plus-one', ...evidence } };
  return { state: 'FAIL_CLOSED', evidence: { reason: 'table-row-terminal-unprovable', ...evidence } };
}

function resolve(primary, selector, ctx = {}) {
  const op = ctx.operation;
  const loc = op ? locate(primary, op) : { ok: false, reason: 'no-operation' };
  if (!loc.ok || loc.ri === null) return { ok: false, target: null, selector, ambiguity: loc.ok ? 'NOT_FOUND' : loc.reason.toUpperCase().replace(/[-:]/g, '_') };
  return { ok: true, target: { kind: 'table_row', text: op.targetText, tableAnchor: op.tableAnchor, at: `${loc.ti}:${loc.ri}` }, selector, ambiguity: null };
}

function preDispatchCheck(op, target) {
  return target && target.kind === 'table_row' && target.text === op.targetText && target.tableAnchor === op.tableAnchor ? { ok: true } : { ok: false, dispatched: false, detail: 'table-row-target-unresolved' };
}

// ---- lower --------------------------------------------------------------------------------------

function lower(op, target) {
  const kind = INTENTS[op.intent];
  const meta = Object.freeze({ family: 'table', action: 'row', rowAction: kind.action, before: kind.before, targetText: target.text, tableAnchor: target.tableAnchor });
  return { kind: 'callCommand', body: commandBody(meta), meta };
}

/**
 * One callCommand (row probes 3-4): the table by its anchor cell (or every table), the target cell
 * by its FULL normalized text, unique in scope -- never Search; then the public ApiTable.AddRow /
 * RemoveRow. Any guard miss returns unchanged. The return value is not trusted (RemoveRow answers
 * false on success, probe 3): the fresh read decides.
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
  if (m.rowAction === 'add') hits[0].t.AddRow(hits[0].cell, m.before === true);
  else hits[0].t.RemoveRow(hits[0].cell);
} catch (_) {}
`;
}

// ---- verify -------------------------------------------------------------------------------------

/** The expected row sequence after the operation, as indices into the BEFORE rows (null = the new row). */
function expectedSequence(op, beforeRows, ri) {
  const idx = beforeRows.map((_, i) => i);
  const kind = INTENTS[op.intent];
  if (kind.action === 'remove') return idx.filter(i => i !== ri);
  const at = kind.before ? ri : ri + 1;
  return [...idx.slice(0, at), null, ...idx.slice(at)];
}

/** L1: rows +-1; the new row on the declared side (empty, target shape); the original rows' texts in order; delete: target gone. */
function verifyPrimitive(op, before, after) {
  const b = locate(before && before.primary, op);
  if (!b.ok || b.ri === null) return { ok: false, evidence: { reason: b.ok ? 'before-target-not-found' : `before-${b.reason}` } };
  const a = locate(after.primary, op);
  if (!a.ok) return { ok: false, evidence: { reason: `after-${a.reason}` } };
  const seq = expectedSequence(op, b.table.rows, b.ri);
  const rowsOk = a.rows === seq.length && a.rows === op.terminalRowCount;
  const mismatches = [];
  if (rowsOk) seq.forEach((src, i) => {
    const got = a.table.rows[i];
    if (src === null) { if (!(got.length === b.table.rows[b.ri].length && isEmptyRow(got))) mismatches.push({ row: i, expected: 'new empty row of the target shape', actual: rowTexts(got) }); }
    else if (JSON.stringify(rowTexts(got)) !== JSON.stringify(rowTexts(b.table.rows[src]))) mismatches.push({ row: i, expected: rowTexts(b.table.rows[src]), actual: rowTexts(got) });
  });
  // (the target's presence / absence is IMPLIED: every row sits at its place with its own texts, and
  // the target text was unique in the table before -- sabotaged separately, no state reaches it)
  const ok = rowsOk && mismatches.length === 0;
  return { ok, evidence: { rowsBefore: b.rows, rowsAfter: a.rows, terminalRowCount: op.terminalRowCount, targetRowBefore: b.ri, targetRowAfter: a.ri, mismatches: mismatches.slice(0, 10) } };
}

/**
 * L3, PRIMITIVE_INTEGRAL: every preserved row's cells unchanged (text, paragraphs, content shape,
 * first-paragraph properties) and their HTML spans; every OTHER table unchanged (cells and spans);
 * the text units outside the tables unchanged in order.
 */
function verifyPreservation(op, before, after) {
  const b = locate(before && before.primary, op);
  if (!b.ok || b.ri === null) return { ok: false, evidence: { reason: b.ok ? 'before-target-not-found' : `before-${b.reason}` } };
  const a = locate(after.primary, op);
  if (!a.ok) return { ok: false, evidence: { reason: `after-${a.reason}` } };
  const seq = expectedSequence(op, b.table.rows, b.ri);
  if (a.rows !== seq.length || a.html.rows.length !== seq.length) return { ok: false, evidence: { reason: 'row-count-not-as-expected' } };
  const span = c => [c.span, c.rowspan];
  const changedRows = [];
  seq.forEach((src, i) => {
    if (src === null) return;
    const cells = JSON.stringify(a.table.rows[i].map(cellKey)) === JSON.stringify(b.table.rows[src].map(cellKey));
    const spans = JSON.stringify(a.html.rows[i].map(span)) === JSON.stringify(b.html.rows[src].map(span));
    if (!cells || !spans) changedRows.push({ row: i, from: src, cells, spans });
  });
  const others = r => r.tables.map((t, ti) => (ti === (r === b.read ? b.ti : a.ti) ? null : JSON.stringify([t.rows.map(row => row.map(cellKey)), r.html[ti].rows.map(row => row.map(span))]))).filter(x => x !== null);
  const otherTables = JSON.stringify(others(a.read)) === JSON.stringify(others(b.read));
  const units = p => textUnits(p).map(u => norm(u.text));
  const outside = JSON.stringify(units(after.primary)) === JSON.stringify(units(before.primary));
  const ok = changedRows.length === 0 && otherTables && outside;
  return { ok, evidence: { changedRows: changedRows.slice(0, 10), otherTablesUnchanged: otherTables, textUnitsOutsideTablesUnchanged: outside, expectationSource: 'derived-from-fresh-pre-mutation-observation' } };
}

/** L2: the group's final projected row count, and this operation's local terminal claim. */
function wholeTaskChecks(op, finalObservation) {
  const loc = locate(finalObservation.primary, op);
  const groupFinal = op.rowGroup && !op.rowGroup.ungroupable ? op.rowGroup.finalRowCount : op.terminalRowCount;
  return [{ name: `${op.id}: ${op.intent} "${op.targetText}" -- table rows ${groupFinal} (the task's final projection) and the operation's terminal claim`, ok: loc.ok && loc.rows === groupFinal && localClaim(loc, op) }];
}

// refs: none -- a row is not a text unit a later step can name
function producedText() {
  return null;
}

module.exports = { plan, annotateTask, classify, resolve, preDispatchCheck, lower, commandBody, verifyPrimitive, verifyPreservation, wholeTaskChecks, producedText, locate, INTENTS };
