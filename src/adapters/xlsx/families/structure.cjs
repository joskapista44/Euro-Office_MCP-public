'use strict';

const { colNum, colName, parseRange, rangeText, MAX_COL, MAX_ROW } = require('../a1.cjs');
const { key } = require('../inventory.cjs');
const { refOnTarget } = require('../dependencies.cjs');
const formula = require('../formula.cjs');
const { shiftFormula, referencesBlock, moveIndex, moveSpan } = require('../shift.cjs');
const { layoutGuard, layoutDiffs } = require('../sheet-layout.cjs');
const drawings = require('../drawings.cjs');

/**
 * XLSX `structure` family: insert_rows / delete_rows / insert_columns / delete_columns (M6 wave 5).
 * docs/evidence/M6-W5-STRUCTURE-AUDIT-AND-CONTRACT.md sections 3 and 5 (the call shape measured by the W5 probe,
 * LIVE-accepted 2026-09-29).
 *
 * The source's retry needed a token the MCP could not carry (audit X1): an identical retry inserted / deleted
 * AGAIN. Here the retry is decided from a fresh read, as the DOCX table-row precedent (M4 W4b): the caller names
 * a content ANCHOR (a unique value in one column / row) and where it is now; the terminal position follows:
 *   insert at N, count k: anchor at N -> APPLY | anchor at N + k and the k new rows blank -> SATISFIED
 *   delete at N, count k (owner decision B, 2026-11-02, owner channel): TWO witnesses, both read before the first write --
 *     the anchor names the FIRST SURVIVING row (now at N + k), the witness the FIRST REMOVED row (now at N); together they
 *     bracket the block. Anchor at N + k AND witness at N -> APPLY. Anchor at N with the witness gone reads like an
 *     already-run delete, but no fresh read can tell that from a wrong anchor / witness (the defect this decision closes:
 *     an anchor given as the first REMOVED row's value sat at N and read SATISFIED, 0 rows deleted) -- so it is REFUSED,
 *     named `xlsx-structure-delete-prior-execution-unprovable`, never SATISFIED: a delete has no SATISFIED path until a
 *     prior execution can be proven by something other than matching data. A retry still never deletes twice.
 *   anything else -> refused (neither the before nor the after state)
 * Refused before the write, too: a formula anywhere the delete would BREAK (a reference into the block, a range
 * entirely inside it -> #REF!; a range spanning the block shrinks and is fine); a merge
 * crossing the block edge in an unmeasured shape (M7 S5 R5: see mergeRefusal); ANY validation, conditional format, chart or pivot that reaches the target sheet, and any
 * rule / object hosted on it (W5 probe: validation does not follow an insert, a CF's formula does not follow at all);
 * a defined name only outside the M7 S4 R1 shape (nameGuard below: one absolute cell -- those follow); an
 * unobservable dependency reader.
 * L1: every cell of the target sheet is where the shift puts it, with its value (literal) and its formula
 * (reference-aware shift, shift.cjs); an insert leaves the new rows / columns blank; merges moved with it.
 * L3: every other sheet's cells keep their values and their formulas re-pointed exactly as the shift implies.
 * A structural step is the only step of its task -- EXCEPT delete_rows, which may repeat with itself (L-19 C,
 * adapter.cjs planOperations): each repeat's `at` / `anchor` names the position as it will be once the task's
 * OWN earlier steps have already run, not the task-start position. Retry SATISFIED, 0 write / 0 command; in a
 * delete_rows batch, a SATISFIED verdict for a step that follows one this SAME run had to APPLY is refused
 * instead (adapter.cjs classifyOperation, `xlsx-structure-satisfied-after-applied-in-batch`) -- SATISFIED steps
 * of one run must be an unbroken prefix, since a retry can only ever find an earlier RUN's prefix already done.
 *
 * M7 S2a (owner GO 2026-09-30, owner channel; sheet-layout.cjs): the AutoFilter, the row / column page breaks and the
 * frozen pane of EVERY sheet are read before and after. Refused before the write: an object of the shifted sheet that is
 * not BEFORE the cut on the step's axis, and any unreadable axis on any sheet (incl. a split / unknown pane and the
 * column breaks until their READ is bound). L3: all of them unchanged on every sheet.
 *
 * fa7244ac/5 (Q4 MEASURED, c19b1065; sheet-layout.cjs autofilterRowInsertWithinData / filterExtendedAsExpected):
 * insert_rows is the ONE exception to the paragraph above's "not BEFORE the cut" refusal -- when the cut lies
 * strictly within the shifted sheet's OWN AutoFilter DATA rows (after its header, at or before its last data row),
 * the engine itself grows the filter's ref by the inserted count and leaves the header and every criterion alone
 * (MEASURED, not assumed), so the step is no longer refused there. L3 then expects EXACTLY that growth on that one
 * sheet; any other difference to it, or any difference on another sheet, still fails. Everything else from the
 * paragraph above is unchanged: a cut at or through the header row, a cut past the last data row (there the ref's
 * own fate is NOT MEASURED), delete_rows, insert_columns, delete_columns and a Table's own AutoFilter all remain
 * refused exactly as before.
 *
 * M7 E R2 / R3a (owner M7 GO 2026-10-03; tier E probe m7-e-probe1-20261003.json, session = reopen): a validation or a
 * conditional format HOSTED on the shifted sheet no longer refuses when it is entirely BEFORE the cut -- its range ends
 * before the first inserted / deleted row (column), and its formulas are constants, strings or references on that sheet
 * that, stretched over the rule's own extent, also end before the cut (no sheet-qualified reference, no defined name).
 * Measured: a list validation, an xlCellValue rule and an expression rule over the rows before the cut stay unchanged
 * through insert rows, delete rows and insert columns. delete_columns over such a rule is NOT measured and still
 * refuses. A rule on ANOTHER sheet that references the shifted sheet still refuses. L3 adds every sheet's rules
 * unchanged. Charts (R4): measured that a series after the cut does NOT follow (it then shows other data).
 *
 * M7 D-B1 / D-B2 (owner GO 2026-10-03, owner channel point 2): drawings (charts included) on the shifted sheet
 * entirely before the cut, by the build-gated anchor read (drawingGuard), and chart series reaching the sheet entirely
 * before the cut (seriesGuard) no longer refuse; delete_columns over either is not measured and refuses. L3 adds every
 * sheet's chart series and every drawing's anchor read unchanged.
 */

const INTENTS = Object.freeze(['insert_rows', 'delete_rows', 'insert_columns', 'delete_columns']);
const FIELDS = Object.freeze(['intent', 'id', 'sheet', 'at', 'count', 'anchor']);
const DELETE_FIELDS = Object.freeze([...FIELDS, 'witness']);
const MAX_COUNT = 100;

/** The lane (a column number for rows, a row number for columns) of an anchor / witness, or null. */
const laneOf = (x, rows) => (rows ? (typeof x.column === 'string' && /^[A-Z]{1,3}$/.test(x.column) ? colNum(x.column) : null) : (Number.isInteger(x.row) && x.row >= 1 ? x.row : null));
const valueOk = v => (typeof v === 'string' && v !== '') || Number.isFinite(v);

/** Owner decision B: the delete's second witness -- the value of the FIRST REMOVED row / column, read where it is NOW (at). */
function witnessOf(step, at, rows, anchorLane, pos) {
  const w = step.witness;
  const along = rows ? 'column' : 'row';
  const what = rows ? 'row' : 'column';
  if (!w || typeof w !== 'object' || Array.isArray(w) || Object.keys(w).some(k => ![along, 'value'].includes(k))) return { detail: `${at}: witness must be { ${along}, value } -- the value of the FIRST ${what} the delete removes (the ${what} at \`at\`) NOW; required: with the anchor (the first SURVIVING ${what}) it brackets the block` };
  const lane = laneOf(w, rows);
  if (lane === null) return { detail: `${at}: witness.${along} must be ${rows ? 'a column letter' : 'a row number'}` };
  if (!valueOk(w.value)) return { detail: `${at}: witness.value must be a non-empty text or a number` };
  if (lane === anchorLane && String(w.value) === String(step.anchor.value)) return { detail: `${at}: the witness and the anchor name the same value in the same ${along} -- the witness is the first REMOVED ${what} (at), the anchor the first SURVIVING one (at + count)` };
  return { witness: Object.freeze({ lane, value: String(w.value), now: pos }) };
}

function plan(step, index) {
  const at = `step ${index}: ${step.intent}`;
  const insert = step.intent.startsWith('insert');
  const unknown = Object.keys(step).filter(k => !(insert ? FIELDS : DELETE_FIELDS).includes(k));
  if (unknown.length) return { ok: false, detail: `${at} does not support field(s): ${unknown.join(', ')}` };
  if (typeof step.sheet !== 'string' || step.sheet === '') return { ok: false, detail: `${at} requires a sheet name` };
  const rows = step.intent.endsWith('_rows');
  const pos = rows ? (Number.isInteger(step.at) && step.at >= 1 && step.at <= MAX_ROW ? step.at : null)
    : (typeof step.at === 'string' && /^[A-Z]{1,3}$/.test(step.at) && colNum(step.at) <= MAX_COL ? colNum(step.at) : null);
  if (pos === null) return { ok: false, detail: `${at}: at must be ${rows ? 'a row number' : "a column letter ('B')"}` };
  if (!(Number.isInteger(step.count) && step.count >= 1 && step.count <= MAX_COUNT)) return { ok: false, detail: `${at}: count must be an integer in [1, ${MAX_COUNT}]` };
  const a = step.anchor;
  const along = rows ? 'column' : 'row';
  if (!a || typeof a !== 'object' || Array.isArray(a) || Object.keys(a).some(k => ![along, 'value'].includes(k))) return { ok: false, detail: `${at}: anchor must be { ${along}, value } -- the value of the ${rows ? 'row' : 'column'} that is at ${step.intent.startsWith('insert') ? 'at' : 'at + count'} NOW` };
  const lane = laneOf(a, rows);
  if (lane === null) return { ok: false, detail: `${at}: anchor.${along} must be ${rows ? 'a column letter' : 'a row number'}` };
  if (!valueOk(a.value)) return { ok: false, detail: `${at}: anchor.value must be a non-empty text or a number` };
  const w = insert ? { witness: null } : witnessOf(step, at, rows, lane, pos);
  if (w.detail) return { ok: false, detail: w.detail };
  const count = insert ? step.count : -step.count;
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`,
      intent: step.intent,
      family: 'structure',
      selector: Object.freeze({ sheet: step.sheet, range: rows ? `A${pos}:${colName(MAX_COL)}${pos + step.count - 1}` : `${colName(pos)}1:${colName(pos + step.count - 1)}${MAX_ROW}` }),
      shift: Object.freeze({ sheet: step.sheet, axis: rows ? 'rows' : 'cols', at: pos, count }),
      anchor: Object.freeze({ lane, value: String(a.value), now: insert ? pos : pos + step.count, terminal: pos + (insert ? step.count : 0) }),
      ...(w.witness ? { witness: w.witness } : {}),
      cells: Object.freeze([]),
      preservationMode: 'PRIMITIVE_INTEGRAL',
    },
  };
}

const rowOf = a => Number(/\d+$/.exec(a)[0]);
const colOf = a => colNum(/^[A-Z]+/.exec(a)[0]);
const cellsOfSheet = (p, sheet) => Object.entries(p.cells).filter(([k]) => k.startsWith(key(sheet, ''))).map(([k, c]) => [k.slice(key(sheet, '').length), c]);

/** Where a mark's value (the anchor, the delete's witness) is now: every literal cell in its column (rows) / row (cols). */
function laneHits(p, op, mark) {
  const rows = op.shift.axis === 'rows';
  return cellsOfSheet(p, op.shift.sheet).filter(([a, c]) => c.f === '' && String(c.v) === mark.value && (rows ? colOf(a) === mark.lane : rowOf(a) === mark.lane)).map(([a]) => (rows ? rowOf(a) : colOf(a)));
}
const anchorHits = (p, op) => laneHits(p, op, op.anchor);

/** The k new rows / columns at N: every READ cell there blank. */
function gapBlank(p, op) {
  const rows = op.shift.axis === 'rows';
  const k = Math.abs(op.shift.count);
  return cellsOfSheet(p, op.shift.sheet).every(([a, c]) => { const i = rows ? rowOf(a) : colOf(a); return i < op.shift.at || i >= op.shift.at + k || (c.f === '' && (c.v === null || c.v === '')); });
}

function blockOf(op) {
  const k = -op.shift.count;
  return op.shift.axis === 'rows' ? { sheet: op.shift.sheet, r1: op.shift.at, r2: op.shift.at + k - 1, c1: 1, c2: MAX_COL } : { sheet: op.shift.sheet, c1: op.shift.at, c2: op.shift.at + k - 1, r1: 1, r2: MAX_ROW };
}

/**
 * M7 S4 R1 (owner GO 2026-10-03, owner channel): a defined name follows the shift the way a formula does (W5 probe +
 * the S4 probe, docs/evidence/M7-S4-R1-NAMES.md). Supported: a name whose text is ONE absolute cell on one sheet
 * (`Sheet!$B$5`, `'My Sheet'!$B$5`), at the scopes in NAME_SCOPES_MEASURED. Refused before the write: any other name
 * shape reaching the shifted sheet (range, mixed / relative, a name through another name, a formula), a name the
 * delete would turn into #REF!, an unmeasured scope, an unreadable name list.
 */
const NAME_SCOPES_MEASURED = Object.freeze({ workbook: true, sheet: true });
const ABS_CELL = /^(?:'((?:[^']|'')+)'|([A-Za-z_À-￿][\w.À-￿]*))!\$([A-Z]{1,3})\$(\d+)$/;

/** The one absolute cell a name's text names ({ sheet, col, row }), or null for any other shape. */
function absoluteCellOf(ref) {
  const m = ABS_CELL.exec(String(ref || '').trim());
  return m ? { sheet: m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2], col: m[3], row: Number(m[4]) } : null;
}

const sameSheet = (a, b) => String(a).toUpperCase() === String(b).toUpperCase();
const nameKey = d => `${d.name.toUpperCase()}\u0000${d.scopeSheet === null ? '' : d.scopeSheet.toUpperCase()}`;

/** A name's expected text after the shift: shifted when it names the shifted sheet, else unchanged. */
function expectedNameRef(op, ref) {
  const s = shiftFormula(`=${ref}`, null, op.shift);
  return { ref: s.broken ? null : s.formula.slice(1), broken: s.broken };
}

/** The definition behind one dependency entry of kind 'name' (same name, same scope). */
const definitionOf = (p, r) => p.definedNames.find(x => x.name === r.name && (x.scopeSheet === null ? r.host === null : r.host !== null && sameSheet(x.scopeSheet, r.host)));

/**
 * GH #331 (T06): an Excel Table of the shifted sheet is read as a defined name over its range, so every structural write
 * on a sheet holding a table was refused as a range-shaped name (baseline LIVE, installed 1.1.0-03cbc20c3e5e, through the
 * table AND outside it: test/live/xlsx-gh331-t06-table-aware-baseline-probe.live.cjs). The table itself is told by
 * read_tables (`tableReads`, cell Bridge /18): a table ENTIRELY BEFORE the cut on the step's axis does not move -- the
 * same rule this family already applies to charts and drawings -- so its name is supported; a table at or after the cut
 * (the write would shift it or cut through it, not measured) is refused, named. A name that is not a readable table of
 * the sheet keeps the range-shape refusal.
 */
function tableOfName(p, sheet, d) {
  const read = p.tableReads ? p.tableReads[sheet] : undefined;
  if (!read || typeof read !== 'object' || !Array.isArray(read.tables)) return null;
  const t = read.tables.find(x => String(x.name).toUpperCase() === String(d.name).toUpperCase());
  const g = t ? parseRange(String(t.ref)) : null;
  if (!g) return null;
  // the table's name refers to its DATA BODY -- the ref without the header row (and the totals row): measured LIVE,
  // table A1:C4 with a header -> name =Sheet!$A$2:$C$4 (xlsx-gh331-t06-table-aware-baseline-probe)
  const body = rangeText({ ...g, r1: g.r1 + (t.headerRow === true ? 1 : 0), r2: g.r2 - (t.totalsRow === true ? 1 : 0), count: 2 });
  const range = String(d.ref || '').replace(/^.*!/, '').replace(/\$/g, '').toUpperCase();
  return range === body ? t : null;
}

function tableRefusal(op, t) {
  const g = parseRange(String(t.ref));
  const before = g && (op.shift.axis === 'rows' ? g.r2 < op.shift.at : g.c2 < op.shift.at);
  return before ? null : { reason: 'xlsx-structure-overlaps-table', table: t.name, ref: t.ref, axis: op.shift.axis, at: op.shift.at };
}

/** One name reaching the shifted sheet: null when supported, else the refusal. */
function nameRefusal(op, d, r, p) {
  const sheet = op.shift.sheet;
  const at = { name: d.name, scope: d.scopeSheet, ref: d.ref };
  if (!NAME_SCOPES_MEASURED[d.scopeSheet === null ? 'workbook' : 'sheet']) return { reason: 'xlsx-structure-name-scope-not-measured', ...at };
  // a sheet-level name of the shifted sheet naming another sheet: its text does not move
  if (!r.sheets.some(x => sameSheet(x, sheet))) return null;
  const table = tableOfName(p, sheet, d);
  if (table) return tableRefusal(op, table);
  const cell = absoluteCellOf(d.ref);
  if (!cell || !sameSheet(cell.sheet, sheet)) return { reason: 'xlsx-structure-name-shape-unsupported', ...at };
  return expectedNameRef(op, d.ref).broken ? { reason: 'xlsx-structure-name-would-break', ...at } : null;
}

/** The names reaching the shifted sheet: each supported, or the first refusal. */
function nameGuard(p, op) {
  const sheet = op.shift.sheet;
  const reach = p.dependencies.refs.filter(r => r.kind === 'name' && (r.sheets.some(x => sameSheet(x, sheet)) || (r.host !== null && sameSheet(r.host, sheet))));
  if (!reach.length) return null;
  if (!Array.isArray(p.definedNames)) return { reason: 'xlsx-structure-names-unobservable' };
  for (const r of reach) {
    const d = definitionOf(p, r);
    const bad = d ? nameRefusal(op, d, r, p) : { reason: 'xlsx-structure-names-unobservable', name: r.name };
    if (bad) return bad;
  }
  return null;
}

/** L3 of the names: the same (name, scope) set; each text where the shift puts it (untouched names unchanged). */
function nameDiffs(op, b, a) {
  if (!Array.isArray(b.definedNames) || !Array.isArray(a.definedNames)) return [{ reason: 'xlsx-names-unobservable' }];
  const after = new Map(a.definedNames.map(d => [nameKey(d), d]));
  const out = [];
  for (const d of b.definedNames) {
    const now = after.get(nameKey(d));
    after.delete(nameKey(d));
    if (!now) { out.push({ name: d.name, scope: d.scopeSheet, change: 'missing' }); continue; }
    const want = expectedNameRef(op, d.ref);
    if (want.broken || !sameText(`=${want.ref}`, `=${now.ref}`)) out.push({ name: d.name, scope: d.scopeSheet, was: d.ref, want: want.ref, now: now.ref });
  }
  for (const d of after.values()) out.push({ name: d.name, scope: d.scopeSheet, change: 'appeared', now: d.ref });
  return out;
}

/**
 * M7 E R2 / R3a: one rule formula -- a constant / string, or same-sheet references that end before `early` (null =
 * fine). A sheet-qualified reference or a defined name is already a dependent of the sheet (guards: `reach`).
 */
function ruleFormulaProblem(f, op, early) {
  const text = f.startsWith('=') ? f : `=${f}`;
  const moved = early.at >= 1 ? shiftFormula(text, op.shift.sheet, early) : { broken: true };
  return moved.broken || formula.canonical(moved.formula) !== formula.canonical(text) ? 'rule-formula-reaches-the-cut' : null;
}

/** M7 E R2 / R3a: why a rule hosted on the shifted sheet may not stay as it is (null = entirely before the cut). */
function ruleRefusal(op, rule) {
  const { axis, at, count } = op.shift;
  if (axis !== 'rows' && count < 0) return 'delete-columns-over-a-rule-not-measured';
  const parts = String(rule.range || '').split(',').map(x => parseRange(x.trim()));
  const end = g => (axis === 'rows' ? g.r2 : g.c2);
  if (parts.some(g => !g || end(g) >= at)) return 'rule-range-at-or-after-the-cut-or-unreadable';
  // a relative reference moves with each cell of the rule's range: the check moves the cut back by that extent
  const early = { ...op.shift, at: at - Math.max(...parts.map(g => (axis === 'rows' ? g.r2 - g.r1 : g.c2 - g.c1))) };
  return [rule.formula, rule.formula2].filter(Boolean).map(f => ruleFormulaProblem(f, op, early)).find(Boolean) || null;
}

/** The pre-write guards: dependency readers complete, nothing unmeasured reaching the sheet, no #REF!, no split merge. */
function guards(p, op) {
  const sheet = op.shift.sheet;
  if (!p.dependencies || !p.dependencies.complete) return { reason: 'xlsx-structure-dependency-observation-incomplete', unobservable: p.dependencies ? p.dependencies.unobservable.slice(0, 20) : 'not-read' };
  const ref = refOnTarget(p.dependencies, [sheet]);
  if (ref) return ref;
  const named = nameGuard(p, op);
  if (named) return named;
  const hostedRules = p.rules && p.rules.bySheet && p.rules.bySheet[sheet] ? p.rules.bySheet[sheet] : [];
  const unsafe = hostedRules.map(r => ({ rule: r, why: ruleRefusal(op, r) })).filter(x => x.why);
  if (unsafe.length) return { reason: 'xlsx-structure-dependent-not-measured', rules: unsafe.slice(0, 20) };
  // any rule (on any sheet) with a sheet-qualified reference or a name to this sheet, any chart or pivot reaching it
  // (M7 D-B: a chart is decided by the drawing guard -- its geometry -- and the series guard -- its data source)
  const reach = p.dependencies.refs.filter(r => r.kind !== 'formula' && r.kind !== 'name' && r.kind !== 'chart' && (r.sheets.includes(sheet) || r.host === sheet));
  if (reach.length) return { reason: 'xlsx-structure-dependent-not-measured', dependents: reach.slice(0, 20) };
  const objects = p.objects && p.objects[sheet];
  if (objects && Array.isArray(objects.pivots) && objects.pivots.length) return { reason: 'xlsx-structure-dependent-not-measured', objects: sheet };
  if (op.shift.count < 0) {
    const block = blockOf(op);
    // only a reference the delete BREAKS (a cell in the block, a range entirely inside it) -- a range that spans the
    // block shrinks (Excel); referencesBlock names the candidates, the shift decides
    const refs = Object.entries(p.cells).filter(([k, c]) => c.f && referencesBlock(c.f, k.split('\u0000')[0], block) && shiftFormula(c.f, k.split('\u0000')[0], op.shift).broken).map(([k]) => k.replace('\u0000', '!'));
    if (refs.length) return { reason: 'xlsx-structure-delete-referenced-by-formulas', cells: refs.slice(0, 20), count: refs.length };
  }
  for (const [a, c] of cellsOfSheet(p, sheet)) {
    if (!c.merge || c.merge === a) continue;
    const bad = mergeRefusal(op, c.merge);
    if (bad) return { reason: bad, cell: a, mergeArea: c.merge };
  }
  return null;
}

/**
 * M7 S5 R5 (owner M7 GO 2026-10-03; audit R5): a merge the cut crosses. MEASURED (W5 probe m6-w5-probe-20260929,
 * session = reopen): an insert strictly inside a merge GROWS it (rows at 3 x2: G2:H3 -> G2:H5; column B: A8:C8 -> A8:D8);
 * a delete inside it or across its far edge CLIPS it (column B: A8:C8 -> A8:B8; rows 3:4: G2:H3 -> G2:H2), the top-left
 * surviving. Supported: exactly those -- the top-left before the cut and at least two cells left. Refused
 * (NOT_MEASURED): a delete that removes the merge's top-left, a merge that collapses to one cell. A merge the cut does
 * not cross moves (M6, measured); one entirely inside a deleted block goes with it (the unchanged M6 path -- its engine
 * shape is NOT_MEASURED, audit R5).
 */
function mergeRefusal(op, area) {
  const g = parseRange(area);
  const rows = op.shift.axis === 'rows';
  const [lo, hi] = rows ? [g.r1, g.r2] : [g.c1, g.c2];
  const { at, count } = op.shift;
  if (count > 0) return null; // an insert either moves the merge or grows it -- both measured
  const last = at - count - 1;
  if (!((lo < at && hi >= at) || (lo <= last && hi > last))) return null;
  if (lo >= at) return 'xlsx-structure-merge-top-left-deleted';
  const span = moveSpan(lo, hi, op.shift);
  const other = rows ? g.c2 - g.c1 + 1 : g.r2 - g.r1 + 1;
  return (span[1] - span[0] + 1) * other < 2 ? 'xlsx-structure-merge-collapses-to-one-cell' : null;
}

/**
 * M7 S2b (owner GO 2026-10-01, owner channel): measured (M7-S2B-DRAWING-PROBE.md) a public row / column insert moves the
 * cells and NOT the drawings. The L3 claims every sheet's drawings, so an unreadable drawing anywhere refuses.
 *
 * M7 D-B1 / D-B2 (owner GO 2026-10-03, owner channel point 2; M7-DB-PROBE.md): a drawing on the shifted sheet no
 * longer refuses when its WHOLE extent is proven before the cut -- the build-gated internal anchor read
 * (drawings.anchorExtent: the measured oneCell mode, an observed size, rotation 0, the computed bottom-right cell) ends
 * on the step's axis before the first inserted / deleted row (column). Measured: a chart before the cut unchanged through
 * insert rows, delete rows and insert columns (M7-E-PROBE R4); drawings stable through an insert + reopen (D-B probe).
 * delete_columns over a drawing is NOT measured and refuses. An unknown mode or an incomplete geometry refuses before
 * the write -- a safety refusal, not a measured EO limit.
 * M7 D-B hydration (owner GO 2026-10-04, owner channel): on a NON-ACTIVE sheet the driver first creates the sheet's
 * view (since B15 /5 through the DS Bridge sheet.model.ensureReady/v1); a hydration that cannot run reads UNOBSERVABLE
 * and refuses here unchanged.
 */
function drawingGuard(p, op) {
  const { sheet, axis, at, count } = op.shift;
  const o = drawings.observe(p.drawingReads, p.sheets.map(x => x.name));
  if (!o.complete) return { reason: 'xlsx-structure-drawings-unobservable', unobservable: o.gaps.slice(0, 20), count: o.gaps.length };
  const raw = Array.isArray(p.drawingReads[sheet]) ? p.drawingReads[sheet] : [];
  if (!raw.length) return null;
  const names = raw.slice(0, 20).map(d => d.name);
  if (axis !== 'rows' && count < 0) return { reason: 'xlsx-structure-drawing-delete-columns-not-measured', drawings: names, count: raw.length };
  const ext = raw.map(d => ({ name: d.name, ...drawings.anchorExtent(d, sheet) }));
  const unknown = ext.filter(e => e.why);
  if (unknown.length) return { reason: 'xlsx-structure-drawing-geometry-unobservable', drawings: unknown.slice(0, 20), count: unknown.length };
  // 0-based last cell -> 1-based, the same index space as the cut
  const reaching = ext.filter(e => (axis === 'rows' ? e.lastRow : e.lastCol) + 1 >= at);
  return reaching.length ? { reason: 'xlsx-structure-drawing-at-or-after-the-cut', drawings: reaching.slice(0, 20), count: reaching.length } : null;
}

/**
 * M7 D-B: the chart series reaching the shifted sheet (owner: "the geometry guard does not replace the series
 * dependency check"). Measured (M7-E-PROBE R4 / C15): a series over cells before the cut stays unchanged, one after
 * it does NOT follow. Supported: the measured series shape -- ONE sheet-qualified cell / range ('=S!$B$2:$B$5') -- ending
 * before the cut. Refused before the write: any other shape reaching the sheet, a range at or after the cut,
 * delete_columns (not measured).
 */
const SERIES_REF = /^=(?:'((?:[^']|'')+)'|([^'!\s][^'!]*))!(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)$/;

function seriesProblem(text, op) {
  const m = SERIES_REF.exec(String(text).trim());
  if (!m) return 'series-shape-not-measured';
  if ((m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2]) !== op.shift.sheet) return 'series-shape-not-measured';
  if (op.shift.axis !== 'rows' && op.shift.count < 0) return 'delete-columns-over-a-series-not-measured';
  const g = parseRange(m[3].replace(/\$/g, ''));
  if (!g) return 'series-shape-not-measured';
  return (op.shift.axis === 'rows' ? g.r2 : g.c2) >= op.shift.at ? 'series-at-or-after-the-cut' : null;
}

function seriesGuard(p, op) {
  const sheet = op.shift.sheet;
  const bad = [];
  for (const r of p.dependencies.refs.filter(x => x.kind === 'chart' && x.sheets.includes(sheet))) {
    const [i, j, f] = r.name.split('.');
    const text = p.objects[r.host].charts[Number(i)].series[Number(j)][f];
    const why = seriesProblem(text, op);
    if (why) bad.push({ host: r.host, chart: Number(i), series: Number(j), field: f, text, why });
  }
  return bad.length ? { reason: 'xlsx-structure-chart-series-not-before-the-cut', series: bad.slice(0, 20), count: bad.length } : null;
}

/** L3 of the charts' data (M7 D-B): every sheet's series texts unchanged (a kept series is before the cut). */
function seriesDiffs(b, a) {
  const of = x => Object.fromEntries(x.sheets.map(s => s.name).map(n => {
    const o = x.objects && x.objects[n];
    return [n, o && Array.isArray(o.charts) ? o.charts.map(c => c.series) : 'unobservable'];
  }));
  const was = of(b);
  const now = of(a);
  return Object.keys(was).filter(n => was[n] === 'unobservable' || JSON.stringify(was[n]) !== JSON.stringify(now[n])).map(n => ({ sheet: n, was: was[n], now: now[n] }));
}

/** L3 of the anchors (M7 D-B): every drawing's anchor read unchanged, by sheet and name (unread on both sides = no claim). */
function anchorDiffs(b, a) {
  const of = x => Object.entries(x.drawingReads || {}).flatMap(([n, l]) => (Array.isArray(l) ? l.map(d => [`${n}\u0000${d.name}`, JSON.stringify(d.anchor === undefined ? null : d.anchor)]) : []));
  const now = new Map(of(a));
  return of(b).filter(([k, v]) => now.get(k) !== v).map(([k, v]) => ({ drawing: k.replace('\u0000', '!'), was: v, now: now.has(k) ? now.get(k) : 'missing' }));
}

function gate(p, op) {
  if (!p || p.format !== 'xlsx') return { ok: false, reason: 'xlsx-observation-missing' };
  if (!p.complete) return { ok: false, reason: p.truncated ? 'xlsx-observation-truncated' : 'xlsx-observation-unreadable-getters', unsupported: p.unsupported };
  const s = p.sheets.filter(x => x.name === op.shift.sheet);
  if (s.length !== 1) return { ok: false, reason: s.length > 1 ? 'xlsx-sheet-ambiguous' : 'xlsx-sheet-not-found', available: p.sheets.map(x => x.name) };
  const hits = anchorHits(p, op);
  if (hits.length !== 1) return { ok: false, reason: hits.length ? 'xlsx-structure-anchor-not-unique' : 'xlsx-structure-anchor-not-found', anchor: op.anchor.value, at: hits.slice(0, 20) };
  // a delete's witness: present once (before the delete) or gone (after it) -- twice is no witness at all
  const witnessAt = op.witness ? laneHits(p, op, op.witness) : null;
  if (witnessAt && witnessAt.length > 1) return { ok: false, reason: 'xlsx-structure-witness-not-unique', witness: op.witness.value, at: witnessAt.slice(0, 20) };
  return { ok: true, pos: hits[0], witnessAt };
}

/**
 * Owner decision B: a delete APPLIES only when BOTH witnesses stand where the step says they are now; the after-shape
 * (anchor at `at`, witness gone) is refused as unprovable -- matching data is not a proof that this delete ever ran.
 */
function classifyDelete(op, g, where) {
  const at = { ...where, witnessAt: g.witnessAt, expectedNow: { anchor: op.anchor.now, witness: op.witness.now } };
  const anchorNow = g.pos === op.anchor.now;
  const anchorAfter = g.pos === op.anchor.terminal;
  if (!anchorNow && !anchorAfter) return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-structure-anchor-state-conflict', ...where, expectedNow: op.anchor.now, expectedAfter: op.anchor.terminal } };
  if (anchorNow && g.witnessAt.length === 1 && g.witnessAt[0] === op.witness.now) return null;
  if (anchorAfter && g.witnessAt.length === 0) {
    return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-structure-delete-prior-execution-unprovable', ...at, readsAsAlreadyApplied: true, written: 0,
      detail: 'the sheet reads as if this delete had already run (the anchor at `at`, the witness gone), but no read can prove THIS delete produced it -- a wrong anchor or witness reads the same; nothing was written. Re-read the sheet; re-plan with the witnesses as they are now if rows still need deleting.' } };
  }
  return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-structure-delete-witnesses-disagree', ...at, detail: 'the anchor (first SURVIVING row/column, now at at + count) and the witness (first REMOVED row/column, now at at) do not both stand where the step says -- the block the step names is not the block on the sheet; nothing was written' } };
}

function classify(p, op) {
  const g = gate(p, op);
  if (!g.ok) return { state: 'FAIL_CLOSED', evidence: g };
  const where = { sheet: op.shift.sheet, axis: op.shift.axis, at: op.shift.at, count: op.shift.count, anchorAt: g.pos };
  if (op.witness) {
    const refused = classifyDelete(op, g, where);
    if (refused) return refused;
  } else {
    if (g.pos === op.anchor.terminal && gapBlank(p, op)) return { state: 'SATISFIED', evidence: { ...where, proof: 'fresh-live-anchor-at-its-terminal-position' } };
    if (g.pos !== op.anchor.now) return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-structure-anchor-state-conflict', ...where, expectedNow: op.anchor.now, expectedAfter: op.anchor.terminal } };
  }
  const bad = guards(p, op) || layoutGuard(p, op.shift) || drawingGuard(p, op) || seriesGuard(p, op);
  if (bad) return { state: 'FAIL_CLOSED', evidence: bad };
  return { state: 'APPLY', evidence: op.witness ? { ...where, witnessAt: op.witness.now } : where };
}

function resolve(p, selector, ctx = {}) {
  const g = gate(p, ctx.operation);
  if (!g.ok) return { ok: false, target: null, selector, ambiguity: g.reason, detail: g };
  return { ok: true, target: Object.freeze({ sheet: ctx.operation.shift.sheet, anchorAt: g.pos }) };
}

/** One callCommand: exact-name sheet check, then the public entire-span Insert / Delete (the W5 probe's IR1 shape). */
function lower(op) {
  const { sheet, axis, at, count } = op.shift;
  const k = Math.abs(count);
  const span = axis === 'rows' ? `A${at}:${colName(MAX_COL)}${at + k - 1}` : `${colName(at)}1:${colName(at + k - 1)}${MAX_ROW}`;
  const call = count > 0 ? (axis === 'rows' ? "Insert('down')" : "Insert('right')") : (axis === 'rows' ? "Delete('up')" : "Delete('left')");
  return {
    kind: 'callCommand',
    body: `
try {
  const name = ${JSON.stringify(sheet)};
  const ws = Api.GetSheet(name);
  if (!ws || ws.GetName() !== name) return { ok: false, stage: 'sheet-identity' };
  ws.GetRange(${JSON.stringify(span)}).${call};
  return { ok: true };
} catch (err) { return { ok: false, stage: 'threw', error: String(err) }; }
`,
    meta: Object.freeze({ family: 'structure', action: op.intent, sheet, axis, at, count }),
  };
}

/** Where a before-cell of the target sheet lands (null = deleted). */
function landing(op, a) {
  const rows = op.shift.axis === 'rows';
  const r = rowOf(a);
  const c = colOf(a);
  const n = moveIndex(rows ? r : c, op.shift);
  if (n === null) return null;
  return rows ? `${colName(c)}${n}` : `${colName(n)}${r}`;
}

const sameText = (want, got) => { const a = formula.canonical(want); const b = formula.canonical(got); return a !== null && a === b; };

/** Does `now` hold `was` after the shift (literal: the value; formula: the shifted text, and on an insert its value)? */
function kept(op, host, was, now) {
  if (!now) return false;
  if (was.f === '') return now.f === '' && JSON.stringify(now.v) === JSON.stringify(was.v);
  const want = shiftFormula(was.f, host, op.shift);
  if (want.broken || now.f === '' || !sameText(want.formula, now.f)) return false;
  return op.shift.count < 0 || JSON.stringify(now.v) === JSON.stringify(was.v);
}

/** L1: the target sheet -- every cell where the shift puts it, the gap blank, merges moved, the anchor at its terminal. */
function verifyPrimitive(op, before, after) {
  const b = before && before.primary;
  const a = after && after.primary;
  const sheet = op.shift.sheet;
  const wrong = [];
  for (const [at, was] of cellsOfSheet(b, sheet)) {
    const to = landing(op, at);
    if (to === null) continue;
    const now = a.cells[key(sheet, to)];
    if ((was.f !== '' || (was.v !== null && was.v !== '')) && !kept(op, sheet, was, now)) wrong.push({ from: at, to, was: { v: was.v, f: was.f }, now: now ? { v: now.v, f: now.f } : 'unread' });
    if (was.merge && was.merge !== at && now) {
      const g = parseRange(was.merge);
      const rows = op.shift.axis === 'rows';
      const s = rows ? moveSpan(g.r1, g.r2, op.shift) : moveSpan(g.c1, g.c2, op.shift);
      const want = s ? rangeText(rows ? { ...g, r1: s[0], r2: s[1], count: 2 } : { ...g, c1: s[0], c2: s[1], count: 2 }) : null;
      if (want && now.merge !== want) wrong.push({ from: at, to, merge: { was: was.merge, want, now: now.merge } });
    }
  }
  const hits = anchorHits(a, op);
  const anchorOk = hits.length === 1 && hits[0] === op.anchor.terminal;
  const gapOk = op.shift.count < 0 || gapBlank(a, op);
  // owner decision B: the delete's witness (unique before it, in the removed block) is gone after it
  const witnessGone = !op.witness || laneHits(a, op, op.witness).length === 0;
  return { ok: wrong.length === 0 && anchorOk && gapOk && witnessGone, evidence: { sheet, axis: op.shift.axis, at: op.shift.at, count: op.shift.count, wrong: wrong.slice(0, 20), wrongCount: wrong.length, anchorAt: hits, gapBlank: gapOk, ...(op.witness ? { witnessGone } : {}), proof: 'fresh-live-every-cell-at-its-shifted-position' } };
}

/** L3: every OTHER sheet -- values kept, formulas re-pointed exactly as the shift implies. */
function verifyPreservation(op, before, after) {
  const b = before && before.primary;
  const a = after && after.primary;
  if (!b || !a || !b.complete || !a.complete) return { ok: false, evidence: { reason: 'xlsx-preservation-observation-incomplete' } };
  const names = x => x.sheets.map(s => s.name).join('\u0000');
  if (names(b) !== names(a)) return { ok: false, evidence: { reason: 'xlsx-sheet-list-changed' } };
  const changed = [];
  let compared = 0;
  for (const [k, was] of Object.entries(b.cells)) {
    const host = k.split('\u0000')[0];
    if (host === op.shift.sheet) continue;
    if (was.f === '' && (was.v === null || was.v === '')) continue;
    compared += 1;
    if (!kept(op, host, was, a.cells[k])) changed.push({ cell: k.replace('\u0000', '!'), was: { v: was.v, f: was.f }, now: a.cells[k] ? { v: a.cells[k].v, f: a.cells[k].f } : 'unread' });
  }
  for (const [k, now] of Object.entries(a.cells)) {
    if (k.split('\u0000')[0] === op.shift.sheet || b.cells[k]) continue;
    if (now.f !== '' || (now.v !== null && now.v !== '')) changed.push({ cell: k.replace('\u0000', '!'), change: 'appeared' });
  }
  const layout = layoutDiffs(b, a, op.shift);
  const sheetsOf = x => x.sheets.map(s => s.name);
  const draw = drawings.drawingDiffs(drawings.observe(b.drawingReads, sheetsOf(b)), drawings.observe(a.drawingReads, sheetsOf(a)));
  const named = nameDiffs(op, b, a);
  const series = seriesDiffs(b, a);
  const anchors = anchorDiffs(b, a);
  // M7 E R2 / R3a: every sheet's rules unchanged (a rule kept on the shifted sheet is entirely before the cut)
  const rulesKept = !!b.rules && !!a.rules && !!b.rules.bySheet && JSON.stringify(b.rules.bySheet) === JSON.stringify(a.rules.bySheet);
  // GH #331: the shifted sheet's tables unchanged (a table is kept only when it lies entirely before the cut)
  const tablesOf = x => JSON.stringify(x.tableReads ? x.tableReads[op.shift.sheet] : null);
  const tablesKept = tablesOf(b) === tablesOf(a);
  return { ok: changed.length === 0 && layout.length === 0 && draw.length === 0 && named.length === 0 && rulesKept && tablesKept && series.length === 0 && anchors.length === 0, evidence: { compared, rulesKept, tablesKept, seriesKept: series.length === 0, seriesDiffs: series.slice(0, 10), anchorsKept: anchors.length === 0, anchorDiffs: anchors.slice(0, 10), changed: changed.slice(0, 20), changedCount: changed.length, layoutKept: layout.length === 0, layoutDiffs: layout.slice(0, 20), drawingsKept: draw.length === 0, drawingDiffs: draw, namesCompared: Array.isArray(b.definedNames) ? b.definedNames.length : null, namesKept: named.length === 0, nameDiffs: named.slice(0, 20) } };
}

function wholeTaskChecks(op, finalObservation) {
  const p = finalObservation && finalObservation.primary;
  const hits = p ? anchorHits(p, op) : [];
  const witnessAt = p && op.witness ? laneHits(p, op, op.witness) : [];
  const ok = hits.length === 1 && hits[0] === op.anchor.terminal && (op.shift.count < 0 || gapBlank(p, op)) && witnessAt.length === 0;
  return [{ name: `${op.shift.sheet}: the anchor '${op.anchor.value}' at ${op.shift.axis === 'rows' ? 'row' : 'column'} ${op.shift.axis === 'rows' ? op.anchor.terminal : colName(op.anchor.terminal)}${op.shift.count > 0 ? ', the new span blank' : ''}${op.witness ? `, the witness '${op.witness.value}' gone` : ''}`, ok, detail: ok ? null : { anchorAt: hits, ...(op.witness ? { witnessAt } : {}) } }];
}

const windowsOf = () => [];
// M7 S2a: every sheet's AutoFilter, print READ (row / column breaks) and pane
// M7 D-B: the build-gated anchor read of every drawing
// GH #331: the shifted sheet's tables (read_tables' axis), for the table-name rule and its L3
const liveExtras = op => ({ dependents: true, filterSheets: ['*'], printSheets: ['*'], colBreaks: true, panes: true, drawingSheets: ['*'], drawingAnchor: true, ...(op && op.shift ? { tableSheets: [op.shift.sheet] } : {}) });

module.exports = { INTENTS, plan, classify, resolve, lower, verifyPrimitive, verifyPreservation, wholeTaskChecks, windowsOf, liveExtras };
