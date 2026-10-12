'use strict';

const { parseRange, rangeText, addr, colName } = require('../a1.cjs');
const { cellAt } = require('../inventory.cjs');
const cells = require('./cells.cjs');
const { aggregate, bodyMismatch } = require('../pivot-aggregate.cjs');

/**
 * XLSX `pivots` family: create_pivot (M7 tier H P1, owner GO 2026-10-03, owner channel point 1: "P1 on the
 * existing approved write path, with its own acceptance"). The write is the public InsertPivotExistingWorksheet +
 * AddFields + AddDataField + RefreshTable path the P-A / P-A2 / P-A3 probes used; the source after a reopen is the
 * content-proven one (pivot-source.cjs, P9).
 *
 * Contract (docs/evidence/M7-H3-PROBE.md, measured): create_pivot { name, sheet, range, rows, columns, sum, host, at } --
 * one row field, one column field, one `Sum of` data field (the only measured layout); the source range a text header
 * row (unique), text keys, a numeric sum field, no formula, at most MAX_DATA_ROWS data rows; sheet / host names that need
 * no quoting; `at` one cell. The table is (2 + row keys + 1) x (1 + column keys + 1) from `at` (P-A3, three shapes).
 * Refused before the write: an unmeasured shape, a pivot of the name anywhere with another identity, an unobservable
 * pivots axis on any sheet, a destination extent that is NOT entirely empty (P-A3: the insert silently overwrites it)
 * or that overlaps the source. SATISFIED: exactly one pivot of the name, on the host, with the source, the layout and the
 * extent -- in a reopened session its source is the content-proven one, so a retry is 0 / 0.
 * L1: the pivot read back with that identity and its table reproducing the source's sums cell by cell. L3: every cell
 * outside the extent, every sheet's charts / names / OTHER pivots unchanged.
 */

const INTENTS = Object.freeze(['create_pivot']);
const SHEET = /^[A-Za-z_][A-Za-z0-9_]*$/;
const NAME = /^[A-Za-z_][A-Za-z0-9_.]{0,99}$/;
const MAX_DATA_ROWS = 50;
const FIELDS = ['intent', 'id', 'name', 'sheet', 'range', 'rows', 'columns', 'sum', 'host', 'at'];

const oneLine = x => typeof x === 'string' && x !== '' && !/[\r\n]/.test(x);

/** The plan's measured-shape checks, in order: [refused?, why]. */
const PLAN_CHECKS = Object.freeze([
  [step => Object.keys(step).some(k => !FIELDS.includes(k)), step => `does not support field(s): ${Object.keys(step).filter(k => !FIELDS.includes(k)).join(', ')}`],
  [step => ![step.sheet, step.host].every(s => typeof s === 'string' && SHEET.test(s)), () => 'sheet / host must be names that need no quoting in a reference (a quoted one is not measured)'],
  [step => typeof step.name !== 'string' || !NAME.test(step.name), () => 'name must be a pivot name (letters, digits, _ . ; starting with a letter or _)'],
  [step => ![step.rows, step.columns, step.sum].every(oneLine) || new Set([step.rows, step.columns, step.sum]).size !== 3, () => 'rows / columns / sum must be three different header names'],
  [step => { const g = parseRange(step.range); return !g || g.r2 - g.r1 < 1 || g.r2 - g.r1 > MAX_DATA_ROWS || g.c2 - g.c1 < 2; }, () => `range must be a plain A1 range: a header row, 1..${MAX_DATA_ROWS} data rows, at least three columns`],
  [step => { const a = parseRange(step.at); return !a || a.count !== 1; }, () => 'at must be one cell'],
]);

function planProblem(step) {
  const hit = PLAN_CHECKS.find(([bad]) => bad(step));
  return hit ? hit[1](step) : null;
}

function plan(step, index) {
  const at = `step ${index}: ${step.intent}`;
  const why = planProblem(step);
  if (why) return { ok: false, detail: `${at}: ${why}` };
  const g = parseRange(step.range);
  const a = parseRange(step.at);
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`,
      intent: step.intent,
      family: 'pivots',
      // the pivot claims its NAME workbook-wide; it reads its source range (a write into it clashes)
      selector: Object.freeze({ sheet: step.sheet, range: rangeText(g), pivot: step.name, claims: Object.freeze([`pivot:*:${step.name}`]) }),
      name: step.name,
      host: step.host,
      at: Object.freeze({ c: a.c1, r: a.r1 }),
      fields: Object.freeze({ rows: step.rows, cols: step.columns, sum: step.sum }),
      grid: Object.freeze({ ...g }),
      cells: Object.freeze([]),
      preservationMode: 'PRIMITIVE_INTEGRAL',
    },
  };
}

const sourceText = op => `${op.selector.sheet}!$${colName(op.grid.c1)}$${op.grid.r1}:$${colName(op.grid.c2)}$${op.grid.r2}`;
const layoutOf = op => ({ rows: [op.fields.rows], cols: [op.fields.cols], data: [`Sum of ${op.fields.sum}`] });
const same = (a, b) => JSON.stringify(a) === JSON.stringify(b);

/** The source cells as a text grid, or the first cell outside the measured shape. */
function sourceGrid(p, op) {
  const g = op.grid;
  const out = [];
  for (let r = g.r1; r <= g.r2; r += 1) {
    const row = [];
    for (let c = g.c1; c <= g.c2; c += 1) {
      const cell = cellAt(p, op.selector.sheet, addr(c, r));
      if (!cell || cell.f !== '' || cell.count === null) return { why: 'xlsx-pivot-source-shape-not-measured', cell: addr(c, r), detail: cell ? 'formula-or-unreadable' : 'unread' };
      row.push(cell.v === null || cell.v === undefined ? '' : String(cell.v));
    }
    out.push(row);
  }
  return { grid: out };
}

/** The source's sums (pivot-aggregate.cjs) and the table extent they give, or a refusal. */
function sourceModel(p, op) {
  const s = sourceGrid(p, op);
  if (s.why) return s;
  const fields = s.grid[0];
  if (fields.some(f => f === '') || new Set(fields).size !== fields.length) return { why: 'xlsx-pivot-source-header-not-unique-text' };
  if (![op.fields.rows, op.fields.cols, op.fields.sum].every(f => fields.includes(f))) return { why: 'xlsx-pivot-field-not-in-header', fields };
  const agg = aggregate(s.grid, fields, op.fields.rows, op.fields.cols, op.fields.sum);
  if (agg.why) return { why: `xlsx-${agg.why}` };
  const extent = { c1: op.at.c, r1: op.at.r, c2: op.at.c + agg.colKeys.size + 1, r2: op.at.r + agg.rowKeys.size + 2 };
  return { agg, extent, extentText: rangeText(extent) };
}

/** Every sheet's pivots, flattened with their host; a sheet whose pivots axis is a reason string is a gap. */
function allPivots(p) {
  const out = [];
  const gaps = [];
  for (const s of p.sheets) {
    const o = p.objects && p.objects[s.name];
    if (!o || !Array.isArray(o.pivots)) gaps.push({ sheet: s.name, why: o ? o.pivots : 'objects-not-read' });
    else o.pivots.forEach(pv => out.push({ host: s.name, ...pv }));
  }
  return { pivots: out, gaps };
}

const cellsOf = e => { const out = []; for (let r = e.r1; r <= e.r2; r += 1) for (let c = e.c1; c <= e.c2; c += 1) out.push(addr(c, r)); return out; };
const blank = cell => !cell || (cell.f === '' && (cell.v === null || cell.v === ''));
const overlaps = (a, b) => a.c1 <= b.c2 && b.c1 <= a.c2 && a.r1 <= b.r2 && b.r1 <= a.r2;

/** The table cells of the extent reproduce the source's sums (labels, body, totals) -- the L1 / SATISFIED content proof. */
function tableMismatch(p, op, m) {
  const e = m.extent;
  const v = (c, r) => { const cell = cellAt(p, op.host, addr(c, r)); return cell && cell.v !== null && cell.v !== undefined ? String(cell.v) : ''; };
  const body = [];
  for (let r = e.r1 + 2; r <= e.r2; r += 1) { const row = []; for (let c = e.c1 + 1; c <= e.c2; c += 1) row.push(v(c, r)); body.push(row); }
  const rowLabels = body.map((_, i) => v(e.c1, e.r1 + 2 + i));
  const colLabels = body[0].map((_, j) => v(e.c1 + 1 + j, e.r1 + 1));
  return bodyMismatch(m.agg, body, rowLabels, colLabels);
}

function gate(p, op) {
  if (!p || p.format !== 'xlsx') return { ok: false, reason: 'xlsx-observation-missing' };
  const missing = [op.selector.sheet, op.host].filter(n => p.sheets.filter(x => x.name === n).length !== 1);
  if (!p.complete || missing.length) return !p.complete ? { ok: false, reason: p.truncated ? 'xlsx-observation-truncated' : 'xlsx-observation-unreadable-getters', unsupported: p.unsupported } : { ok: false, reason: 'xlsx-sheet-not-found', sheets: missing };
  const m = sourceModel(p, op);
  if (m.why) return { ok: false, reason: m.why, ...m };
  const all = allPivots(p);
  if (all.gaps.length) return { ok: false, reason: 'xlsx-pivot-observation-incomplete', gaps: all.gaps.slice(0, 20) };
  return { ok: true, model: m, hits: all.pivots.filter(x => x.name === op.name) };
}

const identityOk = (op, m, pv) => pv.host === op.host && pv.source === sourceText(op) && same(pv.layout, layoutOf(op)) && pv.table === m.extentText;

/** APPLY's own preconditions: the destination extent empty, off the source. */
function destinationProblem(p, op, m) {
  if (op.host === op.selector.sheet && overlaps(m.extent, op.grid)) return { reason: 'xlsx-pivot-destination-overlaps-source', extent: m.extentText };
  const taken = cellsOf(m.extent).filter(a => !blank(cellAt(p, op.host, a)));
  return taken.length ? { reason: 'xlsx-pivot-destination-not-empty', extent: m.extentText, cells: taken.slice(0, 20), count: taken.length } : null;
}

function classify(p, op) {
  const g = gate(p, op);
  if (!g.ok) return { state: 'FAIL_CLOSED', evidence: g };
  const where = { pivot: op.name, source: sourceText(op), host: op.host, extent: g.model.extentText };
  if (g.hits.length > 1) return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-pivot-name-ambiguous', ...where, count: g.hits.length } };
  if (g.hits.length === 1) {
    const ok = identityOk(op, g.model, g.hits[0]) && !tableMismatch(p, op, g.model);
    return ok ? { state: 'SATISFIED', evidence: { ...where, proof: 'fresh-live-pivot-with-the-identity-and-content' } }
      : { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-pivot-exists-with-another-identity', ...where, now: g.hits[0] } };
  }
  const d = destinationProblem(p, op, g.model);
  return d ? { state: 'FAIL_CLOSED', evidence: { ...d, ...where } } : { state: 'APPLY', evidence: where };
}

function resolve(p, selector, ctx = {}) {
  const g = gate(p, ctx.operation);
  if (!g.ok) return { ok: false, target: null, selector, ambiguity: g.reason, detail: g };
  return { ok: true, target: Object.freeze({ sheet: selector.sheet, pivot: selector.pivot, extent: g.model.extentText }) };
}

function lower(op, target) {
  const J = JSON.stringify;
  const g = op.grid;
  return {
    kind: 'callCommand',
    body: `
try {
  const src = Api.GetSheet(${J(op.selector.sheet)});
  const host = Api.GetSheet(${J(op.host)});
  if (!src || src.GetName() !== ${J(op.selector.sheet)} || !host || host.GetName() !== ${J(op.host)}) return { ok: false, stage: 'sheet-identity' };
  const nm = x => { try { return x.GetName(); } catch (e) { return null; } };
  if (Api.GetSheets().some(w => (w.GetAllPivotTables() || []).some(x => nm(x) === ${J(op.name)}))) return { ok: false, stage: 'name-taken' };
  // re-checked at dispatch: the insert silently overwrites a non-empty destination (P-A3)
  const now = host.GetRange(${J(target.extent)}).GetValue();
  if ((Array.isArray(now) ? now.flat() : [now]).some(v => v !== '' && v !== null && v !== undefined)) return { ok: false, stage: 'destination-not-empty' };
  const p = Api.InsertPivotExistingWorksheet(src.GetRange(${J(`${colName(g.c1)}${g.r1}:${colName(g.c2)}${g.r2}`)}), host.GetRange(${J(addr(op.at.c, op.at.r))}));
  if (!p) return { ok: false, stage: 'insert-returned-nothing' };
  p.SetName(${J(op.name)});
  p.AddFields({ rows: ${J(op.fields.rows)}, columns: ${J(op.fields.cols)} });
  p.AddDataField(${J(op.fields.sum)});
  p.RefreshTable();
  return { ok: true };
} catch (err) { return { ok: false, stage: 'threw', error: String(err) }; }
`,
    meta: Object.freeze({ family: 'pivots', action: op.intent, name: op.name, sheet: op.selector.sheet, range: op.selector.range, host: op.host, at: addr(op.at.c, op.at.r), extent: target.extent, fields: op.fields }),
  };
}

function verifyPrimitive(op, _before, after) {
  const g = gate(after && after.primary, op);
  if (!g.ok) return { ok: false, evidence: g };
  const pv = g.hits.length === 1 ? g.hits[0] : null;
  const content = pv ? tableMismatch(after.primary, op, g.model) : 'no-single-pivot';
  const ok = !!pv && identityOk(op, g.model, pv) && !content;
  return { ok, evidence: { pivot: op.name, now: g.hits, want: { host: op.host, source: sourceText(op), layout: layoutOf(op), table: g.model.extentText }, content: content || 'reproduces-the-source', proof: 'fresh-live-pivot-getters-and-cells' } };
}

/** L3: every cell outside the extent; every sheet's charts / names and every OTHER pivot. */
function verifyPreservation(op, before, after) {
  const m = sourceModel(before.primary, op);
  if (m.why) return { ok: false, evidence: { reason: m.why } };
  const rest = cells.verifyPreservation({ ...op, selector: { ...op.selector, sheet: op.host }, cells: cellsOf(m.extent).map(a => ({ a })) }, before, after);
  if (!rest.ok && rest.evidence && rest.evidence.reason) return rest;
  const others = p => JSON.stringify(Object.fromEntries(Object.entries(p.objects || {}).map(([s, o]) => [s, o && Array.isArray(o.pivots) ? { ...o, pivots: o.pivots.filter(x => x.name !== op.name) } : o])));
  const kept = others(before.primary) === others(after.primary);
  return { ok: rest.ok && kept, evidence: { ...rest.evidence, otherObjectsKept: kept } };
}

function wholeTaskChecks(op, finalObservation) {
  const v = verifyPrimitive(op, null, finalObservation);
  return [{ name: `pivot ${op.name} over ${sourceText(op)} at ${op.host}!${addr(op.at.c, op.at.r)}`, ok: v.ok, detail: v.ok ? null : v.evidence }];
}

/** The source and the largest extent the source can give (one key per data row on both axes). */
function windowsOf(op) {
  const n = op.grid.r2 - op.grid.r1;
  return [{ sheet: op.selector.sheet, range: op.selector.range }, { sheet: op.host, range: rangeText({ c1: op.at.c, r1: op.at.r, c2: op.at.c + n + 1, r2: op.at.r + n + 2 }) }];
}
const liveExtras = () => ({ objectSheets: ['*'] });

module.exports = { INTENTS, MAX_DATA_ROWS, SHEET, NAME, plan, classify, resolve, lower, verifyPrimitive, verifyPreservation, wholeTaskChecks, windowsOf, liveExtras, sourceGrid, tableMismatch, cellsOf, overlaps };
