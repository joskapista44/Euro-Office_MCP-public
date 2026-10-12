'use strict';

const { parseRange, rangeText, colName } = require('../a1.cjs');
const { cellAt, isBlank } = require('../inventory.cjs');
const cells = require('./cells.cjs');
const { aggregate } = require('../pivot-aggregate.cjs');
const { layoutOf, parseRef } = require('../pivot-source.cjs');
const { MAX_DATA_ROWS, SHEET, NAME, sourceGrid, tableMismatch, cellsOf, overlaps } = require('./pivots.cjs');

/**
 * XLSX `refresh_pivot` (M7 tier H P3, owner GO 2026-10-04, owner channel): rebind a pivot to a source the CALLER
 * names, then refresh -- the PUBLIC ApiPivotTable.SetSource(range) + RefreshTable() (source audit, sdkjs d8a3d24:
 * SetSource -> asc_set, the same rebuild the UI's "Change Data Source" runs; an EQUAL reference is a no-op there and the
 * refresh re-resolves it by name). LIVE m7-p3c-probe4: after a reopen, next to a same-named impostor of identical content,
 * SetSource binds the NAMED sheet and the body follows it; other pivots unchanged; the binding survives a reopen.
 *
 * Contract: refresh_pivot { pivot, sourceSheet, sourceRange }. The old source is never inferred (a stale pivot is the
 * reason to refresh, and it fails the content proof by design): the pivot is found by NAME through its raw identity
 * (pivotRefs: the approved asc_getDataRef READ, fields, table, layout).
 * Refused before the write: a sheet name that needs quoting (the public setter builds an UNQUOTED reference and is a
 * silent no-op on such a name -- MEASURED, the supported path's limit); a pivot that is missing, ambiguous or not in the
 * measured layout (one row field, one column field, `Sum of <field>`); a source whose header row is not exactly the
 * pivot's field list, or outside the measured shape; a pivots axis unreadable on any sheet; a grown table extent that is
 * not empty (the insert overwrites silently, P-A3) or that overlaps the source.
 * SATISFIED: the data reference is exactly the requested one AND the table is the requested source's aggregation now --
 * so a retry, also reopened, is 0 / 0.
 * L1: the exact data reference, the table extent and every table cell reproducing the fresh source's sums, the cells a
 * shrunk table left blank. L3: every cell outside the old and new extents (all sheets), every sheet's charts / names and
 * every OTHER pivot's raw identity unchanged.
 */

const INTENTS = Object.freeze(['refresh_pivot']);
const FIELDS = ['intent', 'id', 'pivot', 'sourceSheet', 'sourceRange'];

const PLAN_CHECKS = Object.freeze([
  [step => Object.keys(step).some(k => !FIELDS.includes(k)), step => `does not support field(s): ${Object.keys(step).filter(k => !FIELDS.includes(k)).join(', ')}`],
  [step => typeof step.pivot !== 'string' || !NAME.test(step.pivot), () => 'pivot must be a pivot name (letters, digits, _ . ; starting with a letter or _)'],
  [step => typeof step.sourceSheet !== 'string' || !SHEET.test(step.sourceSheet), () => 'sourceSheet must be a name that needs no quoting in a reference (the public SetSource is a silent no-op on a quoted one -- measured)'],
  [step => { const g = parseRange(step.sourceRange); return !g || g.r2 - g.r1 < 1 || g.r2 - g.r1 > MAX_DATA_ROWS || g.c2 - g.c1 < 2; }, () => `sourceRange must be a plain A1 range: a header row, 1..${MAX_DATA_ROWS} data rows, at least three columns`],
]);

function plan(step, index) {
  const at = `step ${index}: ${step.intent}`;
  const hit = PLAN_CHECKS.find(([bad]) => bad(step));
  if (hit) return { ok: false, detail: `${at}: ${hit[1](step)}` };
  const g = parseRange(step.sourceRange);
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`,
      intent: step.intent,
      family: 'pivots',
      selector: Object.freeze({ sheet: step.sourceSheet, range: rangeText(g), pivot: step.pivot, claims: Object.freeze([`pivot:*:${step.pivot}`]) }),
      name: step.pivot,
      grid: Object.freeze({ ...g }),
      cells: Object.freeze([]),
      preservationMode: 'PRIMITIVE_INTEGRAL',
    },
  };
}

const wantRef = op => `${op.selector.sheet}!$${colName(op.grid.c1)}$${op.grid.r1}:$${colName(op.grid.c2)}$${op.grid.r2}`;
const refMatch = (pv, op) => { const r = parseRef(pv.dataRef); return !!r && r.sheet === op.selector.sheet && r.range === op.selector.range; };

/** The one pivot of the name, with its host -- every sheet's raw pivot identities must be readable. */
function findPivot(p, name) {
  const refs = p.pivotRefs || {};
  const gaps = p.sheets.filter(s => !Array.isArray(refs[s.name])).map(s => ({ sheet: s.name, why: refs[s.name] || 'pivots-not-read' }));
  if (gaps.length) return { why: 'xlsx-pivot-observation-incomplete', gaps: gaps.slice(0, 20) };
  const hits = p.sheets.flatMap(s => refs[s.name].filter(x => x.name === name).map(x => ({ host: s.name, pv: x })));
  if (hits.length !== 1) return { why: hits.length ? 'xlsx-pivot-name-ambiguous' : 'xlsx-pivot-not-found', count: hits.length };
  return hits[0];
}

/** The observation itself: complete, with the named source sheet -> null or the refusal. */
function observationProblem(p, op) {
  if (!p || p.format !== 'xlsx') return { reason: 'xlsx-observation-missing' };
  if (!p.complete) return { reason: p.truncated ? 'xlsx-observation-truncated' : 'xlsx-observation-unreadable-getters', unsupported: p.unsupported };
  return p.sheets.filter(x => x.name === op.selector.sheet).length === 1 ? null : { reason: 'xlsx-sheet-not-found', sheets: [op.selector.sheet] };
}

/** The pivot's measured layout, the requested source's sums and the extent they give from the table's anchor. */
function modelOf(p, op, pv) {
  const lay = layoutOf({ fields: pv.fields, rows: pv.layout.rows, cols: pv.layout.cols, data: pv.layout.data });
  const old = parseRange(pv.table);
  if (lay.why || !old) return { reason: lay.why ? `xlsx-${lay.why}` : 'xlsx-pivot-table-unread', pivot: op.name };
  const s = sourceGrid(p, op);
  if (s.why) return { reason: s.why, ...s };
  const agg = aggregate(s.grid, pv.fields, lay.row, lay.col, lay.data);
  if (agg.why) return { reason: `xlsx-${agg.why}`, fields: pv.fields };
  const extent = { c1: old.c1, r1: old.r1, c2: old.c1 + agg.colKeys.size + 1, r2: old.r1 + agg.rowKeys.size + 2 };
  return { old, model: { agg, extent, extentText: rangeText(extent) } };
}

/** The fresh read: the pivot by name and its model against the requested source. */
function gate(p, op) {
  const bad = observationProblem(p, op);
  if (bad) return { ok: false, ...bad };
  const f = findPivot(p, op.name);
  if (f.why) return { ok: false, reason: f.why, ...f };
  const m = modelOf(p, op, f.pv);
  return m.reason ? { ok: false, ...m } : { ok: true, pv: f.pv, host: f.host, ...m };
}

/** A cell holds nothing: read blank, or outside its sheet's used range (the driver reads every used range whole). */
function free(p, sheet, a) {
  const cell = cellAt(p, sheet, a);
  if (cell) return isBlank(cell);
  const s = p.sheets.find(x => x.name === sheet);
  return !!s && (s.used === null || !overlaps(parseRange(s.used), parseRange(a)));
}

const inside = (e, a) => { const x = parseRange(a); return overlaps(e, x); };
const growthOf = g => cellsOf(g.model.extent).filter(a => !inside(g.old, a));
const fresh = (p, op, g) => refMatch(g.pv, op) && g.pv.table === g.model.extentText && !tableMismatch(p, { host: g.host }, g.model);

function destinationProblem(p, op, g) {
  if (g.host === op.selector.sheet && overlaps(g.model.extent, op.grid)) return { reason: 'xlsx-pivot-destination-overlaps-source', extent: g.model.extentText };
  const taken = growthOf(g).filter(a => !free(p, g.host, a));
  return taken.length ? { reason: 'xlsx-pivot-destination-not-empty', extent: g.model.extentText, cells: taken.slice(0, 20), count: taken.length } : null;
}

function classify(p, op) {
  const g = gate(p, op);
  if (!g.ok) return { state: 'FAIL_CLOSED', evidence: g };
  const where = { pivot: op.name, host: g.host, source: wantRef(op), extent: g.model.extentText, was: { dataRef: g.pv.dataRef, table: g.pv.table } };
  if (fresh(p, op, g)) return { state: 'SATISFIED', evidence: { ...where, proof: 'fresh-live-exact-dataref-and-table-content' } };
  const d = destinationProblem(p, op, g);
  return d ? { state: 'FAIL_CLOSED', evidence: { ...d, ...where } } : { state: 'APPLY', evidence: where };
}

function resolve(p, selector, ctx = {}) {
  const g = gate(p, ctx.operation);
  if (!g.ok) return { ok: false, target: null, selector, ambiguity: g.reason, detail: g };
  return { ok: true, target: Object.freeze({ sheet: selector.sheet, pivot: selector.pivot, host: g.host, extent: g.model.extentText, old: g.pv.table, growth: Object.freeze(growthOf(g)) }) };
}

function lower(op, target) {
  const J = JSON.stringify;
  return {
    kind: 'callCommand',
    body: `
try {
  const nm = x => { try { return x.GetName(); } catch (e) { return null; } };
  const hits = [];
  Api.GetSheets().forEach(w => (w.GetAllPivotTables() || []).forEach(x => { if (nm(x) === ${J(op.name)}) hits.push({ w, x }); }));
  if (hits.length !== 1 || hits[0].w.GetName() !== ${J(target.host)}) return { ok: false, stage: 'pivot-identity', count: hits.length };
  const src = Api.GetSheet(${J(op.selector.sheet)});
  if (!src || src.GetName() !== ${J(op.selector.sheet)}) return { ok: false, stage: 'sheet-identity' };
  // re-checked at dispatch: a grown table overwrites a non-empty cell silently (P-A3)
  for (const a of ${J(target.growth)}) { const v = hits[0].w.GetRange(a).GetValue(); if (v !== '' && v !== null && v !== undefined) return { ok: false, stage: 'destination-not-empty', cell: a }; }
  hits[0].x.SetSource(src.GetRange(${J(op.selector.range)}));
  hits[0].x.RefreshTable();
  return { ok: true };
} catch (err) { return { ok: false, stage: 'threw', error: String(err) }; }
`,
    meta: Object.freeze({ family: 'pivots', action: op.intent, name: op.name, sheet: op.selector.sheet, range: op.selector.range, host: target.host, extent: target.extent, old: target.old }),
  };
}

/** L1: the exact data reference, the table extent and content, the host unchanged, the cells a shrink vacated blank. */
function verifyPrimitive(op, before, after) {
  const g = gate(after && after.primary, op);
  if (!g.ok) return { ok: false, evidence: g };
  const b = gate(before && before.primary, op);
  const vacated = b.ok ? cellsOf(b.old).filter(a => !overlaps(g.model.extent, parseRange(a)) && !free(after.primary, g.host, a)) : [];
  const content = tableMismatch(after.primary, { host: g.host }, g.model);
  const ok = [refMatch(g.pv, op), g.pv.table === g.model.extentText, !content, !b.ok || b.host === g.host, vacated.length === 0].every(Boolean);
  return { ok, evidence: { pivot: op.name, host: g.host, dataRef: g.pv.dataRef, want: wantRef(op), table: g.pv.table, extent: g.model.extentText, content: content || 'reproduces-the-source', vacatedNotBlank: vacated.slice(0, 20), proof: 'fresh-live-exact-dataref-and-cells' } };
}

const othersOf = (p, name) => JSON.stringify({
  refs: Object.fromEntries(Object.entries(p.pivotRefs || {}).map(([s, r]) => [s, Array.isArray(r) ? r.filter(x => x.name !== name) : r])),
  objects: Object.fromEntries(Object.entries(p.objects || {}).map(([s, o]) => [s, o ? { ...o, pivots: undefined } : o])),
});

/** L3: every cell outside the old and new extents; every sheet's charts / names and every OTHER pivot's raw identity. */
function verifyPreservation(op, before, after) {
  const b = gate(before.primary, op);
  if (!b.ok) return { ok: false, evidence: { reason: b.reason } };
  const area = [...new Set([...cellsOf(b.old), ...cellsOf(b.model.extent)])];
  const rest = cells.verifyPreservation({ ...op, selector: { ...op.selector, sheet: b.host }, cells: area.map(a => ({ a })) }, before, after);
  if (!rest.ok && rest.evidence && rest.evidence.reason) return rest;
  const kept = othersOf(before.primary, op.name) === othersOf(after.primary, op.name);
  return { ok: rest.ok && kept, evidence: { ...rest.evidence, otherObjectsKept: kept } };
}

function wholeTaskChecks(op, finalObservation) {
  const v = verifyPrimitive(op, null, finalObservation);
  return [{ name: `pivot ${op.name} bound to ${wantRef(op)} and refreshed`, ok: v.ok, detail: v.ok ? null : v.evidence }];
}

const windowsOf = op => [{ sheet: op.selector.sheet, range: op.selector.range }];
const liveExtras = () => ({ objectSheets: ['*'] });

module.exports = { INTENTS, plan, classify, resolve, lower, verifyPrimitive, verifyPreservation, wholeTaskChecks, windowsOf, liveExtras };
