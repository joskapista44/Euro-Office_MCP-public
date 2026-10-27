'use strict';

const { parseRange, rangeText, addr, colName } = require('../a1.cjs');
const { cellAt } = require('../inventory.cjs');
const cells = require('./cells.cjs');

/**
 * create_pivot v2 (GH #302 G3): n row / column / page fields, several data fields with their functions, a row-axis layout
 * and a style name -- the public InsertPivotExistingWorksheet + AddFields + AddDataField / SetFunction + SetRowAxisLayout +
 * SetStyleName + RefreshTable path. A step is v2 when it carries any of `pages` / `data` / `layout` / `style` or a field
 * LIST; the 1 + 1 + Sum shape stays v1 (pivots.cjs, content-proven), unchanged.
 *
 * The G3 LIVE design probe (docs/evidence/gh302-g3-create-pivot-v2-probe-20261026.json) set the rules:
 *   - the table's extent cannot be pre-computed: a page field puts its area ABOVE `at`, and an `at` too close to the top
 *     is SHIFTED down (A1 -> the table at A3) -- so v2 writes only onto an EMPTY pivotSheet (no used range), never the
 *     source sheet; nothing on it can be overwritten
 *   - every requested property reads back, session + reopen: the axes (G2 read), each data field's function AND its
 *     source field (GetSourceName -- the display name 'Sum of Amount2' holds a Count), every field's layout form, the
 *     style name; the public source only in session -- after a reopen the Bridge dataRef identifies it
 * Refused before the write (named): an unmeasured shape, a pivot of the name anywhere with another identity, a sheet whose
 * pivots are unread, a non-empty pivotSheet, the source sheet as pivotSheet, a header that is not unique text, a field
 * not in the header. SATISFIED: exactly one pivot of the name on pivotSheet with the whole identity (the retry after a
 * reopen: the source by the Bridge dataRef). L1: the pivot read back with that identity in the same session. L3: every
 * cell outside pivotSheet, every OTHER pivot / chart / name unchanged.
 */

const SHEET = /^[A-Za-z_][A-Za-z0-9_]*$/;
const NAME = /^[A-Za-z_][A-Za-z0-9_.]{0,99}$/;
const FUNCTIONS = Object.freeze(['Sum', 'Count', 'Average', 'Max', 'Min', 'Product', 'CountNumbers', 'StdDev', 'StdDevP', 'Var', 'VarP']);
const LAYOUTS = Object.freeze(['Compact', 'Outline', 'Tabular']);
const STYLE = /^PivotStyle(Light|Medium|Dark)([1-9]|1[0-9]|2[0-8])$/;
const MAX_DATA_ROWS = 1000;
const LIMITS = Object.freeze({ rows: [1, 4], columns: [0, 3], pages: [0, 3], data: [1, 4] });
const FIELDS = Object.freeze(['intent', 'id', 'name', 'sheet', 'range', 'rows', 'columns', 'pages', 'data', 'layout', 'style', 'pivotSheet', 'at']);

const oneLine = x => typeof x === 'string' && x !== '' && !/[\r\n]/.test(x);
const listOf = x => (x === undefined ? [] : (Array.isArray(x) ? x : [x]));
const same = (a, b) => JSON.stringify(a) === JSON.stringify(b);
const without = (o, keys) => Object.fromEntries(Object.entries(o).filter(([k]) => !keys.includes(k)));
const bareRef = s => String(s || '').replace(/\$/g, '').replace(/^'((?:[^']|'')+)'!/, (_, n) => `${n.replace(/''/g, "'")}!`);

/** a step is v2 when it names anything v1 cannot express */
const isV2 = step => ['pages', 'data', 'layout', 'style'].some(k => step[k] !== undefined) || Array.isArray(step.rows) || Array.isArray(step.columns);

/** The plan's measured-shape checks, in order: [refused?, why]. */
const PLAN_CHECKS = Object.freeze([
  [s => Object.keys(s).some(k => !FIELDS.includes(k)), s => `does not support field(s): ${Object.keys(s).filter(k => !FIELDS.includes(k)).join(', ')} (a v2 pivot names its data fields in \`data\`, not \`sum\`)`],
  [s => ![s.sheet, s.pivotSheet].every(x => typeof x === 'string' && SHEET.test(x)), () => 'sheet / pivotSheet must be names that need no quoting in a reference (a quoted one is not measured)'],
  [s => s.sheet === s.pivotSheet, () => 'pivotSheet must be another sheet than the source (v2 writes onto an empty sheet only)'],
  [s => typeof s.name !== 'string' || !NAME.test(s.name), () => 'name must be a pivot name (letters, digits, _ . ; starting with a letter or _)'],
  [s => ['rows', 'columns', 'pages'].some(k => { const l = listOf(s[k]); return l.length < LIMITS[k][0] || l.length > LIMITS[k][1] || !l.every(oneLine); }), () => `rows 1..${LIMITS.rows[1]}, columns 0..${LIMITS.columns[1]}, pages 0..${LIMITS.pages[1]} header names`],
  [s => { const all = ['rows', 'columns', 'pages'].flatMap(k => listOf(s[k])); return new Set(all).size !== all.length; }, () => 'a field may sit on one of rows / columns / pages only'],
  [s => !Array.isArray(s.data) || s.data.length < LIMITS.data[0] || s.data.length > LIMITS.data[1] || !s.data.every(d => d && typeof d === 'object' && !Array.isArray(d) && oneLine(d.field) && FUNCTIONS.includes(d.function) && Object.keys(d).every(k => k === 'field' || k === 'function')), () => `data must be 1..${LIMITS.data[1]} { field, function } with function one of ${FUNCTIONS.join(', ')}`],
  [s => new Set(s.data.map(d => `${d.field}\u0000${d.function}`)).size !== s.data.length, () => 'the same field with the same function twice'],
  [s => s.layout !== undefined && !LAYOUTS.includes(s.layout), () => `layout must be one of ${LAYOUTS.join(', ')}`],
  [s => s.style !== undefined && !(typeof s.style === 'string' && STYLE.test(s.style)), () => 'style must be a built-in pivot style name (PivotStyleLight1..28 / Medium1..28 / Dark1..28)'],
  [s => { const g = parseRange(s.range); return !g || g.r2 - g.r1 < 1 || g.r2 - g.r1 > MAX_DATA_ROWS || g.c2 - g.c1 < 1; }, () => `range must be a plain A1 range: a header row, 1..${MAX_DATA_ROWS} data rows, at least two columns`],
  [s => { const a = parseRange(s.at); return !a || a.count !== 1; }, () => 'at must be one cell'],
]);

function plan(step, index) {
  const at = `step ${index}: create_pivot`;
  const hit = PLAN_CHECKS.find(([bad]) => bad(step));
  if (hit) return { ok: false, detail: `${at}: ${hit[1](step)}` };
  const g = parseRange(step.range);
  const a = parseRange(step.at);
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`,
      intent: 'create_pivot',
      family: 'pivots',
      v2: true,
      selector: Object.freeze({ sheet: step.sheet, range: rangeText(g), pivot: step.name }),
      name: step.name,
      pivotSheet: step.pivotSheet,
      grid: Object.freeze({ c1: g.c1, r1: g.r1, c2: g.c2, r2: g.r2 }),
      at: Object.freeze({ c: a.c1, r: a.r1 }),
      spec: Object.freeze({ rows: listOf(step.rows), columns: listOf(step.columns), pages: listOf(step.pages), data: step.data.map(d => ({ field: d.field, function: d.function })),
        ...(step.layout !== undefined ? { layout: step.layout } : {}), ...(step.style !== undefined ? { style: step.style } : {}) }),
      cells: Object.freeze([]),
      preservationMode: 'PRIMITIVE_INTEGRAL',
    },
  };
}

const sourceText = op => `${op.selector.sheet}!$${colName(op.grid.c1)}$${op.grid.r1}:$${colName(op.grid.c2)}$${op.grid.r2}`;

/** the header row as text, or the reason it is not unique text holding every named field */
function header(p, op) {
  const out = [];
  for (let c = op.grid.c1; c <= op.grid.c2; c += 1) {
    const cell = cellAt(p, op.selector.sheet, addr(c, op.grid.r1));
    if (!cell || cell.f !== '' || cell.count === null) return { why: 'xlsx-pivot-source-shape-not-measured', cell: addr(c, op.grid.r1) };
    out.push(cell.v === null || cell.v === undefined ? '' : String(cell.v));
  }
  if (out.some(f => f === '') || new Set(out).size !== out.length) return { why: 'xlsx-pivot-source-header-not-unique-text' };
  const named = [...op.spec.rows, ...op.spec.columns, ...op.spec.pages, ...op.spec.data.map(d => d.field)];
  const missing = named.filter(f => !out.includes(f));
  return missing.length ? { why: 'xlsx-pivot-field-not-in-header', fields: missing } : { fields: out };
}

/** every pivot of the workbook with its host, from the per-pivot READ rows (G2) where the axis is a gap; unread sheets named */
function pivotsOf(p) {
  const out = [];
  const gaps = [];
  for (const s of p.sheets) {
    const o = p.objects && p.objects[s.name];
    if (o && Array.isArray(o.pivotDetails)) o.pivotDetails.forEach(d => out.push({ host: s.name, ...d }));
    else if (o && Array.isArray(o.pivots)) o.pivots.forEach(d => out.push({ host: s.name, ...d }));
    else gaps.push({ sheet: s.name, why: o ? o.pivots : 'objects-not-read' });
  }
  return { pivots: out, gaps };
}

function gate(p, op) {
  if (!p || p.format !== 'xlsx') return { ok: false, reason: 'xlsx-observation-missing' };
  const missing = [op.selector.sheet, op.pivotSheet].filter(n => p.sheets.filter(x => x.name === n).length !== 1);
  if (!p.complete) return { ok: false, reason: p.truncated ? 'xlsx-observation-truncated' : 'xlsx-observation-unreadable-getters', unsupported: p.unsupported };
  if (missing.length) return { ok: false, reason: 'xlsx-sheet-not-found', sheets: missing };
  const h = header(p, op);
  if (h.why) return { ok: false, reason: h.why, ...h };
  const all = pivotsOf(p);
  if (all.gaps.length) return { ok: false, reason: 'xlsx-pivot-observation-incomplete', gaps: all.gaps.slice(0, 20) };
  // empty = no NON-BLANK cell in the pivotSheet's read: a new, empty sheet reports used 'A1' (release LIVE run 1), not null
  const prefix = `${op.pivotSheet}\u0000`;
  const taken = Object.entries(p.cells || {}).filter(([k, c]) => k.startsWith(prefix) && c && !(c.f === '' && (c.v === null || c.v === '' || c.v === undefined))).map(([k]) => k.slice(prefix.length));
  return { ok: true, hits: all.pivots.filter(x => x.name === op.name), hostTaken: taken };
}

/** the whole identity, or the first field that differs */
function identityGap(op, pv) {
  const s = op.spec;
  const layout = pv.layout || {};
  const sourceOk = typeof pv.source === 'string' && !pv.source.startsWith('UNOBSERVABLE') ? bareRef(pv.source) === bareRef(sourceText(op)) : bareRef(pv.dataRef) === bareRef(sourceText(op));
  const checks = [
    ['host', pv.host === op.pivotSheet],
    ['source', sourceOk],
    ['rows', same(layout.rows, s.rows)],
    ['columns', same(layout.cols, s.columns)],
    ['pages', same(pv.pages, s.pages)],
    ['data', Array.isArray(pv.dataFields) && same(pv.dataFields.map(d => ({ field: d.field, function: d.function })), s.data)],
    ['layout', s.layout === undefined || pv.rowLayout === s.layout],
    ['style', s.style === undefined || pv.style === s.style],
  ];
  const bad = checks.find(([, ok]) => !ok);
  return bad ? bad[0] : null;
}

function classify(p, op) {
  const g = gate(p, op);
  if (!g.ok) return { state: 'FAIL_CLOSED', evidence: g };
  const where = { pivot: op.name, source: sourceText(op), pivotSheet: op.pivotSheet };
  if (g.hits.length > 1) return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-pivot-name-ambiguous', ...where, count: g.hits.length } };
  if (g.hits.length === 1) {
    const gap = identityGap(op, g.hits[0]);
    return gap ? { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-pivot-exists-with-another-identity', differs: gap, ...where } }
      : { state: 'SATISFIED', evidence: { ...where, proof: 'fresh-live-pivot-with-the-whole-identity' } };
  }
  if (g.hostTaken.length) return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-pivot-v2-destination-sheet-not-empty', cells: g.hostTaken.slice(0, 20), count: g.hostTaken.length, ...where } };
  return { state: 'APPLY', evidence: where };
}

function resolve(p, selector, ctx = {}) {
  const g = gate(p, ctx.operation);
  if (!g.ok) return { ok: false, target: null, selector, ambiguity: g.reason, detail: g };
  return { ok: true, target: Object.freeze({ sheet: selector.sheet, pivot: selector.pivot }) };
}

function lower(op) {
  const J = JSON.stringify;
  const g = op.grid;
  const s = op.spec;
  const fields = { rows: s.rows, ...(s.columns.length ? { columns: s.columns } : {}), ...(s.pages.length ? { pages: s.pages } : {}) };
  return {
    kind: 'callCommand',
    body: `
try {
  const src = Api.GetSheet(${J(op.selector.sheet)});
  const host = Api.GetSheet(${J(op.pivotSheet)});
  if (!src || src.GetName() !== ${J(op.selector.sheet)} || !host || host.GetName() !== ${J(op.pivotSheet)}) return { ok: false, stage: 'sheet-identity' };
  const nm = x => { try { return x.GetName(); } catch (e) { return null; } };
  if (Api.GetSheets().some(w => (w.GetAllPivotTables() || []).some(x => nm(x) === ${J(op.name)}))) return { ok: false, stage: 'name-taken' };
  // re-checked at dispatch: v2 writes only onto an EMPTY sheet (a page area / a shifted table cannot be pre-computed)
  const used = host.GetUsedRange();
  const now = used ? used.GetValue() : '';
  if ((Array.isArray(now) ? now.flat() : [now]).some(v => v !== '' && v !== null && v !== undefined)) return { ok: false, stage: 'destination-not-empty' };
  const p = Api.InsertPivotExistingWorksheet(src.GetRange(${J(`${colName(g.c1)}${g.r1}:${colName(g.c2)}${g.r2}`)}), host.GetRange(${J(addr(op.at.c, op.at.r))}));
  if (!p) return { ok: false, stage: 'insert-returned-nothing' };
  p.SetName(${J(op.name)});
  p.AddFields(${J(fields)});
  for (const d of ${J(s.data)}) { const f = p.AddDataField(d.field); if (!f) return { ok: false, stage: 'data-field-not-added', field: d.field }; if (d.function !== 'Sum') f.SetFunction(d.function); }
${s.layout ? `  p.SetRowAxisLayout(${J(s.layout)}, false);\n` : ''}${s.style ? `  p.SetStyleName(${J(s.style)});\n` : ''}  p.RefreshTable();
  return { ok: true };
} catch (err) { return { ok: false, stage: 'threw', error: String(err) }; }
`,
    meta: Object.freeze({ family: 'pivots', action: 'create_pivot', v2: true, name: op.name, sheet: op.selector.sheet, range: op.selector.range, pivotSheet: op.pivotSheet, at: addr(op.at.c, op.at.r), spec: s }),
  };
}

function verifyPrimitive(op, _before, after) {
  const g = gate(after && after.primary, op);
  if (!g.ok) return { ok: false, evidence: g };
  const pv = g.hits.length === 1 ? g.hits[0] : null;
  const gap = pv ? identityGap(op, pv) : 'no-single-pivot';
  return { ok: !gap, evidence: { pivot: op.name, differs: gap, now: pv ? { host: pv.host, source: pv.source, layout: pv.layout, pages: pv.pages, dataFields: pv.dataFields, rowLayout: pv.rowLayout, style: pv.style } : g.hits.length, want: { source: sourceText(op), ...op.spec }, proof: 'fresh-live-pivot-getters' } };
}

/** L3: every cell outside the (empty before) pivotSheet -- release LIVE run 2: the table's cells there 'appeared' -- and every
 * sheet's charts / names and every OTHER pivot */
function verifyPreservation(op, before, after) {
  // the pivotSheet was EMPTY before the write (refused otherwise): every cell on it now is the pivot's own output
  const prefix = `${op.pivotSheet}\u0000`;
  const own = [...new Set([before, after].flatMap(x => Object.keys((x && x.primary && x.primary.cells) || {})).filter(k => k.startsWith(prefix)))].map(k => ({ a: k.slice(prefix.length) }));
  const rest = cells.verifyPreservation({ ...op, selector: { ...op.selector, sheet: op.pivotSheet }, cells: own }, before, after);
  if (!rest.ok && rest.evidence && rest.evidence.reason) return rest;
  const others = p => JSON.stringify(Object.fromEntries(Object.entries(p.objects || {}).map(([s, o]) => [s, o && typeof o === 'object' ? { charts: o.charts, names: o.names, pivots: Array.isArray(o.pivotDetails) ? o.pivotDetails.filter(x => x.name !== op.name).map(x => without(x, ['source', 'dataRef'])) : (Array.isArray(o.pivots) ? o.pivots.filter(x => x.name !== op.name) : o.pivots) } : o])));
  const kept = others(before.primary) === others(after.primary);
  return { ok: rest.ok && kept, evidence: { ...rest.evidence, otherObjectsKept: kept } };
}

function wholeTaskChecks(op, finalObservation) {
  const v = verifyPrimitive(op, null, finalObservation);
  return [{ name: `pivot ${op.name} over ${sourceText(op)} on ${op.pivotSheet} (v2)`, ok: v.ok, detail: v.ok ? null : v.evidence }];
}

/** the source's header row (the field names); the pivotSheet's emptiness is the sheet list's own `used` */
const windowsOf = op => [{ sheet: op.selector.sheet, range: rangeText({ c1: op.grid.c1, r1: op.grid.r1, c2: op.grid.c2, r2: op.grid.r1 }) }];
const liveExtras = () => ({ objectSheets: ['*'] });

module.exports = { isV2, FUNCTIONS, LAYOUTS, STYLE, plan, classify, resolve, lower, verifyPrimitive, verifyPreservation, wholeTaskChecks, windowsOf, liveExtras, identityGap, sourceText };
