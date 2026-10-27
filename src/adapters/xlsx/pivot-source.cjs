'use strict';

const { XLSX_RUNTIME_IDS } = require('./runtime.cjs');

/**
 * M7 tier H P-A / P9 -- a pivot's SOURCE after a reopen, proven by content (owner GO 2026-10-03, owner channel).
 *
 * The public `GetSource()` throws after any reopen (W3g, P-A probe). The data reference, read through the DS
 * Bridge `pivot.state.read/v1` (B13, Wave B READ /4; formerly the internal pivot's data ref), names a sheet and a range, but after a source-sheet rename it still names the OLD sheet,
 * and a new sheet of that old name then answers for it (P-A2 probe, docs/evidence/M7-H2-PROBE.md). So the name alone
 * is never accepted (owner: "a lapnév feloldhatósága önmagában nem bizonyítja az azonosságot"). The source is proven
 * only when ALL hold:
 *   1. the runtime is the one whose Bridge carries pivot.state.read/v1 (PIVOT_STATE_BRIDGE_RUNTIMES) and the reader
 *      answered every part (a Bridge refusal reaches here as the proof's own 'UNOBSERVABLE:<reason>');
 *   2. the layout is the measured one: one row field, one column field, one data field `Sum of <field>`, the source's
 *      header row equal to the pivot's field names, every grand total present;
 *   3. the pivot's cached body (row x column sums, row totals, column totals, grand total) is reproduced EXACTLY from
 *      the named range's values, with the row / column labels being exactly the source's distinct keys;
 *   4. NO other sheet reproduces the body from the same address -- a renamed true source, a copy of the source or an
 *      impostor with the same content makes it ambiguous.
 * Anything else is 'UNOBSERVABLE:<why>' -- never an empty dependency. A stale cache (the source edited after the last
 * refresh) fails 3 and stays unobservable: a refusal, not a guess.
 */

const { aggregate, bodyMismatch } = require('./pivot-aggregate.cjs');

// The runtime identity whose DS Bridge advertises pivot.state.read/v1 (patch family /4) -- not an internal-model gate.
const PIVOT_STATE_BRIDGE_RUNTIMES = XLSX_RUNTIME_IDS;

const MAX_SOURCE_CELLS = 10000;
const REF = /^(?:'((?:[^']|'')+)'|([^'!]+))!\$?([A-Z]{1,3})\$?([0-9]{1,7}):\$?([A-Z]{1,3})\$?([0-9]{1,7})$/;
const BARE = /^\$?([A-Z]{1,3})\$?([0-9]{1,7}):\$?([A-Z]{1,3})\$?([0-9]{1,7})$/;

const colNum = s => [...s].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);
const box = (c1, r1, c2, r2) => ({ c1: colNum(c1), r1: Number(r1), c2: colNum(c2), r2: Number(r2) });
const fail = why => `UNOBSERVABLE:${why}`;

/** `Sheet!$A$1:$C$7` / `'My sheet'!A1:C7` -> { sheet, range: 'A1:C7', box }; null when another shape. */
function parseRef(text) {
  const m = REF.exec(String(text || ''));
  if (!m) return null;
  const sheet = m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2];
  return { sheet, range: `${m[3]}${m[4]}:${m[5]}${m[6]}`, box: box(m[3], m[4], m[5], m[6]) };
}

/** `I5:K7` (unqualified) -> a box; null when another shape. */
function parseBare(text) {
  const m = BARE.exec(String(text || ''));
  return m ? box(m[1], m[2], m[3], m[4]) : null;
}

const quoted = s => (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(s) ? s : `'${s.replace(/'/g, "''")}'`);
const absolute = r => r.replace(/([A-Z]{1,3})([0-9]{1,7})/g, '$$$1$$$2');

const isMatrix = (v, rows, cols) => Array.isArray(v) && v.length === rows && v.every(r => Array.isArray(r) && r.length === cols);

/** The measured layout (1 row field, 1 column field, `Sum of <field>`) -> { row, col, data } or { why }. */
function layoutOf({ fields, rows, cols, data }) {
  if (![fields, rows, cols, data].every(a => Array.isArray(a) && a.every(x => typeof x === 'string'))) return { why: 'pivot-fields-unread' };
  const m = /^Sum of (.+)$/.exec(data.length === 1 ? data[0] : '');
  if (!m || rows.length !== 1 || cols.length !== 1) return { why: 'pivot-layout-not-measured' };
  const used = [rows[0], cols[0], m[1]];
  const distinct = new Set(fields).size === fields.length && new Set(used).size === 3;
  return distinct && used.every(f => fields.includes(f)) ? { row: rows[0], col: cols[0], data: m[1] } : { why: 'pivot-layout-not-measured' };
}

/** The body sits inside the table with a label row above it and a label column left of it. */
const labelledInside = (bb, tb) => bb.r1 > tb.r1 && bb.c1 > tb.c1 && bb.r2 <= tb.r2 && bb.c2 <= tb.c2;

/** The cached body, its table and the labels around it -> { body, rowLabels, colLabels } or { why }. */
function extentOf(proof) {
  const bb = parseBare(proof.body && proof.body.range);
  const tb = parseBare(proof.table && proof.table.range);
  if (!bb || !tb || !labelledInside(bb, tb)) return { why: 'pivot-extent-shape' };
  const [nr, nc] = [bb.r2 - bb.r1 + 1, bb.c2 - bb.c1 + 1];
  if (nr < 2 || nc < 2 || !isMatrix(proof.body.values, nr, nc) || !isMatrix(proof.table.values, tb.r2 - tb.r1 + 1, tb.c2 - tb.c1 + 1)) return { why: 'pivot-body-unread' };
  const tv = proof.table.values;
  const body = proof.body.values;
  return { body, rowLabels: body.map((_, i) => tv[bb.r1 - tb.r1 + i][bb.c1 - 1 - tb.c1]), colLabels: body[0].map((_, j) => tv[bb.r1 - 1 - tb.r1][bb.c1 - tb.c1 + j]) };
}

/** The data reference and the source read against it -> { ref } or { why }. */
function sourceOf(proof, sheetNames, width) {
  const ref = parseRef(proof.dataRef);
  if (!ref) return { why: 'pivot-dataref-shape' };
  if (!sheetNames.includes(ref.sheet)) return { why: `pivot-source-sheet-missing:${ref.sheet}` };
  const src = proof.source;
  if (!src || src.sheet !== ref.sheet || src.range !== ref.range) return { why: 'pivot-source-read-differs-from-dataref' };
  const b = ref.box;
  if ((b.r2 - b.r1 + 1) * (b.c2 - b.c1 + 1) > MAX_SOURCE_CELLS) return { why: 'pivot-source-too-large' };
  return b.c2 - b.c1 + 1 === width ? { ref } : { why: 'pivot-source-width-differs' };
}

/** 4. every OTHER sheet read at the same address; none may reproduce the body -> null or the reason. */
function ambiguity(proof, ref, sheetNames, reproduces) {
  const others = Array.isArray(proof.others) ? proof.others : null;
  const expected = sheetNames.filter(n => n !== ref.sheet);
  if (!others || others.length !== expected.length || !expected.every(n => others.some(o => o && o.sheet === n))) return 'pivot-candidates-incomplete';
  for (const o of others) {
    if (!Array.isArray(o.values)) return `pivot-candidate-unread:${o.sheet}`;
    if (reproduces(o.values)) return `pivot-source-ambiguous:${o.sheet}`;
  }
  return null;
}

/**
 * The proven source of one pivot, or 'UNOBSERVABLE:<why>'.
 * `proof` = the driver's raw read (or its own 'UNOBSERVABLE:<why>'): { dataRef, fields, rows, cols, data,
 * body: { range, values }, table: { range, values }, source: { sheet, range, values }, others: [{ sheet, values }] };
 * `runtime` = the session's DS runtime identity (core/runtime-identity.cjs); `sheetNames` = every sheet of the workbook.
 */
function provenSource(proof, runtime, sheetNames) {
  if (!PIVOT_STATE_BRIDGE_RUNTIMES.includes(runtime)) return fail('pivot-runtime-not-verified');
  if (typeof proof === 'string' || !proof || typeof proof !== 'object') return typeof proof === 'string' && proof.startsWith('UNOBSERVABLE') ? proof : fail('pivot-proof-shape');
  const lay = layoutOf(proof);
  const s = lay.why ? lay : sourceOf(proof, sheetNames, proof.fields.length);
  const ext = s.why ? s : extentOf(proof);
  if (ext.why) return fail(ext.why);
  const check = values => { const agg = aggregate(values, proof.fields, lay.row, lay.col, lay.data); return agg.why || bodyMismatch(agg, ext.body, ext.rowLabels, ext.colLabels); };
  const bad = check(proof.source.values) || ambiguity(proof, s.ref, sheetNames, values => !check(values));
  return bad ? fail(bad) : `${quoted(s.ref.sheet)}!${absolute(s.ref.range)}`;
}

/**
 * Every pivot of a read's objects with its source: the public one when it answered, else the content-proven one; the
 * driver's raw proof is not kept. If ANY pivot of a sheet stays unproven, that sheet's whole pivots axis becomes its
 * reason string -- exactly the shape every consumer refused on before P9 (a per-pivot gap inside a list would pass the
 * consumers that check the axis only: the W3 LIVE regression's copy of a source sheet, m7-h-regr-w3-20261004).
 */
function withProvenPivots(objects, runtime, sheetNames) {
  const pivot = p => {
    const pub = p && typeof p.source === 'string' && !p.source.startsWith('UNOBSERVABLE') ? p.source : null;
    // M7 P1: the table's extent and the layout ride along when the driver read them (create_pivot's identity)
    const extra = p && typeof p.table === 'string' && p.layout && typeof p.layout === 'object' ? { table: p.table, layout: p.layout } : {};
    return { name: p && p.name, source: pub || provenSource(p && p.proof, runtime, sheetNames), ...extra };
  };
  const axis = all => { const gap = all.find(p => p.source.startsWith('UNOBSERVABLE')); return gap ? `${gap.source} (pivot ${gap.name})` : all; };
  // GH #302 G2: the READ projection's per-pivot rows (read_pivots), built only when the driver read the axes. They never feed
  // the pivots AXIS above (the write guards keep refusing on any unproven pivot of a sheet); an unproven source is that ONE
  // row's named `source`, the list stays whole.
  const detail = (p, proven) => {
    const ax = p.axes;
    const fns = Array.isArray(p.functions) ? p.functions : null;
    // GH #302 G3: the row axis layout form (SetRowAxisLayout sets every field's form; the row fields' common form is the
    // answer, a disagreement or a gap is named) and the Bridge dataRef a create_pivot v2 retry compares the source by
    // (read_pivots does not project it: `source` is the public / content-proven reference)
    const forms = Array.isArray(p.fieldLayout) && Array.isArray(ax.rows) ? ax.rows.map(r => (p.fieldLayout.find(f => f.name === r) || {}).form) : null;
    const rowLayout = !forms ? 'UNOBSERVABLE:field-layout-unread' : (forms.length && forms.every(f => typeof f === 'string' && !f.startsWith('UNOBSERVABLE') && f === forms[0]) ? forms[0] : (forms.length ? 'UNOBSERVABLE:row-layout-mixed-or-unread' : null));
    return { name: proven.name, source: proven.source, ...(typeof p.table === 'string' ? { table: p.table } : {}),
      layout: { rows: ax.rows, cols: ax.cols, data: fns ? fns.map(f => f.name) : p.functions },
      pages: ax.pages, values: ax.values, dataFields: fns || p.functions, rowLayout, style: p.style, dataRef: p.dataRef };
  };
  const sheet = o => {
    const all = o.pivots.map(pivot);
    const withAxes = o.pivots.every(p => p && p.axes && typeof p.axes === 'object');
    return { ...o, pivots: axis(all), ...(withAxes && o.pivots.length ? { pivotDetails: o.pivots.map((p, i) => detail(p, all[i])) } : {}) };
  };
  return Object.fromEntries(Object.entries(objects || {}).map(([n, o]) => [n, o && Array.isArray(o.pivots) ? sheet(o) : o]));
}

module.exports = { PIVOT_STATE_BRIDGE_RUNTIMES, MAX_SOURCE_CELLS, provenSource, withProvenPivots, layoutOf, parseRef };
