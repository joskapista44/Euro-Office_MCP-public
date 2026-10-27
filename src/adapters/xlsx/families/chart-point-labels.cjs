'use strict';

const { sheetNameProblem } = require('../sheet-name.cjs');
const cells = require('./cells.cjs');
const drawings = require('../drawings.cjs');

/**
 * XLSX `chart-point-labels` family: set_chart_point_data_label (GH #302 G5) -- the data-label flags of ONE point of ONE
 * series of one existing, named chart, through the public ApiChart.SetShowPointDataLabel(series, point, serName,
 * catName, val, percent). The READ mirror and the verifier is read_charts' effectiveLabels['<series>:<point>']
 * (chart.dataLabels.effective.read/v1, the point -> series -> chart precedence already resolved).
 * The G5 LIVE design probe (docs/evidence/gh302-g5-point-data-label-probe-20261026.json), two runs:
 *   - every call was effective in session, after a fresh reopen AND in the saved <c:dLbl>: value only, category + value,
 *     a point turned ALL OFF again (it reads level POINT, every flag false), percent on a pie point
 *   - only the requested point changes: the other points of a bar keep level SERIES; an untouched pie reads every point
 *     level POINT, all false already (the pie's own default, the untouched pie control) -- so the set is per point
 *   - the chart-level GetDataLabels() reads all false before and after (the point write does not touch it)
 * The four flags are all required (the setter takes them positionally -- a partial request would have to guess the
 * others). showPercent true only on a chart that reads type 'pie' (the only measured one; refused named elsewhere).
 * SATISFIED: the point already shows exactly these four flags (the EFFECTIVE state, whatever level it comes from).
 * Refused before the write (named): the chart read incomplete, no chart of the name, its effectiveLabels not observable
 * (over the point cap, an unknown point count, the bridge capability missing), no such series/point.
 * L1: the point's four effective flags. L3: every cell, every drawing, every chart of every sheet exactly as read --
 * this chart with only this point's effective label changed.
 */

const INTENTS = Object.freeze(['set_chart_point_data_label']);
const NAME = /^[A-Za-z_][A-Za-z0-9_. -]{0,99}$/;
const FIELDS = Object.freeze(['intent', 'id', 'sheet', 'name', 'series', 'point', 'labels']);
const LABEL_KEYS = Object.freeze(['showSerName', 'showCatName', 'showVal', 'showPercent']);
const PERCENT_TYPES = Object.freeze(['pie']);

const keyOf = op => `${op.series}:${op.point}`;
const flagsOf = e => (e !== null && typeof e === 'object' && LABEL_KEYS.every(k => typeof e[k] === 'boolean') ? Object.fromEntries(LABEL_KEYS.map(k => [k, e[k]])) : null);
const sameFlags = (a, b) => !!a && LABEL_KEYS.every(k => a[k] === b[k]);

function plan(step, index) {
  const at = `step ${index}: set_chart_point_data_label`;
  const unknown = Object.keys(step).filter(k => !FIELDS.includes(k));
  if (unknown.length) return { ok: false, detail: `${at} does not support field(s): ${unknown.join(', ')}` };
  const sheetWhy = sheetNameProblem(step.sheet);
  if (sheetWhy) return { ok: false, detail: `${at}: ${sheetWhy}` };
  if (typeof step.name !== 'string' || !NAME.test(step.name)) return { ok: false, detail: `${at}: name must be a chart name (letters, digits, _ . - space; starting with a letter or _)` };
  for (const k of ['series', 'point']) if (!Number.isInteger(step[k]) || step[k] < 0) return { ok: false, detail: `${at}: ${k} must be a 0-based index (read_charts effectiveLabels '<series>:<point>')` };
  const l = step.labels;
  const extra = l && typeof l === 'object' && !Array.isArray(l) ? Object.keys(l).filter(k => !LABEL_KEYS.includes(k)) : null;
  if (!extra || extra.length || !LABEL_KEYS.every(k => typeof l[k] === 'boolean')) return { ok: false, detail: `${at}: labels must give all four booleans ${LABEL_KEYS.join(', ')} and nothing else` };
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`,
      intent: 'set_chart_point_data_label',
      family: 'chart-point-labels',
      selector: Object.freeze({ sheet: step.sheet, cellFree: true, claims: Object.freeze([`chart:${step.sheet}:${step.name}`]) }),
      name: step.name,
      series: step.series,
      point: step.point,
      labels: Object.freeze(Object.fromEntries(LABEL_KEYS.map(k => [k, l[k]]))),
      cells: Object.freeze([]),
      preservationMode: 'PRIMITIVE_INTEGRAL',
    },
  };
}

/** the one chart of the name and the point's effective label, or the named reason */
function gate(p, op) {
  if (!p || p.format !== 'xlsx') return { ok: false, reason: 'xlsx-observation-missing' };
  if (!p.complete) return { ok: false, reason: p.truncated ? 'xlsx-observation-truncated' : 'xlsx-observation-unreadable-getters', unsupported: p.unsupported };
  if (p.sheets.filter(x => x.name === op.selector.sheet).length !== 1) return { ok: false, reason: 'xlsx-sheet-not-found', available: p.sheets.map(x => x.name) };
  const r = p.chartReads && p.chartReads[op.selector.sheet];
  if (!r || typeof r !== 'object' || !Array.isArray(r.charts)) return { ok: false, reason: 'xlsx-chart-unobservable', why: typeof r === 'string' ? r : 'not-read' };
  const hits = r.charts.filter(c => c.name === op.name);
  if (hits.length !== 1) return { ok: false, reason: 'xlsx-chart-point-label-no-such-chart', sheet: op.selector.sheet, name: op.name, count: hits.length };
  const chart = hits[0];
  if (chart.effectiveLabels === null || typeof chart.effectiveLabels !== 'object') return { ok: false, reason: 'xlsx-chart-point-label-unobservable', why: chart.effectiveLabels };
  if (!Object.hasOwn(chart.effectiveLabels, keyOf(op))) return { ok: false, reason: 'xlsx-chart-point-label-no-such-point', point: keyOf(op), points: Object.keys(chart.effectiveLabels).length };
  const now = flagsOf(chart.effectiveLabels[keyOf(op)]);
  if (!now) return { ok: false, reason: 'xlsx-chart-point-label-unobservable', why: chart.effectiveLabels[keyOf(op)] };
  return { ok: true, chart, now };
}

function classify(p, op) {
  const g = gate(p, op);
  if (!g.ok) return { state: 'FAIL_CLOSED', evidence: g };
  const where = { sheet: op.selector.sheet, name: op.name, point: keyOf(op) };
  if (sameFlags(g.now, op.labels)) return { state: 'SATISFIED', evidence: { ...where, proof: 'fresh-live-effective-label-reads-these-flags' } };
  if (op.labels.showPercent && !PERCENT_TYPES.includes(g.chart.type)) return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-chart-point-label-percent-not-measured-on-type', type: g.chart.type, measured: PERCENT_TYPES } };
  const drawn = drawings.observe(p.drawingReads, p.sheets.map(x => x.name));
  if (!drawn.complete) return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-chart-point-label-drawings-unobservable', unobservable: drawn.gaps.slice(0, 20), count: drawn.gaps.length } };
  return { state: 'APPLY', evidence: where };
}

function resolve(p, selector, ctx = {}) {
  const g = gate(p, ctx.operation);
  if (!g.ok) return { ok: false, target: null, selector, ambiguity: g.reason, detail: g };
  return { ok: true, target: Object.freeze({ sheet: selector.sheet, name: ctx.operation.name, point: keyOf(ctx.operation) }) };
}

/** One callCommand: exact-name sheet, exactly one chart of the name holding the series, the point's four flags. */
function lower(op) {
  const J = JSON.stringify;
  const f = op.labels;
  return {
    kind: 'callCommand',
    body: `
try {
  const ws = Api.GetSheet(${J(op.selector.sheet)});
  if (!ws || ws.GetName() !== ${J(op.selector.sheet)}) return { ok: false, stage: 'sheet-identity' };
  const hit = (ws.GetAllCharts() || []).filter(c => { try { return c.GetName() === ${J(op.name)}; } catch (e) { return false; } });
  if (hit.length !== 1) return { ok: false, stage: 'chart-identity', charts: hit.length };
  if ((hit[0].GetAllSeries() || []).length <= ${op.series}) return { ok: false, stage: 'series-identity' };
  if (hit[0].SetShowPointDataLabel(${op.series}, ${op.point}, ${f.showSerName}, ${f.showCatName}, ${f.showVal}, ${f.showPercent}) === false) return { ok: false, stage: 'setter-rejected' };
  return { ok: true };
} catch (err) { return { ok: false, stage: 'threw', error: String(err) }; }
`,
    meta: Object.freeze({ family: 'chart-point-labels', sheet: op.selector.sheet, name: op.name, series: op.series, point: op.point, labels: op.labels }),
  };
}

function verifyPrimitive(op, before, after) {
  const g = gate(after && after.primary, op);
  if (!g.ok) return { ok: false, evidence: { reason: g.reason, why: g.why || null } };
  return { ok: sameFlags(g.now, op.labels), evidence: { sheet: op.selector.sheet, name: op.name, point: keyOf(op), now: g.now, want: op.labels, proof: 'fresh-live-chart.dataLabels.effective.read' } };
}

/** every chart of every sheet as read, with this point's effective label taken out of this chart */
function chartsWithoutPoint(reads, op) {
  if (!reads) return null;
  return Object.entries(reads).map(([s, r]) => [s, r && Array.isArray(r.charts) ? r.charts.map(c => {
    if (s !== op.selector.sheet || c.name !== op.name || !c.effectiveLabels || typeof c.effectiveLabels !== 'object') return c;
    return { ...c, effectiveLabels: Object.fromEntries(Object.entries(c.effectiveLabels).filter(([k]) => k !== keyOf(op))) };
  }) : r, r && r.pivots]);
}

/** L3: every cell; every drawing; every chart field of every sheet -- this chart with only this point's label changed */
function verifyPreservation(op, before, after) {
  const rest = cells.verifyPreservation(op, before, after);
  if (!rest.ok && rest.evidence && rest.evidence.reason) return rest;
  const b = chartsWithoutPoint(before.primary.chartReads, op);
  const a = chartsWithoutPoint(after.primary.chartReads, op);
  const chartsKept = !!b && !!a && JSON.stringify(b) === JSON.stringify(a);
  const draw = drawings.drawingDiffs(
    drawings.observe(before.primary.drawingReads, before.primary.sheets.map(x => x.name)),
    drawings.observe(after.primary.drawingReads, after.primary.sheets.map(x => x.name)),
    {},
  );
  return { ok: rest.ok && chartsKept && draw.length === 0, evidence: { ...rest.evidence, chartsKept, drawingsKept: draw.length === 0, drawingDiffs: draw } };
}

function wholeTaskChecks(op, finalObservation) {
  const v = verifyPrimitive(op, null, finalObservation);
  return [{ name: `${op.selector.sheet}: chart ${op.name} point ${keyOf(op)} data label`, ok: v.ok, detail: v.ok ? null : v.evidence }];
}

const windowsOf = () => [];
const liveExtras = () => ({ chartSheets: ['*'], drawingSheets: ['*'] });

module.exports = { INTENTS, LABEL_KEYS, PERCENT_TYPES, plan, classify, resolve, lower, verifyPrimitive, verifyPreservation, wholeTaskChecks, windowsOf, liveExtras };
