'use strict';

const { parseRange, rangeText, addr, colName } = require('../a1.cjs');
const { cellAt } = require('../inventory.cjs');
const cells = require('./cells.cjs');
const drawings = require('../drawings.cjs');

/**
 * XLSX `charts` family: set_chart (create) / rename_chart (M6 wave 7).
 *
 * The reader is PUBLIC (the driver's chartReads: every chart's GetName / GetChartType / GetTitle / series name, values
 * and category formulas, plus the pivot NAMES, of every sheet). Audit group C: CH1 / CH18 MIGRATE; CH2 SOURCE_DEFECT --
 * the source never read the data range, so a chart with the same name / type / title / size over ANOTHER range
 * "satisfied" the request; CH4 the size reads 0x0 after a reopen (so the size is not part of the contract); CH20 there
 * is no public chart Delete on this build (W7 probe K5: no Delete / Remove on the chart or the drawing -- delete_chart
 * is not an intent).
 *
 * Contract (the W7 probe, test/live/xlsx-w7-objects-probe.live.cjs, measured the series form; LIVE-accepted):
 *   set_chart { sheet, name, range, type: 'bar' | 'lineNormal' | 'area' | 'pie' | 'doughnut', title } -- the
 *   data shape W3h measured and nothing wider: series in
 *   COLUMNS, the first row text headers (the series names), the first column text categories, a numeric body, a
 *   sheet name that needs no quoting. Identity = type + title + every series' name / values / categories, derived
 *   from the range. SATISFIED: exactly one chart of that name on the sheet with that identity. Refused: a chart of
 *   that name with another identity (no silent modify), more than one, an unmeasured data shape.
 *   rename_chart { sheet, name, newName }: APPLY one `name` and no `newName`; SATISFIED no `name` and one `newName`;
 *   anything else refused. The renamed chart keeps its identity (L1).
 * L1 the chart; L3 every cell, every OTHER chart of every sheet, every pivot name. Retry SATISFIED 0/0.
 *
 * M7 S5 C1 (owner M7 GO 2026-10-03; audit P C1): `orientation: 'rows'` -- the series in ROWS over the same data shape,
 * AddChart(ref, true, ...). MEASURED in session and after a reopen (W7 probe K2, m6-w7-probe-20260929.json): per data
 * row r, name = the row's first cell, values = the row's body, cat = the header row's body. The orientation is part of
 * the identity (it changes every series), so a chart of the same name in the other orientation is refused as another
 * identity -- never silently modified. 'columns' (the default) is unchanged.
 *
 * X-O03 (implementation phase): `type` beyond 'bar' -- LIVE-measured (the
 * X-O03 series-semantics probe) on the SAME multi-column data shape: bar/lineNormal/area/
 * pie/doughnut all hold exactly the bar contract's series description (name/values/cat per
 * column) and persist across a reopen -- these five are open here. `scatter` (its values getter
 * throws internally, unverifiable on the public path) and `stock` (does not persist with this
 * data shape across a reopen) are explicitly NOT opened -- measured failures, not an oversight.
 * The remaining named candidates (stacked/3D/percent variants) are simply NOT YET measured on
 * this data shape; they stay closed until their own LIVE acceptance, per the same rule.
 *
 * X-O04 (implementation phase): `set_chart_legend { sheet, name, legend }` -- LIVE-
 * measured: `SetLegendPos` accepts exactly 'bottom' | 'top' | 'left' | 'right' | 'none', persists
 * across a reopen; an unsupported value (measured: 'topRight', and any invalid string) is
 * REJECTED WITHOUT THROWING -- the setter returns `false` and the position stays unchanged. The
 * write therefore checks the return value explicitly (never assumed from the absence of a throw,
 * the same lesson as W-09's `SetShd`). No creation-time legend field: `set_chart` never measured
 * one, so legend is post-creation-only here, exactly like `set_chart_title` is for the title.
 *
 * X-O05 (LIVE MEASURED; the untouched control chart, same report): `set_chart_axis_title` --
 * SetHorAxisTitle(text) / SetVerAxisTitle(text) (ONE argument, plain text; GetHorAxisTitle/GetVerAxisTitle read it
 * back with a trailing CRLF the same way GetTitle does, trimmed by the SAME `trimTitle`). An untouched chart reads
 * BOTH axis titles as `null` (not a throw), so `trimTitle(null)` (`''`) is the measured "no title" baseline.
 * DATA LABELS ARE NOT A WRITE INTENT (OWNER-GATE / LIMIT / #ref): SetShowDataLabels writes the
 * CHART-level c:dLbls, but every series keeps its own explicit all-0 c:dLbls (written at AddChart), which OOXML lets
 * override the chart level -- the effective (rendered) state is not what GetDataLabels reads, and no public reader
 * exists for the series or point level (ApiSeries has no label method; only ApiChart.SetShowPointDataLabel writes per
 * point). The chart-level labels are still READ (GetDataLabels) so every chart write proves it left them unchanged.
 * Identity: exactly one chart of the name; APPLY when the requested field differs from the fresh read;
 * SATISFIED when it already matches. L1 also proves every OTHER chart field this family tracks (type, title, every
 * series, the other axis title, the data-label flags) stayed put -- the same "seriesKept" pattern set_chart_title
 * already uses, widened to the two new fields.
 */

const INTENTS = Object.freeze(['set_chart', 'rename_chart', 'set_chart_title', 'set_chart_legend', 'set_chart_axis_title']);
const CHART_TYPES = Object.freeze(['bar', 'lineNormal', 'area', 'pie', 'doughnut']);
const LEGEND_POS = Object.freeze(['bottom', 'top', 'left', 'right', 'none']);
const AXES = Object.freeze({ horizontal: 'SetHorAxisTitle', vertical: 'SetVerAxisTitle' });
const LABEL_KEYS = Object.freeze(['showSerName', 'showCatName', 'showVal', 'showPercent']);
const SHEET = /^[A-Za-z_][A-Za-z0-9_]*$/;
const NAME = /^[A-Za-z_][A-Za-z0-9_. -]{0,99}$/;

const oneLine = x => typeof x === 'string' && x !== '' && !/[\r\n]/.test(x);

// the accepted fields per intent (set_chart is the only one that creates and reads cells)
const FIELDS = Object.freeze({
  set_chart: ['intent', 'id', 'sheet', 'name', 'range', 'type', 'title', 'orientation'],
  set_chart_title: ['intent', 'id', 'sheet', 'name', 'title'],
  set_chart_legend: ['intent', 'id', 'sheet', 'name', 'legend'],
  set_chart_axis_title: ['intent', 'id', 'sheet', 'name', 'axis', 'title'],
  rename_chart: ['intent', 'id', 'sheet', 'name', 'newName'],
});

/** The intent-specific checks of plan(): { detail } on a problem, else { g } (set_chart's parsed range, or null). */
function intentProblem(step, at) {
  if (['set_chart', 'set_chart_title', 'set_chart_axis_title'].includes(step.intent) && !oneLine(step.title)) return { detail: `${at}: title must be one non-empty line` };
  if (step.intent === 'set_chart_axis_title' && !Object.hasOwn(AXES, step.axis)) return { detail: `${at}: axis must be 'horizontal' or 'vertical'` };
  if (step.intent === 'set_chart_legend' && !LEGEND_POS.includes(step.legend)) return { detail: `${at}: legend must be one of ${LEGEND_POS.join(', ')}` };
  if (step.intent === 'rename_chart' && (typeof step.newName !== 'string' || !NAME.test(step.newName) || step.newName === step.name)) return { detail: `${at}: newName must be a chart name different from name` };
  if (step.intent !== 'set_chart') return { g: null };
  const g = parseRange(step.range);
  if (!g || g.r2 - g.r1 < 1 || g.c2 - g.c1 < 1) return { detail: `${at} requires a plain A1 range with a header row, a category column and data` };
  if (!CHART_TYPES.includes(step.type)) return { detail: `${at}: type must be one of ${CHART_TYPES.join(', ')}` };
  if ('orientation' in step && !['columns', 'rows'].includes(step.orientation)) return { detail: `${at}: orientation must be 'columns' (default) or 'rows'` };
  return { g };
}

function plan(step, index) {
  const at = `step ${index}: ${step.intent}`;
  const set = step.intent === 'set_chart';
  const retitle = step.intent === 'set_chart_title';
  const relegend = step.intent === 'set_chart_legend';
  const axisTitle = step.intent === 'set_chart_axis_title';
  const cellFree = !set; // every non-creation chart intent touches no cells
  const fields = FIELDS[step.intent] || [];
  const unknown = Object.keys(step).filter(k => !fields.includes(k));
  if (unknown.length) return { ok: false, detail: `${at} does not support field(s): ${unknown.join(', ')}` };
  if (typeof step.sheet !== 'string' || !SHEET.test(step.sheet)) return { ok: false, detail: `${at}: sheet must be a name that needs no quoting in a reference (a quoted sheet reference is not measured yet)` };
  if (typeof step.name !== 'string' || !NAME.test(step.name)) return { ok: false, detail: `${at}: name must be a chart name (letters, digits, _ . - space; starting with a letter or _)` };
  const checked = intentProblem(step, at);
  if (checked.detail) return { ok: false, detail: checked.detail };
  const { g } = checked;
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`,
      intent: step.intent,
      family: 'charts',
      // a chart claims its NAME(s) on the sheet (two steps on one chart name clash); set_chart also reads its data
      // range (a write into it clashes), a rename touches no cells
      selector: Object.freeze({ sheet: step.sheet, range: set ? rangeText(g) : 'A1', chart: step.name, ...(cellFree ? { cellFree: true } : {}),
        claims: Object.freeze([`chart:${step.sheet}:${step.name}`, ...(step.intent === 'rename_chart' ? [`chart:${step.sheet}:${step.newName}`] : [])]) }),
      name: step.name,
      newName: step.intent === 'rename_chart' ? step.newName : null,
      title: set || retitle || axisTitle ? step.title : null,
      legend: relegend ? step.legend : null,
      axis: axisTitle ? step.axis : null,
      type: set ? step.type : null,
      inRows: set && step.orientation === 'rows',
      grid: g ? Object.freeze({ ...g }) : null,
      cells: Object.freeze([]),
      preservationMode: 'PRIMITIVE_INTEGRAL',
    },
  };
}

const bare = x => String(x === null || x === undefined ? '' : x).replace(/^=/, '');
const trimTitle = x => String(x === null || x === undefined ? '' : x).replace(/[\r\n]+$/, '');
const abs = (sheet, c, r) => `${sheet}!$${colName(c)}$${r}`;

/**
 * The series the bar chart over the range has. Columns (W3h): per data column, name = the header, values = the column's
 * body, cat = the first column's body. Rows (W7 K2): per data row, name = the row's first cell, values = the row's body,
 * cat = the header row's body.
 */
function expectedSeries(op) {
  const { sheet } = op.selector;
  const g = op.grid;
  const out = [];
  if (op.inRows) {
    for (let r = g.r1 + 1; r <= g.r2; r += 1) out.push({ name: abs(sheet, g.c1, r), values: `${abs(sheet, g.c1 + 1, r)}:$${colName(g.c2)}$${r}`, cat: `${abs(sheet, g.c1 + 1, g.r1)}:$${colName(g.c2)}$${g.r1}` });
    return out;
  }
  for (let c = g.c1 + 1; c <= g.c2; c += 1) out.push({ name: abs(sheet, c, g.r1), values: `${abs(sheet, c, g.r1 + 1)}:$${colName(c)}$${g.r2}`, cat: `${abs(sheet, g.c1, g.r1 + 1)}:$${colName(g.c1)}$${g.r2}` });
  return out;
}

// X-O05: an untouched chart reads BOTH axis titles null and every data-label flag false (LIVE
// MEASURED, the control chart) -- THAT is the baseline: null -> '' (no title), four booleans as read.
// a maintainer review: anything else (undefined, a missing getter, a non-boolean flag) is UNOBSERVABLE, never defaulted --
// each such read gets its OWN token, so it never compares equal to anything (every identity / kept check fails closed).
let unobservedSeq = 0;
const UNOBSERVED = what => `UNOBSERVABLE:${what}:${unobservedSeq += 1}`;
const isUnobserved = v => typeof v === 'string' && v.startsWith('UNOBSERVABLE:');
const axisTitleOf = t => (t === null ? '' : typeof t === 'string' ? trimTitle(t) : UNOBSERVED('axis-title'));
const labelsOf = l => (l !== null && typeof l === 'object' && LABEL_KEYS.every(k => typeof l[k] === 'boolean')
  ? Object.fromEntries(LABEL_KEYS.map(k => [k, l[k]])) : UNOBSERVED('data-labels'));
const NO_LABELS = Object.freeze(Object.fromEntries(LABEL_KEYS.map(k => [k, false])));
const fingerprint = c => ({
  type: c.type, title: trimTitle(c.title), series: (c.series || []).map(x => ({ name: bare(x.name), values: bare(x.values), cat: bare(x.cat) })),
  horAxisTitle: axisTitleOf(c.horAxisTitle), verAxisTitle: axisTitleOf(c.verAxisTitle), labels: labelsOf(c.dataLabels),
});
const same = (a, b) => JSON.stringify(a) === JSON.stringify(b);
// a freshly-created chart's measured baseline: no axis titles, every data-label flag false (same as above)
const wanted = op => ({ type: op.type, title: op.title, series: expectedSeries(op), horAxisTitle: '', verAxisTitle: '', labels: NO_LABELS });

/** The data shape the contract measured, or why not. */
function shapeProblem(p, op) {
  const { sheet } = op.selector;
  const g = op.grid;
  for (let r = g.r1; r <= g.r2; r += 1) {
    for (let c = g.c1; c <= g.c2; c += 1) {
      const cell = cellAt(p, sheet, addr(c, r));
      if (!cell) return { reason: 'xlsx-chart-cell-unread', cell: addr(c, r) };
      if (r === g.r1 && c === g.c1) continue;
      if (cell.f !== '' || cell.count === null) return { reason: 'xlsx-chart-data-shape-not-measured', cell: addr(c, r), why: 'formula-or-unreadable' };
      const text = cell.count === 0 && cell.v !== null && cell.v !== '';
      if ((r === g.r1 || c === g.c1) ? !text : cell.count !== 1) return { reason: 'xlsx-chart-data-shape-not-measured', cell: addr(c, r), why: r === g.r1 || c === g.c1 ? 'header-or-category-not-text' : 'value-not-numeric' };
    }
  }
  return null;
}

function gate(p, op) {
  if (!p || p.format !== 'xlsx') return { ok: false, reason: 'xlsx-observation-missing' };
  if (!p.complete) return { ok: false, reason: p.truncated ? 'xlsx-observation-truncated' : 'xlsx-observation-unreadable-getters', unsupported: p.unsupported };
  const s = p.sheets.filter(x => x.name === op.selector.sheet);
  if (s.length !== 1) return { ok: false, reason: s.length > 1 ? 'xlsx-sheet-ambiguous' : 'xlsx-sheet-not-found', available: p.sheets.map(x => x.name) };
  const reads = p.chartReads;
  const bad = !reads ? ['not-read'] : p.sheets.map(x => x.name).filter(n => !reads[n] || typeof reads[n] !== 'object' || !Array.isArray(reads[n].charts) || !Array.isArray(reads[n].pivots));
  if (bad.length) return { ok: false, reason: 'xlsx-chart-unobservable', sheets: bad.slice(0, 20), why: reads ? bad.map(n => reads[n]).slice(0, 5) : 'not-read' };
  return { ok: true, charts: reads[op.selector.sheet].charts };
}

const named = (charts, n) => charts.filter(c => c.name === n);

/** M7 S2b: the L3 claims every sheet's drawings -- an unreadable one refuses the write (only APPLY is gated). */
function drawingBlocked(p) {
  const o = drawings.observe(p.drawingReads, p.sheets.map(x => x.name));
  return o.complete ? null : { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-charts-drawings-unobservable', unobservable: o.gaps.slice(0, 20), count: o.gaps.length } };
}

function classify(p, op) {
  const c = classifyChart(p, op);
  return c.state === 'APPLY' ? (drawingBlocked(p) || c) : c;
}

function classifyRename(g, op, where) {
  const from = named(g.charts, op.name).length;
  const to = named(g.charts, op.newName).length;
  if (from === 1 && to === 0) return { state: 'APPLY', evidence: { ...where, newName: op.newName } };
  if (from === 0 && to === 1) return { state: 'SATISFIED', evidence: { ...where, newName: op.newName, proof: 'fresh-live-chart-under-the-new-name' } };
  return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-chart-rename-state-conflict', ...where, newName: op.newName, withName: from, withNewName: to } };
}

function classifyChart(p, op) {
  const g = gate(p, op);
  if (!g.ok) return { state: 'FAIL_CLOSED', evidence: g };
  const where = { sheet: op.selector.sheet, chart: op.name };
  const postCreate = { set_chart_title: classifyTitle, set_chart_legend: classifyLegend, set_chart_axis_title: classifyAxisTitle, rename_chart: classifyRename }[op.intent];
  if (postCreate) return postCreate(g, op, where);
  const shape = shapeProblem(p, op);
  if (shape) return { state: 'FAIL_CLOSED', evidence: shape };
  const hits = named(g.charts, op.name);
  if (hits.length > 1) return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-chart-name-ambiguous', ...where, count: hits.length } };
  if (hits.length === 1) {
    return same(fingerprint(hits[0]), wanted(op))
      ? { state: 'SATISFIED', evidence: { ...where, proof: 'fresh-live-chart-with-the-identity' } }
      : { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-chart-exists-with-another-identity', ...where, now: fingerprint(hits[0]) } };
  }
  return { state: 'APPLY', evidence: { ...where, range: op.selector.range } };
}

/**
 * M7 C3 set_chart_title (owner M7 GO 2026-10-03; tier D probe: SetTitle on an existing chart, the series unchanged, the
 * title read back with a trailing CR LF -- session = reopen): exactly one chart of the name; SATISFIED when its title
 * already reads the requested one.
 */
function classifyTitle(g, op, where) {
  const hits = named(g.charts, op.name);
  if (hits.length !== 1) return { state: 'FAIL_CLOSED', evidence: { reason: hits.length ? 'xlsx-chart-name-ambiguous' : 'xlsx-chart-not-found', ...where, count: hits.length } };
  if (trimTitle(hits[0].title) === op.title) return { state: 'SATISFIED', evidence: { ...where, proof: 'fresh-live-chart-title' } };
  return { state: 'APPLY', evidence: { ...where, title: op.title, now: trimTitle(hits[0].title) } };
}

/**
 * X-O04 set_chart_legend (LIVE-measured): exactly one chart of the name; SATISFIED when its legend
 * position already reads the requested one.
 */
function classifyLegend(g, op, where) {
  const hits = named(g.charts, op.name);
  if (hits.length !== 1) return { state: 'FAIL_CLOSED', evidence: { reason: hits.length ? 'xlsx-chart-name-ambiguous' : 'xlsx-chart-not-found', ...where, count: hits.length } };
  if (hits[0].legend === op.legend) return { state: 'SATISFIED', evidence: { ...where, proof: 'fresh-live-chart-legend' } };
  return { state: 'APPLY', evidence: { ...where, legend: op.legend, now: hits[0].legend } };
}

/** X-O05 (LIVE MEASURED): same shape as classifyTitle, keyed by the requested axis. */
function classifyAxisTitle(g, op, where) {
  const hits = named(g.charts, op.name);
  if (hits.length !== 1) return { state: 'FAIL_CLOSED', evidence: { reason: hits.length ? 'xlsx-chart-name-ambiguous' : 'xlsx-chart-not-found', ...where, count: hits.length } };
  const now = axisTitleOf(op.axis === 'horizontal' ? hits[0].horAxisTitle : hits[0].verAxisTitle);
  if (isUnobserved(now)) return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-chart-axis-title-unobservable', ...where, axis: op.axis } };
  if (now === op.title) return { state: 'SATISFIED', evidence: { ...where, axis: op.axis, proof: 'fresh-live-chart-axis-title' } };
  return { state: 'APPLY', evidence: { ...where, axis: op.axis, title: op.title, now } };
}

function resolve(p, selector, ctx = {}) {
  const g = gate(p, ctx.operation);
  if (!g.ok) return { ok: false, target: null, selector, ambiguity: g.reason, detail: g };
  return { ok: true, target: Object.freeze({ sheet: selector.sheet, chart: selector.chart }) };
}

function lower(op) {
  const { sheet } = op.selector;
  const nameOf = 'c => { try { return c.GetName(); } catch (e) { return null; } }';
  let call;
  if (op.intent === 'set_chart') {
    const g = op.grid;
    const ref = `${sheet}!$${colName(g.c1)}$${g.r1}:$${colName(g.c2)}$${g.r2}`;
    // re-checked at dispatch: a chart of the name that appeared since the read is never doubled
    call = `if ((ws.GetAllCharts() || []).some(c => (${nameOf})(c) === ${JSON.stringify(op.name)})) return { ok: false, stage: 'name-taken' };
  const c = ws.AddChart(${JSON.stringify(ref)}, ${op.inRows ? 'true' : 'false'}, ${JSON.stringify(op.type)}, 2, 3600000, 2160000, ${g.c2}, 0, ${g.r1 - 1}, 0);
  if (!c) return { ok: false, stage: 'add-returned-nothing' };
  c.SetName(${JSON.stringify(op.name)});
  c.SetTitle(${JSON.stringify(op.title)}, 12, false);`;
  } else if (op.intent === 'set_chart_title') {
    call = `const hit = (ws.GetAllCharts() || []).filter(c => (${nameOf})(c) === ${JSON.stringify(op.name)});
  if (hit.length !== 1) return { ok: false, stage: 'chart-identity' };
  hit[0].SetTitle(${JSON.stringify(op.title)}, 12, false);`;
  } else if (op.intent === 'set_chart_legend') {
    // measured: an unsupported position is REJECTED WITHOUT THROWING (the setter returns false,
    // the position stays unchanged) -- the return value is checked explicitly, never assumed.
    call = `const hit = (ws.GetAllCharts() || []).filter(c => (${nameOf})(c) === ${JSON.stringify(op.name)});
  if (hit.length !== 1) return { ok: false, stage: 'chart-identity' };
  if (hit[0].SetLegendPos(${JSON.stringify(op.legend)}) === false) return { ok: false, stage: 'legend-rejected' };`;
  } else if (op.intent === 'set_chart_axis_title') {
    call = `const hit = (ws.GetAllCharts() || []).filter(c => (${nameOf})(c) === ${JSON.stringify(op.name)});
  if (hit.length !== 1) return { ok: false, stage: 'chart-identity' };
  hit[0].${AXES[op.axis]}(${JSON.stringify(op.title)});`;
  } else {
    call = `const all = ws.GetAllCharts() || [];
  const hit = all.filter(c => (${nameOf})(c) === ${JSON.stringify(op.name)});
  if (hit.length !== 1 || all.some(c => (${nameOf})(c) === ${JSON.stringify(op.newName)})) return { ok: false, stage: 'rename-identity' };
  hit[0].SetName(${JSON.stringify(op.newName)});`;
  }
  return {
    kind: 'callCommand',
    body: `
try {
  const name = ${JSON.stringify(sheet)};
  const ws = Api.GetSheet(name);
  if (!ws || ws.GetName() !== name) return { ok: false, stage: 'sheet-identity' };
  ${call}
  return { ok: true };
} catch (err) { return { ok: false, stage: 'threw', error: String(err) }; }
`,
    meta: Object.freeze({ family: 'charts', action: op.intent, sheet, name: op.name, newName: op.newName, title: op.title, legend: op.legend, axis: op.axis, type: op.intent === 'set_chart' ? op.type : null, series: op.intent === 'set_chart' ? expectedSeries(op) : null }),
  };
}

const chartsBefore = (op, before) => {
  const was = before && before.primary && before.primary.chartReads && before.primary.chartReads[op.selector.sheet];
  return was && Array.isArray(was.charts) ? was.charts : null;
};

const PRIMITIVE = {
  set_chart(op, charts, was) {
    const hits = named(charts, op.name);
    const grew = !was || charts.length === was.length + (named(was, op.name).length ? 0 : 1);
    return hits.length === 1 && same(fingerprint(hits[0]), wanted(op)) && grew;
  },
  // the title changed to the requested one; type and every series exactly as before. The whole-task check has no
  // before (null): there the title alone -- the series were proven by L1
  set_chart_title(op, charts, was, before) {
    const hits = named(charts, op.name);
    const old = was ? named(was, op.name) : [];
    const seriesKept = before === null || (old.length === 1 && same({ ...fingerprint(hits[0] || {}), title: null }, { ...fingerprint(old[0]), title: null }));
    return hits.length === 1 && trimTitle(hits[0].title) === op.title && seriesKept;
  },
  // X-O05: the requested axis title changed; type, title, every series, the data labels AND the OTHER axis title
  // exactly as before (nulling only the axis key being set -- same pattern as set_chart_title above)
  set_chart_axis_title(op, charts, was, before) {
    const hits = named(charts, op.name);
    if (hits.length !== 1) return false;
    const key = op.axis === 'horizontal' ? 'horAxisTitle' : 'verAxisTitle';
    const now = axisTitleOf(hits[0][key]);
    const old = was ? named(was, op.name) : [];
    const nullKey = f => ({ ...f, [key]: null });
    const kept = before === null || (old.length === 1 && same(nullKey(fingerprint(hits[0])), nullKey(fingerprint(old[0]))));
    return now === op.title && kept;
  },
  rename_chart(op, charts, was) {
    const to = named(charts, op.newName);
    const old = was ? named(was, op.name) : [];
    const kept = old.length !== 1 || (to.length === 1 && same(fingerprint(to[0]), fingerprint(old[0])));
    return named(charts, op.name).length === 0 && to.length === 1 && kept;
  },
  // the legend changed to the requested position; type, title and every series exactly as before.
  set_chart_legend(op, charts, was, before) {
    const hits = named(charts, op.name);
    const old = was ? named(was, op.name) : [];
    const kept = before === null || old.length !== 1 || same(fingerprint(hits[0] || {}), fingerprint(old[0]));
    return hits.length === 1 && hits[0].legend === op.legend && kept;
  },
};

function verifyPrimitive(op, before, after) {
  const g = gate(after && after.primary, op);
  if (!g.ok) return { ok: false, evidence: g };
  const ok = PRIMITIVE[op.intent](op, g.charts, chartsBefore(op, before), before);
  return { ok, evidence: { sheet: op.selector.sheet, chart: op.name, newName: op.newName, now: g.charts.map(c => c.name), proof: 'fresh-live-chart-getters' } };
}

/** L3: every cell, every OTHER chart of every sheet (name + identity, in order), every pivot name. */
function verifyPreservation(op, before, after) {
  const rest = cells.verifyPreservation(op, before, after);
  if (!rest.ok && rest.evidence && rest.evidence.reason) return rest;
  const b = before.primary.chartReads;
  const a = after.primary.chartReads;
  const mine = (sheet, c) => sheet === op.selector.sheet && (c.name === op.name || (op.newName !== null && c.name === op.newName));
  // `legend` is outside fingerprint() (which also feeds the set_chart identity check, where legend
  // has no creation-time counterpart) -- added here directly so an OTHER chart's legend is still
  // covered by this preservation check.
  const others = reads => (reads ? Object.entries(reads).map(([s, r]) => [s, r && Array.isArray(r.charts) ? r.charts.filter(c => !mine(s, c)).map(c => ({ name: c.name, legend: c.legend, ...fingerprint(c) })) : r, r && r.pivots]) : null);
  const chartsKept = !!b && !!a && same(others(b), others(a));
  // M7 S2b: every sheet's drawings equal; the requested chart change exempts ONLY its requested field -- set_chart adds
  // that one chart, rename_chart changes that chart's name (its anchor, size, rotation, flips must stay)
  const sh = op.selector.sheet;
  const draw = drawings.drawingDiffs(drawings.observe(before.primary.drawingReads, before.primary.sheets.map(x => x.name)), drawings.observe(after.primary.drawingReads, after.primary.sheets.map(x => x.name)), {
    expect: r => (op.intent === 'rename_chart' && r.sheet === sh && r.kind === 'chart' && r.name === op.name ? Object.freeze({ ...r, name: op.newName }) : r),
    added: r => op.intent === 'set_chart' && r.sheet === sh && r.kind === 'chart' && r.name === op.name,
  });
  return { ok: rest.ok && chartsKept && draw.length === 0, evidence: { ...rest.evidence, otherChartsAndPivotsKept: chartsKept, drawingsKept: draw.length === 0, drawingDiffs: draw } };
}

function wholeTaskChecks(op, finalObservation) {
  const v = verifyPrimitive(op, null, finalObservation);
  const what = op.intent === 'set_chart' ? `${op.name} over ${op.selector.range}`
    : op.intent === 'set_chart_title' ? `${op.name} titled ${JSON.stringify(op.title)}`
    : op.intent === 'set_chart_legend' ? `${op.name} legend set to ${op.legend}`
    : op.intent === 'set_chart_axis_title' ? `${op.name} ${op.axis} axis titled ${JSON.stringify(op.title)}`
    : `${op.name} renamed to ${op.newName}`;
  return [{ name: `${op.selector.sheet}: chart ${what}`, ok: v.ok, detail: v.ok ? null : v.evidence }];
}

const windowsOf = op => (op.intent === 'set_chart' ? [{ sheet: op.selector.sheet, range: op.selector.range }] : []);
const liveExtras = () => ({ chartSheets: ['*'], drawingSheets: ['*'] });

module.exports = { INTENTS, plan, classify, resolve, lower, verifyPrimitive, verifyPreservation, wholeTaskChecks, windowsOf, liveExtras };
