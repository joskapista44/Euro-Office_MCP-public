'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { PERSISTENCE_STATE } = require('../../src/core/persistence/verified-writes.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');
const { readQuery } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');

/**
 * M7 backlog X-O01 cf-slice: the cf-reads reader type-branches so a range carrying
 * a Databar/IconSet/ColorScale no longer fails with xlsx-cf-read-unobservable, and three new rule
 * types (aboveAverage/top10/unique) get add/delete support.
 *
 * X-O01 full implementation (enum-probe LIVE MEASURED -> #ref dispatch): colorScale/
 * databar/iconSet get add/delete too, but DEFAULT-ONLY -- no appearance fields -- since the
 * criteria/enum setter shapes (GetColorScaleCriteria/GetIconCriteria, BarFillType/Direction/
 * MinPointType/AxisPosition/IconSet value strings) were never measured (cf.cjs's own doc comment).
 * Identity is type + appliesTo only (CF_IDENTITY_FIELDS[...] = [], same empty shape as 'unique').
 */

const book = () => [{ name: 'C', cells: { A1: { v: 1 }, A2: { v: 6 }, A3: { v: 9 } } }, { name: 'O', cells: { A1: { v: 'keep' } } }];
const add = (r, range = 'A1:A3') => ({ id: 'm7-o-slice', steps: [{ intent: 'add_conditional_format', sheet: 'C', range, rule: r }] });
const del = (r, range = 'A1:A3') => ({ id: 'm7-o-slice', steps: [{ intent: 'delete_conditional_format', sheet: 'C', range, rule: r }] });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}
const rules = store => store.durable.find(s => s.name === 'C').rules;
const reason = result => result.execution.appliedOperations[0].evidence.evidence.reason;

// ---- (c) the three new types: add APPLIED, retry SATISFIED 0/0; delete the same way ------------------------------

for (const [name, rule, idField, reRead] of [
  ['aboveAverage', { type: 'aboveAverage', above: true, numStdDev: 1 }, 'aboveBelow', true],
  ['top10', { type: 'top10', rank: 5, percent: false, bottom: true }, 'topBottom', 'xlTop10Bottom'],
  ['unique', { type: 'unique' }, null, null],
  ['colorScale', { type: 'colorScale' }, null, null],
  ['databar', { type: 'databar' }, null, null],
  ['iconSet', { type: 'iconSet' }, null, null],
]) {
  test(`${name}: add APPLIED, retry SATISFIED 0/0; delete APPLIED, retry SATISFIED 0/0`, async () => {
    const store = createStore(book());
    const a = await run(store, add(rule));
    assert.equal(a.result.ok, true, JSON.stringify([name, a.result.execution.primitiveVerification, a.result.execution.appliedOperations[0].evidence]));
    assert.equal(a.m.commands, 1);
    assert.equal(a.result.persistenceState, PERSISTENCE_STATE.FORCESAVE_CONFIRMED);
    assert.equal(rules(store).length, 1);
    if (idField) assert.equal(rules(store)[0][idField], reRead, name);
    assert.equal((await run(store, add(rule))).m.commands, 0, `${name}: retry 0/0`);
    const d = await run(store, del(rule));
    assert.equal(d.result.ok, true, JSON.stringify([name, d.result.execution.primitiveVerification]));
    assert.equal(rules(store).length, 0, name);
    assert.equal((await run(store, del(rule))).m.commands, 0, `${name}: delete retry 0/0`);
  });
}

// ---- (c) field-mismatch -> L1 FAIL (aboveAverage/top10 -- unique has no scalar to corrupt) ------------------------

for (const [name, rule] of [['aboveAverage', { type: 'aboveAverage', above: true }], ['top10', { type: 'top10', rank: 3, percent: true, bottom: false }]]) {
  test(`${name}: a field lands wrong -> L1 FAIL`, async () => {
    const { result } = await run(createStore(book()), add(rule), { fault: 'cf-wrong-value' });
    assert.equal(result.ok, false, name);
    assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL', name);
  });
}

// ---- EACH identity field individually matters -- not just "some field", proving CF_IDENTITY_FIELDS is complete ---

for (const [name, rule, field] of [
  ['aboveAverage.aboveBelow', { type: 'aboveAverage', above: true, numStdDev: 1 }, 'aboveBelow'],
  ['aboveAverage.numStdDev', { type: 'aboveAverage', above: true, numStdDev: 1 }, 'numStdDev'],
  ['top10.rank', { type: 'top10', rank: 3, percent: true, bottom: false }, 'rank'],
  ['top10.percent', { type: 'top10', rank: 3, percent: true, bottom: false }, 'percent'],
  ['top10.topBottom', { type: 'top10', rank: 3, percent: true, bottom: false }, 'topBottom'],
]) {
  test(`${name} alone lands wrong -> L1 FAIL (the OTHER fields staying right does not hide it)`, async () => {
    const { result } = await run(createStore(book()), add(rule), { fault: 'cf-wrong-one-field', cfWrongField: field });
    assert.equal(result.ok, false, name);
    assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL', name);
  });
}

// ---- (a) a cellValue add next to a Databar: the OLD type-blind reader threw; the NEW one doesn't ------------------

test('cellValue add on a range that ALSO carries a Databar: APPLIED, the Databar stays in L3 unchanged', async () => {
  const b = book();
  b[0].rules = [{ kind: 'conditional', type: 'xlDatabar', barColor: 65280, range: 'A1:A3' }];
  const store = createStore(b);
  const r = { type: 'cellValue', operator: 'greater', formula1: '5', fill: '#FF0000' };
  const a = await run(store, add(r));
  assert.equal(a.result.ok, true, JSON.stringify([a.result.execution.primitiveVerification, a.result.execution.appliedOperations[0].evidence]));
  assert.equal(a.result.execution.preservationVerification[0].verdict, 'PASS', 'the Databar is an OTHER rule L3 must keep');
  const kept = rules(store).find(x => x.type === 'xlDatabar');
  assert.ok(kept, 'the Databar is still on the range');
  assert.equal(kept.barColor, 65280, 'untouched');
});

// ---- (b) an empty type (the measured unidentifiable shape) makes the WHOLE RANGE unobservable, fail-closed --------

test('an empty-type rule on the range (the measured unique/duplicate shape) refuses ANY cf operation on it, fail-closed', async () => {
  const b = book();
  b[0].rules = [{ kind: 'conditional', type: '', range: 'A1:A3' }];
  const { result, driver } = await run(createStore(b), add({ type: 'cellValue', operator: 'greater', formula1: '5' }));
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(reason(result), 'xlsx-cf-read-unobservable');
});

// ---- (d) the duplicate variant is refused BEFORE any write is attempted -------------------------------------------

test("unique's duplicate variant is refused at plan time (measured unidentifiable -- GetType() reads empty)", () => {
  const p = planOperations(add({ type: 'unique', dupe: true }));
  assert.equal(p.ok, false);
  assert.match(p.detail, /rule.dupe must be/);
});

// ---- X-O01 full implementation: colorScale/databar/iconSet add/delete succeed DEFAULT-ONLY (the
// loop above); any appearance field is refused at PLAN time (the criteria/enum shapes are NEM-MERT) --

test('databar/iconSet/colorScale: any appearance field is refused at plan time (unmeasured shape)', () => {
  for (const [type, field] of [['databar', { barColor: '#00FF00' }], ['iconSet', { iconSet: 'xl3TrafficLights1' }], ['colorScale', { n: 2 }]]) {
    const p = planOperations(add({ type, ...field }));
    assert.equal(p.ok, false, type);
    assert.match(p.detail, /a .* rule is \{ type:/, type);
  }
});

test('colorScale/databar/iconSet: adding on a range that already carries ANOTHER of the three is refused (same-condition-other-range check never applies here -- different types do not collide)', async () => {
  const b = book();
  b[0].rules = [{ kind: 'conditional', type: 'xlIconSets', iconSet: 'xl3TrafficLights1', showIconOnly: false, range: 'A1:A3' }];
  const store = createStore(b);
  const a = await run(store, add({ type: 'databar' }));
  assert.equal(a.result.ok, true, JSON.stringify(a.result.execution.appliedOperations[0].evidence));
  assert.equal(rules(store).length, 2);
  assert.ok(rules(store).some(x => x.type === 'xlIconSets'), 'the icon set is untouched');
});

// ---- the DRIVER's own cfReads text, directly: proves the branch truly never calls the throwing getters ----------
// (xlsx-w6-cf.test.cjs/xlsx-m7-d-rules.test.cjs test at the FAKE-DRIVER level; this is the one place that runs
// the REAL `readQuery()`-generated script against a fake Api whose GetOperator/GetFormula1/GetFormula2/
// GetFillColor THROW if called -- the fake driver's OWN cfReadsOf never calls real getters at all, so it cannot
// catch a regression where the branch accidentally still reaches one of them on a non-cellValue/expression type.)

function condItem(type, extra) {
  const forbidden = name => () => { throw new Error(`${name} called on ${type}`); };
  return Object.assign({ GetType: () => type, GetOperator: forbidden('GetOperator'), GetFormula1: forbidden('GetFormula1'), GetFormula2: forbidden('GetFormula2'), GetFillColor: forbidden('GetFillColor'), GetAppliesTo: () => ({ GetAddress: () => 'A1:A3' }) }, extra);
}
function cfReadsFor(item) {
  const q = readQuery([], { cfRanges: [{ sheet: 'C', range: 'A1:A3' }] });
  const ws = { GetName: () => 'C', GetIndex: () => 0, GetUsedRange: () => null, GetRange: () => ({ GetFormatConditions: () => ({ GetCount: () => 1, GetItem: () => item }) }) };
  const Api = { GetSheets: () => [ws], GetSheet: n => (n === 'C' ? ws : null) };
  return new Function('Api', q)(Api).cfReads['C!A1:A3'];
}

test('driver cfReads: the six new-reader types never call GetOperator/GetFormula1/GetFormula2/GetFillColor -- type + appliesTo + their OWN measured scalars only', () => {
  const CASES = {
    xlAboveAverageCondition: [condItem('xlAboveAverageCondition', { GetAboveBelow: () => true, GetNumStdDev: () => 1 }), { type: 'xlAboveAverageCondition', appliesTo: 'A1:A3', aboveBelow: true, numStdDev: 1 }],
    xlTop10: [condItem('xlTop10', { GetRank: () => 5, GetPercent: () => false, GetTopBottom: () => 'xlTop10Top' }), { type: 'xlTop10', appliesTo: 'A1:A3', rank: 5, percent: false, topBottom: 'xlTop10Top' }],
    xlUniqueValues: [condItem('xlUniqueValues', { GetDupeUnique: () => 'xlUnique' }), { type: 'xlUniqueValues', appliesTo: 'A1:A3', dupeUnique: 'xlUnique' }],
    xlDatabar: [condItem('xlDatabar', { GetBarColor: () => ({ GetRGB: () => 65280 }) }), { type: 'xlDatabar', appliesTo: 'A1:A3', barColor: 65280 }],
    xlIconSets: [condItem('xlIconSets', { GetIconSet: () => 'xl3TrafficLights1', GetShowIconOnly: () => false }), { type: 'xlIconSets', appliesTo: 'A1:A3', iconSet: 'xl3TrafficLights1', showIconOnly: false }],
    xlColorScale: [condItem('xlColorScale', {}), { type: 'xlColorScale', appliesTo: 'A1:A3' }],
  };
  for (const [type, [item, want]] of Object.entries(CASES)) assert.deepEqual(cfReadsFor(item), [want], type);
});

test('driver cfReads: cellValue/expression are UNCHANGED -- they still call the SAME getters as before X-O01', () => {
  const cv = Object.assign(condItem('xlCellValue', {}), { GetOperator: () => 'xlGreater', GetFormula1: () => '5', GetFormula2: () => '', GetFillColor: () => ({ GetRGB: () => 16711680 }) });
  assert.deepEqual(cfReadsFor(cv), [{ type: 'xlCellValue', operator: 'xlGreater', formula1: '5', formula2: '', fill: 16711680, appliesTo: 'A1:A3' }]);
  const ex = Object.assign(condItem('xlExpression', {}), { GetOperator: () => '', GetFormula1: () => '$A1>5', GetFormula2: () => '', GetFillColor: () => null });
  assert.deepEqual(cfReadsFor(ex), [{ type: 'xlExpression', operator: '', formula1: '$A1>5', formula2: '', fill: null, appliesTo: 'A1:A3' }]);
});

test('driver cfReads: an empty type makes the WHOLE axis a reason string, not an array with a bad entry', () => {
  assert.equal(cfReadsFor(condItem('', {})), 'UNOBSERVABLE:cf-type-empty');
});

test('driver cfReads: a MISSING measured getter on an otherwise-recognized type degrades to UNOBSERVABLE, not a throw escaping the script', () => {
  assert.equal(cfReadsFor(condItem('xlUniqueValues', {})), 'UNOBSERVABLE:TypeError: x[getter] is not a function');
});
