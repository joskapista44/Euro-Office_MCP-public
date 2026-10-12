'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const charts = require('../../src/adapters/xlsx/families/charts.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * M7 S5 (owner M7 GO 2026-10-03, owner channel; plan tier C): the measured, ready items -- C1 chart series in rows,
 * R5 a structural step across a merge (grow / shrink), R21 a descending sort with blank keys and same-row formulas,
 * S23 a freeze as the LAST step of a task. Every engine shape here was measured LIVE before (W7 K2, W5 probe, W6 probe).
 */

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}
const op = result => result.execution.appliedOperations[0];
const reason = result => op(result).evidence.evidence.reason;

// ---- C1: set_chart orientation 'rows' ------------------------------------------------------------------------------

const chartBook = (extra = []) => [
  { name: 'K', cells: { A1: { v: 'cat' }, B1: { v: 's1' }, C1: { v: 's2' }, A2: { v: 'a' }, B2: { v: 1 }, C2: { v: 4 }, A3: { v: 'b' }, B3: { v: 2 }, C3: { v: 5 }, A4: { v: 'c' }, B4: { v: 3 }, C4: { v: 6 } }, objects: { charts: extra, pivots: [], names: [] } },
  { name: 'O', cells: { A1: { v: 'keep' } }, objects: { charts: [], pivots: [], names: [] } },
];
const addChart = (extra = {}) => ({ id: 'm7-s5-c1', steps: [{ intent: 'set_chart', sheet: 'K', name: 'R1', range: 'A1:C4', type: 'bar', title: 'Rows', ...extra }] });
// the W7 K2 measurement (m6-w7-probe-20260929.json S2_READ, after a reopen), written out -- not derived by the code under test
const K2_ROWS = [
  { name: '=K!$A$2', values: '=K!$B$2:$C$2', cat: '=K!$B$1:$C$1', x: '' },
  { name: '=K!$A$3', values: '=K!$B$3:$C$3', cat: '=K!$B$1:$C$1', x: '' },
  { name: '=K!$A$4', values: '=K!$B$4:$C$4', cat: '=K!$B$1:$C$1', x: '' },
];
const chartsOf = (store, s) => store.durable.find(x => x.name === s).objects.charts;

test('C1 set_chart rows: APPLIED with the K2-measured series; the reopened retry SATISFIED 0/0', async () => {
  const store = createStore(chartBook());
  const a = await run(store, addChart({ orientation: 'rows' }));
  assert.equal(a.result.ok, true, JSON.stringify([a.result.execution.primitiveVerification, op(a.result).evidence]));
  assert.deepEqual(chartsOf(store, 'K')[0].series, K2_ROWS);
  const r = await run(store, addChart({ orientation: 'rows' }));
  assert.deepEqual([r.m.writes, r.m.commands], [0, 0]);
  assert.equal(op(r.result).status, OPERATION_STATUS.SATISFIED);
});

test('C1: the orientation is part of the identity -- a columns chart of that name refuses a rows request (and back), 0/0', async () => {
  const cols = [{ name: '=K!$B$1', values: '=K!$B$2:$B$4', cat: '=K!$A$2:$A$4', x: '' }, { name: '=K!$C$1', values: '=K!$C$2:$C$4', cat: '=K!$A$2:$A$4', x: '' }];
  const asCols = await run(createStore(chartBook([{ name: 'R1', type: 'bar', title: 'Rows', series: cols }])), addChart({ orientation: 'rows' }));
  assert.deepEqual([asCols.m.writes, asCols.m.commands], [0, 0]);
  assert.equal(reason(asCols.result), 'xlsx-chart-exists-with-another-identity');
  const asRows = await run(createStore(chartBook([{ name: 'R1', type: 'bar', title: 'Rows', series: K2_ROWS }])), addChart());
  assert.equal(reason(asRows.result), 'xlsx-chart-exists-with-another-identity');
});

test('C1: the command passes inRows = true to AddChart only for rows; a bad orientation is refused at planning', () => {
  const body = o => charts.lower(planOperations(addChart(o)).operations[0]).body;
  // X-O03: `type` is now interpolated via JSON.stringify (like every other embedded value in
  // this file) instead of a hardcoded 'bar' literal -- double-quoted, same as "K!..." above it.
  assert.match(body({ orientation: 'rows' }), /AddChart\("K!\$A\$1:\$C\$4", true, "bar"/);
  assert.match(body({}), /AddChart\("K!\$A\$1:\$C\$4", false, "bar"/);
  assert.match(body({ orientation: 'columns' }), /AddChart\("K!\$A\$1:\$C\$4", false, "bar"/);
  assert.match(planOperations(addChart({ orientation: 'diagonal' })).detail, /orientation must be/);
});

test('C1 L1 negative: the editor builds the chart in COLUMNS for a rows request -> not APPLIED', async () => {
  const store = createStore(chartBook());
  const { result } = await run(store, addChart({ orientation: 'rows' }), { fault: 'chart-wrong-series' });
  assert.equal(result.ok, false);
  assert.notEqual(op(result).status, OPERATION_STATUS.APPLIED);
});

// ---- R5: a structural step across a merge (grow / clip) ------------------------------------------------------------

const mergeBook = merges => [
  { name: 'P', cells: {
    A1: { v: 'r1' }, A2: { v: 'r2' }, A3: { v: 'r3' }, A4: { v: 'r4' }, A5: { v: 'r5' },
    B1: { v: 'b1' }, C1: { v: 'c1' }, D1: { v: 'd1' }, G2: { v: 'm' }, A8: { v: 'w' },
  }, merges },
  { name: 'Q', cells: { A1: { v: 'keep' } } },
];
const shiftStep = (intent, extra) => ({ id: 'm7-s5-r5', steps: [{ intent, sheet: 'P', ...extra }] });
const mergesOf = store => store.durable.find(s => s.name === 'P').merges;

for (const [name, t, merges, want] of [
  // the W5 probe's four shapes, written out from m6-w5-probe-20260929.json S2_READ (not derived by the code under test)
  ['insert 2 rows at 3 inside G2:H3 -> G2:H5', shiftStep('insert_rows', { at: 3, count: 2, anchor: { column: 'A', value: 'r3' } }), ['G2:H3'], ['G2:H5']],
  ['insert column B inside A8:C8 -> A8:D8', shiftStep('insert_columns', { at: 'B', count: 1, anchor: { row: 1, value: 'b1' } }), ['A8:C8'], ['A8:D8']],
  ['delete column B inside A8:C8 -> A8:B8', shiftStep('delete_columns', { at: 'B', count: 1, anchor: { row: 1, value: 'c1' } }), ['A8:C8'], ['A8:B8']],
  ['delete rows 3:4 across the bottom of G2:H3 -> G2:H2', shiftStep('delete_rows', { at: 3, count: 2, anchor: { column: 'A', value: 'r5' } }), ['G2:H3'], ['G2:H2']],
]) {
  test(`R5 APPLIED: ${name}; retry SATISFIED 0/0`, async () => {
    const store = createStore(mergeBook(merges));
    const { result, m } = await run(store, t);
    assert.equal(result.ok, true, JSON.stringify([result.execution.primitiveVerification, op(result).evidence]));
    assert.equal(m.commands, 1);
    assert.deepEqual(mergesOf(store), want);
    const again = await run(store, t);
    assert.equal(again.m.writes, 0);
    assert.equal(again.m.commands, 0);
    assert.equal(op(again.result).status, OPERATION_STATUS.SATISFIED);
  });
}

for (const [name, t, merges, why] of [
  ['a delete removing the merge\'s top-left row', shiftStep('delete_rows', { at: 2, count: 1, anchor: { column: 'A', value: 'r3' } }), ['G2:H3'], 'xlsx-structure-merge-top-left-deleted'],
  ['a delete removing the merge\'s top-left column', shiftStep('delete_columns', { at: 'A', count: 1, anchor: { row: 1, value: 'b1' } }), ['A8:C8'], 'xlsx-structure-merge-top-left-deleted'],
  ['a merge clipped to one cell', shiftStep('delete_columns', { at: 'B', count: 1, anchor: { row: 1, value: 'c1' } }), ['A8:B8'], 'xlsx-structure-merge-collapses-to-one-cell'],
]) {
  test(`R5 refused BEFORE dispatch: ${name}`, async () => {
    const { result, driver } = await run(createStore(mergeBook(merges)), t);
    assert.equal(result.ok, false);
    assert.equal(driver.calls.dispatch, 0);
    assert.equal(reason(result), why);
  });
}

test('R5 RED: a merge that only moves (no grow / no clip) -> L1 FAIL, nothing persisted', async () => {
  for (const [t, merges] of [
    [shiftStep('insert_rows', { at: 3, count: 2, anchor: { column: 'A', value: 'r3' } }), ['G2:H3']],
    [shiftStep('delete_columns', { at: 'B', count: 1, anchor: { row: 1, value: 'c1' } }), ['A8:C8']],
  ]) {
    const { result } = await run(createStore(mergeBook(merges)), t, { fault: 'merge-keeps-size' });
    assert.equal(result.ok, false);
    assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL', JSON.stringify(result.execution.primitiveVerification));
  }
});

// ---- R21: desc sort with blank keys; same-row formulas in the block ------------------------------------------------

// the W6 probe's SO shapes (m6-w6-probe-20260929.json): D/E desc with a blank key; G/H with H = G*10; J1 outside
const sortBook = () => [
  { name: 'D', cells: {
    D1: { v: 'num' }, E1: { v: 'orig' }, D2: { v: 3 }, E2: { v: 1 }, E3: { v: 2 }, D4: { v: 1 }, E4: { v: 3 }, D5: { v: 2 }, E5: { v: 4 }, D6: { v: 5 }, E6: { v: 5 },
    G1: { v: 'k' }, H1: { v: 'f' }, G2: { v: 4 }, H2: { v: 40, f: '=G2*10' }, G3: { v: 2 }, H3: { v: 20, f: '=G3*10' }, G4: { v: 3 }, H4: { v: 30, f: '=G4*10' }, G5: { v: 1 }, H5: { v: 10, f: '=G5*10' },
    J1: { v: 4, f: '=G2+0' },
  } },
  { name: 'O', cells: { A1: { v: 'keep' } } },
];
const sortTask = extra => ({ id: 'm7-s5-r21', steps: [{ intent: 'sort_range', sheet: 'D', hasHeaders: true, ...extra }] });
const Dc = store => store.durable.find(s => s.name === 'D').cells;

test('R21 APPLIED: desc with a blank key -> the blank row last (E travels with it); retry SATISFIED 0/0', async () => {
  const store = createStore(sortBook());
  const t = sortTask({ range: 'D1:E6', key: 'D', order: 'desc' });
  const { result, m } = await run(store, t);
  assert.equal(result.ok, true, JSON.stringify([result.execution.primitiveVerification, op(result).evidence]));
  assert.equal(m.commands, 1);
  // the measured result, written out: D 5,3,2,1,'' / E 5,1,4,3,2
  assert.deepEqual([2, 3, 4, 5, 6].map(r => (Dc(store)[`D${r}`] || {}).v), [5, 3, 2, 1, undefined]);
  assert.deepEqual([2, 3, 4, 5, 6].map(r => Dc(store)[`E${r}`].v), [5, 1, 4, 3, 2]);
  const again = await run(store, t);
  assert.equal(again.m.commands, 0);
  assert.equal(op(again.result).status, OPERATION_STATUS.SATISFIED);
});

test('R21 APPLIED: same-row formulas move with their row and point at the new row; the outside formula is not re-pointed', async () => {
  const store = createStore(sortBook());
  const t = sortTask({ range: 'G1:H5', key: 'G', order: 'asc' });
  const { result } = await run(store, t);
  assert.equal(result.ok, true, JSON.stringify([result.execution.primitiveVerification, result.execution.preservationVerification]));
  assert.deepEqual([2, 3, 4, 5].map(r => Dc(store)[`G${r}`].v), [1, 2, 3, 4]);
  assert.deepEqual([2, 3, 4, 5].map(r => Dc(store)[`H${r}`].f), ['=G2*10', '=G3*10', '=G4*10', '=G5*10']);
  assert.equal(Dc(store).J1.f, '=G2+0');
  const again = await run(store, t);
  assert.equal(again.m.commands, 0);
});

for (const [name, extra, mutate, why] of [
  ['an ASCENDING sort with a blank key (placement not measured)', { range: 'D1:E6', key: 'D', order: 'asc' }, null, 'xlsx-sort-blank-key-unmeasured'],
  ['a formula in the key column', { range: 'G1:H5', key: 'G', order: 'asc' }, b => { b[0].cells.G3 = { v: 2, f: '=1+1' }; }, 'xlsx-sort-formula-key-unmeasured'],
  ['a formula pointing at another row', { range: 'G1:H5', key: 'G', order: 'asc' }, b => { b[0].cells.H3.f = '=G2*10'; }, 'xlsx-sort-formula-in-block-unmeasured'],
  ['a formula pointing outside the block', { range: 'G1:H5', key: 'G', order: 'asc' }, b => { b[0].cells.H3.f = '=J3*10'; }, 'xlsx-sort-formula-in-block-unmeasured'],
  ['an absolute reference', { range: 'G1:H5', key: 'G', order: 'asc' }, b => { b[0].cells.H3.f = '=$G3*10'; }, 'xlsx-sort-formula-in-block-unmeasured'],
  ['a same-row range', { range: 'G1:H5', key: 'G', order: 'asc' }, b => { b[0].cells.H3.f = '=SUM(G3:G3)'; }, 'xlsx-sort-formula-in-block-unmeasured'],
  ['another sheet', { range: 'G1:H5', key: 'G', order: 'asc' }, b => { b[0].cells.H3.f = '=O!G3*10'; }, 'xlsx-sort-formula-in-block-unmeasured'],
  ['a defined name', { range: 'G1:H5', key: 'G', order: 'asc' }, b => { b[0].cells.H3.f = '=G3*Rate'; }, 'xlsx-sort-formula-in-block-unmeasured'],
]) {
  test(`R21 refused BEFORE dispatch: ${name}`, async () => {
    const b = sortBook();
    if (mutate) mutate(b);
    const { result, driver } = await run(createStore(b), sortTask(extra));
    assert.equal(result.ok, false);
    assert.equal(driver.calls.dispatch, 0);
    assert.equal(reason(result), why);
  });
}

test('R21 allowed: a function call and a string literal around same-row references', async () => {
  const b = sortBook();
  b[0].cells.H3 = { v: 'x', f: '=IF(G3>1,"G9",TRUE)' };
  const { result } = await run(createStore(b), sortTask({ range: 'G1:H5', key: 'G', order: 'asc' }));
  assert.equal(result.ok, true, JSON.stringify(op(result).evidence));
});

test('R21 RED: a formula that moves but keeps its old row -> L1 FAIL; blanks first in desc -> L1 FAIL', async () => {
  const stays = await run(createStore(sortBook()), sortTask({ range: 'G1:H5', key: 'G', order: 'asc' }), { fault: 'sort-formula-stays' });
  assert.equal(stays.result.ok, false);
  assert.equal(stays.result.execution.primitiveVerification[0].verdict, 'FAIL');
  const rev = await run(createStore(sortBook()), sortTask({ range: 'D1:E6', key: 'D', order: 'desc' }), { fault: 'sort-reverse' });
  assert.equal(rev.result.ok, false);
  assert.equal(rev.result.execution.primitiveVerification[0].verdict, 'FAIL');
  const first = await run(createStore(sortBook()), sortTask({ range: 'D1:E6', key: 'D', order: 'desc' }), { fault: 'sort-blanks-first' });
  assert.equal(first.result.ok, false);
  assert.equal(first.result.execution.primitiveVerification[0].verdict, 'FAIL');
});

// ---- S23: a freeze as the LAST step of a mixed task -----------------------------------------------------------------

const freezeBook = () => [
  { name: 'Home', cells: { A1: { v: 'start' } }, pane: { row0: 2, col0: 0 } },
  { name: 'Data', cells: { A1: { v: 'Name' }, B2: { v: 5 } } },
];
const mixed = (steps = null) => ({ id: 'm7-s5-s23', steps: steps || [
  { intent: 'write_range', sheet: 'Data', range: 'C2', values: [[7]] },
  { intent: 'freeze_panes', sheet: 'Data', rows: 1 },
] });
const sheetOf = (store, n) => store.durable.find(s => s.name === n);

test('S23 APPLIED: write, then freeze LAST -- both applied, the active sheet restored; retry SATISFIED 0/0', async () => {
  const store = createStore(freezeBook());
  const { result, m } = await run(store, mixed());
  assert.equal(result.ok, true, JSON.stringify([result.execution.primitiveVerification, result.execution.preservationVerification]));
  assert.equal(m.commands, 2);
  assert.deepEqual(result.execution.appliedOperations.map(o => o.status), [OPERATION_STATUS.APPLIED, OPERATION_STATUS.APPLIED]);
  assert.equal(sheetOf(store, 'Data').cells.C2.v, 7);
  assert.deepEqual(sheetOf(store, 'Data').pane, { row0: 1, col0: 0 });
  assert.equal(store.active, 'Home', 'the originally active sheet is active again');
  const again = await run(store, mixed());
  assert.equal(again.m.writes, 0);
  assert.equal(again.m.commands, 0);
});

test('S23 refused at planning: a freeze that is not the last step; two freezes in one task', () => {
  const notLast = planOperations(mixed([{ intent: 'freeze_panes', sheet: 'Data', rows: 1 }, { intent: 'write_range', sheet: 'Data', range: 'C2', values: [[7]] }]));
  assert.equal(notLast.ok, false);
  assert.match(notLast.detail, /LAST step/);
  const two = planOperations(mixed([{ intent: 'freeze_panes', sheet: 'Data', rows: 1 }, { intent: 'unfreeze_panes', sheet: 'Home' }]));
  assert.equal(two.ok, false);
  assert.match(two.detail, /at most one freeze/);
  assert.equal(planOperations(mixed()).ok, true);
});

test('S23 RED: the freeze leaves the target active / moves another pane -> not ok, the freeze not APPLIED', async () => {
  for (const fault of ['no-restore', 'side-effect']) {
    const { result } = await run(createStore(freezeBook()), mixed(), { fault });
    assert.equal(result.ok, false, fault);
    assert.notEqual(result.execution.appliedOperations[1].status, OPERATION_STATUS.APPLIED, fault);
  }
});

// ---- C3 (tier D): set_chart_title on an existing chart --------------------------------------------------------------

const retitle = title => ({ id: 'm7-d-c3', steps: [{ intent: 'set_chart_title', sheet: 'K', name: 'R1', title }] });
const chartOf = store => store.durable.find(s => s.name === 'K').objects.charts.find(c => c.name === 'R1');

test('C3 APPLIED: the title changes, type and series stay; read back with CR LF (the probe) and still SATISFIED 0/0', async () => {
  const store = createStore(chartBook());
  assert.equal((await run(store, addChart())).result.ok, true);
  const series = JSON.stringify(chartOf(store).series);
  const { result, m } = await run(store, retitle('Első cím'));
  assert.equal(result.ok, true, JSON.stringify([result.execution.primitiveVerification, result.execution.preservationVerification]));
  assert.equal(m.commands, 1);
  assert.equal(chartOf(store).title, 'Első cím\r\n');
  assert.equal(JSON.stringify(chartOf(store).series), series);
  const again = await run(store, retitle('Első cím'));
  assert.equal(again.m.commands, 0);
  assert.equal(op(again.result).status, OPERATION_STATUS.SATISFIED);
});

test('C3 refused: no chart of the name; a title of two lines', async () => {
  const none = await run(createStore(chartBook()), retitle('X'));
  assert.equal(none.driver.calls.dispatch, 0);
  assert.equal(reason(none.result), 'xlsx-chart-not-found');
  assert.equal(planOperations(retitle('a\nb')).ok, false);
});

test('C3 RED: the retitle drops a series -> L1 FAIL', async () => {
  const store = createStore(chartBook());
  await run(store, addChart());
  const { result } = await run(store, retitle('New'), { fault: 'title-drops-series' });
  assert.equal(result.ok, false);
  assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL');
});

test('C3: the command calls SetTitle on the one named chart (the fake applies by metadata, so the call is checked here)', () => {
  const body = charts.lower(planOperations(retitle('Új "cím"')).operations[0]).body;
  assert.match(body, /hit\[0\]\.SetTitle\("Új \\"cím\\"", 12, false\);/);
  assert.doesNotMatch(body, /SetName|AddChart/);
});
