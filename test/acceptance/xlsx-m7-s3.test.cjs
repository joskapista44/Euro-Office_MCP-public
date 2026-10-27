'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { observedStyleChanges } = require('../../src/adapters/xlsx/style-axis.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * M7 S3 / S31 (owner GO 2026-10-01, owner channel): write + format + column width in ONE task. Two overlapping steps
 * are planned only when both declare the axes they modify and the axes are disjoint; the later step then depends on
 * the earlier (its L1 / L3 run on the state the earlier produced). The same axis, and an undeclared or multi-axis
 * effect, stay refused. A write and a size never change style: their L3 compares every style key observed on both
 * sides; a format never changes a value (its L1).
 */

const book = () => [
  { name: 'Data', cells: { E5: { v: 'keep' }, F1: { v: 9 } } },
  { name: 'Other', cells: { A1: { v: 'keep me' } } },
];

const WRITE = { intent: 'write_range', id: 'w', sheet: 'Data', range: 'A1:C3', values: [['Name', 'Qty', 'Price'], ['apple', 3, 1.5], ['pear', 7, 2.25]] };
const BOLD = { intent: 'format_range', id: 'f', sheet: 'Data', range: 'A1:C1', format: { font: { bold: true }, fill: '#DDEEFF' } };
const NF = { intent: 'format_range', id: 'n', sheet: 'Data', range: 'C2:C3', format: { numberFormat: '0.00' } };
const WIDTH = { intent: 'layout_range', id: 'l', sheet: 'Data', columns: 'A:C', width: 18 };
const task = (...steps) => ({ id: 'm7-s3', steps });
const THREE = task(WRITE, BOLD, WIDTH);

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, m: result.metrics };
}

const data = store => store.durable.find(s => s.name === 'Data');
const statuses = result => result.execution.appliedOperations.map(e => `${e.operationId}:${e.status}`);
const failed = result => result.execution.appliedOperations.find(e => e.status !== OPERATION_STATUS.APPLIED && e.status !== OPERATION_STATUS.SATISFIED);

// ---- planning ----

test('planner: write + format + width overlapping on disjoint axes -> planned, each later step depends on the earlier ones', () => {
  const p = planOperations(THREE);
  assert.equal(p.ok, true, p.detail);
  assert.deepEqual(p.operations.map(o => o.footprint), [{ value: 'A1:C3' }, { style: 'A1:C1' }, { columns: [1, 3], rows: '*' }]);
  assert.equal(p.operations[0].dependsOn, undefined);
  assert.deepEqual(p.operations[1].dependsOn, ['w']);
  assert.deepEqual(p.operations[2].dependsOn, ['w', 'f']);
  const apart = planOperations(task(WRITE, { ...BOLD, range: 'E1' }));
  assert.equal(apart.ok, true);
  assert.equal(apart.operations[1].dependsOn, undefined, 'no overlap, no dependency');
});

test('planner: the SAME axis overlapping stays refused', () => {
  const cases = [
    [task(WRITE, { ...WRITE, id: 'w2', range: 'C3', values: [['x']] }), /same axis: value/],
    [task(BOLD, { ...BOLD, id: 'f2', range: 'B1', format: { wrap: true } }), /same axis: style/],
    [task(WIDTH, { intent: 'layout_range', id: 'h', sheet: 'Data', columns: 'C', hidden: true }), /same axis: columns/],
    [task({ ...BOLD, format: { wrap: true } }, WIDTH), /same axis: rows/],
    [task(WIDTH, { intent: 'layout_range', id: 'r', sheet: 'Data', rows: '2', height: 30 }), /same axis: rows/],
  ];
  for (const [t, re] of cases) {
    const p = planOperations(t);
    assert.equal(p.ok, false, JSON.stringify(t.steps.map(s => s.id)));
    assert.match(p.detail, /overlapping-writes-refused/);
    assert.match(p.detail, re);
  }
});

test('planner: the area is per axis -- steps that share NO cell still meet on the row axis (owner review of 93ff8899)', () => {
  const wrapA1 = { intent: 'format_range', id: 'f', sheet: 'Data', range: 'A1', format: { wrap: true } };
  const widthB = { intent: 'layout_range', id: 'l', sheet: 'Data', columns: 'B', width: 18 };
  const widthD = { intent: 'layout_range', id: 'l2', sheet: 'Data', columns: 'D', width: 12 };
  const cases = [
    [task(wrapA1, widthB), 'A1 wrap + column B width: row 1 auto height'],
    [task(widthB, widthD), 'two widths on separate columns: both re-fit every wrapped row'],
    [task(widthB, { intent: 'layout_range', id: 'r', sheet: 'Data', rows: '40', height: 30 }), 'a width and a far row height'],
  ];
  for (const [t, why] of cases) {
    const p = planOperations(t);
    assert.equal(p.ok, false, why);
    assert.match(p.detail, /overlapping-writes-refused: .*\(same axis: rows\)/, why);
  }
  const hidden = planOperations(task({ intent: 'layout_range', id: 'h1', sheet: 'Data', columns: 'B', hidden: true }, { intent: 'layout_range', id: 'h2', sheet: 'Data', columns: 'D', hidden: true }));
  assert.equal(hidden.ok, true, 'two hidden columns apart: columns do not meet, no row effect');
  assert.equal(planOperations(task(wrapA1, { ...widthB, sheet: 'Other' })).ok, true, 'another sheet does not meet');
  assert.equal(planOperations(task({ ...wrapA1, format: { fill: '#010203' } }, widthB)).ok, true, 'a fill has no row effect');
  const height = row => ({ intent: 'layout_range', id: 'r', sheet: 'Data', rows: row, height: 30 });
  assert.equal(planOperations(task({ ...wrapA1, range: 'C5:D6' }, height('6'))).ok, false, 'a wrap claims ITS rows: row 6 meets');
  assert.equal(planOperations(task({ ...wrapA1, range: 'C5:D6' }, height('8'))).ok, true, 'a wrap does not claim row 8');
});

test('planner: a step that declares no axes (merge, clear, sort ...) gets no exemption', () => {
  for (const other of [
    { intent: 'merge_range', id: 'm', sheet: 'Data', range: 'A1:B1' },
    { intent: 'clear_range', id: 'c', sheet: 'Data', range: 'B2' },
    { intent: 'sort_range', id: 's', sheet: 'Data', range: 'A2:C3', keyColumn: 'B', order: 'asc', hasHeaders: false },
  ]) {
    const p = planOperations(task(WRITE, other));
    assert.equal(p.ok, false, other.intent);
    assert.match(p.detail, /overlapping-writes-refused: .*\(axes not declared\)/, other.intent);
  }
});

test('liveScope of the composition reads style on the target sheet (format AND width ask for it) and the layout extent', () => {
  const s = liveScope(THREE);
  assert.deepEqual(s.styleSheets, ['Data']);
  // OD-2a / D-W: a width step also asks the canonical value of its requested width
  assert.deepEqual(s.layout, [{ sheet: 'Data', maxCol: 3, maxRow: 0, canon: [18] }]);
  assert.deepEqual(liveScope(task(WIDTH)).styleSheets, ['Data'], 'a width alone reads style too (its L3 compares it)');
  assert.equal(liveScope(task(WRITE)).styleSheets, undefined, 'a standalone write reads no style (the M6 contract)');
});

// ---- execution: the three task states ----

test('FRESH: write + bold header + width in one session -> 3 APPLIED (L1 + L3 each), whole task PASS; reopened retry 0 write / 0 command', async () => {
  const store = createStore(book());
  const { result, m } = await run(store, THREE);
  assert.equal(result.ok, true, JSON.stringify([statuses(result), failed(result)]));
  assert.deepEqual(statuses(result), ['w:APPLIED', 'f:APPLIED', 'l:APPLIED']);
  assert.equal(m.writes, 3);
  assert.equal(m.commands, 3);
  assert.equal(result.execution.preservationVerification.every(v => v.verdict === 'PASS'), true);
  const d = data(store);
  assert.equal(d.cells.B2.v, 3);
  assert.equal(d.cells.A1.bold, true);
  assert.equal(d.cells.C1.fill, 0xDDEEFF);
  assert.equal(d.layout.cols.B.width, 18);
  const again = await run(store, THREE);
  assert.equal(again.result.ok, true);
  assert.deepEqual(statuses(again.result), ['w:SATISFIED', 'f:SATISFIED', 'l:SATISFIED']);
  assert.equal(again.m.writes, 0);
  assert.equal(again.m.commands, 0);
});

test('PARTLY satisfied: the values already there -> write SATISFIED, format + width APPLIED; values + format there -> only the width', async () => {
  const store = createStore(book());
  assert.equal((await run(store, task(WRITE))).result.ok, true);
  const one = await run(store, THREE);
  assert.equal(one.result.ok, true);
  assert.deepEqual(statuses(one.result), ['w:SATISFIED', 'f:APPLIED', 'l:APPLIED']);
  assert.equal(one.m.writes, 2);

  const store2 = createStore(book());
  assert.equal((await run(store2, task(WRITE, BOLD))).result.ok, true);
  const two = await run(store2, THREE);
  assert.equal(two.result.ok, true);
  assert.deepEqual(statuses(two.result), ['w:SATISFIED', 'f:SATISFIED', 'l:APPLIED']);
  assert.equal(two.m.writes, 1);
  assert.equal(two.m.commands, 1);
});

test('FULLY satisfied before the task: 0 write / 0 command', async () => {
  const store = createStore(book());
  assert.equal((await run(store, THREE)).result.ok, true);
  const again = await run(store, THREE);
  assert.equal(again.result.ok, true);
  assert.equal(again.m.writes, 0);
  assert.equal(again.m.commands, 0);
});

test('a later, requested number format does not make the earlier write wrong (GetValue reads the raw value, W4 LIVE)', async () => {
  const store = createStore(book());
  const { result } = await run(store, task(WRITE, NF, WIDTH));
  assert.equal(result.ok, true, JSON.stringify([statuses(result), failed(result)]));
  assert.equal(data(store).cells.C2.nf, '0.00');
  assert.equal(data(store).cells.C2.v, 1.5);
});

test('X-F02: font no longer depends on step order -- it reaches an EMPTY cell even BEFORE the write targets it', async () => {
  const store = createStore(book());
  const { result, m } = await run(store, task({ ...BOLD, format: { font: { bold: true } } }, WRITE));
  assert.equal(result.ok, true, JSON.stringify([statuses(result), failed(result)]));
  assert.equal(m.writes, 2);
  assert.equal(data(store).cells.A1.font.bold, true);
  assert.equal(data(store).cells.A1.v, 'Name', 'the later write still lands');
});

// ---- negative controls: an unrequested change on another axis still fails ----

test('NEG: the format ALSO changes a value -> the format is never APPLIED (its L1), the width never runs', async () => {
  const store = createStore(book());
  const { result } = await run(store, THREE, { fault: 'format-changes-value' });
  assert.equal(result.ok, false);
  assert.deepEqual(statuses(result).slice(0, 1), ['w:APPLIED']);
  const f = result.execution.appliedOperations.find(e => e.operationId === 'f');
  assert.notEqual(f.status, OPERATION_STATUS.APPLIED);
  assert.equal(result.execution.appliedOperations.find(e => e.operationId === 'l').status, OPERATION_STATUS.SKIPPED_DUE_TO_PRIOR_FAILURE);
});

test('NEG: the width ALSO changes a format -> the width is never APPLIED (its L3 compares style)', async () => {
  const store = createStore(book());
  const { result } = await run(store, THREE, { fault: 'layout-changes-style', styleCell: 'B1' });
  assert.equal(result.ok, false);
  const l = result.execution.appliedOperations.find(e => e.operationId === 'l');
  assert.notEqual(l.status, OPERATION_STATUS.APPLIED);
  const l3 = result.execution.preservationVerification.find(v => v.verdict !== 'PASS');
  assert.ok(JSON.stringify(l3).includes('Data!B1'), JSON.stringify(l3));
});

test('NEG: a width alone that changes a format fails too (the widened L3 is not composition-only)', async () => {
  const store = createStore(book());
  const { result } = await run(store, task(WIDTH), { fault: 'layout-changes-style', styleCell: 'E5' });
  assert.equal(result.ok, false);
});

test('NEG: in a composition the write ALSO changes a style key -> the write is never APPLIED (its L3, style observed)', async () => {
  const store = createStore(book());
  const { result, m } = await run(store, THREE, { fault: 'write-changes-style' });
  assert.equal(result.ok, false);
  assert.notEqual(result.execution.appliedOperations.find(e => e.operationId === 'w').status, OPERATION_STATUS.APPLIED);
  assert.equal(m.writes, 0, 'nothing counted as written');
});

test('LIMIT, stated: a standalone write reads no style, so it claims nothing about it (the M6 contract)', async () => {
  const store = createStore(book());
  const { result } = await run(store, task(WRITE), { fault: 'write-changes-style' });
  assert.equal(result.ok, true);
});

test('NEG: a style getter throws after the dispatch (style was observable before) -> preservation not proven, never APPLIED', async () => {
  for (const [script, t, op] of [
    [{ styleGetterThrows: true }, THREE, 'w'],
    [{ fillDropped: true }, THREE, 'w'],
    [{ styleGetterThrows: true }, task(WIDTH), 'l'],
    [{ fillDropped: true }, task(WIDTH), 'l'],
  ]) {
    const store = createStore(book());
    const { result } = await run(store, t, script);
    assert.equal(result.ok, false, JSON.stringify(script));
    assert.notEqual(result.execution.appliedOperations.find(e => e.operationId === op).status, OPERATION_STATUS.APPLIED, JSON.stringify(script));
  }
});

// ---- the shared style comparison ----

test('observedStyleChanges: observable -> observable exact; observable -> unobservable / gone = a change; unobservable before = no claim', () => {
  const k = 'Data\u0000A1';
  const base = { wrap: false, font: { bold: null, size: 11 }, align: { h: null, v: 0 }, border: 'UNOBSERVABLE:runtime-not-verified' };
  const o = cell => ({ cells: { [k]: cell } });
  assert.deepEqual(observedStyleChanges(o(base), o(base)).changed, []);
  assert.deepEqual(observedStyleChanges(o({ ...base, font: 'UNOBSERVABLE:empty-cell' }), o(base)).changed, [], 'the font of an empty cell becomes readable');
  assert.deepEqual(observedStyleChanges(o(base), o({ ...base, font: { bold: true, size: 11 } })).changed, [{ cell: 'Data!A1', keys: ['font.bold'] }]);
  assert.deepEqual(observedStyleChanges(o(base), o({ v: 'x' })).changed, [{ cell: 'Data!A1', change: 'style-unread-after' }]);
  assert.deepEqual(observedStyleChanges(o(base), o({ ...base, align: 'UNOBSERVABLE:threw' })).changed, [{ cell: 'Data!A1', keys: ['align.h', 'align.v'] }], 'a getter that throws after');
  assert.deepEqual(observedStyleChanges(o({ ...base, fill: 255 }), o(base)).changed, [{ cell: 'Data!A1', keys: ['fill'] }], 'one key gone after');
  assert.deepEqual(observedStyleChanges(o(base), o({ ...base, border: { t: { s: 7, c: 0 } } })).changed, [], 'an axis unobservable BEFORE makes no claim');
  assert.deepEqual(observedStyleChanges(o({ v: 'x' }), o(base)), { compared: 0, changed: [] }, 'no style read before: no claim');
});
