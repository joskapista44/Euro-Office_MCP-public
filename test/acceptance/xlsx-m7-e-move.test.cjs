'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/** M7 E R7 (tier E probe m7-e-probe1-20261003.json MV / MX): a move of a referenced block, the measured shapes only. */

const book = (extraP = {}, extraQ = {}) => [
  { name: 'P', cells: { H1: { v: 1 }, H2: { v: 3, f: '=H1*3' }, C1: { v: 4, f: '=SUM(H1:H2)' }, C2: { v: 6, f: '=H2*2' }, C3: { v: 2, f: '=$H$1+1' }, C4: { v: 2, f: '=COUNT(N1)' }, ...extraP } },
  { name: 'Q', cells: { B1: { v: 4, f: '=P!H2+1' }, A9: { v: 'keep' }, ...extraQ } },
];
const NAMES = () => [{ name: 'N1', ref: 'P!$H$1:$H$2' }];
const CELLS = [[{ t: 'number', v: '1' }], [{ t: 'formula', f: '=H1*3' }]];
const move = (target, { targetSheet = 'P', cells = CELLS } = {}) => {
  const end = target.replace(/\d+/, n => String(Number(n) + 1));
  return { id: 'm7-e-r7', steps: [{ intent: 'move_range', sheet: 'P', range: 'H1:H2', target, ...(targetSheet !== 'P' ? { targetSheet } : {}),
    expect: { version: 1, fileId: 'disposable', source: { sheet: 'P', range: 'H1:H2' }, target: { sheet: targetSheet, range: `${target}:${end}` }, cells } }] };
};

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}
const S = (store, n) => store.durable.find(s => s.name === n).cells;
const op = result => result.execution.appliedOperations[0];
const reason = result => op(result).evidence.evidence.reason;
const withNames = b => { const s = createStore(b); s.names = NAMES(); return s; };

test('R7: a block referenced in every measured shape moves; every reference and the name follow; retry SATISFIED 0/0', async () => {
  const store = withNames(book());
  const a = await run(store, move('J5'));
  assert.equal(a.result.ok, true, JSON.stringify([op(a.result).evidence, a.result.execution.primitiveVerification, a.result.execution.preservationVerification]));
  assert.equal(a.m.commands, 1);
  assert.deepEqual([S(store, 'P').C1.f, S(store, 'P').C2.f, S(store, 'P').C3.f, S(store, 'Q').B1.f, S(store, 'P').J6.f], ['=SUM(J5:J6)', '=J6*2', '=$J$5+1', '=P!J6+1', '=J5*3']);
  assert.equal(store.names[0].ref, 'P!$J$5:$J$6');
  const again = await run(store, move('J5'));
  assert.equal(again.m.commands, 0);
  assert.equal(op(again.result).status, OPERATION_STATUS.SATISFIED);
});

test('R7: a cross-sheet move gives a bare reference the target sheet', async () => {
  const b = book();
  b[0].cells = { H1: { v: 1 }, H2: { v: 3 }, C2: { v: 6, f: '=H2*2' } };
  b[1].cells = { A9: { v: 'keep' } };
  const store = createStore(b);
  const cells = [[{ t: 'number', v: '1' }], [{ t: 'number', v: '3' }]];
  const a = await run(store, move('D1', { targetSheet: 'Q', cells }));
  assert.equal(a.result.ok, true, JSON.stringify([op(a.result).evidence, a.result.execution.preservationVerification]));
  assert.equal(S(store, 'P').C2.f, '=Q!D2*2');
});

test('R7 RED: the editor leaves the outside references and the name where they were -> L3 FAIL, not APPLIED', async () => {
  const { result } = await run(withNames(book()), move('J5'), { fault: 'move-no-repoint' });
  assert.equal(result.ok, false);
  const ev = result.execution.preservationVerification[0].evidence.evidence;
  assert.ok(ev.repointedWrong.length > 0, JSON.stringify(ev));
});

test('R7 RED: the editor leaves an outside reference where it was (no name involved) -> L3 FAIL on that cell', async () => {
  const b = book();
  delete b[0].cells.C4;
  const { result } = await run(createStore(b), move('J5'), { fault: 'move-no-repoint' });
  assert.equal(result.ok, false);
  const ev = result.execution.preservationVerification[0].evidence.evidence;
  assert.deepEqual([ev.dependentsKept, ev.repointedWrong.map(x => x.cell).sort()], [true, ['P!C1', 'P!C2', 'P!C3', 'Q!B1']]);
});

for (const [name, b, t, why] of [
  ['a reference that only partly overlaps the block', book({ C5: { v: 4, f: '=SUM(H1:H3)' } }), move('J5'), 'partial-or-whole-line-reference'],
  ['a whole-row reference into the block', book({ C5: { v: 4, f: '=SUM(1:1)' } }), move('J5'), 'partial-or-whole-line-reference'],
  ['a sheet-qualified reference into a block moved to another sheet', (() => { const x = book(); x[0].cells = { H1: { v: 1 }, H2: { v: 3 } }; return x; })(), move('D1', { targetSheet: 'Q', cells: [[{ t: 'number', v: '1' }], [{ t: 'number', v: '3' }]] }), 'qualified-reference-into-a-block-moved-to-another-sheet'],
  ['an inner reference of a block moved to another sheet', (() => { const x = book(); x[0].cells = { H1: { v: 1 }, H2: { v: 3, f: '=H1*3' } }; x[1].cells = { A9: { v: 'keep' } }; return x; })(), move('D1', { targetSheet: 'Q' }), 'inner-reference-of-a-block-moved-to-another-sheet'],
]) {
  test(`R7 refused before the write: ${name}`, async () => {
    const r = await run(createStore(b), t);
    assert.equal(r.driver.calls.dispatch, 0, JSON.stringify(op(r.result).evidence));
    assert.equal(reason(r.result), 'xlsx-move-reference-shape-not-measured');
    assert.equal(op(r.result).evidence.evidence.cells[0].why, why);
  });
}
