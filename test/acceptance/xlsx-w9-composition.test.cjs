'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * M6 wave 9 composition (offline): several families in ONE task on the fake. What composes is planned and
 * verified per step; the reopened retry is SATISFIED 0/0 for every step. What would make the terminal state depend
 * on the step order is refused at PLANNING -- by cell overlap, or by a shared object (a sheet's page layout, a chart
 * name); a sheet-level or object step claims no cells.
 */

const book = () => [
  { name: 'K', cells: { A1: { v: 'cat' }, B1: { v: 's1' }, A2: { v: 'a' }, B2: { v: 1 }, A3: { v: 'b' }, B3: { v: 2 } }, objects: { charts: [{ name: 'Old', type: 'bar', title: 't', series: [{ name: '=K!$B$1', values: '=K!$B$2:$B$3', cat: '=K!$A$2:$A$3', x: '' }] }], pivots: [], names: [] } },
  { name: 'O', cells: { A1: { v: 'keep' } } },
];
const W = { intent: 'write_range', sheet: 'O', range: 'C1:D1', values: [['x', 2]] };
const N = { intent: 'set_defined_name', name: 'W9N', refersTo: '=O!$C$1' };
const C = { intent: 'set_chart', sheet: 'K', name: 'New', range: 'A1:B3', type: 'bar', title: 'T' };
const P = { intent: 'set_page_layout', sheet: 'K', orientation: 'landscape' };
const task = (...steps) => ({ id: 'm6-w9', steps: steps.map((s, i) => ({ id: `s${i}`, ...s })) });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}

test('four families in one task: APPLIED, every step verified; the reopened retry SATISFIED 0/0 for all four', async () => {
  const store = createStore(book());
  const t = task(W, N, C, P);
  const a = await run(store, t, { rules: true });
  assert.equal(a.result.ok, true, JSON.stringify([a.result.execution.primitiveVerification, a.result.execution.preservationVerification, a.result.execution.appliedOperations.map(o => o.evidence)]));
  assert.equal(a.m.commands, 4);
  const r = await run(store, t, { rules: true });
  assert.equal(r.result.ok, true);
  assert.equal(r.m.commands, 0);
  assert.deepEqual(r.result.execution.appliedOperations.map(o => o.status), Array(4).fill(OPERATION_STATUS.SATISFIED));
});

test('what composes: defining a name and using it; two formulas using one name; deleting a name no formula uses', () => {
  for (const t of [
    task({ intent: 'set_defined_name', name: 'NX', refersTo: '=O!$A$1' }, { intent: 'write_range', sheet: 'O', range: 'B1', formulas: [['=NX*2']] }),
    task({ intent: 'write_range', sheet: 'O', range: 'B1', formulas: [['=NX*2']] }, { intent: 'write_range', sheet: 'O', range: 'B2', formulas: [['=NX+1']] }),
    task({ intent: 'delete_defined_name', name: 'NX' }, { intent: 'delete_defined_name', name: 'NY' }, { intent: 'write_range', sheet: 'O', range: 'B1', formulas: [['=A1*2']] }),
  ]) assert.equal(planOperations(t).ok, true, JSON.stringify([t, planOperations(t).detail]));
});

test('a sheet-level step claims no cells: set_page_layout + a write to A1 of the same sheet compose', () => {
  const p = planOperations(task(P, { intent: 'write_range', sheet: 'K', range: 'A1', values: [['cat']] }));
  assert.equal(p.ok, true, p.detail);
});

for (const [name, t, re] of [
  ['two layouts of one sheet', task(P, { ...P, orientation: 'portrait' }), /on page:K$/],
  ['a rename TO a name another step creates (disjoint ranges)', task({ intent: 'rename_chart', sheet: 'K', name: 'Old', newName: 'New' }, { ...C, range: 'D5:E7' }), /on chart:K:New$/],
  ['a rename FROM a name another step creates', task({ ...C, name: 'Old' }, { intent: 'rename_chart', sheet: 'K', name: 'Old', newName: 'Z' }), /on (K|chart:K:Old)$/],
  ['two creates of one chart name over different ranges', task(C, { ...C, range: 'D5:E7' }), /on chart:K:New$/],
  ['a write into the data a chart step reads', task(C, { intent: 'write_range', sheet: 'K', range: 'B2', values: [[9]] }), /on K/],
  ['a formula using a name another step deletes', task({ intent: 'delete_defined_name', name: 'NX' }, { intent: 'write_range', sheet: 'O', range: 'B1', formulas: [['=nx*2']] }), /s0 and s1 on name:nx$/],
  ['a name another step deletes, built at runtime (INDIRECT: any name may be used)', task({ intent: 'delete_defined_name', name: 'NX' }, { intent: 'write_range', sheet: 'O', range: 'B1', formulas: [['=INDIRECT("N"&"X")']] }), /s0 and s1 on name:nx$/],
  ['a conditional format whose condition names a name another step deletes', task({ intent: 'delete_defined_name', name: 'NX' }, { intent: 'add_conditional_format', sheet: 'O', range: 'A1:A3', rule: { type: 'cellValue', operator: 'greater', formula1: 'NX', fill: '#FF0000' } }), /s0 and s1 on name:nx$/],
  ['the same, the other order', task({ intent: 'write_range', sheet: 'O', range: 'B1', formulas: [['=SUM(NX)']] }, { intent: 'delete_defined_name', name: 'NX' }), /s0 and s1 on name:nx$/],
]) {
  test(`refused at PLANNING (order-dependent): ${name}`, () => {
    const p = planOperations(t);
    assert.equal(p.ok, false, JSON.stringify(t));
    assert.match(p.detail, /^overlapping-writes-refused/);
    assert.match(p.detail, re);
  });
}

test('the one-step families stay alone in a composition', () => {
  for (const s of [{ intent: 'create_sheet', name: 'Z' }, { intent: 'freeze_panes', sheet: 'K', rows: 1 }, { intent: 'copy_range', sheet: 'K', range: 'A1:B3', target: 'D1' }]) {
    const p = planOperations(task(s, W));
    assert.equal(p.ok, false, JSON.stringify(s));
    assert.match(p.detail, s.intent === 'freeze_panes' ? /LAST step/ : /only step of its task/); // M7 S5 S23: a freeze may close a task
  }
});

test('a large mixed batch (36 writes + a name over them + a chart + a layout + a formula using the name): 40 APPLIED, the retry 40 SATISFIED 0/0', async () => {
  const store = createStore(book());
  const steps = [];
  for (let i = 1; i <= 36; i += 1) steps.push({ intent: 'write_range', sheet: 'O', range: `C${i}:D${i}`, values: [[`r${i}`, i]] });
  steps.push({ intent: 'set_defined_name', name: 'BIG', refersTo: '=O!$C$1:$D$36' }, C, { ...P, sheet: 'O' }, { intent: 'write_range', sheet: 'K', range: 'D1', formulas: [['=SUM(BIG)']] });
  const t = task(...steps);
  const a = await run(store, t, { rules: true });
  assert.equal(a.result.ok, true, JSON.stringify(a.result.execution.primitiveVerification.filter(v => v.verdict !== 'PASS')));
  assert.equal(a.m.commands, 40);
  const r = await run(store, t, { rules: true });
  assert.equal(r.m.commands, 0);
  assert.equal(r.result.execution.appliedOperations.filter(o => o.status === OPERATION_STATUS.SATISFIED).length, 40);
});
