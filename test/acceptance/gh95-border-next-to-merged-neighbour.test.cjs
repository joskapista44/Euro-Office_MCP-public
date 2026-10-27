'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * GH #95 (owner request via alice; LIVE probe test/live/p92b-left-border-next-to-merge-probe.live.cjs reproduced it on
 * all four shapes): border.mode 'cells' sides ['left'] on F1:F2, whose LEFT neighbours E1/E2 are members of the merges
 * A1:E1 / A2:E2. The write draws F's left = the merge's right edge (shared), so E's right side legitimately changes.
 * The L3 ring check must accept that shared edge; before the fix the GH #83 merge-member branch of sideMayChange
 * narrowed the RING cell E1 to its own merge's requested perimeter sides (['left'] on column A only) and called the
 * shared right edge an unexpected change -> INTEGRAL_PRESERVATION_NOT_PROVEN (DISPATCHED_OUTCOME_UNKNOWN).
 */

const book = () => [
  { name: 'Data', cells: {
    A1: { v: 'label 1' }, B1: { v: '' }, C1: { v: '' }, D1: { v: '' }, E1: { v: '' }, F1: { v: 1 },
    A2: { v: 'label 2' }, B2: { v: '' }, C2: { v: '' }, D2: { v: '' }, E2: { v: '' }, F2: { v: 2 },
  }, merges: ['A1:E1', 'A2:E2'] },
];

async function run(store, steps) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, {});
  const t = { id: 'gh95', steps };
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, m: result.metrics };
}

const LEFT = { mode: 'cells', sides: ['left'], style: 'medium', color: '#000000' };
const step = border => ({ intent: 'format_range', sheet: 'Data', range: 'F1:F2', format: { border } });

test('GH #95: a left border next to a merged neighbour -- APPLIED (the shared edge on the merge member is expected), retry SATISFIED', async () => {
  const store = createStore(book());
  const { result } = await run(store, [step(LEFT)]);
  const op = result.execution.appliedOperations[0];
  assert.equal(op.status, OPERATION_STATUS.APPLIED, JSON.stringify(op.evidence || op).slice(0, 1500));
  assert.equal(result.ok, true);
  const again = await run(store, [step(LEFT)]);
  assert.equal(again.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
  assert.equal(again.m.commands, 0);
});

test('GH #95 control: the GH #83 target-member narrowing is unchanged -- a range CONTAINING the merge still writes only its perimeter', async () => {
  const store = createStore(book());
  const { result } = await run(store, [{ intent: 'format_range', sheet: 'Data', range: 'A1:E1', format: { border: { mode: 'cells', sides: ['top', 'bottom', 'left', 'right'], style: 'thin', color: '#000000' } } }]);
  assert.equal(result.ok, true);
  const c = a => store.durable.find(s => s.name === 'Data').cells[a].border;
  assert.equal(c('B1').l, undefined, 'interior edge untouched');
  assert.equal(c('C1').r, undefined, 'interior edge untouched');
});

test('GH #95 NEG control: the ring merge member\'s NON-shared side changing still fails L3 (the fix only admits the facing edge)', async () => {
  const store = createStore(book());
  let n = 0;
  const t = { id: 'gh95n', steps: [step(LEFT)] };
  // 'format-leaks': the SAME spec (a left border) also lands on E1 -- E1's LEFT side is NOT the shared edge with F1
  const driver = createFakeXlsxDriver(store, { fault: 'format-leaks', leakCell: 'E1' });
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  assert.notEqual(result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
  assert.equal(result.ok, false);
});
