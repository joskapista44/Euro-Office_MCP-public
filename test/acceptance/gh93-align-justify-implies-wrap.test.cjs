'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * GH #93 follow-up (release rehearsal 1.0.1-73f0e4023bd9: align justify DISPATCHED_OUTCOME_UNKNOWN twice, retry
 * SATISFIED): SetAlignHorizontal('justify') also turns the cell's wrap ON (LIVE-measured,
 * test/live/justify-text-cell-probe.live.cjs: h 3 + wrap true on text / longer text / number / empty cells, in the
 * session AND after a reopen). The planner now expects that coupled wrap:true for a horizontal justify, so the
 * primitive verification no longer reads it as an unrequested change.
 */

const book = () => [{ name: 'Data', cells: { B3: { v: 'y' }, C3: { v: 5 } }, merges: [] }];
async function run(store, align) {
  let n = 0;
  const t = { id: 'gh93j', steps: [{ intent: 'format_range', sheet: 'Data', range: 'B3:C3', format: { align } }] };
  const driver = createFakeXlsxDriver(store, {});
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, m: result.metrics };
}
const cell = (store, a) => store.durable.find(s => s.name === 'Data').cells[a];

test('GH #93: horizontal justify -- APPLIED (the coupled wrap:true is expected), retry SATISFIED, 0 commands', async () => {
  const store = createStore(book());
  const { result } = await run(store, { horizontal: 'justify' });
  const op = result.execution.appliedOperations[0];
  assert.equal(op.status, OPERATION_STATUS.APPLIED, JSON.stringify(op.evidence || op).slice(0, 1200));
  assert.equal(cell(store, 'B3').wrap, true);
  const again = await run(store, { horizontal: 'justify' });
  assert.equal(again.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
  assert.equal(again.m.commands, 0);
});

test('GH #93 control: center does NOT expect a wrap change (wrap stays false)', async () => {
  const store = createStore(book());
  const { result } = await run(store, { horizontal: 'center' });
  assert.equal(result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
  assert.notEqual(cell(store, 'B3').wrap, true);
});
