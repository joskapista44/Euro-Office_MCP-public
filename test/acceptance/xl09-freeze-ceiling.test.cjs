'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const freeze = require('../../src/adapters/xlsx/families/freeze.cjs');
const W = require('../../src/server/write-tools.cjs');
const { req } = require('../fakes/write-wiring.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * X-L09 (owner GO "A", owner channel; LIVE: test/live/xlsx-xl09-freeze-above-cap-probe.live.cjs): the freeze_panes ceiling
 * was OUR implementation limit (1000 rows / 100 columns), not the engine's -- a freeze at 1001 / 2000 / 100000 rows, 101 / 200 / 1000 columns
 * and 1001 x 101 together lands in the saved file as exactly that pane. The ceiling is now the MEASURED range, 100000 rows / 1000 columns; ABOVE it
 * the plan (and the product WRITE tool schema) still refuse, named.
 */

const book = () => [{ name: 'Data', cells: { A1: { v: 1 } } }, { name: 'Home', cells: { A1: { v: 2 } } }];
const task = step => ({ id: 'xl09', steps: [{ intent: 'freeze_panes', sheet: 'Data', ...step }] });

async function run(store, t, script = {}) {
  let n = 0;
  const bodies = [];
  const driver = createFakeXlsxDriver(store, script);
  const dispatch = driver.dispatch.bind(driver);
  driver.dispatch = async (h, c) => { bodies.push(c && c.body); return dispatch(h, c); };
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, bodies, m: result.metrics };
}
const pane = (store, name = 'Data') => store.durable.find(s => s.name === name).pane;

for (const [label, spec, want] of [
  ['rows 1001 (just above the OLD ceiling)', { rows: 1001 }, { row0: 1001, col0: 0 }],
  ['rows 100000 (the new ceiling)', { rows: 100000 }, { row0: 100000, col0: 0 }],
  ['columns 101 (just above the OLD ceiling)', { columns: 101 }, { row0: 0, col0: 101 }],
  ['columns 1000 (the new ceiling)', { columns: 1000 }, { row0: 0, col0: 1000 }],
  ['1001 rows x 101 columns together', { rows: 1001, columns: 101 }, { row0: 1001, col0: 101 }],
  ['1000 x 100 (the OLD ceiling, unchanged)', { rows: 1000, columns: 100 }, { row0: 1000, col0: 100 }],
]) {
  test(`X-L09: freeze_panes ${label} -- APPLIED, the pane is exactly that; retry SATISFIED 0 commands`, async () => {
    const store = createStore(book());
    const a = await run(store, task(spec));
    assert.equal(a.result.ok, true, JSON.stringify(a.result.execution.appliedOperations[0].evidence));
    assert.equal(a.result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
    assert.deepEqual(pane(store), want);
    const again = await run(store, task(spec));
    assert.equal(again.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
    assert.equal(again.m.commands, 0);
  });
}

test('X-L09: the lowered body freezes at the measured address (the cell at column x row), nothing else changed', () => {
  const body = spec => freeze.lower(planOperations(task(spec)).operations[0]).body;
  assert.match(body({ rows: 100000, columns: 1000 }), /fp\.FreezeAt\(ws\.GetRange\("ALL100000"\)\)/);
  assert.match(body({ rows: 100000 }), /fp\.FreezeRows\(100000\)/);
  assert.match(body({ columns: 1000 }), /fp\.FreezeColumns\(1000\)/);
});

test('X-L09: ABOVE the new ceiling the plan still refuses, named, with the bound in the message', () => {
  for (const [spec, re] of [[{ rows: 100001 }, /rows must be an integer in \[1, 100000\]/], [{ columns: 1001 }, /columns must be an integer in \[1, 1000\]/], [{ rows: 100001, columns: 1001 }, /must be an integer in/]]) {
    const p = planOperations(task(spec));
    assert.equal(p.ok, false, JSON.stringify(spec));
    assert.match(p.detail, re, JSON.stringify(spec));
  }
});

test('X-L09: the product WRITE tool schema moves with the plan -- the new ceiling passes, one above is refused before any dispatch', () => {
  const validate = spec => W.validateWriteRequest(req([{ intent: 'freeze_panes', id: 'f', sheet: 'Data', ...spec }]));
  for (const spec of [{ rows: 100000 }, { columns: 1000 }, { rows: 100000, columns: 1000 }, { rows: 1 }]) assert.equal(validate(spec).ok, true, JSON.stringify(spec));
  for (const spec of [{ rows: 100001 }, { columns: 1001 }, { rows: 0 }, { columns: 0 }]) assert.equal(validate(spec).ok, false, JSON.stringify(spec));
});

test('X-L09 RED mutant: a freeze that lands one row off is still caught at the new ceiling (L1 reads the pane back)', async () => {
  const store = createStore(book());
  const a = await run(store, task({ rows: 100000 }), { fault: 'freeze-off-by-one' });
  assert.equal(a.result.ok, false);
});
