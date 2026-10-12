'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const format = require('../../src/adapters/xlsx/families/format.cjs');
const drv = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * B05-WRITE explicit runtime gate. History: added 2026-10-09 (owner GO the owner, owner channel; internal tracker) as the
 * runtime list of the direct-internal setters; widened to [/5, /6] by M11.5a.2. M11.5a.4 (owner GO owner channel,
 * internal tracker, DS5_POLICY=B05_ONLY_FAIL_CLOSED): B05-WRITE is a Bridge WRITE now, the capability exists only
 * on /6, so THIS write path alone is /6-only -- refused after the prewrite READ, before any write, everywhere else.
 * The capability contract itself: xlsx-b05w-bridge-write-cutover.test.cjs.
 */

const RT6 = '9.3.4|euro-office-bridge-v1/9.3.4.60/6';
const RT7 = '9.3.4|euro-office-bridge-v1/9.3.4.60/7';
const RT8 = '9.3.4|euro-office-bridge-v1/9.3.4.60/8';
const book = () => [{ name: 'Data', cells: { A1: { v: 'x' }, B1: { v: 'y' } }, merges: [] }];
const task = spec => ({ id: 'b05w', steps: [{ intent: 'format_range', sheet: 'Data', range: 'A1', format: spec }] });
const INDENT = { align: { horizontal: 'left', indent: 2 } };
const SHRINK = { align: { shrink: true } };

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}
const op = r => r.execution.appliedOperations[0];

for (const [label, spec] of [['indent', INDENT], ['shrink', SHRINK]]) {
  for (const rt of ['9.3.4|euro-office-bridge-v1/9.3.4.60/5', '9.3.4|euro-office-bridge-v1/9.3.4.60/4', '9.3.4|euro-office-bridge-v1/9.3.4.60/9', '9.3.4|euro-office-bridge-v1/9.3.4.60/2']) {
    test(`${label} on ${rt}: FAIL_CLOSED before the write -- 0 dispatch / 0 command / 0 Bridge write, nothing changed`, async () => {
      const store = createStore(book()); const before = JSON.stringify(store.durable);
      const { result, driver, m } = await run(store, task(spec), { runtime: rt });
      assert.equal(result.ok, false);
      assert.equal(op(result).status, OPERATION_STATUS.FAILED);
      assert.equal(op(result).evidence.evidence.reason, 'xlsx-align-ext-bridge-write-runtime-not-proven');
      assert.deepEqual([driver.calls.dispatch, m.writes, m.commands, driver.calls.bridgeWrites], [0, 0, 0, 0]);
      assert.equal(JSON.stringify(store.durable), before, 'DOCUMENT_MUTATION=0');
    });
  }

  test(`POSITIVE CONTROL: ${label} on /6 -> APPLIED once through ONE Bridge WRITE, then a fresh retry SATISFIED 0/0`, async () => {
    const store = createStore(book());
    const first = await run(store, task(spec), { runtime: RT6 });
    assert.equal(op(first.result).status, OPERATION_STATUS.APPLIED, JSON.stringify(first.result.execution));
    assert.deepEqual([first.m.commands, first.driver.calls.bridgeWrites], [1, 1]);
    const again = await run(store, task(spec), { runtime: RT6 });
    assert.equal(op(again.result).status, OPERATION_STATUS.SATISFIED);
    assert.deepEqual([again.m.writes, again.m.commands, again.driver.calls.bridgeWrites], [0, 0, 0]);
  });
}

test('the gate decides only a WRITE: an already-holding state stays SATISFIED 0/0 on /5', async () => {
  const b = book(); b[0].cells.A1.shrink = true;
  const { result, driver } = await run(createStore(b), task(SHRINK), { runtime: '9.3.4|euro-office-bridge-v1/9.3.4.60/5' });
  assert.equal(op(result).status, OPERATION_STATUS.SATISFIED);
  assert.equal(driver.calls.dispatch, 0);
});

test('scope (B05_ONLY): a PUBLIC align.horizontal on /5 is untouched by this gate', async () => {
  const { result, driver } = await run(createStore(book()), task({ align: { horizontal: 'right' } }), { runtime: '9.3.4|euro-office-bridge-v1/9.3.4.60/5' });
  assert.equal(op(result).status, OPERATION_STATUS.APPLIED);
  assert.equal(driver.calls.dispatch, 1);
});

test('the driver second line: the SAME /6+/7+/8 list; a B05 command is refused WITHOUT an editor elsewhere; an unmarked one is not', async () => {
  assert.deepEqual([...format.B05_BRIDGE_WRITE_RUNTIMES], [...drv.B05_BRIDGE_WRITE_RUNTIMES]);
  assert.deepEqual([...drv.B05_BRIDGE_WRITE_RUNTIMES], [RT6, RT7, RT8]);
  const d = drv.createEuroOfficeXlsxDriver({ url: 'u', user: 'u', pass: 'p', loadPlaywright: () => ({ ok: false }) });
  const lowered = format.lower(planOperations(task(INDENT)).operations[0], { sheet: 'Data' });
  const r = await d.dispatch(null, lowered);
  assert.deepEqual([r.ok, r.dispatched, r.outcome], [false, false, 'b05-bridge-write-runtime-not-proven']);
  assert.equal(format.lower(planOperations(task({ align: { horizontal: 'left' } })).operations[0], { sheet: 'Data' }).meta.bridgeAlignExtWrite, false);
});
