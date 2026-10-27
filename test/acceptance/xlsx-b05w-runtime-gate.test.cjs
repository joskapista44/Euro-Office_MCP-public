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
 *
 * DS9346 (owner D5=B, VALTAS, docs/DS9346-RUNTIME-RETIREMENT.md): the entire 9.3.4.60 family (/5 through /16) is
 * retired -- the accepted list is now the ONE 9.3.5.0/16 entry. The off-list examples below (9.3.4.60/4, /5, /17,
 * /2 -- never accepted, now simply joined by every other 9.3.4.60 revision) are UNCHANGED: their role was already
 * "not on the accepted list," which this retirement only widens.
 */

const RT6 = '9.3.5|euro-office-bridge-v1/9.3.5.0/16';
const RT17 = '9.3.5|euro-office-bridge-v1/9.3.5.0/17';
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

// DS9346: every one of these is simply not the one accepted runtime (9.3.5.0/16) -- a single-entry
// accepted list has no "safely unclaimed number" to chase as the list grows.
for (const [label, spec] of [['indent', INDENT], ['shrink', SHRINK]]) {
  for (const rt of ['9.3.4|euro-office-bridge-v1/9.3.4.60/5', '9.3.4|euro-office-bridge-v1/9.3.4.60/4', '9.3.4|euro-office-bridge-v1/9.3.4.60/17', '9.3.4|euro-office-bridge-v1/9.3.4.60/2']) {
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

// GH #326/#327: the format-side and driver-side B05 lists stay pinned to each other, now both /16 + /17.
test('the driver second line: the SAME two-entry (9.3.5.0/16 + /17) list; a B05 command is refused WITHOUT an editor elsewhere; an unmarked one is not', async () => {
  assert.deepEqual([...format.B05_BRIDGE_WRITE_RUNTIMES], [...drv.B05_BRIDGE_WRITE_RUNTIMES]);
  assert.deepEqual([...drv.B05_BRIDGE_WRITE_RUNTIMES], [RT6, RT17]);
  const d = drv.createEuroOfficeXlsxDriver({ url: 'u', user: 'u', pass: 'p', loadPlaywright: () => ({ ok: false }) });
  const lowered = format.lower(planOperations(task(INDENT)).operations[0], { sheet: 'Data' });
  const r = await d.dispatch(null, lowered);
  assert.deepEqual([r.ok, r.dispatched, r.outcome], [false, false, 'b05-bridge-write-runtime-not-proven']);
  assert.equal(format.lower(planOperations(task({ align: { horizontal: 'left' } })).operations[0], { sheet: 'Data' }).meta.bridgeAlignExtWrite, false);
});
