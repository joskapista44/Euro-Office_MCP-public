'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { SessionHandle } = require('../../../src/core/session/handle.cjs');
const { Metrics } = require('../../../src/core/metrics/collect.cjs');
const { withSession } = require('../../../src/core/session/lifecycle.cjs');
const { createFakeSession } = require('../../fakes/fake-session.cjs');
const { createFakeAdapter } = require('../../fakes/fake-adapter.cjs');
const { createFakeExecutor } = require('../../fakes/fake-executor.cjs');

/** Owner requirement 16: exactly one core owner opens and closes a session. */

test('the handle an adapter receives exposes no lifecycle method', () => {
  const handle = new SessionHandle({ id: 'h', fileId: 'f', adapterKind: 'fake', metrics: new Metrics() });
  for (const forbidden of ['open', 'close', 'openSession', 'closeSession', 'persist', 'forceSave', 'canSave', 'save']) {
    assert.equal(typeof handle[forbidden], 'undefined', `the handle must not expose ${forbidden}`);
  }
});

test('the driver is not reachable from the handle itself', () => {
  const handle = new SessionHandle({ id: 'h', fileId: 'f', adapterKind: 'fake', metrics: new Metrics() });
  const exposed = Object.keys(handle).concat(Object.getOwnPropertyNames(handle));
  for (const key of exposed) {
    const value = handle[key];
    if (value && typeof value === 'object' && typeof value.forceSave === 'function') {
      assert.fail(`the handle exposes the transport driver through "${key}"`);
    }
  }
});

test('the handle is sealed — an adapter cannot bolt a lifecycle method onto it', () => {
  const handle = new SessionHandle({ id: 'h', fileId: 'f', adapterKind: 'fake', metrics: new Metrics() });
  assert.throws(() => { 'use strict'; handle.close = () => {}; }, TypeError);
});

test('one task opens exactly one session and closes it exactly once', async () => {
  let t = 0;
  const driver = createFakeSession({});
  const adapter = createFakeAdapter({});
  const task = { id: 'ownership', operations: [{ id: 'op1' }, { id: 'op2' }, { id: 'op3' }] };

  const result = await withSession(
    { driver, fileId: 'f', now: () => (t += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } },
    adapter,
    createFakeExecutor(task, {})
  );

  assert.equal(driver.calls.open, 1, 'OPEN_COUNT');
  assert.equal(driver.calls.close, 1, 'CLOSE_COUNT');
  assert.equal(result.metrics.operationsApplied, 3, 'three operations shared one session — no per-operation reopen');
});

test('the executor is handed a handle, never the ability to open another session', async () => {
  let t = 0;
  const driver = createFakeSession({});
  let seen = null;

  await withSession(
    { driver, fileId: 'f', now: () => (t += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } },
    createFakeAdapter({}),
    async (handle, ctx) => {
      seen = { handle, ctxKeys: Object.keys(ctx) };
      const { buildExecutionResult } = require('../../../src/core/execution/result.cjs');
      return buildExecutionResult({ executionOk: true, outcome: 'noop', appliedOperations: ctx.appliedOperations });
    }
  );

  assert.equal(typeof seen.handle.close, 'undefined');
  assert.deepEqual(seen.ctxKeys.sort(), ['adapter', 'appliedOperations', 'lifecycle']);
});
