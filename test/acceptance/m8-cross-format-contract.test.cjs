'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { PERSISTENCE_STATE } = require('../../src/core/persistence/verified-writes.cjs');
const docx = require('../../src/adapters/docx/adapter.cjs');
const xlsx = require('../../src/adapters/xlsx/adapter.cjs');
const FD = require('../fakes/fake-docx-driver.cjs');
const { normalizeInventory } = require('../../src/adapters/docx/inventory.cjs');
const FX = require('../fakes/fake-xlsx-driver.cjs');

/**
 * M8 cross-format (docs/M8-CROSS-FORMAT-REGRESSION-PLAN.md, matrix rows 1-6): the SAME engine contract asserted through
 * BOTH real adapters, on their fake transports. The engine-only tests (lifecycle-matrix, uncertain-persistence,
 * m7-save-contract) prove it with a fake adapter; this file proves no adapter bends it. The DOCX fake runs with
 * `saveObserver` -- the real DOCX driver's save-ack pairing (M7, PR #17) -- so both formats reach FORCESAVE_CONFIRMED.
 */

const FULL_WRITE_TRACE = ['CREATED', 'OPENING', 'READY', 'OBSERVING', 'EXECUTING', 'VERIFYING', 'PERSISTING', 'CLOSING', 'CLOSED', 'FINALIZED'];
const NO_BARRIER_TRACE = ['CREATED', 'OPENING', 'READY', 'OBSERVING', 'EXECUTING', 'VERIFYING', 'CLOSING', 'CLOSED', 'FINALIZED'];
const OPEN_FAIL_TRACE = ['CREATED', 'OPENING', 'CLOSING', 'CLOSED', 'FINALIZED'];

const RICH = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const FORMATS = {
  DOCX: {
    store: () => FD.createStore([{ text: 'Para one.', props: { ...RICH } }, { text: 'Para two.', props: { ...RICH } }]),
    driver: (store, script) => FD.createFakeDocxDriver(store, { saveObserver: true, ...script }),
    adapter: docx.docxAdapter,
    task: { id: 'm8-d', steps: [{ intent: 'format_text', selector: { text: 'Para two.' }, format: { bold: true } }] },
    options: () => undefined,
  },
  XLSX: {
    store: () => FX.createStore([{ name: 'Data', cells: { A1: { v: 'x' } } }]),
    driver: (store, script) => FX.createFakeXlsxDriver(store, script),
    adapter: xlsx.xlsxAdapter,
    task: { id: 'm8-x', steps: [{ intent: 'write_range', sheet: 'Data', range: 'B2', values: [[5]] }] },
    options: t => ({ observation: { live: xlsx.liveScope(t) } }),
  },
};

async function run(f, store, script = {}) {
  let t = 0;
  const driver = f.driver(store, script);
  const options = { driver, fileId: 'm8', now: () => (t += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, f.adapter, createExecutor(f.task, f.options(f.task)));
  return { result, driver, m: result.metrics };
}

for (const [name, f] of Object.entries(FORMATS)) {
  test(`${name}: write -> the full lifecycle trace, FORCESAVE_CONFIRMED; reopened retry -> no-barrier trace, 0 / 0, NOT_REQUIRED`, async () => {
    const store = f.store();
    const first = await run(f, store);
    assert.equal(first.result.ok, true, first.result.outcome);
    assert.deepEqual(first.result.lifecycleTrace, FULL_WRITE_TRACE);
    assert.equal(first.result.persistenceState, PERSISTENCE_STATE.FORCESAVE_CONFIRMED);
    assert.deepEqual([first.m.writes, first.m.commands], [1, 1]);
    const second = await run(f, store);
    assert.equal(second.result.ok, true, second.result.outcome);
    assert.deepEqual(second.result.lifecycleTrace, NO_BARRIER_TRACE);
    assert.deepEqual([second.m.writes, second.m.commands, second.driver.calls.forceSave], [0, 0, 0]);
    assert.equal(second.result.persistenceState, PERSISTENCE_STATE.NOT_REQUIRED);
  });

  test(`${name}: open failure -> session-open-failed, no close, no write, the open-fail trace`, async () => {
    const { result, driver, m } = await run(f, f.store(), { openFails: true });
    assert.equal(result.ok, false);
    assert.equal(result.outcome, 'session-open-failed');
    assert.deepEqual([m.writes, m.commands, driver.calls.close, driver.calls.forceSave], [0, 0, 0, 0]);
    assert.deepEqual(result.lifecycleTrace, OPEN_FAIL_TRACE);
  });

  test(`${name}: close failure -> the document / persistence truth kept, cleanup FAILED, ok:false, task-cleanup-failed`, async () => {
    const { result, driver } = await run(f, f.store(), { closeFails: true });
    assert.equal(result.ok, false);
    assert.equal(result.outcome, 'task-cleanup-failed');
    assert.equal(result.documentOutcome, 'SUCCESS');
    assert.equal(result.cleanupOutcome, 'FAILED');
    assert.equal(driver.calls.close, 1);
    assert.deepEqual(result.lifecycleTrace, FULL_WRITE_TRACE);
  });

  test(`${name}: the dispatch throws -> UNCERTAIN_MAY_PERSIST, no barrier, no forceSave, the durable store untouched`, async () => {
    const store = f.store();
    const before = JSON.stringify(store.durable);
    const { result, driver, m } = await run(f, store, { fault: 'throw' });
    assert.equal(result.ok, false);
    assert.equal(result.persistenceState, PERSISTENCE_STATE.UNCERTAIN_MAY_PERSIST);
    assert.deepEqual([m.writes, driver.calls.forceSave, driver.calls.close], [0, 0, 1]);
    assert.deepEqual(result.lifecycleTrace, NO_BARRIER_TRACE);
    assert.equal(JSON.stringify(store.durable), before);
  });
}

/** Rows 3 and 4: each format's OWN task shape for the branch (the one its existing suite measured). */
const docxParas = () => [{ text: 'Para one.', props: { ...RICH } }, { text: 'Para two.', props: { ...RICH } }];
const docxSelector = () => {
  const b = normalizeInventory({ html: FD.htmlOf(docxParas()), paragraphs: docxParas() }).blocks.find(x => x.text === 'Para two.');
  return { text: 'Para two.', kind: b.kind, before: b.locator.context.before, after: b.locator.context.after };
};
const BRANCH = {
  DOCX: {
    l3: { task: () => ({ id: 'm8-d-l3', steps: [{ intent: 'rewrite_paragraph', selector: docxSelector(), text: 'Para two, new.' }] }), script: { fault: 'drop-props' } },
    prewrite: { task: () => FORMATS.DOCX.task, script: { propsUnreadable: true } },
  },
  XLSX: {
    l3: { task: () => FORMATS.XLSX.task, script: { fault: 'side-effect' } },
    prewrite: { task: () => ({ id: 'm8-x-pre', steps: [{ intent: 'clear_range', sheet: 'Data', range: 'A1' }] }), script: { boldUnobservable: 'always' } },
  },
};
async function runTask(f, task, script) {
  let t = 0;
  const store = f.store();
  const before = JSON.stringify(store.durable);
  const driver = f.driver(store, script);
  const options = { driver, fileId: 'm8', now: () => (t += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, f.adapter, createExecutor(task, f.options(task)));
  return { result, driver, m: result.metrics, untouched: JSON.stringify(store.durable) === before };
}

for (const [name, f] of Object.entries(FORMATS)) {
  test(`${name}: a post-dispatch L3 (preservation) FAIL -> UNCERTAIN_MAY_PERSIST, no forceSave, the durable store untouched`, async () => {
    const b = BRANCH[name].l3;
    const { result, driver, untouched } = await runTask(f, b.task(), b.script);
    assert.equal(result.ok, false);
    assert.match(result.outcome, /integral_preservation_not_proven/);
    assert.equal(result.persistenceState, PERSISTENCE_STATE.UNCERTAIN_MAY_PERSIST);
    assert.equal(driver.calls.forceSave, 0);
    assert.equal(untouched, true);
  });

  test(`${name}: an unobservable required axis -> refused BEFORE the write: 0 dispatch / 0 write / 0 command, NOT_REQUIRED`, async () => {
    const b = BRANCH[name].prewrite;
    const { result, driver, m, untouched } = await runTask(f, b.task(), b.script);
    assert.equal(result.ok, false);
    assert.deepEqual([driver.calls.dispatch, m.writes, m.commands, driver.calls.forceSave], [0, 0, 0, 0]);
    assert.equal(result.persistenceState, PERSISTENCE_STATE.NOT_REQUIRED);
    assert.equal(untouched, true);
  });
}

test('the two adapters agree on every engine outcome of these four branches (the cross-format invariant itself)', async () => {
  const outcomes = {};
  for (const [name, f] of Object.entries(FORMATS)) {
    outcomes[name] = [];
    for (const script of [{}, { openFails: true }, { closeFails: true }, { fault: 'throw' }]) {
      const { result } = await run(f, f.store(), script);
      outcomes[name].push([result.ok, result.outcome.replace(/dispatch_threw$/, 'X'), result.persistenceState, result.cleanupOutcome, result.lifecycleTrace.join('>')]);
    }
  }
  assert.deepEqual(outcomes.DOCX, outcomes.XLSX);
});
