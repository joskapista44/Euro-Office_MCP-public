'use strict';

/**
 * M11.4a.2 acceptance (offline): the WITHHELD product WRITE tool `xlsx_execute_task` wired to the CANONICAL executor
 * (owner decision, owner channel 2026-10-10):
 *
 *   handler -> execute -> shared session gate -> runtime check -> product WRITE driver wrapper -> broker configSource
 *   -> withSession -> xlsxAdapter -> createExecutor(task, sequential) -> canonical verification / persistence
 *   -> mapExecutionResult
 *
 * The transport is the fake XLSX editor (test/fakes/fake-xlsx-driver.cjs) behind the createDriver seam, so the canonical
 * lifecycle, executor, verification layers and persistence barrier run for real; the broker-route cases use the REAL
 * product driver with a recording fake broker. Nothing here is a LIVE measurement and nothing is exposed.
 * Set M11_4A2_EVIDENCE=<file> to also write the measured matrices as JSON.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');

const W = require('../../src/server/write-tools.cjs');
const { buildServerTools } = require('../../src/server/server-tools.cjs');
const { createSessionGate } = require('../../src/server/session-gate.cjs');
const { createBrokerConfigSource } = require('../../src/server/broker-config-source.cjs');
const { exposedTools } = require('../../src/server/tools.cjs');
const { harness, req, createStore, book, seedChart } = require('../fakes/write-wiring.cjs');
const { REFUSALS } = require('../fakes/write-refusals.cjs');
const { BROKER, TOKEN, editConfig, xlsxEditor } = require('../fakes/read-editors.cjs');

const EVIDENCE = {};
test.after(() => { if (process.env.M11_4A2_EVIDENCE) fs.writeFileSync(process.env.M11_4A2_EVIDENCE, `${JSON.stringify(EVIDENCE, null, 1)}\n`); });

const STEP = Object.freeze({
  write_range: { intent: 'write_range', sheet: 'Budget', range: 'F1:G2', values: [['x', 1], ['y', 2]] },
  insert_rows: { intent: 'insert_rows', sheet: 'Budget', at: 3, count: 2, anchor: { column: 'A', value: 'pear' } },
  delete_rows: { intent: 'delete_rows', sheet: 'Budget', at: 2, count: 1, anchor: { column: 'A', value: 'pear' } },
  rename_sheet: { intent: 'rename_sheet', sheet: 'Budget', name: 'Budget_2026' },
  unmerge_range: { intent: 'unmerge_range', sheet: 'Budget', range: 'C1:D1' },
  set_chart_title: { intent: 'set_chart_title', sheet: 'K', name: 'R1', title: 'Revenue' },
});
const fresh = () => seedChart(createStore(book()));
const statuses = r => (r.operations || []).map(o => o.status);

// ---- WRITE open mode and the broker route (the REAL product driver, a recording fake broker) ----------------------

/** A recording broker: answer(body) -> { status, body } | 'THROW'. */
function broker(answer) {
  const calls = [];
  const fetchImpl = async (url, init) => {
    const body = JSON.parse(init.body);
    calls.push(body);
    const a = answer(body);
    if (a === 'THROW') throw new Error('connect ECONNREFUSED');
    return { status: a.status, json: async () => a.body };
  };
  return { calls, source: createBrokerConfigSource({ brokerUrl: BROKER, readToken: () => TOKEN, fetchImpl }) };
}

/** The product registry with the REAL drivers: only the broker, the editor and the runtime check are faked. */
function realProduct(b) {
  const ed = xlsxEditor();
  const runtime = { check: () => ({ ok: true }), load: ed.loadPlaywright };
  const built = buildServerTools({ config: { schemaVersion: 2, broker: { url: BROKER } }, configIdentity: 'fixture',
    readTool: { configSource: b.source, runtime }, writeTool: { configSource: b.source, runtime } });
  return { ...built, ed };
}

test('WRITE_MODE: the WRITE tool asks the broker for requestedMode WRITE; a denial is controlled, one call, no fallback', async () => {
  const cases = [
    ['OPEN_ACCESS_DENIED', { status: 403, body: { ok: false, code: 'PRODUCT_OCS_CONFIG_DENY_LEVEL' } }, 'ACCESS_DENIED'],
    ['OPEN_TARGET_READ_ONLY', { status: 403, body: { ok: false, code: 'PRODUCT_OCS_CONFIG_DENY_CONFIG_BELOW_REQUEST' } }, 'TARGET_READ_ONLY'],
    ['BROKER_UNAVAILABLE', 'THROW', 'BROKER_UNAVAILABLE'],
  ];
  EVIDENCE.brokerRoute = {};
  for (const [label, answer, code] of cases) {
    const b = broker(() => answer);
    const p = realProduct(b);
    const r = await p.registry.xlsx_execute_task.handler(req([STEP.write_range]));
    assert.deepEqual([r.ok, r.stage, r.code, r.dispatched], [false, 'OPEN', code, false], `${label}: ${JSON.stringify(r)}`);
    assert.equal(b.calls.length, 1, `${label}: exactly one broker call -- no READ re-open, no other credential, no other fileId`);
    assert.equal(b.calls[0].requestedMode, 'WRITE', label);
    assert.equal(b.calls[0].fileId, '1303626', label);
    assert.equal(p.ed.launches(), 0, `${label}: no browser on a refused config`);
    EVIDENCE.brokerRoute[label] = { RESULT: r.code, STAGE: r.stage, BROKER_CALLS: b.calls.length, REQUESTED_MODE: b.calls[0].requestedMode, ALTERNATE_CREDENTIAL_FALLBACK: 0, BROWSER_LAUNCH: p.ed.launches(), DISPATCH_COUNT: 0 };
  }
});

test('WRITE_MODE wrapper: the canonical driver is opened with operation WRITE; every other method is the driver\'s own', async () => {
  const h = harness({ store: await fresh() });
  const r = await h.tool.handler(req([STEP.write_range]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(h.opens, [{ fileId: '1303626', operation: 'WRITE' }]);
  const seen = [];
  const inner = { open: async a => { seen.push(a); return { ok: true }; }, close: function close() { return this === inner; } };
  const wrapped = W.productWriteDriver(inner);
  await wrapped.open({ url: 'u', user: 'x', pass: 'y', fileId: '7' });
  assert.deepEqual(seen, [{ fileId: '7', operation: 'WRITE' }], 'url / user / pass of the shared lifecycle never reach the product driver');
  assert.equal(wrapped.close(), true, 'other methods keep their own receiver');
});

test('READ_MODE_REGRESSION: xlsx_read / docx_read still ask READ and pass through the same shared registry', async () => {
  for (const fmt of ['xlsx', 'docx']) {
    const b = broker(body => ({ status: 200, body: { ok: true, status: 200, config: editConfig(body.fileId, fmt) } }));
    const ed = fmt === 'xlsx' ? xlsxEditor() : require('../fakes/read-editors.cjs').docxEditor();
    const runtime = { check: () => ({ ok: true }), load: ed.loadPlaywright };
    const { registry } = buildServerTools({ config: { schemaVersion: 2, broker: { url: BROKER } }, configIdentity: 'fixture', readTool: { configSource: b.source, runtime } });
    const op = fmt === 'xlsx' ? 'workbook_overview' : 'read_text';
    const r = await registry[`${fmt}_read`].handler({ fileId: '1302148', operation: op });
    assert.equal(r.ok, true, `${fmt}: ${JSON.stringify(r)}`);
    assert.deepEqual(b.calls.map(c => c.requestedMode), ['READ'], fmt);
    assert.equal(ed.api.dispatch + ed.api.forceSave, 0, `${fmt}: READ stays non-mutating`);
  }
});

// ---- the six intents: APPLIED, then SATISFIED in a new session with no write ----------------------------------------

test('SIX_INTENTS: first invocation APPLIED (L1/L2/L3 PASS, FORCESAVE_CONFIRMED); second, new session: SATISFIED, write 0', async () => {
  EVIDENCE.intents = {};
  for (const [intent, step] of Object.entries(STEP)) {
    const store = await fresh();
    const h1 = harness({ store });
    const r1 = await h1.tool.handler(req([step]));
    const c1 = h1.calls();
    assert.equal(r1.ok, true, `${intent}: ${JSON.stringify(r1)}`);
    assert.deepEqual(statuses(r1), ['APPLIED'], intent);
    assert.equal(r1.operations[0].l1, 'PASS', intent);
    assert.equal(r1.wholeTask, 'PASS', intent);
    assert.notEqual(r1.operations[0].l3, 'FAIL', intent);
    assert.equal(r1.persistence.state, 'FORCESAVE_CONFIRMED', intent);
    assert.equal(c1.dispatch, 1, intent);
    assert.equal(c1.forceSave, 1, intent);
    const h2 = harness({ store });
    const r2 = await h2.tool.handler(req([step]));
    const c2 = h2.calls();
    assert.equal(r2.ok, true, `${intent} retry: ${JSON.stringify(r2)}`);
    assert.deepEqual(statuses(r2), ['SATISFIED'], intent);
    assert.equal(c2.dispatch, 0, `${intent}: the retry dispatches nothing`);
    assert.equal(c2.forceSave, 0, `${intent}: a no-op starts no save`);
    assert.equal(r2.persistence.state, 'NOT_REQUIRED', intent);
    assert.equal(c2.open, 1, `${intent}: the retry is a NEW session`);
    EVIDENCE.intents[intent] = { FIRST_RESULT: r1.operations[0].status, FIRST_DISPATCH_COUNT: c1.dispatch, FIRST_VERIFY: { L1: r1.operations[0].l1, L2: r1.wholeTask, L3: r1.operations[0].l3 },
      FIRST_PERSISTENCE: r1.persistence.state, SECOND_RESULT: r2.operations[0].status, SECOND_DISPATCH_COUNT: c2.dispatch, SECOND_WRITE_COUNT: c2.dispatch, SECOND_FORCESAVE: c2.forceSave };
  }
});

// ---- multi-step ---------------------------------------------------------------------------------------------------

test('MULTISTEP: two intents, three steps, one intent twice -- every step APPLIED, whole task L2 PASS, retry SATISFIED', async () => {
  const combos = {
    two_intents: [{ ...STEP.write_range, id: 'w' }, { ...STEP.unmerge_range, id: 'u' }],
    three_steps: [{ ...STEP.write_range, id: 'w' }, { ...STEP.unmerge_range, id: 'u' }, { ...STEP.set_chart_title, id: 'c' }],
    same_intent_twice: [{ ...STEP.write_range, id: 'w1' }, { intent: 'write_range', id: 'w2', sheet: 'Budget', range: 'H5', values: [['z']] }],
  };
  EVIDENCE.multistep = {};
  for (const [name, steps] of Object.entries(combos)) {
    const store = await fresh();
    const h1 = harness({ store });
    const r1 = await h1.tool.handler(req(steps));
    assert.equal(r1.ok, true, `${name}: ${JSON.stringify(r1)}`);
    assert.deepEqual(statuses(r1), steps.map(() => 'APPLIED'), name);
    assert.equal(r1.wholeTask, 'PASS', name);
    assert.equal(h1.calls().forceSave, 1, `${name}: one barrier for the task`);
    const h2 = harness({ store });
    const r2 = await h2.tool.handler(req(steps));
    assert.deepEqual(statuses(r2), steps.map(() => 'SATISFIED'), name);
    assert.equal(h2.calls().dispatch, 0, name);
    EVIDENCE.multistep[name] = { FIRST: statuses(r1), L2: r1.wholeTask, DISPATCH: h1.calls().dispatch, SECOND: statuses(r2), SECOND_DISPATCH: h2.calls().dispatch };
  }
  // dependent steps on the SAME cells: the canonical planner refuses them (no step declares a disjoint axis) ->
  // whole task refused pre-dispatch, nothing opened
  const h = harness({ store: await fresh() });
  const r = await h.tool.handler(req([{ ...STEP.unmerge_range, id: 'u' }, { intent: 'write_range', id: 'w', sheet: 'Budget', range: 'D1', values: [['after']] }]));
  assert.deepEqual([r.ok, r.stage, r.code], [false, 'PRE_DISPATCH', 'REFUSED_COMPOSITION'], JSON.stringify(r));
  assert.equal(h.n.drivers, 0);
  EVIDENCE.multistep.dependent_same_cells = { RESULT: r.code, DRIVERS: h.n.drivers };
});

// ---- persistence --------------------------------------------------------------------------------------------------

test('PERSISTENCE: EDITOR_CLEAN only from the fresh canSave poll; a forceSave timeout is a FAILED barrier, never EDITOR_CLEAN', async () => {
  const unconfirmed = harness({ store: await fresh(), script: { forceSaveUnconfirmed: 'no-save-ack-observed' } });
  const u = await unconfirmed.tool.handler(req([STEP.write_range]));
  assert.equal(u.ok, false, 'an unconfirmed save never completes the task');
  assert.equal(u.outcome, 'task-save-not-confirmed');
  assert.equal(u.persistence.state, 'EDITOR_CLEAN');
  assert.equal(u.persistence.editorCleanSource, 'fresh-canSave-poll-after-forceSave');
  assert.equal(u.persistence.storedContentVerified, false);
  // NEG control: the editor never reports clean -> the barrier times out -> FAILED, not an assumed EDITOR_CLEAN
  const stuck = harness({ store: await fresh(), wrap: d => { d.canSave = async () => ({ ok: true, canSave: true }); return d; },
    persistence: { timeoutMs: 3, pollMs: 1, wait: async () => {} } });
  const t = await stuck.tool.handler(req([STEP.write_range]));
  assert.equal(t.ok, false);
  assert.equal(t.outcome, 'task-persistence-barrier-failed');
  assert.equal(t.persistence.state, 'FAILED');
  assert.equal(t.persistence.reason, 'forceSave-timeout');
  assert.equal(t.persistence.editorCleanSource, null);
  // refused and no-op tasks start no save
  const refused = harness({ store: await fresh() });
  await refused.tool.handler(req([{ intent: 'merge_range', sheet: 'Budget', range: 'A1:B1' }]));
  assert.equal(refused.calls().forceSave, 0);
  EVIDENCE.persistence = { FORCESAVE_UNCONFIRMED: { RESULT: u.outcome, SAVE: u.persistence.state, EDITOR_CLEAN_SOURCE: u.persistence.editorCleanSource },
    FORCESAVE_TIMEOUT_NEG: { RESULT: t.outcome, SAVE: t.persistence.state, REASON: t.persistence.reason }, REFUSED_FORCESAVE: refused.calls().forceSave };
});

// ---- failure injection ----------------------------------------------------------------------------------------------

const lost = () => { throw new Error('Target page, context or browser has been closed'); };
const W2 = [{ ...STEP.write_range, id: 'w' }, { ...STEP.set_chart_title, id: 'c' }];
const FAILURES = {
  READY_TIMEOUT: { wrap: d => { d.open = async () => ({ ok: false, detail: 'live spreadsheet model not ready' }); return d; },
    want: { stage: 'OPEN', code: 'OPEN_FAILED', dispatch: 0, write: false, verify: 'NONE', persist: 0, save: null, close: 'NONE (no handle)', retry: 'rerun the full task' } },
  DISPATCH_THROW: { script: { fault: 'throw' },
    want: { ops: ['DISPATCHED_OUTCOME_UNKNOWN'], uncertainty: 'DISPATCH_THREW', dispatch: 1, write: 'UNKNOWN (editor)', verify: 'NOT_PROVEN', persist: 0, save: 'UNCERTAIN_NOT_PERSISTED', close: 'SUCCESS', retry: 'rerun classifies from a fresh observation' } },
  L1_MISMATCH: { script: { fault: 'ignore' }, steps: W2,
    want: { ops: ['DISPATCHED_OUTCOME_UNKNOWN', 'SKIPPED_DUE_TO_PRIOR_FAILURE'], uncertainty: 'VERIFY_CANNOT_PROVE_TERMINAL_STATE', dispatch: 1, write: 'UNKNOWN (editor)', verify: 'L1 not PASS', persist: 0, save: 'UNCERTAIN_NOT_PERSISTED', close: 'SUCCESS', retry: 'rerun: s0 re-classified, s1 runs' } },
  L3_FAILURE: { script: { fault: 'side-effect' },
    want: { ops: ['DISPATCHED_OUTCOME_UNKNOWN'], uncertainty: 'INTEGRAL_PRESERVATION_NOT_PROVEN', l3: 'FAIL', dispatch: 1, write: 'YES (editor), side effect', verify: 'L1 PASS, L3 FAIL', persist: 0, save: 'UNCERTAIN_NOT_PERSISTED', close: 'SUCCESS', retry: 'rerun after inspection' } },
  FORCESAVE_UNCONFIRMED: { script: { forceSaveUnconfirmed: 'no-save-ack-observed' },
    want: { ops: ['APPLIED'], outcome: 'task-save-not-confirmed', dispatch: 1, write: 'YES (editor)', verify: 'L1/L2 PASS', persist: 1, save: 'EDITOR_CLEAN', close: 'SUCCESS', retry: 'rerun -> SATISFIED if stored, else APPLY' } },
  STORAGE_WRITE_REJECTED: { script: { forceSaveUnconfirmed: 'completion-not-successful' },
    want: { ops: ['APPLIED'], code: 'STORAGE_WRITE_REJECTED', outcome: 'task-save-not-confirmed', dispatch: 1, write: 'YES (editor), save rejected', verify: 'L1/L2 PASS', persist: 1, save: 'EDITOR_CLEAN', close: 'SUCCESS', retry: 'no automatic retry; no other location / credential' } },
  CLOSE_FAILURE: { script: { closeFails: true },
    want: { ops: ['APPLIED'], outcome: 'task-cleanup-failed', dispatch: 1, write: 'YES (editor)', verify: 'L1/L2 PASS', persist: 1, save: 'EDITOR_CLEAN', close: 'FAILED', retry: 'rerun -> SATISFIED if stored' } },
  BROWSER_LOSS: { wrap: d => { for (const m of ['readLive', 'canSave', 'forceSave', 'close']) { const real = d[m].bind(d); d[m] = async (...a) => (d.calls.dispatch >= 1 ? lost() : real(...a)); } return d; },
    want: { ops: ['DISPATCHED_OUTCOME_UNKNOWN'], dispatch: 1, write: 'UNKNOWN (editor)', verify: 'NOT_PROVEN', persist: 0, save: 'UNCERTAIN_NOT_PERSISTED', close: 'FAILED', retry: 'rerun classifies from a fresh observation' } },
};

test('FAILURE_MATRIX: canonical outcomes per injection point -- never APPLIED-and-ok, no rollback claimed', async () => {
  EVIDENCE.failures = {};
  for (const [name, c] of Object.entries(FAILURES)) {
    const store = await fresh();
    const h = harness({ store, script: c.script || {}, wrap: c.wrap || null });
    const r = await h.tool.handler(req(c.steps || [STEP.write_range]));
    const calls = h.calls();
    const w = c.want;
    assert.equal(r.ok, false, `${name}: ${JSON.stringify(r)}`);
    if (w.stage) assert.equal(r.stage, w.stage, name);
    if (w.code) assert.equal(r.code, w.code, name);
    if (w.outcome) assert.equal(r.outcome, w.outcome, name);
    if (w.ops) assert.deepEqual(statuses(r), w.ops, name);
    if (w.uncertainty) assert.equal(r.operations[0].uncertainty, w.uncertainty, name);
    if (w.l3) assert.equal(r.operations[0].l3, w.l3, name);
    assert.equal(calls.dispatch, w.dispatch, `${name}: dispatch count`);
    assert.equal(calls.forceSave, w.persist, `${name}: persist attempts`);
    if (r.persistence) assert.equal(r.persistence.state, w.save, name);
    if (w.close === 'FAILED' || w.close === 'SUCCESS') assert.equal(r.cleanupOutcome, w.close, name);
    if (w.ops && w.ops[0] === 'DISPATCHED_OUTCOME_UNKNOWN') assert.equal(r.mutationCertainty, 'UNCERTAIN', name);
    EVIDENCE.failures[name] = { RESULT: r.code || r.outcome, OPERATIONS: statuses(r), DISPATCH_COUNT: calls.dispatch, WRITE_OCCURRED: w.write, VERIFY_STATE: w.verify,
      PERSIST_ATTEMPT: calls.forceSave, SAVE_RESULT: r.persistence ? r.persistence.state : null, CLOSE_RESULT: r.cleanupOutcome || w.close, RETRY_EXPECTATION: w.retry, ROLLBACK_AVAILABLE: 'NO' };
  }
});

test('SHUTDOWN_DURING_WRITE: bounded close through the driver; the dispatched write is DISPATCHED_OUTCOME_UNKNOWN, not success', async () => {
  let release;
  const held = new Promise(res => { release = res; });
  let entered;
  const inDispatch = new Promise(res => { entered = res; });
  const h = harness({ store: await fresh(), wrap: d => { const real = d.dispatch.bind(d); d.dispatch = async (...a) => { entered(); await held; return real(...a); }; return d; } });
  const pending = h.tool.handler(req([STEP.write_range]));
  await inDispatch;
  const sd = await h.gate.shutdown({ timeoutMs: 50 });
  release();
  const r = await pending;
  assert.equal(sd.sessions, 1);
  assert.ok(['CLOSED', 'TIMEOUT'].includes(sd.cleanup), JSON.stringify(sd));
  assert.equal(r.ok, false, JSON.stringify(r));
  assert.deepEqual(statuses(r), ['DISPATCHED_OUTCOME_UNKNOWN']);
  assert.equal(r.mutationCertainty, 'UNCERTAIN');
  assert.equal(h.calls().forceSave, 0);
  const after = await h.tool.handler(req([STEP.write_range]));
  assert.equal(after.code, 'BUSY', 'a closing gate admits nothing new');
  EVIDENCE.failures.SHUTDOWN_DURING_WRITE = { SHUTDOWN: sd, RESULT: r.outcome, OPERATIONS: statuses(r), DISPATCH_COUNT: h.calls().dispatch, PERSIST_ATTEMPT: 0, SAVE_RESULT: r.persistence.state, ROLLBACK_AVAILABLE: 'NO' };
});

// ---- BUSY: one session per process, shared by READ and WRITE --------------------------------------------------------

test('BUSY: an active READ or WRITE session -> a second READ/WRITE is BUSY, no queue, no second browser; free again after close', async () => {
  const gate = createSessionGate();
  let release;
  const held = new Promise(res => { release = res; });
  const store = await fresh();
  const h = harness({ store, gate, wrap: d => { const real = d.open.bind(d); d.open = async a => { await held; return real(a); }; return d; } });
  const first = h.tool.handler(req([STEP.write_range]));
  await new Promise(res => setImmediate(res));
  assert.equal(gate.active(), 1);
  const second = await h.tool.handler(req([STEP.unmerge_range]));
  assert.equal(second.code, 'BUSY');
  assert.equal(h.n.drivers, 1, 'no second driver / browser');
  const driversWhileBusy = h.n.drivers;
  // a READ tool on the SAME gate (the product registry shares one gate)
  const { registry } = buildServerTools({ config: { schemaVersion: 2, broker: { url: BROKER } }, configIdentity: 'fixture', readTool: { gate, runtime: { check: () => ({ ok: true }), load: () => ({ ok: false }) } } });
  const read = await registry.xlsx_read.handler({ fileId: '1302148', operation: 'workbook_overview' });
  assert.equal(read.code, 'BUSY');
  release();
  const r1 = await first;
  assert.equal(r1.ok, true, JSON.stringify(r1));
  assert.equal(gate.active(), 0);
  const next = await h.tool.handler(req([STEP.unmerge_range]));
  assert.equal(next.ok, true, 'a new request is admitted after the close');
  EVIDENCE.busy = { MAX_ACTIVE_SESSIONS: gate.max, SECOND_WRITE: second.code, SECOND_READ: read.code, SECOND_BROWSER_LAUNCH: driversWhileBusy - 1, QUEUE: 'NONE', AFTER_CLOSE: next.ok };
});

test('BUSY shared in the product registry: the WRITE tool and the READ tools hold the same gate', () => {
  const { registry, gate } = buildServerTools({ config: { schemaVersion: 2, broker: { url: BROKER } }, configIdentity: 'fixture' });
  assert.equal(typeof registry.xlsx_execute_task.handler, 'function');
  const lease = gate.tryEnter();
  assert.ok(lease);
  assert.equal(gate.tryEnter(), null, 'one gate, max 1');
  lease.release();
});

// ---- refusal regression on the WIRED path ---------------------------------------------------------------------------

test('REFUSAL_REGRESSION on the wired path: every M11.4a.1 refusal stops before the runtime check, the driver and the executor', async () => {
  for (const [label, args, code] of REFUSALS) {
    const h = harness();
    const r = await h.tool.handler(JSON.parse(JSON.stringify(args)));
    assert.equal(r.code, code, label);
    assert.equal(r.stage, 'PRE_DISPATCH', label);
    assert.deepEqual([h.n.runtime, h.n.drivers, h.calls().open, h.calls().dispatch, h.calls().forceSave], [0, 0, 0, 0, 0], label);
  }
  EVIDENCE.refusals = { COUNT: REFUSALS.length, ALL_PRE_DISPATCH: true };
});

test('SERVED_SURFACE (M11.4b.1): the wired xlsx_execute_task is served, no DOCX WRITE tool exists', () => {
  assert.deepEqual(exposedTools(), ['product_status', 'xlsx_read', 'docx_read', 'xlsx_execute_task']);
  const { tools, registry } = buildServerTools({ config: { schemaVersion: 2, broker: { url: BROKER } }, configIdentity: 'fixture' });
  assert.deepEqual(Object.keys(tools), ['product_status', 'xlsx_read', 'docx_read', 'xlsx_execute_task']);
  assert.equal(registry.docx_execute_task, undefined);
});
