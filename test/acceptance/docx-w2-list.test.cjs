'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS, UNCERTAINTY_REASON } = require('../../src/core/execution/result.cjs');
const { PERSISTENCE_STATE } = require('../../src/core/persistence/verified-writes.cjs');
const { docxAdapter, planOperations } = require('../../src/adapters/docx/adapter.cjs');
const list = require('../../src/adapters/docx/families/list.cjs');
const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');

/**
 * M4 Wave 2 offline acceptance: the list family through the TARGET path (M1 -> M2 -> M3 -> DOCX
 * adapter) on the fake transport whose list model follows the live probes (numbering level +
 * identity per paragraph, one <ul>/<ol> per numbering run). FIRST applies, SECOND is a 0-write no-op.
 */

const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const H = 'Intro heading';
const A = 'Alpha item.';
const B = 'Beta item.';
const C = 'Gamma item.';
const E1 = 'Existing one.';
const E2 = 'Existing two.';
const Z = 'Closing paragraph.';
const doc = () => [
  { text: H, level: 1, props: { ...PLAIN, bold: true } },
  { text: A, props: { ...PLAIN } },
  { text: B, props: { ...PLAIN } },
  { text: C, props: { ...PLAIN } },
  { text: E1, props: { ...PLAIN }, list: { type: 'bullet', level: 0, numId: 'e1' } },
  { text: E2, props: { ...PLAIN }, list: { type: 'bullet', level: 0, numId: 'e1' } },
  { text: Z, props: { ...PLAIN } },
];

async function run(store, task, script = {}) {
  let t = 0;
  const driver = createFakeDocxDriver(store, script);
  const options = { driver, fileId: 'd', now: () => (t += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, docxAdapter, createExecutor(task));
  assert.equal(driver.calls.open, 1, 'OPEN_COUNT=1');
  assert.equal(driver.calls.close, 1, 'CLOSE_COUNT=1');
  if (result.execution.operationAccounting) assert.equal(result.execution.operationAccounting.balanced, true);
  return { result, driver, m: result.metrics, x: result.execution };
}
const task = (...steps) => ({ id: 'w2', steps });
const status = (x, i = 0) => x.appliedOperations[i].status;
const reason = (x, i = 0) => {
  const e = x.appliedOperations[i].evidence;
  return (e && e.evidence && e.evidence.reason) || (e && e.reason);
};
const stateOf = (store, text) => { const p = store.durable.find(x => x.text === text); return p && p.list ? { type: p.list.type, level: p.list.level, numId: p.list.numId } : { type: 'normal' }; };

async function firstAndSecond(store, t) {
  const first = await run(store, t);
  assert.equal(first.result.editOk, true, JSON.stringify(first.x.stopDetail || first.x.outcome));
  assert.ok(first.x.appliedOperations.every(e => e.status === OPERATION_STATUS.APPLIED), JSON.stringify(first.x.appliedOperations.map(e => e.status)));
  assert.equal(first.m.writes, t.steps.length);
  assert.equal(first.m.commands, t.steps.length);
  assert.ok(first.x.primitiveVerification.every(v => v.verdict === 'PASS'), 'L1');
  assert.equal(first.x.preservationVerification.length, t.steps.length, 'integral L3 for every P1 list operation');
  assert.ok(first.x.preservationVerification.every(v => v.verdict === 'PASS'), 'L3');
  assert.equal(first.x.wholeTaskVerification.verdict, 'PASS', 'L2');
  assert.equal(first.result.persistenceState, PERSISTENCE_STATE.EDITOR_CLEAN);
  assert.equal(first.driver.calls.forceSave, 1, 'one barrier, at task end');
  const second = await run(store, t);
  assert.equal(second.result.ok, true, JSON.stringify(second.x.stopDetail || second.x.outcome));
  assert.ok(second.x.appliedOperations.every(e => e.status === OPERATION_STATUS.SATISFIED), JSON.stringify(second.x.appliedOperations.map(e => e.status)));
  assert.equal(second.m.writes, 0);
  assert.equal(second.m.commands, 0);
  assert.equal(second.driver.calls.forceSave, 0);
  assert.equal(second.result.persistenceState, PERSISTENCE_STATE.NOT_REQUIRED);
  assert.equal(second.x.wholeTaskVerification.verdict, 'PASS');
  return { first, second };
}

// ---- projection (planning) ----------------------------------------------------------------------

test('projection: terminal state per target, relative steps need an absolute anchor, apply_list continuation', () => {
  const p = planOperations(task(
    { intent: 'apply_list', selector: { text: A }, type: 'numbered', level: 0 },
    { intent: 'indent_list_item', selector: { text: A } },
    { intent: 'apply_list', selector: { text: B }, type: 'numbered', level: 0 },
    { intent: 'outdent_list_item', selector: { text: C } },
  ));
  assert.equal(p.ok, true, p.detail);
  const [a, i, b, o] = p.operations;
  assert.deepEqual(a.terminal, { type: 'numbered', level: 1 }, 'the apply step retries against the TASK terminal state');
  assert.deepEqual(i.terminal, { type: 'numbered', level: 1 });
  assert.equal(b.continueFromText, A, 'consecutive apply_list with the same type+level continues the numbering');
  assert.equal(o.relativeUnprovable, true, 'outdent without an anchor for its target');
  assert.equal(i.relativeUnprovable, undefined);
});

// ---- FIRST / SECOND per intent ------------------------------------------------------------------

test('apply_list: FIRST applies (integral), SECOND is a 0-write no-op', async () => {
  const store = createStore(doc());
  await firstAndSecond(store, task({ intent: 'apply_list', selector: { text: A }, type: 'numbered', level: 0 }));
  assert.deepEqual(stateOf(store, A).type, 'numbered');
});

test('apply_list x2 (same type+level): the second continues the first numbering (L2 continuity)', async () => {
  const store = createStore(doc());
  await firstAndSecond(store, task(
    { intent: 'apply_list', selector: { text: A }, type: 'numbered', level: 0 },
    { intent: 'apply_list', selector: { text: B }, type: 'numbered', level: 0 },
  ));
  assert.equal(stateOf(store, A).numId, stateOf(store, B).numId);
});

test('set_list_level on an existing item: FIRST / SECOND, the other item untouched', async () => {
  const store = createStore(doc());
  await firstAndSecond(store, task({ intent: 'set_list_level', selector: { text: E1 }, level: 2 }));
  assert.equal(stateOf(store, E1).level, 2);
  assert.equal(stateOf(store, E2).level, 0);
});

test('indent_list_item anchored by apply_list in the same task: FIRST / SECOND', async () => {
  const store = createStore(doc());
  await firstAndSecond(store, task(
    { intent: 'apply_list', selector: { text: A }, type: 'bullet', level: 0 },
    { intent: 'indent_list_item', selector: { text: A } },
  ));
  assert.deepEqual(stateOf(store, A).level, 1);
});

test('outdent_list_item: level 1 -> 0, and at level 0 it leaves the list (removal path)', async () => {
  const s1 = createStore(doc());
  await firstAndSecond(s1, task({ intent: 'apply_list', selector: { text: A }, type: 'numbered', level: 1 }, { intent: 'outdent_list_item', selector: { text: A } }));
  assert.equal(stateOf(s1, A).level, 0);
  const s2 = createStore(doc());
  await firstAndSecond(s2, task({ intent: 'set_list_level', selector: { text: E1 }, level: 0 }, { intent: 'outdent_list_item', selector: { text: E1 } }));
  assert.equal(stateOf(s2, E1).type, 'normal');
});

test('a task whose terminal state already holds is SATISFIED without a write (apply + outdent-at-0 on a normal paragraph)', async () => {
  const { x, m } = await run(createStore(doc()), task({ intent: 'apply_list', selector: { text: A }, type: 'bullet', level: 0 }, { intent: 'outdent_list_item', selector: { text: A } }));
  assert.deepEqual(x.appliedOperations.map(e => e.status), [OPERATION_STATUS.SATISFIED, OPERATION_STATUS.SATISFIED]);
  assert.equal(m.writes, 0);
});

test('set_list_type on a single-item list: FIRST / SECOND', async () => {
  const store = createStore(doc());
  await firstAndSecond(store, task({ intent: 'apply_list', selector: { text: C }, type: 'bullet', level: 0 }, { intent: 'set_list_type', selector: { text: C }, type: 'numbered' }));
  assert.equal(stateOf(store, C).type, 'numbered');
});

test('remove_list on an existing item: FIRST / SECOND (the removed item is a normal paragraph now)', async () => {
  const store = createStore(doc());
  await firstAndSecond(store, task({ intent: 'remove_list', selector: { text: E1 } }));
  assert.equal(stateOf(store, E1).type, 'normal');
  assert.equal(stateOf(store, E2).type, 'bullet', 'the other item keeps its list');
});

test('REF: apply_list -> set_list_type on its result (selector.ref, dependsOn), retry SATISFIED', async () => {
  const store = createStore(doc());
  await firstAndSecond(store, task(
    { intent: 'apply_list', selector: { text: B }, type: 'bullet', level: 0 },
    { intent: 'set_list_type', selector: { ref: 's0' }, type: 'numbered', dependsOn: ['s0'] },
  ));
  assert.equal(stateOf(store, B).type, 'numbered');
});

// W2.1 closed the cross-family boundary: test/acceptance/docx-w21-composability.test.cjs.

// ---- negatives: FAIL_CLOSED before any dispatch -------------------------------------------------

test('negatives: unprovable relative level / precondition / ambiguous / absent / unreadable -> FAIL_CLOSED, 0 dispatch', async () => {
  const cases = [
    [task({ intent: 'indent_list_item', selector: { text: E1 } }), doc(), {}, 'list-relative-level-unprovable'],
    [task({ intent: 'set_list_level', selector: { text: A }, level: 1 }), doc(), {}, 'list-precondition-not-met'],
    [task({ intent: 'apply_list', selector: { text: A }, type: 'bullet' }), [...doc(), { text: A, props: { ...PLAIN } }], {}, 'list-target-ambiguous'],
    [task({ intent: 'apply_list', selector: { text: 'Nope.' }, type: 'bullet' }), doc(), {}, 'list-target-not_found'],
    [task({ intent: 'apply_list', selector: { text: A }, type: 'bullet' }), doc(), { numberingUnreadable: true }, 'list-state-unreadable'],
  ];
  for (const [t, d, script, want] of cases) {
    const { x, driver } = await run(createStore(d), t, script);
    assert.equal(status(x), OPERATION_STATUS.FAILED, want);
    assert.equal(reason(x), want);
    assert.equal(driver.calls.dispatch, 0, want);
  }
});

test('negatives: context/physical selectors and bad parameters -> plan fails before any dispatch', async () => {
  for (const step of [
    { intent: 'apply_list', selector: { text: A, before: H }, type: 'bullet' },
    { intent: 'apply_list', selector: { text: A, index: 1 }, type: 'bullet' },
    { intent: 'apply_list', selector: { text: A }, type: 'roman' },
    { intent: 'set_list_level', selector: { text: E1 }, level: 9 },
  ]) {
    const { x, driver } = await run(createStore(doc()), task(step));
    assert.equal(x.outcome, 'plan-failed', JSON.stringify(step));
    assert.equal(driver.calls.dispatch, 0);
  }
});

// ---- the write happened but is not the terminal contract ----------------------------------------

test('faults at L1: no mutation / wrong level -> never APPLIED, no barrier', async () => {
  for (const fault of ['ignore', 'list-wrong-level']) {
    const { x, result, driver } = await run(createStore(doc()), task({ intent: 'apply_list', selector: { text: A }, type: 'bullet', level: 0 }), { fault });
    assert.equal(status(x), OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN, fault);
    assert.equal(x.primitiveVerification[0].verdict, 'FAIL', fault);
    assert.equal(x.appliedOperations[0].uncertainty.reason, UNCERTAINTY_REASON.VERIFY_CANNOT_PROVE_TERMINAL_STATE, fault);
    assert.equal(result.persistenceState, PERSISTENCE_STATE.UNCERTAIN_NOT_PERSISTED, fault);
    assert.equal(driver.calls.forceSave, 0, fault);
  }
});

test('faults at L3: the target holds but another item changed / a shared numbering template -> NOT APPLIED', async () => {
  const touched = await run(createStore(doc()), task({ intent: 'apply_list', selector: { text: A }, type: 'bullet', level: 0 }), { fault: 'list-touch-other' });
  assert.equal(touched.x.primitiveVerification[0].verdict, 'PASS');
  assert.equal(touched.x.preservationVerification[0].verdict, 'FAIL');
  assert.equal(touched.x.appliedOperations[0].uncertainty.reason, UNCERTAINTY_REASON.INTEGRAL_PRESERVATION_NOT_PROVEN);
  // set_list_type on an item of a two-item list changes the shared numbering level: collateral, not a pass
  const shared = await run(createStore(doc()), task({ intent: 'set_list_type', selector: { text: E1 }, type: 'numbered' }));
  assert.equal(shared.x.primitiveVerification[0].verdict, 'PASS');
  assert.equal(shared.x.preservationVerification[0].verdict, 'FAIL', 'the other item changed type too');
  const changed = await run(createStore(doc()), task({ intent: 'apply_list', selector: { text: A }, type: 'bullet' }), { fault: 'list-change-text' });
  assert.notEqual(status(changed.x), OPERATION_STATUS.APPLIED, 'text changed by a list operation -> not applied');
});

test('fault at L2: the continuation opened a NEW numbering -> whole-task FAIL', async () => {
  const { x, result } = await run(createStore(doc()), task(
    { intent: 'apply_list', selector: { text: A }, type: 'numbered', level: 0 },
    { intent: 'apply_list', selector: { text: B }, type: 'numbered', level: 0 },
  ), { fault: 'list-new-numbering', faultOn: [2] });
  assert.ok(x.primitiveVerification.every(v => v.verdict === 'PASS'), 'each item alone looks right');
  assert.equal(x.wholeTaskVerification.verdict, 'FAIL', 'but the two items are not one list');
  assert.equal(result.ok, false);
  assert.equal(x.outcome, 'whole-task-verification-failed');
});

test('lowering: one callCommand through the common dispatch, no editor backdoor, level re-checked live', () => {
  const body = list.commandBody({ listAction: 'remove_list', targetText: E1, beforeType: 'bullet', beforeLevel: 0, bridgeNumberingRemove: true });
  // M11.5d.4: the removal is the driver's B18 Bridge client (word /2), never a direct internal call
  assert.match(body, /__eoB18Remove\(m\.targetText, m\.beforeLevel\)/);
  assert.doesNotMatch(body, /RemoveNumPr|\.Paragraph\./);
  assert.match(body, /liveLevel !== m\.beforeLevel\) return/);
  assert.doesNotMatch(body, /frame\.evaluate|window\.|put_ListType/);
});

// ---- M11.5d.4: the B18 Bridge WRITE runtime guard (word /2 only; was the EO_INTERNAL_DEPENDENCY guard) --------------

const WORD_V1 = '9.3.4|euro-office-bridge-word-v1/9.3.4.60/1';

test('B18 guard A: the runtime is word /2 -> the removal runs (remove_list FIRST / SECOND)', async () => {
  const store = createStore(doc());
  await firstAndSecond(store, task({ intent: 'remove_list', selector: { text: E1 } }));
  assert.equal(stateOf(store, E1).type, 'normal');
});

test('B18 guard B: the runtime is NOT word /2 -> FAIL_CLOSED before any dispatch (remove, outdent at 0); never a direct internal fallback', async () => {
  for (const t of [
    task({ intent: 'remove_list', selector: { text: E1 } }),
    task({ intent: 'set_list_level', selector: { text: E2 }, level: 0 }, { intent: 'outdent_list_item', selector: { text: E2 } }),
  ]) {
    const { x, driver, m } = await run(createStore(doc()), t, { runtime: WORD_V1 });
    const i = t.steps.length - 1;
    assert.equal(status(x, i), OPERATION_STATUS.FAILED);
    assert.equal(reason(x, i), 'docx-numbering-remove-bridge-write-runtime-not-proven');
    assert.equal(driver.calls.dispatch, t.steps.length - 1, 'no dispatch for the removal step');
    assert.equal(m.writes, t.steps.length - 1);
  }
  // a step that does not remove numbering is not blocked on word /1
  const ok = await run(createStore(doc()), task({ intent: 'set_list_level', selector: { text: E1 }, level: 1 }), { runtime: WORD_V1 });
  assert.equal(status(ok.x), OPERATION_STATUS.APPLIED);
});

test('B18 guard C: the removal command does nothing -> L1 FAIL, uncertain, no barrier', async () => {
  const { x, result, driver } = await run(createStore(doc()), task({ intent: 'remove_list', selector: { text: E1 } }), { fault: 'ignore' });
  assert.equal(status(x), OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN);
  assert.equal(x.primitiveVerification[0].verdict, 'FAIL');
  assert.equal(x.appliedOperations[0].uncertainty.reason, UNCERTAINTY_REASON.VERIFY_CANNOT_PROVE_TERMINAL_STATE);
  assert.equal(result.persistenceState, PERSISTENCE_STATE.UNCERTAIN_NOT_PERSISTED);
  assert.equal(driver.calls.forceSave, 0);
});
