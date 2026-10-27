'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS, UNCERTAINTY_REASON } = require('../../src/core/execution/result.cjs');
const { PERSISTENCE_STATE } = require('../../src/core/persistence/verified-writes.cjs');
const { validateAdapter } = require('../../src/core/adapter/contract.cjs');
const { docxAdapter } = require('../../src/adapters/docx/adapter.cjs');
const { normalizeInventory } = require('../../src/adapters/docx/inventory.cjs');
const { createStore, createFakeDocxDriver, htmlOf } = require('../fakes/fake-docx-driver.cjs');

/**
 * M4 rewrite_paragraph vertical slice through the TARGET path: M1 lifecycle -> M2 observation /
 * verification -> M3 executor -> DOCX adapter, on a fake transport that keeps DURABLE and LIVE state
 * apart (W3.0-measured behaviour). The LIVE acceptance on the real editor is a separate harness.
 */

const RICH = { bold: true, italic: true, underline: false, strike: false, fontFamily: 'Arial', fontSize: 11, language: 'hu-HU' };
const OLD = 'W010 HYPERLINK TARGET';
const NEW = 'M4 REWRITE TERMINAL';
const doc = () => [
  { text: 'Intro heading', level: 1, props: { ...RICH, bold: false } },
  { text: 'First body paragraph.', props: { ...RICH, bold: false, italic: false } },
  { text: OLD, props: RICH },
  { text: 'Closing paragraph.', props: { ...RICH, italic: false } },
];

/** The selector a caller builds from a FRESH inventory, exactly like the source W9.2-R harness. */
function selectorFrom(paragraphs) {
  const primary = normalizeInventory({ html: htmlOf(paragraphs), paragraphs });
  const b = primary.blocks.find(x => x.text === OLD);
  return { text: OLD, kind: b.kind, before: b.locator.context.before, after: b.locator.context.after };
}

async function run(store, task, script = {}) {
  let t = 0;
  const driver = createFakeDocxDriver(store, script);
  const options = {
    driver, fileId: 'disposable', now: () => (t += 1),
    ready: { wait: async () => {} }, persistence: { wait: async () => {} },
  };
  const result = await withSession(options, docxAdapter, createExecutor(task));
  assert.equal(driver.calls.open, 1, 'OPEN_COUNT=1');
  assert.equal(driver.calls.close, 1, 'CLOSE_COUNT=1');
  assert.equal(result.execution.operationAccounting.balanced, true);
  return { result, driver, m: result.metrics };
}

const rewriteTask = selector => ({ id: 'm4-rewrite', steps: [{ intent: 'rewrite_paragraph', selector, text: NEW }] });

test('the DOCX adapter satisfies the M0 contract and owns no lifecycle member', () => {
  assert.deepEqual(validateAdapter(docxAdapter), { ok: true, problems: [] });
});

test('FIRST session: APPLY -> 1 dispatch -> L1 -> L3 preservation -> L2 -> barrier -> close', async () => {
  const store = createStore(doc());
  const { result, driver, m } = await run(store, rewriteTask(selectorFrom(doc())));
  // M7 save contract: the DOCX driver has no save observer -> the edit is ok, the task waits on an unconfirmed save
  assert.equal(result.editOk, true, 'FIRST_EDIT_OK');
  assert.equal(result.ok, false, 'no confirmed save on DOCX (a separate open item)');
  assert.equal(result.outcome, 'task-save-not-confirmed');
  assert.equal(m.writes, 1, 'FIRST_WRITES=1');
  assert.equal(m.commands, 1, 'exactly the necessary dispatch');
  const e = result.execution.appliedOperations[0];
  assert.equal(e.status, OPERATION_STATUS.APPLIED);
  assert.equal(e.verification, 'LIVE_VERIFY', 'primitive layer authority');
  assert.equal(result.authority, 'LIVE_VERIFY', 'the barrier never raises the authority: no general PERSISTED');
  assert.equal(result.saveEvidence.reason, 'barrier:driver-provides-no-save-confirmation');
  assert.equal(result.execution.wholeTaskVerification.failedCount, 0, 'WHOLE_TASK_FAILED=0');
  assert.ok(result.execution.wholeTaskVerification.checkCount >= 3);
  assert.equal(result.execution.preservationVerification[0].verdict, 'PASS');
  assert.deepEqual(result.execution.preservationVerification[0].evidence.evidence.after, RICH, 'bold, italic, hu-HU and the other four survive');
  assert.equal(result.persistenceState, PERSISTENCE_STATE.EDITOR_CLEAN, 'PERSISTENCE=PASS');
  assert.equal(result.cleanupOutcome, 'SUCCESS', 'CLOSE=PASS');
  assert.equal(driver.calls.forceSave, 1);
  assert.equal(store.durable.find(p => p.text === NEW).props.language, 'hu-HU');
  assert.equal(store.durable.some(p => p.text === OLD), false);
});

test('SECOND session (restart retry): fresh durable read -> SATISFIED, 0 dispatch, 0 writes, no barrier', async () => {
  const store = createStore(doc());
  const selector = selectorFrom(doc());
  await run(store, rewriteTask(selector));
  // A NEW session, the SAME logical task, no receipt and no process memory passed in.
  const { result, driver, m } = await run(store, rewriteTask(selector));
  assert.equal(result.ok, true);
  assert.equal(result.execution.noOp, true, 'SECOND_NOOP=true');
  assert.equal(m.writes, 0, 'SECOND_WRITES=0');
  assert.equal(m.commands, 0, 'SECOND_COMMANDS=0');
  assert.equal(driver.calls.forceSave, 0);
  assert.equal(result.persistenceState, PERSISTENCE_STATE.NOT_REQUIRED);
  assert.equal(result.execution.wholeTaskVerification.verdict, 'PASS');
  const e = result.execution.appliedOperations[0];
  assert.equal(e.status, OPERATION_STATUS.SATISFIED);
  assert.equal(e.evidence.proof, 'fresh-live-terminal-context-proven', 'RETRY_AUTHORITY=FRESH_DURABLE_STATE');
  assert.ok(m.freshReads >= 1, 'the retry read before it classified');
});

test('NEGATIVE CONTROL for the retry authority: if the first write never reached DURABLE, the retry APPLIES again', async () => {
  const store = createStore(doc());
  const selector = selectorFrom(doc());
  // First session writes LIVE but its barrier is sabotaged at the transport: nothing is saved.
  let t = 0;
  const driver = createFakeDocxDriver(store);
  driver.forceSave = async () => ({ ok: true }); // pretends, saves nothing
  driver.canSave = async () => ({ ok: true, canSave: false });
  await withSession({ driver, fileId: 'd', now: () => (t += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } },
    docxAdapter, createExecutor(rewriteTask(selector)));
  const { result, m } = await run(store, rewriteTask(selector));
  assert.equal(result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED, 'the retry saw the DURABLE old text');
  assert.equal(m.writes, 1);
});

test('retry identity: without structural anchors a missing original is FAIL_CLOSED, never a blind SATISFIED', async () => {
  const store = createStore(doc());
  await run(store, rewriteTask(selectorFrom(doc())));
  const { result, driver } = await run(store, rewriteTask({ text: OLD }));
  assert.equal(result.execution.outcome, 'operation-fail-closed');
  assert.equal(result.execution.appliedOperations[0].evidence.evidence.reason, 'rewrite-terminal-identity-insufficient-selector');
  assert.equal(driver.calls.dispatch, 0);
});

test('lost run properties -> NOT APPLIED: preservation is PRIMITIVE_INTEGRAL, the write is uncertain, no barrier', async () => {
  const store = createStore(doc());
  const { result, driver } = await run(store, rewriteTask(selectorFrom(doc())), { fault: 'drop-props' });
  const e = result.execution.appliedOperations[0];
  assert.equal(result.ok, false);
  assert.equal(e.status, OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN, 'FALSE_APPLIED_ON_PROPERTY_LOSS=NO');
  assert.notEqual(e.verification, 'LIVE_VERIFY');
  assert.equal(e.uncertainty.reason, UNCERTAINTY_REASON.INTEGRAL_PRESERVATION_NOT_PROVEN);
  assert.equal(e.uncertainty.verificationOutcome, 'PASS', 'the text part (L1) did pass');
  assert.equal(result.execution.primitiveVerification[0].verdict, 'PASS');
  assert.equal(result.execution.preservationVerification[0].verdict, 'FAIL', 'INTEGRAL_PRESERVATION_BEFORE_APPLIED');
  assert.equal(result.metrics.writes, 0, 'no verified write is counted');
  assert.equal(result.persistenceState, PERSISTENCE_STATE.UNCERTAIN_MAY_PERSIST, 'uncertain-only: hybrid policy, no barrier');
  assert.equal(driver.calls.forceSave, 0);
  assert.equal(store.durable.some(p => p.text === OLD), true, 'durable untouched');
});

test('property loss after a verified prefix -> prefix APPLIED, loss uncertain, PERSISTED_UNCERTAIN', async () => {
  const SECOND = 'M4 SECOND TERMINAL';
  const paragraphs = doc();
  const primary = normalizeInventory({ html: htmlOf(paragraphs), paragraphs });
  const closing = primary.blocks.find(x => x.text === 'Closing paragraph.');
  const task = { id: 'm4-two', steps: [
    { intent: 'rewrite_paragraph', selector: selectorFrom(paragraphs), text: NEW },
    // no `before` anchor: the first rewrite changes this paragraph's preceding neighbour
    { intent: 'rewrite_paragraph', selector: { text: 'Closing paragraph.', kind: closing.kind }, text: SECOND },
  ] };
  const store = createStore(paragraphs);
  const { result, driver } = await run(store, task, { fault: 'drop-props', dropPropsOn: [2] });
  const [first, second] = result.execution.appliedOperations;
  assert.equal(first.status, OPERATION_STATUS.APPLIED);
  assert.equal(second.status, OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN);
  assert.equal(second.uncertainty.reason, UNCERTAINTY_REASON.INTEGRAL_PRESERVATION_NOT_PROVEN);
  assert.equal(result.persistenceState, PERSISTENCE_STATE.EDITOR_CLEAN);
  assert.equal(driver.calls.forceSave, 1, 'the verified prefix persists; persistence is not rollback');
  // NEGATIVE CONTROL: the same two-step task without the fault -> both APPLIED, PERSISTED.
  const neg = await run(createStore(paragraphs), task);
  assert.deepEqual(neg.result.execution.appliedOperations.map(e => e.status), [OPERATION_STATUS.APPLIED, OPERATION_STATUS.APPLIED]);
  assert.equal(neg.result.persistenceState, PERSISTENCE_STATE.EDITOR_CLEAN);
});

test('retry after an un-persisted property loss re-applies from the durable old state', async () => {
  const store = createStore(doc());
  const selector = selectorFrom(doc());
  await run(store, rewriteTask(selector), { fault: 'drop-props' });
  const { result, m } = await run(store, rewriteTask(selector));
  assert.equal(result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
  assert.equal(m.writes, 1);
  assert.equal(store.durable.find(p => p.text === NEW).props.language, 'hu-HU');
});

// ---- owner decision R2: P1 transformative retry is never a blind SATISFIED -----------------------
// A normal success, B normal persisted retry and C retry after an un-persisted loss are the FIRST /
// SECOND / "retry after an un-persisted property loss" cases above. D, E, F and the limits follow.

const LOST = { ...RICH, bold: false, italic: false, language: 'en-US' };
const durableWith = props => doc().map(p => (p.text === OLD ? { text: NEW, props } : p));
const recovering = (selector, extra = {}) => ({ ...rewriteTask(selector), recovery: { unresolvedUncertainOperations: ['s0'] }, ...extra });
const withExpectation = (task, expectedPreservation) => ({ ...task, steps: task.steps.map(s => ({ ...s, expectedPreservation })) });
const reason = r => r.execution.appliedOperations[0].evidence.evidence.reason;

test('DERIVED expectation: APPLY carries the seven properties of the fresh pre-mutation target; L3 checks against it', async () => {
  const { result } = await run(createStore(doc()), rewriteTask(selectorFrom(doc())));
  const l3 = result.execution.preservationVerification[0];
  assert.equal(l3.verdict, 'PASS');
  assert.deepEqual(l3.evidence.evidence.expected, RICH);
  assert.equal(l3.evidence.evidence.expectationSource, 'derived-from-fresh-pre-mutation-observation');
});

test('D. property-lost state durable, baseline gone, explicit unresolved-uncertainty recovery -> FAIL_CLOSED', async () => {
  const { result, driver } = await run(createStore(durableWith(LOST)), recovering(selectorFrom(doc())));
  assert.equal(result.execution.outcome, 'operation-fail-closed');
  assert.equal(result.execution.appliedOperations[0].status, OPERATION_STATUS.FAILED);
  assert.equal(reason(result), 'preservation-baseline-unavailable', 'BLIND_TERMINAL_SATISFIED=REJECTED');
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(driver.calls.forceSave, 0);
});

test('D (end to end): a real uncertain run is persisted by its verified prefix; the recovery names it -> FAIL_CLOSED for it only', async () => {
  const SECOND = 'M4 SECOND TERMINAL';
  const paragraphs = doc();
  const closing = normalizeInventory({ html: htmlOf(paragraphs), paragraphs }).blocks.find(x => x.text === 'Closing paragraph.');
  // s0 anchors on `before` only: its `after` neighbour is the paragraph s1 rewrites, so an `after`
  // anchor would make s0's own retry fail closed (rewrite-terminal-target-not_found) -- safe, but
  // not what this case measures.
  const s0 = selectorFrom(paragraphs);
  const task = { id: 'm4-two', steps: [
    { intent: 'rewrite_paragraph', selector: { text: s0.text, kind: s0.kind, before: s0.before }, text: NEW },
    { intent: 'rewrite_paragraph', selector: { text: 'Closing paragraph.', kind: closing.kind, after: '' }, text: SECOND },
  ] };
  const store = createStore(paragraphs);
  const first = await run(store, task, { fault: 'drop-props', dropPropsOn: [2] });
  assert.equal(first.result.persistenceState, PERSISTENCE_STATE.EDITOR_CLEAN, 'the loss reached durable state');
  const uncertain = first.result.execution.appliedOperations.filter(e => e.status === OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN).map(e => e.operationId);
  assert.deepEqual(uncertain, ['s1']);
  // The caller hands the earlier ExecutionResult's uncertain ids to the next invocation.
  const { result, driver } = await run(store, { ...task, recovery: { unresolvedUncertainOperations: uncertain } });
  const [a, b] = result.execution.appliedOperations;
  assert.equal(a.status, OPERATION_STATUS.SATISFIED, 's0 was verified: its normal retry stays SATISFIED');
  assert.equal(b.status, OPERATION_STATUS.FAILED);
  assert.equal(b.evidence.evidence.reason, 'preservation-baseline-unavailable');
  assert.equal(driver.calls.dispatch, 0);
});

test('E. explicit expected preservation: SATISFIED only if the terminal properties match it', async () => {
  const selector = selectorFrom(doc());
  const good = await run(createStore(durableWith(RICH)), withExpectation(recovering(selector), { bold: true, italic: true, language: 'hu-HU' }));
  assert.equal(good.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
  assert.equal(good.result.execution.appliedOperations[0].evidence.preservation, 'terminal-properties-match-explicit-expectation');
  assert.equal(good.m.writes, 0);
  const bad = await run(createStore(durableWith(LOST)), withExpectation(recovering(selector), { bold: true, italic: true, language: 'hu-HU' }));
  assert.equal(reason(bad.result), 'rewrite-terminal-preservation-mismatch');
});

test('F. terminal text present, properties wrong, expectation known -> never SATISFIED (with or without recovery)', async () => {
  const selector = selectorFrom(doc());
  const exp = { bold: true, italic: true, language: 'hu-HU' };
  for (const task of [withExpectation(rewriteTask(selector), exp), withExpectation(recovering(selector), exp)]) {
    const { result, driver } = await run(createStore(durableWith(LOST)), task);
    assert.equal(reason(result), 'rewrite-terminal-preservation-mismatch');
    assert.equal(driver.calls.dispatch, 0);
  }
  // The other terminal branch (the selector itself resolves to the terminal text) holds the same rule.
  const direct = await run(createStore(durableWith(LOST)), withExpectation({ id: 'd', steps: [{ intent: 'rewrite_paragraph', selector: { text: NEW }, text: NEW }] }, exp));
  assert.equal(reason(direct.result), 'rewrite-terminal-preservation-mismatch');
});

test('an explicit expectation that contradicts the fresh baseline refuses APPLY (no write)', async () => {
  const { result, driver } = await run(createStore(doc()), withExpectation(rewriteTask(selectorFrom(doc())), { language: 'en-US' }));
  assert.equal(reason(result), 'rewrite-explicit-expectation-contradicts-baseline');
  assert.equal(driver.calls.dispatch, 0);
});

test('recovery with the original still durable (the loss was not persisted) -> APPLY again, not FAIL_CLOSED', async () => {
  const { result, m } = await run(createStore(doc()), recovering(selectorFrom(doc())));
  assert.equal(result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
  assert.equal(m.writes, 1);
});

test('malformed recovery or expectation fails the plan closed, before any dispatch', async () => {
  const selector = selectorFrom(doc());
  for (const task of [
    { ...rewriteTask(selector), recovery: { unresolvedUncertainOperations: 's0' } },
    { ...rewriteTask(selector), recovery: { unresolvedUncertainOperations: ['s9'] } },
    withExpectation(rewriteTask(selector), { colour: 'red' }),
    withExpectation(rewriteTask(selector), {}),
  ]) {
    const { result, driver } = await run(createStore(doc()), task);
    assert.equal(result.execution.outcome, 'plan-failed', JSON.stringify(task.recovery || task.steps[0].expectedPreservation));
    assert.equal(driver.calls.dispatch, 0);
  }
});

test('LIMIT: without recovery context or expectation, a durable property-lost terminal state is SATISFIED', async () => {
  // Information-theoretic limit (owner decision R2, point 9): a fresh invocation that carries no
  // recovery context is indistinguishable from a new task, and the framework keeps no task state
  // across restarts (no journal, no document flag -- point 5). Pinned so any change is visible.
  const { result, m } = await run(createStore(durableWith(LOST)), rewriteTask(selectorFrom(doc())));
  assert.equal(result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
  assert.equal(m.writes, 0);
});

test('unreadable property baseline -> refused BEFORE dispatch (certain, no mutation)', async () => {
  const store = createStore(doc());
  const { result, driver } = await run(store, rewriteTask(selectorFrom(doc())), { propsUnreadable: true });
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(result.execution.appliedOperations[0].status, OPERATION_STATUS.FAILED);
  assert.equal(result.mutationCertainty, 'CERTAIN');
});

test('ambiguous original target -> FAIL_CLOSED, zero dispatch', async () => {
  const twice = [...doc(), { text: OLD, props: RICH }];
  const store = createStore(twice);
  const { result, driver } = await run(store, rewriteTask({ text: OLD }));
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(result.execution.appliedOperations[0].evidence.evidence.reason, 'rewrite-ambiguous-old-target');
});

test('old and new text both present -> FAIL_CLOSED (contradicting states), zero dispatch', async () => {
  const both = [...doc(), { text: NEW, props: RICH }];
  const store = createStore(both);
  const { result, driver } = await run(store, rewriteTask(selectorFrom(both)));
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(result.execution.appliedOperations[0].evidence.evidence.reason, 'rewrite-conflicting-old-and-new-state');
});

test('dispatch "succeeds" but the editor did not change -> never APPLIED; uncertain, not persisted', async () => {
  const store = createStore(doc());
  const { result, driver } = await run(store, rewriteTask(selectorFrom(doc())), { fault: 'ignore' });
  const e = result.execution.appliedOperations[0];
  assert.equal(e.status, OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN);
  assert.equal(e.uncertainty.reason, UNCERTAINTY_REASON.VERIFY_CANNOT_PROVE_TERMINAL_STATE);
  assert.equal(result.persistenceState, PERSISTENCE_STATE.UNCERTAIN_MAY_PERSIST);
  assert.equal(driver.calls.forceSave, 0);
});

test('stronger than the source L2: new text inserted but old text kept -> not APPLIED', async () => {
  const store = createStore(doc());
  const { result } = await run(store, rewriteTask(selectorFrom(doc())), { fault: 'insert-only' });
  assert.notEqual(result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
  assert.equal(result.ok, false);
});

test('a registered but not-yet-migrated intent is refused at planning, before any dispatch', async () => {
  const store = createStore(doc());
  const { result, driver } = await run(store, { id: 'x', steps: [{ intent: 'delete_table_row', selector: { text: OLD } }] });
  assert.equal(result.execution.outcome, 'plan-failed');
  assert.equal(driver.calls.dispatch, 0);
});

test('DOCX observation is LIVE-primary: no package I/O in the slice, every read is a counted live read', async () => {
  const store = createStore(doc());
  const { m, driver } = await run(store, rewriteTask(selectorFrom(doc())));
  assert.equal(m.packageReads, 0);
  assert.equal(driver.calls.readPackage, 0);
  assert.equal(m.freshReads, driver.calls.readLive);
});
