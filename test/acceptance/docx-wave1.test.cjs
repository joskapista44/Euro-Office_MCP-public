'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS, UNCERTAINTY_REASON } = require('../../src/core/execution/result.cjs');
const { PERSISTENCE_STATE } = require('../../src/core/persistence/verified-writes.cjs');
const { docxAdapter, planOperations } = require('../../src/adapters/docx/adapter.cjs');
const { normalizeInventory } = require('../../src/adapters/docx/inventory.cjs');
const paragraphFamily = require('../../src/adapters/docx/families/paragraph.cjs');
const { createStore, createFakeDocxDriver, htmlOf } = require('../fakes/fake-docx-driver.cjs');

/**
 * M4 Wave 1 offline acceptance: format_text + the five paragraph-structure intents through the
 * TARGET path (M1 lifecycle -> M2 observation/verification -> M3 executor -> DOCX adapter) on the
 * fake transport that keeps DURABLE and LIVE state apart. Every FIRST also runs a SECOND in a new
 * session (restart retry: SATISFIED, 0 write, 0 command, no barrier).
 */

const RICH = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const H = 'Intro heading';
const P1 = 'First body paragraph.';
const P2 = 'Second body paragraph.';
const P3 = 'Third body paragraph.';
const P4 = 'Closing paragraph.';
const doc = () => [
  { text: H, level: 1, props: { ...RICH, bold: true } },
  { text: P1, props: { ...RICH } },
  { text: P2, props: { ...RICH } },
  { text: P3, props: { ...RICH } },
  { text: P4, props: { ...RICH } },
];
const order = store => store.durable.map(p => p.text);

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
const one = step => ({ id: 'w1', steps: [step] });
const status = (x, i = 0) => x.appliedOperations[i].status;
const reason = (x, i = 0) => {
  const e = x.appliedOperations[i].evidence;
  return (e && e.evidence && e.evidence.reason) || (e && e.reason);
};

/** FIRST (APPLY, 1 command, 1 write, L1, L2, barrier, close) then SECOND (new session, 0/0, no barrier). */
async function firstAndSecond(store, task, { integral = false } = {}) {
  const first = await run(store, task);
  assert.equal(first.result.editOk, true, JSON.stringify(first.x.stopDetail || first.x.outcome));
  assert.ok(first.x.appliedOperations.every(e => e.status === OPERATION_STATUS.APPLIED));
  assert.equal(first.m.writes, task.steps.length);
  assert.equal(first.m.commands, task.steps.length);
  assert.equal(first.x.wholeTaskVerification.verdict, 'PASS');
  assert.equal(first.result.persistenceState, PERSISTENCE_STATE.EDITOR_CLEAN);
  assert.equal(first.driver.calls.forceSave, 1, 'one barrier, at task end');
  assert.equal(first.result.cleanupOutcome, 'SUCCESS');
  const integralCount = integral === true ? task.steps.length : (integral || 0);
  const l3 = first.x.preservationVerification || [];
  assert.equal(l3.length, integralCount, 'integral L3 only for P1 operations');
  assert.ok(l3.every(v => v.verdict === 'PASS'));
  const second = await run(store, task);
  assert.equal(second.result.ok, true);
  assert.ok(second.x.appliedOperations.every(e => e.status === OPERATION_STATUS.SATISFIED));
  assert.equal(second.m.writes, 0);
  assert.equal(second.m.commands, 0);
  assert.equal(second.driver.calls.forceSave, 0);
  assert.equal(second.result.persistenceState, PERSISTENCE_STATE.NOT_REQUIRED);
  assert.equal(second.x.wholeTaskVerification.verdict, 'PASS');
  return { first, second };
}

// ---- format_text (P1) ---------------------------------------------------------------------------

const formatTask = (format, selector = { text: P2 }) => one({ intent: 'format_text', selector, format });

test('format_text: FIRST applies with integral preservation, SECOND is a 0-write no-op', async () => {
  const store = createStore(doc());
  await firstAndSecond(store, formatTask({ bold: true, alignment: 'center', fontSize: 14 }), { integral: true });
  const p = store.durable.find(x => x.text === P2);
  assert.equal(p.format.bold, true);
  assert.equal(p.format.alignment, 'center');
  assert.deepEqual(order(store), [H, P1, P2, P3, P4]);
});

test('format_text: part of the format already holds -> APPLY, then the whole terminal state is proven', async () => {
  const d = doc().map(p => (p.text === P2 ? { ...p, props: { ...p.props, bold: true } } : p));
  await firstAndSecond(createStore(d), formatTask({ bold: true, italic: true }), { integral: true });
});

test('format_text negatives: ambiguous target, physical locator, unobservable property -> FAIL_CLOSED, 0 dispatch', async () => {
  const dup = await run(createStore([...doc(), { text: P2, props: { ...RICH } }]), formatTask({ bold: true }));
  assert.equal(reason(dup.x), 'format-target-ambiguous');
  const stale = await run(createStore(doc()), formatTask({ bold: true }, { text: P2, index: 2 }));
  assert.equal(reason(stale.x), 'format-target-physical_locator_refused', 'stale targeting rejection');
  const blind = doc().map(p => (p.text === P2 ? { ...p, format: { alignment: null } } : p));
  const unobs = await run(createStore(blind), formatTask({ alignment: 'center' }));
  assert.equal(reason(unobs.x), 'format-terminal-property-unobservable');
  const blindShading = doc().map(p => (p.text === P2 ? { ...p, format: { shading: null } } : p));
  const unobsShading = await run(createStore(blindShading), formatTask({ shading: 'FFFF00' }));
  assert.equal(reason(unobsShading.x), 'format-terminal-property-unobservable');
  for (const r of [dup, stale, unobs, unobsShading]) assert.equal(r.driver.calls.dispatch, 0);
});

test('format_text: dispatch ok but no mutation / only part applied -> never APPLIED (uncertain, no barrier)', async () => {
  for (const fault of ['ignore', 'format-partial']) {
    const { x, result, driver } = await run(createStore(doc()), formatTask({ bold: true, italic: true }), { fault });
    assert.equal(status(x), OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN, fault);
    assert.equal(x.appliedOperations[0].uncertainty.reason, UNCERTAINTY_REASON.VERIFY_CANNOT_PROVE_TERMINAL_STATE);
    assert.equal(result.persistenceState, PERSISTENCE_STATE.UNCERTAIN_MAY_PERSIST);
    assert.equal(driver.calls.forceSave, 0);
  }
});

test('format_text is P1: formatting holds but another paragraph was lost -> NOT APPLIED (integral)', async () => {
  const { x, driver } = await run(createStore(doc()), formatTask({ bold: true }), { fault: 'format-drop-other' });
  assert.equal(x.primitiveVerification[0].verdict, 'PASS', 'L1 (the formatting) held');
  assert.equal(x.preservationVerification[0].verdict, 'FAIL');
  assert.equal(x.appliedOperations[0].uncertainty.reason, UNCERTAINTY_REASON.INTEGRAL_PRESERVATION_NOT_PROVEN);
  assert.equal(driver.calls.forceSave, 0);
  const changed = await run(createStore(doc()), formatTask({ bold: true }), { fault: 'format-change-text' });
  assert.notEqual(status(changed.x), OPERATION_STATUS.APPLIED, 'target text changed by a format -> not applied');
});

test('format_text: images are part of the integral round-trip', async () => {
  const images = [{ src: 'img/logo.png', alt: 'logo' }];
  await firstAndSecond(createStore(doc(), { images }), formatTask({ italic: true }), { integral: true });
});

// W-09: indentation (points externally, twips at the Builder boundary) + shading (6-digit hex).
test('format_text: indentation (negative first-line = hanging) applies with integral preservation', async () => {
  const store = createStore(doc());
  await firstAndSecond(store, formatTask({ indentLeft: 36, indentRight: 18, indentFirstLine: -12 }), { integral: true });
  const p = store.durable.find(x => x.text === P2);
  assert.equal(p.format.indentLeft, 36);
  assert.equal(p.format.indentRight, 18);
  assert.equal(p.format.indentFirstLine, -12);
});

test('format_text: shading (hex colour) applies with integral preservation', async () => {
  const store = createStore(doc());
  await firstAndSecond(store, formatTask({ shading: 'ffff00' }), { integral: true });
  const p = store.durable.find(x => x.text === P2);
  assert.equal(p.format.shading, 'FFFF00', 'normalized to uppercase, case-insensitive input accepted');
});

// W-09-3: tabs, the four per-side borders, and the four pagination flags --
// MEASURED (W09-1a..1e): tabs REPLACE whole (one key, so the generic per-key write already gives
// this), each border side is independent/additive (four separate keys, same reason), false/0
// pagination is explicitly refused at plan time (not yet measured), units are points externally.
test('format_text: tabs apply with integral preservation; a second SetTabs call replaces the whole list', async () => {
  const store = createStore(doc());
  await firstAndSecond(store, formatTask({ tabs: [{ pos: 36, align: 'left' }, { pos: 72, align: 'right' }] }), { integral: true });
  const p = store.durable.find(x => x.text === P2);
  assert.deepEqual(p.format.tabs, [{ pos: 36, align: 'left' }, { pos: 72, align: 'right' }]);
  // a second, different tabs request on the same paragraph: REPLACE, not append (W09-1e measured)
  await firstAndSecond(store, formatTask({ tabs: [{ pos: 18, align: 'left' }] }), { integral: true });
  assert.deepEqual(store.durable.find(x => x.text === P2).format.tabs, [{ pos: 18, align: 'left' }]);
});

test('format_text: per-side borders apply with integral preservation; two sides on one paragraph both survive', async () => {
  const store = createStore(doc());
  await firstAndSecond(store, formatTask({ borderTop: { style: 'single', size: 1, color: '#FF0000' } }), { integral: true });
  let p = store.durable.find(x => x.text === P2);
  assert.deepEqual(p.format.borderTop, { style: 'single', size: 1, color: 'FF0000' }, 'stored without the leading #, matching the driver read shape');
  // a second, different side on the SAME paragraph: ADD/merge, not overwrite (W09-1e measured)
  await firstAndSecond(store, formatTask({ borderLeft: { style: 'single', size: 0.5, color: '#0000FF' } }), { integral: true });
  p = store.durable.find(x => x.text === P2);
  assert.deepEqual(p.format.borderTop, { style: 'single', size: 1, color: 'FF0000' }, 'the first side survives the second write');
  assert.deepEqual(p.format.borderLeft, { style: 'single', size: 0.5, color: '0000FF' });
});

test('format_text: pagination flags (keepNext/keepLines/pageBreakBefore/widowControl) apply true, with integral preservation', async () => {
  const store = createStore(doc());
  await firstAndSecond(store, formatTask({ keepNext: true, keepLines: true, pageBreakBefore: true, widowControl: true }), { integral: true });
  const p = store.durable.find(x => x.text === P2);
  assert.equal(p.format.keepNext, true);
  assert.equal(p.format.keepLines, true);
  assert.equal(p.format.pageBreakBefore, true);
  assert.equal(p.format.widowControl, true);
});

test('format_text plan: a non-boolean flag and malformed tabs/border shapes -> plan fails before any dispatch (NOT MEASURED, refused)', async () => {
  for (const step of [
    // GH #290: keepNext:false, a center tab, widowControl:false and tabs:[] are released (gh290-format-text-*.test.cjs);
    // decimal stays refused
    { intent: 'format_text', selector: { text: P2 }, format: { keepNext: 0 } },
    { intent: 'format_text', selector: { text: P2 }, format: { tabs: {} } },
    { intent: 'format_text', selector: { text: P2 }, format: { tabs: [{ pos: 36, align: 'decimal' }] } },
    { intent: 'format_text', selector: { text: P2 }, format: { tabs: [{ pos: -1, align: 'left' }] } },
    { intent: 'format_text', selector: { text: P2 }, format: { borderTop: { style: 'double', size: 1, color: '#FF0000' } } },
    { intent: 'format_text', selector: { text: P2 }, format: { borderTop: { style: 'single', size: 0, color: '#FF0000' } } },
    { intent: 'format_text', selector: { text: P2 }, format: { borderTop: { style: 'single', size: 1, color: 'red' } } },
  ]) {
    const { x, driver } = await run(createStore(doc()), one(step));
    assert.equal(x.outcome, 'plan-failed', JSON.stringify(step));
    assert.equal(driver.calls.dispatch, 0);
  }
});

test('format_text negatives: tabs/border unobservable on the target -> FAIL_CLOSED, 0 dispatch', async () => {
  const blindTabs = doc().map(p => (p.text === P2 ? { ...p, format: { tabs: undefined } } : p));
  const unobsTabs = await run(createStore(blindTabs), formatTask({ tabs: [{ pos: 36, align: 'left' }] }));
  assert.equal(reason(unobsTabs.x), 'format-terminal-property-unobservable');
  const blindBorder = doc().map(p => (p.text === P2 ? { ...p, format: { borderTop: undefined } } : p));
  const unobsBorder = await run(createStore(blindBorder), formatTask({ borderTop: { style: 'single', size: 1, color: '#FF0000' } }));
  assert.equal(reason(unobsBorder.x), 'format-terminal-property-unobservable');
  for (const r of [unobsTabs, unobsBorder]) assert.equal(r.driver.calls.dispatch, 0);
});

// RED-proof target: otherParagraphsNewFieldsUnchanged (formatting.cjs verifyPreservation) --
// the PRE-EXISTING whole-paragraph L3 only ever checked other paragraphs' TEXT/EXISTENCE; this
// is NEW coverage for the new fields' FORMAT. P1's text is untouched (old L3 would PASS) but its
// tabs were mutated as a side effect -- only the new check catches this.
test('format_text is P1: tabs formatting holds but ANOTHER paragraph\'s new-field format was side-effected -> NOT APPLIED (integral)', async () => {
  const { x, driver } = await run(createStore(doc()), formatTask({ tabs: [{ pos: 36, align: 'left' }] }), { fault: 'format-other-new-field-side-effect' });
  assert.equal(x.primitiveVerification[0].verdict, 'PASS', 'L1 (the tabs on the target) held');
  assert.equal(x.preservationVerification[0].verdict, 'FAIL', 'L3: a sibling paragraph\'s new-field format changed');
  assert.equal(x.appliedOperations[0].uncertainty.reason, UNCERTAINTY_REASON.INTEGRAL_PRESERVATION_NOT_PROVEN);
  assert.equal(driver.calls.forceSave, 0);
});

// W1.1 migrated the substring target (`matchText`): its acceptance is test/acceptance/docx-w11-partial-format.test.cjs.
test('format_text plan: unknown property, contained-selector target, empty format -> plan fails before any dispatch', async () => {
  for (const step of [
    { intent: 'format_text', selector: { text: P2 }, format: { colour: 'red' } },
    { intent: 'format_text', matchText: 'Second', format: { bold: true } },
    { intent: 'format_text', selector: { text: 'Second', contains: true }, format: { bold: true } },
    { intent: 'format_text', selector: { text: P2 }, format: {} },
    { intent: 'format_text', selector: { text: P2 }, format: { shading: 'yellow' } },
  ]) {
    const { x, driver } = await run(createStore(doc()), one(step));
    assert.equal(x.outcome, 'plan-failed', JSON.stringify(step));
    assert.equal(driver.calls.dispatch, 0);
  }
});

// ---- insert_paragraph_before / after (P0, directional) ------------------------------------------

test('insert_paragraph_after: TARGET NEW -- directional adjacency, SECOND no-op', async () => {
  const store = createStore(doc());
  await firstAndSecond(store, one({ intent: 'insert_paragraph_after', selector: { text: P2 }, text: 'NEW' }));
  assert.deepEqual(order(store), [H, P1, P2, 'NEW', P3, P4]);
});

test('insert_paragraph_before: NEW TARGET -- directional adjacency, SECOND no-op', async () => {
  const store = createStore(doc());
  await firstAndSecond(store, one({ intent: 'insert_paragraph_before', selector: { text: P2 }, text: 'NEW' }));
  assert.deepEqual(order(store), [H, P1, 'NEW', P2, P3, P4]);
});

test('NEGATIVE CONTROL insert: the editor puts it on the WRONG side -> never APPLIED', async () => {
  for (const intent of ['insert_paragraph_before', 'insert_paragraph_after']) {
    const { x } = await run(createStore(doc()), one({ intent, selector: { text: P2 }, text: 'NEW' }), { fault: 'wrong-side' });
    assert.equal(status(x), OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN, intent);
    assert.equal(x.primitiveVerification[0].evidence.evidence.directional, false);
  }
});

test('NEGATIVE CONTROL insert verifier: the same after-state with the side flipped FAILS', () => {
  const before = { primary: normalizeInventory({ html: htmlOf(doc()), paragraphs: doc() }) };
  const afterDoc = [...doc().slice(0, 3), { text: 'NEW', props: RICH }, ...doc().slice(3)];
  const after = { primary: normalizeInventory({ html: htmlOf(afterDoc), paragraphs: afterDoc }) };
  const op = side => ({ id: 'x', intent: `insert_paragraph_${side}`, selector: { text: P2 }, text: 'NEW' });
  assert.equal(paragraphFamily.verifyPrimitive(op('after'), before, after).ok, true);
  assert.equal(paragraphFamily.verifyPrimitive(op('before'), before, after).ok, false);
  assert.equal(paragraphFamily.wholeTaskChecks(op('before'), after).every(c => c.ok), false, 'L2 is directional too');
});

test('insert: text already present in the WRONG position -> FAIL_CLOSED, not a blind SATISFIED', async () => {
  const d = [...doc().slice(0, 2), { text: 'NEW', props: RICH }, ...doc().slice(2)]; // NEW before P2
  const { x, driver } = await run(createStore(d), one({ intent: 'insert_paragraph_after', selector: { text: P2 }, text: 'NEW' }));
  assert.equal(reason(x), 'insert-produced-text-in-wrong-position');
  assert.equal(driver.calls.dispatch, 0);
});

test('insert negatives: ambiguous anchor, duplicate produced text, inserted twice, no mutation', async () => {
  const amb = await run(createStore([...doc(), { text: P2, props: RICH }]), one({ intent: 'insert_paragraph_after', selector: { text: P2 }, text: 'NEW' }));
  assert.equal(reason(amb.x), 'insert-anchor-ambiguous');
  const twice = await run(createStore(doc()), one({ intent: 'insert_paragraph_after', selector: { text: P2 }, text: 'NEW' }), { fault: 'insert-twice' });
  assert.equal(status(twice.x), OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN);
  const none = await run(createStore(doc()), one({ intent: 'insert_paragraph_after', selector: { text: P2 }, text: 'NEW' }), { fault: 'ignore' });
  assert.equal(status(none.x), OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN);
});

// ---- delete_paragraph (P0, surviving context) ----------------------------------------------------

/** The selector a caller builds from a FRESH inventory. */
function selectorOf(paragraphs, text) {
  const b = normalizeInventory({ html: htmlOf(paragraphs), paragraphs }).blocks.find(x => x.text === text);
  return { text, kind: b.kind, before: b.locator.context.before, after: b.locator.context.after };
}

test('delete_paragraph: target absent + neighbours closed up; SECOND proves it from the surviving context', async () => {
  const store = createStore(doc());
  const { first, second } = await firstAndSecond(store, one({ intent: 'delete_paragraph', selector: selectorOf(doc(), P2) }));
  assert.deepEqual(order(store), [H, P1, P3, P4]);
  assert.equal(first.x.primitiveVerification[0].evidence.evidence.survivingContext.adjacent, true);
  assert.equal(second.x.appliedOperations[0].evidence.proof, 'fresh-live-surviving-context-adjacent');
});

test('delete: an absent target without context anchors is FAIL_CLOSED (a typo is absent too)', async () => {
  const store = createStore(doc());
  await run(store, one({ intent: 'delete_paragraph', selector: selectorOf(doc(), P2) }));
  const bare = await run(store, one({ intent: 'delete_paragraph', selector: { text: P2 } }));
  assert.equal(reason(bare.x), 'delete-terminal-identity-insufficient-selector');
  const typo = await run(createStore(doc()), one({ intent: 'delete_paragraph', selector: { text: 'Secnd body paragraph.', before: P1, after: P3 } }));
  assert.equal(reason(typo.x), 'delete-surviving-context-not-found', 'P1 and P3 are not adjacent: nothing was deleted between them');
});

test('delete negatives: the wrong paragraph deleted / nothing deleted -> never APPLIED; ambiguous -> FAIL_CLOSED', async () => {
  for (const fault of ['delete-neighbour', 'ignore']) {
    const { x } = await run(createStore(doc()), one({ intent: 'delete_paragraph', selector: selectorOf(doc(), P2) }), { fault });
    assert.equal(status(x), OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN, fault);
  }
  const amb = await run(createStore([...doc(), { text: P2, props: RICH }]), one({ intent: 'delete_paragraph', selector: { text: P2 } }));
  assert.equal(reason(amb.x), 'delete-target-ambiguous');
});

// ---- move_paragraph_before / after (P0, adjacency + direction + old position closed) ------------

test('move_paragraph_after: target once, destination once, adjacent after it, old position closed; SECOND no-op', async () => {
  const store = createStore(doc());
  const { first } = await firstAndSecond(store, one({ intent: 'move_paragraph_after', selector: { text: P1 }, destinationSelector: { text: P3 } }));
  assert.deepEqual(order(store), [H, P2, P3, P1, P4]);
  const ev = first.x.primitiveVerification[0].evidence.evidence;
  assert.equal(ev.directional, true);
  assert.equal(ev.oldPositionClosed, true);
});

test('move_paragraph_before: TARGET DESTINATION; SECOND no-op', async () => {
  const store = createStore(doc());
  await firstAndSecond(store, one({ intent: 'move_paragraph_before', selector: { text: P4 }, destinationSelector: { text: P2 } }));
  assert.deepEqual(order(store), [H, P1, P4, P2, P3]);
});

test('NEGATIVE CONTROLS move: wrong side, not adjacent, duplicate target left behind -> never APPLIED', async () => {
  for (const fault of ['wrong-side', 'move-not-adjacent', 'no-delete', 'ignore']) {
    const { x } = await run(createStore(doc()), one({ intent: 'move_paragraph_after', selector: { text: P1 }, destinationSelector: { text: P3 } }), { fault });
    assert.equal(status(x), OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN, fault);
  }
});

test('move negatives: duplicate target, duplicate destination, source = destination, destination ref -> fail closed, 0 dispatch', async () => {
  const dupT = await run(createStore([...doc(), { text: P1, props: RICH }]), one({ intent: 'move_paragraph_after', selector: { text: P1 }, destinationSelector: { text: P3 } }));
  assert.equal(reason(dupT.x), 'move-target-ambiguous');
  const dupD = await run(createStore([...doc(), { text: P3, props: RICH }]), one({ intent: 'move_paragraph_after', selector: { text: P1 }, destinationSelector: { text: P3 } }));
  assert.equal(reason(dupD.x), 'move-destination-ambiguous');
  const same = await run(createStore(doc()), one({ intent: 'move_paragraph_after', selector: { text: P1 }, destinationSelector: { text: P1 } }));
  assert.equal(reason(same.x), 'move-source-equals-destination');
  const dref = await run(createStore(doc()), { id: 'x', steps: [
    { intent: 'insert_paragraph_after', selector: { text: P2 }, text: 'NEW' },
    { intent: 'move_paragraph_after', selector: { text: P4 }, destinationSelector: { ref: 's0' } },
  ] });
  assert.equal(dref.x.outcome, 'plan-failed');
  for (const r of [dupT, dupD, same, dref]) assert.equal(r.driver.calls.dispatch, 0);
});

test('move verifier: "both exist" is not a pass', () => {
  const before = { primary: normalizeInventory({ html: htmlOf(doc()), paragraphs: doc() }) };
  const same = before; // nothing moved: target and destination both exist
  const op = { id: 'x', intent: 'move_paragraph_after', selector: { text: P1 }, destinationSelector: { text: P3 } };
  assert.equal(paragraphFamily.verifyPrimitive(op, before, same).ok, false);
});

// ---- dependency / ref: REF_IS_NOT_STALE_LOCATOR --------------------------------------------------

test('REF: A creates a paragraph, B inserts after A\'s RESULT resolved from the fresh post-A state', async () => {
  const store = createStore(doc());
  const task = { id: 'ref', steps: [
    { intent: 'insert_paragraph_after', selector: { text: P2 }, text: 'NEW A' },
    { intent: 'insert_paragraph_after', selector: { ref: 's0' }, text: 'NEW B' },
  ] };
  const planned = planOperations(task);
  assert.equal(planned.operations[1].refText, 'NEW A', 'the ref names A\'s semantic text, not a position');
  const { first } = await firstAndSecond(store, task);
  assert.deepEqual(order(store), [H, P1, P2, 'NEW A', 'NEW B', P3, P4]);
  assert.ok(first.m.freshReads >= 3, 'B was targeted from a fresh read after A');
});

test('DEPENDENCY: A inserts; B moves another paragraph after A\'s text (dependsOn, semantic destination)', async () => {
  const store = createStore(doc());
  await firstAndSecond(store, { id: 'dep', steps: [
    { intent: 'insert_paragraph_after', selector: { text: P1 }, text: 'NEW A' },
    { intent: 'move_paragraph_after', selector: { text: P4 }, destinationSelector: { text: 'NEW A' }, dependsOn: ['s0'] },
  ] });
  assert.deepEqual(order(store), [H, P1, 'NEW A', P4, P2, P3]);
});

test('ref negatives: a ref to a non-producing operation, or to no earlier operation -> plan fails', async () => {
  for (const steps of [
    [{ intent: 'delete_paragraph', selector: { text: P3 } }, { intent: 'insert_paragraph_after', selector: { ref: 's0' }, text: 'X' }],
    [{ intent: 'insert_paragraph_after', selector: { ref: 'nope' }, text: 'X' }],
  ]) {
    const { x, driver } = await run(createStore(doc()), { id: 'r', steps });
    assert.equal(x.outcome, 'plan-failed');
    assert.equal(driver.calls.dispatch, 0);
  }
});

// ---- mixed Wave-1 chain in ONE session -----------------------------------------------------------

test('MIXED: rewrite -> format(ref) -> insert(ref) -> move(dependsOn), one session, one barrier; retry all SATISFIED', async () => {
  const store = createStore(doc());
  const s0 = selectorOf(doc(), P2);
  const task = { id: 'mixed', steps: [
    { intent: 'rewrite_paragraph', selector: { text: s0.text, kind: s0.kind, before: s0.before }, text: 'REWRITTEN' },
    { intent: 'format_text', selector: { ref: 's0' }, format: { bold: true } },
    { intent: 'insert_paragraph_after', selector: { ref: 's0' }, text: 'INSERTED' },
    { intent: 'move_paragraph_after', selector: { text: P4 }, destinationSelector: { text: 'INSERTED' }, dependsOn: ['s2'] },
  ] };
  await firstAndSecond(store, task, { integral: 2 }); // rewrite and format are P1; insert and move are P0
  assert.deepEqual(order(store), [H, P1, 'REWRITTEN', 'INSERTED', P4, P3]);
  assert.equal(store.durable.find(p => p.text === 'REWRITTEN').format.bold, true);
});
