'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { PERSISTENCE_STATE } = require('../../src/core/persistence/verified-writes.cjs');
const { docxAdapter, planOperations } = require('../../src/adapters/docx/adapter.cjs');
const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');

/**
 * W-10 offline acceptance: define_style/modify_style through the TARGET path on the fake
 * transport, modeled directly on test/acceptance/docx-w11-page-setup.test.cjs's own shape. The
 * fake model follows the LIVE probes (W10-1a..1c #ref/#ref, W10-2 #ref/#ref): styles
 * keyed by NAME (never styleId), text{bold,size,color} + paragraph{align,indentLeft,spacingAfter}.
 */

const paragraphs = () => [{ text: 'Body heading', level: 1, props: {} }, { text: 'Body first.', props: {} }];
const store = (styles = {}, paras = paragraphs()) => createStore(paras, { doc: { styles } });

async function run(s, task, script = {}) {
  let t = 0;
  const driver = createFakeDocxDriver(s, script);
  const options = { driver, fileId: 'd', now: () => (t += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  // W-10-8: the caller no longer names the styles its task targets -- the executor's
  // optional `adapter.observationRequestFor(task)` hook (docxAdapter, executor.cjs) derives it from
  // the task itself, the SAME derivation `styleNamesOf` always did, now run automatically. A plain
  // `createExecutor(task)`, with no options at all, is the real production call shape.
  const result = await withSession(options, docxAdapter, createExecutor(task));
  if (result.execution.operationAccounting) assert.equal(result.execution.operationAccounting.balanced, true);
  return { result, driver, m: result.metrics, x: result.execution };
}
const task = (...steps) => ({ id: 'w10', steps });
const define = (name, fields, extra = {}) => ({ intent: 'define_style', name, ...fields, ...extra });
const modify = (name, fields, extra = {}) => ({ intent: 'modify_style', name, ...fields, ...extra });
const status = (x, i = 0) => x.appliedOperations[i].status;
const reason = (x, i = 0) => { const e = x.appliedOperations[i].evidence; return (e && e.evidence && e.evidence.reason) || (e && e.reason); };
const styleOf = (s, name) => s.durableDoc.styles[name];
const snapshot = s => JSON.stringify([s.durable.map(p => p.text), s.durableDoc.styles]);

async function firstAndSecond(s, t) {
  const first = await run(s, t);
  assert.equal(first.result.editOk, true, JSON.stringify(first.x.stopDetail || first.x.outcome || first.x.appliedOperations.map(e => [e.status, e.evidence])));
  assert.ok(first.x.appliedOperations.every(e => e.status === OPERATION_STATUS.APPLIED), JSON.stringify(first.x.appliedOperations.map(e => [e.status, e.evidence])));
  assert.equal(first.m.writes, t.steps.length);
  assert.equal(first.m.commands, t.steps.length, 'FIRST_COMMANDS: one command per operation');
  assert.ok(first.x.primitiveVerification.every(v => v.verdict === 'PASS'), JSON.stringify(first.x.primitiveVerification));
  assert.ok(first.x.preservationVerification.every(v => v.verdict === 'PASS'), JSON.stringify(first.x.preservationVerification));
  assert.equal(first.x.wholeTaskVerification.verdict, 'PASS', JSON.stringify(first.x.wholeTaskVerification));
  assert.equal(first.result.persistenceState, PERSISTENCE_STATE.EDITOR_CLEAN);
  const second = await run(s, t);
  assert.equal(second.result.ok, true, JSON.stringify(second.x.stopDetail || second.x.outcome || second.x.appliedOperations.map(e => [e.status, e.evidence])));
  assert.ok(second.x.appliedOperations.every(e => e.status === OPERATION_STATUS.SATISFIED), JSON.stringify(second.x.appliedOperations.map(e => [e.status, e.evidence])));
  assert.deepEqual([second.m.writes, second.m.commands, second.driver.calls.forceSave], [0, 0, 0]);
  assert.equal(second.x.wholeTaskVerification.verdict, 'PASS');
  return { first, second };
}

test('define_style: write; the style now exists with every requested field. RETRY is idempotent (GH #264): the same call again finds the name AND every requested field matching -> SATISFIED, 0 writes', async () => {
  const s = store();
  const t = task(define('W10Report', { text: { bold: true, size: 14, color: 'FF0000' }, paragraph: { align: 'center', indentLeft: 10, spacingAfter: 6 } }));
  const first = await run(s, t);
  assert.equal(first.result.editOk, true, JSON.stringify(first.x.stopDetail || first.x.appliedOperations.map(e => [e.status, e.evidence])));
  assert.ok(first.x.appliedOperations.every(e => e.status === OPERATION_STATUS.APPLIED));
  assert.equal(first.m.writes, 1);
  //: the fake stores size/indentLeft/spacingAfter in the SAME raw units as the real
  // driver's file (half-points / twips), mirroring commandBody's own write-side conversion --
  // 14pt -> 28 half-pt, 10pt -> 200 twips, 6pt -> 120 twips.
  assert.deepEqual(styleOf(s, 'W10Report'), { text: { bold: true, size: 28, color: 'FF0000' }, paragraph: { align: 'center', indentLeft: 200, spacingAfter: 120 } });
  const before = snapshot(s);
  const second = await run(s, t);
  assert.equal(status(second.x), OPERATION_STATUS.SATISFIED, JSON.stringify(second.x.appliedOperations));
  assert.equal(second.m.writes, 0);
  assert.equal(snapshot(s), before, 'a satisfied define writes nothing');
});

test('define_style: a name that already exists with the SAME requested fields is SATISFIED; with a DIFFERENT field it is FAIL_CLOSED with a NAMED reason listing the differing fields -- and it never writes or overwrites', async () => {
  const s = store({ Existing: { text: { bold: true, size: 28 }, paragraph: { align: 'center' } } });
  const before = snapshot(s);
  const same = await run(s, task(define('Existing', { text: { bold: true, size: 14 } })));
  assert.equal(status(same.x), OPERATION_STATUS.SATISFIED, JSON.stringify(same.x.appliedOperations));
  assert.equal(same.m.writes, 0);
  const diff = await run(s, task(define('Existing', { text: { bold: true, size: 16 }, paragraph: { align: 'left' } })));
  assert.equal(status(diff.x), OPERATION_STATUS.FAILED, JSON.stringify(diff.x.appliedOperations));
  assert.equal(reason(diff.x), 'define_style-name-already-exists-with-different-fields:Existing:text.size,paragraph.align');
  assert.equal(diff.m.writes, 0);
  assert.equal(before, snapshot(s), 'a refused define must never write');
  // a requested field the style does not carry at all (read back as unset) also differs -- never a guessed match
  const unset = await run(s, task(define('Existing', { text: { color: 'FF0000' } })));
  assert.equal(status(unset.x), OPERATION_STATUS.FAILED);
  assert.match(reason(unset.x), /text\.color$/);
  // an UNOBSERVABLE requested field is never a match either
  const u = store({ Existing: { text: { bold: true }, paragraph: {} } });
  const un = await run(u, task(define('Existing', { text: { bold: true } })), { styleUnobservableFields: ['text.bold'] });
  assert.equal(status(un.x), OPERATION_STATUS.FAILED, JSON.stringify(un.x.appliedOperations));
  assert.match(reason(un.x), /define_style-name-already-exists-with-different-fields:Existing:text\.bold$/);
});

test('modify_style: write + retry SATISFIED 0/0; every OTHER style unchanged', async () => {
  const s = store({ Target: { text: {}, paragraph: {} }, Other: { text: { bold: true }, paragraph: {} } });
  await firstAndSecond(s, task(modify('Target', { text: { size: 24 } })));
  assert.equal(styleOf(s, 'Target').text.size, 48); // 24pt -> 48 half-pt, raw storage
  assert.deepEqual(styleOf(s, 'Other'), { text: { bold: true }, paragraph: {} });
});

test('modify_style: a name that does not exist is FAIL_CLOSED, never creates one', async () => {
  const s = store();
  const before = snapshot(s);
  const { x, m } = await run(s, task(modify('Ghost', { text: { bold: true } })));
  assert.equal(status(x), OPERATION_STATUS.FAILED, JSON.stringify(x.appliedOperations));
  assert.equal(before, snapshot(s));
  assert.equal(m.writes, 0);
});

test('modify_style: a DIFFERENT value on the same field is a real write, never a false SATISFIED', async () => {
  // the SEED is raw storage (half-points): 20 half-pt = 10pt pre-existing.
  const s = store({ Target: { text: { size: 20 }, paragraph: {} } });
  const before = snapshot(s);
  const { x, m, driver } = await run(s, task(modify('Target', { text: { size: 20 } }))); // requests 20pt -- DIFFERENT from the pre-existing 10pt
  assert.notEqual(before, snapshot(s));
  assert.equal(status(x), OPERATION_STATUS.APPLIED);
  assert.equal(m.writes, 1);
  assert.equal(driver.calls.dispatch, 1);
  assert.equal(styleOf(s, 'Target').text.size, 40); // 20pt -> 40 half-pt, raw storage
});

test('plan rejects: no text/paragraph field at all', () => {
  const planned = planOperations(task({ intent: 'define_style', name: 'NoFields' }));
  assert.equal(planned.ok, false);
  assert.match(planned.detail, /requires at least one of text, paragraph/);
});

test('plan (GH #292 S8): text.color "auto" is accepted (SetColor(Api.AutoColor()), LIVE); a non-hex, non-auto colour is still refused', () => {
  assert.equal(planOperations(task(define('W10Auto', { text: { color: 'auto' } }))).ok, true);
  const bad = planOperations(task(define('W10Auto', { text: { color: 'red' } })));
  assert.equal(bad.ok, false);
  assert.match(bad.detail, /6-hex-digit string or 'auto'/);
});

test('plan (GH #292 S3): define_style basedOn is accepted as a style name; on modify_style, blank, or naming the style itself it is refused', () => {
  const planned = planOperations(task(define('W10Based', { text: { bold: true } }, { basedOn: 'Normal' })));
  assert.equal(planned.ok, true, planned.detail);
  assert.equal(planOperations(task(define('W10Based', {}, { basedOn: 'Normal' }))).ok, true, 'basedOn alone is a style definition');
  for (const [step, re] of [[define('W10Based', { text: { bold: true } }, { basedOn: '  ' }), /basedOn must be a non-empty style name/],
    [define('W10Based', { text: { bold: true } }, { basedOn: 'W10Based' }), /basedOn names the style itself/],
    [{ intent: 'modify_style', name: 'Normal', text: { bold: true }, basedOn: 'Heading 1' }, /modify_style.basedOn is not yet measured/]]) {
    const r = planOperations(task(step));
    assert.equal(r.ok, false);
    assert.match(r.detail, re);
  }
});

test('plan rejects: an unsupported field inside text/paragraph', () => {
  const planned = planOperations(task(define('W10Bad', { text: { highlight: 'yellow' } })));
  assert.equal(planned.ok, false);
  assert.match(planned.detail, /field-unsupported/);
});

test('faults: a sibling style changes too (SD-other-style-changed) -> L3 preservation FAIL', async () => {
  const { x } = await run(store({ Target: { text: {}, paragraph: {} }, Other: { text: {}, paragraph: {} } }), task(modify('Target', { text: { bold: true } })), { fault: 'sd-other-style-changed' });
  assert.equal(x.preservationVerification[0].verdict, 'FAIL', 'a changed sibling style must not pass preservation');
});

test('faults: the document body changes too (sd-body-change) -> L3 preservation FAIL', async () => {
  const { x } = await run(store({ Target: { text: {}, paragraph: {} } }), task(modify('Target', { text: { bold: true } })), { fault: 'sd-body-change' });
  assert.equal(x.preservationVerification[0].verdict, 'FAIL', 'a changed body must not pass preservation');
});

test('UNOBSERVABLE: a requested field that never becomes readable after the write -> L1 FAIL (never a guessed PASS)', async () => {
  const { x } = await run(store(), task(define('W10Unobs', { text: { bold: true } })), { styleUnobservableFields: ['text.bold'] });
  assert.equal(x.primitiveVerification[0].verdict, 'FAIL', 'an unobservable requested field must never read back as matching');
});

test('UNOBSERVABLE pre-dispatch: an unobservable field on modify_style does NOT block classify (it dispatches a real write, never a guessed SATISFIED) -- the write is still attempted, even though verification afterward stays uncertain for the SAME reason', async () => {
  const { x, driver } = await run(store({ Target: { text: { bold: true }, paragraph: {} } }), task(modify('Target', { text: { bold: true } })), { styleUnobservableFields: ['text.bold'] });
  assert.equal(driver.calls.dispatch, 1, 'pre-dispatch must NOT have blocked the write (classify must not have returned FAIL_CLOSED)');
  assert.notEqual(status(x), OPERATION_STATUS.SATISFIED, 'never a guessed SATISFIED from an unobservable field');
  assert.equal(status(x), OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN, JSON.stringify(x.appliedOperations));
});

test('UNIT BOUNDARY: text.size is carried in PT end-to-end (request -> lower *2 -> raw half-point storage -> read /2) -- round-trips through the SAME unit transition the real driver applies, never the raw half-point value leaking into the pt-based field', async () => {
  const s = store();
  const { x } = await run(s, task(define('W10Unit', { text: { size: 26 } }))); // a maintainer's exact LIVE example: 26pt
  assert.equal(status(x), OPERATION_STATUS.APPLIED, JSON.stringify(x.appliedOperations));
  assert.ok(x.primitiveVerification.every(v => v.verdict === 'PASS'), JSON.stringify(x.primitiveVerification));
  // the RAW storage (what a real saved file's rPr.sz would carry) is half-points -- 52, not 26.
  assert.equal(styleOf(s, 'W10Unit').text.size, 52, 'raw storage must be half-points (the write-side conversion)');
  // a RETRY with the SAME pt request must read back as SATISFIED -- proving the read side
  // converts 52 back to 26 before the family ever compares it to the requested pt value. Before
  // this unit boundary existed, the fake's round-trip was trivially consistent (no conversion
  // on either side) and could NOT have caught the real driver's missing /2 -- this test is the
  // dedicated regression lock zero asked for.
  const retry = await run(s, task(modify('W10Unit', { text: { size: 26 } })));
  assert.equal(status(retry.x), OPERATION_STATUS.SATISFIED, JSON.stringify(retry.x.appliedOperations));
  assert.equal(retry.m.writes, 0);
});

test('vm: commandBody is syntactically well-formed for both define and modify, any field combination', () => {
  const vm = require('node:vm');
  const { commandBody } = require('../../src/adapters/docx/families/style-definition.cjs');
  for (const mode of ['define', 'modify']) {
    for (const fields of [{ text: { bold: true, size: 14, color: 'FF0000' } }, { paragraph: { align: 'center', indentLeft: 10, spacingAfter: 6 } }, { text: { bold: true }, paragraph: { align: 'left' } }]) {
      assert.doesNotThrow(() => new vm.Script(`(function(){${commandBody({ family: 'styleDefinition', action: 'styleDefinition', mode, name: 'X', fields })}})`), JSON.stringify({ mode, fields }));
    }
  }
});

/**
 * W-10-8 (point 2): styleNames-passing must not be the CALLER's job -- a plain
 * `createExecutor(task)`, no options at all, must still see a style the SAME task just defined,
 * because `docxAdapter.observationRequestFor(task)` (adapter.cjs) now derives it automatically
 * from the task's own steps (the executor's optional hook, executor.cjs). A define-then-modify
 * task in ONE run is the right probe: the modify step's FRESH targeting observation (taken AFTER
 * the define step dispatched) is exactly the read that `GetAllStyles()` alone cannot see.
 */
test('RED/GREEN: a plain createExecutor(task), with NO options, still sees a style this SAME task just defined -- the adapter derives styleNames, the caller names nothing', async () => {
  const s = store();
  let t = 0;
  const driver = createFakeDocxDriver(s, {});
  const options = { driver, fileId: 'd', now: () => (t += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const t2 = task(define('W10AutoSeen', { text: { bold: true } }), modify('W10AutoSeen', { text: { size: 20 } }));
  const result = await withSession(options, docxAdapter, createExecutor(t2)); // no options -- the real production call shape
  assert.equal(result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED, JSON.stringify(result.execution.appliedOperations));
  assert.equal(result.execution.appliedOperations[1].status, OPERATION_STATUS.APPLIED, JSON.stringify(result.execution.appliedOperations));
  assert.deepEqual(styleOf(s, 'W10AutoSeen').text, { bold: true, size: 40 }); // 20pt -> 40 half-pt, the SAME raw storage the UNIT BOUNDARY test above exercises
});

test('RED/GREEN: reverting the adapter hook reproduces the exact pre-fix blindness -- WITHOUT it, even the define step\'s OWN self-verification cannot read the style it just created, and the task stops there', async () => {
  const s = store();
  let t = 0;
  const driver = createFakeDocxDriver(s, {});
  const options = { driver, fileId: 'd', now: () => (t += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  // a shallow clone WITHOUT observationRequestFor -- exactly the adapter shape before this fix
  const blindAdapter = { ...docxAdapter };
  delete blindAdapter.observationRequestFor;
  const t2 = task(define('W10Blind', { text: { bold: true } }), modify('W10Blind', { text: { size: 20 } }));
  const result = await withSession(options, blindAdapter, createExecutor(t2));
  // this is the EXACT W10-6 finding, reproduced: without a targeted name list, no read (not even
  // the define step's own L1 verify, in the SAME task) can see a style GetAllStyles() never listed.
  assert.equal(result.execution.appliedOperations[0].status, OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN, JSON.stringify(result.execution.appliedOperations));
  assert.equal(result.execution.appliedOperations[0].evidence.evidence.reason, 'style-not-found-after-write');
  assert.equal(result.execution.appliedOperations[1].status, OPERATION_STATUS.SKIPPED_DUE_TO_PRIOR_FAILURE, 'the uncertain define stops the task before modify is even attempted');
});

test('observationRequestFor(task): {} for a task with no define_style/modify_style steps -- every other family\'s request is unchanged by this hook\'s existence', () => {
  const t2 = task({ intent: 'rewrite_paragraph', selector: { paragraphOrdinal: 0 }, text: 'x' });
  assert.deepEqual(docxAdapter.observationRequestFor(t2), {});
});
