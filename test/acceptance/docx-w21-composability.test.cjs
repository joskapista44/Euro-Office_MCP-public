'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { PERSISTENCE_STATE } = require('../../src/core/persistence/verified-writes.cjs');
const { docxAdapter, planOperations } = require('../../src/adapters/docx/adapter.cjs');
const { normalizeInventory, resolveTextUnit } = require('../../src/adapters/docx/inventory.cjs');
const { createStore, createFakeDocxDriver, htmlOf, runPayload, formatOf } = require('../fakes/fake-docx-driver.cjs');

/**
 * M4 W2.1 cross-family composability (owner brief W2.1): an earlier Wave-1 operation whose semantic
 * result a LATER list step declaredly turns into a list item keeps its final postconditions -- checked
 * on the representation-free text unit -- and nothing else is relaxed.
 */

const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const H = 'Intro heading';
const A = 'Alpha item.';
const B = 'Beta item.';
const Z = 'Closing paragraph.';
const NEW = 'Rewritten KEY item.';
const doc = () => [
  { text: H, level: 1, props: { ...PLAIN, bold: true } },
  { text: A, props: { ...PLAIN } },
  { text: B, props: { ...PLAIN, italic: true } }, // a non-trivial P1 property baseline for the rewrite
  { text: Z, props: { ...PLAIN } },
];
const RW = { intent: 'rewrite_paragraph', selector: { text: B, kind: 'paragraph', before: A }, text: NEW };

async function run(store, task, script = {}) {
  let t = 0;
  const driver = createFakeDocxDriver(store, script);
  const options = { driver, fileId: 'd', now: () => (t += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, docxAdapter, createExecutor(task));
  assert.equal(driver.calls.open, 1, 'OPEN_COUNT=1');
  assert.equal(driver.calls.close, 1, 'CLOSE_COUNT=1');
  return { result, driver, m: result.metrics, x: result.execution };
}
const task = (...steps) => ({ id: 'w21', steps });
const failedL2 = x => ((x.wholeTaskVerification && x.wholeTaskVerification.evidence && x.wholeTaskVerification.evidence.checks) || []).filter(c => !c.ok).map(c => c.name);

async function oneSessionAndRetry(store, t) {
  const first = await run(store, t);
  assert.equal(first.result.editOk, true, JSON.stringify({ outcome: first.x.outcome, l2: failedL2(first.x) }));
  assert.ok(first.x.appliedOperations.every(e => e.status === OPERATION_STATUS.APPLIED));
  assert.equal(first.m.writes, t.steps.length);
  assert.ok(first.x.primitiveVerification.every(v => v.verdict === 'PASS'), 'L1');
  assert.ok(first.x.preservationVerification.every(v => v.verdict === 'PASS'), 'L3 (every step is P1 here)');
  assert.equal(first.x.wholeTaskVerification.verdict, 'PASS', 'L2: earlier postconditions + list terminal state');
  assert.equal(first.driver.calls.forceSave, 1, 'ONE persistence barrier');
  assert.equal(first.result.persistenceState, PERSISTENCE_STATE.EDITOR_CLEAN);
  const second = await run(store, t);
  assert.equal(second.result.ok, true, JSON.stringify({ outcome: second.x.outcome, l2: failedL2(second.x) }));
  assert.ok(second.x.appliedOperations.every(e => e.status === OPERATION_STATUS.SATISFIED), JSON.stringify(second.x.appliedOperations.map(e => e.status)));
  assert.equal(second.m.writes, 0);
  assert.equal(second.m.commands, 0);
  assert.equal(second.driver.calls.forceSave, 0);
  assert.equal(second.result.persistenceState, PERSISTENCE_STATE.NOT_REQUIRED);
  return { first, second };
}
const durable = (store, text) => store.durable.find(p => p.text === text);
const chars = p => runPayload(p).runs.flatMap(r => [...r.text].map(() => r.format));

// ---- the planner metadata -----------------------------------------------------------------------

test('planner: successorTransforms names exactly the LATER, DECLARED list steps on the same semantic result', () => {
  const p = planOperations(task(RW, { intent: 'format_text', selector: { ref: 's0' }, format: { bold: true } }, { intent: 'apply_list', selector: { ref: 's0' }, type: 'bullet' }, { intent: 'insert_paragraph_after', selector: { text: Z }, text: 'Other.' }));
  assert.equal(p.ok, true, p.detail);
  assert.deepEqual(p.operations[0].successorTransforms, ['s2']);
  assert.deepEqual(p.operations[1].successorTransforms, ['s2']);
  assert.equal(p.operations[3].successorTransforms, undefined, 'no successor -> the strict check stays');
  const none = planOperations(task(RW, { intent: 'format_text', selector: { ref: 's0' }, format: { bold: true } }));
  assert.ok(none.operations.every(o => o.successorTransforms === undefined), 'a non-transforming successor declares nothing');
});

test('semantic identity: paragraph -> list item is the same text unit; ambiguity is not resolved by text alone', () => {
  const d = [{ text: A, props: PLAIN }, { text: NEW, props: PLAIN, list: { type: 'bullet', level: 0, numId: 'n1' } }];
  const inv = normalizeInventory({ html: htmlOf(d), paragraphs: d.map(p => ({ text: p.text, props: p.props, format: formatOf(p), ...runPayload(p), numbering: p.list ? { present: true, level: 0, numId: 'n1' } : { present: false } })) });
  const u = resolveTextUnit(inv, NEW);
  assert.equal(u.ok, true);
  assert.equal(u.unit.kind, 'list_item');
  assert.equal(u.unit.prev, A, 'context from the neighbouring UNIT, not the stale block');
  assert.ok(u.unit.node.runs, 'the list item carries its own run/format read');
  const dup = [...d, { text: NEW, props: PLAIN }];
  const inv2 = normalizeInventory({ html: htmlOf(dup), paragraphs: dup.map(p => ({ text: p.text, props: p.props, format: formatOf(p), ...runPayload(p) })) });
  assert.equal(resolveTextUnit(inv2, NEW).reason, 'text-unit-ambiguous');
});

// ---- required mixed tasks (one session, retry 0/0) ----------------------------------------------

test('A: rewrite -> apply_list on its result (ref): one session, earlier guarantees kept, retry SATISFIED 0/0', async () => {
  const store = createStore(doc());
  await oneSessionAndRetry(store, task(RW, { intent: 'apply_list', selector: { ref: 's0' }, type: 'bullet', level: 0 }));
  const p = durable(store, NEW);
  assert.equal(p.list.type, 'bullet');
  assert.equal(p.props.italic, true, 'the rewrite P1 property survived the list conversion');
});

test('B: format -> apply_list on the same paragraph (ref): the requested formatting survives the conversion', async () => {
  const store = createStore(doc());
  await oneSessionAndRetry(store, task({ intent: 'format_text', selector: { text: B }, format: { bold: true } }, { intent: 'apply_list', selector: { ref: 's0' }, type: 'numbered', level: 0 }));
  const p = durable(store, B);
  assert.equal(p.list.type, 'numbered');
  assert.equal(formatOf(p).bold, true);
});

test('C: rewrite -> format(substring) -> apply_list -> set_list_level, one session, retry SATISFIED 0/0', async () => {
  const store = createStore(doc());
  await oneSessionAndRetry(store, task(
    RW,
    { intent: 'format_text', selector: { ref: 's0' }, matchText: 'KEY', format: { bold: true } },
    { intent: 'apply_list', selector: { ref: 's0' }, type: 'bullet', level: 0 },
    { intent: 'set_list_level', selector: { ref: 's0' }, level: 1 },
  ));
  const p = durable(store, NEW);
  assert.deepEqual({ type: p.list.type, level: p.list.level }, { type: 'bullet', level: 1 });
  const c = chars(p);
  const at = NEW.indexOf('KEY');
  for (let i = 0; i < NEW.length; i += 1) assert.equal(c[i].bold, i >= at && i < at + 3, `bold at ${i}`);
  assert.ok(c.every(f => f.italic === true), 'the rewrite P1 property kept on every character');
});

// ---- negative controls: composability must not hide a real regression --------------------------

test('NEG: rewrite text lost in the list step -> never APPLIED', async () => {
  const { x, result } = await run(createStore(doc()), task(RW, { intent: 'apply_list', selector: { ref: 's0' }, type: 'bullet' }), { fault: 'list-change-text', faultOn: [2] });
  assert.notEqual(x.appliedOperations[1].status, OPERATION_STATUS.APPLIED);
  assert.equal(result.ok, false);
});

test('NEG: rewrite P1 properties / requested formatting lost in the list conversion -> the list step L3 FAILS', async () => {
  for (const t of [
    task(RW, { intent: 'apply_list', selector: { ref: 's0' }, type: 'bullet' }),
    task({ intent: 'format_text', selector: { text: B }, format: { bold: true } }, { intent: 'apply_list', selector: { ref: 's0' }, type: 'numbered' }),
  ]) {
    const { x, result } = await run(createStore(doc()), t, { fault: 'list-drop-format', faultOn: [2] });
    assert.equal(x.primitiveVerification[1].verdict, 'PASS', 'the list state itself is right');
    assert.equal(x.preservationVerification[1].verdict, 'FAIL', 'but the target formatting did not survive');
    assert.notEqual(x.appliedOperations[1].status, OPERATION_STATUS.APPLIED);
    assert.equal(result.ok, false);
  }
});

test('NEG: list terminal type/level wrong -> FAIL', async () => {
  const { x, result } = await run(createStore(doc()), task(RW, { intent: 'apply_list', selector: { ref: 's0' }, type: 'bullet', level: 0 }), { fault: 'list-wrong-level', faultOn: [2] });
  assert.equal(x.primitiveVerification[1].verdict, 'FAIL');
  assert.equal(result.ok, false);
});

test('NEG: an UNDECLARED representation change does not relax the earlier L2', async () => {
  // end to end: the rewrite alone, but the transport turned the paragraph into a list item -> not APPLIED
  const { x, result } = await run(createStore(doc()), task(RW), { fault: 'rewrite-into-list' });
  assert.equal(result.ok, false);
  assert.notEqual(x.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
  // and the L2 itself: the same final state, WITHOUT a declared successor, fails the structural check;
  // WITH one (the planner metadata) it passes -- the relaxation is bound to the declaration
  const text = require('../../src/adapters/docx/families/text.cjs');
  const d = [{ text: A, props: PLAIN }, { text: NEW, props: PLAIN, list: { type: 'bullet', level: 0, numId: 'n1' } }, { text: Z, props: PLAIN }];
  const final = { primary: normalizeInventory({ html: htmlOf(d), paragraphs: d.map(p => ({ text: p.text, props: p.props, format: formatOf(p), ...runPayload(p), numbering: p.list ? { present: true, level: 0, numId: 'n1' } : { present: false } })) }) };
  const [op] = planOperations(task(RW)).operations;
  assert.equal(text.wholeTaskChecks(op, final).find(c => /structural context/.test(c.name)).ok, false, 'undeclared -> strict');
  assert.equal(text.wholeTaskChecks({ ...op, successorTransforms: ['s1'] }, final).find(c => /structural context/.test(c.name)).ok, true, 'declared -> text unit');
});

test('NEG: a HEADING requirement is not relaxed by a list successor', async () => {
  const d = [{ text: H, level: 2, props: PLAIN }, { text: A, props: PLAIN }, { text: Z, props: PLAIN }];
  const { x, result } = await run(createStore(d), task(
    { intent: 'rewrite_paragraph', selector: { text: H, kind: 'heading' }, text: 'New heading' },
    { intent: 'apply_list', selector: { ref: 's0' }, type: 'bullet' },
  ));
  assert.equal(result.ok, false);
  assert.ok(failedL2(x).some(n => /s0: final text in its structural context/.test(n)), JSON.stringify(failedL2(x)));
});

test('NEG: the successor-aware format L2 still checks the formatting on the list item (lost -> FAIL)', () => {
  const formatting = require('../../src/adapters/docx/families/formatting.cjs');
  const mk = bold => { const d = [{ text: A, props: PLAIN }, { text: B, props: { ...PLAIN, bold }, list: { type: 'numbered', level: 0, numId: 'n1' } }]; return { primary: normalizeInventory({ html: htmlOf(d), paragraphs: d.map(p => ({ text: p.text, props: p.props, format: formatOf(p), ...runPayload(p), numbering: p.list ? { present: true, level: 0, numId: 'n1' } : { present: false } })) }) }; };
  const [op] = planOperations(task({ intent: 'format_text', selector: { text: B }, format: { bold: true } }, { intent: 'apply_list', selector: { ref: 's0' }, type: 'numbered' })).operations;
  const holds = f => formatting.wholeTaskChecks(op, f).find(c => /formatting holds/.test(c.name)).ok;
  assert.equal(holds(mk(true)), true);
  assert.equal(holds(mk(false)), false, 'bold lost on the list item');
  const [range] = planOperations(task({ intent: 'format_text', selector: { text: B }, matchText: 'Beta', format: { bold: true } }, { intent: 'apply_list', selector: { ref: 's0' }, type: 'numbered' })).operations;
  assert.equal(formatting.wholeTaskChecks(range, mk(false)).find(c => /formatting holds/.test(c.name)).ok, false, 'range formatting lost');
});
