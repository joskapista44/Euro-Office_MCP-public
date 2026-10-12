'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { docxAdapter, planOperations } = require('../../src/adapters/docx/adapter.cjs');
const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');

/**
 * M5 constraints + selector.precondition + sectionHeading/paragraphOrdinal (owner, owner channel
 * 28855; source f8c6d6c semantics, constraints.cjs). Every rule: refused BEFORE dispatch where the
 * target is known (0 dispatch / 0 write), checked after EVERY write on the fresh before/after pair,
 * and a satisfied retry after reopening is 0 write / 0 command.
 */

const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const H1 = 'Hitelbiztosítékok';
const H2 = 'Védett melléklet';
const P1 = 'Első bekezdés.';
const P2 = 'Második bekezdés.';
const SUB = 'Alfejezet';
const P3 = 'Harmadik bekezdés.';
const OUT = 'Kívül eső bekezdés.';
const cell = text => ({ span: 1, paras: [{ text, props: { ...PLAIN } }] });
// H1 > P1, P2, (H2-level sub) SUB > P3 ; then a same-level heading H2 > OUT ; a table after the body
const paragraphs = () => [
  { text: H1, level: 1, props: { ...PLAIN, bold: true } }, { text: P1, props: { ...PLAIN } }, { text: P2, props: { ...PLAIN } },
  { text: SUB, level: 2, props: { ...PLAIN, bold: true } }, { text: P3, props: { ...PLAIN } },
  { text: H2, level: 1, props: { ...PLAIN, bold: true } }, { text: OUT, props: { ...PLAIN } },
];
const store = () => createStore(paragraphs(), { doc: { tables: [{ rows: [[cell('T00'), cell('T01')]] }] } });

async function run(s, task, script = {}) {
  let t = 0;
  const driver = createFakeDocxDriver(s, script);
  const result = await withSession({ driver, fileId: 'd', now: () => (t += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } }, docxAdapter, createExecutor(task));
  return { result, driver, m: result.metrics, x: result.execution };
}
const ALL = { section: H1, doNotTouchTables: true, preserveFormattingExceptRequested: true };
const snapshot = s => JSON.stringify([s.durable.map(p => [p.text, p.props]), s.durableDoc]);
const rw = (selector, text) => ({ intent: 'rewrite_paragraph', selector, text });
const stopOf = x => { const e = x.appliedOperations[0].evidence; return e && (e.ambiguity || (e.evidence && e.evidence.reason)); };

async function appliedThenSatisfied(s, task) {
  const first = await run(s, task);
  assert.equal(first.result.editOk, true, JSON.stringify(first.x.appliedOperations.map(e => [e.status, e.evidence])));
  assert.ok(first.x.appliedOperations.every(e => e.status === OPERATION_STATUS.APPLIED));
  assert.ok(first.x.preservationVerification.length >= task.steps.length && first.x.preservationVerification.every(v => v.verdict === 'PASS'), JSON.stringify(first.x.preservationVerification));
  assert.equal(first.x.wholeTaskVerification.verdict, 'PASS');
  const second = await run(s, task);
  assert.equal(second.result.ok, true, JSON.stringify(second.x.appliedOperations.map(e => [e.status, e.evidence])));
  assert.ok(second.x.appliedOperations.every(e => e.status === OPERATION_STATUS.SATISFIED));
  assert.deepEqual([second.m.writes, second.m.commands, second.driver.calls.dispatch], [0, 0, 0], 'retry after reopening: 0 write / 0 command');
  return first;
}

async function refusedBeforeDispatch(s, task, want) {
  const before = snapshot(s);
  const { x, driver, m } = await run(s, task);
  assert.equal(x.appliedOperations[0].status, OPERATION_STATUS.FAILED, want);
  assert.equal(stopOf(x), want);
  assert.deepEqual([driver.calls.dispatch, m.writes], [0, 0], `${want}: 0 dispatch / 0 write`);
  assert.equal(snapshot(s), before, `${want}: nothing written`);
}

async function failedAfterWrite(s, task, script, constraint) {
  const { result, x } = await run(s, task, script);
  assert.equal(result.ok, false, `${constraint}: the task fails`);
  const l3 = x.preservationVerification.find(v => v.verdict !== 'PASS');
  assert.ok(l3, `${constraint}: an L3 failed`);
  const ev = l3.evidence && (l3.evidence.constraints || l3.evidence.evidence || l3.evidence);
  assert.equal(JSON.stringify(ev).includes(`"${constraint}":{"ok":false`), true, `${constraint} is the failing check: ${JSON.stringify(l3.evidence)}`);
  assert.notEqual(x.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
}

// ---- positive: every field migrated, first write + reopened retry ------------------------------------

test('all three constraints + precondition, text selector inside the section: APPLIED; retry 0/0', async () => {
  await appliedThenSatisfied(store(), { id: 't', constraints: ALL, steps: [rw({ kind: 'paragraph', text: P1, precondition: { text: P1 } }, 'Első bekezdés, új.')] });
});

test('sectionHeading + paragraphOrdinal (1-based, paragraphs only, nested subsection counted): ordinal 3 = P3 under the subheading', async () => {
  const s = store();
  await appliedThenSatisfied(s, { id: 't', constraints: ALL, steps: [rw({ sectionHeading: H1, paragraphOrdinal: 3, precondition: { text: P3 } }, 'Harmadik, új.')] });
  assert.deepEqual(s.durable.map(p => p.text), [H1, P1, P2, SUB, 'Harmadik, új.', H2, OUT]);
});

test('format_text under preserveFormattingExceptRequested: the requested key changes, nothing else; retry 0/0', async () => {
  const s = store();
  await appliedThenSatisfied(s, { id: 't', constraints: ALL, steps: [{ intent: 'format_text', selector: { sectionHeading: H1, paragraphOrdinal: 2 }, format: { bold: true } }] });
  assert.equal(s.durable.find(p => p.text === P2).props.bold, true);
  assert.equal(s.durable.find(p => p.text === P1).props.bold, false);
});

test('the constraints hold across a whole multi-step task (every write checked, not only the end)', async () => {
  await appliedThenSatisfied(store(), { id: 't', constraints: ALL, steps: [
    rw({ kind: 'paragraph', text: P1 }, 'Egy.'),
    { intent: 'insert_paragraph_after', selector: { text: 'Egy.' }, text: 'Beszúrt.' },
    { intent: 'format_text', selector: { text: P2 }, format: { italic: true } },
  ] });
});

// ---- negative: refused before the first dispatch ------------------------------------------------------

test('section: a target outside, the heading itself, a missing / duplicated heading -> refused, 0 dispatch', async () => {
  await refusedBeforeDispatch(store(), { id: 't', constraints: { section: H1 }, steps: [rw({ text: OUT }, 'x')] }, 'TARGET_OUTSIDE_CONSTRAINED_SECTION');
  await refusedBeforeDispatch(store(), { id: 't', constraints: { section: H1 }, steps: [rw({ text: H1 }, 'x')] }, 'TARGET_OUTSIDE_CONSTRAINED_SECTION');
  await refusedBeforeDispatch(store(), { id: 't', constraints: { section: 'Nincs ilyen' }, steps: [rw({ text: P1 }, 'x')] }, 'CONSTRAINT_SECTION_NOT_FOUND');
  const dup = createStore([...paragraphs(), { text: H1, level: 1, props: { ...PLAIN } }]);
  await refusedBeforeDispatch(dup, { id: 't', constraints: { section: H1 }, steps: [rw({ text: P1 }, 'x')] }, 'CONSTRAINT_SECTION_AMBIGUOUS');
  // a move whose DESTINATION is outside
  await refusedBeforeDispatch(store(), { id: 't', constraints: { section: H1 }, steps: [{ intent: 'move_paragraph_after', selector: { text: P1 }, destinationSelector: { text: OUT } }] }, 'TARGET_OUTSIDE_CONSTRAINED_SECTION');
});

test('selector.precondition mismatch -> refused before dispatch (was: silently APPLIED before M5)', async () => {
  await refusedBeforeDispatch(store(), { id: 't', steps: [rw({ text: P1, precondition: { text: 'Valami más.' } }, 'x')] }, 'PRECONDITION_TEXT_MISMATCH');
});

test('sectionHeading + paragraphOrdinal: out of range, 0, a string, an unknown heading -> refused, 0 dispatch', async () => {
  await refusedBeforeDispatch(store(), { id: 't', steps: [rw({ sectionHeading: H1, paragraphOrdinal: 9 }, 'x')] }, 'rewrite-paragraph-not-found');
  await refusedBeforeDispatch(store(), { id: 't', steps: [rw({ sectionHeading: H1, paragraphOrdinal: 0 }, 'x')] }, 'rewrite-paragraph-ordinal-invalid');
  await refusedBeforeDispatch(store(), { id: 't', steps: [rw({ sectionHeading: H1, paragraphOrdinal: '2' }, 'x')] }, 'rewrite-paragraph-ordinal-invalid');
  await refusedBeforeDispatch(store(), { id: 't', steps: [rw({ sectionHeading: 'Nincs', paragraphOrdinal: 1 }, 'x')] }, 'rewrite-section-not-found');
});

test('planning refusals: shape, unknown keys, table intents, header/footer under a section, ordinal after a count change', () => {
  const detail = t => { const p = planOperations(t); assert.equal(p.ok, false, JSON.stringify(t)); return p.detail; };
  const S = [rw({ text: P1 }, 'x')];
  assert.match(detail({ constraints: [], steps: S }), /task-constraints-invalid/);
  assert.match(detail({ constraints: { zzq: 1 }, steps: S }), /task-constraints-unsupported: zzq/);
  assert.match(detail({ constraints: { section: '  ' }, steps: S }), /task-section-constraint-invalid/);
  assert.match(detail({ constraints: { doNotTouchTables: 'yes' }, steps: S }), /task-table-constraint-invalid/);
  assert.match(detail({ constraints: { preserveFormattingExceptRequested: 1 }, steps: S }), /task-formatting-constraint-invalid/);
  assert.match(detail({ constraints: { doNotTouchTables: true }, steps: [{ intent: 'replace_table_cell', selector: { tableCell: { text: 'T00' } }, text: 'y' }] }), /task-table-write-forbidden-by-constraint/);
  assert.match(detail({ constraints: { section: H1 }, steps: [{ intent: 'edit_header', section: { index: 0, of: 1 }, text: 'h' }] }), /constraint-section-not-enforceable/);
  assert.match(detail({ steps: [rw({ text: P1, precondition: 'x' }, 'y')] }), /precondition-invalid/);
  assert.match(detail({ steps: [rw({ text: P1, precondition: { text: 'a', contains: true } }, 'y')] }), /precondition-unsupported: contains/);
  assert.match(detail({ steps: [rw({ text: P1, precondition: {} }, 'y')] }), /precondition-text-required/);
  assert.match(detail({ steps: [{ intent: 'insert_paragraph_after', selector: { text: P1 }, text: 'n' }, rw({ sectionHeading: H1, paragraphOrdinal: 2 }, 'y')] }), /ordinal-selector-with-count-changing-step/);
  // AFTER the ordinal step too: on a reopened retry the later step has already changed the count
  assert.match(detail({ steps: [rw({ sectionHeading: H1, paragraphOrdinal: 2 }, 'y'), { intent: 'delete_paragraph', selector: { text: P1 } }] }), /ordinal-selector-with-count-changing-step/);
  assert.match(detail({ steps: [rw({ sectionHeading: H1, paragraphOrdinal: 2 }, '')] }), /ordinal-selector-on-count-changing-step/);
  // not migrated on these intents / fields: still refused, named (never dropped)
  assert.match(detail({ steps: [{ intent: 'delete_paragraph', selector: { sectionHeading: H1, paragraphOrdinal: 1 } }] }), /field-not-enforced-yet: .*selector\.sectionHeading/);
  assert.match(detail({ steps: [{ intent: 'move_paragraph_after', selector: { text: P1 }, destinationSelector: { text: P2, precondition: { text: P2 } } }] }), /field-not-enforced-yet: .*selector\.precondition/);
  assert.match(detail({ steps: [rw({ sectionHeading: H1, paragraphOrdinal: 1, text: P1 }, 'y')] }), /text would be ignored beside the sectionHeading selector/);
});

// ---- negative: a write with an out-of-scope side effect -> the post-write L3 fails ---------------------

const sideText = text => ({ sideEffect: ({ live }) => { const p = live.find(x => x.text === text); if (p) p.text = `${text} (módosítva)`; } });
const sideBold = text => ({ sideEffect: ({ live }) => { const p = live.find(x => x.text === text); if (p) p.props = { ...p.props, bold: true }; } });
const sideCell = { sideEffect: ({ liveDoc }) => { liveDoc.tables[0].rows[0][0].paras[0].text = 'T00 (módosítva)'; } };

// ---- M5 gate 3: insert_paragraph_before / _after with sectionHeading + paragraphOrdinal ----------------

const ins = (intent, ordinal, text, pre) => ({ intent, selector: { sectionHeading: H1, paragraphOrdinal: ordinal, ...(pre ? { precondition: { text: pre } } : {}) }, text });
const texts = s => s.durable.map(p => p.text);
// an insert is P0: L3 exists only under constraints -- the helper asserts L1 + L2 + the 0/0 retry
async function insertAppliedThenSatisfied(st, task) {
  const first = await run(st, task);
  assert.equal(first.result.editOk, true, JSON.stringify(first.x.appliedOperations.map(e => [e.status, e.evidence])));
  assert.ok(first.x.appliedOperations.every(e => e.status === OPERATION_STATUS.APPLIED));
  assert.ok(first.x.primitiveVerification.every(v => v.verdict === 'PASS'), JSON.stringify(first.x.primitiveVerification));
  assert.ok(!first.x.preservationVerification || first.x.preservationVerification.every(v => v.verdict === 'PASS'));
  assert.equal(first.x.wholeTaskVerification.verdict, 'PASS', JSON.stringify(first.x.wholeTaskVerification));
  const second = await run(st, task);
  assert.ok(second.x.appliedOperations.every(e => e.status === OPERATION_STATUS.SATISFIED), JSON.stringify(second.x.appliedOperations.map(e => [e.status, e.evidence])));
  assert.deepEqual([second.m.writes, second.m.commands, second.driver.calls.dispatch], [0, 0, 0], 'retry after reopening: 0 write / 0 command');
  assert.equal(second.x.wholeTaskVerification.verdict, 'PASS');
}

test('gate 3: ordinal insert BEFORE / AFTER lands at the terminal ordinal; retry after reopening is SATISFIED 0/0 (the anchor moved, the proof is positional)', async () => {
  const b = store();
  await insertAppliedThenSatisfied(b, { id: 't', steps: [ins('insert_paragraph_before', 2, 'Új előtte.')] });
  assert.deepEqual(texts(b).slice(0, 4), [H1, P1, 'Új előtte.', P2], 'before ordinal 2: the new paragraph IS ordinal 2, P2 moved to 3');
  const a = store();
  await insertAppliedThenSatisfied(a, { id: 't', steps: [ins('insert_paragraph_after', 2, 'Új utána.', P2)] });
  assert.deepEqual(texts(a).slice(0, 5), [H1, P1, P2, 'Új utána.', SUB], 'after ordinal 2 (P2, precondition held): ordinal 3');
  // under all three constraints too (the insert is inside the constrained section)
  await insertAppliedThenSatisfied(store(), { id: 't', constraints: ALL, steps: [ins('insert_paragraph_before', 1, 'Új első.', P1)] });
});

test('gate 3 NEGATIVE: precondition mismatch / out of range -> refused before dispatch; the text already elsewhere -> never SATISFIED', async () => {
  await refusedBeforeDispatch(store(), { id: 't', steps: [ins('insert_paragraph_before', 2, 'Új.', 'Nem ez.')] }, 'PRECONDITION_TEXT_MISMATCH');
  await refusedBeforeDispatch(store(), { id: 't', steps: [ins('insert_paragraph_after', 9, 'Új.')] }, 'insert-anchor-paragraph_not_found');
  // the produced text exists, but NOT at the terminal ordinal -> fail closed, no write
  await refusedBeforeDispatch(store(), { id: 't', steps: [ins('insert_paragraph_before', 1, OUT)] }, 'insert-produced-text-in-wrong-position');
  // the terminal position holds the text but the anchor's precondition does not -> not SATISFIED
  const s = store();
  await insertAppliedThenSatisfied(s, { id: 't', steps: [ins('insert_paragraph_after', 1, 'Közbe.', P1)] });
  await refusedBeforeDispatch(s, { id: 't', steps: [ins('insert_paragraph_after', 1, 'Közbe.', P2)] }, 'insert-anchor-precondition-mismatch-at-terminal');
});

test('gate 3 planning: an ordinal insert only as the task\'s single ordinal AND single count-changing step', () => {
  const detail = t => { const r = planOperations({ id: 't', ...t }); return r.ok ? 'OK' : r.detail; };
  assert.equal(detail({ steps: [ins('insert_paragraph_before', 2, 'x')] }), 'OK');
  assert.match(detail({ steps: [ins('insert_paragraph_before', 2, 'x'), rw({ sectionHeading: H1, paragraphOrdinal: 1 }, 'y')] }), /ordinal-insert-with-another-ordinal-or-count-changing-step/);
  assert.match(detail({ steps: [ins('insert_paragraph_before', 2, 'x'), { intent: 'delete_paragraph', selector: { text: OUT } }] }), /ordinal-insert-with-another-ordinal-or-count-changing-step/);
  assert.match(detail({ steps: [rw({ sectionHeading: H1, paragraphOrdinal: 1 }, 'y'), ins('insert_paragraph_after', 2, 'x')] }), /ordinal-selector-with-count-changing-step/);
});

test('post-write: outside the section changed -> FAIL; CONTROL without the constraint the same write passes', async () => {
  await failedAfterWrite(store(), { id: 't', constraints: { section: H1 }, steps: [rw({ text: P1 }, 'Egy.')] }, sideText(OUT), 'outsideSectionUnchanged');
  const control = await run(store(), { id: 't', steps: [rw({ text: P1 }, 'Egy.')] }, sideText(OUT));
  assert.equal(control.result.editOk, true, 'the family checks alone do not see the outside change -- the constraint is what catches it');
});

test('post-write: a table changed -> FAIL under doNotTouchTables; CONTROL without it passes', async () => {
  await failedAfterWrite(store(), { id: 't', constraints: { doNotTouchTables: true }, steps: [rw({ text: P1 }, 'Egy.')] }, sideCell, 'tablesUnchanged');
  const control = await run(store(), { id: 't', steps: [rw({ text: P1 }, 'Egy.')] }, sideCell);
  assert.equal(control.result.editOk, true);
});

test('post-write: another paragraph formatted -> FAIL under preserveFormattingExceptRequested; CONTROL without it passes', async () => {
  await failedAfterWrite(store(), { id: 't', constraints: { preserveFormattingExceptRequested: true }, steps: [rw({ text: P1 }, 'Egy.')] }, sideBold(P2), 'formattingUnchangedExceptRequested');
  const control = await run(store(), { id: 't', steps: [rw({ text: P1 }, 'Egy.')] }, sideBold(P2));
  assert.equal(control.result.editOk, true);
  // a paragraph-structure step (no preservation contract of its own) is checked too
  await failedAfterWrite(store(), { id: 't', constraints: { section: H1 }, steps: [{ intent: 'insert_paragraph_after', selector: { text: P1 }, text: 'Új.' }] }, sideText(OUT), 'outsideSectionUnchanged');
});
