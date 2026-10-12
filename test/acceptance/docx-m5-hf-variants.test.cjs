'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { PERSISTENCE_STATE } = require('../../src/core/persistence/verified-writes.cjs');
const { docxAdapter } = require('../../src/adapters/docx/adapter.cjs');
const hf = require('../../src/adapters/docx/families/header-footer.cjs');
const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');

/**
 * M5 first/even (owner decision after owner channel: option A). A title / even write turns
 * the variant's visibility flag on ONLY if every unrequested slot that the flag switches is provably
 * unchanged -- title: the OTHER kind of the same section; even: BOTH kinds of EVERY section except
 * the target. Unknown, inherited or unobservable -> refused before dispatch, 0 dispatch / 0 write.
 * An already active flag has no activation side effect; targeting and preservation still apply.
 * The fake models the target probe 2 facts: the title flag per section, the even/odd flag
 * document-wide, no slot created by a flag change.
 */

const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const part = (...texts) => ({ paras: texts.map(t => ({ text: t, props: { ...PLAIN }, alignment: 'left' })) });
const sec = (h = {}, f = {}, titlePage = false) => ({ header: { default: null, even: null, title: null, ...h }, footer: { default: null, even: null, title: null, ...f }, titlePage });
const paragraphs = () => [{ text: 'Body heading', level: 1, props: { ...PLAIN, bold: true } }, { text: 'Body first.', props: { ...PLAIN } }];
const store = (sections, evenAndOdd = false) => createStore(paragraphs(), { doc: { sections, evenAndOdd } });

async function run(s, task, script = {}) {
  let t = 0;
  const driver = createFakeDocxDriver(s, script);
  const result = await withSession({ driver, fileId: 'd', now: () => (t += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } }, docxAdapter, createExecutor(task));
  return { result, driver, m: result.metrics, x: result.execution };
}
const task = (...steps) => ({ id: 'm5hf', steps });
const eh = (text, variant, index = 0, of = 1) => ({ intent: 'edit_header', section: { index, of }, text, variant });
const ef = (text, variant, index = 0, of = 1) => ({ intent: 'edit_footer', section: { index, of }, text, variant });
const partOf = (s, i, k, v) => { const p = s.durableDoc.sections[i][k][v]; return p ? p.paras.map(x => x.text) : null; };
const reason = x => { const e = x.appliedOperations[0].evidence; return (e && e.evidence && e.evidence.reason) || (e && e.reason); };
const snapshot = s => JSON.stringify([s.durable.map(p => p.text), s.durableDoc]);

async function firstAndSecond(s, t) {
  const first = await run(s, t);
  assert.equal(first.result.editOk, true, JSON.stringify(first.x.appliedOperations.map(e => [e.status, e.evidence])));
  assert.equal(first.x.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
  assert.deepEqual([first.m.writes, first.m.commands], [1, 1]);
  assert.ok(first.x.preservationVerification.every(v => v.verdict === 'PASS'), JSON.stringify(first.x.preservationVerification));
  assert.equal(first.x.wholeTaskVerification.verdict, 'PASS');
  assert.equal(first.result.persistenceState, PERSISTENCE_STATE.EDITOR_CLEAN);
  const second = await run(s, t);
  assert.equal(second.x.appliedOperations[0].status, OPERATION_STATUS.SATISFIED, JSON.stringify(second.x.appliedOperations.map(e => e.evidence)));
  assert.deepEqual([second.m.writes, second.m.commands, second.driver.calls.forceSave], [0, 0, 0]);
  return first;
}

async function refused(s, t, want) {
  const before = snapshot(s);
  const { x, driver, m } = await run(s, t);
  assert.equal(x.appliedOperations[0].status, OPERATION_STATUS.FAILED, want);
  assert.equal(reason(x), want);
  assert.deepEqual([driver.calls.dispatch, m.writes], [0, 0], `${want}: 0 dispatch / 0 write`);
  assert.equal(snapshot(s), before, `${want}: nothing written`);
}

// ---- positive ------------------------------------------------------------------------------------

test('title header and title footer, one section, nothing to switch -> part + title flag on, nothing else; SECOND 0/0', async () => {
  for (const [st, k] of [[eh('T HEADER', 'title'), 'header'], [ef('T FOOTER', 'title'), 'footer']]) {
    const s = store([sec()]);
    await firstAndSecond(s, task(st));
    assert.deepEqual(partOf(s, 0, k, 'title'), [st.text]);
    assert.equal(s.durableDoc.sections[0].titlePage, true);
    assert.equal(s.durableDoc.evenAndOdd, false, 'the even flag untouched');
    for (const [kk, v] of [['header', 'default'], ['footer', 'default'], [k === 'header' ? 'footer' : 'header', 'title'], ['header', 'even'], ['footer', 'even']]) assert.equal(partOf(s, 0, kk, v), null, `no ${kk}/${v} slot created`);
  }
});

test('title: the other kind has default AND title parts with identical content -> provably unchanged -> activated', async () => {
  const s = store([sec({}, { default: part('SAME'), title: part('SAME') })]);
  await firstAndSecond(s, task(eh('T HEADER', 'title')));
  assert.equal(s.durableDoc.sections[0].titlePage, true);
  assert.deepEqual(partOf(s, 0, 'footer', 'title'), ['SAME']);
});

test('even header and even footer, one section, nothing to switch -> part + document even flag on; SECOND 0/0', async () => {
  for (const [st, k] of [[eh('E HEADER', 'even'), 'header'], [ef('E FOOTER', 'even'), 'footer']]) {
    const s = store([sec()]);
    await firstAndSecond(s, task(st));
    assert.deepEqual(partOf(s, 0, k, 'even'), [st.text]);
    assert.equal(s.durableDoc.evenAndOdd, true);
    assert.equal(s.durableDoc.sections[0].titlePage, false, 'the title flag untouched');
  }
});

test('even footer on section 0 of 2: every OTHER slot of EVERY section provably unchanged (own default == own even) -> activated', async () => {
  const s = store([
    sec({ default: part('HD'), even: part('HD') }, {}),
    sec({ default: part('HD'), even: part('HD') }, { default: part('F1'), even: part('F1') }),
  ]);
  await firstAndSecond(s, task(ef('E0 FOOTER', 'even', 0, 2)));
  assert.equal(s.durableDoc.evenAndOdd, true);
  assert.deepEqual(partOf(s, 0, 'footer', 'even'), ['E0 FOOTER']);
  assert.deepEqual(partOf(s, 1, 'footer', 'even'), ['F1'], 'section 1 unchanged');
});

test('already active flag: no activation side effect -> a plain write even with a default footer present', async () => {
  const s = store([sec({}, { default: part('F') }, true)]);
  await firstAndSecond(s, task(eh('T HEADER', 'title')));
  assert.equal(s.durableDoc.sections[0].titlePage, true);
  assert.deepEqual(partOf(s, 0, 'footer', 'default'), ['F']);
  const e = store([sec({ default: part('H') }, { default: part('F') })], true);
  await firstAndSecond(e, task(ef('E FOOTER', 'even')));
  assert.equal(e.durableDoc.evenAndOdd, true);
});

// ---- negative (0 dispatch / 0 write) ---------------------------------------------------------------

test('title activation would change the first page of the OTHER kind -> refused', async () => {
  await refused(store([sec({}, { default: part('F') })]), task(eh('T', 'title')), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:title-activation-changes-unrequested-section-0-footer');
  await refused(store([sec({}, { default: part('F'), title: part('G') })]), task(eh('T', 'title')), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:title-activation-changes-unrequested-section-0-footer');
  await refused(store([sec({ default: part('H') }, {})]), task(ef('T', 'title')), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:title-activation-changes-unrequested-section-0-header');
});

test('EVEN: ANOTHER section\'s unrequested footer would change -> refused (the document-wide flag)', async () => {
  const s = store([
    sec({ default: part('HD'), even: part('HD') }, {}),
    sec({ default: part('HD'), even: part('HD') }, { default: part('F1') }),
  ]);
  await refused(s, task(eh('E0', 'even', 0, 2)), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:even-activation-changes-unrequested-section-1-footer');
});

test('EVEN: the same kind of the target in another section is checked too (not only the other kind)', async () => {
  // target: section 1 footer; section 0 has a default footer and no even footer -> its even pages would change
  const s = store([sec({}, { default: part('F0') }), sec({}, { default: part('F1'), even: part('F1') })]);
  await refused(s, task(ef('E1', 'even', 1, 2)), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:even-activation-changes-unrequested-section-0-footer');
});

test('INHERITED slot (what it shows is not observable) -> refused, never guessed', async () => {
  // section 1 has no footer of its own; section 0 has default and even footers -> section 1 inherits
  const s = store([
    sec({ default: part('HD'), even: part('HD') }, { default: part('F0'), even: part('F0') }),
    sec({ default: part('HD'), even: part('HD') }, {}),
  ]);
  await refused(s, task(eh('E0', 'even', 0, 2)), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:even-activation-inherited-section-1-footer');
});

test('an even flag reported differently by two sections -> refused as inconsistent (unit, synthetic observation)', () => {
  const s = { sections: [{ header: { default: null, even: null, title: null }, footer: { default: null, even: null, title: null }, titlePage: false, evenAndOdd: false },
    { header: { default: null, even: null, title: null }, footer: { default: null, even: null, title: null }, titlePage: false, evenAndOdd: true }] };
  const op = { intent: 'edit_header', part: 'header', variant: 'even', sectionIndex: 0, sectionCount: 2, text: 'x' };
  const st = hf.partState({ document: s }, op);
  assert.equal(st.flag, null);
  assert.equal(hf.shapeGuard(st, op), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:even-flag-inconsistent-across-sections');
});

test('the flag not set by the command -> L1 fails, never APPLIED, no barrier', async () => {
  const s = store([sec()]);
  const { x } = await run(s, task(eh('T', 'title')), { fault: 'hf-flag-not-set' });
  assert.notEqual(x.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
  assert.equal(s.durableDoc.sections[0].titlePage, false);
});

test('another section\'s title flag changed by the write -> L3 fails (default and title writes alike)', async () => {
  for (const st of [eh('D', 'default', 0, 2), eh('T', 'title', 0, 2)]) {
    const s = store([sec({ default: part('H0'), title: part('T0') }, {}), sec({ default: part('H1'), title: part('T1') }, {})]);
    const { x } = await run(s, task(st), { fault: 'hf-other-title-flag' });
    assert.notEqual(x.appliedOperations[0].status, OPERATION_STATUS.APPLIED, st.variant);
    const l3 = x.preservationVerification[0];
    assert.equal(l3.verdict, 'FAIL', st.variant);
    assert.deepEqual(l3.evidence.evidence.unrequestedFlagChanges, ['1:titlePage'], st.variant);
  }
});
