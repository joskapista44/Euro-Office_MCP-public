'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { PERSISTENCE_STATE } = require('../../src/core/persistence/verified-writes.cjs');
const { docxAdapter, planOperations } = require('../../src/adapters/docx/adapter.cjs');
const ps = require('../../src/adapters/docx/families/page-setup.cjs');
const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');

/**
 * W11 offline acceptance: set_page_setup through the TARGET path on the fake transport, modeled
 * directly on test/acceptance/docx-w5-header-footer.test.cjs's own shape. The fake model follows
 * the LIVE probes (W11-1a reflection #ref, W11-1b write #ref, the two follow-ups #ref/
 * #ref): per section, size/margins/columns/startPageNumber, with the measured call orders
 * (size: width,height two-arg; margins: left,top,right,bottom -- NOT the task's top,right,bottom,
 * left field order).
 */

const paragraphs = () => [{ text: 'Body heading', level: 1, props: {} }, { text: 'Body first.', props: {} }];
// flat fields -- matching how cloneDoc() reads them in BOTH its contexts (fixture->live AND
// live->durable on save); a nested shape would be silently discarded on the first forceSave.
const sectionFixture = (fields = {}) => ({ ...fields });
const store = (sections, paras = paragraphs()) => createStore(paras, { doc: sections ? { sections } : {} });

async function run(s, task, script = {}) {
  let t = 0;
  const driver = createFakeDocxDriver(s, script);
  const options = { driver, fileId: 'd', now: () => (t += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, docxAdapter, createExecutor(task));
  if (result.execution.operationAccounting) assert.equal(result.execution.operationAccounting.balanced, true);
  return { result, driver, m: result.metrics, x: result.execution };
}
const task = (...steps) => ({ id: 'w11', steps });
const sps = (fields, index = 0, of = 1, extra = {}) => ({ intent: 'set_page_setup', section: { index, of }, ...fields, ...extra });
const status = (x, i = 0) => x.appliedOperations[i].status;
const reason = (x, i = 0) => { const e = x.appliedOperations[i].evidence; return (e && e.evidence && e.evidence.reason) || (e && e.reason); };
const sectionOf = (s, i) => s.durableDoc.sections[i];
const snapshot = s => JSON.stringify([s.durable.map(p => p.text), s.durableDoc.sections]);

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

test('startPageNumber: write + retry SATISFIED 0/0; every other field of the section unchanged', async () => {
  const s = store([sectionFixture()]);
  await firstAndSecond(s, task(sps({ startPageNumber: 5 })));
  assert.equal(sectionOf(s, 0).startPageNumber, 5);
  assert.deepEqual([sectionOf(s, 0).pageWidth, sectionOf(s, 0).pageHeight, sectionOf(s, 0).margins, sectionOf(s, 0).columnsCount], [12240, 15840, { top: 1440, right: 1440, bottom: 1440, left: 1440 }, 1]);
});

test('columns: count+space write; BOTH axes re-readable (GetColumnsCount AND GetColumnsSpaces, #ref)', async () => {
  const s = store([sectionFixture()]);
  await firstAndSecond(s, task(sps({ columns: { count: 2, space: 720 } })));
  assert.equal(sectionOf(s, 0).columnsCount, 2);
  assert.deepEqual(sectionOf(s, 0).columnsSpaces, [720]);
});

test('COLUMNS_SPACE_FALSE_SATISFIED: same count, DIFFERENT space -> NOT SATISFIED, a real write is dispatched', async () => {
  const s = store([sectionFixture({ columnsCount: 2, columnsSpace: 720 })]); // already count=2, space=720
  const before = snapshot(s);
  const { x, m, driver } = await run(s, task(sps({ columns: { count: 2, space: 500 } }))); // SAME count, DIFFERENT space
  assert.notEqual(before, snapshot(s), 'the space difference must cause a real write, not a false SATISFIED');
  assert.equal(status(x), OPERATION_STATUS.APPLIED);
  assert.equal(m.writes, 1);
  assert.equal(driver.calls.dispatch, 1);
  assert.deepEqual(sectionOf(s, 0).columnsSpaces, [500]);
});

test('faults: columns space corrupted on write -> L1 FAIL (the exact bug class #ref caught)', async () => {
  const { x } = await run(store([sectionFixture()]), task(sps({ columns: { count: 2, space: 720 } })), { fault: 'ps-columns-space-wrong' });
  assert.equal(x.primitiveVerification[0].verdict, 'FAIL', 'a wrong space must not read back as correct');
});

test('columns space specifically unobservable (count readable, space not) -> FAIL_CLOSED before dispatch', async () => {
  const { x, driver } = await run(store([sectionFixture()]), task(sps({ columns: { count: 2, space: 720 } })), { columnsSpacesUnreadable: true });
  assert.equal(status(x), OPERATION_STATUS.FAILED);
  assert.equal(reason(x), 'FAIL_CLOSED_UNSUPPORTED_SHAPE:page-setup-field-unobservable:columns');
  assert.equal(driver.calls.dispatch, 0);
});

test('size: two-arg SetPageSize call, portrait only -- landscape (width > height) refused at PLANNING, 0 dispatch', async () => {
  const s = store([sectionFixture()]); // fixture default is 12240x15840 -- use a DIFFERENT target so the first round is a real APPLY, not an accidental SATISFIED
  await firstAndSecond(s, task(sps({ size: { width: 11000, height: 14000 } })));
  assert.deepEqual([sectionOf(s, 0).pageWidth, sectionOf(s, 0).pageHeight], [11000, 14000]);
  const before = snapshot(s);
  const { x, driver } = await run(s, task(sps({ size: { width: 15840, height: 12240 } })));
  assert.equal(x.outcome, 'plan-failed', JSON.stringify(x));
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(snapshot(s), before, 'nothing written on a plan-time refusal');
});

test('MARGINS_ORDER: four DISTINCT values land in the measured (left,top,right,bottom) call order, not the task field order -- this is the exact bug an identical-value fixture could not have caught', async () => {
  const s = store([sectionFixture()]);
  await firstAndSecond(s, task(sps({ margins: { top: 1102, right: 1203, bottom: 1304, left: 1001 } })));
  assert.deepEqual(sectionOf(s, 0).margins, { top: 1102, right: 1203, bottom: 1304, left: 1001 }, 'the TASK field names, read back correctly despite the reordered call');
  // NEGATIVE CONTROL: the fake driver's own fault that swaps the call positions must be CAUGHT by
  // L1 (primitiveVerification FAILs) -- proving the test is sensitive to the exact bug class, not
  // just to "some write happened".
  const s2 = store([sectionFixture()]);
  const { x } = await run(s2, task(sps({ margins: { top: 1102, right: 1203, bottom: 1304, left: 1001 } })), { fault: 'ps-margins-swapped' });
  assert.equal(x.primitiveVerification[0].verdict, 'FAIL', 'a swapped-order write must not read back as correct');
});

test('ALL FOUR FIELDS in one step: every axis changes exactly once, no cross-axis bleed', async () => {
  const s = store([sectionFixture()]);
  await firstAndSecond(s, task(sps({ size: { width: 12240, height: 15840 }, margins: { top: 1000, right: 1100, bottom: 1200, left: 1300 }, columns: { count: 3, space: 500 }, startPageNumber: 7 })));
  const sec = sectionOf(s, 0);
  assert.deepEqual([sec.pageWidth, sec.pageHeight, sec.margins, sec.columnsCount, sec.startPageNumber], [12240, 15840, { top: 1000, right: 1100, bottom: 1200, left: 1300 }, 3, 7]);
});

test('SECTION_TARGETING: two sections, each with a distinct page setup -- writing section 1 leaves section 0 untouched', async () => {
  const sections = () => [sectionFixture({ startPageNumber: 1 }), sectionFixture({ startPageNumber: 1 })];
  const s = store(sections());
  await firstAndSecond(s, task(sps({ startPageNumber: 9 }, 1, 2)));
  assert.deepEqual([sectionOf(s, 0).startPageNumber, sectionOf(s, 1).startPageNumber], [1, 9]);
});

test('negatives before dispatch -> FAIL_CLOSED or plan-failed, 0 dispatch, nothing written', async () => {
  const failed = [
    [task(sps({ startPageNumber: 5 }, 0, 2)), [sectionFixture()], 'FAIL_CLOSED:section-descriptor-mismatch'],
    [task(sps({ startPageNumber: 5 })), [sectionFixture(), sectionFixture()], 'FAIL_CLOSED:section-descriptor-mismatch'],
  ];
  for (const [t, sections, want] of failed) {
    const s = store(sections);
    const before = snapshot(s);
    const { x, driver } = await run(s, t);
    assert.equal(status(x), OPERATION_STATUS.FAILED, want);
    assert.equal(reason(x), want);
    assert.equal(driver.calls.dispatch, 0, want);
    assert.equal(snapshot(s), before, `${want}: nothing written`);
  }
  // sections unreadable -> FAIL_CLOSED, never a guess
  const u = await run(store(), task(sps({ startPageNumber: 5 })), { sectionsUnreadable: true });
  assert.equal(reason(u.x), 'sections-unreadable');
  assert.equal(u.driver.calls.dispatch, 0);
  // page-setup fields specifically unobservable -> FAIL_CLOSED (section-unreadable path), not a guessed number
  const un = await run(store([sectionFixture()]), task(sps({ startPageNumber: 5 })), { pageSetupUnreadable: true });
  assert.equal(status(un.x), OPERATION_STATUS.FAILED);
  for (const st of [
    { intent: 'set_page_setup', startPageNumber: 5 }, // no section: never implicit
    sps({}), // no field at all
    sps({ margins: { top: 1 } }), // partial margins
    sps({ columns: { count: 0, space: 0 } }), // count must be positive
    sps({ startPageNumber: 0 }),
    sps({ size: { width: 15840, height: 12240 } }), // landscape, not yet measured
    { intent: 'set_page_setup', section: { index: 0, of: 1 }, startPageNumber: 5, selector: { text: 'x' } }, // no selector at all
    { intent: 'set_page_setup', section: { index: 0, of: 1 }, startPageNumber: 5, orientation: 'landscape' }, // unknown field
  ]) {
    const { x, driver } = await run(store(), task(st));
    assert.equal(x.outcome, 'plan-failed', JSON.stringify(st));
    assert.equal(driver.calls.dispatch, 0);
  }
});

test('faults at L3 (P1): a sibling field on the SAME section changing, or another section changing -> NOT APPLIED', async () => {
  // ps-other-field-changed mutates startPageNumber as the SIBLING -- request columns here, so the
  // requested field and the sibling fault never collide (which would corrupt L1, not just L3).
  {
    const { x } = await run(store([sectionFixture()]), task(sps({ columns: { count: 2, space: 720 } })), { fault: 'ps-other-field-changed' });
    assert.equal(x.primitiveVerification[0].verdict, 'PASS', 'ps-other-field-changed: the requested field itself is right');
    assert.equal(x.preservationVerification[0].verdict, 'FAIL', 'ps-other-field-changed');
  }
  {
    const sections = [sectionFixture(), sectionFixture()];
    const { x } = await run(store(sections), task(sps({ startPageNumber: 5 }, 0, sections.length)), { fault: 'ps-other-section' });
    assert.equal(x.primitiveVerification[0].verdict, 'PASS', 'ps-other-section: the requested field itself is right');
    assert.equal(x.preservationVerification[0].verdict, 'FAIL', 'ps-other-section');
  }
});

test('faults at L3: the body changed -> NOT APPLIED', async () => {
  const { x } = await run(store([sectionFixture()]), task(sps({ startPageNumber: 5 })), { fault: 'ps-body-change' });
  assert.equal(x.primitiveVerification[0].verdict, 'PASS');
  assert.equal(x.preservationVerification[0].verdict, 'FAIL');
});

test('retry: a correct terminal section is SATISFIED without a write; an unresolved earlier run -> FAIL_CLOSED', async () => {
  const s = store([sectionFixture({ startPageNumber: 5 })]);
  const { x, m } = await run(s, task(sps({ startPageNumber: 5 })));
  assert.equal(status(x), OPERATION_STATUS.SATISFIED);
  assert.deepEqual([m.writes, m.commands], [0, 0]);
  const r = await run(store([sectionFixture({ startPageNumber: 5 })]), { ...task(sps({ startPageNumber: 5 })), recovery: { unresolvedUncertainOperations: ['s0'] } });
  assert.equal(reason(r.x), 'preservation-baseline-unavailable');
  assert.equal(r.driver.calls.dispatch, 0);
});

test('lowering: public section API only; margins REORDERED to (left,top,right,bottom); size as a two-arg call; no internal/cursor call', () => {
  const op = planOperations(task(sps({ size: { width: 12240, height: 15840 }, margins: { top: 1, right: 2, bottom: 3, left: 4 }, columns: { count: 2, space: 720 }, startPageNumber: 5 }))).operations[0];
  const low = ps.lower(op, { kind: 'section_page_setup', sectionIndex: 0, sectionCount: 1, fields: op.fields });
  assert.deepEqual([low.meta.sectionIndex, low.meta.sectionCount], [0, 1]);
  assert.match(low.body, /secs\.length !== m\.sectionCount/);
  assert.match(low.body, /sec\.SetPageSize\(g\.width, g\.height\)/);
  assert.match(low.body, /sec\.SetPageMargins\(g\.left, g\.top, g\.right, g\.bottom\)/);
  assert.match(low.body, /sec\.SetEqualColumns\(g\.count, g\.space\)/);
  assert.match(low.body, /sec\.SetStartPageNumber\(m\.fields\.startPageNumber\)/);
  assert.doesNotMatch(low.body, /evaluate|asc_|private_|\.Document\b|\.Search\(|MoveCursor|SetDocPosType/);
  assert.equal(op.preservationMode, 'PRIMITIVE_INTEGRAL');
});
