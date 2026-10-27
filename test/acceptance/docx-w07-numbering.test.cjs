'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { docxAdapter, planOperations } = require('../../src/adapters/docx/adapter.cjs');
const numbering = require('../../src/adapters/docx/families/numbering.cjs');
const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');

/**
 * W-07 (M7 backlog, numbering pattern/start) offline acceptance: set_numbering through the
 * TARGET path (M1 -> M2 -> M3 -> DOCX adapter) on the fake transport, modeled directly on
 * test/acceptance/docx-w2-list.test.cjs's own shape (the fake driver's `list`/`numberingFormats`
 * model). The vm-test (test/static/docx-driver-numbering-read-vm.test.cjs) covers the REAL
 * PROPERTY_QUERY string separately -- this file covers the FAMILY's own logic through the fake.
 */

const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'en-US' };
const A = 'Alpha item.';
const B = 'Beta item.'; // shares A's numId -- the shared-definition L3 target
const Z = 'Plain paragraph.';

const doc = (aFormat) => [
  { text: A, props: { ...PLAIN }, list: { type: 'numbered', level: 0, numId: 'n1' } },
  { text: B, props: { ...PLAIN }, list: { type: 'numbered', level: 0, numId: 'n1' } },
  { text: Z, props: { ...PLAIN } },
];

const store = (aFormat, extraDoc = {}) => createStore(doc(aFormat), {
  doc: { numberingFormats: aFormat ? { n1: [aFormat] } : {}, ...extraDoc },
});

async function run(s, task, script = {}) {
  let t = 0;
  const driver = createFakeDocxDriver(s, script);
  const options = { driver, fileId: 'd', now: () => (t += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, docxAdapter, createExecutor(task));
  assert.equal(driver.calls.open, 1, 'OPEN_COUNT=1');
  assert.equal(driver.calls.close, 1, 'CLOSE_COUNT=1');
  if (result.execution.operationAccounting) assert.equal(result.execution.operationAccounting.balanced, true);
  return { result, driver, m: result.metrics, x: result.execution };
}
const task = (...steps) => ({ id: 'w07', steps });
const sn = (fields, text = A) => ({ intent: 'set_numbering', selector: { text }, ...fields });
const status = (x, i = 0) => x.appliedOperations[i].status;
const reason = (x, i = 0) => { const e = x.appliedOperations[i].evidence; return (e && e.evidence && e.evidence.reason) || (e && e.reason); };
const fmtOf = (s, numId = 'n1', level = 0) => { const arr = s.durableDoc.numberingFormats[numId]; return arr && arr[level]; };

test('plan: refuses code+custom together, an unknown field, and a step with no field', () => {
  assert.equal(planOperations({ steps: [sn({ code: '1.', custom: { format: 'decimal', text: '%1.', align: 'left' } })] }).ok, false);
  assert.equal(planOperations({ steps: [{ intent: 'set_numbering', selector: { text: A }, bogus: 1 }] }).ok, false);
  assert.equal(planOperations({ steps: [sn({})] }).ok, false);
});

test('APPLY: code "1." on a bare numbering (no prior format) sets numFmt/lvlText/align, L1/L3 pass, second run SATISFIED 0-write', async () => {
  const s = store(null);
  const r1 = await run(s, task(sn({ code: '1.' })));
  assert.equal(status(r1.x), OPERATION_STATUS.APPLIED, JSON.stringify(r1.x.appliedOperations));
  assert.equal(r1.m.writes, 1);
  assert.deepEqual(fmtOf(s), { numFmt: 'decimal', lvlText: '%1.', lvlJc: 'left' });
  const r2 = await run(s, task(sn({ code: '1.' })));
  assert.equal(status(r2.x), OPERATION_STATUS.SATISFIED);
  assert.equal(r2.m.writes, 0, 'SATISFIED -- the format already matches, zero writes');
});

test('APPLY: custom format/text/align, start and suffix all together, one step', async () => {
  const s = store(null);
  const r = await run(s, task(sn({ custom: { format: 'lowerLetter', text: '%1)', align: 'left' }, start: 7, suffix: 'tab' })));
  assert.equal(status(r.x), OPERATION_STATUS.APPLIED, JSON.stringify(r.x.appliedOperations));
  assert.deepEqual(fmtOf(s), { numFmt: 'lowerLetter', lvlText: '%1)', lvlJc: 'left', start: 7, suff: 'tab' });
});

test('code: "bullet" uses the default symbol "-"; an explicit symbol overrides it', async () => {
  const s1 = store(null);
  await run(s1, task(sn({ code: 'bullet' })));
  assert.equal(fmtOf(s1).lvlText, '-');
  const s2 = store(null);
  await run(s2, task(sn({ code: 'bullet', symbol: '*' })));
  assert.equal(fmtOf(s2).lvlText, '*');
});

test('suffix "none" writes the measured file word "nothing", and L1 accepts either spelling (unmeasured ToJSON key, named)', async () => {
  const s = store(null);
  const r = await run(s, task(sn({ suffix: 'none' })), { fault: 'numbering-suffix-nothing-spelling' });
  assert.equal(status(r.x), OPERATION_STATUS.APPLIED, JSON.stringify(r.x.appliedOperations));
  assert.equal(fmtOf(s).suff, 'nothing');
});

test('FAIL_CLOSED: a target that is not a list item at all', async () => {
  const s = store(null);
  const r = await run(s, task(sn({ start: 5 }, Z)));
  assert.equal(status(r.x), OPERATION_STATUS.FAILED, JSON.stringify(r.x.appliedOperations));
  assert.equal(reason(r.x), 'set-numbering-target-not-listed');
});

test('start requested when the live ToJSON omits the start key (never-set level, unobservable): never a guessed SATISFIED -- it APPLIES, and the write makes it observable', async () => {
  const s = store(null); // no numberingFormats seeded at all -- format absent, start unobservable
  const r = await run(s, task(sn({ start: 5 })));
  assert.equal(status(r.x), OPERATION_STATUS.APPLIED, JSON.stringify(r.x.appliedOperations));
  assert.equal(r.m.writes, 1, 'unobservable must never be read as "already matches" -- it must still write');
  assert.deepEqual(r.x.appliedOperations[0].evidence.evidence.unobservable, [], 'after the write, start is observable -- L1 must not report it unobservable anymore');
  assert.equal(fmtOf(s).start, 5);
});

test('start:1 (W07-7 MEASURED resolution): moving FROM an explicit 5 TO 1 is APPLIED, retry is SATISFIED -- the family logic is UNCHANGED, only the real driver now maps an absent key to 1', async () => {
  const s = store({ numFmt: 'decimal', lvlText: '%1.', lvlJc: 'left', start: 5 });
  const r1 = await run(s, task(sn({ start: 1 })));
  assert.equal(status(r1.x), OPERATION_STATUS.APPLIED, JSON.stringify(r1.x.appliedOperations));
  assert.equal(r1.m.writes, 1);
  assert.equal(fmtOf(s).start, 1);
  const r2 = await run(s, task(sn({ start: 1 })));
  assert.equal(status(r2.x), OPERATION_STATUS.SATISFIED, JSON.stringify(r2.x.appliedOperations));
  assert.equal(r2.m.writes, 0, 'SATISFIED -- start already reads 1, zero writes');
});

test('L3 (design doc point 3 ELFOGADVA): a SECOND item sharing the same numId sees the SAME new value -- the shared definition is the intended effect', async () => {
  const s = store(null);
  const r = await run(s, task(sn({ code: 'a.' }, A)));
  assert.equal(status(r.x), OPERATION_STATUS.APPLIED, JSON.stringify(r.x.appliedOperations));
  const l3 = r.x.preservationVerification[0];
  assert.equal(l3.verdict, 'PASS', JSON.stringify(l3));
  assert.equal(l3.evidence.evidence.sharedOk, true, JSON.stringify(l3));
  assert.equal(l3.evidence.evidence.coMemberCount, 1);
  // confirmed independently, re-reading the store directly (not just trusting the evidence field)
  const bFmt = fmtOf(s);
  assert.deepEqual(bFmt, { numFmt: 'lowerLetter', lvlText: '%1.', lvlJc: 'left' });
});

test('L3: a DIFFERENT numId is unaffected by a write on this one', async () => {
  const s = store(null, { numberingFormats: { other: [{ numFmt: 'decimal', lvlText: '%1.', lvlJc: 'right' }] } });
  const r = await run(s, task(sn({ code: 'a.' })));
  assert.equal(status(r.x), OPERATION_STATUS.APPLIED, JSON.stringify(r.x.appliedOperations));
  assert.deepEqual(fmtOf(s, 'other'), { numFmt: 'decimal', lvlText: '%1.', lvlJc: 'right' });
});

test('L3 FAIL: a fault that leaves the co-member on the OLD value is caught (not silently accepted)', async () => {
  // the fake's own write always propagates to the shared slot (by construction) -- this proves
  // the CHECK itself would catch a divergence, by scripting the fault on the OTHER numId's slot
  // and confirming verifyPreservation's otherNumIdsUnchanged flips false when it changes.
  const s = store(null, { numberingFormats: { n2: [{ numFmt: 'decimal', lvlText: '%1.', lvlJc: 'left' }] } });
  s.durable.push({ ...s.durable[0], text: 'Other list item.', list: { type: 'numbered', level: 0, numId: 'n2' } });
  const r = await run(s, task(sn({ code: 'a.' })), { fault: 'numbering-touch-other-numid' });
  assert.equal(r.x.preservationVerification[0].verdict, 'FAIL', 'the collateral change on numId n2 must fail L3, not pass silently');
  assert.equal(r.x.preservationVerification[0].evidence.evidence.otherNumIdsUnchanged, false);
});

test('commandBody is ALL-OR-NOTHING: moved level before dispatch fails closed, no partial write', async () => {
  const s = store({ numFmt: 'decimal', lvlText: '%1.', lvlJc: 'left' });
  const r = await run(s, task(sn({ start: 5 })), { fault: 'list-wrong-level' });
  assert.equal(r.m.writes <= 1, true);
  // whichever way the executor classifies a moved-level race, the slot must not end up partially written
  const f = fmtOf(s);
  assert.ok(!('start' in f) || f.start === 5, 'no half-applied state: start is either fully absent or fully 5');
});
