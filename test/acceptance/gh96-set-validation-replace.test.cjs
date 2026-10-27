'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const validation = require('../../src/adapters/xlsx/families/validation.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * GH #96 (owner request): set_validation `replace: true` -- an explicit opt-in to replace an EXISTING,
 * DIFFERENT validation on (or fragmenting) the target with the requested one (Delete then Add), instead of
 * today's fail-closed refusal (xlsx-validation-exists-different / xlsx-validation-overlaps-another-range).
 * Without `replace` those refusals are UNCHANGED (test/acceptance/xlsx-w6-validation.test.cjs, untouched).
 *
 * The exact Delete+Add mechanics on a fragmented (per-cell) pre-existing state is LIVE-pending (a separate,
 * parallel unit measures it) -- this is prep work against the stated most-likely shape, ahead of that
 * measurement. The fake's own write-application (`applyValidation`) already removes any rule overlapping the
 * written range before adding the new one regardless of the exact command body, so these tests exercise
 * plan/classify/verify's behaviour, not the literal Delete() call text (covered separately, see the command-
 * body test below).
 */

const book = () => [{ name: 'V', cells: { B2: { v: 1 }, B3: { v: 2 }, B4: { v: 3 }, B5: { v: 4 } } }];
const set = (range = 'B2:B5', extra = {}) => ({ id: 'gh96', steps: [{ intent: 'set_validation', sheet: 'V', range, type: 'wholeNumber', operator: 'between', alertStyle: 'stop', min: 1, max: 10, ...extra }] });
const rule = (range, f1 = '1', f2 = '10') => ({ kind: 'validation', type: 1, range, formula: f1, formula2: f2, api: { type: 'xlValidateWholeNumber', alertStyle: 'xlValidAlertStop', operator: 'xlBetween' } });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}
const rules = store => store.durable.find(s => s.name === 'V').rules || [];
const reason = result => result.execution.appliedOperations[0].evidence.evidence.reason;

test('replace:true over an EXACT, different rule -- APPLIED (Delete+Add), retry SATISFIED 0/0', async () => {
  const b = book();
  b[0].rules = [rule('B2:B5', '1', '99')]; // the SAME shape xlsx-w6-validation.test.cjs's "no silent Modify" test uses
  const store = createStore(b);
  const a = await run(store, set('B2:B5', { replace: true }));
  assert.equal(a.result.ok, true, JSON.stringify(a.result.execution.appliedOperations[0].evidence));
  assert.equal(a.result.execution.appliedOperations[0].status, OPERATION_STATUS.APPLIED);
  assert.deepEqual(rules(store).map(r => [r.range, r.formula, r.formula2]), [['B2:B5', '1', '10']]);
  const a2 = await run(store, set('B2:B5', { replace: true }));
  assert.equal(a2.m.commands, 0);
  assert.equal(a2.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('replace:true over FRAGMENTED (per-cell) pre-existing rules -- APPLIED, the whole range ends up with ONE unified rule', async () => {
  const b = book();
  // three separate, per-cell rules covering the same overall range the new request targets
  b[0].rules = [rule('B2', '1', '5'), rule('B3:B4', '2', '8'), rule('B5', '0', '3')];
  const store = createStore(b);
  const a = await run(store, set('B2:B5', { replace: true }));
  assert.equal(a.result.ok, true, JSON.stringify(a.result.execution.appliedOperations[0].evidence));
  assert.deepEqual(rules(store).map(r => r.range), ['B2:B5']);
});

test('replace:false (default) is UNCHANGED: the exact-different and fragmented cases still refuse before dispatch', async () => {
  for (const pre of [[rule('B2:B5', '1', '99')], [rule('B2', '1', '5'), rule('B3:B4', '2', '8')]]) {
    const b = book();
    b[0].rules = pre;
    const { result, driver } = await run(createStore(b), set('B2:B5'));
    assert.equal(result.ok, false, JSON.stringify(pre));
    assert.equal(driver.calls.dispatch, 0, JSON.stringify(pre));
    assert.match(reason(result), /xlsx-validation-(exists-different|overlaps-another-range)/, JSON.stringify(pre));
  }
});

// a maintainer's P96 LIVE probe (917246c's own commit message): a range spanning per-cell validations refuses
// xlsx-validation-read-unobservable at gate() -- the AGGREGATE public reader itself fails there, so
// classify()'s 'other'-overlap check (which depends on gate() succeeding first) never even runs. `replace`
// must tolerate THIS specific gap (gate(), not just classify()'s own refusals) to actually reach the
// fragmented case it is meant to fix. The fake driver has no per-window-granular unobservable control, so
// this is unit-tested directly against gate()/classify() with a constructed observation.
test('gate()/classify() (replace:true): an UNOBSERVABLE aggregate range read (per-cell fragments) is tolerated, using the SEPARATE internal rules reader', () => {
  const op = planOperations(set('B2:B5', { replace: true })).operations[0];
  const primary = {
    format: 'xlsx', complete: true, sheets: [{ name: 'V' }],
    rules: { complete: true, bySheet: { V: [
      { kind: 'validation', range: 'B2', type: 1, formula: '1', formula2: '5' },
      { kind: 'validation', range: 'B3:B4', type: 1, formula: '2', formula2: '8' },
      { kind: 'validation', range: 'B5', type: 1, formula: '0', formula2: '3' },
    ] } },
    validationReads: { 'V!B2:B5': 'UNOBSERVABLE:threw' }, // the AGGREGATE read -- this is the measured failure
  };
  const g = validation.classify(primary, op);
  assert.equal(g.state, 'APPLY', JSON.stringify(g));
  assert.ok(g.evidence.replacing, JSON.stringify(g));

  // NEG. CONTROL: replace:false must still hard-refuse on the SAME observation (the tolerance is replace-only).
  const opNoReplace = planOperations(set('B2:B5')).operations[0];
  const g2 = validation.classify(primary, opNoReplace);
  assert.equal(g2.state, 'FAIL_CLOSED', JSON.stringify(g2));
  assert.equal(g2.evidence.reason, 'xlsx-validation-read-unobservable', JSON.stringify(g2));
});

test('replace:true with NO pre-existing rule at all -- behaves exactly like a plain set_validation (APPLIED, one rule)', async () => {
  const store = createStore(book());
  const { result } = await run(store, set('B2:B5', { replace: true }));
  assert.equal(result.ok, true, JSON.stringify(result.execution.appliedOperations[0].evidence));
  assert.deepEqual(rules(store).map(r => r.range), ['B2:B5']);
});

test('replace:true when the rule ALREADY matches exactly -- SATISFIED, 0 dispatch (no gratuitous Delete+Add)', async () => {
  const b = book();
  b[0].rules = [rule('B2:B5', '1', '10')];
  const store = createStore(b);
  const { result, driver } = await run(store, set('B2:B5', { replace: true }));
  assert.equal(result.ok, true);
  assert.equal(result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
  assert.equal(driver.calls.dispatch, 0);
});

test('clear_validation ignores `replace` entirely -- the field is not accepted on that intent', () => {
  const p = planOperations({ id: 'gh96', steps: [{ intent: 'clear_validation', sheet: 'V', range: 'B2:B5', replace: true }] });
  assert.equal(p.ok, false);
  assert.match(p.detail, /does not support field/);
});

test('planner: replace must be a boolean when present', () => {
  const p = planOperations(set('B2:B5', { replace: 'yes' }));
  assert.equal(p.ok, false);
  assert.match(p.detail, /replace must be true or false/);
});

test('RED: replace:true still fails L1 if the write is ignored (the fault family is unaffected by the new field)', async () => {
  const b = book();
  b[0].rules = [rule('B2:B5', '1', '99')];
  const store = createStore(b);
  const { result } = await run(store, set('B2:B5', { replace: true }), { fault: 'ignore' });
  assert.equal(result.ok, false);
  assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL');
});

// The fake driver's own write-application always collapses a replace into ONE unified rule (its
// `applyValidation` unconditionally removes every overlapping rule before adding), so the full-stack tests
// above never actually exercise a DIVERGENCE between the aggregate range-level read and a per-cell read -- the
// exact point of the per-cell check (a fragmented pre-existing state may not collapse the way a fresh range
// does, per the module doc). Unit-tested directly against verifyPrimitive with constructed observations.
test("verifyPrimitive (replace:true): the per-cell check catches a cell the AGGREGATE read alone would miss", () => {
  const op = planOperations(set('B2:B5', { replace: true })).operations[0];
  const want = op.want;
  const matchingRead = { ...want, absent: false };
  const primaryGood = {
    format: 'xlsx', complete: true, sheets: [{ name: 'V' }],
    rules: { complete: true, bySheet: { V: [{ kind: 'validation', range: 'B2:B5', ...want }] } },
    validationReads: {
      'V!B2:B5': matchingRead, // the aggregate range-level read: matches
      'V!B2': matchingRead, 'V!B3': matchingRead, 'V!B4': { ...matchingRead, formula2: '99' }, 'V!B5': matchingRead, // B4 is STALE
    },
  };
  const r1 = validation.verifyPrimitive(op, null, { primary: primaryGood });
  assert.equal(r1.ok, false, JSON.stringify(r1.evidence));
  assert.deepEqual(r1.evidence.cellMisses, ['B4']);

  const primaryAllGood = { ...primaryGood, validationReads: { ...primaryGood.validationReads, 'V!B4': matchingRead } };
  const r2 = validation.verifyPrimitive(op, null, { primary: primaryAllGood });
  assert.equal(r2.ok, true, JSON.stringify(r2.evidence));
  assert.deepEqual(r2.evidence.cellMisses, []);

  // NEG. CONTROL: replace:false never runs the per-cell path at all -- the SAME stale B4 is invisible to it
  // (the field simply is not in the evidence), proving the two code paths are genuinely distinct.
  const opNoReplace = planOperations(set('B2:B5')).operations[0];
  const r3 = validation.verifyPrimitive(opNoReplace, null, { primary: primaryGood });
  assert.equal(r3.ok, true, JSON.stringify(r3.evidence));
  assert.equal('cellMisses' in r3.evidence, false);
});

test('the command body: replace:true calls Delete() before Add(); replace:false/absent calls Add() only', () => {
  const log = [];
  const ws = { GetName: () => 'V', GetRange: a => ({ GetValidation: () => ({ Add: (...x) => log.push([a, 'Add', ...x]), Delete: () => log.push([a, 'Delete']) }) }) };
  const runBody = op => new Function('Api', validation.lower(op).body)({ GetSheet: () => ws });

  log.length = 0;
  assert.deepEqual(runBody(planOperations(set('B2:B5', { replace: true })).operations[0]), { ok: true });
  assert.deepEqual(log, [['B2:B5', 'Delete'], ['B2:B5', 'Add', 'xlValidateWholeNumber', 'xlValidAlertStop', 'xlBetween', '1', '10']]);

  log.length = 0;
  assert.deepEqual(runBody(planOperations(set('B2:B5')).operations[0]), { ok: true });
  assert.deepEqual(log, [['B2:B5', 'Add', 'xlValidateWholeNumber', 'xlValidAlertStop', 'xlBetween', '1', '10']]);
});

// a maintainer review: a rule that EXTENDS beyond the target would be deleted/shrunk OUTSIDE the target by the range Delete(),
// and the replace L3 strips every touching rule -- so that outside change would go unseen. replace refuses it, named.
test('replace:true refuses a touching rule that extends BEYOND the target range (named, 0 dispatch)', async () => {
  const b = book();
  b[0].rules = [rule('B2:B5', '1', '99')];
  const { result, driver } = await run(createStore(b), set('B2:B3', { replace: true }));
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.match(reason(result), /xlsx-validation-replace-rule-extends-beyond-range/);
});
