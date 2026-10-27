'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const validation = require('../../src/adapters/xlsx/families/validation.cjs');
const cf = require('../../src/adapters/xlsx/families/cf.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * M7 R17 / R18 (owner M7 GO 2026-10-03; tier D probe docs/evidence/M7-D-PROBE.md): the measured validation types /
 * operators / alert styles and the measured conditional-format types / operators. Every expected Add call below is the
 * probe's own call, written out (m7-d-probe1-20261003.json, ACT.v / ACT.c), not derived by the code under test.
 */

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}
const op = result => result.execution.appliedOperations[0];

// ---- R17 validation --------------------------------------------------------------------------------------------------

const vbook = () => [{ name: 'V', cells: { B2: { v: 1 }, B3: { v: 2 }, Z1: { v: 'a' }, Z2: { v: 'b' }, Z3: { v: 'c' } } }, { name: 'O', cells: { A1: { v: 'keep' } } }];
const vset = extra => ({ id: 'm7-d-r17', steps: [{ intent: 'set_validation', sheet: 'V', range: 'B2:B5', ...extra }] });

for (const [name, extra, addArgs] of [
  ['list inline', { type: 'list', alertStyle: 'stop', items: ['a', 'b', 'c'] }, ['xlValidateList', 'xlValidAlertStop', 'xlBetween', '"a,b,c"']],
  ['list range', { type: 'list', alertStyle: 'stop', listRange: 'Z1:Z3' }, ['xlValidateList', 'xlValidAlertStop', 'xlBetween', '=$Z$1:$Z$3']],
  ['decimal between, warning', { type: 'decimal', alertStyle: 'warning', operator: 'between', min: 0.5, max: 9.5 }, ['xlValidateDecimal', 'xlValidAlertWarning', 'xlBetween', '0.5', '9.5']],
  ['text length <=, information', { type: 'textLength', alertStyle: 'information', operator: 'lessEqual', value: 5 }, ['xlValidateTextLength', 'xlValidAlertInformation', 'xlLessEqual', '5']],
  ['whole notBetween', { type: 'wholeNumber', alertStyle: 'stop', operator: 'notBetween', min: 1, max: 3 }, ['xlValidateWholeNumber', 'xlValidAlertStop', 'xlNotBetween', '1', '3']],
  ['whole greater', { type: 'wholeNumber', alertStyle: 'stop', operator: 'greater', value: 4 }, ['xlValidateWholeNumber', 'xlValidAlertStop', 'xlGreater', '4']],
  // GH #301 D1: the four remaining one-value operators (ONE table, validation.cjs), same types/value field as
  // greater/lessEqual above -- spanning bounds (wholeNumber/decimal/textLength), time and date to prove the
  // shared ONE table reaches every caller, not just bounds.
  ['whole equal', { type: 'wholeNumber', alertStyle: 'stop', operator: 'equal', value: 7 }, ['xlValidateWholeNumber', 'xlValidAlertStop', 'xlEqual', '7']],
  ['decimal notEqual', { type: 'decimal', alertStyle: 'warning', operator: 'notEqual', value: 2.5 }, ['xlValidateDecimal', 'xlValidAlertWarning', 'xlNotEqual', '2.5']],
  ['text length greaterEqual', { type: 'textLength', alertStyle: 'information', operator: 'greaterEqual', value: 3 }, ['xlValidateTextLength', 'xlValidAlertInformation', 'xlGreaterEqual', '3']],
  ['time less, stop', { type: 'time', alertStyle: 'stop', operator: 'less', value: '09:30' }, ['xlValidateTime', 'xlValidAlertStop', 'xlLess', String((9 * 3600 + 30 * 60) / 86400)]],
  ['date greaterEqual, warning', { type: 'date', alertStyle: 'warning', operator: 'greaterEqual', value: '2024-01-02' }, ['xlValidateDate', 'xlValidAlertWarning', 'xlGreaterEqual', '45293']],
  // L-22a (allowed EURO GO; L-22 probe test/live/xlsx-l22-validation-types-probe.live.cjs, MEASURED_OK): min/max
  // are 24-hour text, written as a day fraction (06:00 -> 0.25, 18:00 -> 0.75) -- the SAME Add() call shape.
  ['time between, stop', { type: 'time', alertStyle: 'stop', operator: 'between', min: '06:00', max: '18:00' }, ['xlValidateTime', 'xlValidAlertStop', 'xlBetween', '0.25', '0.75']],
  ['time greater, warning', { type: 'time', alertStyle: 'warning', operator: 'greater', value: '12:00:30' }, ['xlValidateTime', 'xlValidAlertWarning', 'xlGreater', String((12 * 3600 + 30) / 86400)]],
  // L-22d v2: date -- ISO 'YYYY-MM-DD' text, written as the 1900-system serial. EVERY real calendar date is accepted now,
  // including a genuinely locale-ambiguous one (2024-03-04 / 2024-11-09, M/D vs D/M) -- identity is the bridge's own raw
  // serial, not the locale-ambiguous public text (see the family header comment).
  ['date between, stop (locale-ambiguous dates)', { type: 'date', alertStyle: 'stop', operator: 'between', min: '2024-03-04', max: '2024-11-09' }, ['xlValidateDate', 'xlValidAlertStop', 'xlBetween', '45355', '45605']],
  ['date lessEqual, warning', { type: 'date', alertStyle: 'warning', operator: 'lessEqual', value: '2024-01-02' }, ['xlValidateDate', 'xlValidAlertWarning', 'xlLessEqual', '45293']],
  // L-22b (same standing GO; the same probe, custom, MEASURED): NO operator; the Add() call's own operator
  // argument is the measured placeholder xlBetween (the editor ignores it for this type).
  ['custom absolute formula, stop', { type: 'custom', alertStyle: 'stop', formula: '=$H$2>0' }, ['xlValidateCustom', 'xlValidAlertStop', 'xlBetween', '=$H$2>0']],
  // L-22e (card an internal branch; measured LIVE r1-r3, #ref/#ref/#ref):
  // relative and mixed references are accepted now too -- the fake's own default activeCell (the range's own
  // first cell, no quirk simulated) means these pass through unshifted here; the dedicated RED/GREEN tests below
  // simulate the measured active-cell quirk for real and prove the re-anchoring itself.
  ['custom relative formula, stop', { type: 'custom', alertStyle: 'stop', formula: '=B2>0' }, ['xlValidateCustom', 'xlValidAlertStop', 'xlBetween', '=B2>0']],
  ['custom mixed formula, stop', { type: 'custom', alertStyle: 'stop', formula: '=$B2>0' }, ['xlValidateCustom', 'xlValidAlertStop', 'xlBetween', '=$B2>0']],
]) {
  test(`R17 ${name}: the measured Add call; APPLIED, retry SATISFIED 0/0`, async () => {
    const p = planOperations(vset(extra));
    assert.equal(p.ok, true, p.detail);
    const log = [];
    const ws = { GetName: () => 'V', GetRange: () => ({ GetValidation: () => ({ Add: (...x) => log.push(x) }) }) };
    new Function('Api', validation.lower(p.operations[0]).body)({ GetSheet: () => ws });
    assert.deepEqual(log, [addArgs]);
    const store = createStore(vbook());
    const { result, m } = await run(store, vset(extra));
    assert.equal(result.ok, true, JSON.stringify([result.execution.primitiveVerification, op(result).evidence]));
    assert.equal(m.commands, 1);
    const again = await run(store, vset(extra));
    assert.equal(again.m.commands, 0);
    assert.equal(op(again.result).status, OPERATION_STATUS.SATISFIED);
  });
}

// GH #301 D1: with equal/notEqual/less/greaterEqual now also measured, all eight of Excel's one-/two-formula
// validation operators (between, notBetween, greater, lessEqual, equal, notEqual, less, greaterEqual) are
// accepted on every bound-taking type -- there is no real operator left to use as the "unmeasured" NEG.
// CONTROL below (same situation, and the same fix, as X-L13/7's UNMEASURED_OP in xlsx-m7-s1.test.cjs).
const UNMEASURED_VALIDATION_OP = 'xlOperatorNotYetMeasured';

test('R17 refused at planning: date, unmeasured operators, bad bounds and items', () => {
  for (const [extra, re] of [
    // L-22d: date was refused by name; it is a type now, and a NUMBER (or any non-ISO text) is refused as a bound.
    [{ type: 'date', alertStyle: 'stop', operator: 'between', min: 45364, max: 45621 }, /YYYY-MM-DD/],
    [{ type: 'date', alertStyle: 'stop', operator: 'between', min: '2024-3-13', max: '2024-11-25' }, /YYYY-MM-DD/],
    [{ type: 'date', alertStyle: 'stop', operator: 'between', min: '2024-02-30', max: '2024-11-25' }, /YYYY-MM-DD/],
    [{ type: 'date', alertStyle: 'stop', operator: 'between', min: '03/13/2024', max: '11/25/2024' }, /YYYY-MM-DD/],
    [{ type: 'date', alertStyle: 'stop', operator: 'between', min: '2024-11-25', max: '2024-03-13' }, /min <= max/],
    [{ type: 'date', alertStyle: 'stop', operator: 'between', min: '1900-02-13', max: '2024-11-25' }, /1900-03-01/],
    [{ type: 'date', alertStyle: 'stop', operator: UNMEASURED_VALIDATION_OP, value: '2024-03-13' }, /operator must be one of/],
    [{ type: 'date', alertStyle: 'stop', operator: 'greater', value: '2024-03-13', items: ['a'] }, /items belongs to type list/],
    // L-22d v2: a locale-ambiguous date (M/D vs D/M disagree) is NO LONGER refused at planning -- see the
    // matrix's own 'date between, stop (locale-ambiguous dates)' row above for the accepted shape.
    [{ type: 'wholeNumber', alertStyle: 'stop', operator: UNMEASURED_VALIDATION_OP, value: 1 }, /operator must be one of/],
    [{ type: 'decimal', alertStyle: 'stop', operator: 'between', min: 2, max: 1 }, /min <= max/],
    [{ type: 'textLength', alertStyle: 'stop', operator: 'greater', value: -1 }, /non-negative/],
    [{ type: 'list', alertStyle: 'stop', items: ['a,b'] }, /without a comma/],
    [{ type: 'list', alertStyle: 'stop', items: ['a'], listRange: 'Z1:Z3' }, /exactly one of/],
    [{ type: 'wholeNumber', alertStyle: 'stop', operator: 'greater', value: 4, items: ['a'] }, /items belongs to type list/],
    // L-22a: time -- 24-hour text only (a number, or an AM/PM-suffixed / out-of-range string, is refused the
    // same way); min <= max is the FRACTION order (chronological), not a string one.
    [{ type: 'time', alertStyle: 'stop', operator: 'between', min: 6, max: 18 }, /24-hour/],
    [{ type: 'time', alertStyle: 'stop', operator: 'between', min: '6:00 AM', max: '18:00' }, /24-hour/],
    [{ type: 'time', alertStyle: 'stop', operator: 'between', min: '25:00', max: '18:00' }, /24-hour/],
    [{ type: 'time', alertStyle: 'stop', operator: 'between', min: '18:00', max: '06:00' }, /min <= max/],
    [{ type: 'time', alertStyle: 'stop', operator: UNMEASURED_VALIDATION_OP, value: '12:00' }, /operator must be one of/],
    // L-22b/e: custom -- no formula at all, another sheet, a name, INDIRECT/OFFSET, an external/array `[` --
    // every case the probe did NOT measure, refused named, never guessed. Relative/mixed references are no
    // longer in this list -- see the matrix above + the dedicated RED/GREEN tests below (L-22e).
    [{ type: 'custom', alertStyle: 'stop' }, /non-empty string starting with =/],
    [{ type: 'custom', alertStyle: 'stop', formula: '=H2>0', replace: true }, /does not support replace: true/],
    [{ type: 'custom', alertStyle: 'stop', formula: '=$H2>0', replace: true }, /does not support replace: true/],
    [{ type: 'custom', alertStyle: 'stop', formula: '=H$2>0', replace: true }, /does not support replace: true/],
    [{ type: 'custom', alertStyle: 'stop', formula: "=Other!$H$2>0" }, /must not reference another sheet/],
    [{ type: 'custom', alertStyle: 'stop', formula: "='Sheet 2'!$H$2>0" }, /must not reference another sheet/],
    [{ type: 'custom', alertStyle: 'stop', formula: '=MyNamedRange>0' }, /name, INDIRECT\/OFFSET or another unsupported function/],
    [{ type: 'custom', alertStyle: 'stop', formula: '=INDIRECT("H2")>0' }, /name, INDIRECT\/OFFSET or another unsupported function/],
    [{ type: 'custom', alertStyle: 'stop', formula: '=OFFSET($H$2,0,0)>0' }, /name, INDIRECT\/OFFSET or another unsupported function/],
    [{ type: 'custom', alertStyle: 'stop', formula: "=[Book1.xlsx]Sheet1!$H$2>0" }, /must not reference another sheet/], // '!' is checked first -- also refused, just by the sheet rule
    [{ type: 'custom', alertStyle: 'stop', formula: '=Table1[Column1]>0' }, /external reference or a structured-table/], // brackets, no sheet qualifier -- the '[' rule itself
    [{ type: 'custom', alertStyle: 'stop', formula: '=AND($H$2>0,$H$3<10)' }, /name, INDIRECT\/OFFSET or another unsupported function/],
    [{ type: 'custom', alertStyle: 'stop', operator: 'between', formula: '=$H$2>0', min: 1 }, /does not take operator/],
    [{ type: 'custom', alertStyle: 'stop', formula: 'H2>0' }, /must be a non-empty string starting with =/],
  ]) {
    const p = planOperations(vset(extra));
    assert.equal(p.ok, false, JSON.stringify(extra));
    assert.match(p.detail, re);
  }
});

test('R17 RED: the second bound lands wrong -> L1 FAIL', async () => {
  const { result } = await run(createStore(vbook()), vset({ type: 'decimal', alertStyle: 'warning', operator: 'between', min: 0.5, max: 9.5 }), { fault: 'validation-wrong-bounds' });
  assert.equal(result.ok, false);
  assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL');
});

test('R17 RED (time, L-22a): the second bound lands wrong -> L1 FAIL, same as every other type', async () => {
  const { result } = await run(createStore(vbook()), vset({ type: 'time', alertStyle: 'stop', operator: 'between', min: '06:00', max: '18:00' }), { fault: 'validation-wrong-bounds' });
  assert.equal(result.ok, false);
  assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL');
});

test('R17 RED (time, L-22a): a garbled, unparseable locale echo on a FRESH write fails L1 -- never read as a match', async () => {
  const { result } = await run(createStore(vbook()), vset({ type: 'time', alertStyle: 'stop', operator: 'between', min: '06:00', max: '18:00' }), { fault: 'validation-time-garbled-echo' });
  assert.equal(result.ok, false);
  assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL');
});

test('R17 RED (time, L-22a): a PRE-EXISTING rule whose locale text cannot be parsed back is a NAMED gap, not a guessed match or difference, 0 dispatch', async () => {
  // a rule already on the range (e.g. written by another tool, or predating this feature) -- not something
  // THIS write produced; seeded directly, the same way a human-authored pre-existing rule would look.
  const bookWithGarbled = () => [{ name: 'V', cells: { B2: { v: 1 }, B3: { v: 2 } },
    rules: [{ kind: 'validation', type: 1, range: 'B2:B5', formula: 'not-a-time', formula2: 'not-a-time-2', api: { type: 'xlValidateTime', alertStyle: 'xlValidAlertStop', operator: 'xlBetween' } }] }];
  const { result, driver } = await run(createStore(bookWithGarbled()), vset({ type: 'time', alertStyle: 'stop', operator: 'between', min: '06:00', max: '18:00' }));
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0, 'refused before any dispatch -- never guessed');
  assert.equal(result.execution.appliedOperations[0].evidence.evidence.reason, 'xlsx-validation-time-unparseable');
});

test('R17 RED (date, L-22d v2): the second bound lands wrong -> L1 FAIL, same as every other type', async () => {
  const { result } = await run(createStore(vbook()), vset({ type: 'date', alertStyle: 'stop', operator: 'between', min: '2024-03-13', max: '2024-11-25' }), { fault: 'validation-wrong-bounds' });
  assert.equal(result.ok, false);
  assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL');
});

test('R17 RED (date, L-22d v2): the BRIDGE serial lands wrong (the public text still matches the request) -> L1 FAIL -- identity is the bridge, not the text', async () => {
  const { result } = await run(createStore(vbook()), vset({ type: 'date', alertStyle: 'stop', operator: 'between', min: '2024-03-13', max: '2024-11-25' }), { fault: 'validation-date-bridge-wrong' });
  assert.equal(result.ok, false);
  assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL');
});

test('R17 (date, L-22d v2): a garbled public echo (not N/N/YYYY at all) does NOT fail L1 -- identity is the bridge serial, which is correct; unparseable text is ignored by the cross-check, not a conflict', async () => {
  const { result, m } = await run(createStore(vbook()), vset({ type: 'date', alertStyle: 'stop', operator: 'between', min: '2024-03-13', max: '2024-11-25' }), { fault: 'validation-date-garbled-echo' });
  assert.equal(result.ok, true, JSON.stringify(result.execution.primitiveVerification));
  assert.equal(m.commands, 1);
});

for (const locale of ['MD', 'DM']) {
  test(`R17 date (L-22d v2): a genuinely locale-ambiguous date verifies under a ${locale} locale -- identity is the bridge serial, APPLIED + retry SATISFIED`, async () => {
    const store = createStore(vbook());
    const ext = { type: 'date', alertStyle: 'stop', operator: 'between', min: '2024-03-04', max: '2024-11-09' };
    const { result, m } = await run(store, vset(ext), { dateLocale: locale });
    assert.equal(result.ok, true, JSON.stringify([result.execution.primitiveVerification, op(result).evidence]));
    assert.equal(m.commands, 1);
    const again = await run(store, vset(ext), { dateLocale: locale });
    assert.equal(again.m.commands, 0);
    assert.equal(op(again.result).status, OPERATION_STATUS.SATISFIED);
  });
}

test('R17 (date, L-22d v2): a PRE-EXISTING rule whose public text names two possible dates, but the bridge serial IS the wanted one, is SATISFIED -- 0 dispatch (identity is the bridge, the ambiguous text is not a gap by itself)', async () => {
  const preexisting = () => [{ name: 'V', cells: { B2: { v: 1 } }, rules: [{ kind: 'validation', type: 1, range: 'B2:B5', formula: '3/4/2024', formula2: '11/9/2024', bridgeFormula: '45355', bridgeFormula2: '45605', api: { type: 'xlValidateDate', alertStyle: 'xlValidAlertStop', operator: 'xlBetween' } }] }];
  const { result, driver } = await run(createStore(preexisting()), vset({ type: 'date', alertStyle: 'stop', operator: 'between', min: '2024-03-04', max: '2024-11-09' }));
  assert.equal(result.ok, true, JSON.stringify(op(result).evidence));
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(op(result).status, OPERATION_STATUS.SATISFIED);
});

test('R17 RED (date, L-22d v2): a PRE-EXISTING rule whose BRIDGE serial is a different date is refused exists-different -- the public text is unparseable here, so this isolates the bridge-mismatch path from the read-conflict one', async () => {
  const wrongBridge = () => [{ name: 'V', cells: { B2: { v: 1 } }, rules: [{ kind: 'validation', type: 1, range: 'B2:B5', formula: 'not-a-date', formula2: 'not-a-date-2', bridgeFormula: '1', bridgeFormula2: '2', api: { type: 'xlValidateDate', alertStyle: 'xlValidAlertStop', operator: 'xlBetween' } }] }];
  const { result, driver } = await run(createStore(wrongBridge()), vset({ type: 'date', alertStyle: 'stop', operator: 'between', min: '2024-03-13', max: '2024-11-25' }));
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0, 'a pre-existing different rule is refused, not overwritten');
  assert.equal(op(result).evidence.evidence.reason, 'xlsx-validation-exists-different');
});

test('R17 RED (date, L-22d v2): a PRE-EXISTING rule whose public text names a date that is NEITHER the bridge serial NOR unparseable is a NAMED read-conflict, 0 dispatch -- the two sources disagree, never silently resolved', async () => {
  // bridge = 2024-03-13 (45364); public text = 3/23/2024, a single unambiguous reading (day 23 > 12) that disagrees with the bridge
  const conflict = () => [{ name: 'V', cells: { B2: { v: 1 } }, rules: [{ kind: 'validation', type: 1, range: 'B2:B5', formula: '3/23/2024', formula2: '11/25/2024', bridgeFormula: '45364', bridgeFormula2: '45621', api: { type: 'xlValidateDate', alertStyle: 'xlValidAlertStop', operator: 'xlBetween' } }] }];
  const { result, driver } = await run(createStore(conflict()), vset({ type: 'date', alertStyle: 'stop', operator: 'between', min: '2024-03-13', max: '2024-11-25' }));
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(op(result).evidence.evidence.reason, 'xlsx-validation-date-read-conflict');
});

test('R17 (date, L-22d v2): a PRE-EXISTING rule whose public text is not N/N/YYYY at all is NOT a conflict (ignored) -- the bridge serial alone decides SATISFIED, 0 dispatch', async () => {
  const garbledPreexisting = () => [{ name: 'V', cells: { B2: { v: 1 } }, rules: [{ kind: 'validation', type: 1, range: 'B2:B5', formula: 'not-a-date', formula2: 'not-a-date-2', bridgeFormula: '45364', bridgeFormula2: '45621', api: { type: 'xlValidateDate', alertStyle: 'xlValidAlertStop', operator: 'xlBetween' } }] }];
  const { result, driver } = await run(createStore(garbledPreexisting()), vset({ type: 'date', alertStyle: 'stop', operator: 'between', min: '2024-03-13', max: '2024-11-25' }));
  assert.equal(result.ok, true, JSON.stringify(op(result).evidence));
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(op(result).status, OPERATION_STATUS.SATISFIED);
});

test('R17 RED (date, L-22d v2): the bridge/dependents axis unobservable (PATCH_VERSION_MISMATCH) refuses named, 0 dispatch', async () => {
  const { result, driver } = await run(createStore(vbook()), vset({ type: 'date', alertStyle: 'stop', operator: 'between', min: '2024-03-13', max: '2024-11-25' }), { dependencyBuild: false });
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(op(result).evidence.evidence.reason, 'xlsx-validation-rules-unobservable');
});

test('R17 RED (date, L-22d v2): the bridge validations axis itself a reason string refuses named, 0 dispatch', async () => {
  const { result, driver } = await run(createStore(vbook()), vset({ type: 'date', alertStyle: 'stop', operator: 'between', min: '2024-03-13', max: '2024-11-25' }), { validationsFault: 'UNOBSERVABLE:bridge-validations-threw' });
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(op(result).evidence.evidence.reason, 'xlsx-validation-rules-unobservable');
});

test('R17 RED (custom, L-22b): a different formula lands -> L1 FAIL, same as every other type', async () => {
  const { result } = await run(createStore(vbook()), vset({ type: 'custom', alertStyle: 'stop', formula: '=$H$2>0' }), { fault: 'validation-custom-formula-wrong' });
  assert.equal(result.ok, false);
  assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL');
});

test('R17 GREEN (custom, L-22e): a relative formula\'s readback SHIFTED from an unrelated active cell (the measured defect, simulated for real) still APPLIES -- the adapter re-anchors it, never guesses', async () => {
  const { result, m } = await run(createStore(vbook()), vset({ type: 'custom', alertStyle: 'stop', formula: '=B2>0' }), { fault: 'validation-custom-read-shifted-from-elsewhere' });
  assert.equal(result.ok, true, JSON.stringify(result.execution.primitiveVerification));
  assert.equal(m.commands, 1);
});

test('R17 GREEN (custom, L-22e): the SAME shifted-readback fault, retried -- SATISFIED, 0 dispatch (the re-anchor is deterministic across sessions)', async () => {
  const store = createStore(vbook());
  const step = vset({ type: 'custom', alertStyle: 'stop', formula: '=B2>0' });
  const first = await run(store, step, { fault: 'validation-custom-read-shifted-from-elsewhere' });
  assert.equal(first.result.ok, true);
  const retry = await run(store, step, { fault: 'validation-custom-read-shifted-from-elsewhere' });
  assert.equal(retry.result.ok, true, JSON.stringify(retry.result.execution.primitiveVerification));
  assert.equal(retry.m.commands, 0, 'SATISFIED -- the re-anchored formula still matches on a fresh read');
});

test('R17 RED (custom, L-22e): a FULLY ABSOLUTE formula is unaffected by an unrelated active cell -- the shift fault only touches relative/mixed formulas', async () => {
  const { result, m } = await run(createStore(vbook()), vset({ type: 'custom', alertStyle: 'stop', formula: '=$H$2>0' }), { fault: 'validation-custom-read-shifted-from-elsewhere' });
  assert.equal(result.ok, true, JSON.stringify(result.execution.primitiveVerification));
  assert.equal(m.commands, 1);
});

test('R17 RED (custom, L-22e): GetActiveCell() itself unobservable on a relative formula -> L1 FAIL, never guessed as matching', async () => {
  const { result } = await run(createStore(vbook()), vset({ type: 'custom', alertStyle: 'stop', formula: '=B2>0' }), { fault: 'validation-custom-active-cell-unobservable' });
  assert.equal(result.ok, false);
  assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL');
});

test('R17 RED (custom, L-22e): the SAME GetActiveCell()-unobservable fault does NOT affect an absolute formula (it never needs the active cell)', async () => {
  const { result } = await run(createStore(vbook()), vset({ type: 'custom', alertStyle: 'stop', formula: '=$H$2>0' }), { fault: 'validation-custom-active-cell-unobservable' });
  assert.equal(result.ok, true, JSON.stringify(result.execution.primitiveVerification));
});

test('R17 RED (custom, L-22e): set_validation(replace: true) with a relative formula is refused at planning, named -- the GH #96 per-cell verify was never measured with an active-cell capture', () => {
  const p = planOperations(vset({ type: 'custom', alertStyle: 'stop', formula: '=B2>0', replace: true }));
  assert.equal(p.ok, false);
  assert.match(p.detail, /does not support replace: true/);
});

test('R17 (custom, L-22e): set_validation(replace: true) with an ABSOLUTE formula is unaffected -- unchanged since L-22b/GH#96', () => {
  const p = planOperations(vset({ type: 'custom', alertStyle: 'stop', formula: '=$H$2>0', replace: true }));
  assert.equal(p.ok, true, p.detail);
});

test('R17 RED (custom, L-22b): a PRE-EXISTING RELATIVE custom rule is refused named xlsx-validation-exists-different, never silently replaced -- the SAME defect the probe measured, kept refused', async () => {
  const bookWithRelative = () => [{ name: 'V', cells: { B2: { v: 1 }, B3: { v: 2 } },
    rules: [{ kind: 'validation', type: 1, range: 'B2:B5', formula: '=A1>0', formula2: '', api: { type: 'xlValidateCustom', alertStyle: 'xlValidAlertStop', operator: 'xlBetween' } }] }];
  const { result, driver } = await run(createStore(bookWithRelative()), vset({ type: 'custom', alertStyle: 'stop', formula: '=$H$2>0' }));
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0, 'refused before any dispatch -- never silently fixed or guessed');
  assert.equal(result.execution.appliedOperations[0].evidence.evidence.reason, 'xlsx-validation-exists-different');
});

test('R17 (custom, L-22b): the editor\'s own canonicalization (a space after =, upper-cased refs) still SATISFIES a retry -- a bare string compare would wrongly see a difference', async () => {
  const store = createStore(vbook());
  const step = vset({ type: 'custom', alertStyle: 'stop', formula: '=$h$2>0' });
  const first = await run(store, step);
  assert.equal(first.result.ok, true, JSON.stringify(first.result.execution.appliedOperations[0].evidence));
  // the editor re-renders the stored formula with its own canonicalization (formula.cjs's own measured rule);
  // simulate that directly on the store, the way a real reopen would show it, and confirm a retry still matches
  store.durable[0].rules[0].formula = '= $H$2>0';
  const retry = await run(store, step);
  assert.equal(retry.result.ok, true, JSON.stringify(retry.result.execution.appliedOperations[0].evidence));
  assert.equal(retry.m.commands, 0, 'SATISFIED -- the canonical forms match, even though the raw strings do not');
});

// ---- R18 conditional formats -----------------------------------------------------------------------------------------

const cbook = () => [{ name: 'C', cells: { A1: { v: 1 }, A2: { v: 6 }, A3: { v: 9 } } }, { name: 'O', cells: { A1: { v: 'keep' } } }];
const cadd = rule => ({ id: 'm7-d-r18', steps: [{ intent: 'add_conditional_format', sheet: 'C', range: 'A1:A3', rule }] });
const cdel = rule => ({ id: 'm7-d-r18', steps: [{ intent: 'delete_conditional_format', sheet: 'C', range: 'A1:A3', rule }] });

for (const [name, rule, addArgs, fill] of [
  ['between + fill', { type: 'cellValue', operator: 'between', formula1: 2, formula2: 7, fill: '#FFC8C8' }, ['xlCellValue', 'xlBetween', '2', '7'], [255, 200, 200]],
  ['less + fill', { type: 'cellValue', operator: 'less', formula1: 3, fill: '#C8FFC8' }, ['xlCellValue', 'xlLess', '3'], [200, 255, 200]],
  ['equal, no fill', { type: 'cellValue', operator: 'equal', formula1: 6 }, ['xlCellValue', 'xlEqual', '6'], null],
  ['notBetween, no fill', { type: 'cellValue', operator: 'notBetween', formula1: 2, formula2: 7 }, ['xlCellValue', 'xlNotBetween', '2', '7'], null],
  ['expression + fill', { type: 'expression', formula: '=$A1>5', fill: '#C8C8FF' }, ['xlExpression', null, '=$A1>5'], [200, 200, 255]],
]) {
  test(`R18 ${name}: the measured Add call; add APPLIED + retry 0/0; delete APPLIED + retry 0/0`, async () => {
    const p = planOperations(cadd(rule));
    assert.equal(p.ok, true, p.detail);
    const log = [];
    const ws = { GetName: () => 'C', GetRange: () => ({ GetFormatConditions: () => ({ Add: (...x) => { log.push(['Add', ...x]); return { SetFillColor: c => log.push(['Fill', c]) }; } }) }) };
    new Function('Api', cf.lower(p.operations[0]).body)({ GetSheet: () => ws, CreateColorFromRGB: (...x) => x });
    assert.deepEqual(log, [['Add', ...addArgs], ...(fill ? [['Fill', fill]] : [])]);
    const store = createStore(cbook());
    const a = await run(store, cadd(rule));
    assert.equal(a.result.ok, true, JSON.stringify([a.result.execution.primitiveVerification, op(a.result).evidence]));
    assert.equal((await run(store, cadd(rule))).m.commands, 0);
    const d = await run(store, cdel(rule));
    assert.equal(d.result.ok, true, JSON.stringify(d.result.execution.primitiveVerification));
    assert.equal((await run(store, cdel(rule))).m.commands, 0);
  });
}

test('R18 refused: unmeasured types / operators; an expression without a fill; between without formula2', () => {
  for (const [rule, re] of [
    // M7 X-O01 widened the measured set: top10/unique/aboveAverage now have their
    // OWN shape (xlsx-m7-d-o-slice.test.cjs). X-O01 full implementation widened it
    // again: a BARE { type: 'databar' } is now ACCEPTED (default-only create) -- any appearance
    // field on it stays refused, since the criteria/enum shapes are unmeasured.
    // X-O01 (2026-10-12) measured the appearance setters: barColor plans now; an UNMEASURED value stays refused
    [{ type: 'databar', direction: 'leftToRight' }, /leftToRight is not measured/],
    [{ type: 'cellValue', operator: 'notEqual', formula1: 1 }, /operator must be one of/],
    [{ type: 'expression', formula: '=$A1>5' }, /needs a fill/],
    [{ type: 'cellValue', operator: 'between', formula1: 1 }, /needs formula2/],
    [{ type: 'cellValue', operator: 'less', formula1: 1, formula2: 2 }, /formula1 only/],
  ]) {
    const p = planOperations(cadd(rule));
    assert.equal(p.ok, false, JSON.stringify(rule));
    assert.match(p.detail, re);
  }
});

test('R18 RED: formula2 lands wrong / the fill lands wrong -> L1 FAIL', async () => {
  for (const [fault, rule] of [['cf-wrong-formula2', { type: 'cellValue', operator: 'between', formula1: 2, formula2: 7, fill: '#FFC8C8' }], ['cf-wrong-fill', { type: 'cellValue', operator: 'equal', formula1: 6 }]]) {
    const { result } = await run(createStore(cbook()), cadd(rule), { fault });
    assert.equal(result.ok, false, fault);
    assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL', fault);
  }
});
