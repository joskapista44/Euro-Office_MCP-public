'use strict';

/**
 * L-22d: set_validation type='date' on the PRODUCT surface (the WRITE tool with the fake XLSX editor). The family-level behaviour (the serial
 * conversion, the unambiguous-read-back rule, the named gaps) is in test/acceptance/xlsx-m7-d-rules.test.cjs; this file is the product-tool
 * level, mirroring l22a-validation-time-release.test.cjs. The analysis behind the restriction: docs/L22D-VALIDATION-DATE-ANALYSIS.md.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const W = require('../../src/server/write-tools.cjs');
const T = require('../../src/server/tools.cjs');
const { liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { harness, req } = require('../fakes/write-wiring.cjs');
const { createStore } = require('../fakes/fake-xlsx-driver.cjs');

const J = x => JSON.stringify(x);
const book = () => [{ name: 'Data', cells: { A1: { v: 1 } }, objects: { charts: [], pivots: [], names: [] } }];
const base = { intent: 'set_validation', id: 's', sheet: 'Data', range: 'A1:A3', type: 'date', alertStyle: 'stop' };
const dateStep = (over = {}) => ({ ...base, operator: 'between', min: '2024-03-13', max: '2024-11-25', ...over });
const greaterStep = (over = {}) => ({ ...base, operator: 'greater', value: '2025-01-13', ...over });
const statuses = r => (r.operations || []).map(o => o.status);

test("L-22d: set_validation type='date' is released -- the allowlist is unchanged, the type enum gains 'date'", () => {
  assert.ok(W.XLSX_WRITE_ALLOWLIST.includes('set_validation'));
  assert.equal(T.writeReleaseGate().ok, true, J(T.writeReleaseGate().violations));
  assert.ok(JSON.stringify(W.STEP_SCHEMAS.set_validation.properties.type.enum).includes('"date"'));
});

test("L-22d v2: the schema accepts ISO 'YYYY-MM-DD' text for min/max/value and refuses every other shape before any dispatch -- a locale-ambiguous date is accepted now (identity is the bridge serial)", () => {
  const v = s => W.validateWriteRequest(req([s]));
  assert.equal(v(dateStep()).ok, true);
  assert.equal(v(greaterStep()).ok, true, 'one-bound operator');
  assert.equal(v(dateStep({ min: '2024-05-05', max: '2024-10-10' })).ok, true, 'day == month, unambiguous either way');
  assert.equal(v(dateStep({ min: '2024-03-04', max: '2024-11-25' })).ok, true, 'genuinely locale-ambiguous (M/D vs D/M) -- NO LONGER refused at planning');
  assert.equal(v(greaterStep({ value: '2024-11-09' })).ok, true, 'the same, one-bound operator');
  const bad = [
    [dateStep({ min: 45364, max: 45621 }), 'a serial number, not date text'],
    [dateStep({ min: '2024-3-13', max: '2024-11-25' }), 'month not zero-padded'],
    [dateStep({ min: '03/13/2024', max: '11/25/2024' }), 'a locale format is never an input'],
    [dateStep({ min: '2024-02-30', max: '2024-11-25' }), 'not a calendar date'],
    [dateStep({ min: '2024-11-25', max: '2024-03-13' }), 'min after max'],
    [dateStep({ min: '1900-02-13', max: '2024-11-25' }), 'before 1900-03-01 (the 1900 leap-year quirk)'],
    [dateStep({ min: '06:00', max: '18:00' }), 'time text is not date text'],
    [dateStep({ min: '2024-03-13', max: '2024-11-25', extra: true }), 'an unknown field'],
  ];
  for (const [s, why] of bad) assert.equal(v(s).ok, false, why);
  const num = v({ ...base, type: 'wholeNumber', operator: 'between', min: '2024-03-13', max: '2024-11-25' });
  assert.equal(num.ok, false, 'date text is refused for every OTHER type');
});

for (const locale of ['MD', 'DM']) {
  test(`L-22d: APPLIED, retry SATISFIED 0 dispatch (between) under a ${locale} locale`, async () => {
    const h = harness({ store: createStore(book()), script: { dateLocale: locale } });
    const first = await h.tool.handler(req([dateStep()]));
    assert.deepEqual(statuses(first), ['APPLIED'], J(first).slice(0, 400));
    const d0 = h.calls().dispatch;
    const retry = await h.tool.handler(req([dateStep()]));
    assert.deepEqual(statuses(retry), ['SATISFIED']);
    assert.equal(h.calls().dispatch, d0, 'the retry dispatches nothing');
  });
}

test('L-22d: APPLIED + retry SATISFIED (greater, one bound)', async () => {
  const h = harness({ store: createStore(book()) });
  const first = await h.tool.handler(req([greaterStep()]));
  assert.deepEqual(statuses(first), ['APPLIED'], J(first).slice(0, 400));
  const d0 = h.calls().dispatch;
  assert.deepEqual(statuses(await h.tool.handler(req([greaterStep()]))), ['SATISFIED']);
  assert.equal(h.calls().dispatch, d0);
});

test('L-22d: a nonexistent sheet is refused named xlsx-sheet-not-found, nothing dispatched', async () => {
  const h = harness({ store: createStore(book()) });
  const d0 = h.calls().dispatch;
  const r = await h.tool.handler(req([dateStep({ sheet: 'NoSuchSheet' })]));
  assert.equal(r.ok, false, J(r).slice(0, 300));
  assert.match(J(r), /xlsx-sheet-not-found/);
  assert.equal(h.calls().dispatch, d0);
});

test('L-22d v2: a pre-existing DATE rule whose public text names two possible dates is SATISFIED when the bridge serial is the wanted one, 0 dispatch -- identity is the bridge, not the ambiguous text', async () => {
  const seeded = () => [{ name: 'Data', cells: { A1: { v: 1 } }, objects: { charts: [], pivots: [], names: [] },
    rules: [{ kind: 'validation', type: 1, range: 'A1:A3', formula: '3/4/2024', formula2: '11/9/2024', bridgeFormula: '45355', bridgeFormula2: '45605', api: { type: 'xlValidateDate', alertStyle: 'xlValidAlertStop', operator: 'xlBetween' } }] }];
  const h = harness({ store: createStore(seeded()) });
  const d0 = h.calls().dispatch;
  const r = await h.tool.handler(req([dateStep({ min: '2024-03-04', max: '2024-11-09' })]));
  assert.equal(r.ok, true, J(r).slice(0, 300));
  assert.deepEqual(statuses(r), ['SATISFIED']);
  assert.equal(h.calls().dispatch, d0, 'never dispatched -- the bridge already matches');
});

test('L-22d v2: a pre-existing DATE rule whose BRIDGE serial disagrees with the wanted one is refused xlsx-validation-exists-different, 0 dispatch', async () => {
  const seeded = () => [{ name: 'Data', cells: { A1: { v: 1 } }, objects: { charts: [], pivots: [], names: [] },
    rules: [{ kind: 'validation', type: 1, range: 'A1:A3', formula: 'not-a-date', formula2: 'not-a-date-2', bridgeFormula: '1', bridgeFormula2: '2', api: { type: 'xlValidateDate', alertStyle: 'xlValidAlertStop', operator: 'xlBetween' } }] }];
  const h = harness({ store: createStore(seeded()) });
  const d0 = h.calls().dispatch;
  const r = await h.tool.handler(req([dateStep()]));
  assert.equal(r.ok, false, J(r).slice(0, 300));
  assert.match(J(r), /xlsx-validation-exists-different/);
  assert.equal(h.calls().dispatch, d0, 'never guessed a match or a difference');
});

test('L-22d: the large-book path is unaffected -- set_validation stays in TIERED_INTENTS for type date too', () => {
  assert.equal(liveScope({ id: 't', steps: [dateStep()] }).tiered, true);
});
