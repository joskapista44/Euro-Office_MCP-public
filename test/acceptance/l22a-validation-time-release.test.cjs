'use strict';

/**
 * L-22a: set_validation type='time' on the PRODUCT surface (the WRITE tool with the fake XLSX editor). The
 * family-level behaviour (the measured Add() shape, the readback-fraction parse, the unparseable gap) is in
 * test/acceptance/xlsx-m7-d-rules.test.cjs (R17 loop + the L-22a RED cases); this file is the product-tool level,
 * mirroring gh138-u2-print-titles-release.test.cjs / gh139-u2-move-chart-release.test.cjs. The LIVE proof is
 * test/live/release-l22a-validation-time-product.live.cjs.
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
const base = { intent: 'set_validation', id: 's', sheet: 'Data', range: 'A1:A3', type: 'time', alertStyle: 'stop' };
const timeStep = (over = {}) => ({ ...base, operator: 'between', min: '06:00', max: '18:00', ...over });
const timeGreaterStep = (over = {}) => ({ ...base, operator: 'greater', value: '12:00:30', ...over });
const statuses = r => (r.operations || []).map(o => o.status);

test("L-22a: set_validation type='time' is released -- the allowlist is unchanged (set_validation itself was already released, L-22a widens its type enum)", () => {
  assert.ok(W.XLSX_WRITE_ALLOWLIST.includes('set_validation'));
  assert.ok(T.ACCEPTED_XLSX_WRITE_ALLOWLIST.includes('set_validation'));
  assert.equal(T.writeReleaseGate().ok, true, J(T.writeReleaseGate().violations));
});

test("L-22a: the schema accepts 24-hour 'HH:MM' / 'HH:MM:SS' text for min/max/value, and refuses every other shape before any dispatch", () => {
  const v = s => W.validateWriteRequest(req([s]));
  assert.equal(v(timeStep()).ok, true);
  assert.equal(v(timeStep({ min: '06:00:30', max: '18:00:00' })).ok, true, 'seconds form');
  assert.equal(v(timeGreaterStep()).ok, true, 'one-bound operator');
  const bad = [
    [timeStep({ min: 6, max: 18 }), 'a plain number, not time text'],
    [timeStep({ min: '6:00', max: '18:00' }), 'hour not zero-padded'],
    [timeStep({ min: '6:00 AM', max: '18:00' }), 'AM/PM is a WRITE input, never accepted (24-hour text only)'],
    [timeStep({ min: '25:00', max: '18:00' }), 'an out-of-range hour'],
    [timeStep({ min: '06:60', max: '18:00' }), 'an out-of-range minute'],
    [timeStep({ min: '06:00', max: '18:00', extra: true }), 'an unknown field'],
  ];
  for (const [s, why] of bad) assert.equal(v(s).ok, false, why);
});

test('L-22a: APPLIED to the requested bounds, retry SATISFIED 0 dispatch (between)', async () => {
  const h = harness({ store: createStore(book()) });
  const first = await h.tool.handler(req([timeStep()]));
  assert.deepEqual(statuses(first), ['APPLIED'], J(first).slice(0, 400));
  const d0 = h.calls().dispatch;
  const retry = await h.tool.handler(req([timeStep()]));
  assert.deepEqual(statuses(retry), ['SATISFIED']);
  assert.equal(h.calls().dispatch, d0, 'the retry dispatches nothing');
});

test('L-22a: APPLIED + retry SATISFIED (greater, one bound)', async () => {
  const h = harness({ store: createStore(book()) });
  const step = timeGreaterStep();
  const first = await h.tool.handler(req([step]));
  assert.deepEqual(statuses(first), ['APPLIED'], J(first).slice(0, 400));
  const d0 = h.calls().dispatch;
  const retry = await h.tool.handler(req([step]));
  assert.deepEqual(statuses(retry), ['SATISFIED']);
  assert.equal(h.calls().dispatch, d0);
});

test('L-22a: a nonexistent sheet is refused named xlsx-sheet-not-found, nothing dispatched', async () => {
  const h = harness({ store: createStore(book()) });
  const d0 = h.calls().dispatch;
  const r = await h.tool.handler(req([timeStep({ sheet: 'NoSuchSheet' })]));
  assert.equal(r.ok, false, J(r).slice(0, 300));
  assert.match(J(r), /xlsx-sheet-not-found/);
  assert.equal(h.calls().dispatch, d0, 'refused before any dispatch');
});

test('L-22a: a pre-existing rule on the range (another type) is refused named xlsx-validation-exists-different, nothing dispatched', async () => {
  const seeded = () => [{ name: 'Data', cells: { A1: { v: 1 } },
    objects: { charts: [], pivots: [], names: [] },
    rules: [{ kind: 'validation', type: 1, range: 'A1:A3', formula: '1', formula2: '10', api: { type: 'xlValidateWholeNumber', alertStyle: 'xlValidAlertStop', operator: 'xlBetween' } }] }];
  const h = harness({ store: createStore(seeded()) });
  const d0 = h.calls().dispatch;
  const r = await h.tool.handler(req([timeStep()]));
  assert.equal(r.ok, false, J(r).slice(0, 300));
  assert.match(J(r), /xlsx-validation-exists-different/);
  assert.equal(h.calls().dispatch, d0, 'refused before any dispatch');
});

test('L-22a: a pre-existing TIME rule whose locale text cannot be parsed back is refused named xlsx-validation-time-unparseable, nothing dispatched', async () => {
  const seeded = () => [{ name: 'Data', cells: { A1: { v: 1 } },
    objects: { charts: [], pivots: [], names: [] },
    rules: [{ kind: 'validation', type: 1, range: 'A1:A3', formula: 'not-a-time', formula2: 'not-a-time-2', api: { type: 'xlValidateTime', alertStyle: 'xlValidAlertStop', operator: 'xlBetween' } }] }];
  const h = harness({ store: createStore(seeded()) });
  const d0 = h.calls().dispatch;
  const r = await h.tool.handler(req([timeStep()]));
  assert.equal(r.ok, false, J(r).slice(0, 300));
  assert.match(J(r), /xlsx-validation-time-unparseable/);
  assert.equal(h.calls().dispatch, d0, 'refused before any dispatch -- never guessed a match or a difference');
});

test('L-22a: the large-book path is unaffected -- set_validation stays in TIERED_INTENTS for type time too (the tiered observation is keyed by intent, not by type)', () => {
  const sc = s => liveScope({ id: 't', steps: [s] });
  assert.equal(sc(timeStep()).tiered, true, 'unchanged from every other set_validation type');
});
