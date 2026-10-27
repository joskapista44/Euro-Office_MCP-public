'use strict';

/**
 * L-22b: set_validation type='custom' (ABSOLUTE-reference-only) on the PRODUCT surface (the WRITE tool with the
 * fake XLSX editor). The family-level behaviour (the measured Add() shape, the reference classifier, the
 * canonical-form compare) is in test/acceptance/xlsx-m7-d-rules.test.cjs (R17 loop + the L-22b RED cases); this
 * file is the product-tool level, mirroring l22a-validation-time-release.test.cjs. The LIVE proof is
 * test/live/release-l22b-validation-custom-product.live.cjs.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const W = require('../../src/server/write-tools.cjs');
const T = require('../../src/server/tools.cjs');
const { liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { harness, req } = require('../fakes/write-wiring.cjs');
const { createStore } = require('../fakes/fake-xlsx-driver.cjs');

const J = x => JSON.stringify(x);
const book = () => [{ name: 'Data', cells: { H2: { v: 5 } }, objects: { charts: [], pivots: [], names: [] } }];
const customStep = (over = {}) => ({ intent: 'set_validation', id: 's', sheet: 'Data', range: 'A1:A3', type: 'custom', alertStyle: 'stop', formula: '=$H$2>0', ...over });
const statuses = r => (r.operations || []).map(o => o.status);

test("L-22b: set_validation type='custom' is released -- the allowlist is unchanged (set_validation itself was already released, L-22b widens its type enum)", () => {
  assert.ok(W.XLSX_WRITE_ALLOWLIST.includes('set_validation'));
  assert.ok(T.ACCEPTED_XLSX_WRITE_ALLOWLIST.includes('set_validation'));
  assert.equal(T.writeReleaseGate().ok, true, J(T.writeReleaseGate().violations));
});

test('L-22b/e: the schema accepts an absolute, relative or mixed formula, and refuses every other shape before any dispatch', () => {
  const v = s => W.validateWriteRequest(req([s]));
  assert.equal(v(customStep()).ok, true);
  assert.equal(v(customStep({ formula: '=$H$2:$H$5>0' })).ok, true, 'an absolute range');
  // L-22e: relative and mixed references are accepted now too -- L-22b's own fully-absolute-only
  // restriction is lifted (the family's reAnchoredFormula() corrects the read-side quirk instead).
  assert.equal(v(customStep({ formula: '=H2>0' })).ok, true, 'relative');
  assert.equal(v(customStep({ formula: '=$H2>0' })).ok, true, 'mixed (column-absolute)');
  assert.equal(v(customStep({ formula: '=H$2>0' })).ok, true, 'mixed (row-absolute)');
  const bad = [
    [customStep({ formula: 'H2>0' }), 'no = prefix'],
    [customStep({ formula: '=' }), 'too short'],
    [customStep({ formula: '=H2>0', replace: true }), 'relative + replace: true -- not measured'],
    [customStep({ formula: '=Other!$H$2>0' }), 'another sheet'],
    [customStep({ formula: '=MyRange>0' }), 'a name'],
    [customStep({ operator: 'between', min: 1, max: 2 }), 'operator/min/max belong to the other types'],
    [customStep({ formula: undefined }), 'formula missing'],
    [customStep({ formula: '=$H$2>0', extra: true }), 'an unknown field'],
  ];
  for (const [s, why] of bad) assert.equal(v(s).ok, false, why);
});

test('L-22b: APPLIED with the exact absolute formula, retry SATISFIED 0 dispatch', async () => {
  const h = harness({ store: createStore(book()) });
  const first = await h.tool.handler(req([customStep()]));
  assert.deepEqual(statuses(first), ['APPLIED'], J(first).slice(0, 400));
  const d0 = h.calls().dispatch;
  const retry = await h.tool.handler(req([customStep()]));
  assert.deepEqual(statuses(retry), ['SATISFIED']);
  assert.equal(h.calls().dispatch, d0, 'the retry dispatches nothing');
});

test('L-22b: a nonexistent sheet is refused named xlsx-sheet-not-found, nothing dispatched', async () => {
  const h = harness({ store: createStore(book()) });
  const d0 = h.calls().dispatch;
  const r = await h.tool.handler(req([customStep({ sheet: 'NoSuchSheet' })]));
  assert.equal(r.ok, false, J(r).slice(0, 300));
  assert.match(J(r), /xlsx-sheet-not-found/);
  assert.equal(h.calls().dispatch, d0, 'refused before any dispatch');
});

test('L-22b: a pre-existing RELATIVE custom rule on the range is refused named xlsx-validation-exists-different, never silently replaced', async () => {
  const seeded = () => [{ name: 'Data', cells: { H2: { v: 5 } },
    objects: { charts: [], pivots: [], names: [] },
    rules: [{ kind: 'validation', type: 1, range: 'A1:A3', formula: '=A1>0', formula2: '', api: { type: 'xlValidateCustom', alertStyle: 'xlValidAlertStop', operator: 'xlBetween' } }] }];
  const h = harness({ store: createStore(seeded()) });
  const d0 = h.calls().dispatch;
  const r = await h.tool.handler(req([customStep()]));
  assert.equal(r.ok, false, J(r).slice(0, 300));
  assert.match(J(r), /xlsx-validation-exists-different/);
  assert.equal(h.calls().dispatch, d0, 'refused before any dispatch');
});

test('L-22b: the large-book path is unaffected -- set_validation stays in TIERED_INTENTS for type custom too', () => {
  const sc = s => liveScope({ id: 't', steps: [s] });
  assert.equal(sc(customStep()).tiered, true, 'unchanged from every other set_validation type');
});
