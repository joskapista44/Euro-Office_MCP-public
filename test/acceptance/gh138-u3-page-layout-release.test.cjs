'use strict';

/**
 * GH #138 U3: set_page_layout on the PRODUCT surface (the WRITE tool with the fake XLSX editor). The family-level behaviour is in xlsx-w8-page.test.cjs;
 * this file is the product-tool level, mirroring gh138-u1-defined-name-release.test.cjs. The READ mirror is read_page_setup's page half (L-10b); the LIVE
 * proof is test/live/release-gh138-u3-page-layout-product.live.cjs.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const W = require('../../src/server/write-tools.cjs');
const T = require('../../src/server/tools.cjs');
const { harness, req } = require('../fakes/write-wiring.cjs');
const { createStore } = require('../fakes/fake-xlsx-driver.cjs');

const J = x => JSON.stringify(x);
const book = () => [{ name: 'P', cells: { A1: { v: 1 } } }, { name: 'O', cells: { A1: { v: 'keep' } }, page: { orientation: 'xlPortrait', printGridlines: true, printHeadings: false } }];
const lay = (extra = { orientation: 'landscape', printGridlines: true }) => ({ intent: 'set_page_layout', id: 's', sheet: 'P', ...extra });
const statuses = r => (r.operations || []).map(o => o.status);

test('GH #138 U3: set_page_layout is released (allowlist + accepted set + schema + constraints), a normal step, and the release gate passes', () => {
  assert.ok(W.XLSX_WRITE_ALLOWLIST.includes('set_page_layout'));
  assert.ok(T.ACCEPTED_XLSX_WRITE_ALLOWLIST.includes('set_page_layout'));
  assert.ok(!W.SOLE_STEP.includes('set_page_layout'));
  assert.deepEqual(Object.keys(W.STEP_SCHEMAS), [...W.XLSX_WRITE_ALLOWLIST]);
  assert.deepEqual(Object.keys(W.CONSTRAINTS), [...W.XLSX_WRITE_ALLOWLIST]);
  assert.equal(T.writeReleaseGate().ok, true, J(T.writeReleaseGate().violations));
});

test('GH #138 U3: the schema + constraints + the family plan refuse a malformed step before any dispatch', () => {
  const v = s => W.validateWriteRequest(req([s]));
  assert.equal(v(lay()).ok, true);
  assert.equal(v(lay({ topMargin: 25.4, leftMargin: 0, printHeadings: false })).ok, true);
  const bad = [
    [lay({}), 'no field at all'],
    [lay({ orientation: 'sideways' }), 'an unknown orientation'],
    [lay({ topMargin: 200.5 }), 'a margin over the 200 mm cap'],
    [lay({ leftMargin: -1 }), 'a negative margin'],
    [lay({ printGridlines: 'yes' }), 'a non-boolean flag'],
    [lay({ paperSize: 'A4' }), 'a print-family field (not released, B10)'],
    [{ intent: 'set_page_layout', id: 's', orientation: 'landscape' }, 'no sheet'],
  ];
  for (const [s, why] of bad) assert.equal(v(s).ok, false, why);
});

test('GH #138 U3: APPLIED (the requested fields only, the other sheet untouched), the retry SATISFIED 0 dispatch', async () => {
  const h = harness({ store: createStore(book()) });
  const r = await h.tool.handler(req([lay({ orientation: 'landscape', printGridlines: true, topMargin: 25 })]));
  assert.equal(r.ok, true, J(r).slice(0, 500));
  assert.deepEqual(statuses(r), ['APPLIED']);
  const P = h.store.durable.find(x => x.name === 'P').page;
  assert.equal(P.orientation, 'xlLandscape');
  assert.equal(P.printGridlines, true);
  assert.equal(P.topMargin, 25);
  assert.deepEqual(h.store.durable.find(x => x.name === 'O').page, { orientation: 'xlPortrait', printGridlines: true, printHeadings: false });
  const d0 = h.calls().dispatch;
  const again = await h.tool.handler(req([lay({ orientation: 'landscape', printGridlines: true, topMargin: 25 })]));
  assert.deepEqual(statuses(again), ['SATISFIED']);
  assert.equal(h.calls().dispatch, d0, 'the retry dispatches nothing');
});

test('GH #138 U3: NEG -- a sheet that does not exist is refused named (xlsx-sheet-not-found), 0 dispatch', async () => {
  const h = harness({ store: createStore(book()) });
  const r = await h.tool.handler(req([lay({ sheet: 'Nope', orientation: 'landscape' })]));
  assert.equal(r.ok, false, J(r).slice(0, 300));
  assert.equal(h.calls().dispatch, 0);
  assert.match(J(r), /xlsx-sheet-not-found/);
});

// GH #138 U4 (2026-10-17) lifted the B10 boundary this test originally pinned -- set_print_scale /
// set_paper_size and the rest of the print family are now released (test/acceptance/
// gh138-u4-print-family-release.test.cjs). Kept here, flipped, so a future narrowing of U3's own
// scope (set_page_layout) is still distinguishable from U4's.
test('GH #138 U3: set_page_layout alone does NOT reach the rest of the print family (that is U4\'s own release, a separate step)', () => {
  assert.ok(W.XLSX_WRITE_ALLOWLIST.includes('set_page_layout'));
  assert.ok(W.XLSX_WRITE_ALLOWLIST.includes('set_print_scale'), 'U4 released this -- if it is ever reverted, set_page_layout must not silently carry it back in');
});
