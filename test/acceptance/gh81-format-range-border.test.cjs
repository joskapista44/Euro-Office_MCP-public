'use strict';

/**
 * GitHub #81 (owner request via the owner channel: NAD draft, A80:F87 inner cell borders): format_range is released
 * on the product surface with format.border too -- edge modes 'outline' / 'inside' / 'all' only, the four originally
 * measured styles, '#RRGGBB' or 'auto' colour, and border as the step's ONLY format key. Owner GO (paired with
 * the fill 'none' release): + mode 'remove' (X-B04, format-remove.cjs) -- the named sides gone, no style/color
 * (removeExpectation refuses them as extra keys). Owner GO (same card, unit 5): + the other 9 measured border
 * styles (X-B06, CLOSED_PROVEN: m7-b06-probe1 measured all 9, m7-p47-live2 LIVE-accepted double/hair/dashDot
 * specifically). Offline, the product WRITE tool with the fake XLSX editor (test/fakes/write-wiring.cjs).
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const { harness, req, createStore, book } = require('../fakes/write-wiring.cjs');

const bd = (range, border, id = 'bd') => ({ intent: 'format_range', id, sheet: 'Budget', range, format: { border } });

for (const mode of ['all', 'inside', 'outline']) {
  test(`GH #81: border mode '${mode}' is APPLIED with L1/L3 PASS; a retry is SATISFIED with 0 dispatch`, async () => {
    const h = harness({ store: createStore(book()) });
    const step = bd('F10:H12', { mode, style: 'thin', color: '#000000' }); // clear of the fake book's C1:D1 merge (and its ring)
    const r = await h.tool.handler(req([step]));
    assert.equal(r.ok, true, JSON.stringify(r));
    assert.deepEqual(r.operations.map(o => [o.status, o.l1, o.l3]), [['APPLIED', 'PASS', 'PASS']]);
    const d0 = h.calls().dispatch;
    const again = await h.tool.handler(req([step]));
    assert.deepEqual([again.ok, again.operations[0].status, h.calls().dispatch - d0], [true, 'SATISFIED', 0], JSON.stringify(again));
  });
}

test("owner GO: border mode 'remove' is APPLIED (no style/color), kept axes unaffected; a retry is SATISFIED with 0 dispatch", async () => {
  const h = harness({ store: createStore(book()) });
  const drawn = bd('F10:H12', { mode: 'cells', sides: ['top', 'bottom', 'left', 'right'], style: 'thin', color: '#000000' }, 'draw');
  const colored = await h.tool.handler(req([drawn]));
  assert.equal(colored.ok, true, JSON.stringify(colored));
  const removed = bd('F10:H12', { mode: 'remove', sides: ['top', 'bottom'] }, 'rm');
  const r = await h.tool.handler(req([removed]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(r.operations.map(o => [o.status, o.l1, o.l3]), [['APPLIED', 'PASS', 'PASS']]);
  const d0 = h.calls().dispatch;
  const again = await h.tool.handler(req([removed]));
  assert.deepEqual([again.ok, again.operations[0].status, h.calls().dispatch - d0], [true, 'SATISFIED', 0], JSON.stringify(again));
});

for (const style of ['double', 'hair', 'dashDotDot', 'dashDot', 'dotted', 'mediumDashDotDot', 'slantDashDot', 'mediumDashDot', 'mediumDashed']) {
  test(`owner GO: border style '${style}' (X-B06) is APPLIED on the product surface; a retry is SATISFIED with 0 dispatch`, async () => {
    const h = harness({ store: createStore(book()) });
    const step = bd('F10:H12', { mode: 'cells', sides: ['top'], style, color: '#000000' });
    const r = await h.tool.handler(req([step]));
    assert.equal(r.ok, true, JSON.stringify(r));
    assert.deepEqual(r.operations.map(o => [o.status, o.l1, o.l3]), [['APPLIED', 'PASS', 'PASS']]);
    const d0 = h.calls().dispatch;
    const again = await h.tool.handler(req([step]));
    assert.deepEqual([again.ok, again.operations[0].status, h.calls().dispatch - d0], [true, 'SATISFIED', 0], JSON.stringify(again));
  });
}

test("GH #81: border mode 'cells' with all four sides (the bounded-eligible grid) is APPLIED; a retry is SATISFIED", async () => {
  const h = harness({ store: createStore(book()) });
  const step = bd('F10:H12', { mode: 'cells', sides: ['top', 'bottom', 'left', 'right'], style: 'thin', color: '#000000' });
  const r = await h.tool.handler(req([step]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(r.operations.map(o => [o.status, o.l1, o.l3]), [['APPLIED', 'PASS', 'PASS']]);
  const d0 = h.calls().dispatch;
  const again = await h.tool.handler(req([step]));
  assert.deepEqual([again.ok, again.operations[0].status, h.calls().dispatch - d0], [true, 'SATISFIED', 0], JSON.stringify(again));
});

test("owner GO: style/color are still required for every non-'remove' mode -- enforced by the canonical planner now, not the static schema (required: ['mode','style','color'] -> ['mode'])", async () => {
  const h = harness({ store: createStore(book()) });
  for (const mode of ['cells', 'outline', 'inside', 'all']) {
    const noStyle = await h.tool.handler(req([{ intent: 'format_range', sheet: 'Budget', range: 'A1:C3', format: { border: { mode, ...(mode === 'cells' ? { sides: ['top'] } : {}), color: '#000000' } } }]));
    assert.deepEqual([noStyle.ok, noStyle.dispatched], [false, false], `${mode} without style`);
    assert.match(noStyle.reason, /border\.style must be one of/, `${mode} without style`);
    const noColor = await h.tool.handler(req([{ intent: 'format_range', sheet: 'Budget', range: 'A1:C3', format: { border: { mode, ...(mode === 'cells' ? { sides: ['top'] } : {}), style: 'thin' } } }]));
    assert.deepEqual([noColor.ok, noColor.dispatched], [false, false], `${mode} without color`);
    assert.match(noColor.reason, /border\.color must be/, `${mode} without color`);
  }
  assert.equal(h.calls().dispatch, 0);
});

test("GH #81 + GH #86 (b): every released border mode has a bounded variant ('cells' since #81, the shared-edge modes since #86 b)", () => {
  const fmt = require('../../src/adapters/xlsx/families/format.cjs');
  const variant = border => fmt.boundedExtras(fmt.plan({ intent: 'format_range', sheet: 'Budget', range: 'F10:H12', format: { border } }, 0).operation);
  assert.notEqual(variant({ mode: 'cells', sides: ['top', 'bottom', 'left', 'right'], style: 'thin', color: '#000000' }), null);
  for (const mode of ['outline', 'inside', 'all']) assert.notEqual(variant({ mode, style: 'thin', color: '#000000' }), null, mode);
  assert.notEqual(variant({ mode: 'remove', sides: ['top'] }), null, 'remove');
});

test('GH #81: border shapes outside the released subset are refused BEFORE dispatch', async () => {
  const h = harness({ store: createStore(book()) });
  const refused = [
    { numberFormat: '0%', border: { mode: 'all', style: 'thin', color: '#000000' } }, // border is the ONLY key
    { border: { mode: 'remove' } }, // 'remove' needs sides too
    { border: { mode: 'remove', sides: ['top'], style: 'thin', color: '#000000' } }, // 'remove' takes no style/color
    { border: { mode: 'cells', style: 'thin', color: '#000000' } }, // 'cells' needs sides
    { border: { style: 'thin', color: '#000000' } }, // mode is required
    { border: { mode: 'all', style: 'nonexistent-style', color: '#000000' } }, // 'double' etc. are released now (X-B06)
    { border: { mode: 'all', style: 'thin', color: 'red' } },
    { border: { mode: 'outline', sides: ['diagonalDown'], style: 'thin', color: '#000000' } },
    { border: { mode: 'inside', sides: ['top'], style: 'thin', color: '#000000' } }, // sides only with 'outline'
    { border: { mode: 'all', style: 'thin', color: '#000000', extra: 1 } },
  ];
  for (const format of refused) {
    const r = await h.tool.handler(req([{ intent: 'format_range', sheet: 'Budget', range: 'A1:C3', format }]));
    assert.deepEqual([r.ok, r.dispatched], [false, false], JSON.stringify(format) + ' ' + JSON.stringify(r));
  }
  const mixed = await h.tool.handler(req([{ intent: 'format_range', sheet: 'Budget', range: 'A1', format: refused[0] }]));
  assert.match(mixed.reason, /exactly one of numberFormat \/ border/);
  const empty = await h.tool.handler(req([{ intent: 'format_range', sheet: 'Budget', range: 'A1', format: {} }]));
  assert.deepEqual([empty.ok, empty.dispatched], [false, false]);
  assert.equal(h.calls().dispatch, 0);
});

test('GH #81: a non-released key next to border still names the release boundary', async () => {
  const h = harness({ store: createStore(book()) });
  const r = await h.tool.handler(req([{ intent: 'format_range', sheet: 'Budget', range: 'A1', format: { border: { mode: 'all', style: 'thin', color: '#000000' }, textDirection: 'rtl' } }]));
  assert.deepEqual([r.ok, r.code, r.dispatched], [false, 'REFUSED_TEMPORARY_BOUNDARY', false]);
  assert.match(r.reason, /format\.textDirection/); // L-01 released wrap -- the held-key example is an unreleased key now
});
