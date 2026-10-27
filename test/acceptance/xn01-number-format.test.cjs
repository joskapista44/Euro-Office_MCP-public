'use strict';

/**
 * X-N01 (owner decision, owner channel 2026-10-12: number format is a CAPABILITY; value and number format are separate
 * semantics -- value 0.1 + a percentage format -> 10% after reopen; the value is never turned into 10): the canonical
 * format_range is released on the product surface with format.numberFormat ONLY. Offline, the product WRITE tool with
 * the fake XLSX editor (test/fakes/write-wiring.cjs).
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const { harness, req, createStore, book } = require('../fakes/write-wiring.cjs');

const nf = (range, numberFormat, id = 'nf') => ({ intent: 'format_range', id, sheet: 'Budget', range, format: { numberFormat } });
const durable = (h, a) => h.store.durable.find(x => x.name === 'Budget').cells[a];

test('X-N01: a number format on an existing number: the value stays, the format is set; a retry is SATISFIED with 0 dispatch', async () => {
  const h = harness({ store: createStore(book()) });
  const before = durable(h, 'B2').v;
  const r = await h.tool.handler(req([nf('B2', '0%')]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(r.operations.map(o => [o.status, o.l1, o.l3]), [['APPLIED', 'PASS', 'PASS']]);
  assert.deepEqual([durable(h, 'B2').v, durable(h, 'B2').nf], [before, '0%']);
  const d0 = h.calls().dispatch;
  const again = await h.tool.handler(req([nf('B2', '0%')]));
  assert.deepEqual([again.ok, again.operations[0].status, h.calls().dispatch - d0], [true, 'SATISFIED', 0]);
});

test('X-N01: value 0.1 and its percentage format in ONE task (separate steps, separate semantics): stored 0.1 + 0%', async () => {
  const h = harness({ store: createStore(book()) });
  const r = await h.tool.handler(req([{ intent: 'write_range', id: 'v', sheet: 'Budget', range: 'G1', values: [[0.1]] }, nf('G1', '0%')]));
  assert.equal(r.ok, true, JSON.stringify(r));
  assert.deepEqual(r.operations.map(o => o.status), ['APPLIED', 'APPLIED']);
  assert.deepEqual([durable(h, 'G1').v, durable(h, 'G1').nf], [0.1, '0%'], 'the value is NOT rewritten to 10');
});

test('X-N01: any other format key is refused BEFORE dispatch with the release-boundary reason; the shape is strict', async () => {
  // GitHub #88/#89/#93 released fill, font and align, L-01 wrap -- every OTHER key is still behind the boundary
  const h = harness({ store: createStore(book()) });
  for (const format of [{ textDirection: 'rtl' }]) {
    const r = await h.tool.handler(req([{ intent: 'format_range', sheet: 'Budget', range: 'B2', format }]));
    assert.deepEqual([r.ok, r.stage, r.code, r.dispatched], [false, 'PRE_DISPATCH', 'REFUSED_TEMPORARY_BOUNDARY', false], JSON.stringify(format));
    assert.match(r.reason, /released with format\.numberFormat, format\.border, format\.fill, format\.font, format\.align and format\.wrap only/);
  }
  // two released keys together is a DIFFERENT refusal (the one-key-per-step rule, REFUSED_SHAPE, not a boundary)
  const twoReleased = await h.tool.handler(req([{ intent: 'format_range', sheet: 'Budget', range: 'B2', format: { numberFormat: '0%', font: { bold: true } } }]));
  assert.deepEqual([twoReleased.ok, twoReleased.code, twoReleased.dispatched], [false, 'REFUSED_SHAPE', false]);
  assert.match(twoReleased.reason, /exactly one of numberFormat \/ border \/ fill \/ font \/ align/);
  for (const bad of [{ intent: 'format_range', sheet: 'Budget', range: 'B2' }, { intent: 'format_range', sheet: 'Budget', range: 'B2', format: { numberFormat: '' } }, { intent: 'format_range', sheet: 'Budget', range: 'B2', format: { font: {} } }, { intent: 'format_range', sheet: 'Budget', range: 'B2', format: { align: {} } }, { intent: 'format_range', sheet: 'Budget', range: 'B2', format: { align: { direction: 'rtl' } } }]) {
    const r = await h.tool.handler(req([bad]));
    assert.deepEqual([r.ok, r.dispatched], [false, false], JSON.stringify(bad));
  }
  assert.equal(h.calls().dispatch, 0);
});
