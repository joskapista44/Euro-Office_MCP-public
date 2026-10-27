'use strict';

/**
 * GH #325 (G1 minimum, docs/GH302-PIVOT-CHART-DRAWING-LIVE-MATRIX-20261025.md rows C-03/D-02): the
 * PUBLIC GetWidth()/GetHeight() getters read an exact 0 x 0 for a chart or a shape after a reopen,
 * while the stored anchor is right -- `read_drawings` passed that raw 0 through unchanged (a wrong
 * value read as a real one). This unit proves the fix on the product READ surface
 * (src/server/read-tools.cjs XLSX_PROJECTIONS.read_drawings): an exact 0 x 0 now degrades width AND
 * height to 'UNOBSERVABLE:size-zero-after-reopen', never passed through as a guessed real size.
 *
 * The richer option the research doc also lists (read the saved package's own xdr:ext instead,
 * same precedent as ooxml/point-data-label-reader.cjs) needs the file's raw bytes; this read path
 * (readOnce -> driver.readLive -> this projection) never has them -- out of this unit's scope.
 *
 * No new driver reader; this file only proves the projection's own size-handling, offline, through
 * the real read_drawings handler (test/fakes/fake-xlsx-driver.cjs already simulates a zero-size
 * chart/shape -- the SAME shape its own DRAWING_FAULTS.drawing-size-lost uses for the L3 side).
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const rt = require('../../src/server/read-tools.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');
const { fakeBroker, FILE } = require('../fakes/read-editors.cjs');

const book = () => [
  {
    name: 'Sheet1',
    cells: { A1: { v: 'x', f: '' } },
    objects: {
      names: [],
      pivots: [],
      // the exact 0 x 0 the LIVE research measured on a chart after reopen (C-03)
      charts: [{ name: 'ZeroChart', w: 0, h: 0, pos: { fromCol: 0, colOffset: 0, fromRow: 0, rowOffset: 0 } }],
    },
    drawings: [
      // the SAME exact 0 x 0 on a shape after reopen (D-02)
      { kind: 'shape', name: 'ZeroShape', pos: { fromCol: 1, colOffset: 0, fromRow: 1, rowOffset: 0 }, width: 0, height: 0, text: '', strokeWidth: 0, fillLineColor: 'UNOBSERVABLE:no-public-content-getter' },
      // control: a real, non-zero shape -- untouched by the fix
      { kind: 'shape', name: 'RealShape', pos: { fromCol: 2, colOffset: 0, fromRow: 2, rowOffset: 0 }, width: 914400, height: 457200, text: '', strokeWidth: 0, fillLineColor: 'UNOBSERVABLE:no-public-content-getter' },
      // negative control: zero on ONE side only -- a bad shape, not the measured reopen limit; the
      // card's own scope (drawings.cjs's sizeOf()) never special-cases this either, so the fix must not
      {
        kind: 'shape', name: 'OneSidedZero', pos: { fromCol: 3, colOffset: 0, fromRow: 3, rowOffset: 0 }, width: 0, height: 457200, text: '', strokeWidth: 0, fillLineColor: 'UNOBSERVABLE:no-public-content-getter',
      },
    ],
  },
];

function harness() {
  const store = createStore(book());
  const b = fakeBroker('xlsx');
  const createDriver = () => createFakeXlsxDriver(store);
  return rt.createReadTool('xlsx', { configSource: b.source, loadPlaywright: () => {}, createDriver });
}

test('GH #325: an exact 0 x 0 chart or shape degrades width/height to UNOBSERVABLE, never the raw 0', async () => {
  const t = harness();
  const r = await t.handler({ fileId: FILE, operation: 'read_drawings', sheet: 'Sheet1' });
  assert.equal(r.ok, true, JSON.stringify(r));
  const byName = Object.fromEntries(r.result.sheets[0].drawings.map(d => [d.name, d]));

  assert.deepEqual([byName.ZeroChart.width, byName.ZeroChart.height], ['UNOBSERVABLE:size-zero-after-reopen', 'UNOBSERVABLE:size-zero-after-reopen']);
  assert.deepEqual([byName.ZeroShape.width, byName.ZeroShape.height], ['UNOBSERVABLE:size-zero-after-reopen', 'UNOBSERVABLE:size-zero-after-reopen']);
  // every OTHER field of the degraded records is untouched
  assert.equal(byName.ZeroChart.kind, 'chart');
  assert.equal(byName.ZeroShape.kind, 'shape');

  // a real size passes through unchanged
  assert.deepEqual([byName.RealShape.width, byName.RealShape.height], [914400, 457200]);

  // a zero on ONE side only is NOT the measured reopen limit -- stays the raw (bad) value, same as
  // drawings.cjs's own sizeOf() convention (an EXACT 0 x 0 only)
  assert.deepEqual([byName.OneSidedZero.width, byName.OneSidedZero.height], [0, 457200]);
});
