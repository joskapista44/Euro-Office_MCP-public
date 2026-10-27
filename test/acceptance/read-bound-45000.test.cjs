'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const { tierOf, readQuery, MAX_READ_CELLS, MAX_TIERED_CELLS } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { validateReadRequest, xlsxRequest } = require('../../src/server/read-tools.cjs');
const { liveScope } = require('../../src/adapters/xlsx/adapter.cjs');

/**
 * Read bound 5000 -> 45000 (owner decision C; measured full format+style read: 5000 cells 0.6 s | 10000 1.2 s | 20000 2.3 s | 45000 5.0 s).
 * The tiered band is now (45 000, 50 000]; above 50 000 stays refused. MAX_FORMULA_INDEX is a DIFFERENT concept (formulas outside the windows)
 * and does not move.
 */

const sheet = cells => [{ name: 'S', used: `A1:A${cells}` }]; // one column: the bounding box is exactly `cells`
const writeTask = { id: 't', steps: [{ intent: 'write_range', id: 'w', sheet: 'S', range: 'A1', values: [['x']] }] };
const eligible = liveScope(writeTask); // a tiered family
const notMigrated = { windows: [{ sheet: 'S', range: 'A1' }] }; // no `tiered` mark: a family without a tiered path

test('the two bounds are 45 000 and 50 000 (the hard ceiling did not move)', () => {
  assert.equal(MAX_READ_CELLS, 45000);
  assert.equal(MAX_TIERED_CELLS, 50000);
});

test('eligible request: 44999 / 45000 full | 45001 / 50000 tiered | 50001 refused', () => {
  const tier = n => tierOf(sheet(n), eligible);
  assert.deepEqual([tier(44999).tier, tier(45000).tier], ['full', 'full']);
  assert.deepEqual([tier(45001).tier, tier(50000).tier], ['tiered', 'tiered']);
  assert.equal(tier(50001).tier, 'refused');
  assert.equal(tier(50001).reason, 'workbook-over-50000-cells');
});

test('a family without a tiered path: full up to 45 000 (it was refused from 5001), refused from 45 001 with its reason', () => {
  const tier = n => tierOf(sheet(n), notMigrated);
  assert.equal(tier(5001).tier, 'full', 'the old refusal point now runs the full read');
  assert.equal(tier(20000).tier, 'full');
  assert.equal(tier(45000).tier, 'full');
  const over = tier(45001);
  assert.equal(over.tier, 'refused');
  assert.match(over.reason, /family-not-migrated-to-the-tiered-observation/);
});

test('the windows cap follows the read bound: 45 000 window cells fit, 45 001 do not (tiered path)', () => {
  const wide = n => ({ ...eligible, windows: [{ sheet: 'S', range: `A1:A${n}` }] });
  const book = sheet(46000); // tiered band
  assert.equal(tierOf(book, wide(45000)).tier, 'tiered');
  const over = tierOf(book, wide(45001));
  assert.equal(over.tier, 'refused');
  assert.equal(over.reason, 'task-windows-over-45000-cells');
});

test('readQuery row limit is 45 000 (the only generated token that changed) and the formula-index ceiling stays 5000', () => {
  assert.match(readQuery([], {}), /const LIMIT = 45000;/);
  assert.doesNotMatch(readQuery([], {}), /const LIMIT = 5000;/);
  const src = fs.readFileSync(path.join(__dirname, '../../src/drivers/eurooffice/xlsx-driver.cjs'), 'utf8');
  assert.match(src, /^const MAX_FORMULA_INDEX = 5000;$/m, 'a different concept: formulas outside the windows');
});

test('read_range: a 45 000-cell range passes the pre-dispatch limit, 45 001 is refused LIMIT', () => {
  const ok = validateReadRequest('xlsx', { fileId: '1', operation: 'read_range', sheet: 'S', range: 'A1:A45000' });
  const over = validateReadRequest('xlsx', { fileId: '1', operation: 'read_range', sheet: 'S', range: 'A1:A45001' });
  assert.notEqual(ok.code, 'LIMIT', JSON.stringify(ok));
  assert.equal(over.code, 'LIMIT');
  assert.match(JSON.stringify(over), /45000/);
  assert.equal(xlsxRequest('read_range', { sheet: 'S', range: 'A1:A45000', format: false, style: false }).windows[0].range, 'A1:A45000');
});
