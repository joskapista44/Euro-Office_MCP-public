'use strict';

/**
 * GH #110 (the reporter's LIVE workbook, 1.0.1-06bb9c52b554): xlsx_read read_range on a 14-cell window came back
 * complete:false, every cell read:false, no reason. Another sheet of the workbook reported a full-width used range
 * (A1:XFC172, ~2.8 M cells); tierOf sums every sheet into the bound, so the read was refused (workbook-over-50000-cells)
 * and fell back to the full read, which truncated before the window. The 50 000-cell ceiling guards the T2 digest scan;
 * a read asserts no preservation and needs no digest. A read-only window request above the ceiling is now served by the
 * window read alone (tier 'window', window cells still capped at MAX_READ_CELLS); and a read that is not complete says why.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const { xlsxRequest, createReadTool } = require('../../src/server/read-tools.cjs');
const { tierOf, readQuery, MAX_TIERED_CELLS } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { normalizeInventory } = require('../../src/adapters/xlsx/inventory.cjs');
const { liveScope } = require('../../src/adapters/xlsx/adapter.cjs');

// the reporter's shape: the window's own sheet is small, ANOTHER sheet is full-width
const book = [{ name: 'Elszamolas', used: 'A1:X60' }, { name: 'Ajanlat', used: 'A1:XFC172' }];
const rr = (range, extra = {}) => xlsxRequest('read_range', { sheet: 'Elszamolas', range, format: false, style: false, ...extra });

test('the reporter shape: a 14-cell window is no longer refused -- tier window, bound named', () => {
  const t = tierOf(book, rr('U17:U30'));
  assert.ok(t.bound > MAX_TIERED_CELLS, 'the control: the book IS over the ceiling');
  assert.deepEqual([t.tier, t.windowCells], ['window', 14]);
});

test('NEG: a window over MAX_READ_CELLS is still refused on a book over the ceiling', () => {
  const t = tierOf(book, rr('A1:Z1800'));
  assert.equal(t.tier, 'refused');
  assert.match(t.reason, /workbook-over-50000-cells/);
});

test('NEG: a WRITE task on the same book keeps the ceiling (the window tier is for read-only requests only)', () => {
  const task = { id: 't', steps: [{ intent: 'write_range', id: 'w1', sheet: 'Elszamolas', range: 'U17', values: [['x']] }] };
  const t = tierOf(book, liveScope(task));
  assert.deepEqual([t.tier, t.reason], ['refused', 'workbook-over-50000-cells']);
});

test('NEG: below the ceiling read_range keeps the tiered path (GitHub #72), below the bound the full read', () => {
  assert.equal(tierOf([{ name: 'Elszamolas', used: 'A1:HD220' }], rr('U17:U30')).tier, 'tiered');
  assert.equal(tierOf([{ name: 'Elszamolas', used: 'A1:C3' }], rr('A1:B2')).tier, 'full');
});

test('a window-tier payload normalizes complete (no digest is expected of it)', () => {
  const inv = normalizeInventory({ sheets: book.map((s, index) => ({ ...s, index })), cells: [{ sheet: 'Elszamolas', a: 'U17', v: '5', f: '' }], truncated: false, unsupported: [], objects: {}, tier: 'window', bound: 2819316 });
  assert.equal(inv.complete, true);
  assert.equal(inv.tier, 'window');
});

// ---- the read tool end to end, on a stand-in driver that runs the driver's own tier decision ---------------------

function standIn(payloadFor) {
  return { open: async () => ({ ok: true, handle: {} }), readLive: async (_h, request) => ({ ok: true, payload: payloadFor(request) }), close: async () => ({ ok: true }) };
}
const toolOn = payloadFor => createReadTool('xlsx', { configSource: {}, runtime: { check: () => ({ ok: true }), load: () => null }, createDriver: () => standIn(payloadFor) });

test('an incomplete read says why: the limit note reaches the read_range response', async () => {
  const unsupported = [{ field: 'observation', reason: 'IMPLEMENTATION_LIMIT:workbook-over-50000-cells (bound 2819316)' }];
  const tool = toolOn(() => ({ sheets: book.map((s, index) => ({ ...s, index })), cells: [], truncated: true, unsupported, objects: {} }));
  const out = await tool.handler({ fileId: '1186748', operation: 'read_range', sheet: 'Elszamolas', range: 'U17:U18' });
  assert.equal(out.complete, false);
  assert.deepEqual(out.incomplete, ['truncated', 'observation: IMPLEMENTATION_LIMIT:workbook-over-50000-cells (bound 2819316)']);
});

test('NEG: a complete read carries no incomplete field', async () => {
  const tool = toolOn(() => ({ sheets: [{ name: 'Elszamolas', index: 0, used: 'A1:X60' }], cells: [{ sheet: 'Elszamolas', a: 'U17', v: '5', f: '' }], truncated: false, unsupported: [], objects: {} }));
  const out = await tool.handler({ fileId: '1186748', operation: 'read_range', sheet: 'Elszamolas', range: 'U17' });
  assert.equal(out.complete, true);
  assert.equal('incomplete' in out, false);
});

test('the window read query reads windows only (usedRanges:false) -- the T1 body the window tier runs', () => {
  assert.equal(typeof readQuery([{ sheet: 'Elszamolas', range: 'U17:U30' }], { usedRanges: false }), 'string');
});
