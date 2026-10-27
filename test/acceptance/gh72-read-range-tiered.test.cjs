'use strict';
// GitHub #72: xlsx_read read_range above the full-read bound read the whole workbook, truncated before the
// requested window and returned every cell `read: false`; and no cell said whether its value is text or a number.
const test = require('node:test');
const assert = require('node:assert/strict');
const { xlsxRequest } = require('../../src/server/read-tools.cjs');
const { tierOf, MAX_READ_CELLS, MAX_TIERED_CELLS } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');

const big = used => [{ name: 'S', used: 'A1:A1' }, { name: 'Far', used }];
const req = extra => xlsxRequest('read_range', { sheet: 'S', range: 'A1:B2', format: false, style: false, ...extra });

test('#72 read_range is eligible for the tiered observation: above the bound it reads the window, not the workbook', () => {
  assert.ok(MAX_READ_CELLS < 212 * 220 && 212 * 220 < MAX_TIERED_CELLS);
  for (const extra of [{}, { format: true }, { style: true }, { format: true, style: true }]) {
    const t = tierOf(big('A1:HD220'), req(extra));
    assert.equal(t.tier, 'tiered', `${JSON.stringify(extra)}: ${JSON.stringify(t)}`);
  }
});

test('#72 NEG: at or below the bound the full read runs unchanged; above the hard limit the window read alone (GH #110)', () => {
  assert.equal(tierOf(big('A1:C3'), req({ style: true })).tier, 'full');
  // GH #110 superseded '#72: above MAX_TIERED_CELLS it stays refused': the ceiling guards the digest a read never needs
  const r = tierOf(big('A1:ZZ500'), req({}));
  assert.ok(r.bound > MAX_TIERED_CELLS);
  assert.equal(r.tier, 'window');
});

test('#72 NEG: a window over the read bound is still refused on the tiered path', () => {
  const t = tierOf(big('A1:C100'), xlsxRequest('read_range', { sheet: 'S', range: 'A1:Z1800', format: false, style: false })); // 46 800 window cells > 45 000
  assert.equal(t.tier, 'refused');
  assert.match(t.reason, /task-windows-over-/);
});
