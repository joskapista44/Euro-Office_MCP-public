'use strict';

/**
 * GH #365 option C U2 review: the driver wiring (xlsx-driver.cjs's veryHiddenReads
 * axis) and the projection (sheet-state-read.cjs's visibilityOf) were both correct in isolation,
 * but `src/adapters/xlsx/inventory.cjs`'s normalizeInventory() never forwarded
 * `payload.veryHiddenReads` into the frozen `inv` object the family actually receives -- the real
 * seam between driver payload and projection. The prior tests (gh133-read-sheet-state-release.test.cjs
 * line ~94-111) called projectSheetStateRead() directly with a HAND-BUILT `inv`-shaped object,
 * skipping normalizeInventory() entirely; they could not have caught this.
 *
 * These tests go through the REAL normalizeInventory(), the same function
 * src/server/read-tools.cjs's live read path calls (readLive -> normalizeInventory -> projection,
 * see read-tools.cjs:860-865) -- not a hand-built `inv`.
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const { normalizeInventory } = require('../../src/adapters/xlsx/inventory.cjs');
const { projectSheetStateRead, HIDDEN_VS_VERYHIDDEN_DETAIL } = require('../../src/adapters/xlsx/families/sheet-state-read.cjs');

function basePayload(overrides = {}) {
  return {
    sheets: [{ name: 'Vis', index: 0, used: null }, { name: 'VHid', index: 1, used: null }],
    cells: [],
    unsupported: [],
    active: 'Vis',
    visibility: { Vis: true, VHid: false },
    ...overrides,
  };
}

test('SEAM, through the real normalizeInventory(): veryHiddenReads {VHid:true} -> visibility "veryHidden"', () => {
  const payload = basePayload({ veryHiddenReads: { VHid: true } });
  const inv = normalizeInventory(payload);
  assert.ok(inv, 'normalizeInventory must accept the payload');
  assert.deepEqual(inv.veryHiddenReads, { VHid: true }, 'the seam must forward veryHiddenReads into inv');
  const out = projectSheetStateRead(inv);
  const vhid = out.sheets.find(s => s.name === 'VHid');
  assert.deepEqual(vhid, { name: 'VHid', index: 1, visibility: 'veryHidden' });
});

test('SEAM, through the real normalizeInventory(): veryHiddenReads {VHid:false} -> visibility "hidden", NO visibilityDetail', () => {
  const payload = basePayload({ veryHiddenReads: { VHid: false } });
  const inv = normalizeInventory(payload);
  const out = projectSheetStateRead(inv);
  const vhid = out.sheets.find(s => s.name === 'VHid');
  assert.deepEqual(vhid, { name: 'VHid', index: 1, visibility: 'hidden' });
  assert.equal('visibilityDetail' in vhid, false, 'a positively-known plain-hidden sheet carries no visibilityDetail');
});

test('SEAM, through the real normalizeInventory(): veryHiddenReads as a whole-axis degrade string ("UNOBSERVABLE:CAPABILITY_MISSING") falls back to the pre-U2 shape', () => {
  const payload = basePayload({ veryHiddenReads: 'UNOBSERVABLE:CAPABILITY_MISSING' });
  const inv = normalizeInventory(payload);
  // a bare string is not an object -- the seam's own two-shape convention (same as cfReads/compiledStyleReads) means
  // the whole-axis degrade is NOT carried into inv; the projection must not see it as a boolean for any sheet.
  assert.equal(inv.veryHiddenReads, null);
  const out = projectSheetStateRead(inv);
  const vhid = out.sheets.find(s => s.name === 'VHid');
  assert.deepEqual(vhid, { name: 'VHid', index: 1, visibility: 'hidden', visibilityDetail: HIDDEN_VS_VERYHIDDEN_DETAIL });
});

test('SEAM, through the real normalizeInventory(): veryHiddenReads not asked at all (undefined) -> the old shape, byte-identical to before this unit existed', () => {
  const payload = basePayload(); // no veryHiddenReads key at all
  const inv = normalizeInventory(payload);
  assert.equal(inv.veryHiddenReads, null);
  const out = projectSheetStateRead(inv);
  assert.deepEqual(out, {
    sheets: [
      { name: 'Vis', index: 0, visibility: 'visible' },
      { name: 'VHid', index: 1, visibility: 'hidden', visibilityDetail: HIDDEN_VS_VERYHIDDEN_DETAIL },
    ],
    active: 'Vis',
  });
});

test('SEAM: a visible sheet never carries a veryHiddenReads entry at all (the driver\'s own "never asked" convention) -- the projection must not misread it', () => {
  const payload = basePayload({ veryHiddenReads: { VHid: true } }); // VHid only, Vis absent
  const inv = normalizeInventory(payload);
  const out = projectSheetStateRead(inv);
  const vis = out.sheets.find(s => s.name === 'Vis');
  assert.deepEqual(vis, { name: 'Vis', index: 0, visibility: 'visible' });
});
