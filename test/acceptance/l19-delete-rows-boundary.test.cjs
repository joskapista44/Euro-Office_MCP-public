'use strict';

/**
 * L-19 (kiosztó: zero, kártya an internal branch): a leltár L-19 sora
 * "Sor torlese: max 100 sor, csak onallo lepeskent"-et mond. MERT ALLAPOT a kodban (main 7215aba7):
 * write-tools.cjs ROWS-schema `count: { minimum:1, maximum:100 }` ÉS delete_rows a SOLE_STEP listájában
 * -- EZ A FÁJL ezt a KÉT, SÉMA-SZINTŰ határt pinneli a produktum write-tool-on (validateWriteRequest),
 * nem kitalált számmal, hanem a séma SAJÁT maximumával.
 *
 * Ami EBBŐL NINCS MÉRVE, és amit ez a fájl NEM old fel: hogy a DRIVER mekkora SHEET-aggregátumot bír
 * delete_rows alatt (insert_rows/delete_rows -- a `structure` család -- NEM TAGJA a TIERED_INTENTS-nek,
 * adapter.cjs; ugyanaz az osztály, amit L-11/3 a filter_range-nél pinnelt), és mennyi ideig tart egy
 * nagy lapon. Az a LIVE proof: test/live/release-l19-delete-rows-boundary-product.live.cjs (zero futtatja).
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const { validateWriteRequest } = require('../../src/server/write-tools.cjs');
const { liveScope } = require('../../src/adapters/xlsx/adapter.cjs');

const J = x => JSON.stringify(x);
const FILE = '1303626';
const req = steps => ({ fileId: FILE, task: { id: 't', steps } });
const del = (extra = {}) => ({ intent: 'delete_rows', id: 's', sheet: 'S', at: 10, count: 100, anchor: { column: 'A', value: 'anchor' }, ...extra });

test('count = 100 (the schema\'s own ceiling): validateWriteRequest lets it through', () => {
  const r = validateWriteRequest(req([del()]));
  assert.equal(r.ok, true, J(r));
});

test('count = 101 (one over the schema ceiling): named REFUSED_SHAPE, the schema\'s own maximum, not a guessed number', () => {
  const r = validateWriteRequest(req([del({ count: 101 })]));
  assert.equal(r.ok, false);
  assert.equal(r.code, 'REFUSED_SHAPE');
  assert.match(r.reason, /count/);
  assert.match(r.reason, /out of range/);
});

test('count = 0 and a fractional count are refused the same way (the schema minimum / integer check, not just the ceiling)', () => {
  for (const count of [0, 50.5]) {
    const r = validateWriteRequest(req([del({ count })]));
    assert.equal(r.ok, false, `count=${count}`);
    assert.equal(r.code, 'REFUSED_SHAPE', `count=${count}`);
  }
});

test('delete_rows next to another (non-delete_rows) step: refused, naming delete_rows -- L-19 C (2026-10-17) narrowed the rule to "delete_rows may repeat with itself", so the wording changed from the old blanket SOLE_STEP message; the refusal itself did not', () => {
  const other = { intent: 'write_range', id: 'w', sheet: 'S', range: 'A1', values: [['x']] };
  const r = validateWriteRequest(req([del(), other]));
  assert.equal(r.ok, false);
  assert.equal(r.code, 'REFUSED_COMPOSITION');
  assert.match(r.reason, /delete_rows may only share a task with other delete_rows steps/);
  const r2 = validateWriteRequest(req([other, del()]));
  assert.equal(r2.ok, false);
  assert.equal(r2.code, 'REFUSED_COMPOSITION');
});

test('delete_rows (structure family) stays outside the tiered observation -- the same class L-11/3 pinned for filter_range (adapter.cjs TIERED_INTENTS)', () => {
  assert.notEqual(liveScope({ id: 't', steps: [del()] }).tiered, true);
});
