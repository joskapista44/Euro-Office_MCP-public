'use strict';

/**
 * GH #86 (b) (owner GO): the shared-edge border modes (outline / inside / all / remove) get a
 * bounded variant above the 5000-cell read bound. The proof under BOUNDED: every ring cell across a shared edge is in
 * the target + ring window, read with style -- so the shared-edge check (sharedEdgeMismatch) still sees it, and a
 * drift on a ring cell still fails L3. Direct verifyPreservation on hand-built BOUNDED observations (the offline
 * stand-in has no SetBorders; the LIVE harness proves the editor side).
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const formatFam = require('../../src/adapters/xlsx/families/format.cjs');
const { planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { tierOf } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');

const opOf = format => planOperations({ id: 't', steps: [{ intent: 'format_range', id: 'f1', sheet: 'S', range: 'B5', format }] }).operations[0];
const NONE = { s: 0, c: 0 };
const THIN = c => ({ s: 7, c }); // 'thin' = style code 7 (format-maps STYLE)
const RING = ['A4', 'B4', 'C4', 'A5', 'B5', 'C5', 'A6', 'B6', 'C6'];
const cell = (border = {}) => ({ f: '', v: 0, wrap: false, border: { t: NONE, b: NONE, l: NONE, r: NONE, ...border } });
const obs = cells => ({ primary: { format: 'xlsx', complete: true, sheets: [{ name: 'S', used: 'A1:J4600' }], tier: 'tiered', digests: [], bounded: true,
  cells: Object.fromEntries(Object.entries(cells).map(([a, c]) => [`S\u0000${a}`, c])) } });
const grid = over => Object.fromEntries(RING.map(a => [a, cell((over || {})[a])]));

test('eligibility: every shared-edge mode has a bounded variant, and the tier above 45000 is tiered+bounded', () => {
  for (const border of [{ mode: 'outline', style: 'thin', color: '#FF0000' }, { mode: 'inside', style: 'thin', color: '#FF0000' }, { mode: 'all', style: 'thin', color: '#FF0000' }, { mode: 'remove', sides: ['top'] }]) {
    const steps = [{ intent: 'format_range', id: 'f1', sheet: 'S', range: 'B5:C6', format: { border } }];
    const t = tierOf([{ name: 'S', used: 'A1:J4600' }], liveScope({ id: 't', steps }));
    assert.deepEqual([t.tier, t.bounded], ['tiered', true], border.mode);
  }
});

// B5 outline in red (#FF0000 = 16711680): the four ring cells across B5's edges show the facing side
const RED = 16711680;
const outlined = (over = {}) => grid({ B5: { t: THIN(RED), b: THIN(RED), l: THIN(RED), r: THIN(RED) }, B4: { b: THIN(RED) }, B6: { t: THIN(RED) }, A5: { r: THIN(RED) }, C5: { l: THIN(RED) }, ...over });

test('outline under BOUNDED: the shared ring edges hold -> L3 PASS, scope BOUNDED, the ring compared with style', () => {
  const r = formatFam.verifyPreservation(opOf({ border: { mode: 'outline', style: 'thin', color: '#FF0000' } }), obs(grid()), obs(outlined()));
  assert.equal(r.ok, true, JSON.stringify(r.evidence));
  assert.equal(r.evidence.preservationScope, 'BOUNDED');
  assert.ok(r.evidence.styleCompared >= 8, `ring compared (${r.evidence.styleCompared})`);
});

test('SABOTAGE outline under BOUNDED: a ring cell missing its shared edge -> L3 FAIL, named', () => {
  const after = outlined({ B4: {} });
  const r = formatFam.verifyPreservation(opOf({ border: { mode: 'outline', style: 'thin', color: '#FF0000' } }), obs(grid()), obs(after));
  assert.equal(r.ok, false);
  assert.match(JSON.stringify(r.evidence), /B4/);
});

test('SABOTAGE outline under BOUNDED: a CORNER ring cell (no shared edge) gains a border -> L3 FAIL (drift)', () => {
  const r = formatFam.verifyPreservation(opOf({ border: { mode: 'outline', style: 'thin', color: '#FF0000' } }), obs(grid()), obs(outlined({ A4: { r: THIN(RED) } })));
  assert.equal(r.ok, false);
  assert.match(JSON.stringify(r.evidence), /A4/);
});

// remove all four sides of B5: before both cells of every edge carry a thin black side, after none
const boxed = () => grid({ B5: { t: THIN(0), b: THIN(0), l: THIN(0), r: THIN(0) }, B4: { b: THIN(0) }, B6: { t: THIN(0) }, A5: { r: THIN(0) }, C5: { l: THIN(0) } });
const REMOVE = { border: { mode: 'remove', sides: ['top', 'bottom', 'left', 'right'] } };

test('remove under BOUNDED: both cells of every removed edge cleared -> L3 PASS, scope BOUNDED', () => {
  const r = formatFam.verifyPreservation(opOf(REMOVE), obs(boxed()), obs(grid()));
  assert.equal(r.ok, true, JSON.stringify(r.evidence));
  assert.equal(r.evidence.preservationScope, 'BOUNDED');
});

test('SABOTAGE remove under BOUNDED: the neighbour keeps its half of the edge -> L3 FAIL, named', () => {
  const r = formatFam.verifyPreservation(opOf(REMOVE), obs(boxed()), obs(grid({ C5: { l: THIN(0) } })));
  assert.equal(r.ok, false);
  assert.match(JSON.stringify(r.evidence), /C5/);
});
