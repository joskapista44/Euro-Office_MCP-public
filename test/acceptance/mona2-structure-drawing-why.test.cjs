'use strict';

/**
 * alice's LIVE delete_rows (2026-11-01, sheet with one logo image at A1:B5, cut at row 29) refused
 * `xlsx-structure-drawing-geometry-unobservable` with no hint WHICH drawing or WHY: structure.cjs drawingGuard
 * carries [{ name, why }] on `drawings`, failureOf() passed scalars only. Same root cause as GH #96 / L-18/1,
 * a third shape. Carried on failure.drawings (classification), bounded: 10 entries, name / why only, why <= 160.
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const W = require('../../src/server/write-tools.cjs');

const ev = (reason, drawings) => ({ status: 'FAILED', evidence: { stage: 'classification', evidence: { reason, drawings, count: Array.isArray(drawings) ? drawings.length : 0 } } });

test('mona2: failureOf names the drawing and its why on a structure drawing refusal', () => {
  const f = W.failureOf(ev('xlsx-structure-drawing-geometry-unobservable', [{ name: 'Moneta Hitel logo', why: 'anchor-mode-not-measured:0' }]), { l1: null, l3: null });
  assert.equal(f.stage, 'CLASSIFICATION');
  assert.equal(f.reason, 'xlsx-structure-drawing-geometry-unobservable');
  assert.deepEqual(f.drawings, [{ name: 'Moneta Hitel logo', why: 'anchor-mode-not-measured:0' }]);
});

test('mona2: a drawing entry without why (at-or-after-the-cut) keeps its name only', () => {
  const f = W.failureOf(ev('xlsx-structure-drawing-at-or-after-the-cut', [{ name: 'Chart 1', lastRow: 40, lastCol: 5 }]), { l1: null, l3: null });
  assert.deepEqual(f.drawings, [{ name: 'Chart 1' }]);
});

test('mona2: the carried list is bounded and only scalar name / why survive', () => {
  const many = Array.from({ length: 25 }, (_, i) => ({ name: `d${i}`, why: 'x'.repeat(500), extra: { nested: true } }));
  const f = W.failureOf(ev('xlsx-structure-drawing-geometry-unobservable', many), { l1: null, l3: null });
  assert.equal(f.drawings.length, 10);
  assert.equal(f.drawings[0].why.length, 160);
  assert.deepEqual(Object.keys(f.drawings[0]).sort(), ['name', 'why']);
  assert.equal(f.drawingsTotal, 25);
});

test('mona2: no drawings key when the list is absent, empty or not an array', () => {
  for (const d of [undefined, [], 'x', { name: 'a' }]) {
    const f = W.failureOf(ev('xlsx-structure-drawing-geometry-unobservable', d), { l1: null, l3: null });
    assert.equal('drawings' in f, false, JSON.stringify(d));
  }
});
