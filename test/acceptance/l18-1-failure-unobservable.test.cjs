'use strict';

/**
 * L-18/1 review (a maintainer's LIVE run on the prod DS, card
 * an internal branch): delete_cells's MERGE_IN scenario
 * refused `xlsx-delete-cells-dependency-observation-incomplete` with no hint WHICH axis was unobservable --
 * the classifier's own evidence carries `dependencies.cjs`'s `unobservable` list ([{ axis, why, host }]),
 * but failureOf() only passed scalars, so the product answer dropped it, same root cause as GH #96's
 * `unsupported` list, a different shape. Not scoped to delete-cells.cjs alone: charts.cjs's own
 * drawing/chart-source guards (`xlsx-charts-drawings-unobservable`) carry the same shape and were silently
 * dropped too. The list is now carried on failure.unobservable (classification only, the one stage both
 * families use it from), bounded: at most 10 entries, each `why` cut to 160 characters.
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const W = require('../../src/server/write-tools.cjs');

const ev = (reason, unobservable) => ({ status: 'FAILED', evidence: { stage: 'classification', evidence: { reason, unobservable } } });

test('L-18/1: failureOf carries the dependencies/rules unobservable list on a classification refusal', () => {
  const f = W.failureOf(ev('xlsx-delete-cells-dependency-observation-incomplete', [{ axis: 'defined-names', why: 'UNOBSERVABLE:names-reader-absent', host: null }]), { l1: null, l3: null });
  assert.equal(f.reason, 'xlsx-delete-cells-dependency-observation-incomplete');
  assert.deepEqual(f.unobservable, [{ axis: 'defined-names', why: 'UNOBSERVABLE:names-reader-absent' }]);
});

test('L-18/1: the host field survives when present (a per-sheet gap, not a workbook-wide one)', () => {
  const f = W.failureOf(ev('xlsx-charts-drawings-unobservable', [{ axis: 'objects:Sheet2', why: 'UNOBSERVABLE:objects-not-read', host: 'Sheet2' }]), { l1: null, l3: null });
  assert.deepEqual(f.unobservable, [{ axis: 'objects:Sheet2', why: 'UNOBSERVABLE:objects-not-read', host: 'Sheet2' }]);
});

test('L-18/1: the carried list is bounded (10 entries, 160-char why) and only scalar axis/why/host survive', () => {
  const many = Array.from({ length: 25 }, (_, i) => ({ axis: `axis${i}`, why: 'x'.repeat(500), host: null, extra: { nested: true } }));
  const f = W.failureOf(ev('xlsx-delete-cells-dependency-observation-incomplete', many), { l1: null, l3: null });
  assert.equal(f.unobservable.length, 10);
  assert.equal(f.unobservable[0].why.length, 160);
  assert.deepEqual(Object.keys(f.unobservable[0]).sort(), ['axis', 'why']);
  assert.equal(f.unobservableTotal, 25);
});

test('L-18/1: no unobservable key when the list is absent, empty or not an array', () => {
  for (const u of [undefined, [], 'UNOBSERVABLE:x', { axis: 'a' }]) {
    const f = W.failureOf(ev('xlsx-delete-cells-dependency-observation-incomplete', u), { l1: null, l3: null });
    assert.equal('unobservable' in f, false, JSON.stringify(u));
  }
});
