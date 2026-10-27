'use strict';

/**
 * GH #302 G7: set_shape_rotation -- one named shape's rotation through the public SetRotation, verified by read_drawings'
 * rotation within the one measured 1/60000-degree truncation unit (docs/evidence/gh302-g7-shape-rotation-probe-20261026.json:
 * 30 -> 29.999983333333333, 45 exact, -30 stored negative, 400 stored past a full turn).
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const W = require('../../src/server/write-tools.cjs');
const { planOperations } = require('../../src/adapters/xlsx/adapter.cjs');
const fam = require('../../src/adapters/xlsx/families/shape-rotation.cjs');
const { req } = require('../fakes/write-wiring.cjs');

const shapeRaw = (name, rotation, extra = {}) => ({ kind: 'shape', name, parent: 'S', pos: { fromCol: 1, colOffset: 0, fromRow: 1, rowOffset: 0 }, width: 914400, height: 457200, rotation, flipH: false, flipV: false, text: '', strokeWidth: 0, fillLineColor: 'UNOBSERVABLE:x', ...extra });
const obs = list => ({ format: 'xlsx', complete: true, sheets: [{ name: 'S', index: 0, used: null }], cells: {}, drawingReads: { S: list } });
const stepOf = (extra = {}) => ({ intent: 'set_shape_rotation', sheet: 'S', name: 'Box 1', rotation: 30, ...extra });
const opOf = s => { const p = fam.plan(s, 0); assert.equal(p.ok, true, p.detail); return p.operation; };

test('G7 sameRotation: the measured truncation (one 1/60000 unit) passes, two units do not', () => {
  assert.equal(fam.sameRotation(29.999983333333333, 30), true, 'the measured 30 read');
  assert.equal(fam.sameRotation(358.9999833333333, 359), true, 'the measured 359 read');
  assert.equal(fam.sameRotation(45, 45), true);
  assert.equal(fam.sameRotation(30 - 2 / 60000, 30), false);
  assert.equal(fam.sameRotation(31, 30), false);
  assert.equal(fam.sameRotation('UNOBSERVABLE:x', 30), false);
});

test('G7 schema + plan: 0 <= rotation < 360; negative / full turn / non-number / extra field refused', () => {
  assert.equal(W.validateWriteRequest(req([stepOf()])).ok, true);
  assert.equal(W.validateWriteRequest(req([stepOf({ rotation: 0 })])).ok, true);
  assert.equal(W.validateWriteRequest(req([stepOf({ rotation: 12.5 })])).ok, true);
  for (const rotation of [-30, 360, 400, '30']) {
    assert.equal(fam.plan(stepOf({ rotation }), 0).ok, false, JSON.stringify(rotation));
    assert.equal(W.validateWriteRequest(req([stepOf({ rotation })])).ok, false, JSON.stringify(rotation));
  }
  assert.equal(W.validateWriteRequest(req([stepOf({ flip: true })])).ok, false);
});

test('G7 classify: another rotation -> APPLY; this one (within the unit) -> SATISFIED; no such shape / an image -> named FAIL_CLOSED', () => {
  const op = opOf(stepOf());
  assert.equal(fam.classify(obs([shapeRaw('Box 1', 0)]), op).state, 'APPLY');
  assert.equal(fam.classify(obs([shapeRaw('Box 1', 29.999983333333333)]), op).state, 'SATISFIED');
  assert.equal(fam.classify(obs([shapeRaw('Other', 0)]), op).evidence.reason, 'xlsx-shape-rotation-no-such-shape');
  const image = { kind: 'image', name: 'Box 1', parent: 'S', pos: { fromCol: 1, colOffset: 0, fromRow: 1, rowOffset: 0 }, width: 914400, height: 457200, rotation: 0, flipH: false, flipV: false };
  assert.equal(fam.classify(obs([image]), op).evidence.reason, 'xlsx-shape-rotation-not-a-shape');
  assert.equal(fam.classify(obs([shapeRaw('Box 1', 'UNOBSERVABLE:threw')]), op).state, 'FAIL_CLOSED', 'an unreadable rotation is a drawing gap');
});

test('G7 lower: one shape of the name, SetRotation with the requested degrees', () => {
  const body = fam.lower(opOf(stepOf({ rotation: 12.5 }))).body;
  assert.match(body, /shape-identity/);
  assert.match(body, /shapes\[0\]\.SetRotation\(12\.5\) === false/);
});

test('G7 verify: L1 the read rotation within the unit; L3 keeps every other drawing and this one\'s other fields', () => {
  const op = opOf(stepOf());
  const other = shapeRaw('Other', 0, { pos: { fromCol: 5, colOffset: 0, fromRow: 5, rowOffset: 0 } });
  const before = { primary: obs([shapeRaw('Box 1', 0), other]) };
  const good = obs([shapeRaw('Box 1', 29.999983333333333), other]);
  assert.equal(fam.verifyPrimitive(op, before, { primary: good }).ok, true);
  assert.equal(fam.verifyPrimitive(op, before, { primary: obs([shapeRaw('Box 1', 0), other]) }).ok, false, 'not rotated');
  assert.equal(fam.verifyPreservation(op, before, { primary: good }).ok, true);
  assert.equal(fam.verifyPreservation(op, before, { primary: obs([shapeRaw('Box 1', 29.999983333333333), shapeRaw('Other', 10, { pos: other.pos })]) }).ok, false, 'another shape rotated');
  assert.equal(fam.verifyPreservation(op, before, { primary: obs([shapeRaw('Box 1', 29.999983333333333, { width: 900000 }), other]) }).ok, false, 'this shape changed more than its rotation');
});

test('G7 routing: set_shape_rotation plans through the shape-rotation family', () => {
  const p = planOperations({ id: 't', steps: [stepOf()] });
  assert.equal(p.ok, true, p.detail);
  assert.equal(p.operations[0].family, 'shape-rotation');
});
