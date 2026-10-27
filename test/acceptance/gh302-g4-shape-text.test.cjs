'use strict';

/**
 * GH #302 G4: shape text WRITE -- insert_shape's optional `text` and set_shape_text on an existing named shape. The
 * write is the public GetContent() paragraph path, the verifier read_drawings' text (drawing.text.read/v1) in the
 * measured "p\r\n|q\r\n" shape (docs/evidence/gh302-g4-shape-text-probe-20261026.json).
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const W = require('../../src/server/write-tools.cjs');
const { planOperations } = require('../../src/adapters/xlsx/adapter.cjs');
const shapeText = require('../../src/adapters/xlsx/shape-text.cjs');
const fam = require('../../src/adapters/xlsx/families/shape-text.cjs');
const ins = require('../../src/adapters/xlsx/families/drawings-insert.cjs');
const { req } = require('../fakes/write-wiring.cjs');

const shapeRaw = (name, text, extra = {}) => ({ kind: 'shape', name, parent: 'S', pos: { fromCol: 1, colOffset: 0, fromRow: 1, rowOffset: 0 }, width: 914400, height: 457200, rotation: 0, flipH: false, flipV: false, text, strokeWidth: 0, fillLineColor: 'UNOBSERVABLE:x', ...extra });
const obs = list => ({ format: 'xlsx', complete: true, sheets: [{ name: 'S', index: 0, used: null }], cells: {}, drawingReads: { S: list } });
const stepOf = (extra = {}) => ({ intent: 'set_shape_text', sheet: 'S', name: 'Box 1', text: ['Quarterly', '', 'notes'], ...extra });
const opOf = s => { const p = fam.plan(s, 0); assert.equal(p.ok, true, p.detail); return p.operation; };

test('G4 encode: the measured read shape -- every paragraph with \\r\\n, joined by |', () => {
  assert.equal(shapeText.encode(['one']), 'one\r\n');
  assert.equal(shapeText.encode(['a', 'b']), 'a\r\n|b\r\n');
  assert.equal(shapeText.encode(['x', '', 'z']), 'x\r\n|\r\n|z\r\n');
});

test('G4 schema + plan: a paragraph list; a line break, a control character, all-empty, too many paragraphs, an extra field refused', () => {
  assert.equal(W.validateWriteRequest(req([stepOf()])).ok, true);
  for (const [text, why] of [[['a\nb'], /line break/], [['bell\u0007'], /line break or control/], [['', ''], /non-empty/], [Array.from({ length: 11 }, () => 'p'), /1\.\.10/], ['plain string', /list/]]) {
    const p = fam.plan(stepOf({ text }), 0);
    assert.equal(p.ok, false, JSON.stringify(text));
    assert.match(p.detail, why);
  }
  assert.equal(W.validateWriteRequest(req([stepOf({ font: 'bold' })])).ok, false);
  assert.equal(W.validateWriteRequest(req([{ intent: 'insert_shape', sheet: 'S', name: 'B', geometry: 'rect', at: { col: 1, row: 1 }, size: { width: 100, height: 50 }, fill: '#AABBCC', text: ['hello', 'world'] }])).ok, true);
});

test('G4 set_shape_text classify: another text -> APPLY; this text already -> SATISFIED (0 write); no such shape / an image / an unobservable text -> named FAIL_CLOSED', () => {
  const op = opOf(stepOf());
  assert.equal(fam.classify(obs([shapeRaw('Box 1', 'old\r\n')]), op).state, 'APPLY');
  assert.equal(fam.classify(obs([shapeRaw('Box 1', shapeText.encode(['Quarterly', '', 'notes']))]), op).state, 'SATISFIED');
  assert.equal(fam.classify(obs([shapeRaw('Other', 'x\r\n')]), op).evidence.reason, 'xlsx-shape-text-no-such-shape');
  const image = { kind: 'image', name: 'Box 1', parent: 'S', pos: { fromCol: 1, colOffset: 0, fromRow: 1, rowOffset: 0 }, width: 914400, height: 457200, rotation: 0, flipH: false, flipV: false };
  assert.equal(fam.classify(obs([image]), op).evidence.reason, 'xlsx-shape-text-not-a-shape');
  assert.equal(fam.classify(obs([shapeRaw('Box 1', 'UNOBSERVABLE:CAPABILITY_MISSING')]), op).state, 'FAIL_CLOSED');
});

test('G4 set_shape_text lower: one shape of the name, cleared, the paragraphs written into element 0 + pushed', () => {
  const body = fam.lower(opOf(stepOf())).body;
  assert.match(body, /shape-identity/);
  assert.match(body, /c\.RemoveAllElements\(\);/);
  assert.match(body, /c\.GetElement\(0\)\.AddText\(__p\[0\]\)/);
  assert.match(body, /\["Quarterly","","notes"\]/);
});

test('G4 set_shape_text verify: L1 exact text; L3 keeps every other drawing and this one\'s other fields', () => {
  const op = opOf(stepOf());
  const want = shapeText.encode(['Quarterly', '', 'notes']);
  const other = shapeRaw('Other', 'keep\r\n', { pos: { fromCol: 5, colOffset: 0, fromRow: 5, rowOffset: 0 } });
  const before = { primary: obs([shapeRaw('Box 1', 'old\r\n'), other]) };
  assert.equal(fam.verifyPrimitive(op, before, { primary: obs([shapeRaw('Box 1', want), other]) }).ok, true);
  assert.equal(fam.verifyPrimitive(op, before, { primary: obs([shapeRaw('Box 1', 'Quarterly\r\n|notes\r\n'), other]) }).ok, false, 'the empty middle paragraph lost');
  assert.equal(fam.verifyPreservation(op, before, { primary: obs([shapeRaw('Box 1', want), other]) }).ok, true);
  assert.equal(fam.verifyPreservation(op, before, { primary: obs([shapeRaw('Box 1', want), shapeRaw('Other', 'changed\r\n', { pos: other.pos })]) }).ok, false, 'another shape\'s text changed');
  assert.equal(fam.verifyPreservation(op, before, { primary: obs([shapeRaw('Box 1', want, { rotation: 30 }), other]) }).ok, false, 'this shape changed more than its text');
});

test('G4 insert_shape text: planned, written after SetName, part of the identity', () => {
  const p = ins.plan({ intent: 'insert_shape', sheet: 'S', name: 'Box 2', geometry: 'rect', at: { col: 1, row: 1 }, size: { width: 72, height: 36 }, fill: '#DDEEFF', text: ['hello', 'world'] }, 0);
  assert.equal(p.ok, true, p.detail);
  const body = ins.lower(p.operation).body;
  assert.match(body, /s\.SetName\("Box 2"\);\s*const c = s\.GetContent\(\);/);
  assert.match(body, /\["hello","world"\]/);
  const bad = ins.plan({ intent: 'insert_shape', sheet: 'S', name: 'Box 2', geometry: 'rect', at: { col: 1, row: 1 }, size: { width: 72, height: 36 }, fill: '#DDEEFF', text: ['a\r\nb'] }, 0);
  assert.equal(bad.ok, false);
  const noText = ins.plan({ intent: 'insert_shape', sheet: 'S', name: 'Box 3', geometry: 'rect', at: { col: 1, row: 1 }, size: { width: 72, height: 36 }, fill: '#DDEEFF' }, 0);
  assert.doesNotMatch(ins.lower(noText.operation).body, /GetContent/, 'no text, no content write (v1 unchanged)');
});

test('G4 routing: set_shape_text plans through the shape-text family', () => {
  const p = planOperations({ id: 't', steps: [stepOf()] });
  assert.equal(p.ok, true, p.detail);
  assert.equal(p.operations[0].family, 'shape-text');
});
