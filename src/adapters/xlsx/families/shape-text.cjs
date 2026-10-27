'use strict';

const cells = require('./cells.cjs');
const drawings = require('../drawings.cjs');
const shapeText = require('../shape-text.cjs');

/**
 * XLSX `shape-text` family: set_shape_text (GH #302 G4) -- replace the WHOLE text of one existing shape, named, with a
 * paragraph list: GetContent().RemoveAllElements() (which leaves one empty paragraph -- the G4 LIVE probe), then the same
 * write insert_shape's `text` uses (shape-text.cjs). The READ mirror and the verifier is read_drawings' `text`
 * (drawing.text.read/v1, the measured "p\r\n|q\r\n" shape).
 * Refused before the write (named): a drawing read that is incomplete, no shape of the name on the sheet (or not a shape),
 * a text the read cannot observe (the bridge capability missing on this runtime -- the write would be unverifiable).
 * SATISFIED: the shape already reads exactly this text (0 write). L1: it reads exactly this text. L3: every cell, and every
 * drawing of every sheet unchanged -- this shape with only its text changed.
 */

const INTENTS = Object.freeze(['set_shape_text']);
const SHEET = /^[A-Za-z_][A-Za-z0-9_]*$/;
const NAME = /^[A-Za-z_][A-Za-z0-9_. -]{0,99}$/;
const FIELDS = Object.freeze(['intent', 'id', 'sheet', 'name', 'text']);

function plan(step, index) {
  const at = `step ${index}: set_shape_text`;
  const unknown = Object.keys(step).filter(k => !FIELDS.includes(k));
  if (unknown.length) return { ok: false, detail: `${at} does not support field(s): ${unknown.join(', ')}` };
  if (typeof step.sheet !== 'string' || !SHEET.test(step.sheet)) return { ok: false, detail: `${at}: sheet must be a name that needs no quoting in a reference` };
  if (typeof step.name !== 'string' || !NAME.test(step.name)) return { ok: false, detail: `${at}: name must be a drawing name (letters, digits, _ . - space; starting with a letter or _)` };
  const why = shapeText.textProblem(step.text);
  if (why) return { ok: false, detail: `${at}: ${why}` };
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`,
      intent: 'set_shape_text',
      family: 'shape-text',
      selector: Object.freeze({ sheet: step.sheet, cellFree: true, claims: Object.freeze([`drawing:${step.sheet}:${step.name}`]) }),
      name: step.name,
      text: Object.freeze([...step.text]),
      cells: Object.freeze([]),
      preservationMode: 'PRIMITIVE_INTEGRAL',
    },
  };
}

/** the one shape of the name and its read text, or the named reason */
function gate(p, op) {
  if (!p || p.format !== 'xlsx') return { ok: false, reason: 'xlsx-observation-missing' };
  if (!p.complete) return { ok: false, reason: p.truncated ? 'xlsx-observation-truncated' : 'xlsx-observation-unreadable-getters', unsupported: p.unsupported };
  if (p.sheets.filter(x => x.name === op.selector.sheet).length !== 1) return { ok: false, reason: 'xlsx-sheet-not-found', available: p.sheets.map(x => x.name) };
  if (!p.drawingReads) return { ok: false, reason: 'xlsx-shape-text-unobservable', why: 'not-read' };
  const o = drawings.observe(p.drawingReads, p.sheets.map(x => x.name));
  if (!o.complete) return { ok: false, reason: 'xlsx-shape-text-drawings-unobservable', unobservable: o.gaps.slice(0, 20), count: o.gaps.length };
  const hits = (o.bySheet[op.selector.sheet] || []).filter(r => r.name === op.name);
  if (hits.length !== 1) return { ok: false, reason: 'xlsx-shape-text-no-such-shape', sheet: op.selector.sheet, name: op.name, count: hits.length };
  if (hits[0].kind !== 'shape') return { ok: false, reason: 'xlsx-shape-text-not-a-shape', kind: hits[0].kind };
  if (typeof hits[0].text !== 'string' || hits[0].text.startsWith('UNOBSERVABLE')) return { ok: false, reason: 'xlsx-shape-text-unobservable', why: hits[0].text };
  return { ok: true, shape: hits[0], observed: o };
}

function classify(p, op) {
  const g = gate(p, op);
  if (!g.ok) return { state: 'FAIL_CLOSED', evidence: g };
  const where = { sheet: op.selector.sheet, name: op.name };
  return g.shape.text === shapeText.encode(op.text) ? { state: 'SATISFIED', evidence: { ...where, proof: 'fresh-live-shape-reads-this-text' } } : { state: 'APPLY', evidence: where };
}

function resolve(p, selector, ctx = {}) {
  const g = gate(p, ctx.operation);
  if (!g.ok) return { ok: false, target: null, selector, ambiguity: g.reason, detail: g };
  return { ok: true, target: Object.freeze({ sheet: selector.sheet, name: ctx.operation.name }) };
}

/** One callCommand: exact-name sheet, exactly one SHAPE of the name, its content cleared, the paragraphs written. */
function lower(op) {
  const J = JSON.stringify;
  return {
    kind: 'callCommand',
    body: `
try {
  const ws = Api.GetSheet(${J(op.selector.sheet)});
  if (!ws || ws.GetName() !== ${J(op.selector.sheet)}) return { ok: false, stage: 'sheet-identity' };
  const nm = d => { try { return d.GetName(); } catch (e) { return null; } };
  const all = (ws.GetAllDrawings() || []).filter(d => nm(d) === ${J(op.name)});
  const shapes = (ws.GetAllShapes() || []).filter(d => nm(d) === ${J(op.name)});
  if (all.length !== 1 || shapes.length !== 1) return { ok: false, stage: 'shape-identity', drawings: all.length, shapes: shapes.length };
  const c = shapes[0].GetContent();
  c.RemoveAllElements();
  ${shapeText.writeJs(op.text)}
  return { ok: true };
} catch (err) { return { ok: false, stage: 'threw', error: String(err) }; }
`,
    meta: Object.freeze({ family: 'shape-text', sheet: op.selector.sheet, name: op.name, text: op.text }),
  };
}

function verifyPrimitive(op, before, after) {
  const g = gate(after && after.primary, op);
  if (!g.ok) return { ok: false, evidence: { reason: g.reason, why: g.why || null } };
  const want = shapeText.encode(op.text);
  return { ok: g.shape.text === want, evidence: { sheet: op.selector.sheet, name: op.name, now: g.shape.text, want, proof: 'fresh-live-drawing.text.read' } };
}

/** L3: every cell; every drawing of every sheet -- this one with ONLY its text changed */
function verifyPreservation(op, before, after) {
  const rest = cells.verifyPreservation(op, before, after);
  if (!rest.ok && rest.evidence && rest.evidence.reason) return rest;
  const sh = op.selector.sheet;
  const draw = drawings.drawingDiffs(
    drawings.observe(before.primary.drawingReads, before.primary.sheets.map(x => x.name)),
    drawings.observe(after.primary.drawingReads, after.primary.sheets.map(x => x.name)),
    { expect: r => (r.sheet === sh && r.name === op.name ? Object.freeze({ ...r, text: shapeText.encode(op.text) }) : r) },
  );
  return { ok: rest.ok && draw.length === 0, evidence: { ...rest.evidence, drawingsKept: draw.length === 0, drawingDiffs: draw } };
}

function wholeTaskChecks(op, finalObservation) {
  const v = verifyPrimitive(op, null, finalObservation);
  return [{ name: `${op.selector.sheet}: shape ${op.name} text`, ok: v.ok, detail: v.ok ? null : v.evidence }];
}

const windowsOf = () => [];
const liveExtras = () => ({ drawingSheets: ['*'] });

module.exports = { INTENTS, plan, classify, resolve, lower, verifyPrimitive, verifyPreservation, wholeTaskChecks, windowsOf, liveExtras };
