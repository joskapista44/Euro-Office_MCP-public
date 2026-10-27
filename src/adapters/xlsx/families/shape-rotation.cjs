'use strict';

const cells = require('./cells.cjs');
const drawings = require('../drawings.cjs');

/**
 * XLSX `shape-rotation` family: set_shape_rotation (GH #302 G7) -- the rotation of one existing, named shape through the
 * public ApiDrawing.SetRotation(degrees). The READ mirror and the verifier is read_drawings' `rotation` (GetRotation).
 * The G7 LIVE design probe (docs/evidence/gh302-g7-shape-rotation-probe-20261026.json), nine shapes, session + reopen +
 * the saved <a:xfrm rot> + the installed product's read:
 *   - the file keeps 1/60000 degree units and the engine TRUNCATES: 30 -> rot 1799999 -> reads 29.999983333333333;
 *     45 / 90 / 180 / 12.5 exact; 359 -> 358.9999833333333. So the verifier allows exactly ONE stored unit
 *     (1/60000 degree) between the request and the read -- the measured truncation, nothing wider.
 *   - -30 is stored NEGATIVE (rot -1799999) and 400 as rot 24000000 (past a full turn): neither is the canonical
 *     0..360 form, so a request is 0 <= rotation < 360 (refused named otherwise).
 *   - 30 then 0 reads 0 (an existing rotation can be taken back).
 * Refused before the write (named): the drawing read incomplete, no shape of the name (or not a shape), its rotation
 * unreadable. SATISFIED: the shape already reads this rotation (within the one unit). L1: the read rotation. L3: every
 * cell; every drawing of every sheet unchanged -- this shape with only its rotation changed.
 */

const INTENTS = Object.freeze(['set_shape_rotation']);
const SHEET = /^[A-Za-z_][A-Za-z0-9_]*$/;
const NAME = /^[A-Za-z_][A-Za-z0-9_. -]{0,99}$/;
const FIELDS = Object.freeze(['intent', 'id', 'sheet', 'name', 'rotation']);
const UNITS = 60000; // the file's 1/60000 degree (ST_Angle)

/** the measured truncation: at most one stored unit between the requested and the read rotation */
const sameRotation = (read, want) => Number.isFinite(read) && Math.abs(Math.round(read * UNITS) - Math.round(want * UNITS)) <= 1;

function plan(step, index) {
  const at = `step ${index}: set_shape_rotation`;
  const unknown = Object.keys(step).filter(k => !FIELDS.includes(k));
  if (unknown.length) return { ok: false, detail: `${at} does not support field(s): ${unknown.join(', ')}` };
  if (typeof step.sheet !== 'string' || !SHEET.test(step.sheet)) return { ok: false, detail: `${at}: sheet must be a name that needs no quoting in a reference` };
  if (typeof step.name !== 'string' || !NAME.test(step.name)) return { ok: false, detail: `${at}: name must be a drawing name (letters, digits, _ . - space; starting with a letter or _)` };
  if (typeof step.rotation !== 'number' || !Number.isFinite(step.rotation) || step.rotation < 0 || step.rotation >= 360) return { ok: false, detail: `${at}: rotation must be degrees, 0 <= rotation < 360 (a negative or a full-turn value is stored outside the canonical range, measured)` };
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`,
      intent: 'set_shape_rotation',
      family: 'shape-rotation',
      selector: Object.freeze({ sheet: step.sheet, cellFree: true, claims: Object.freeze([`drawing:${step.sheet}:${step.name}`]) }),
      name: step.name,
      rotation: step.rotation,
      cells: Object.freeze([]),
      preservationMode: 'PRIMITIVE_INTEGRAL',
    },
  };
}

/** the one shape of the name, or the named reason */
function gate(p, op) {
  if (!p || p.format !== 'xlsx') return { ok: false, reason: 'xlsx-observation-missing' };
  if (!p.complete) return { ok: false, reason: p.truncated ? 'xlsx-observation-truncated' : 'xlsx-observation-unreadable-getters', unsupported: p.unsupported };
  if (p.sheets.filter(x => x.name === op.selector.sheet).length !== 1) return { ok: false, reason: 'xlsx-sheet-not-found', available: p.sheets.map(x => x.name) };
  if (!p.drawingReads) return { ok: false, reason: 'xlsx-shape-rotation-unobservable', why: 'not-read' };
  const o = drawings.observe(p.drawingReads, p.sheets.map(x => x.name));
  if (!o.complete) return { ok: false, reason: 'xlsx-shape-rotation-drawings-unobservable', unobservable: o.gaps.slice(0, 20), count: o.gaps.length };
  const hits = (o.bySheet[op.selector.sheet] || []).filter(r => r.name === op.name);
  if (hits.length !== 1) return { ok: false, reason: 'xlsx-shape-rotation-no-such-shape', sheet: op.selector.sheet, name: op.name, count: hits.length };
  if (hits[0].kind !== 'shape') return { ok: false, reason: 'xlsx-shape-rotation-not-a-shape', kind: hits[0].kind };
  return { ok: true, shape: hits[0] };
}

function classify(p, op) {
  const g = gate(p, op);
  if (!g.ok) return { state: 'FAIL_CLOSED', evidence: g };
  const where = { sheet: op.selector.sheet, name: op.name };
  return sameRotation(g.shape.rotation, op.rotation) ? { state: 'SATISFIED', evidence: { ...where, now: g.shape.rotation, proof: 'fresh-live-shape-reads-this-rotation' } } : { state: 'APPLY', evidence: { ...where, now: g.shape.rotation } };
}

function resolve(p, selector, ctx = {}) {
  const g = gate(p, ctx.operation);
  if (!g.ok) return { ok: false, target: null, selector, ambiguity: g.reason, detail: g };
  return { ok: true, target: Object.freeze({ sheet: selector.sheet, name: ctx.operation.name }) };
}

/** One callCommand: exact-name sheet, exactly one SHAPE of the name, SetRotation. */
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
  if (shapes[0].SetRotation(${J(op.rotation)}) === false) return { ok: false, stage: 'setter-rejected' };
  return { ok: true };
} catch (err) { return { ok: false, stage: 'threw', error: String(err) }; }
`,
    meta: Object.freeze({ family: 'shape-rotation', sheet: op.selector.sheet, name: op.name, rotation: op.rotation }),
  };
}

function verifyPrimitive(op, before, after) {
  const g = gate(after && after.primary, op);
  if (!g.ok) return { ok: false, evidence: { reason: g.reason } };
  return { ok: sameRotation(g.shape.rotation, op.rotation), evidence: { sheet: op.selector.sheet, name: op.name, now: g.shape.rotation, want: op.rotation, tolerance: '1/60000 degree (the measured truncation)', proof: 'fresh-live-drawing-GetRotation' } };
}

/** L3: every cell; every drawing of every sheet -- this one with ONLY its rotation changed (to the read value within the unit) */
function verifyPreservation(op, before, after) {
  const rest = cells.verifyPreservation(op, before, after);
  if (!rest.ok && rest.evidence && rest.evidence.reason) return rest;
  const sh = op.selector.sheet;
  const a = drawings.observe(after.primary.drawingReads, after.primary.sheets.map(x => x.name));
  const mineAfter = a.complete ? (a.bySheet[sh] || []).find(r => r.name === op.name) : null;
  const rotated = mineAfter && sameRotation(mineAfter.rotation, op.rotation) ? mineAfter.rotation : op.rotation;
  const draw = drawings.drawingDiffs(drawings.observe(before.primary.drawingReads, before.primary.sheets.map(x => x.name)), a,
    { expect: r => (r.sheet === sh && r.name === op.name ? Object.freeze({ ...r, rotation: rotated }) : r) });
  return { ok: rest.ok && draw.length === 0, evidence: { ...rest.evidence, drawingsKept: draw.length === 0, drawingDiffs: draw } };
}

function wholeTaskChecks(op, finalObservation) {
  const v = verifyPrimitive(op, null, finalObservation);
  return [{ name: `${op.selector.sheet}: shape ${op.name} rotation`, ok: v.ok, detail: v.ok ? null : v.evidence }];
}

const windowsOf = () => [];
const liveExtras = () => ({ drawingSheets: ['*'] });

module.exports = { INTENTS, UNITS, sameRotation, plan, classify, resolve, lower, verifyPrimitive, verifyPreservation, wholeTaskChecks, windowsOf, liveExtras };
