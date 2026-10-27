'use strict';

const cells = require('./cells.cjs');
const drawings = require('../drawings.cjs');

/**
 * XLSX `drawings-insert` family: insert_shape / insert_image (X-O08, product-slice dispatch,
 * docs/X-O08-DRAWING-INSERT-DESIGN-20261007.md v5). Owner mandate for this new intent family: the
 * FORMATTING COMPLETION backlog's own standing authorization (reaffirmed 2026-10-06 -- "the
 * formatting-completion workstream continues"), NOT the general fejlesztesi-stop measurement-only
 * exemption (which this family would otherwise fall outside of, being termekfunkcio).
 *
 * v1 scope, per the design doc's FINAL decisions (X-O08-1a through X-O08-6):
 *   insert_shape { sheet, name, geometry: 'rect', at: { col, row }, size: { width, height } (whole
 *     points, * 12700 to EMU at the Builder boundary), fill: '#RRGGBB' }
 *   insert_image { sheet, name, at: { col, row }, size: { width, height } (points), source:
 *     'data:image/png;base64,...' } -- PNG ONLY (the only mime type LIVE-measured, X-O08-1a/1b);
 *     an external http(s):// URL is REFUSED AT PLAN TIME, never fetched (outside the measured/
 *     approved egress path, X-O08's own design doc 2b).
 *   POSITION is cell-aligned ONLY: `colOffset`/`rowOffset` are not exposed, always 0 at the
 *     Builder boundary (X-O08-3 measured this EXACT, on two different cells -- the one value that
 *     never needed the ±1 EMU tolerance question).
 *   SIZE is NEVER an exact-match condition, in EITHER unit (X-O08-3 measured non-round EMU drifts
 *     by -1; X-O08-4 measured ALL SIXTEEN non-round POINT-derived sizes drift the same way,
 *     systematically, on both axes) -- `size` is accepted and WRITTEN, but never compared for
 *     exact equality anywhere in this family (not in identity, not in verifyPrimitive). The ±1 EMU
 *     drift is the SAME open owner-tolerance item as the chart-size question (X-O06) -- not
 *     resolved here.
 *   NAME is the retry-identity key, FAIL_CLOSED on a collision with another identity (no silent
 *     modify) -- the SAME rule `charts.cjs`'s `set_chart` already uses. Identity here is LIMITED
 *     to what the driver's `drawingReads` record actually carries: kind (shape/image) + position
 *     (fromCol/fromRow, offset always 0). SOURCE (the image content) is NOT part of identity --
 *     the driver's drawing reader has no image-content field at all (confirmed by reading that
 *     fragment directly, not guessed); a retry therefore cannot tell two different-image
 *     same-position-same-name images apart, and does not claim to.
 *     FILL/STROKE colour (GH #139 D4/D6, X-O08 drawing.fillLineColor.read/v1, DS /14+) is part of
 *     identity ONLY when the read actually carries it -- on an unaccepted (pre-/14) runtime it
 *     degrades to its own named gap and colour stays unprovable, same as before this capability
 *     existed (see colourCheck() below and colourCheck's updated doc).
 *
 * CORRECTION to the design doc's decision 4: position was declared session-
 * unreadable because a LIVE probe (X-O08-2) reflected `shape.GetPosition()`'s PROTOTYPE for
 * Get-prefixed METHODS and found zero. Reading `xlsx-driver.cjs`'s own, already-shipped drawing
 * reader fragment directly (not probed, READ) shows `GetPosition()` returns a plain DATA object
 * (`pos.fromCol` / `pos.colOffset` / `pos.fromRow` / `pos.rowOffset`, own enumerable properties,
 * not methods) -- exactly the shape `drawings.cjs`'s `normalize()` already consumes, and exactly
 * what the ALREADY-SHIPPED `charts` family already relies on (same driver code, same `drawings.
 * observe()` call, proven for the 'chart' kind). The probe's reflection technique (prototype-
 * method-only) was the wrong tool to inspect a plain-property object -- not a genuine absence.
 * Position is therefore included in this family's session-side classify/verifyPrimitive via
 * `drawings.observe()`, the SAME proven mechanism `charts.cjs` already uses, not FILE-only.
 *
 * NOT-MEASURED, stated rather than assumed: whether the driver's OWN `drawingReads` fragment,
 * read in the SAME session immediately after an `AddShape`/`AddImage` dispatch (no reopen), shows
 * the same ±1 EMU size drift the FILE bytes showed in X-O08-3/4/5 (those probes measured the SAVED
 * FILE after a `settle()`, never the same-session read before any save) -- irrelevant to this
 * family's own correctness since size is never exact-matched either way, but the gap is named,
 * not glossed over. Flagged as a LIVE acceptance question, not resolved by this patch.
 *
 * v2 (16fd90b9, a LIVE measurement round): `insert_shape` gains an OPTIONAL `stroke: { width,
 * color }`. WIDTH is the measured READABLE axis -- `shape.SetLine(Api.CreateStroke(...))`
 * written AFTER creation (the exact mechanism LIVE-measured, not AddShape's own stroke argument,
 * which was never independently re-measured for readback), then `shape.GetLine().GetWidth()`
 * PUBLIC-readable same-session AND after reopen -- so it is part of this
 * family's L1 identity/verifyPrimitive, the SAME way position already is. COLOUR (both the
 * existing `fill` and this new `stroke.color`) was WRITE_ONLY through the PUBLIC getter:
 * `GetFill()` returns only an opaque internal object (`GetType()` says 'solid', the RGB itself is
 * in `UniFill`, direct-internal, not a public accessor) -- written, confirmed landed in the FILE
 * bytes, but not independently re-readable that way.
 *
 * GH #139 D4/D6 (2026-10-17, this capability closes that specific PUBLIC-getter limit, not a new
 * one): the driver's `fillLineColor` read (X-O08, DS Bridge drawing.fillLineColor.read/v1, the
 * internal model -- NOT GetFill()/GetLine(), which stay opaque, unchanged) reads the fill/line RGB
 * back on an accepted /14+ runtime, and LIVE-measured (three independent prod samples, including a
 * same-session-after-write one) matches this family's own write-side hex exactly: uppercase 6-hex,
 * no '#', the SAME normalization `plan()` already applies to `op.fill`/`op.stroke.color`.
 * `colourCheck()` below compares it and, WHEN the read actually carries an object (not
 * 'UNOBSERVABLE:<reason>' -- a pre-/14 runtime or a throw), folds the match into `identityOf()`'s
 * L1 proof. `FILL_COLOUR_AXIS`/`STROKE_COLOUR_AXIS` below are therefore no longer
 * ALWAYS named in `wholeTaskChecks`' `notProven` -- only when the read did not carry the
 * capability on THIS call (the SAME `notProven` vocabulary format.cjs's X-N01b scope already uses,
 * applied here to an L1/write axis instead of L3/preservation, now conditional on observed
 * capability rather than a permanent, unconditional limit).
 */
const FILL_COLOUR_AXIS = 'shape-fill-colour';
const STROKE_COLOUR_AXIS = 'shape-stroke-colour';

/**
 * GH #139 D4/D6: the fillLineColor read against the op's OWN written values (not identity-wise
 * comparable at all on a 'UNOBSERVABLE:<reason>' read -- a gap on this ONE field, never treated as
 * a mismatch). `fillProven`/`strokeProven` say whether THIS call's read actually carried the
 * capability (-> drops the matching axis from `notProven`); `fillOk`/`strokeOk` default to `true`
 * when unprovable, so an unprovable axis never fails `identityOf()` -- only a PROVEN mismatch does.
 * Stroke is checked ONLY when `op.stroke` was written (the SAME scope STROKE_COLOUR_AXIS already
 * names) -- a shape with no `stroke` writes the product's own hardcoded default
 * (`Api.CreateStroke(0, Api.CreateNoFill())`, `line: 'noLine'`), but that default is not this
 * family's OWN claim and is not asserted here, same as before this capability existed.
 */
function colourCheck(op, hit) {
  const flc = hit && hit.fillLineColor;
  if (!flc || typeof flc !== 'object') return { fillProven: false, fillOk: true, strokeProven: false, strokeOk: true };
  const fillOk = flc.fill === op.fill;
  if (!op.stroke) return { fillProven: true, fillOk, strokeProven: false, strokeOk: true };
  const strokeOk = flc.line === op.stroke.color;
  return { fillProven: true, fillOk, strokeProven: true, strokeOk };
}

const shapeText = require('../shape-text.cjs');

const INTENTS = Object.freeze(['insert_shape', 'insert_image']);
const SHEET = /^[A-Za-z_][A-Za-z0-9_]*$/;
const NAME = /^[A-Za-z_][A-Za-z0-9_. -]{0,99}$/;
const GEOMETRIES = Object.freeze(['rect']);
const FILL_RE = /^#?[0-9A-Fa-f]{6}$/;
const PNG_DATA_URI = /^data:image\/png;base64,[A-Za-z0-9+/]+=*$/;
const EMU_PER_PT = 12700;

const FIELDS = Object.freeze({
  insert_shape: ['intent', 'id', 'sheet', 'name', 'geometry', 'at', 'size', 'fill', 'stroke', 'text'],
  insert_image: ['intent', 'id', 'sheet', 'name', 'at', 'size', 'source'],
});

const ptToEmu = pt => Math.round(pt * EMU_PER_PT);

function posProblem(at) {
  if (!at || typeof at !== 'object') return 'at must be an object { col, row }';
  if (!Number.isInteger(at.col) || at.col < 0) return 'at.col must be a non-negative integer';
  if (!Number.isInteger(at.row) || at.row < 0) return 'at.row must be a non-negative integer';
  return null;
}

function sizeProblem(size) {
  if (!size || typeof size !== 'object') return 'size must be an object { width, height }';
  if (!Number.isFinite(size.width) || size.width <= 0) return 'size.width must be a positive number of points';
  if (!Number.isFinite(size.height) || size.height <= 0) return 'size.height must be a positive number of points';
  return null;
}

/** stroke.width is the measured READABLE axis (GetLine().GetWidth(), X-O08-7 LIVE r3) -- held to
 * the SAME positive-points shape `size` already uses. stroke.color is WRITE_ONLY (never verified,
 * see STROKE_COLOUR_AXIS), still validated so a malformed value fails at plan time, not
 * silently written as whatever CreateRGBColor does with a bad string. */
function strokeProblem(stroke) {
  if (!stroke || typeof stroke !== 'object') return 'stroke must be an object { width, color }';
  if (!Number.isFinite(stroke.width) || stroke.width <= 0) return 'stroke.width must be a positive number of points';
  if (typeof stroke.color !== 'string' || !FILL_RE.test(stroke.color)) return 'stroke.color must be a 6-hex colour, e.g. \'#RRGGBB\'';
  return null;
}

function plan(step, index) {
  const at = `step ${index}: ${step.intent}`;
  const isShape = step.intent === 'insert_shape';
  const fields = FIELDS[step.intent] || [];
  const unknown = Object.keys(step).filter(k => !fields.includes(k));
  if (unknown.length) return { ok: false, detail: `${at} does not support field(s): ${unknown.join(', ')}` };
  if (typeof step.sheet !== 'string' || !SHEET.test(step.sheet)) return { ok: false, detail: `${at}: sheet must be a name that needs no quoting in a reference` };
  if (typeof step.name !== 'string' || !NAME.test(step.name)) return { ok: false, detail: `${at}: name must be a drawing name (letters, digits, _ . - space; starting with a letter or _)` };
  const posErr = posProblem(step.at);
  if (posErr) return { ok: false, detail: `${at}: ${posErr}` };
  const sizeErr = sizeProblem(step.size);
  if (sizeErr) return { ok: false, detail: `${at}: ${sizeErr}` };
  let fill = null;
  let source = null;
  let stroke = null;
  if (isShape) {
    if (!GEOMETRIES.includes(step.geometry)) return { ok: false, detail: `${at}: geometry must be one of ${GEOMETRIES.join(', ')}` };
    if (typeof step.fill !== 'string' || !FILL_RE.test(step.fill)) return { ok: false, detail: `${at}: fill must be a 6-hex colour, e.g. '#RRGGBB'` };
    fill = step.fill.replace(/^#/, '').toUpperCase();
    if (step.stroke !== undefined) {
      const strokeErr = strokeProblem(step.stroke);
      if (strokeErr) return { ok: false, detail: `${at}: ${strokeErr}` };
      stroke = Object.freeze({ width: step.stroke.width, color: step.stroke.color.replace(/^#/, '').toUpperCase() });
    }
    // GH #302 G4: an optional paragraph list written into the new shape (shape-text.cjs, LIVE design probe)
    if (step.text !== undefined) {
      const textErr = shapeText.textProblem(step.text);
      if (textErr) return { ok: false, detail: `${at}: ${textErr}` };
    }
  } else {
    if (typeof step.source !== 'string' || !PNG_DATA_URI.test(step.source)) return { ok: false, detail: `${at}: source must be a data:image/png;base64,... URI -- only PNG is LIVE-measured, and an external URL is never accepted` };
    source = step.source;
  }
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`,
      intent: step.intent,
      family: 'drawings-insert',
      // a drawing claims its NAME on the sheet (two steps on one name clash, same as charts.cjs); no cell footprint
      selector: Object.freeze({ sheet: step.sheet, cellFree: true, claims: Object.freeze([`drawing:${step.sheet}:${step.name}`]) }),
      name: step.name,
      geometry: isShape ? step.geometry : null,
      fill,
      stroke,
      text: isShape && step.text !== undefined ? Object.freeze([...step.text]) : null,
      source,
      at: Object.freeze({ col: step.at.col, row: step.at.row }),
      size: Object.freeze({ width: step.size.width, height: step.size.height }),
      cells: Object.freeze([]),
      preservationMode: 'PRIMITIVE_INTEGRAL',
    },
  };
}

function gate(p, op) {
  if (!p || p.format !== 'xlsx') return { ok: false, reason: 'xlsx-observation-missing' };
  if (!p.complete) return { ok: false, reason: p.truncated ? 'xlsx-observation-truncated' : 'xlsx-observation-unreadable-getters', unsupported: p.unsupported };
  const s = p.sheets.filter(x => x.name === op.selector.sheet);
  if (s.length !== 1) return { ok: false, reason: s.length > 1 ? 'xlsx-sheet-ambiguous' : 'xlsx-sheet-not-found', available: p.sheets.map(x => x.name) };
  if (!p.drawingReads) return { ok: false, reason: 'xlsx-drawing-insert-unobservable', why: 'not-read' };
  return { ok: true };
}

const kindOf = op => (op.intent === 'insert_shape' ? 'shape' : 'image');

/** Identity, LIMITED to what the driver's drawingReads record carries: kind + position (+ stroke
 * width, v2, when the operation itself wrote one -- X-O08-7 LIVE r3 confirmed it PUBLIC-readable).
 * SOURCE is NOT comparable -- the driver's reader has no image-content field. FILL/STROKE colour
 * (GH #139 D4/D6) IS now part of identity, but ONLY when the read actually proves it -- colourCheck()
 * returns fillOk/strokeOk=true on an unprovable (pre-/14) read, so this never fails on that axis
 * alone (see colourCheck()'s updated doc). */
function identityOf(op, hit) {
  const c = colourCheck(op, hit);
  return hit.kind === kindOf(op) && hit.fromCol === op.at.col && hit.fromRow === op.at.row && hit.colOffset === 0 && hit.rowOffset === 0
    && (!op.stroke || hit.strokeWidth === ptToEmu(op.stroke.width)) && c.fillOk && c.strokeOk
    // GH #302 G4: the written text, as the read gives it back (shape-text.cjs encode)
    && (!op.text || hit.text === shapeText.encode(op.text));
}

function classify(p, op) {
  const g = gate(p, op);
  if (!g.ok) return { state: 'FAIL_CLOSED', evidence: g };
  const o = drawings.observe(p.drawingReads, p.sheets.map(x => x.name));
  if (!o.complete) return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-drawing-insert-drawings-unobservable', unobservable: o.gaps.slice(0, 20), count: o.gaps.length } };
  const where = { sheet: op.selector.sheet, name: op.name };
  // no separate "two drawings of this name" branch: drawings.cjs's observeSheet() already turns a
  // duplicate name into an incompleteness GAP (o.complete === false, caught above) -- a second
  // check here could never fire against the current drawings.cjs and would be untested dead code.
  const hits = (o.bySheet[op.selector.sheet] || []).filter(r => r.name === op.name);
  if (hits.length === 1) {
    return identityOf(op, hits[0])
      ? { state: 'SATISFIED', evidence: { ...where, proof: 'fresh-live-drawing-with-the-identity' } }
      : { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-drawing-insert-exists-with-another-identity', ...where, now: { kind: hits[0].kind, fromCol: hits[0].fromCol, fromRow: hits[0].fromRow } } };
  }
  return { state: 'APPLY', evidence: { ...where, at: op.at } };
}

function resolve(p, selector, ctx = {}) {
  const g = gate(p, ctx.operation);
  if (!g.ok) return { ok: false, target: null, selector, ambiguity: g.reason, detail: g };
  return { ok: true, target: Object.freeze({ sheet: selector.sheet, name: selector.name }) };
}

function lower(op) {
  const { sheet } = op.selector;
  const isShape = op.intent === 'insert_shape';
  const widthEmu = ptToEmu(op.size.width);
  const heightEmu = ptToEmu(op.size.height);
  const nameOf = 'd => { try { return d.GetName(); } catch (e) { return null; } }';
  let call;
  if (isShape) {
    const hex = op.fill;
    const r = parseInt(hex.slice(0, 2), 16);
    const gC = parseInt(hex.slice(2, 4), 16);
    const b = parseInt(hex.slice(4, 6), 16);
    // v2 stroke (X-O08-7 LIVE r3): written AFTER creation via SetLine -- the EXACT mechanism
    // measured READABLE for width (AddShape's own stroke constructor argument was never
    // independently re-measured for readback, so it is not reused here for the non-default case).
    let setLine = '';
    if (op.stroke) {
      const strokeWidthEmu = ptToEmu(op.stroke.width);
      const sr = parseInt(op.stroke.color.slice(0, 2), 16);
      const sg = parseInt(op.stroke.color.slice(2, 4), 16);
      const sb = parseInt(op.stroke.color.slice(4, 6), 16);
      setLine = `
  const strokeFill = Api.CreateSolidFill(Api.CreateRGBColor(${sr}, ${sg}, ${sb}));
  s.SetLine(Api.CreateStroke(${strokeWidthEmu}, strokeFill));`;
    }
    call = `const fill = Api.CreateSolidFill(Api.CreateRGBColor(${r}, ${gC}, ${b}));
  const stroke = Api.CreateStroke(0, Api.CreateNoFill());
  const s = ws.AddShape(${JSON.stringify(op.geometry)}, ${widthEmu}, ${heightEmu}, fill, stroke, ${op.at.col}, 0, ${op.at.row}, 0);
  if (!s) return { ok: false, stage: 'add-returned-nothing' };
  s.SetName(${JSON.stringify(op.name)});${setLine}${op.text ? `
  const c = s.GetContent();
  ${shapeText.writeJs(op.text)}` : ''}`;
  } else {
    call = `const img = ws.AddImage(${JSON.stringify(op.source)}, ${widthEmu}, ${heightEmu}, ${op.at.col}, 0, ${op.at.row}, 0);
  if (!img) return { ok: false, stage: 'add-returned-nothing' };
  img.SetName(${JSON.stringify(op.name)});`;
  }
  return {
    kind: 'callCommand',
    body: `
try {
  const name = ${JSON.stringify(sheet)};
  const ws = Api.GetSheet(name);
  if (!ws || ws.GetName() !== name) return { ok: false, stage: 'sheet-identity' };
  if ((ws.GetAllDrawings() || []).some(d => (${nameOf})(d) === ${JSON.stringify(op.name)})) return { ok: false, stage: 'name-taken' };
  ${call}
  return { ok: true };
} catch (err) { return { ok: false, stage: 'threw', error: String(err) }; }
`,
    meta: Object.freeze({ family: 'drawings-insert', action: op.intent, sheet, name: op.name, kind: kindOf(op), at: op.at, sizeEmu: Object.freeze({ width: widthEmu, height: heightEmu }), fill: op.fill, stroke: op.stroke, text: op.text, source: op.source }),
  };
}

const NO_COLOUR_PROOF = Object.freeze({ fillProven: false, fillOk: true, strokeProven: false, strokeOk: true });

function verifyPrimitive(op, before, after) {
  const p = after && after.primary;
  const g = gate(p, op);
  if (!g.ok) return { ok: false, colour: NO_COLOUR_PROOF, evidence: g };
  const o = drawings.observe(p.drawingReads, p.sheets.map(x => x.name));
  if (!o.complete) return { ok: false, colour: NO_COLOUR_PROOF, evidence: { reason: 'xlsx-drawing-insert-drawings-unobservable', unobservable: o.gaps.slice(0, 20), count: o.gaps.length } };
  const hits = (o.bySheet[op.selector.sheet] || []).filter(r => r.name === op.name);
  const hit = hits[0] || null;
  const ok = hits.length === 1 && identityOf(op, hit);
  return { ok, colour: hit ? colourCheck(op, hit) : NO_COLOUR_PROOF, evidence: { sheet: op.selector.sheet, name: op.name, now: hit, proof: 'fresh-live-drawing-getters' } };
}

/** L3: every cell, every OTHER drawing of every sheet (the SAME `drawings.drawingDiffs` pattern
 * `charts.cjs` already uses) -- the requested addition is the ONLY exemption. */
function verifyPreservation(op, before, after) {
  const rest = cells.verifyPreservation(op, before, after);
  if (!rest.ok && rest.evidence && rest.evidence.reason) return rest;
  const sh = op.selector.sheet;
  const draw = drawings.drawingDiffs(
    drawings.observe(before.primary.drawingReads, before.primary.sheets.map(x => x.name)),
    drawings.observe(after.primary.drawingReads, after.primary.sheets.map(x => x.name)),
    { added: r => r.sheet === sh && r.name === op.name && r.kind === kindOf(op) },
  );
  return { ok: rest.ok && draw.length === 0, evidence: { ...rest.evidence, drawingsKept: draw.length === 0, drawingDiffs: draw } };
}

function wholeTaskChecks(op, finalObservation) {
  const v = verifyPrimitive(op, null, finalObservation);
  const what = `${kindOf(op)} ${op.name} at col ${op.at.col}/row ${op.at.row}`;
  // GH #139 D4/D6: fill-colour / stroke-colour are named in `notProven` ONLY when THIS call's
  // fillLineColor read did not carry the capability (a pre-/14 runtime or a throw) -- a PROVEN,
  // matching read drops the axis (it is now folded into `ok` via identityOf()/colourCheck()
  // instead); a PROVEN mismatch already made `ok` false above, and is NOT additionally named here
  // (the SAME `notProven` vocabulary format.cjs's X-N01b scope uses for an axis that is not
  // independently checked -- a proven, failing check is not that; it is a genuine L1 FAIL).
  // `operationId: op.id` lets write-tools.cjs's mapExecutionResult thread this into the per-
  // operation RECEIPT (operations[i].semanticsNotProven) -- without it, the limit would only ever
  // reach the executor's own internal result, never the MCP caller.
  const notProven = op.intent !== 'insert_shape' ? [] : [
    ...(!v.colour.fillProven ? [FILL_COLOUR_AXIS] : []),
    ...(op.stroke && !v.colour.strokeProven ? [STROKE_COLOUR_AXIS] : []),
  ];
  return [{ name: `${op.selector.sheet}: ${what}`, ok: v.ok, detail: v.ok ? null : v.evidence, operationId: op.id, ...(notProven.length ? { notProven } : {}) }];
}

const windowsOf = () => [];
const liveExtras = () => ({ drawingSheets: ['*'] });

module.exports = { INTENTS, plan, classify, resolve, lower, verifyPrimitive, verifyPreservation, wholeTaskChecks, windowsOf, liveExtras };
