'use strict';

/**
 * M7 S2b (owner GO 2026-10-01, owner channel, on the fb7b1128 probe -- docs/evidence/M7-S2B-DRAWING-PROBE.md): the
 * drawings of a sheet (image / shape / chart) through the PUBLIC surface, normalized for a guard and a workbook-wide
 * L3. Pure.
 *
 * Reader (the driver's drawingReads[sheet], every sheet): a list of { kind, name, parent, pos: { fromCol, colOffset,
 * fromRow, rowOffset }, width, height, rotation, flipH, flipV, text } -- each field the getter's value or
 * 'UNOBSERVABLE:<why>'; the whole sheet 'UNOBSERVABLE:<why>' when the list itself cannot be read.
 *
 * Observability is an explicit, per-kind REQUIRED field list -- a probe-style "whatever getters exist" inventory is not
 * a guard. A gap is never an empty set or "unchanged":
 *   - a missing / throwing / partial read of a required field                         -> gap
 *   - a kind other than image / shape / chart (group, OLE, ...): no claim is made     -> gap
 *   - a name missing, empty, the string 'null', or shared by two drawings of a sheet  -> gap (identity ambiguous)
 *   - the parent sheet not the sheet that listed it                                   -> gap
 *   - a size that is missing / throwing / not finite / negative / zero on ONE side only    -> gap
 *   - an EXACT 0 x 0 size (any kind): only the SIZE is not observed (see ZERO_SIZE) -- cause NOT proven
 *   - flipH / flipV: a boolean OR null (null is a MEASURED valid value: image / chart) -- never mixed with a missing
 *     getter or with false
 * Sizes are normalized to an INTEGER EMU (the measured float noise 2160000.0000000005 -> 2160000; 1 EMU still differs).
 * No claim on image content, internal identity, groups / OLE or unmeasured anchor modes (D-B1 / D-B2: only the
 * build-gated anchor read of anchorExtent below, for the structure guard; D-B3 / D-B4 not approved).
 */

const KINDS = Object.freeze(['image', 'shape', 'chart']);
const COMMON = Object.freeze(['kind', 'name', 'parent', 'pos', 'width', 'height', 'rotation', 'flipH', 'flipV']);
const REQUIRED = Object.freeze({ image: COMMON, shape: Object.freeze([...COMMON, 'text']), chart: COMMON });
/**
 * An EXACT 0 x 0 size, any kind (owner decisions: A for the chart, owner channel answer to 30264; (2) for the image and
 * the shape, answer to 30280, after the LIVE run measured every image / shape 0 x 0 in a later session -- the cause is
 * NOT proven): only the SIZE axis is not observed; every other required field is still checked exactly and NO size
 * preservation is claimed. The L3 then: size positive before AND after -> integer EMU equality; positive before, 0 x 0
 * after -> a difference (never APPLIED); 0 x 0 before -> no size claim. Every other bad size stays a gap.
 * 'gap' = the earlier, stricter contract (kept for the record and the tests).
 */
const ZERO_SIZE = 'size-unobserved';
const SIZE_UNOBSERVED = 'SIZE_UNOBSERVED';

const unobs = x => typeof x === 'string' && x.startsWith('UNOBSERVABLE');
const nonNegInt = v => Number.isInteger(v) && v >= 0;

const POS_KEYS = Object.freeze(['fromCol', 'colOffset', 'fromRow', 'rowOffset']);
const flipOk = v => v === null || typeof v === 'boolean';

/** The first identity / presence problem of a raw drawing: [field, why] or null. */
function presenceProblem(raw, sheet) {
  if (!raw || typeof raw !== 'object') return ['drawing', 'not-an-object'];
  if (!KINDS.includes(raw.kind)) return ['kind', unobs(raw.kind) ? raw.kind : `kind-not-supported:${String(raw.kind)}`];
  const absent = REQUIRED[raw.kind].find(f => raw[f] === undefined || unobs(raw[f]));
  if (absent) return [absent, raw[absent] === undefined ? 'not-read' : raw[absent]];
  if (typeof raw.name !== 'string' || raw.name === '' || raw.name === 'null') return ['name', `name-unusable:${JSON.stringify(raw.name)}`];
  return raw.parent === sheet ? null : ['parent', `parent-mismatch:${JSON.stringify(raw.parent)}`];
}

/** The first shape problem (position, numbers, flips, shape text): [field, why] or null. */
function shapeProblem(raw) {
  const p = raw.pos;
  const checks = [
    ['pos', 'position-unexpected-shape', () => !!p && typeof p === 'object' && POS_KEYS.every(k => nonNegInt(p[k]))],
    ['size', 'size-or-rotation-not-a-number', () => [raw.width, raw.height, raw.rotation].every(Number.isFinite) && raw.width >= 0 && raw.height >= 0],
    ['flipH', 'flip-not-boolean-or-null', () => flipOk(raw.flipH)],
    ['flipV', 'flip-not-boolean-or-null', () => flipOk(raw.flipV)],
    ['text', 'text-not-a-string', () => raw.kind !== 'shape' || typeof raw.text === 'string'],
  ];
  const bad = checks.find(([, , ok]) => !ok());
  return bad ? [bad[0], bad[1]] : null;
}

/** Integer EMU; an EXACT 0 x 0 (decisions A / (2)) -> SIZE_UNOBSERVED; a zero on one side only -> null (a gap). */
function sizeOf(raw, zeroSize) {
  const width = Math.round(raw.width);
  const height = Math.round(raw.height);
  if (width !== 0 && height !== 0) return { width, height };
  // only the measured EXACT 0 x 0 is a size-observation limit; one zero side is a bad shape
  const exactZero = zeroSize === 'size-unobserved' && raw.width === 0 && raw.height === 0;
  return exactZero ? { width: SIZE_UNOBSERVED, height: SIZE_UNOBSERVED } : null;
}

/** One raw drawing of `sheet` -> { record } or { gap }. */
function normalize(raw, sheet, { zeroSize = ZERO_SIZE } = {}) {
  const gap = ([field, why]) => ({ gap: { sheet, field, why, name: raw && typeof raw.name === 'string' ? raw.name : null } });
  const problem = presenceProblem(raw, sheet) || shapeProblem(raw);
  if (problem) return gap(problem);
  const size = sizeOf(raw, zeroSize);
  if (!size) return gap(['size', 'size-zero-unexplained']);
  const p = raw.pos;
  return { record: Object.freeze({ sheet, kind: raw.kind, name: raw.name, fromCol: p.fromCol, colOffset: p.colOffset, fromRow: p.fromRow, rowOffset: p.rowOffset,
    ...size, rotation: raw.rotation, flipH: raw.flipH, flipV: raw.flipV, ...(raw.kind === 'shape' ? { text: raw.text } : {}) }) };
}

/** One sheet's list -> { recs, gaps }; two drawings sharing a name make the identity ambiguous (a gap). */
function observeSheet(list, sheet, opts) {
  if (!Array.isArray(list)) return { recs: [], gaps: [{ sheet, field: 'drawings', why: unobs(list) ? list : 'not-read' }] };
  const out = list.map(raw => normalize(raw, sheet, opts));
  const recs = out.filter(r => r.record).map(r => r.record);
  const gaps = out.filter(r => r.gap).map(r => r.gap);
  const seen = recs.reduce((m, r) => m.set(r.name, (m.get(r.name) || 0) + 1), new Map());
  for (const [name, k] of seen) if (k > 1) gaps.push({ sheet, field: 'name', why: `name-ambiguous:${k}`, name });
  return { recs, gaps };
}

/** Every sheet's drawings: { complete, gaps, bySheet: { sheet: [record] } }. `sheets` = the workbook's sheet names. */
function observe(drawingReads, sheets, opts = {}) {
  const gaps = [];
  const bySheet = {};
  for (const sheet of sheets) {
    const o = observeSheet(drawingReads ? drawingReads[sheet] : undefined, sheet, opts);
    gaps.push(...o.gaps);
    if (Array.isArray(drawingReads && drawingReads[sheet])) bySheet[sheet] = o.recs;
  }
  return { complete: gaps.length === 0, gaps, bySheet };
}

/**
 * Order-free, multiplicity-keeping comparison of two record lists: the records left unmatched on either side. Exact
 * matches first; then a BEFORE drawing whose size was not observed matches an after record equal on every OTHER field
 * (no size claim). A before size that was observed is never matched by an unobserved after size.
 */
function multisetDiff(before, after) {
  const k = r => JSON.stringify(r);
  const noSize = r => JSON.stringify({ ...r, width: null, height: null });
  const left = after.map(r => ({ r, used: false }));
  const unmatched = [];
  for (const r of before) { const m = left.find(x => !x.used && k(x.r) === k(r)); if (m) m.used = true; else unmatched.push(r); }
  const out = [];
  for (const r of unmatched) {
    const m = r.width === SIZE_UNOBSERVED ? left.find(x => !x.used && noSize(x.r) === noSize(r)) : null;
    if (m) m.used = true; else out.push({ change: 'missing-or-changed', drawing: r });
  }
  for (const x of left) if (!x.used) out.push({ change: 'appeared-or-changed', drawing: x.r });
  return out;
}

/**
 * L3: every sheet's drawings equal before and after, an observation gap on either side a difference. The requested
 * change is described, never inferred:
 *   mapSheet(name)     the sheet a before-drawing should be on after (rename_sheet: the parent changes on purpose)
 *   dropSheet(name)    a sheet whose drawings are gone on purpose (delete_sheet)
 *   skipAfter(name)    an after-sheet not compared here (copy_sheet's copy: L1 proves it)
 *   expect(records)    the before-records transformed by the requested chart change (only its requested fields)
 *   added(record)      an after-record that is the requested addition (set_chart's new chart)
 */
function drawingDiffs(b, a, { mapSheet = n => n, dropSheet = () => false, skipAfter = () => false, expect = r => r, added = () => false } = {}) {
  if (!b || !a) return [{ axis: 'drawings', change: 'observation-missing' }];
  if (!b.complete || !a.complete) return [{ axis: 'drawings', change: 'unobservable', before: b.gaps.slice(0, 10), after: a.gaps.slice(0, 10) }];
  const flat = (o, keep) => Object.entries(o.bySheet).filter(([sheet]) => keep(sheet)).flatMap(([sheet, recs]) => recs.map(r => ({ sheet, r })));
  const before = flat(b, n => !dropSheet(n)).map(({ sheet, r }) => expect(Object.freeze({ ...r, sheet: mapSheet(sheet) })));
  const after = flat(a, n => !skipAfter(n)).map(({ r }) => r).filter(r => !added(r));
  return multisetDiff(before, after).map(d => ({ axis: 'drawings', ...d })).slice(0, 20);
}

/**
 * M7 D-B1 / D-B2 (owner GO 2026-10-03, owner channel point 2; M7-DB-PROBE.md): the LAST cell a drawing occupies,
 * 0-based, from the build-gated internal anchor read (raw.anchor = { type, to: { col, colOff, row, rowOff } }). Only the
 * MEASURED mode: Type 1 (oneCell -- every public creation). `to` is taken as occupied whatever its offset (an offset 0
 * is not treated as "ends before it": conservative). Also required, else the geometry is incomplete: a usable record
 * (normalize: no gap), an OBSERVED size (not the 0 x 0 SIZE_UNOBSERVED), rotation 0 (a rotated extent is not measured),
 * and `to` not before the drawing's own from-cell. Returns { lastCol, lastRow } or { why }.
 */
const MEASURED_ANCHOR_TYPES = Object.freeze([1]);

const anchorTo = a => (a && typeof a === 'object' ? a.to : null);
const toComplete = to => !!to && typeof to === 'object' && ['col', 'row'].every(k => nonNegInt(to[k])) && ['colOff', 'rowOff'].every(k => Number.isFinite(to[k]) && to[k] >= 0);
/** The first reason a normalized record + its raw anchor is not a measured, complete geometry, or null. */
function anchorProblem(r, a) {
  const checks = [
    [() => r.width !== SIZE_UNOBSERVED, () => 'size-not-observed'],
    [() => r.rotation === 0, () => 'rotation-not-measured'],
    [() => !!a && typeof a === 'object', () => (unobs(a) ? a : 'anchor-not-read')],
    [() => MEASURED_ANCHOR_TYPES.includes(a.type), () => `anchor-mode-not-measured:${JSON.stringify(a.type)}`],
    [() => toComplete(a.to), () => 'anchor-to-incomplete'],
    [() => a.to.col >= r.fromCol && a.to.row >= r.fromRow, () => 'anchor-to-before-from'],
  ];
  const bad = checks.find(([ok]) => !ok());
  return bad ? bad[1]() : null;
}

function anchorExtent(raw, sheet, opts) {
  const n = normalize(raw, sheet, opts);
  const why = n.gap ? `drawing-unobservable:${n.gap.field}` : anchorProblem(n.record, raw.anchor);
  if (why) return { why };
  const to = anchorTo(raw.anchor);
  return { lastCol: to.col, lastRow: to.row };
}

module.exports = { KINDS, REQUIRED, ZERO_SIZE, SIZE_UNOBSERVED, MEASURED_ANCHOR_TYPES, normalize, observe, multisetDiff, drawingDiffs, anchorExtent };
