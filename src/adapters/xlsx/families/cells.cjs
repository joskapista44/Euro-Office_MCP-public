'use strict';

const { parseRange, cellsOf, rangeText, overlaps } = require('../a1.cjs');
const { cellAt, key, isBlank } = require('../inventory.cjs');
const formula = require('../formula.cjs');
const { nameUses } = require('../dependencies.cjs');
const { observedStyleChanges } = require('../style-axis.cjs');

/**
 * XLSX `cells` family: write_range with VALUES (M6 wave 1) and FORMULAS (M6 wave 2).
 *
 * Formulas (wave-2 LIVE probe, docs/evidence/M6-W2-FORMULA-PROBE.md): written with SetValue('=...');
 * identity is canonical (formula.cjs -- the editor re-renders the text, literals kept verbatim); a sheet
 * reference must match a live sheet name EXACTLY before dispatch (a case-only difference is silently
 * dropped by the editor); an error value is a separate whole-task failure, named per cell.
 *
 * Source oracle (joskapista44/euro-mcp-public fedaf4a, read-only; audit rows B-C01, B-C06, B-C10,
 * B-C16): bulk-writer.cjs (SetValue per cell), verification-contract.cjs (cell equality, numeric proof
 * through COUNT of the single-cell ApiRange), range-reader.cjs (public getters only).
 *
 * The contract differs from the source where the source audit found a defect, stated:
 *   - a value that the editor would COERCE is refused BEFORE dispatch: a string starting with
 *     = + - @ ' (formula / injection / apostrophe prefix), a numeric-, boolean- or date-looking string,
 *     an empty string and surrounding whitespace. The source let `=...` strings through and failed only
 *     after the write (B-C01);
 *   - writing over a cell that holds a formula, and writing into a merged area, are refused (the source
 *     had no such guard);
 *   - L3 is real: every OTHER read cell of the workbook keeps its formula, and a non-formula cell keeps
 *     its value (a formula cell's value may follow its inputs); the source read nothing outside the range;
 *   - the retry is the terminal state read fresh (classifyOperation), never a retry token.
 *
 * M7 S26 / S28 (owner M7 GO 2026-10-03; tier D probe docs/evidence/M7-D-PROBE.md, session = reopen): a TYPED value in
 * the values matrix --
 *   { text: '<string>' }                  -> SetNumberFormat('@') THEN SetValue(text): the exact string is kept
 *                                           (001, 1/2, TRUE, ' a ', 2026-10-03 measured; COUNT 0, nf '@');
 *                                           a text starting with = + - @ stays refused (=x reads back as a formula);
 *                                           the reverse order coerces (measured), so the order is the contract
 *   { date: 'YYYY-MM-DD', format }         -> SetValue(serial) THEN SetNumberFormat(format), format one of the measured
 *   { date: 'YYYY-MM-DDTHH:MM', format }      masks (DATE_MASKS); the serial is computed here (1900 system,
 *   { time: 'HH:MM', format: 'hh:mm' }        2026-10-03 = 46298 measured)
 * The number format is a DECLARED change on exactly those cells: L1 proves value + COUNT + nf; L3 (writePreservation)
 * proves every other number format read on the sheet unchanged.
 */

const MAX_CELLS = 1000;
const MAX_TEXT = 32767;
const FIELDS = Object.freeze(['intent', 'id', 'sheet', 'range', 'values', 'formulas']);
const MAX_FORMULA = 8192;
const ERROR_VALUE = /^#(DIV\/0!|NAME\?|REF!|VALUE!|N\/A|NUM!|NULL!|GETTING_DATA)$/;

/** A formula the planner accepts: '=' first, tokenizable, bounded, no external [book] reference. */
function formulaProblem(f) {
  if (typeof f !== 'string' || !f.startsWith('=')) return 'a formula must be a string starting with =';
  if (f.length > MAX_FORMULA) return `formula longer than ${MAX_FORMULA} characters`;
  if (formula.canonical(f) === null) return 'formula has an unterminated string literal or quoted name';
  if (formula.sheetRefs(f).external) return 'external workbook references are refused';
  return null;
}

// A string the editor would turn into something else (measured source behaviour + conservative).
const COERCIBLE = [
  [/^[=+\-@']/, 'starts with = + - @ or an apostrophe (formula, injection or prefix)'],
  [/^[-+]?(\d[\d,]*(\.\d*)?|\.\d+)([eE][-+]?\d+)?%?$/, 'looks like a number'],
  [/^(true|false)$/i, 'looks like a boolean'],
  [/^\d{1,4}[-./]\d{1,2}([-./]\d{1,4})?$/, 'looks like a date'],
  [/^\d{1,2}:\d{2}(:\d{2})?(\s?[ap]m)?$/i, 'looks like a time'],
];

// X-F11: three masks ADDED on the strength of the X-F09 S23 probe measurement
// (m7-s23-probe1, LIVE, byte-exact through two reopens) -- 'dd/mm/yyyy hh:mm' into datetime,
// '[$-409]mmmm d, yyyy' and '[$-40E]yyyy. mmmm d.' into date. That probe wrote the mask onto an
// ALREADY-NUMERIC cell (SetNumberFormat only); it did not measure the typed write_range order
// (SetValue(serial) THEN SetNumberFormat) with these specific masks -- that is what the p39 LIVE
// harness (m7-f11-date-masks) checks. '[h]:mm:ss' (a duration, not a date) stays OUT on purpose:
// a maintainer's own scope line excludes it, a separate backlog item.
const DATE_MASKS = Object.freeze({ date: ['yyyy-mm-dd', 'dd/mm/yyyy', '[$-409]mmmm d, yyyy', '[$-40E]yyyy. mmmm d.'], datetime: ['yyyy-mm-dd hh:mm', 'dd/mm/yyyy hh:mm'], time: ['hh:mm'] });
const DAY_MS = 86400000;
const EPOCH = Date.UTC(1899, 11, 30);

/** A typed value -> { v, nf, order, kind } or { problem }. order 'nf-first' (text) | 'value-first' (dates, as measured). */
function typedValue(t) {
  const keys = Object.keys(t).sort().join(',');
  if (keys === 'text') return textValue(t.text);
  if (keys === 'date,format') return dateValue(t.date, t.format);
  if (keys === 'format,time') return timeValue(t.time, t.format);
  return { problem: 'a typed value is { text } | { date, format } | { time, format }' };
}

function textValue(text) {
  if (typeof text !== 'string' || text === '') return { problem: 'text must be a non-empty string' };
  if (text.length > MAX_TEXT) return { problem: `text longer than ${MAX_TEXT} characters` };
  // a text '=x' reads back with GetFormula '=x', which the observation cannot tell from a formula (measured); the
  // + - @ prefixes are not measured -- all four stay refused
  if (/^[=+\-@]/.test(text)) return { problem: 'a text starting with = + - @ is not supported (read back as a formula / not measured)' };
  return { v: text, nf: '@', order: 'nf-first', kind: 'text' };
}

function dateValue(date, format) {
  const m = /^(\d{4})-(\d{2})-(\d{2})(?:T(\d{2}):(\d{2}))?$/.exec(String(date));
  if (!m) return { problem: "date must be 'YYYY-MM-DD' or 'YYYY-MM-DDTHH:MM'" };
  const [y, mo, d, h, mi] = m.slice(1).map(x => (x === undefined ? undefined : Number(x)));
  const ms = Date.UTC(y, mo - 1, d);
  const back = new Date(ms);
  if (y < 1900 || back.getUTCMonth() !== mo - 1 || back.getUTCDate() !== d || (h !== undefined && (h > 23 || mi > 59))) return { problem: `date ${date} is not a valid date after 1900` };
  const masks = h === undefined ? DATE_MASKS.date : DATE_MASKS.datetime;
  if (!masks.includes(format)) return { problem: `format for this date must be one of the measured masks: ${masks.join(' | ')}` };
  return { v: (ms - EPOCH) / DAY_MS + (h === undefined ? 0 : (h * 60 + mi) / 1440), nf: format, order: 'value-first', kind: 'date' };
}

function timeValue(time, format) {
  const m = /^(\d{2}):(\d{2})$/.exec(String(time));
  if (!m || Number(m[1]) > 23 || Number(m[2]) > 59) return { problem: "time must be 'HH:MM'" };
  if (!DATE_MASKS.time.includes(format)) return { problem: `format for a time must be one of: ${DATE_MASKS.time.join(' | ')}` };
  return { v: (Number(m[1]) * 60 + Number(m[2])) / 1440, nf: format, order: 'value-first', kind: 'date' };
}

function valueProblem(v) {
  if (typeof v === 'number') return Number.isFinite(v) ? null : 'not a finite number';
  if (typeof v !== 'string') return `unsupported value type ${v === null ? 'null' : typeof v} (values only: finite numbers and plain text)`;
  if (v === '') return 'empty string (clearing is a separate operation)';
  if (v.length > MAX_TEXT) return `longer than ${MAX_TEXT} characters`;
  if (v !== v.trim()) return 'leading or trailing whitespace';
  for (const [re, why] of COERCIBLE) if (re.test(v)) return why;
  return null;
}

function plan(step, index) {
  const unknown = Object.keys(step).filter(k => !FIELDS.includes(k));
  if (unknown.length) return { ok: false, detail: `step ${index}: write_range does not support field(s): ${unknown.join(', ')}` };
  if (typeof step.sheet !== 'string' || step.sheet === '') return { ok: false, detail: `step ${index}: write_range requires a sheet name` };
  const g = parseRange(step.range);
  if (!g) return { ok: false, detail: `step ${index}: write_range requires a plain A1 range on the named sheet (no Sheet!A1)` };
  if (g.count > MAX_CELLS) return { ok: false, detail: `step ${index}: write_range range has ${g.count} cells, over the ${MAX_CELLS} ceiling` };
  const isMatrix = m => Array.isArray(m) && m.length === g.rows && m.every(row => Array.isArray(row) && row.length === g.cols);
  const vals = step.values;
  const fs = step.formulas;
  if (fs !== undefined && !isMatrix(fs)) return { ok: false, detail: `step ${index}: write_range formulas must be a ${g.rows}x${g.cols} matrix matching ${rangeText(g)}` };
  if (!(fs !== undefined && vals === undefined) && !isMatrix(vals)) {
    return { ok: false, detail: `step ${index}: write_range values must be a ${g.rows}x${g.cols} matrix matching ${rangeText(g)}` };
  }
  const addresses = cellsOf(g);
  const cells = [];
  for (let r = 0; r < g.rows; r += 1) {
    for (let c = 0; c < g.cols; c += 1) {
      const a = addresses[r * g.cols + c];
      const cell = planCell(a, fs ? fs[r][c] : null, vals ? vals[r][c] : null);
      if (cell.detail) return { ok: false, detail: `step ${index}: write_range ${cell.detail}` };
      cells.push(cell);
    }
  }
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`,
      intent: 'write_range',
      selector: Object.freeze({ sheet: step.sheet, range: rangeText(g), uses: usesOf(cells) }),
      cells: Object.freeze(cells),
      footprint: Object.freeze({ value: rangeText(g) }), // M7 S3: what the step modifies, per axis
      preservationMode: 'PRIMITIVE_INTEGRAL',
    },
  };
}

/** One planned cell -- a formula, a typed value (M7 S26/S28) or a plain value -- or { detail } why it is refused. */
function planCell(a, f, v) {
  if (f !== null && f !== undefined) {
    if (v !== null && v !== undefined) return { detail: `cell ${a} has both a value and a formula (the value under a formula must be null)` };
    const why = formulaProblem(f);
    return why ? { detail: `formula at ${a} refused before dispatch: ${why}` } : Object.freeze({ a, formula: f });
  }
  if (v !== null && typeof v === 'object' && !Array.isArray(v)) {
    const t = typedValue(v);
    return t.problem ? { detail: `value at ${a} refused before dispatch: ${t.problem}` } : Object.freeze({ a, v: t.v, nf: t.nf, order: t.order, kind: t.kind });
  }
  const why = valueProblem(v);
  return why ? { detail: `value at ${a} refused before dispatch: ${why}` } : Object.freeze({ a, v });
}

/** The defined names the written formulas use ('name:<lower>'); an unparseable formula may use any ('name:*'). */
function usesOf(cellList) {
  const out = new Set();
  for (const c of cellList) {
    if (!c.formula) continue;
    const u = nameUses(c.formula);
    if (u.problem) out.add('name:*');
    else for (const n of u.names) out.add(`name:${n.toLowerCase()}`);
  }
  return Object.freeze([...out]);
}

const canonicalNumber = s => typeof s === 'string' && /^-?(?:0|[1-9]\d*)(?:\.\d+)?(?:[eE][+-]?\d+)?$/.test(s) && Number.isFinite(Number(s));

/** Does a live cell hold exactly the requested value (or formula)? null = unreadable (count not read). */
function matchesCell(cell, want) {
  if (want.formula !== undefined) return !!cell && formula.sameFormula(want.formula, cell.f);
  return want.nf === undefined ? matches(cell, want.v) : matchesTyped(cell, want);
}

/** M7 S26/S28: the exact text (COUNT 0) or the serial (COUNT 1), AND the declared number format. */
function matchesTyped(cell, want) {
  if (!cell || cell.f !== '' || !('nf' in cell) || cell.nf !== want.nf) return false;
  if (cell.count === null) return null;
  if (want.kind === 'text') return cell.count === 0 && String(cell.v) === want.v;
  const n = typeof cell.v === 'number' ? cell.v : (typeof cell.v === 'string' && cell.v.trim() !== '' ? Number(cell.v) : NaN);
  return cell.count === 1 && Math.abs(n - want.v) < 1e-8; // the read-back serial carries 15 significant digits
}

function matches(cell, v) {
  if (!cell || cell.f !== '') return false;
  if (cell.count === null) return null;
  if (typeof v === 'number') return cell.count === 1 && (typeof cell.v === 'number' ? cell.v === v : canonicalNumber(cell.v) && Number(cell.v) === v);
  return cell.count === 0 && String(cell.v) === v;
}

function sheetOf(primary, name) {
  const exact = primary.sheets.filter(s => s.name === name);
  if (exact.length === 1) return { ok: true, sheet: exact[0] };
  const folded = primary.sheets.filter(s => s.name.toLowerCase() === String(name).toLowerCase());
  return { ok: false, ambiguity: exact.length > 1 ? 'AMBIGUOUS' : folded.length ? 'SHEET_NAME_CASE_MISMATCH' : 'NOT_FOUND', available: primary.sheets.map(s => s.name) };
}

/** The fail-closed gates shared by classification and targeting, on a FRESH observation. */
const want = (op, a) => op.cells.find(c => c.a === a) || {};

function gate(primary, op) {
  if (!primary || primary.format !== 'xlsx') return { ok: false, reason: 'xlsx-observation-missing' };
  if (!primary.complete) return { ok: false, reason: primary.truncated ? 'xlsx-observation-truncated' : 'xlsx-observation-unreadable-getters', unsupported: primary.unsupported };
  const s = sheetOf(primary, op.selector.sheet);
  if (!s.ok) return { ok: false, reason: `xlsx-sheet-${s.ambiguity.toLowerCase().replace(/_/g, '-')}`, available: s.available };
  const cells = {};
  for (const { a } of op.cells) {
    const cell = cellAt(primary, op.selector.sheet, a);
    if (!cell) return { ok: false, reason: 'xlsx-target-cell-unread', cell: a };
    if (cell.merge && cell.merge !== a) return { ok: false, reason: 'xlsx-write-into-merged-area-refused', cell: a, mergeArea: cell.merge };
    if (cell.f !== '' && !(want(op, a).formula !== undefined && formula.sameFormula(want(op, a).formula, cell.f))) {
      return { ok: false, reason: 'xlsx-write-over-formula-refused', cell: a, formula: cell.f };
    }
    if (want(op, a).nf !== undefined && !('nf' in cell)) return { ok: false, reason: 'xlsx-target-format-unread', cell: a };
    cells[a] = cell;
  }
  // a sheet reference must match a live sheet name EXACTLY: the editor silently drops a formula whose
  // sheet name differs only in case (set b, =monthly_plan!B16 -> blank, no error value)
  const names = new Set(primary.sheets.map(x => x.name));
  for (const c of op.cells) {
    if (c.formula === undefined) continue;
    const missing = formula.sheetRefs(c.formula).names.filter(n => !names.has(n));
    if (missing.length) return { ok: false, reason: 'xlsx-formula-sheet-reference-not-exact', cell: c.a, references: missing, available: [...names] };
  }
  return { ok: true, cells };
}

function classify(primary, op) {
  const g = gate(primary, op);
  if (!g.ok) return { state: 'FAIL_CLOSED', evidence: g };
  const verdicts = op.cells.map(c => matchesCell(g.cells[c.a], c));
  if (verdicts.includes(null)) return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-cell-type-unreadable' } };
  if (verdicts.every(Boolean)) return { state: 'SATISFIED', evidence: { sheet: op.selector.sheet, range: op.selector.range, proof: 'fresh-live-cells-already-terminal' } };
  return { state: 'APPLY', evidence: { sheet: op.selector.sheet, range: op.selector.range, differing: verdicts.filter(x => !x).length } };
}

function resolve(primary, selector, ctx = {}) {
  const op = ctx.operation;
  const g = gate(primary, op);
  if (!g.ok) return { ok: false, target: null, selector, ambiguity: g.reason, detail: g };
  return { ok: true, target: Object.freeze({ sheet: selector.sheet, range: selector.range, before: g.cells }) };
}

// M7 S26/S28: the measured order per typed cell ('@' first for text; the value first for a date mask)
const TYPED_LOOP = `for (const p of plan) {
    const r = ws.GetRange(p.a);
    if (p.order === 'nf-first') { r.SetNumberFormat(p.nf); r.SetValue(p.v); }
    else if (p.order === 'value-first') { r.SetValue(p.v); r.SetNumberFormat(p.nf); }
    else r.SetValue(p.v);
  }`;

/**
 * One callCommand for the whole range: an exact-name sheet check, then SetValue per cell (the source's
 * measured write path). `meta` describes the same intent for test transports that cannot run the
 * editor API; the real driver executes `body` only.
 */
function lower(op) {
  const sheet = op.selector.sheet;
  const cells = op.cells.map(c => (c.formula !== undefined ? { a: c.a, v: c.formula } : (c.nf !== undefined ? { a: c.a, v: c.v, nf: c.nf, order: c.order } : { a: c.a, v: c.v })));
  const typed = op.cells.some(c => c.nf !== undefined); // without a typed value the body is the M6 one, byte for byte
  return {
    kind: 'callCommand',
    body: `
try {
  const name = ${JSON.stringify(sheet)};
  const ws = Api.GetSheet(name);
  if (!ws || ws.GetName() !== name) return { ok: false, stage: 'sheet-identity' };
  const plan = ${JSON.stringify(cells)};
  ${typed ? TYPED_LOOP : 'for (const p of plan) ws.GetRange(p.a).SetValue(p.v);'}
  return { ok: true, written: plan.length };
} catch (err) { return { ok: false, stage: 'threw', error: String(err) }; }
`,
    meta: Object.freeze({ family: 'cells', action: 'write', sheet, cells: Object.freeze(cells) }),
  };
}

function verifyPrimitive(op, _before, after) {
  const primary = after && after.primary;
  const wrong = op.cells.filter(c => matchesCell(cellAt(primary, op.selector.sheet, c.a), c) !== true).map(({ a }) => a);
  return { ok: wrong.length === 0, evidence: { sheet: op.selector.sheet, range: op.selector.range, cells: op.cells.length, wrong, proof: 'fresh-live-cell-readback' } };
}

/**
 * L3: nothing outside the target range changed. The sheet list is identical (names, order); every
 * cell read before is still there with the same formula, and a non-formula cell keeps its value; no
 * cell outside the target appeared with content. Both observations must be complete.
 */
/**
 * A cell read before and not read after is unchanged only if it was blank before AND lies outside its
 * sheet's used range after: the driver reads every used range whole, so a cell outside it holds no
 * content. This happens when a clear removes the only content that stretched the used range (W3 LIVE,
 *: clearing a formula cell at J1 left 41 once-read blank cells of column J unread -> a false
 * L3 FAIL). A cell that HAD content, or still sits inside the used range, stays 'unreadable-after'.
 */
function droppedBlank(after, k, cell) {
  if (!isBlank(cell)) return false;
  const [sheet, a] = k.split('\u0000');
  const s = after.sheets.find(x => x.name === sheet);
  return !!s && (s.used === null || !overlaps(parseRange(s.used), parseRange(a)));
}

function verifyPreservation(op, before, after) {
  const b = before && before.primary;
  const a = after && after.primary;
  if (!b || !a || !b.complete || !a.complete) return { ok: false, evidence: { reason: 'xlsx-preservation-observation-incomplete' } };
  const names = x => x.sheets.map(s => s.name).join('\u0000');
  if (names(b) !== names(a)) return { ok: false, evidence: { reason: 'xlsx-sheet-list-changed', before: b.sheets.map(s => s.name), after: a.sheets.map(s => s.name) } };
  const tierGap = tierChange(b, a);
  if (tierGap) return tierGap;
  const target = new Set(op.cells.map(c => key(op.selector.sheet, c.a)));
  const changed = [];
  for (const [k, cell] of Object.entries(b.cells)) {
    if (target.has(k)) continue;
    const now = a.cells[k];
    if (!now) { if (!droppedBlank(a, k, cell)) changed.push({ cell: k.replace('\u0000', '!'), change: 'unreadable-after' }); continue; }
    if (now.f !== cell.f) changed.push({ cell: k.replace('\u0000', '!'), change: 'formula' });
    else if (cell.f === '' && JSON.stringify(now.v) !== JSON.stringify(cell.v)) changed.push({ cell: k.replace('\u0000', '!'), change: 'value' });
  }
  for (const [k, cell] of Object.entries(a.cells)) {
    if (target.has(k) || b.cells[k]) continue;
    if (cell.f !== '' || (cell.v !== null && cell.v !== '')) changed.push({ cell: k.replace('\u0000', '!'), change: 'appeared' });
  }
  return withDigest(b, a, { ok: changed.length === 0, evidence: { compared: Object.keys(b.cells).length - target.size, changed: changed.slice(0, 20), changedCount: changed.length } });
}

/** M7 S32 slice 1: both observations must be of ONE tier (the tier is fixed per task by its bound). */
function tierChange(b, a) {
  const tb = b.tier || 'full';
  const ta = a.tier || 'full';
  return tb === ta ? null : { ok: false, evidence: { reason: 'xlsx-preservation-tier-changed', before: tb, after: ta } };
}

/**
 * Tiered: the cell-level result above covers the task's windows (cell by cell -- every OTHER window of the task
 * included); the content outside every window is proven by the per-sheet digest, identical before and after.
 */
function withDigest(b, a, cellLevel) {
  if (b.tier !== 'tiered') return cellLevel;
  const digest = digestChanges(b.digests, a.digests);
  return { ok: cellLevel.ok && digest.changed.length === 0, evidence: { ...cellLevel.evidence, tier: 'tiered', digest } };
}

/**
 * The digests of two tiered observations: per sheet, the same excluded windows, entry count and SHA-256. They say
 * WHETHER the content outside the windows changed, not where.
 */
function digestChanges(before, after) {
  const changed = [];
  for (const d of before) {
    const n = after.find(x => x.sheet === d.sheet);
    if (!n) { changed.push({ sheet: d.sheet, change: 'digest-missing-after' }); continue; }
    if (JSON.stringify(n.excluded) !== JSON.stringify(d.excluded)) changed.push({ sheet: d.sheet, change: 'excluded-windows-differ', before: d.excluded, after: n.excluded });
    else if (n.hash !== d.hash || n.count !== d.count) changed.push({ sheet: d.sheet, change: 'content-outside-windows', before: { count: d.count, hash: d.hash }, after: { count: n.count, hash: n.hash } });
  }
  return { compared: before.length, changed };
}

/**
 * write_range's own L3 (M7 S3): the content L3 above, AND -- wherever the task reads style (a composed format / layout
 * step asks for it) -- every style key observed on both sides unchanged, the written cells included: a write changes
 * values, never style. A standalone write reads no style and claims nothing about it (the M6 contract).
 */
function writePreservation(op, before, after) {
  const rest = verifyPreservation(op, before, after);
  if (!rest.ok && rest.evidence && rest.evidence.reason) return rest;
  const typed = new Set(op.cells.filter(c => c.nf !== undefined).map(c => key(op.selector.sheet, c.a)));
  const s = observedStyleChanges(before.primary, after.primary);
  const style = s.changed.filter(x => !(typed.has(x.cell.replace('!', '\u0000')) && x.keys && x.keys.every(k => k === 'nf')));
  const nf = typed.size ? nfChanges(before.primary, after.primary, typed) : [];
  return { ok: rest.ok && style.length === 0 && nf.length === 0, evidence: { ...rest.evidence, styleCompared: s.compared, styleChanged: style.slice(0, 20), styleChangedCount: style.length, ...(typed.size ? { declaredNumberFormat: typed.size, nfChanged: nf.slice(0, 20) } : {}) } };
}

/** M7 S26/S28: every number format read on both sides, outside the typed cells, unchanged. */
function nfChanges(b, a, typed) {
  const out = [];
  for (const [k, was] of Object.entries(b.cells || {})) {
    if (typed.has(k) || !('nf' in was)) continue;
    const now = a.cells && a.cells[k];
    if (!now || !('nf' in now) || now.nf !== was.nf) out.push({ cell: k.replace('\u0000', '!'), was: was.nf, now: now ? now.nf : 'unread' });
  }
  return out;
}

/**
 * The write is proven by identity; the EVALUATION is a separate check (owner, W2 GO): a written formula
 * whose value is an error (#DIV/0!, #NAME?, #REF! ...) fails the whole task, named per cell.
 */
function wholeTaskChecks(op, finalObservation) {
  const v = verifyPrimitive(op, null, finalObservation);
  const checks = [{ name: `every cell of ${op.selector.sheet}!${op.selector.range} holds the requested value or formula`, ok: v.ok, detail: v.ok ? null : { wrong: v.evidence.wrong } }];
  const fcells = op.cells.filter(c => c.formula !== undefined);
  if (fcells.length) {
    const primary = finalObservation && finalObservation.primary;
    const errors = fcells.map(c => ({ a: c.a, value: (cellAt(primary, op.selector.sheet, c.a) || {}).v })).filter(x => typeof x.value === 'string' && ERROR_VALUE.test(x.value));
    checks.push({ name: `no formula in ${op.selector.sheet}!${op.selector.range} evaluates to an error value`, ok: errors.length === 0, detail: errors.length ? { evaluationErrors: errors } : null });
  }
  return checks;
}

/** The windows a live read must cover for this operation (the target range itself). */
// a typed value reads the number format of its OWN cells too -- empty target cells outside the used range are read
// through the window only (LIVE m7-d-acc1: formatSheets alone left A2 unread)
const windowsOf = op => [{ sheet: op.selector.sheet, range: op.selector.range, ...(op.cells.some(c => c.nf !== undefined) ? { format: true } : {}) }];

/** A typed value needs the number format read on its sheet (the public GetNumberFormat, W4a). */
const liveExtras = op => (op.cells.some(c => c.nf !== undefined) ? { formatSheets: [op.selector.sheet] } : {});

module.exports = { liveExtras, typedValue, DATE_MASKS, plan, classify, resolve, lower, verifyPrimitive, verifyPreservation, writePreservation, wholeTaskChecks, windowsOf, valueProblem, formulaProblem, matches, MAX_CELLS };
