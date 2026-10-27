'use strict';

const { colNum, colName, parseRange, rangeText } = require('../a1.cjs');
const { key } = require('../inventory.cjs');
const { refOnTarget } = require('../dependencies.cjs');
const { moveIndex, mapRefs } = require('../shift.cjs');

/**
 * XLSX `delete-cells` family: delete_cells (L-18/1). docs/L18-DELETE-CELLS-SHIFT-PLAN.md -- the LIVE
 * feasibility probe (5 rounds, test/live/xlsx-l18-delete-cells-shift-probe.live.cjs) measured that the
 * public `ApiRange.Delete('up'|'left')`,
 * called on a SUB-RANGE block (not a whole row/column, unlike the `structure` family's insert/delete_rows/columns),
 * stays SCOPED to the block's own column band ('up') / row band ('left') -- everything outside that band, on the
 * SAME sheet, is untouched. A formula pointing directly at a cell the block destroys correctly turns into `#REF!`
 * (measured, `refBroken` in the probe); a formula in the shift path, referencing a cell BEHIND the block, is
 * correctly re-pointed (measured: `=C4+1` at C5 -> `=C2+1` at C3 after an up-delete of C2:D3).
 *
 * A merge the block fully covers is deleted WITH it (measured: DELETED_WITH_BLOCK -- not a refusal, same as the
 * `structure` family's "entirely inside a deleted block goes with it" path). A merge the block only PARTIALLY
 * covers is, measured, SILENTLY TRUNCATED by the DS (a 2-row merge C3:D4 straddling an up-delete of C2:D3 became a
 * 1-row C2:D2) -- Excel itself refuses to edit part of a merge, so this family refuses it too, named,
 * PRE_DISPATCH (`xlsx-delete-cells-merge-partial-overlap`): the DS gives no protection of its own here. A merge
 * further along the SHIFT PATH (below the block for 'up', to its right for 'left') that overlaps the block's own
 * band -- whether crossing the band's edge or sitting entirely within it -- is a SECOND, unmeasured shape:
 * refused named, PRE_DISPATCH (`xlsx-delete-cells-merge-crosses-band`). Only a shift-path merge
 * entirely OUTSIDE the band is inert and left alone.
 *
 * RETRY SAFETY: a delete is not naturally idempotent -- a second identical call
 * would shift AGAIN, destroying more data. This family reuses the `structure` family's OWN, LIVE-proven mechanism
 * (M6 wave 5, insert_rows/delete_rows/insert_columns/delete_columns) rather than inventing a new one: the caller
 * names a content ANCHOR -- a value that lives in a column (shift 'up') or row (shift 'left') WITHIN the block's
 * own band -- and classify() reads it FRESH, from the SAME session the write itself uses. If the anchor is already
 * at its TERMINAL position (where the shift would put it), the step is SATISFIED with zero writes -- the exact
 * shape a retry needs. If it is at its pre-shift ("now") position, the step APPLIES. Anything else is a named,
 * pre-dispatch refusal (`xlsx-delete-cells-anchor-state-conflict`) -- never a guess.
 *
 * Scope the probe did NOT measure, refused named, PRE_DISPATCH, blanket (not the `structure` family's fine
 * per-cut allowances -- those are each their OWN LIVE measurement this family has none of yet): any defined name,
 * validation, conditional format, chart, pivot or AutoFilter reaching the sheet AT ALL
 * (`xlsx-delete-cells-dependent-not-measured`); a formula reference of any shape other than a bare, same-sheet
 * CELL reference (`xlsx-delete-cells-formula-shape-not-measured`) -- a RANGE/row/column reference into the
 * shifting band is refused rather than guessed. Large workbooks: this family is not in the adapter's
 * TIERED_INTENTS and carries no bounded variant, so above the read bound it refuses generically
 * (IMPLEMENTATION_LIMIT, the same mechanism filter_range's own large-book boundary already uses -- not a guard
 * this file writes itself). `Insert` (L-18/2) is out of scope.
 */

const INTENTS = Object.freeze(['delete_cells']);
const FIELDS = Object.freeze(['intent', 'id', 'sheet', 'range', 'shift', 'anchor']);
const MAX_CELLS = 200;

function plan(step, index) {
  const at = `step ${index}: ${step.intent}`;
  const unknown = Object.keys(step).filter(k => !FIELDS.includes(k));
  if (unknown.length) return { ok: false, detail: `${at} does not support field(s): ${unknown.join(', ')}` };
  if (typeof step.sheet !== 'string' || step.sheet === '') return { ok: false, detail: `${at} requires a sheet name` };
  const g = parseRange(step.range);
  if (!g) return { ok: false, detail: `${at} requires a plain A1 range` };
  if (g.count > MAX_CELLS) return { ok: false, detail: `${at} range has ${g.count} cells, over the ${MAX_CELLS} ceiling` };
  if (!['up', 'left'].includes(step.shift)) return { ok: false, detail: `${at}: shift must be 'up' or 'left'` };
  const rows = step.shift === 'up';
  const a = step.anchor;
  if (!a || typeof a !== 'object' || Array.isArray(a)) return { ok: false, detail: `${at}: anchor must be an object` };
  const along = rows ? 'column' : 'row';
  const other = rows ? 'row' : 'column';
  if (a[other] !== undefined) return { ok: false, detail: `${at}: anchor.${other} is not used for shift '${step.shift}' (use anchor.${along})` };
  const lane = rows ? (typeof a.column === 'string' && /^[A-Z]{1,3}$/.test(a.column) ? colNum(a.column) : null) : (Number.isInteger(a.row) && a.row >= 1 ? a.row : null);
  if (lane === null) return { ok: false, detail: `${at}: anchor.${along} must be ${rows ? 'a column letter' : 'a row number'}` };
  if (rows ? (lane < g.c1 || lane > g.c2) : (lane < g.r1 || lane > g.r2)) return { ok: false, detail: `${at}: anchor.${along} must be within the block's own ${along} band (${rows ? `${colName(g.c1)}-${colName(g.c2)}` : `${g.r1}-${g.r2}`})` };
  if (!((typeof a.value === 'string' && a.value !== '') || Number.isFinite(a.value))) return { ok: false, detail: `${at}: anchor.value must be a non-empty text or a number` };
  const height = g.r2 - g.r1 + 1;
  const width = g.c2 - g.c1 + 1;
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`,
      intent: step.intent,
      family: 'delete-cells',
      selector: Object.freeze({ sheet: step.sheet, range: rangeText(g) }),
      block: Object.freeze({ sheet: step.sheet, r1: g.r1, r2: g.r2, c1: g.c1, c2: g.c2 }),
      shift: Object.freeze({ sheet: step.sheet, axis: rows ? 'rows' : 'cols', at: rows ? g.r1 : g.c1, count: -(rows ? height : width) }),
      band: Object.freeze(rows ? { lo: g.c1, hi: g.c2 } : { lo: g.r1, hi: g.r2 }),
      anchor: Object.freeze({ lane, value: String(a.value), now: (rows ? g.r2 : g.c2) + 1, terminal: rows ? g.r1 : g.c1 }),
      cells: Object.freeze([]),
      preservationMode: 'PRIMITIVE_INTEGRAL',
    },
  };
}

const rowOf = a => Number(/\d+$/.exec(a)[0]);
const colOf = a => colNum(/^[A-Z]+/.exec(a)[0]);
const cellsOfSheet = (p, sheet) => Object.entries(p.cells).filter(([k]) => k.startsWith(key(sheet, ''))).map(([k, c]) => [k.slice(key(sheet, '').length), c]);

/** Where the anchor value is now: every literal cell in the anchor's column (shift 'up') / row ('left'). */
function anchorHits(p, op) {
  const rows = op.shift.axis === 'rows';
  return cellsOfSheet(p, op.shift.sheet).filter(([a, c]) => c.f === '' && String(c.v) === op.anchor.value && (rows ? colOf(a) === op.anchor.lane : rowOf(a) === op.anchor.lane)).map(([a]) => (rows ? rowOf(a) : colOf(a)));
}

function gate(p, op) {
  if (!p || p.format !== 'xlsx') return { ok: false, reason: 'xlsx-observation-missing' };
  if (!p.complete) return { ok: false, reason: p.truncated ? 'xlsx-observation-truncated' : 'xlsx-observation-unreadable-getters', unsupported: p.unsupported };
  const s = p.sheets.filter(x => x.name === op.shift.sheet);
  if (s.length !== 1) return { ok: false, reason: s.length > 1 ? 'xlsx-sheet-ambiguous' : 'xlsx-sheet-not-found', available: p.sheets.map(x => x.name) };
  const hits = anchorHits(p, op);
  if (hits.length !== 1) return { ok: false, reason: hits.length ? 'xlsx-delete-cells-anchor-not-unique' : 'xlsx-delete-cells-anchor-not-found', anchor: op.anchor.value, at: hits.slice(0, 20) };
  return { ok: true, pos: hits[0] };
}

/** Does `area` (a merge range) overlap `block` at all? */
function overlapsBlock(area, block) {
  const g = parseRange(area);
  return g.r1 <= block.r2 && g.r2 >= block.r1 && g.c1 <= block.c2 && g.c2 >= block.c1;
}

/** Is `area` ENTIRELY inside `block`? */
function insideBlock(area, block) {
  const g = parseRange(area);
  return g.r1 >= block.r1 && g.r2 <= block.r2 && g.c1 >= block.c1 && g.c2 <= block.c2;
}

/**
 * Measured (round 5, M7 S5 R5-style naming): a merge entirely inside the deleted block goes with it
 * (DELETED_WITH_BLOCK -- no refusal). A merge that only partially overlaps the block -- or that overlaps the
 * block's band on the OTHER axis without being fully contained (an unmeasured shape) -- is silently truncated by
 * the DS and MUST be refused (Excel itself refuses to edit part of a merge).
 *
 * A later review caught a SECOND, independent gap this didn't cover: a merge that
 * never touches the deleted block at all, but sits further along the SHIFT PATH (below the block for 'up',
 * to its right for 'left') and overlaps the block's own band -- e.g. a B6:C6 merge when the block is C2:D3
 * ('up', band columns C:D): B6:C6 never overlaps rows 2-3, so the check above never sees it, but the shift
 * moves C6 (in-band) up while leaving B6 (out-of-band) in place, tearing the merge apart exactly the same way
 * a block-straddling merge does. A merge ENTIRELY INSIDE the band in the shift path (would shift as a whole,
 * intact) is a DIFFERENT, equally unmeasured shape -- v1 refuses it too, named the same way, until a LIVE
 * harness proves it shifts correctly. Only a shift-path merge entirely OUTSIDE the band is genuinely inert
 * (never touched, same as any other out-of-band cell) and is left alone.
 */
function inShiftPath(area, block, rows) {
  const g = parseRange(area);
  return rows ? g.r1 > block.r2 : g.c1 > block.c2;
}
function overlapsBand(area, band, rows) {
  const g = parseRange(area);
  return rows ? (g.c1 <= band.hi && g.c2 >= band.lo) : (g.r1 <= band.hi && g.r2 >= band.lo);
}
function mergeGuard(p, op) {
  const sheet = op.shift.sheet;
  const rows = op.shift.axis === 'rows';
  for (const [a, c] of cellsOfSheet(p, sheet)) {
    if (!c.merge || c.merge === a) continue;
    if (insideBlock(c.merge, op.block)) continue;
    if (overlapsBlock(c.merge, op.block)) return { reason: 'xlsx-delete-cells-merge-partial-overlap', cell: a, mergeArea: c.merge };
    if (inShiftPath(c.merge, op.block, rows) && overlapsBand(c.merge, op.band, rows)) return { reason: 'xlsx-delete-cells-merge-crosses-band', cell: a, mergeArea: c.merge };
  }
  return null;
}

/** Blanket, PRE_DISPATCH: refuse if ANY defined name, rule (validation/CF), chart, pivot or AutoFilter reaches the
 * sheet at all -- none of those axes has a measurement for a block-shift yet (unlike `structure`'s per-cut guards). */
function dependentGuard(p, op) {
  const sheet = op.shift.sheet;
  if (!p.dependencies || !p.dependencies.complete) return { reason: 'xlsx-delete-cells-dependency-observation-incomplete', unobservable: p.dependencies ? p.dependencies.unobservable.slice(0, 20) : 'not-read' };
  const ref = refOnTarget(p.dependencies, [sheet]);
  if (ref) return ref;
  const reach = p.dependencies.refs.filter(r => r.kind !== 'formula' && (r.sheets.includes(sheet) || r.host === sheet));
  if (reach.length) return { reason: 'xlsx-delete-cells-dependent-not-measured', dependents: reach.slice(0, 20), count: reach.length };
  if (!p.rules || !p.rules.complete || !p.rules.bySheet) return { reason: 'xlsx-delete-cells-rule-observation-incomplete', unobservable: p.rules ? p.rules.unobservable.slice(0, 20) : 'not-read' };
  if ((p.rules.bySheet[sheet] || []).length) return { reason: 'xlsx-delete-cells-dependent-not-measured', rules: sheet };
  const objects = p.objects && p.objects[sheet];
  if (objects && ((Array.isArray(objects.charts) && objects.charts.length) || (Array.isArray(objects.pivots) && objects.pivots.length))) return { reason: 'xlsx-delete-cells-dependent-not-measured', objects: sheet };
  const fr = p.filterReads && p.filterReads[sheet];
  if (fr === undefined) return { reason: 'xlsx-delete-cells-filter-observation-incomplete', sheet };
  if (fr && typeof fr === 'object' && fr.range) return { reason: 'xlsx-delete-cells-dependent-not-measured', autoFilter: sheet };
  return null;
}

/**
 * A formula carrying any reference shape other than EXACTLY ONE unqualified, same-sheet CELL reference is
 * not measured for a block-shift -- refused named rather than guessed. The measured shapes themselves are
 * NOT bare cell references alone (`=C4+1`, `=C3`) -- the constraint is on the REFERENCE (one, cell-kind,
 * unqualified), not on the formula text as a whole; arbitrary surrounding arithmetic is fine, exactly as
 * measured. Scans EVERY cell of the WHOLE workbook (`p.cells`, like `structure.cjs`'s own block-reference
 * scan) so a formula hosted on ANOTHER sheet that reaches INTO the shifting sheet is caught too -- scoping
 * the scan to `sheet` alone (an earlier draft's bug) would have missed exactly that cross-sheet case, the
 * only one `structure.cjs`'s own precedent explicitly widens for.
 */
function formulaShapeGuard(p, op) {
  const sheet = op.shift.sheet;
  const bad = [];
  for (const [k, c] of Object.entries(p.cells)) {
    if (!c.f) continue;
    const host = k.split('\u0000')[0];
    const addr = k.split('\u0000')[1];
    const text = c.f.startsWith('=') ? c.f : `=${c.f}`;
    let refCount = 0; let allCell = true; let crossesToSheet = false;
    const parsed = mapRefs(text, ref => {
      refCount += 1;
      if (ref.kind !== 'cell') allCell = false;
      if (ref.sheet !== null && String(ref.sheet).toUpperCase() === sheet.toUpperCase()) crossesToSheet = true;
      return 'x';
    });
    if (host === sheet) {
      // same-sheet: exactly one unqualified cell-kind reference -- the measured shapes (=C3, =C4+1)
      if (parsed === null || refCount !== 1 || !allCell || crossesToSheet) bad.push(addr);
    } else if (crossesToSheet) {
      // another sheet's formula reaching INTO the shifting sheet: not measured at all, any shape
      bad.push(`${host}!${addr}`);
    }
  }
  return bad.length ? { reason: 'xlsx-delete-cells-formula-shape-not-measured', cells: bad.slice(0, 20), count: bad.length } : null;
}

function classify(p, op) {
  const g = gate(p, op);
  if (!g.ok) return { state: 'FAIL_CLOSED', evidence: g };
  const where = { sheet: op.shift.sheet, axis: op.shift.axis, at: op.shift.at, count: op.shift.count, anchorAt: g.pos };
  if (g.pos === op.anchor.terminal) return { state: 'SATISFIED', evidence: { ...where, proof: 'fresh-live-anchor-at-its-terminal-position' } };
  if (g.pos !== op.anchor.now) return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-delete-cells-anchor-state-conflict', ...where, expectedNow: op.anchor.now, expectedAfter: op.anchor.terminal } };
  const bad = mergeGuard(p, op) || dependentGuard(p, op) || formulaShapeGuard(p, op);
  if (bad) return { state: 'FAIL_CLOSED', evidence: bad };
  return { state: 'APPLY', evidence: where };
}

function resolve(p, selector, ctx = {}) {
  const g = gate(p, ctx.operation);
  if (!g.ok) return { ok: false, target: null, selector, ambiguity: g.reason, detail: g };
  return { ok: true, target: Object.freeze({ sheet: ctx.operation.shift.sheet, anchorAt: g.pos }) };
}

/** One callCommand: exact-name sheet check, then the public block-range Delete('up'|'left') -- the LIVE-measured
 * shape (the probe's deleteBody, test/live/xlsx-l18-delete-cells-shift-probe.live.cjs). */
function lower(op) {
  const { sheet } = op.shift;
  const call = op.shift.axis === 'rows' ? "Delete('up')" : "Delete('left')";
  return {
    kind: 'callCommand',
    body: `
try {
  const name = ${JSON.stringify(sheet)};
  const ws = Api.GetSheet(name);
  if (!ws || ws.GetName() !== name) return { ok: false, stage: 'sheet-identity' };
  ws.GetRange(${JSON.stringify(rangeText(op.block))}).${call};
  return { ok: true };
} catch (err) { return { ok: false, stage: 'threw', error: String(err) }; }
`,
    meta: Object.freeze({ family: 'delete-cells', action: op.intent, sheet, range: rangeText(op.block), shift: op.shift.axis === 'rows' ? 'up' : 'left' }),
  };
}

/** Where a before-cell BEYOND the block, within the shifting band, lands (null = deleted/out of band). */
function landing(op, a) {
  const rows = op.shift.axis === 'rows';
  const r = rowOf(a);
  const c = colOf(a);
  if (rows ? (c < op.band.lo || c > op.band.hi) : (r < op.band.lo || r > op.band.hi)) return a; // outside the band: never moves
  const n = moveIndex(rows ? r : c, op.shift);
  if (n === null) return null;
  return rows ? `${colName(c)}${n}` : `${colName(n)}${r}`;
}

/**
 * The shifted text of a formula the formulaShapeGuard already proved carries EXACTLY ONE unqualified,
 * same-sheet cell reference (any surrounding arithmetic kept verbatim -- the measured shapes are `=C3`
 * AND `=C4+1`, not bare-cell-only). The reference's OWN target decides the outcome, independent of where
 * the formula CELL itself sits: a target inside the band is re-pointed by the shift (or #REF! if the shift
 * deletes it); a target outside the band never moves. This is what makes a single helper correct for BOTH
 * a shift-path formula (G5-style, itself in the band) and a cross-reference formula sitting OUTSIDE the
 * band entirely (G1 = '=C3', measured to turn into '=#REF!' when C3 is inside the deleted block) --
 * conflating "does the CELL move" with "does its REFERENCE move" was an earlier draft's bug. Returns null
 * only if the text does not parse (should not happen: the guard already refused anything else upstream).
 */
function shiftedBareCellFormula(op, text) {
  const rows = op.shift.axis === 'rows';
  const out = mapRefs(text, ref => {
    const [, cd, c, rd, r] = ref.a;
    const idx = rows ? Number(r) : colNum(c.toUpperCase());
    const inBand = rows ? (colNum(c.toUpperCase()) >= op.band.lo && colNum(c.toUpperCase()) <= op.band.hi) : (Number(r) >= op.band.lo && Number(r) <= op.band.hi);
    if (!inBand) return `${cd}${c}${rd}${r}`;
    const n = moveIndex(idx, op.shift);
    if (n === null) return '#REF!';
    return rows ? `${cd}${c}${rd}${n}` : `${cd}${colName(n)}${rd}${r}`;
  });
  return out;
}

const sameText = (a, b) => String(a).replace(/\s+/g, '').toUpperCase() === String(b).replace(/\s+/g, '').toUpperCase();

/** Does `now` hold what `was` should become after the shift? `was.f`/`now.f` formulas are compared via
 * shiftedBareCellFormula; `was.f === ''` literal cells compare their value verbatim. */
function kept(op, was, now) {
  if (!now) return false;
  if (was.f === '') return now.f === '' && JSON.stringify(now.v) === JSON.stringify(was.v);
  const want = shiftedBareCellFormula(op, was.f.startsWith('=') ? was.f : `=${was.f}`);
  return want !== null && now.f !== '' && sameText(want, now.f.startsWith('=') ? now.f : `=${now.f}`);
}

/**
 * L1: the target sheet.
 *   - every cell WITHIN the band, BEYOND the block: where the shift puts it (value kept, or formula
 *     re-pointed/#REF!'d -- `kept()`/`shiftedBareCellFormula`).
 *   - every formula cell on this sheet, REGARDLESS of its own position (OUTSIDE the band included --
 *     the G1-style cross-reference case): its text re-derived the SAME way, checked at its OWN,
 *     unmoved address (formulaShapeGuard already proved there is exactly one, same-sheet, bare-ish
 *     reference, so this is never a duplicate claim with the band loop above -- a cell cannot be both
 *     "within the band" and "outside the band").
 *   - merges moved or (fully-covered) gone; the anchor at its terminal position.
 */
function verifyPrimitive(op, before, after) {
  const b = before && before.primary;
  const af = after && after.primary;
  const sheet = op.shift.sheet;
  const rows = op.shift.axis === 'rows';
  const wrong = [];
  for (const [addr, was] of cellsOfSheet(b, sheet)) {
    const r = rowOf(addr); const c = colOf(addr);
    const inBand = rows ? (c >= op.band.lo && c <= op.band.hi) : (r >= op.band.lo && r <= op.band.hi);
    if (inBand) {
      const to = landing(op, addr);
      if (to === null) continue; // inside the deleted block itself: no literal landing to check (merge checked separately)
      const now = af.cells[key(sheet, to)];
      if ((was.f !== '' || (was.v !== null && was.v !== '')) && !kept(op, was, now)) wrong.push({ from: addr, to, was: { v: was.v, f: was.f }, now: now ? { v: now.v, f: now.f } : 'unread' });
    } else if (was.f !== '') {
      // outside the band: the cell itself never moves, but its (single, measured-shape) reference may
      const now = af.cells[key(sheet, addr)];
      if (!kept(op, was, now)) wrong.push({ from: addr, to: addr, was: { v: was.v, f: was.f }, now: now ? { v: now.v, f: now.f } : 'unread' });
    }
  }
  for (const [addr, was] of cellsOfSheet(b, sheet)) {
    if (!was.merge || was.merge === addr) continue;
    if (insideBlock(was.merge, op.block)) {
      // DELETED_WITH_BLOCK: no cell of this sheet should still report this merge's area
      const survives = cellsOfSheet(af, sheet).some(([, c2]) => c2.merge === was.merge);
      if (survives) wrong.push({ merge: was.merge, expected: 'deleted-with-block', found: 'still-present' });
    }
  }
  const hits = anchorHits(af, op);
  const anchorOk = hits.length === 1 && hits[0] === op.anchor.terminal;
  return { ok: wrong.length === 0 && anchorOk, evidence: { sheet, axis: op.shift.axis, at: op.shift.at, count: op.shift.count, wrong: wrong.slice(0, 20), wrongCount: wrong.length, anchorAt: hits, proof: 'fresh-live-every-band-cell-at-its-shifted-position' } };
}

/** L3: every cell OUTSIDE the band on this sheet (LITERAL cells only -- a formula there is covered by
 * verifyPrimitive's own out-of-band check, since its reference may legitimately change even though the
 * cell itself does not move), and every cell of every OTHER sheet, unchanged (the dependent guard already
 * refused any cross-sheet formula reaching this sheet, so no other-sheet formula needs the same carve-out). */
function verifyPreservation(op, before, after) {
  const b = before && before.primary;
  const a = after && after.primary;
  if (!b || !a || !b.complete || !a.complete) return { ok: false, evidence: { reason: 'xlsx-delete-cells-preservation-observation-incomplete' } };
  const names = x => x.sheets.map(s => s.name).join('\u0000');
  if (names(b) !== names(a)) return { ok: false, evidence: { reason: 'xlsx-sheet-list-changed' } };
  const sheet = op.shift.sheet;
  const rows = op.shift.axis === 'rows';
  const changed = [];
  let compared = 0;
  for (const [k, was] of Object.entries(b.cells)) {
    const host = k.split('\u0000')[0];
    if (host === sheet) {
      const addrPart = k.split('\u0000')[1];
      const r = rowOf(addrPart); const c = colOf(addrPart);
      const inBand = rows ? (c >= op.band.lo && c <= op.band.hi) : (r >= op.band.lo && r <= op.band.hi);
      if (inBand || was.f !== '') continue; // covered by verifyPrimitive (band cells, and every formula)
    }
    if (was.f === '' && (was.v === null || was.v === '')) continue;
    compared += 1;
    const now = a.cells[k];
    const sameCell = now && now.f === was.f && JSON.stringify(now.v) === JSON.stringify(was.v) && (was.merge === now.merge || (!was.merge && !now.merge));
    if (!sameCell) changed.push({ cell: k.replace('\u0000', '!'), was: { v: was.v, f: was.f, merge: was.merge }, now: now ? { v: now.v, f: now.f, merge: now.merge } : 'unread' });
  }
  return { ok: changed.length === 0, evidence: { compared, changed: changed.slice(0, 20), changedCount: changed.length } };
}

function wholeTaskChecks(op, finalObservation) {
  const p = finalObservation && finalObservation.primary;
  const hits = p ? anchorHits(p, op) : [];
  const ok = hits.length === 1 && hits[0] === op.anchor.terminal;
  return [{ name: `${op.shift.sheet}: the anchor '${op.anchor.value}' at ${op.shift.axis === 'rows' ? 'row' : 'column'} ${op.shift.axis === 'rows' ? op.anchor.terminal : colName(op.anchor.terminal)}`, ok, detail: ok ? null : { anchorAt: hits } }];
}

const windowsOf = () => [];
const liveExtras = () => ({ dependents: true, filterSheets: ['*'] });

module.exports = { INTENTS, plan, classify, resolve, lower, verifyPrimitive, verifyPreservation, wholeTaskChecks, windowsOf, liveExtras };
