'use strict';

const { parseRange, parseCell, cellsOf, rangeText } = require('../a1.cjs');
const { cellAt, key, isBlank } = require('../inventory.cjs');
const cells = require('./cells.cjs');

/**
 * XLSX `merge` family: merge_range (no "across") and unmerge_range (M6 wave 4).
 *
 * Measured (W4b LIVE probe, docs/evidence/M6-W4-APPEARANCE-AUDIT-AND-CONTRACT.md section 5): the public
 * `ApiRange.Merge(false)` / `UnMerge()`; the readback is the PUBLIC `MergeArea` of EVERY cell (the driver reads
 * it for every read cell). A merge over an inner merge absorbs it; an unmerge dissolves inner merges; a
 * non-blank non-top-left cell is SILENTLY emptied by a merge -- so that is refused before the write.
 *
 * The contract differs from the source (audit B-L14..L18) where the source had a defect, stated:
 *   - the readback is per cell, not the top-left's MergeArea only (B-L18: a false "already satisfied" unmerge
 *     while inner merges remained);
 *   - content guard: every non-top-left cell must be blank -- and a FORMULA is never blank, even when its
 *     result is an empty string (owner rule); an unread cell refuses;
 *   - a merge area that reaches OUTSIDE the target refuses both intents (the engine would merge / unmerge
 *     more than the request names);
 *   - L3: every other cell keeps value, formula AND its merge area.
 * Retry = the terminal state on a fresh read: every target cell reports the target area (merge) or itself
 * (unmerge) -> SATISFIED, 0 write / 0 command.
 *
 * X-L06 (behavioural LIVE probe, adc47f6): merge_range's optional `across: true` field --
 * the SAME public `Merge(true)` (not a separate method; measured, no `MergeAcross`) on the WHOLE target
 * range, but the engine splits it into ONE MERGE PER ROW, persisting across a reopen. The content guard
 * (every non-left-most cell must be blank) applies PER ROW here, not to the whole range -- the engine
 * silently empties a row's non-left-most non-blank cell exactly as it does for a plain merge (X-L07's
 * named safety boundary), so this family refuses rather than weakening. unmerge_range needs no `across`
 * of its own: UnMerge() on the whole range already dissolves every inner (per-row) merge (see above).
 *
 * X-L13/3 (owner EURO-MCP capability-backlog, matrix category (a) MIGRALHATO): migrated to the TIERED
 * observation above 45000 used cells (adapter.cjs TIERED_INTENTS) -- 45000 ALATT unchanged (the full read). Above it,
 * every check ABOVE (gate/classify/verifyPrimitive, all window-only) stays exactly as strict; only the OUTSIDE-the-
 * target merge-area re-check in verifyPreservation() below changes -- see its own comment for why and what it relies
 * on instead.
 */

const INTENTS = Object.freeze(['merge_range', 'unmerge_range']);
const MAX_CELLS = 1000;
const FIELDS = Object.freeze(['intent', 'id', 'sheet', 'range', 'across']);

/** X-L06: the per-row sub-range text for row `r` within parsed range `grid` (A1, no $). */
const rowRangeOf = (grid, r) => rangeText({ c1: grid.c1, r1: r, c2: grid.c2, r2: r, count: grid.cols });
/** X-L06: the left-most cell address of row `r` within parsed range `grid`. */
const leftMostOf = (grid, r) => rangeText({ c1: grid.c1, r1: r, c2: grid.c1, r2: r, count: 1 });

/** X-L06: why the `across` field is not acceptable on this step, or null. across needs at least two COLUMNS: a
 * single-column row has nothing to merge within itself, and its "already merged" state would be indistinguishable
 * from "never touched" (an unmerged cell's own reported area IS its own address, same as a degenerate one-cell
 * row-range) -- refused rather than guessing what Merge(true) does on a shape the X-L06 probe never exercised. */
function acrossProblem(step, g) {
  if (!('across' in step)) return null;
  if (step.intent !== 'merge_range') return `${step.intent} does not support field(s): across`;
  if (typeof step.across !== 'boolean') return 'across must be true or false';
  return step.across === true && g.cols < 2 ? 'merge_range across:true needs at least two columns (a single-column row has nothing to merge within itself)' : null;
}

function plan(step, index) {
  const unknown = Object.keys(step).filter(k => !FIELDS.includes(k));
  if (unknown.length) return { ok: false, detail: `step ${index}: ${step.intent} does not support field(s): ${unknown.join(', ')}` };
  if (typeof step.sheet !== 'string' || step.sheet === '') return { ok: false, detail: `step ${index}: ${step.intent} requires a sheet name` };
  const g = parseRange(step.range);
  if (!g) return { ok: false, detail: `step ${index}: ${step.intent} requires a plain A1 range on the named sheet (no Sheet!A1)` };
  if (g.count > MAX_CELLS) return { ok: false, detail: `step ${index}: ${step.intent} range has ${g.count} cells, over the ${MAX_CELLS} ceiling` };
  if (step.intent === 'merge_range' && g.count < 2) return { ok: false, detail: `step ${index}: merge_range needs at least two cells` };
  const acrossIssue = acrossProblem(step, g);
  if (acrossIssue) return { ok: false, detail: `step ${index}: ${acrossIssue}` };
  const range = rangeText(g);
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`,
      intent: step.intent,
      family: 'merge',
      selector: Object.freeze({ sheet: step.sheet, range }),
      cells: Object.freeze(cellsOf(g).map(a => Object.freeze({ a }))),
      topLeft: cellsOf(g)[0],
      across: step.intent === 'merge_range' ? !!step.across : false,
      preservationMode: 'PRIMITIVE_INTEGRAL',
    },
  };
}

/** Is merge area `area` (A1, no $) inside the target range? */
function inside(area, target) {
  const a = parseRange(area);
  const t = parseRange(target);
  return !!a && !!t && a.r1 >= t.r1 && a.r2 <= t.r2 && a.c1 >= t.c1 && a.c2 <= t.c2;
}

function gate(primary, op) {
  if (!primary || primary.format !== 'xlsx') return { ok: false, reason: 'xlsx-observation-missing' };
  if (!primary.complete) return { ok: false, reason: primary.truncated ? 'xlsx-observation-truncated' : 'xlsx-observation-unreadable-getters', unsupported: primary.unsupported };
  const s = primary.sheets.filter(x => x.name === op.selector.sheet);
  if (s.length !== 1) return { ok: false, reason: s.length > 1 ? 'xlsx-sheet-ambiguous' : 'xlsx-sheet-not-found', available: primary.sheets.map(x => x.name) };
  const found = {};
  for (const { a } of op.cells) {
    const cell = cellAt(primary, op.selector.sheet, a);
    if (!cell || typeof cell.merge !== 'string') return { ok: false, reason: 'xlsx-target-cell-unread', cell: a };
    if (!inside(cell.merge, op.selector.range)) return { ok: false, reason: 'xlsx-merge-area-reaches-outside-the-target', cell: a, mergeArea: cell.merge };
    found[a] = cell;
  }
  return { ok: true, cells: found };
}

/** X-L06: merge_range with across:true -- SATISFIED when every cell already reports ITS OWN row's
 * area; the content guard is PER ROW (a row's non-left-most non-blank cell refuses that row's merge,
 * and since Merge(true) is a single call over the whole range, any one row's risk refuses the whole op). */
function classifyAcross(cellsByAddr, op, areas) {
  const grid = parseRange(op.selector.range);
  if (op.cells.every(({ a }, i) => areas[i] === rowRangeOf(grid, parseCell(a).r))) {
    return { state: 'SATISFIED', evidence: { sheet: op.selector.sheet, range: op.selector.range, proof: 'fresh-live-every-cell-in-its-own-row-area' } };
  }
  const content = op.cells.filter(({ a }) => a !== leftMostOf(grid, parseCell(a).r) && !isBlank(cellsByAddr[a])).map(({ a }) => ({ cell: a, formula: cellsByAddr[a].f !== '' }));
  if (content.length) return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-merge-data-loss-risk', cells: content.slice(0, 20), count: content.length } };
  return { state: 'APPLY', evidence: { sheet: op.selector.sheet, range: op.selector.range, innerMerges: [...new Set(areas.filter((x, i) => x !== op.cells[i].a))] } };
}

function classify(primary, op) {
  const g = gate(primary, op);
  if (!g.ok) return { state: 'FAIL_CLOSED', evidence: g };
  const range = op.selector.range;
  const areas = op.cells.map(({ a }) => g.cells[a].merge);
  if (op.intent === 'merge_range') {
    if (op.across) return classifyAcross(g.cells, op, areas);
    if (areas.every(x => x === range)) return { state: 'SATISFIED', evidence: { sheet: op.selector.sheet, range, proof: 'fresh-live-every-cell-in-the-target-area' } };
    // the content guard: the engine empties every non-top-left cell silently (W4b)
    const content = op.cells.filter(({ a }) => a !== op.topLeft && !isBlank(g.cells[a])).map(({ a }) => ({ cell: a, formula: g.cells[a].f !== '' }));
    if (content.length) return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-merge-data-loss-risk', cells: content.slice(0, 20), count: content.length } };
    return { state: 'APPLY', evidence: { sheet: op.selector.sheet, range, innerMerges: [...new Set(areas.filter((x, i) => x !== op.cells[i].a))] } };
  }
  if (op.cells.every(({ a }, i) => areas[i] === a)) return { state: 'SATISFIED', evidence: { sheet: op.selector.sheet, range, proof: 'fresh-live-every-cell-unmerged' } };
  return { state: 'APPLY', evidence: { sheet: op.selector.sheet, range, merges: [...new Set(areas.filter((x, i) => x !== op.cells[i].a))] } };
}

function resolve(primary, selector, ctx = {}) {
  const g = gate(primary, ctx.operation);
  if (!g.ok) return { ok: false, target: null, selector, ambiguity: g.reason, detail: g };
  return { ok: true, target: Object.freeze({ sheet: selector.sheet, range: selector.range, before: g.cells }) };
}

/** One callCommand: an exact-name sheet check, then the public Merge(false) / UnMerge() on the range. */
function lower(op) {
  const { sheet, range } = op.selector;
  // X-L06: the SAME public Merge call, just the measured `true` argument -- the engine itself
  // splits the whole-range call into one merge per row (no separate MergeAcross method exists).
  const call = op.intent === 'merge_range' ? `r.Merge(${op.across ? 'true' : 'false'});` : 'r.UnMerge();';
  return {
    kind: 'callCommand',
    body: `
try {
  const name = ${JSON.stringify(sheet)};
  const ws = Api.GetSheet(name);
  if (!ws || ws.GetName() !== name) return { ok: false, stage: 'sheet-identity' };
  const r = ws.GetRange(${JSON.stringify(range)});
  ${call}
  return { ok: true };
} catch (err) { return { ok: false, stage: 'threw', error: String(err) }; }
`,
    meta: Object.freeze({ family: 'merge', action: op.intent, sheet, range, across: op.across }),
  };
}

/**
 * L1: every target cell reports the target area (merge) or itself (unmerge), AND every target cell keeps its
 * value / formula (the top-left its content; the others were blank by the guard and stay blank).
 */
function verifyPrimitive(op, before, after) {
  const b = before && before.primary;
  const a = after && after.primary;
  const grid = op.across ? parseRange(op.selector.range) : null;
  const expectedMerge = addrStr => {
    if (op.intent !== 'merge_range') return addrStr;
    return op.across ? rowRangeOf(grid, parseCell(addrStr).r) : op.selector.range;
  };
  const wrong = op.cells.filter(({ a: addr }) => {
    const now = cellAt(a, op.selector.sheet, addr);
    return !now || now.merge !== expectedMerge(addr);
  }).map(({ a: addr }) => addr);
  const contentChanged = op.cells.filter(({ a: addr }) => {
    const was = cellAt(b, op.selector.sheet, addr);
    const now = cellAt(a, op.selector.sheet, addr);
    return !was || !now || was.f !== now.f || (was.f === '' && JSON.stringify(was.v) !== JSON.stringify(now.v));
  }).map(({ a: addr }) => addr);
  return { ok: wrong.length === 0 && contentChanged.length === 0, evidence: { sheet: op.selector.sheet, range: op.selector.range, cells: op.cells.length, wrong: wrong.slice(0, 20), contentChanged: contentChanged.slice(0, 20), proof: 'fresh-live-per-cell-merge-area' } };
}

/**
 * L3: as write_range (every other cell keeps value / formula), AND every other read cell keeps its merge area.
 *
 * X-L13/3 (owner EURO-MCP capability-backlog): `merge_range`/`unmerge_range` migrated to the TIERED
 * observation above 45000 used cells (adapter.cjs TIERED_INTENTS). The digest (T2, xlsx-driver.cjs digestQuery) carries
 * no merge-area field at all -- only type/value/formula -- so a cell's merge area OUTSIDE the task's windows cannot be
 * re-read from a tiered primary (`before.primary.cells` / `after.primary.cells` hold ONLY the window cells under tier
 * 'tiered'; inventory.cjs's own comment: "a tiered read's cells are only the windows"). Letting the loop below run
 * unchanged would not fail -- `Object.entries(b.cells)` would simply never see an outside cell, and an empty result
 * looks identical to "compared, nothing changed".
 *
 * Naming this ONLY in a nested evidence string (an earlier version of this fix) is not enough -- the surrounding
 * receipt (statuses APPLIED, l3 verdict PASS) would still read as an unqualified FULL proof, exactly the shape the
 * X-N01b owner rule forbids for a reduced proof. The fix is the SAME machine-readable taxonomy format.cjs's own
 * `scopeOf()` already uses for the numberFormat-only bounded case (`preservationScope` / `proven` / `notProven` --
 * write-tools.cjs's `l3Scope()` reads these generically, by field name, not by family): a tiered merge/unmerge is
 * BOUNDED, `notProven: ['merge-topology-outside-target']`. The `ok` verdict is NOT weakened by this -- it still rests on
 * the shared digest comparison above (`cells.verifyPreservation`'s own `withDigest()`, proving value/formula
 * preservation outside the windows) and on the architectural invariant `gate()` already enforces before ANY write --
 * a target cell's merge area can never reach outside the target range, so no outside cell can share one with it
 * either. Below 45000 (FULL tier) this function is UNCHANGED: no `preservationScope` field at all, exactly as before
 * (write_range/clear_range's own FULL-tier receipts carry none either).
 */
function verifyPreservation(op, before, after) {
  const rest = cells.verifyPreservation(op, before, after);
  if (!rest.ok && rest.evidence && rest.evidence.reason) return rest;
  const b = before.primary;
  const a = after.primary;
  if (b.tier === 'tiered') {
    return { ok: rest.ok, evidence: { ...rest.evidence, mergeChanged: 'NOT_RECHECKED:tiered-digest-carries-no-merge-area-field', mergeChangedCount: null,
      preservationScope: 'BOUNDED',
      proven: ['target-merge-area', 'target-content', 'sheet-values-and-formulas-outside-windows-digest'],
      notProven: ['merge-topology-outside-target'] } };
  }
  const target = new Set(op.cells.map(c => key(op.selector.sheet, c.a)));
  const mergeChanged = Object.entries(b.cells).filter(([k, c]) => !target.has(k) && a.cells[k] && a.cells[k].merge !== c.merge).map(([k, c]) => ({ cell: k.replace('\u0000', '!'), before: c.merge, after: a.cells[k].merge }));
  return { ok: rest.ok && mergeChanged.length === 0, evidence: { ...rest.evidence, mergeChanged: mergeChanged.slice(0, 20), mergeChangedCount: mergeChanged.length } };
}

function wholeTaskChecks(op, finalObservation) {
  const v = verifyPrimitive(op, finalObservation, finalObservation);
  const label = op.intent !== 'merge_range' ? 'holds no merged cell' : (op.across ? 'is merged per row' : 'is one merged area');
  return [{ name: `${op.selector.sheet}!${op.selector.range} ${label}`, ok: v.evidence.wrong.length === 0, detail: v.evidence.wrong.length ? { wrong: v.evidence.wrong } : null }];
}

/** The target window (a merge may cover cells outside every used range). */
const windowsOf = op => [{ sheet: op.selector.sheet, range: op.selector.range }];

module.exports = { INTENTS, plan, classify, resolve, lower, verifyPrimitive, verifyPreservation, wholeTaskChecks, windowsOf, inside };
