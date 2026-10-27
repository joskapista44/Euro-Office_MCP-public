'use strict';

const formula = require('./formula.cjs');
const { addr } = require('./a1.cjs');

/**
 * The workbook's dependents and rules, from the three dependency readers (M6 wave 3; owner decision
 * 2026-09-28, owner channel -- the READ wiring). Pure.
 *
 * Readers (the driver fills them; docs/evidence/M6-W3C-DEPS-AND-COPY-PROPOSAL.md, W3d/W3e):
 *   names        INTERNAL, build-gated: asc_getDefinedNames(All) -> [{ name, ref, scopeSheet }]
 *                (scopeSheet: null = workbook level, else the sheet the scope index maps to)
 *   validations  INTERNAL, build-gated: each worksheet model's dataValidations.elems ->
 *                { <sheet>: [{ type, ranges: [{ r1, c1, r2, c2 }], formula1, formula2 }] }
 *   conditional  PUBLIC: whole-sheet GetFormatConditions -> GetItem(i) ->
 *                { <sheet>: [{ type, appliesTo, formula1, formula2 }] }
 *   objects      PUBLIC: per sheet { charts, pivots, names } (the W3c readers)
 * A reader that is absent, threw, or returned an unexpected shape is a reason string ('UNOBSERVABLE:..'),
 * never an empty list (owner rule). So is a reference that cannot be resolved: an unknown or ambiguous
 * name, a sheet that does not exist, an external / structured reference, INDIRECT, a name
 * whose text starts with '=' (W3e: the engine never resolves it), an unmeasured rule or series shape.
 * Anything UNOBSERVABLE makes the whole dependency set incomplete: delete and copy refuse before the write.
 * A '#REF!' points at nothing: it is no dependency (dependencies.broken keeps where it sits, refOnTarget refuses it
 * on the sheets a write changes).
 */

/**
 * Conditional-format types whose reference-carrying getters were measured LIVE: xlExpression (W3e); xlCellValue
 * (W6 probe C1/C2, the whole-sheet public read: GetType / GetOperator / GetFormula1 / GetFormula2 / GetAppliesTo, in
 * session and after a reopen; M7 S1, owner GO 2026-09-30) -- only with the operators that read measured.
 * GH #141 (docs/GH141-CF-DEPENDENCY-OBSERVABILITY-20261018.md): xlAboveAverageCondition / xlTop10 /
 * xlUniqueValues added -- not a new measurement, a structural fact: cf.cjs's own doc comment records that
 * these types' Formula1/Formula2 getters THROW (the engine has no formula concept for them), so there is
 * nothing unmeasured to guess at; only appliesTo (already LIVE-measured for every CF type via cf.cjs's
 * deleteBody) is read for them. xlColorScale / xlDatabar / xlIconSets stay OUT: their CFVO stops/criteria can
 * be formula-typed in Excel, and no reader in this codebase has ever called the getter that would show it.
 */
const MEASURED_CF_TYPES = Object.freeze(['xlExpression', 'xlCellValue', 'xlAboveAverageCondition', 'xlTop10', 'xlUniqueValues']);
// X-L13/7: the remaining six xlCellValue operators round-trip through BOTH readers this module
// consumes (the whole-sheet `conditional` reader AND `cfReads`) after a reopen -- operator /
// formula1 / formula2 byte-identical to what was written, on a dedicated sheet each. All eight of
// Excel's xlFormatConditionOperator values are now measured; none remain to add.
const MEASURED_CELL_VALUE_OPERATORS = Object.freeze({
  xlGreater: 1, xlBetween: 2,
  xlLess: 1, xlLessEqual: 1, xlGreaterEqual: 1, xlEqual: 1, xlNotEqual: 1, xlNotBetween: 2,
});
const ERROR_LITERALS = /#(N\/A|DIV\/0!|VALUE!|NAME\?|NUM!|NULL!|GETTING_DATA)/gi;
const unobs = x => typeof x === 'string' && x.startsWith('UNOBSERVABLE');

/** Is an identifier NOT a defined-name use: a function, a cell / column reference, a boolean? */
const notAName = (id, next, prev) => next === '(' || /^[A-Za-z]{1,3}\d+$/.test(id) || /^(TRUE|FALSE)$/i.test(id) || (/^[A-Za-z]{1,3}$/.test(id) && (next === ':' || prev === ':'));

/** The text outside string literals and quoted sheet names, or why it cannot be read as references. */
function referenceText(text) {
  const toks = formula.tokenize(String(text).replace(/^=/, ''));
  if (!toks) return { problem: 'unparsable' };
  // A broken reference (#REF!, also with a trailing cell / range such as '#REF!A1' or '#REF!$A$1:$B$2', and with a sheet
  // prefix such as 'Data!#REF!') points at NOTHING, so it is no dependency: the whole token is dropped and the rest of the
  // formula is read as before. Only the TOKEN goes -- a trailing 'A1' is NOT left behind to be misread as a same-sheet reference.
  // (formula.sheetRefs still names the 'Data' of 'Data!#REF!': a conservative sheet dependency, never a missing one.)
  // L-18/1 measured the old rule (any #REF! made the whole workbook's dependencies incomplete) blocking every delete on a book
  // that held one broken formula anywhere. Where a #REF! sits is kept in dependencies.broken: refOnTarget below.
  const other = toks.map(t => (t.kind === 'other' ? t.text : ' ')).join('')
    .replace(/(?:[A-Za-z_\u00C0-\uFFFF][\w.\u00C0-\uFFFF]*!)?#REF!(?:\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?|\$?[A-Z]{1,3}:\$?[A-Z]{1,3}|\$?\d+:\$?\d+)?/gi, ' ');
  if (/#REF/i.test(other)) return { problem: 'broken-reference' };
  if (/\[/.test(other)) return { problem: 'external-or-structured-reference' };
  if (/\bINDIRECT\s*\(/i.test(other)) return { problem: 'indirect' };
  return { other };
}

/** Identifiers in a formula text that are neither functions, cell / column references, nor booleans: defined-name uses. */
function nameUses(text) {
  const r = referenceText(text);
  if (r.problem) return r;
  const s = r.other.replace(ERROR_LITERALS, ' ').replace(/[A-Za-z_\u00C0-\uFFFF][\w.\u00C0-\uFFFF]*!/g, '!').replace(/\$/g, '');
  const out = [];
  for (const m of s.matchAll(/(^|[^\w.])([A-Za-z_\\\u00C0-\uFFFF][\w.\u00C0-\uFFFF]*)/g)) {
    const id = m[2];
    const at = m.index + m[1].length;
    if (!notAName(id, s.slice(at + id.length).replace(/^\s+/, '')[0], s[at - 1])) out.push(id);
  }
  return { names: [...new Set(out)] };
}

/** The one definition a name use in `host` resolves to (sheet scope first, then workbook), or why not. */
function definitionOf(n, host, ctx) {
  const defs = ctx.names.filter(d => d.name.toLowerCase() === n.toLowerCase());
  const scoped = defs.filter(d => d.scopeSheet !== null && d.scopeSheet.toLowerCase() === String(host || '').toLowerCase());
  const chosen = scoped.length ? scoped : defs.filter(d => d.scopeSheet === null);
  if (chosen.length === 0) return { problem: `unknown-name:${n}` };
  if (chosen.length > 1) return { problem: `ambiguous-name:${n}` };
  return { def: chosen[0] };
}

/** The sheets a name use reaches, following the name's own text (cycles and depth bounded). */
function throughName(n, host, ctx, depth, seen) {
  const found = definitionOf(n, host, ctx);
  if (found.problem) return found;
  const d = found.def;
  const k = `${d.name}\u0000${d.scopeSheet}`;
  if (seen.has(k)) return { problem: `name-cycle:${n}` };
  if (defProblem(d)) return { problem: `${n}->${defProblem(d)}` };
  const inner = resolveRef(d.ref, d.scopeSheet || host, ctx, depth + 1, new Set([...seen, k]));
  return inner.problem ? { problem: `${n}->${inner.problem}` } : inner;
}

/**
 * Resolve one reference text in the context of `host` (the sheet it lives on): the sheets it names,
 * directly or through defined names, or the problem that makes it unresolvable.
 */
function resolveRef(text, host, ctx, depth = 0, seen = new Set()) {
  const t = String(text === null || text === undefined ? '' : text).trim().replace(/^=/, '');
  if (t === '') return { sheets: [] };
  if (depth > 8) return { problem: 'name-chain-too-deep' };
  const refs = formula.sheetRefs(`=${t}`);
  const uses = nameUses(t);
  // GH #298 C: a structured reference stays unresolvable here (its table's sheet is not read) -- fail-closed as before
  const early = !refs.ok ? 'unparsable' : refs.external || refs.structured ? 'external-or-structured-reference' : uses.problem;
  if (early) return { problem: early };
  const sheets = new Set();
  for (const n of refs.names) {
    const hit = ctx.sheets.find(x => x.toLowerCase() === n.toLowerCase());
    if (!hit) return { problem: `missing-sheet:${n}` };
    sheets.add(hit);
  }
  for (const n of uses.names) {
    const inner = throughName(n, host, ctx, depth, seen);
    if (inner.problem) return inner;
    inner.sheets.forEach(x => sheets.add(x));
  }
  return { sheets: [...sheets] };
}

/** A defined name whose own text cannot be trusted as a reference (W3e: '=' prefixed -> never resolved). */
const defProblem = d => (/^\s*=/.test(d.ref) ? 'name-text-starts-with-equals' : null);

const rangeA1 = g => (g.r1 === g.r2 && g.c1 === g.c2 ? addr(g.c1 + 1, g.r1 + 1) : `${addr(g.c1 + 1, g.r1 + 1)}:${addr(g.c2 + 1, g.r2 + 1)}`);
const stripBook = x => String(x || '').replace(/^\[[^\]]*\]/, '');

/** Shape checks of the raw reader output; null when well-formed, else why not. */
function namesShape(v) {
  if (unobs(v)) return v;
  if (!Array.isArray(v)) return 'UNOBSERVABLE:names-not-a-list';
  const bad = v.find(n => !n || typeof n.name !== 'string' || n.name === '' || typeof n.ref !== 'string' || !(n.scopeSheet === null || typeof n.scopeSheet === 'string'));
  return bad ? 'UNOBSERVABLE:names-unexpected-shape' : null;
}
const rect = g => !!g && [g.r1, g.c1, g.r2, g.c2].every(Number.isInteger);
const textOrNull = f => f === null || typeof f === 'string';
function validationsShape(v, sheets) {
  if (unobs(v)) return v;
  if (!v || typeof v !== 'object' || Array.isArray(v)) return 'UNOBSERVABLE:validations-not-a-map';
  if (sheets.some(n => !Array.isArray(v[n]))) return 'UNOBSERVABLE:validations-sheet-missing';
  const bad = sheets.some(n => v[n].some(r => !r || !Array.isArray(r.ranges) || !r.ranges.length || !r.ranges.every(rect) || !textOrNull(r.formula1) || !textOrNull(r.formula2)));
  return bad ? 'UNOBSERVABLE:validations-unexpected-shape' : null;
}
/** An xlCellValue rule read in full: a measured operator, formula1 present, formula2 present iff the operator takes two. */
function cellValueShape(r) {
  const n = MEASURED_CELL_VALUE_OPERATORS[r.operator];
  if (!n) return 'UNOBSERVABLE:conditional-operator-not-measured';
  const has = f => typeof f === 'string' && f !== '';
  if (!has(r.formula1) || (n === 2 ? !has(r.formula2) : !(r.formula2 === '' || r.formula2 === null))) return 'UNOBSERVABLE:conditional-partial-read';
  return null;
}
function conditionalShape(list) {
  if (unobs(list)) return list;
  if (!Array.isArray(list)) return 'UNOBSERVABLE:conditional-not-a-list';
  if (list.some(r => !r || !MEASURED_CF_TYPES.includes(r.type))) return 'UNOBSERVABLE:conditional-type-not-measured';
  if (list.some(r => typeof r.appliesTo !== 'string' || r.appliesTo === '' || !textOrNull(r.formula1) || !textOrNull(r.formula2))) return 'UNOBSERVABLE:conditional-unexpected-shape';
  return list.map(r => (r.type === 'xlCellValue' ? cellValueShape(r) : null)).find(Boolean) || null;
}

/** Rules per sheet in one comparable shape: [{ kind, range, formula, formula2, type }] ordered by range. */
function rulesOf(validations, conditional, sheet) {
  const v = validations[sheet].map(r => ({ kind: 'validation', type: r.type, range: r.ranges.map(rangeA1).join(','), formula: r.formula1 || '', formula2: r.formula2 || '' }));
  const c = conditional[sheet].map(r => ({ kind: 'conditional', type: r.type, ...(r.type === 'xlCellValue' ? { operator: r.operator } : {}), range: stripBook(r.appliesTo).replace(/^.*!/, '').replace(/\$/g, ''), formula: r.formula1 || '', formula2: r.formula2 || '' }));
  const order = (a, b) => (a.range < b.range ? -1 : a.range > b.range ? 1 : 0);
  return [...v.sort(order), ...c.sort(order)];
}

/** Which readers answered in a usable shape: a reason string per gap, null when fine. */
function readerGaps(sheets, raw) {
  const cf = raw.conditional && typeof raw.conditional === 'object' && !unobs(raw.conditional) ? raw.conditional : null;
  const cfAbsent = unobs(raw.conditional) ? raw.conditional : 'UNOBSERVABLE:conditional-reader-absent';
  return {
    names: namesShape(raw.names === undefined ? 'UNOBSERVABLE:names-reader-absent' : raw.names),
    validations: validationsShape(raw.validations === undefined ? 'UNOBSERVABLE:validations-reader-absent' : raw.validations, sheets),
    cf,
    conditional: Object.fromEntries(sheets.map(n => [n, cf ? conditionalShape(cf[n] === undefined ? 'UNOBSERVABLE:conditional-sheet-missing' : cf[n]) : cfAbsent])),
  };
}

/** Collects resolved references and gaps. */
function collector(ctx) {
  const unobservable = [];
  const refs = [];
  const broken = [];
  const note = (axis, why, host = null) => unobservable.push({ axis, why, host });
  const add = (kind, host, id, text) => {
    // a #REF! is no dependency (referenceText drops it), but WHERE it sits is kept: refOnTarget decides per write
    if (/#REF!/i.test(String(text || ''))) broken.push({ kind, host, name: id, text: String(text) });
    const r = resolveRef(text, host, ctx);
    if (r.problem) note(`${kind}:${host || '*'}:${id}`, `UNOBSERVABLE:${r.problem}`, host);
    else if (r.sheets.length) refs.push({ kind, host, name: id, sheets: r.sheets });
  };
  return { unobservable, refs, broken, note, add };
}

/** A chart series field: empty, a reference ('=...', measured), or an unmeasured shape (a gap). */
function scanCharts(col, n, charts) {
  charts.forEach((c, i) => (c.series || []).forEach((x, j) => {
    for (const f of ['name', 'values', 'cat', 'x']) {
      const v = x[f];
      if (v === '' || v === null || v === undefined) continue;
      if (typeof v === 'string' && v.startsWith('=')) col.add('chart', n, `${i}.${j}.${f}`, v);
      else col.note(`chart:${n}:${i}.${j}.${f}`, 'UNOBSERVABLE:series-shape-not-measured', n);
    }
  }));
}

/** The public objects of every sheet: a sheet not read, or an axis that failed, is a gap. */
function scanObjects(col, sheets, objects) {
  for (const n of sheets) {
    const o = objects && objects[n];
    if (!o) { col.note(`objects:${n}`, 'UNOBSERVABLE:objects-not-read', n); continue; }
    ['charts', 'pivots', 'names'].filter(axis => !Array.isArray(o[axis])).forEach(axis => col.note(`${axis}:${n}`, unobs(o[axis]) ? o[axis] : 'UNOBSERVABLE:not-a-list', n));
    if (Array.isArray(o.charts)) scanCharts(col, n, o.charts);
    if (Array.isArray(o.pivots)) o.pivots.forEach(p => col.add('pivot', n, p.name, p.source));
  }
}

/** Every rule's two formulas as dependents of its sheet. */
function scanRules(col, kind, bySheet, sheets) {
  for (const n of sheets) bySheet[n].forEach((r, i) => { col.add(kind, n, `${i}.formula1`, r.formula1); col.add(kind, n, `${i}.formula2`, r.formula2); });
}

/**
 * The analysis: { names, dependencies: { complete, unobservable[], refs[] }, rules: { complete, unobservable[], bySheet } }.
 * `sheets` names every sheet; `cells` maps '<sheet>\0<A1>' -> { f }; `raw` = { names, validations,
 * conditional, objects } (absent readers -> UNOBSERVABLE).
 */
function analyze(sheets, cells, raw) {
  const gaps = readerGaps(sheets, raw);
  const col = collector({ sheets, names: gaps.names ? [] : raw.names });
  if (gaps.names) col.note('defined-names', gaps.names);
  if (gaps.validations) col.note('data-validation', gaps.validations);
  for (const [n, why] of Object.entries(gaps.conditional)) if (why) col.note(`conditional-formatting:${n}`, why, n);
  // a formula's direct sheet references stay with the formula guard too; here a formula through a
  // name reaches the name's sheets, and a broken / unresolvable formula text is a gap
  for (const [k, c] of Object.entries(cells)) if (c.f) { const [host, a] = k.split('\u0000'); col.add('formula', host, a, c.f); }
  if (!gaps.names) for (const d of raw.names) { if (defProblem(d)) col.note(`name:${d.scopeSheet || '*'}:${d.name}`, `UNOBSERVABLE:${defProblem(d)}`, d.scopeSheet); else col.add('name', d.scopeSheet, d.name, d.ref); }
  if (!gaps.validations) scanRules(col, 'validation', raw.validations, sheets);
  scanRules(col, 'conditional', gaps.cf || {}, sheets.filter(n => !gaps.conditional[n]));
  scanObjects(col, sheets, raw.objects);
  const rulesComplete = !gaps.validations && sheets.every(n => !gaps.conditional[n]);
  return {
    names: gaps.names ? null : raw.names.map(d => ({ name: d.name, ref: d.ref, scopeSheet: d.scopeSheet })),
    dependencies: { complete: col.unobservable.length === 0, unobservable: col.unobservable, refs: col.refs, broken: col.broken },
    rules: { complete: rulesComplete, unobservable: col.unobservable.filter(u => u.axis === 'data-validation' || u.axis.startsWith('conditional-formatting:')),
      bySheet: rulesComplete ? Object.fromEntries(sheets.map(n => [n, rulesOf(raw.validations, gaps.cf, n)])) : null },
  };
}

/**
 * The ONE pre-write guard for a broken reference: a #REF! (cell formula, validation, conditional format, defined name,
 * chart series, pivot source) hosted on a sheet the write changes -- copy_sheet: the copied sheet. What the engine does
 * with such a text under the write (shift, delete, rename, copy) is not measured, so it is refused, named apart from the
 * reader-gap refusals; a #REF! on any OTHER sheet is no dependency and does not block. Null when none; a workbook-level
 * name (host null) sits on no sheet.
 */
function refOnTarget(deps, targets) {
  const want = new Set((targets || []).filter(t => typeof t === 'string').map(t => t.toLowerCase()));
  const list = deps && Array.isArray(deps.broken) ? deps.broken : null;
  if (list === null) return { reason: 'xlsx-ref-on-target-unmeasured', refs: 'not-read' };
  const hit = list.filter(b => b.host !== null && want.has(String(b.host).toLowerCase()));
  return hit.length ? { reason: 'xlsx-ref-on-target-unmeasured', refs: hit.slice(0, 20), count: hit.length } : null;
}

module.exports = { analyze, resolveRef, nameUses, refOnTarget, MEASURED_CF_TYPES };
