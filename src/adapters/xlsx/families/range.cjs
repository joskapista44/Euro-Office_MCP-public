'use strict';

const { parseRange, parseCell, cellsOf, rangeText, addr, MAX_COL, MAX_ROW } = require('../a1.cjs');
const { cellAt, key, isBlank } = require('../inventory.cjs');
const formula = require('../formula.cjs');
const { refOnTarget } = require('../dependencies.cjs');
const cells = require('./cells.cjs');
const { translateFormula, referencesBlock, mapRefs, moveFormula } = require('../shift.cjs');
const { EXPECT_VERSION, checkExpect, matches, cellOf } = require('../move-expect.cjs');

/**
 * XLSX `range` family: copy_range / move_range (M6 wave 5). docs/evidence/M6-W5-STRUCTURE-AUDIT-AND-CONTRACT.md
 * sections 3 and 5 (the W5 probe measured Copy's translation and format transfer and Cut's re-pointing).
 *
 * copy_range { sheet, range, target, targetSheet? }: the public ApiRange.Copy onto the target's top-left; the target
 *   block has the source's shape. Retry = the natural terminal state: every target cell equals the translated
 *   source -> SATISFIED. A target that is neither blank nor that is REFUSED (the source overwrote silently, B-S10).
 * move_range { sheet, range, target, targetSheet?, expect }: the public Cut. `expect` (owner decision 2026-09-30, Talk
 *   owner channel) is the WHOLE block's expected content, carried by the task and built by moveExpectation() from a fresh
 *   live read BEFORE the first write: { version: 1, fileId, source: { sheet, range }, target: { sheet, range },
 *   cells: rows x cols of { t: 'blank' } | { t: 'number', v } | { t: 'text', v } | { t: 'formula', f } } -- bound to
 *   the document, both sheets and both ranges. APPLY: the source equals it cell by cell (re-checked on the fresh read
 *   right before the write), the target blank. SATISFIED 0/0: the source blank AND the whole target block equal to it
 *   cell by cell (type, blank, formula TEXT; a formula's computed value is not content). Anything else is refused --
 *   never a second Cut (the source's X1 retry Cut the blank source over the moved data). This is terminal-state
 *   evidence, not proof of where the content came from.
 * Refused before the write, stated: a same-sheet overlap; a merged cell in source or target; a copied formula
 * whose translation leaves the sheet (#REF!); for a move, a formula OUTSIDE the block that references it and a
 * formula INSIDE it that references the block (Cut re-points them -- measured in W5, still refused here).
 * L1 the target cells (copy: the translated source; move: the source as it was) and, for a move, the source blank;
 * L3 every other cell of every sheet. Formats are not compared here (Copy carries the format -- W5 probe CP).
 *
 * M7 S1 (owner GO 2026-09-30, owner channel): what Copy / Cut do to the workbook's DEPENDENTS is not measured, so
 * both are refused before the write when a defined name, a validation, a conditional format, a chart series, a pivot
 * or an AutoFilter reaches the source or the target block -- directly, from another sheet, or through a defined name
 * -- and when any of those axes cannot be observed or a reference cannot be resolved (an unresolvable reference is
 * never "not affected"). L3 also keeps every defined name, every rule, every sheet's charts / pivots / sheet names
 * and the AutoFilter of both sheets unchanged; an axis unreadable after the write fails L3 (no APPLIED).
 *
 * M7 E R7 (owner M7 GO 2026-10-03; tier E probe m7-e-probe1-20261003.json MV / MX, ONE Cut, session = reopen): a move
 * whose block is referenced no longer refuses when every reference into the block is a measured shape (shift.cjs
 * moveFormula): a cell or range ENTIRELY inside the block, relative or absolute, bare or sheet-qualified, from inside
 * or outside the block, on any sheet -- it moves with the block; a cross-sheet move gives a bare reference the target
 * sheet's prefix. A defined name entirely inside the source block moves too (measured). Refused before the write (NOT
 * measured / silent change): a reference that only partly overlaps the block (it does not move, its value changes), a
 * whole-row / whole-column reference into it, a sheet-qualified reference into a block moved to another sheet, an
 * inner reference of a block moved to another sheet, a name partly over it. Rules, charts (C15 measured: a series
 * does NOT follow a Cut), pivots and AutoFilters reaching a block still refuse (S1). L1 / the retry: the target holds
 * the carried expectation WITH its references moved; L3: every outside formula holds exactly its moved text and its
 * value, every name its moved reference, everything else unchanged. Never a second Cut.
 */

const INTENTS = Object.freeze(['copy_range', 'move_range']);
const MAX_CELLS = 500;
function plan(step, index) {
  const at = `step ${index}: ${step.intent}`;
  const fields = ['intent', 'id', 'sheet', 'range', 'target', 'targetSheet'];
  const unknown = Object.keys(step).filter(k => !fields.includes(k));
  if (unknown.length) return { ok: false, detail: `${at} does not support field(s): ${unknown.join(', ')}` };
  if (typeof step.sheet !== 'string' || step.sheet === '') return { ok: false, detail: `${at} requires a sheet name` };
  if ('targetSheet' in step && (typeof step.targetSheet !== 'string' || step.targetSheet === '')) return { ok: false, detail: `${at}: targetSheet must be a sheet name` };
  const g = parseRange(step.range);
  if (!g) return { ok: false, detail: `${at} requires a plain A1 range` };
  if (g.count > MAX_CELLS) return { ok: false, detail: `${at} range has ${g.count} cells, over the ${MAX_CELLS} ceiling` };
  const t = parseCell(step.target);
  if (!t) return { ok: false, detail: `${at}: target must be ONE cell (the top-left of the target block)` };
  const h = g.r2 - g.r1;
  const w = g.c2 - g.c1;
  if (t.c + w > MAX_COL || t.r + h > MAX_ROW) return { ok: false, detail: `${at}: the target block leaves the sheet` };
  const tg = { c1: t.c, r1: t.r, c2: t.c + w, r2: t.r + h, count: g.count };
  const tsheet = step.targetSheet || step.sheet;
  if (tsheet === step.sheet && tg.c1 <= g.c2 && g.c1 <= tg.c2 && tg.r1 <= g.r2 && g.r1 <= tg.r2) return { ok: false, detail: `${at}: the source and the target overlap on the same sheet` };
  // L-06/b (owner decision, zero 2026-10-15): the caller no longer supplies `expect` -- classify()
  // builds it from its OWN fresh read, directly before dispatch (moveExpectationFromOp below), the
  // SAME session the write itself uses. `expect` stays null here, mutated by classify() -- the SAME
  // technique copy_chart's classifyCopy already uses for op.grid (charts.cjs).
  const expect = null;
  const src = cellsOf(g);
  const pairs = src.map(a => { const c = parseCell(a); return Object.freeze({ from: a, to: addr(c.c - g.c1 + tg.c1, c.r - g.r1 + tg.r1) }); });
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`,
      intent: step.intent,
      family: 'range',
      selector: Object.freeze({ sheet: step.sheet, range: rangeText(g) }),
      dest: Object.freeze({ sheet: tsheet, range: rangeText(tg) }),
      delta: Object.freeze({ dr: tg.r1 - g.r1, dc: tg.c1 - g.c1 }),
      pairs: Object.freeze(pairs),
      expect,
      // the planner's overlap check uses selector; L3 builds its own exclusion from pairs (verifyPreservation below)
      cells: Object.freeze([...src.map(a => Object.freeze({ a })), ...(tsheet === step.sheet ? cellsOf(tg).map(a => Object.freeze({ a })) : [])]),
      preservationMode: 'PRIMITIVE_INTEGRAL',
    },
  };
}

const blockOf = (sheet, range) => { const g = parseRange(range); return { sheet, r1: g.r1, r2: g.r2, c1: g.c1, c2: g.c2 }; };
const sameText = (want, got) => { const a = formula.canonical(want); const b = formula.canonical(got); return a !== null && a === b; };

/** What a target cell must hold after the op, from the source cell as read before. */
function expected(op, s) {
  if (s.f === '') return { v: s.v, f: '' };
  if (op.intent === 'move_range') return { f: moveFormula(s.f, op.selector.sheet, moveOf(op)).formula };
  const t = translateFormula(s.f, op.delta.dr, op.delta.dc);
  return t.broken ? { broken: true } : { f: t.formula };
}

/** The carried expectation's cell for a pair (row-major, the source's shape). */
function expectOf(op, i) {
  const cols = op.expect.cells[0].length;
  return op.expect.cells[Math.floor(i / cols)][i % cols];
}

/** M7 E R7: the Cut as moveFormula sees it. */
const moveOf = op => { const g = parseRange(op.selector.range); return { from: { sheet: op.selector.sheet, r1: g.r1, c1: g.c1, r2: g.r2, c2: g.c2 }, to: { sheet: op.dest.sheet }, dr: op.delta.dr, dc: op.delta.dc }; };

/** The carried expectation's cell at the target, its references moved with the block (R7). */
function movedExpect(op, i) {
  const e = expectOf(op, i);
  return e.t === 'formula' ? { t: 'formula', f: moveFormula(e.f, op.selector.sheet, moveOf(op)).formula } : e;
}

/** M7 E R7: why one formula's references cannot be moved as measured, or null. */
function moveProblem(m, innerAcross) {
  if (m.formula === null) return 'unparsable';
  if (m.partial) return 'partial-or-whole-line-reference';
  if (m.qualifiedAcross) return 'qualified-reference-into-a-block-moved-to-another-sheet';
  return innerAcross && m.moved ? 'inner-reference-of-a-block-moved-to-another-sheet' : null;
}

/**
 * M7 E R7: every formula of the observation after a move -- `cells` maps an OUTSIDE cell key to its moved text (only
 * those whose text changes); `problems` lists the reference shapes the probe did not measure (refused before the write).
 */
function moveExpectations(p, op) {
  const mv = moveOf(op);
  const cross = !same(op.selector.sheet, op.dest.sheet);
  const inside = new Set(op.pairs.map(x => key(op.selector.sheet, x.from)));
  const out = { cells: new Map(), problems: [] };
  for (const [k, c] of Object.entries(p.cells)) {
    if (!c.f) continue;
    const m = moveFormula(c.f, k.split('\u0000')[0], mv);
    const why = moveProblem(m, inside.has(k) && cross);
    if (why) out.problems.push({ cell: k.replace('\u0000', '!'), why });
    else if (!inside.has(k) && m.moved) out.cells.set(k, m.formula);
  }
  return out;
}

/** M7 E R7: a defined name whose reference moves with the block (entirely inside it, qualified), or null. */
function movedName(op, n) {
  if (op.intent !== 'move_range') return null;
  const m = moveFormula(`=${n.ref}`, n.scopeSheet || '', moveOf(op));
  return m.formula !== null && m.moved > 0 && !m.partial && !m.qualifiedAcross && !refParts(n.ref).some(x => x.sheet === null) ? m.formula.slice(1) : null;
}

// L-06/b: landedName (the retry's "already landed" check) removed -- classify() refuses a blank
// source before dependentGuard runs, so that branch is unreachable now (see dependentGuard above).

const holds = (now, e) => !!now && !e.broken && (e.f === '' ? now.f === '' && JSON.stringify(now.v) === JSON.stringify(e.v) : now.f !== '' && sameText(e.f, now.f));

function gate(p, op) {
  if (!p || p.format !== 'xlsx') return { ok: false, reason: 'xlsx-observation-missing' };
  if (!p.complete) return { ok: false, reason: p.truncated ? 'xlsx-observation-truncated' : 'xlsx-observation-unreadable-getters', unsupported: p.unsupported };
  for (const n of new Set([op.selector.sheet, op.dest.sheet])) {
    const s = p.sheets.filter(x => x.name === n);
    if (s.length !== 1) return { ok: false, reason: s.length > 1 ? 'xlsx-sheet-ambiguous' : 'xlsx-sheet-not-found', sheet: n };
  }
  const src = {};
  const tgt = {};
  for (const { from, to } of op.pairs) {
    const s = cellAt(p, op.selector.sheet, from);
    const t = cellAt(p, op.dest.sheet, to);
    if (!s || !t) return { ok: false, reason: 'xlsx-range-cell-unread', cell: s ? to : from };
    if ((s.merge && s.merge !== from) || (t.merge && t.merge !== to)) return { ok: false, reason: 'xlsx-range-merged-cell-refused', cell: s.merge && s.merge !== from ? from : to };
    src[from] = s; tgt[to] = t;
  }
  return { ok: true, src, tgt };
}

function classifyCopy(p, op, g, where, tgtBlank) {
  const exp = op.pairs.map(({ from }) => expected(op, g.src[from]));
  const broken = op.pairs.filter((x, i) => exp[i].broken).map(x => x.from);
  if (broken.length) return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-copy-formula-leaves-the-sheet', cells: broken.slice(0, 20) } };
  if (op.pairs.every(({ to }, i) => holds(g.tgt[to], exp[i]))) return { state: 'SATISFIED', evidence: { ...where, proof: 'fresh-live-target-equals-the-translated-source' } };
  if (!tgtBlank) return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-copy-target-occupied', cells: op.pairs.filter(({ to }) => !isBlank(g.tgt[to])).map(x => x.to).slice(0, 20) } };
  return { state: 'APPLY', evidence: where };
}

const refuse = (reason, extra) => ({ state: 'FAIL_CLOSED', evidence: { reason, ...extra } });

/**
 * L-06/b (owner decision, zero 2026-10-15): the carried expectation is now built FRESH by classify()
 * itself (moveExpectationFromOp below), from the SAME observation this function already has -- so it
 * is bound to THIS document and THIS read BY CONSTRUCTION, and the old explicit fileId-mismatch check
 * (owner decision 2026-09-30, when the caller supplied `expect` and it could be stale) can never fire
 * here and is gone. A blank source (a retry, or any other already-applied-looking state) is caught
 * BEFORE this function is even called (classify()'s own early refusal, xlsx-move-already-applied-
 * not-provable) -- a server-built expectation would be built from the ALREADY-MOVED state and has
 * nothing to prove the move against, so v1 never attempts it. This function therefore only ever
 * sees a NOT-blank source.
 */
function classifyMove(p, op, g, where, tgtBlank) {
  const srcOff = op.pairs.filter(({ from }, i) => !matches(g.src[from], expectOf(op, i))).map(x => x.from);
  if (srcOff.length) return refuse('xlsx-move-source-differs-from-expectation', { ...where, cells: srcOff.slice(0, 20) });
  if (!tgtBlank) return refuse('xlsx-move-target-occupied', { cells: op.pairs.filter(({ to }) => !isBlank(g.tgt[to])).map(x => x.to).slice(0, 20) });
  const mx = moveExpectations(p, op);
  if (mx.problems.length) return refuse('xlsx-move-reference-shape-not-measured', { cells: mx.problems.slice(0, 20), count: mx.problems.length });
  return { state: 'APPLY', evidence: { ...where, repointed: [...mx.cells.keys()].map(k => k.replace('\u0000', '!')).slice(0, 20) } };
}

const unobs = x => typeof x === 'string' && x.startsWith('UNOBSERVABLE');

/** Every reference part of a formula text: { sheet (null = bare), abs (every row / column part '$') }; null = unparsable. */
function refParts(text) {
  const parts = [];
  const f = String(text).startsWith('=') ? String(text) : `=${text}`;
  const out = mapRefs(f, ref => {
    const flags = ref.kind === 'cell' ? [ref.a[1], ref.a[3]] : ref.kind === 'range' ? [ref.a[1], ref.a[3], ref.b[1], ref.b[3]] : [ref.parts[1], ref.parts[3]];
    parts.push({ sheet: ref.sheet, abs: flags.every(x => x === '$') });
    return '';
  });
  return out === null ? null : parts;
}

/** The bounding offsets of a rule's applies-to areas from its first area's top-left (the formula's anchor). */
function spreadOf(areas) {
  const a = areas[0];
  const r1 = Math.min(...areas.map(x => x.r1)); const r2 = Math.max(...areas.map(x => x.r2));
  const c1 = Math.min(...areas.map(x => x.c1)); const c2 = Math.max(...areas.map(x => x.c2));
  return { dr: [r1 - a.r1, r2 - a.r1], dc: [c1 - a.c1, c2 - a.c1] };
}

const intersects = (g, b) => g.r1 <= b.r2 && g.r2 >= b.r1 && g.c1 <= b.c2 && g.c2 >= b.c1;

const same = (a, b) => String(a).toLowerCase() === String(b).toLowerCase();

/** The guard's scan state: the two blocks, what reached them (hits) and what could not be judged (problems). */
function scanState(op) {
  const blocks = [blockOf(op.selector.sheet, op.selector.range), blockOf(op.dest.sheet, op.dest.range)];
  const st = { blocks, hits: [], problems: [] };
  // one reference text of a dependent living on `host` (null = workbook level); a name's own text must be absolute
  st.check = (kind, host, id, text, { spread = null, absolute = false } = {}) => {
    if (text === null || text === undefined || String(text).trim() === '') return;
    const f = String(text).startsWith('=') ? String(text) : `=${text}`;
    const parts = refParts(f);
    const why = parts === null ? 'unparsable'
      : host === null && parts.some(x => x.sheet === null) ? 'bare-reference-without-a-host-sheet'
        : absolute && parts.some(x => !x.abs) ? 'relative-name-reference' : null;
    if (why) st.problems.push({ kind, host, id, why });
    else if (blocks.some(b => referencesBlock(f, host === null ? '' : host, b, spread))) st.hits.push({ kind, host, id });
  };
  return st;
}

/** Validations and conditional formats: the applies-to area, and both formulas repeated over it. */
function scanRules(st, bySheet) {
  for (const [sheet, rules] of Object.entries(bySheet)) {
    rules.forEach((r, i) => {
      const areas = String(r.range).split(',').map(x => parseRange(x.trim()));
      if (areas.some(x => !x)) { st.problems.push({ kind: r.kind, host: sheet, id: i, why: 'applies-to-unparsable' }); return; }
      if (st.blocks.some(b => same(b.sheet, sheet) && areas.some(a => intersects(a, b)))) st.hits.push({ kind: r.kind, host: sheet, id: i, on: 'applies-to', range: r.range });
      const spread = spreadOf(areas);
      st.check(r.kind, sheet, `${i}.formula`, r.formula, { spread });
      st.check(r.kind, sheet, `${i}.formula2`, r.formula2, { spread });
    });
  }
}

/** Charts (every series field) and pivots (the source; a pivot HOSTED on a block's sheet: its extent is not observable). */
function scanObjects(st, objects) {
  for (const [sheet, o] of Object.entries(objects || {})) {
    (o.charts || []).forEach((c, i) => (c.series || []).forEach((x, j) => ['name', 'values', 'cat', 'x'].forEach(f => st.check('chart', sheet, `${i}.${j}.${f}`, x[f]))));
    for (const pv of o.pivots || []) {
      if (st.blocks.some(b => same(b.sheet, sheet))) st.hits.push({ kind: 'pivot', host: sheet, id: pv.name, on: 'hosted-extent-unobservable' });
      st.check('pivot', sheet, pv.name, pv.source);
    }
  }
}

/** The AutoFilter of both blocks' sheets (the W6 reader): unreadable -> a problem, its range over a block -> a hit. */
function scanFilters(st, filterReads) {
  for (const sheet of new Set(st.blocks.map(b => b.sheet))) {
    const fr = filterReads ? filterReads[sheet] : undefined;
    const g = fr && !unobs(fr) && fr.range ? parseRange(String(fr.range).replace(/^.*!/, '').replace(/\$/g, '')) : null;
    if (!fr || unobs(fr)) st.problems.push({ kind: 'autofilter', host: sheet, why: fr || 'not-read' });
    else if (fr.range && !g) st.problems.push({ kind: 'autofilter', host: sheet, why: 'range-unparsable' });
    else if (g && st.blocks.some(b => same(b.sheet, sheet) && intersects(g, b))) st.hits.push({ kind: 'autofilter', host: sheet, range: fr.range });
  }
}

/** Why the dependency observation cannot be used at all, or null. */
function readerGap(p) {
  const d = p.dependencies;
  if (!d || !p.rules) return { reason: 'xlsx-range-dependents-not-read' };
  if (!d.complete) return { reason: 'xlsx-range-dependency-observation-incomplete', unobservable: d.unobservable.slice(0, 20), count: d.unobservable.length };
  if (!p.rules.complete || !p.rules.bySheet) return { reason: 'xlsx-range-rule-observation-incomplete', unobservable: p.rules.unobservable.slice(0, 20) };
  return Array.isArray(p.definedNames) ? null : { reason: 'xlsx-range-dependency-observation-incomplete', unobservable: ['defined-names'] };
}

/**
 * M7 S1 pre-write guard: null when no dependent reaches either block, else why not (affected or unobservable).
 * Reads: the dependency analysis (names, validations, conditional formats, charts, pivots -- the W3 readers), the
 * AutoFilter of both sheets (the W6 reader).
 */
function dependentGuard(p, op) {
  const gap = readerGap(p);
  if (gap) return gap;
  // both blocks' sheets (source and destination) are changed by the write
  const ref = refOnTarget(p.dependencies, [op.selector.sheet, op.dest.sheet]);
  if (ref) return ref;
  const st = scanState(op);
  // L-06/b: the retry/"already landed" branch (landedName) is unreachable now -- classify() refuses
  // a blank source BEFORE this guard ever runs (see classify() below), so a move_range reaching here
  // is always a first write, still moving TO the target.
  const follows = n => movedName(op, n) !== null;
  for (const n of p.definedNames) if (!follows(n)) st.check('name', n.scopeSheet, n.name, n.ref, { absolute: true });
  scanRules(st, p.rules.bySheet);
  scanObjects(st, p.objects);
  scanFilters(st, p.filterReads);
  if (st.problems.length) return { reason: 'xlsx-range-dependent-unobservable', dependents: st.problems.slice(0, 20), count: st.problems.length };
  return st.hits.length ? { reason: 'xlsx-range-dependent-affected', dependents: st.hits.slice(0, 20), count: st.hits.length } : null;
}

/**
 * L-06/b (owner decision, zero 2026-10-15): no `expect` on the move_range step -- the caller never
 * supplies one. classify() builds it itself, from THIS SAME observation, directly before dispatch
 * (the SAME session the write itself uses, no second open/read/close) -- mutating `op` the SAME way
 * classifyCopy already mutates op.grid for copy_chart (charts.cjs). req is reconstructed from the
 * already-resolved op (its own source/dest are the single source of truth), and moveExpectation
 * re-plans/re-gates from it -- pure, idempotent, the SAME p.
 */
function classify(p, op) {
  const g = gate(p, op);
  if (!g.ok) return { state: 'FAIL_CLOSED', evidence: g };
  const where = { from: `${op.selector.sheet}!${op.selector.range}`, to: `${op.dest.sheet}!${op.dest.range}` };
  if (op.intent === 'move_range') {
    const srcBlank = op.pairs.every(({ from }) => isBlank(g.src[from]));
    // v1: a blank source (a retry, or any other already-applied-looking state) has nothing to build
    // a fresh expectation FROM -- it would be built from the ALREADY-MOVED state. Refused named,
    // 0/0, never a second Cut, rather than guessed SATISFIED. Checked BEFORE dependentGuard, so a
    // retry costs nothing beyond the gate already run.
    if (srcBlank) return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-move-already-applied-not-provable', ...where } };
    const dstG = parseRange(op.dest.range);
    const req = { sheet: op.selector.sheet, range: op.selector.range, target: addr(dstG.c1, dstG.r1), targetSheet: op.dest.sheet };
    const mx = moveExpectation(p, req);
    if (!mx.ok) return { state: 'FAIL_CLOSED', evidence: { reason: mx.reason, ...where, detail: mx } };
    op.expect = Object.freeze(mx.expect);
  }
  const dg = dependentGuard(p, op);
  if (dg) return { state: 'FAIL_CLOSED', evidence: dg };
  const tgtBlank = op.pairs.every(({ to }) => isBlank(g.tgt[to]));
  return (op.intent === 'copy_range' ? classifyCopy : classifyMove)(p, op, g, where, tgtBlank);
}

function resolve(p, selector, ctx = {}) {
  const g = gate(p, ctx.operation);
  if (!g.ok) return { ok: false, target: null, selector, ambiguity: g.reason, detail: g };
  const op = ctx.operation;
  if (op.intent !== 'copy_range') return { ok: true, target: Object.freeze({ from: selector, to: op.dest }) };
  const formulas = copyFormulas(op, g.src);
  if (op.dest.sheet === op.selector.sheet) return { ok: true, target: Object.freeze({ from: selector, to: op.dest, formulas }) };
  // ANOTHER sheet (zero l065-l068): the editor's Copy leaves a delayed state that overwrites a formula written after it at the next save -- so no Copy is dispatched; the
  // values and the planner-translated formulas are written (what the probe's P7 measured faithful); the format is NOT carried (named, never silent)
  const w = copyValues(op, g.src);
  if (!w.ok) return { ok: false, target: null, selector, ambiguity: w.reason, detail: w };
  return { ok: true, target: Object.freeze({ from: selector, to: op.dest, formulas, values: w.values, crossSheet: true }) };
}

/** The non-formula, non-blank source cells of a CROSS-SHEET copy as typed values ({ 'B1': 'h1', 'B2': 1 }); a cell whose type the reader cannot tell, or a text that reads as a formula, is refused, named. */
function copyValues(op, src) {
  const values = {};
  for (const { from, to } of op.pairs) {
    const s = src[from];
    if (!s || s.f !== '' || isBlank(s)) continue;
    const e = cellOf(s);
    if (e === null) return { ok: false, reason: 'xlsx-copy-cross-sheet-value-type-not-observable', cell: from };
    if (e.t === 'number') { const n = Number(e.v); if (!Number.isFinite(n)) return { ok: false, reason: 'xlsx-copy-cross-sheet-value-type-not-observable', cell: from }; values[to] = n; continue; }
    if (e.v.startsWith('=')) return { ok: false, reason: 'xlsx-copy-cross-sheet-text-reads-as-formula', cell: from };
    values[to] = e.v;
  }
  return { ok: true, values: Object.freeze(values) };
}

/** The formula text each copied FORMULA cell must carry at its target ({ 'E3': '=E2+L1' }): the planner's own translation of the source text as read now. */
function copyFormulas(op, src) {
  const out = {};
  for (const { from, to } of op.pairs) {
    const s = src[from];
    if (!s || s.f === '') continue;
    const t = translateFormula(s.f, op.delta.dr, op.delta.dc);
    if (!t.broken) out[to] = `=${t.formula.slice(1).trim()}`;
  }
  return Object.freeze(out);
}

/**
 * L-06 r5/r6 (zero l064 / l065 / l066 probes, LIVE on the prod DS): ApiRange.Copy shows a formula in session but the SAVED file carries the source's LITERAL text
 * ("= A2+H1") and its stale value -- the in-session L1 cannot see it, the next fresh read (the retry) can. An explicit formula write after the Copy persists
 * correctly (same sheet: P2). Onto ANOTHER sheet the editor's own in-session text was not trustworthy either (l066: the explicit rewrite of the editor text left
 * B4 literal) -- so the text written is the PLANNER's translation of the source formula as read before the write (resolve -> target.formulas, the same
 * translateFormula the classify / L1 equality uses), never the editor's. Each written cell is read back in the body; a target that did not take a formula fails
 * the command, named.
 */
const REWRITE_BODY = (formulas, values) => `${values ? `  const values = ${JSON.stringify(values)};\n  for (const to of Object.keys(values)) t.GetRange(to).SetValue(values[to]);\n` : ''}  const formulas = ${JSON.stringify(formulas)};
  for (const to of Object.keys(formulas)) {
    const cell = t.GetRange(to);
    cell.SetValue(formulas[to]);
    if (String(cell.GetFormula() || '').trim().charAt(0) !== '=') return { ok: false, stage: 'xlsx-copy-formula-not-written', cell: to };
  }
`;

/** One callCommand: both sheets checked by exact name, then the public Copy / Cut onto the target's top-left. */
function lower(op, target) {
  const call = op.intent === 'copy_range' ? 'Copy' : 'Cut';
  const copy = op.intent === 'copy_range';
  const cross = copy && op.dest.sheet !== op.selector.sheet && !!target && target.crossSheet === true;
  return {
    kind: 'callCommand',
    body: `
try {
  const s = Api.GetSheet(${JSON.stringify(op.selector.sheet)});
  const t = Api.GetSheet(${JSON.stringify(op.dest.sheet)});
  if (!s || s.GetName() !== ${JSON.stringify(op.selector.sheet)} || !t || t.GetName() !== ${JSON.stringify(op.dest.sheet)}) return { ok: false, stage: 'sheet-identity' };
${cross ? '' : `  s.GetRange(${JSON.stringify(op.selector.range)}).${call}(t.GetRange(${JSON.stringify(op.pairs[0].to)}));\n`}${copy ? REWRITE_BODY((target && target.formulas) || {}, cross ? (target.values || {}) : null) : ''}  return { ok: true };
} catch (err) { return { ok: false, stage: 'threw', error: String(err) }; }
`,
    meta: Object.freeze({ family: 'range', action: op.intent, sheet: op.selector.sheet, range: op.selector.range, targetSheet: op.dest.sheet, pairs: op.pairs, delta: op.delta, ...(copy ? { explicitFormulaWrite: true, formulaWrites: (target && target.formulas) || {}, ...(cross ? { crossSheetWrite: true, valueWrites: target.values || {} } : {}) } : {}) }),
  };
}

/** L1: every target cell holds what the op puts there; a move also leaves the source blank. */
function verifyPrimitive(op, before, after) {
  const b = before && before.primary;
  const a = after && after.primary;
  const wrong = [];
  for (const { from, to } of op.pairs) {
    const was = cellAt(b, op.selector.sheet, from);
    if (!was) { wrong.push({ from, to, why: 'source-unread-before' }); continue; }
    const now = cellAt(a, op.dest.sheet, to);
    const i = op.pairs.findIndex(x => x.from === from);
    const ok = isBlank(was) ? isBlank(now) : holds(now, expected(op, was));
    if (!ok || (op.intent === 'move_range' && !matches(now, movedExpect(op, i)))) wrong.push({ from, to, was: { v: was.v, f: was.f }, now: now ? { v: now.v, f: now.f } : 'unread' });
  }
  const notCleared = op.intent === 'move_range' ? op.pairs.filter(({ from }) => !(op.dest.sheet === op.selector.sheet && op.pairs.some(x => x.to === from)) && !isBlank(cellAt(a, op.selector.sheet, from))).map(x => x.from) : [];
  return { ok: wrong.length === 0 && notCleared.length === 0, evidence: { from: `${op.selector.sheet}!${op.selector.range}`, to: `${op.dest.sheet}!${op.dest.range}`, wrong: wrong.slice(0, 20), sourceNotCleared: notCleared.slice(0, 20), proof: 'fresh-live-per-cell' } };
}

/**
 * L3: every other cell of every sheet. Excluded (L1's): the target block, and for a MOVE the source block (it is
 * emptied). A COPY's source is protected here, on the same and on another target sheet (owner review 8527132).
 */
function verifyPreservation(op, before, after) {
  // M7 E R7: an outside formula a move re-points is compared here, to its moved text and its unchanged value
  const moved = op.intent === 'move_range' && before && before.primary ? moveExpectations(before.primary, op).cells : new Map();
  const repointed = [...moved].filter(([k, f]) => { const was = before.primary.cells[k]; const now = after && after.primary && after.primary.cells[k]; return !now || !sameText(f, now.f) || JSON.stringify(now.v) !== JSON.stringify(was.v); }).map(([k, f]) => ({ cell: k.replace('\u0000', '!'), want: f }));
  const skip = new Set([...op.pairs.map(x => key(op.dest.sheet, x.to)), ...(op.intent === 'move_range' ? op.pairs.map(x => key(op.selector.sheet, x.from)) : []), ...moved.keys()]);
  const strip = x => (x && x.primary && x.primary.cells ? { ...x, primary: { ...x.primary, cells: Object.fromEntries(Object.entries(x.primary.cells).filter(([k]) => !skip.has(k))) } } : x);
  const rest = cells.verifyPreservation({ ...op, cells: [] }, strip(before), strip(after));
  if (!rest.ok && rest.evidence && rest.evidence.reason) return rest;
  const diffs = dependentDiffs(op, before && before.primary, after && after.primary);
  return { ok: rest.ok && diffs.length === 0 && repointed.length === 0, evidence: { ...rest.evidence, dependentsKept: diffs.length === 0, dependentDiffs: diffs.slice(0, 20), repointedWrong: repointed.slice(0, 20) } };
}

/**
 * M7 S1 L3 over the dependents: every defined name, every rule, every sheet's objects (charts / pivots / sheet names)
 * and both sheets' AutoFilter equal before and after; an axis that is not an observable value on EITHER side is a
 * difference (never compared as "unchanged").
 */
function dependentDiffs(op, b, a) {
  if (!b || !a) return [{ axis: 'observation-missing' }];
  const names = (x, moveIt) => (Array.isArray(x.definedNames) ? JSON.stringify(x.definedNames.map(n => `${n.scopeSheet || ''}\u0000${n.name.toLowerCase()}\u0000${formula.canonical(`=${(moveIt && movedName(op, n)) || n.ref}`)}`).sort()) : null);
  const rules = x => (x.rules && x.rules.complete && x.rules.bySheet ? JSON.stringify(x.rules.bySheet) : null);
  // one axis: a difference unless BOTH sides are observable and equal
  const kept = (vb, va) => vb !== null && va !== null && vb === va;
  const objectsOf = o => (o && ['charts', 'pivots', 'names'].every(ax => Array.isArray(o[ax])) ? JSON.stringify(o) : null);
  const filterOf = (x, n) => { const f = x.filterReads ? x.filterReads[n] : undefined; return f && !unobs(f) ? JSON.stringify(f) : null; };
  return [
    ...(kept(names(b, true), names(a, false)) ? [] : [{ axis: 'defined-names' }]),
    ...(kept(rules(b), rules(a)) ? [] : [{ axis: 'rules' }]),
    ...[...new Set([...Object.keys(b.objects || {}), ...Object.keys(a.objects || {})])].filter(n => !kept(objectsOf((b.objects || {})[n]), objectsOf((a.objects || {})[n]))).map(n => ({ axis: 'objects', sheet: n })),
    ...[...new Set([op.selector.sheet, op.dest.sheet])].filter(n => !kept(filterOf(b, n), filterOf(a, n))).map(n => ({ axis: 'autofilter', sheet: n })),
  ];
}

function wholeTaskChecks(op, finalObservation) {
  const p = finalObservation && finalObservation.primary;
  const name = `${op.intent} ${op.selector.sheet}!${op.selector.range} -> ${op.dest.sheet}!${op.dest.range}`;
  if (op.intent === 'copy_range') {
    const ok = !!p && op.pairs.every(({ from, to }) => { const s = cellAt(p, op.selector.sheet, from); return s && (isBlank(s) ? isBlank(cellAt(p, op.dest.sheet, to)) : holds(cellAt(p, op.dest.sheet, to), expected(op, s))); });
    // another sheet: no Copy, so the cells' FORMAT (bold, number format, fill ...) is not carried -- named, never silent
    const cross = op.dest.sheet !== op.selector.sheet;
    return [{ name, ok, detail: null, operationId: op.id, ...(cross ? { notProven: ['cross-sheet-copy-format-not-carried'] } : {}) }];
  }
  // L-06/b: wholeTaskChecks receives a FRESHLY re-planned op (adapter.verifyWholeTask's own
  // planOperations(task) call, src/adapters/xlsx/adapter.cjs -- a DIFFERENT instance than the one
  // classify() mutated during the main dispatch), so op.expect is null here, and `finalObservation`
  // is POST-write (the source is ALREADY blank by now) -- re-running classify() against it would hit
  // the SAME "blank source -> refuse" branch this file's own classify() uses for a real retry, never
  // repopulating expect. There is therefore no PRE-write snapshot left to compare the target against
  // at this layer: that comparison already happened at L1 (verifyPrimitive), on the SAME op instance
  // classify() mutated with a PRE-write expect, in the SAME dispatch. This check is deliberately
  // WEAKER than the old owner-review-8527132 one (per-cell content equality) -- it only confirms the
  // source is left fully blank, and NAMES the narrowed scope via notProven rather than silently
  // dropping it.
  const kept = !p ? [] : op.pairs.filter(({ from }) => !isBlank(cellAt(p, op.selector.sheet, from))).map(x => x.from);
  const ok = !!p && kept.length === 0;
  return [{ name: `${name}: source left blank (target content equality is L1's job, not re-provable here)`, ok, detail: ok ? null : (p ? { sourceNotBlank: kept.slice(0, 20) } : { reason: 'no-final-observation' }), notProven: ['move-target-content-equality-not-reprovable-at-l2'] }];
}

/** Source and target blocks are windows (the target may lie outside every used range). */
const windowsOf = op => [{ sheet: op.selector.sheet, range: op.selector.range }, { sheet: op.dest.sheet, range: op.dest.range }];

/** M7 S1: the dependency readers (whole workbook) and the AutoFilter of both sheets. */
const liveExtras = op => ({ dependents: true, filterSheets: [...new Set([op.selector.sheet, op.dest.sheet])] });

/** The preparatory read of a move: both blocks as windows (a blank cell outside every used range is otherwise unread). */
function moveReadScope(req) {
  const e = plan({ ...req, intent: 'copy_range' }, 0);
  if (!e.ok) return { windows: [] };
  return { windows: windowsOf(e.operation) };
}

/**
 * The carried expectation, from a fresh LIVE read taken BEFORE the first write (owner decision 2026-09-30): the source
 * block as it is now, bound to the document (runtime.fileId), both sheets and both ranges. Refuses what the read cannot
 * prove: an unread or merged cell, an unobservable type (COUNT unread, a boolean / error literal), an all-blank block.
 */
/** The preparatory read's gates: a valid request, both blocks readable, the document named. */
function prepareExpectation(p, req) {
  const e = plan({ ...req, intent: 'copy_range' }, 0);
  if (!e.ok) return { fail: { ok: false, reason: 'xlsx-move-request-invalid', detail: e.detail } };
  const g = gate(p, e.operation);
  if (!g.ok) return { fail: { ok: false, reason: g.reason, detail: g } };
  const fileId = p.runtime && typeof p.runtime.fileId === 'string' && p.runtime.fileId !== '' ? p.runtime.fileId : null;
  return fileId ? { op: e.operation, g, fileId } : { fail: { ok: false, reason: 'xlsx-move-document-identity-unobservable' } };
}

function moveExpectation(p, req) {
  const pre = prepareExpectation(p, req);
  if (pre.fail) return pre.fail;
  const { op, g, fileId } = pre;
  const flat = op.pairs.map(({ from }) => ({ from, cell: cellOf(g.src[from]) }));
  const unobservable = flat.find(x => x.cell === null);
  if (unobservable) return { ok: false, reason: 'xlsx-move-expect-type-unobservable', cell: unobservable.from, count: g.src[unobservable.from].count };
  const pr = parseRange(op.selector.range);
  const cols = pr.c2 - pr.c1 + 1;
  const cells = Array.from({ length: pr.r2 - pr.r1 + 1 }, (_, r) => flat.slice(r * cols, (r + 1) * cols).map(x => x.cell));
  const expect = { version: EXPECT_VERSION, fileId, source: { sheet: op.selector.sheet, range: op.selector.range }, target: { sheet: op.dest.sheet, range: op.dest.range }, cells };
  const chk = checkExpect(expect, expect.source, expect.target, cells.length, cols);
  if (chk.error) return { ok: false, reason: /nothing to prove/.test(chk.error) ? 'xlsx-move-expect-nothing-to-prove' : 'xlsx-move-expect-invalid', detail: chk.error };
  return { ok: true, expect };
}

module.exports = { INTENTS, EXPECT_VERSION, moveReadScope, moveExpectation, plan, classify, resolve, lower, verifyPrimitive, verifyPreservation, wholeTaskChecks, windowsOf, liveExtras };
