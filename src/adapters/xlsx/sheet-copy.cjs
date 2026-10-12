'use strict';

const { isBlank } = require('./inventory.cjs');
const formula = require('./formula.cjs');

/**
 * copy_sheet's comparison (M6 wave 3, EO_INTERNAL_DEPENDENCY WRITE): is one sheet the faithful copy of
 * another -- cells, objects WITH their references, validation / conditional-formatting rules? The
 * reference rule (owner: "the correct references, not only count and presence"): a reference to the
 * SOURCE sheet itself must name the copy; any other reference is unchanged. What cannot be observed
 * is a difference, never an empty set. Pure.
 */

/** A reference text (formula, series, pivot source, name) of the copy against its source's: a reference to the source sheet itself must now name the copy; any other reference is unchanged. */
function refKept(a, b, src, dst) {
  const norm = x => (x === null || x === undefined || x === '' ? '' : `=${String(x).replace(/^=/, '').replace(/^\s+/, '')}`);
  const x = norm(a);
  const y = norm(b);
  if (x === '' || y === '') return x === y;
  const own = formula.sheetRefs(x).names.some(n => n.toLowerCase() === src.toLowerCase());
  return own ? formula.canonicalRenamed(x, src, dst) === formula.canonical(y) : formula.canonical(x) === formula.canonical(y);
}

const OBJECT_AXES = ['charts', 'pivots', 'names'];
const objectsComplete = o => !!o && OBJECT_AXES.every(k => Array.isArray(o[k]));

function chartDiffs(c, t, i, src, dst) {
  if (!t) return [];
  if (c.type !== t.type || c.series.length !== t.series.length) return [{ object: 'chart', i, change: 'shape' }];
  const out = [];
  c.series.forEach((x, j) => { for (const f of ['name', 'values', 'cat', 'x']) if (!refKept(x[f], t.series[j][f], src, dst)) out.push({ object: 'chart', i, series: j, field: f, source: x[f], copy: t.series[j][f] }); });
  return out;
}

function objectDiffs(so, to, src, dst) {
  const out = [];
  for (const k of OBJECT_AXES) if (so[k].length !== to[k].length) out.push({ object: k, change: 'count', source: so[k].length, copy: to[k].length });
  so.charts.forEach((c, i) => out.push(...chartDiffs(c, to.charts[i], i, src, dst)));
  so.pivots.forEach((pv, i) => { const t = to.pivots[i]; if (t && !refKept(pv.source, t.source, src, dst)) out.push({ object: 'pivot', i, source: pv.source, copy: t.source }); });
  so.names.forEach((n, i) => { const t = to.names[i]; if (t && (t.name !== n.name || !refKept(n.refersTo, t.refersTo, src, dst))) out.push({ object: 'name', name: n.name, source: n.refersTo, copy: t.refersTo }); });
  return out;
}

/** Cell by cell: value, formula (self-reference mapped), number format, fill, merge, bold. */
function cellDiffs(p, src, dst) {
  const diffs = [];
  const cellsOn = n => new Map(Object.entries(p.cells).filter(([k]) => k.startsWith(`${n}\u0000`)).map(([k, c]) => [k.split('\u0000')[1], c]));
  const sc = cellsOn(src);
  const tc = cellsOn(dst);
  for (const [a, c] of sc) {
    const x = tc.get(a);
    if (!x) { if (!isBlank(c)) diffs.push({ a, change: 'missing' }); continue; }
    if (!refKept(c.f, x.f, src, dst)) diffs.push({ a, field: 'formula', source: c.f, copy: x.f });
    else if (c.f === '' && JSON.stringify(c.v) !== JSON.stringify(x.v)) diffs.push({ a, field: 'value' });
    for (const k of ['nf', 'fill', 'merge']) if (JSON.stringify(c[k]) !== JSON.stringify(x[k])) diffs.push({ a, field: k });
    if ('bold' in c && (typeof c.bold !== 'boolean' || c.bold !== x.bold)) diffs.push({ a, field: typeof c.bold === 'boolean' ? 'bold' : 'bold-unobservable' });
  }
  for (const [a, x] of tc) if (!sc.has(a) && !isBlank(x)) diffs.push({ a, change: 'extra' });
  return { diffs, compared: sc.size };
}

/** Validation / conditional-formatting rules: an unobservable side is a difference, never an empty set. */
function ruleDiffs(p, src, dst) {
  const rs = p.rules && p.rules.bySheet ? p.rules.bySheet[src] : undefined;
  const rt = p.rules && p.rules.bySheet ? p.rules.bySheet[dst] : undefined;
  if (!Array.isArray(rs) || !Array.isArray(rt)) return [{ change: 'rules-unobservable' }];
  const same = rs.length === rt.length && rs.every((r, i) => r.kind === rt[i].kind && r.type === rt[i].type && r.range === rt[i].range
    && refKept(r.formula, rt[i].formula, src, dst) && refKept(r.formula2, rt[i].formula2, src, dst));
  // the diff names the rules on both sides (bounded): without them a LIVE mismatch says only THAT, not WHAT differed
  const brief = list => list.slice(0, 10).map(r => ({ kind: r.kind, type: r.type, range: r.range, formula: r.formula, formula2: r.formula2 }));
  return same ? [] : [{ change: 'rules', source: brief(rs), copy: brief(rt), sourceCount: rs.length, copyCount: rt.length }];
}

/** Is sheet `dst` the copy of `src` -- cells, objects (with their references) and rules? */
function copyMatches(p, src, dst) {
  const cells = cellDiffs(p, src, dst);
  const so = p.objects[src];
  const to = p.objects[dst];
  const objs = objectsComplete(so) && objectsComplete(to) ? objectDiffs(so, to, src, dst) : [{ change: 'objects-unobservable' }];
  const diffs = [...cells.diffs, ...objs, ...ruleDiffs(p, src, dst)];
  return { ok: diffs.length === 0, diffs: diffs.slice(0, 20), count: diffs.length, compared: cells.compared };
}

/** Sheet-level names on the source whose reference names a sheet that does not exist (a copy would multiply them). */
const brokenNames = (p, n) => (p.objects[n].names || []).filter(x => formula.sheetRefs(`=${String(x.refersTo || '').replace(/^=/, '')}`).names.some(ref => !p.sheets.some(sh => sh.name.toLowerCase() === ref.toLowerCase())));


module.exports = { copyMatches, objectsComplete, brokenNames, refKept, objectDiffs };
