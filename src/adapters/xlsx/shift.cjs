'use strict';

const { tokenize } = require('./formula.cjs');
const { colNum, colName, MAX_COL, MAX_ROW } = require('./a1.cjs');

/**
 * Reference-aware formula rewriting for the M6 wave-5 structural and range families. Pure.
 *
 * The source's shifters were regexes over the whole formula text (audit X4): they rewrote function names with
 * digits (LOG10), text inside string literals and references to OTHER sheets. Here a reference is rewritten only
 * where it IS a reference: outside string literals, not a function name (followed by '('), not part of a longer
 * identifier, and on the sheet the operation acts on (a qualified 'Sheet!' prefix, or the host sheet when bare).
 *
 * Semantics (Excel's; the W5 probe measured the engine the same way, 2026-09-29): an insert of k rows at row N moves every row >= N
 * by +k (a range whose end is >= N grows); a delete of rows N..N+k-1 turns a single reference into the block into
 * #REF!, shrinks a range that overlaps the block and moves every row after it by -k. Columns the same. Whole-row
 * ('3:4') and whole-column ('B:C') references follow the same rules on their own axis.
 */

const IDENT = /[A-Za-z0-9_.À-￿]/;
const CELL = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)/;
const ROWS = /^(\$?)(\d+):(\$?)(\d+)/;
const COLS = /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![A-Za-z0-9_(])/;

/** Scan one 'other' token for references; `ctx.sheetBefore` is the sheet a qname token right before named. */
function scan(text, onRef, sheetBefore) {
  let out = '';
  let i = 0;
  let pendingSheet = sheetBefore;
  while (i < text.length) {
    const prev = i === 0 ? '' : text[i - 1];
    // a bare sheet prefix: Name!
    const sm = /^([A-Za-z_À-￿][\w.À-￿]*)!/.exec(text.slice(i));
    if (sm && !(prev && IDENT.test(prev))) { out += sm[0]; pendingSheet = sm[1]; i += sm[0].length; continue; }
    const rest = text.slice(i);
    const boundary = prev === '!' || !(prev && IDENT.test(prev));
    let m;
    if (boundary && (m = ROWS.exec(rest)) && !IDENT.test(rest[m[0].length] || '')) {
      out += onRef({ kind: 'rows', sheet: pendingSheet, parts: m }); i += m[0].length; pendingSheet = null; continue;
    }
    if (boundary && (m = COLS.exec(rest))) {
      out += onRef({ kind: 'cols', sheet: pendingSheet, parts: m }); i += m[0].length; pendingSheet = null; continue;
    }
    if (boundary && (m = CELL.exec(rest))) {
      const after = rest[m[0].length] || '';
      if (!/[A-Za-z0-9_(]/.test(after) && Number(m[4]) >= 1) {
        const r2 = after === ':' ? CELL.exec(rest.slice(m[0].length + 1)) : null;
        const tail = r2 ? rest[m[0].length + 1 + r2[0].length] || '' : '';
        if (r2 && !/[A-Za-z0-9_(]/.test(tail)) {
          out += onRef({ kind: 'range', sheet: pendingSheet, a: m, b: r2 }); i += m[0].length + 1 + r2[0].length; pendingSheet = null; continue;
        }
        out += onRef({ kind: 'cell', sheet: pendingSheet, a: m }); i += m[0].length; pendingSheet = null; continue;
      }
    }
    out += text[i];
    if (!/[A-Za-z0-9_.$]/.test(text[i])) pendingSheet = null;
    i += 1;
  }
  return out;
}

/** Walk a formula's tokens; `fn(ref)` returns the replacement text of each reference. null = unparsable. */
function mapRefs(formula, fn) {
  const s = String(formula);
  if (!s.startsWith('=')) return s;
  const toks = tokenize(s.slice(1));
  if (!toks) return null;
  let qsheet = null;
  return '=' + toks.map(t => {
    if (t.kind === 'qname') { qsheet = t.text.slice(1, -1).replace(/''/g, "'"); return t.text; }
    if (t.kind === 'str') { qsheet = null; return t.text; }
    const before = qsheet !== null && t.text.startsWith('!') ? qsheet : null;
    qsheet = null;
    if (before !== null) return '!' + scan(t.text.slice(1), fn, before);
    return scan(t.text, fn, null);
  }).join('');
}

const onSheet = (ref, host, sheet) => (ref.sheet === null ? host : ref.sheet).toUpperCase() === String(sheet).toUpperCase();

/**
 * One axis position after a structural op: { at, count } insert (count > 0) or delete (count < 0, the block
 * at..at-count-1). Returns the new index, or null when a delete removes it.
 */
function moveIndex(i, op) {
  if (op.count > 0) return i >= op.at ? i + op.count : i;
  const last = op.at - op.count - 1;
  if (i < op.at) return i;
  if (i <= last) return null;
  return i + op.count;
}

/** A span [a, b] after the op: grows on an insert inside it, shrinks on a delete; null when entirely deleted. */
function moveSpan(a, b, op) {
  if (op.count > 0) return [a >= op.at ? a + op.count : a, b >= op.at ? b + op.count : b];
  const last = op.at - op.count - 1;
  if (a >= op.at && b <= last) return null;
  const na = a < op.at ? a : (a <= last ? op.at : a + op.count);
  const nb = b < op.at ? b : (b <= last ? op.at - 1 : b + op.count);
  return [na, nb];
}

/**
 * The formula after a structural op on `op.sheet` ({ axis: 'rows' | 'cols', at, count }), for a formula that
 * lives on `host`. Returns { formula, broken } -- broken = a reference fell into a deleted block (#REF!).
 */
function shiftFormula(formula, host, op) {
  let broken = false;
  const text = mapRefs(formula, ref => {
    const raw = ref.kind === 'range' ? `${ref.a[0]}:${ref.b[0]}` : ref.parts ? ref.parts[0] : ref.a[0];
    if (!onSheet(ref, host, op.sheet)) return raw;
    if (ref.kind === 'cell') {
      const [, cd, c, rd, r] = ref.a;
      const idx = op.axis === 'rows' ? Number(r) : colNum(c.toUpperCase());
      const n = moveIndex(idx, op);
      if (n === null) { broken = true; return '#REF!'; }
      return op.axis === 'rows' ? `${cd}${c}${rd}${n}` : `${cd}${colName(n)}${rd}${r}`;
    }
    if (ref.kind === 'range') {
      const [, cd1, c1, rd1, r1] = ref.a;
      const [, cd2, c2, rd2, r2] = ref.b;
      const span = op.axis === 'rows' ? moveSpan(Number(r1), Number(r2), op) : moveSpan(colNum(c1.toUpperCase()), colNum(c2.toUpperCase()), op);
      if (!span) { broken = true; return '#REF!'; }
      return op.axis === 'rows' ? `${cd1}${c1}${rd1}${span[0]}:${cd2}${c2}${rd2}${span[1]}` : `${cd1}${colName(span[0])}${rd1}${r1}:${cd2}${colName(span[1])}${rd2}${r2}`;
    }
    const [, d1, x1, d2, x2] = ref.parts;
    if ((ref.kind === 'rows') !== (op.axis === 'rows')) return raw;
    const n = ref.kind === 'rows' ? [Number(x1), Number(x2)] : [colNum(x1.toUpperCase()), colNum(x2.toUpperCase())];
    const span = moveSpan(n[0], n[1], op);
    if (!span) { broken = true; return '#REF!'; }
    return ref.kind === 'rows' ? `${d1}${span[0]}:${d2}${span[1]}` : `${d1}${colName(span[0])}:${d2}${colName(span[1])}`;
  });
  return { formula: text, broken };
}

/**
 * Does the formula (on `host`) reference any cell of `block` ({ sheet, r1, c1, r2, c2 })? `spread` ({ dr: [lo, hi],
 * dc: [lo, hi] }, optional) widens every RELATIVE row / column part by those offsets: a validation / conditional
 * format formula is written for its applies-to area's first cell and repeats over the whole area, so a relative
 * reference reaches every offset of the area (a superset, fail-closed).
 */
function referencesBlock(formula, host, block, spread = null) {
  const dr = spread ? spread.dr : [0, 0];
  const dc = spread ? spread.dc : [0, 0];
  let hit = false;
  mapRefs(formula, ref => {
    const raw = ref.kind === 'range' ? `${ref.a[0]}:${ref.b[0]}` : ref.parts ? ref.parts[0] : ref.a[0];
    if (!onSheet(ref, host, block.sheet)) return raw;
    // one part: [lo, hi] after the spread (an absolute part does not move)
    const part = (v, abs, d) => (abs ? [v, v] : [v + d[0], v + d[1]]);
    let r; let c;
    if (ref.kind === 'cell') { r = part(Number(ref.a[4]), ref.a[3] === '$', dr); c = part(colNum(ref.a[2].toUpperCase()), ref.a[1] === '$', dc); }
    else if (ref.kind === 'range') {
      const ra = part(Number(ref.a[4]), ref.a[3] === '$', dr); const rb = part(Number(ref.b[4]), ref.b[3] === '$', dr);
      const ca = part(colNum(ref.a[2].toUpperCase()), ref.a[1] === '$', dc); const cb = part(colNum(ref.b[2].toUpperCase()), ref.b[1] === '$', dc);
      r = [Math.min(ra[0], rb[0]), Math.max(ra[1], rb[1])]; c = [Math.min(ca[0], cb[0]), Math.max(ca[1], cb[1])];
    }
    else if (ref.kind === 'rows') { const a = part(Number(ref.parts[2]), ref.parts[1] === '$', dr); const b = part(Number(ref.parts[4]), ref.parts[3] === '$', dr); r = [Math.min(a[0], b[0]), Math.max(a[1], b[1])]; c = [1, MAX_COL]; }
    else { const a = part(colNum(ref.parts[2].toUpperCase()), ref.parts[1] === '$', dc); const b = part(colNum(ref.parts[4].toUpperCase()), ref.parts[3] === '$', dc); c = [Math.min(a[0], b[0]), Math.max(a[1], b[1])]; r = [1, MAX_ROW]; }
    if (r[0] <= block.r2 && r[1] >= block.r1 && c[0] <= block.c2 && c[1] >= block.c1) hit = true;
    return raw;
  });
  return hit;
}

/**
 * A copied formula: every RELATIVE row / column part moves by (dr, dc) -- on any sheet, as a paste does; an
 * absolute part stays. Out of the sheet -> #REF! (broken).
 */
function translateFormula(formula, dr, dc) {
  let broken = false;
  const mv = (abs, v, d, max) => { if (abs) return v; const n = v + d; if (n < 1 || n > max) { broken = true; return null; } return n; };
  const cell = m => { const [, cd, c, rd, r] = m; const nc = mv(cd === '$', colNum(c.toUpperCase()), dc, MAX_COL); const nr = mv(rd === '$', Number(r), dr, MAX_ROW); return nc === null || nr === null ? '#REF!' : `${cd}${colName(nc)}${rd}${nr}`; };
  const text = mapRefs(formula, ref => {
    if (ref.kind === 'cell') return cell(ref.a);
    if (ref.kind === 'range') { const a = cell(ref.a); const b = cell(ref.b); return a === '#REF!' || b === '#REF!' ? '#REF!' : `${a}:${b}`; }
    const [, d1, x1, d2, x2] = ref.parts;
    if (ref.kind === 'rows') { const a = mv(d1 === '$', Number(x1), dr, MAX_ROW); const b = mv(d2 === '$', Number(x2), dr, MAX_ROW); return a === null || b === null ? '#REF!' : `${d1}${a}:${d2}${b}`; }
    const a = mv(d1 === '$', colNum(x1.toUpperCase()), dc, MAX_COL); const b = mv(d2 === '$', colNum(x2.toUpperCase()), dc, MAX_COL);
    return a === null || b === null ? '#REF!' : `${d1}${colName(a)}:${d2}${colName(b)}`;
  });
  return { formula: text, broken };
}

/**
 * M7 E R7 (tier E probe m7-e-probe1-20261003.json MV / MX, one Cut, session = reopen): a formula after a Cut of
 * `mv.from` ({ sheet, r1, c1, r2, c2 }) to `mv.to` ({ sheet }) by (mv.dr, mv.dc). MEASURED: a reference ENTIRELY inside
 * the block -- relative or absolute, a cell or a range, bare or sheet-qualified -- moves with it; a cross-sheet move
 * gives a bare reference the target sheet's prefix. A reference that only PARTLY overlaps the block does not move (its
 * value then changes): `partial`. Not measured, reported: a whole-row / whole-column reference into the block (also
 * `partial`), a sheet-qualified reference into a block moved to ANOTHER sheet (`qualifiedAcross`). `moved` counts the
 * re-pointed references. null formula = unparsable.
 */
/** M7 E R7: where a cell / range reference lies against block b: 'out' | 'partial' | 'in'. */
function placeOf(ref, b) {
  if (ref.kind === 'rows' || ref.kind === 'cols') {
    const pos = x => (ref.kind === 'rows' ? Number(x) : colNum(x.toUpperCase()));
    const [lo, hi] = [pos(ref.parts[2]), pos(ref.parts[4])].sort((x, y) => x - y);
    const [x1, x2] = ref.kind === 'rows' ? [b.r1, b.r2] : [b.c1, b.c2];
    return hi >= x1 && lo <= x2 ? 'partial' : 'out';
  }
  const at = m => ({ c: colNum(m[2].toUpperCase()), r: Number(m[4]) });
  const p1 = at(ref.a);
  const p2 = ref.kind === 'range' ? at(ref.b) : p1;
  const r = [Math.min(p1.r, p2.r), Math.max(p1.r, p2.r)];
  const c = [Math.min(p1.c, p2.c), Math.max(p1.c, p2.c)];
  if (r[1] < b.r1 || r[0] > b.r2 || c[1] < b.c1 || c[0] > b.c2) return 'out';
  return r[0] < b.r1 || r[1] > b.r2 || c[0] < b.c1 || c[1] > b.c2 ? 'partial' : 'in';
}

function moveFormula(formula, host, mv) {
  const cross = String(mv.from.sheet).toUpperCase() !== String(mv.to.sheet).toUpperCase();
  const quote = n => (/^[A-Za-z_][\w.]*$/.test(n) ? n : `'${String(n).replace(/'/g, "''")}'`);
  const shifted = m => `${m[1]}${colName(colNum(m[2].toUpperCase()) + mv.dc)}${m[3]}${Number(m[4]) + mv.dr}`;
  const out = { partial: false, qualifiedAcross: false, moved: 0 };
  out.formula = mapRefs(formula, ref => {
    const raw = ref.kind === 'range' ? `${ref.a[0]}:${ref.b[0]}` : ref.parts ? ref.parts[0] : ref.a[0];
    const place = onSheet(ref, host, mv.from.sheet) ? placeOf(ref, mv.from) : 'out';
    if (place === 'partial') out.partial = true;
    else if (place === 'in' && cross && ref.sheet !== null) out.qualifiedAcross = true;
    if (place !== 'in' || out.qualifiedAcross) return raw;
    out.moved += 1;
    const moved = ref.kind === 'range' ? `${shifted(ref.a)}:${shifted(ref.b)}` : shifted(ref.a);
    return cross ? `${quote(mv.to.sheet)}!${moved}` : moved;
  });
  return out;
}

module.exports = { shiftFormula, referencesBlock, translateFormula, moveIndex, moveSpan, mapRefs, moveFormula };
