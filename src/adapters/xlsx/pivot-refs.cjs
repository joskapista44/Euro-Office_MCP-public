'use strict';

const readText = x => typeof x === 'string' && !x.startsWith('UNOBSERVABLE');
const readList = x => Array.isArray(x) && x.every(y => typeof y === 'string');
const readLayout = l => !!l && ['rows', 'cols', 'data'].every(k => readList(l[k]));
const wholeRef = r => [readText(r.dataRef), readList(r.fields), readText(r.table), readLayout(r.layout)].every(Boolean);

/** One sheet's raw pivot identities, or the reason the sheet's list is not usable. */
function refsOf(o) {
  if (!o || !Array.isArray(o.pivots)) return o && typeof o.pivots === 'string' ? o.pivots : 'UNOBSERVABLE:pivots-not-read';
  const out = [];
  for (const p of o.pivots) {
    const r = { name: p && p.name, dataRef: p && p.dataRef, fields: p && p.fields, table: p && p.table, layout: p && p.layout };
    if (!wholeRef(r)) return `UNOBSERVABLE:pivot-ref-unread (pivot ${r.name})`;
    out.push(r);
  }
  return out;
}

/**
 * M7 P3 refresh_pivot (owner GO 2026-10-04, owner channel): every read pivot's RAW identity by host sheet --
 * { name, dataRef (the approved asc_getDataRef READ), fields, table, layout } -- independent of the content proof (pivot-source.cjs).
 * A refresh binds a source the CALLER names and verifies the exact reference; it never infers which sheet an old, stale
 * source was (a stale pivot fails the proof, which is exactly the case a refresh exists for).
 */
const pivotRefsOf = objects => Object.fromEntries(Object.entries(objects || {}).map(([n, o]) => [n, refsOf(o)]));

module.exports = { pivotRefsOf };
