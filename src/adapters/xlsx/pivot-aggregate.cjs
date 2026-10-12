'use strict';

/**
 * M7 tier H P9 -- the arithmetic of the pivot-source proof (pivot-source.cjs): the measured layout's sums (one row field,
 * one column field, `Sum of <field>`) recomputed from a range's values and compared with a pivot's cached body. Pure.
 */

const num = x => (x === '' || x === null || x === undefined || typeof x === 'boolean' ? NaN : Number(x));
const same = (a, b) => Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(a), Math.abs(b));

/** The source values -> { rows: Map(rowKey -> Map(colKey -> sum)), rowKeys, colKeys } or a reason. */
function aggregate(values, fields, rowField, colField, dataField) {
  if (!Array.isArray(values) || values.length < 2 || !values.every(r => Array.isArray(r) && r.length === fields.length)) return { why: 'pivot-source-shape' };
  if (values[0].some((h, i) => String(h) !== fields[i])) return { why: 'pivot-source-header-mismatch' };
  const [ri, ci, di] = [rowField, colField, dataField].map(f => fields.indexOf(f));
  const sums = new Map();
  const rowKeys = new Set();
  const colKeys = new Set();
  for (const r of values.slice(1)) {
    const [rk, ck, v] = [String(r[ri]), String(r[ci]), num(r[di])];
    if (rk === '' || ck === '') return { why: 'pivot-source-blank-key-not-measured' };
    if (!Number.isFinite(v)) return { why: 'pivot-source-value-not-measured' };
    rowKeys.add(rk);
    colKeys.add(ck);
    const row = sums.get(rk) || new Map();
    row.set(ck, (row.get(ck) || 0) + v);
    sums.set(rk, row);
  }
  return { sums, rowKeys, colKeys };
}

const sameSet = (labels, keys) => labels.length === keys.size && new Set(labels).size === labels.length && labels.every(l => keys.has(l));

/** One body row against its source row: -> { total } or { why }. */
function rowCheck(row, got, cl, colTotal) {
  let total = 0;
  for (let j = 0; j < cl.length; j += 1) {
    const want = row.has(cl[j]) ? row.get(cl[j]) : null;
    if (want === null) { if (got[j] !== '' && got[j] !== null) return { why: 'pivot-body-differs' }; continue; }
    if (!same(num(got[j]), want)) return { why: 'pivot-body-differs' };
    total += want;
    colTotal[j] += want;
  }
  return same(num(got[cl.length]), total) ? { total } : { why: 'pivot-row-total-differs' };
}

/** Does the aggregate reproduce the body (labels from the table) exactly? -> null, or the reason it does not. */
function bodyMismatch(agg, body, rowLabels, colLabels) {
  const [nr, nc] = [body.length, body[0].length];
  const rl = rowLabels.slice(0, nr - 1).map(String);
  const cl = colLabels.slice(0, nc - 1).map(String);
  if (!sameSet(rl, agg.rowKeys) || !sameSet(cl, agg.colKeys)) return 'pivot-labels-differ';
  const colTotal = new Array(nc - 1).fill(0);
  let grand = 0;
  for (let i = 0; i < nr - 1; i += 1) {
    const r = rowCheck(agg.sums.get(rl[i]), body[i], cl, colTotal);
    if (r.why) return r.why;
    grand += r.total;
  }
  if (colTotal.some((t, j) => !same(num(body[nr - 1][j]), t))) return 'pivot-column-total-differs';
  return same(num(body[nr - 1][nc - 1]), grand) ? null : 'pivot-grand-total-differs';
}

module.exports = { aggregate, bodyMismatch };
