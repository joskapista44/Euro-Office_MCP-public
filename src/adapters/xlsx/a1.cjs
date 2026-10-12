'use strict';

/**
 * A1 addressing for the XLSX adapter. Pure; no editor knowledge.
 *
 * Only plain rectangular A1 ranges on ONE sheet are accepted: a sheet-qualified address (`Sheet!A1`)
 * is refused, because the sheet is a separate, exactly matched field of the task.
 */

const MAX_COL = 16384; // XFD
const MAX_ROW = 1048576;

function colName(n) {
  let s = '';
  for (let x = n; x > 0; x = Math.floor((x - 1) / 26)) s = String.fromCharCode(65 + ((x - 1) % 26)) + s;
  return s;
}

function colNum(s) {
  let n = 0;
  for (const ch of s) n = n * 26 + ch.charCodeAt(0) - 64;
  return n;
}

function parseCell(text) {
  const m = String(text).match(/^\$?([A-Z]{1,3})\$?([1-9]\d{0,6})$/);
  if (!m) return null;
  const c = colNum(m[1]);
  const r = Number(m[2]);
  if (c > MAX_COL || r > MAX_ROW) return null;
  return { c, r };
}

/** 'B2' | 'B2:D5' -> { c1, r1, c2, r2, rows, cols, count } ; anything else -> null. */
function parseRange(text) {
  if (typeof text !== 'string' || text.includes('!')) return null;
  const parts = text.trim().toUpperCase().split(':');
  if (parts.length > 2) return null;
  const a = parseCell(parts[0]);
  const b = parseCell(parts[1] || parts[0]);
  if (!a || !b || b.c < a.c || b.r < a.r) return null;
  const rows = b.r - a.r + 1;
  const cols = b.c - a.c + 1;
  return { c1: a.c, r1: a.r, c2: b.c, r2: b.r, rows, cols, count: rows * cols };
}

const addr = (c, r) => `${colName(c)}${r}`;
const rangeText = g => (g.count === 1 ? addr(g.c1, g.r1) : `${addr(g.c1, g.r1)}:${addr(g.c2, g.r2)}`);

/** The cell addresses of a parsed range, row-major. */
function cellsOf(g) {
  const out = [];
  for (let r = g.r1; r <= g.r2; r += 1) for (let c = g.c1; c <= g.c2; c += 1) out.push(addr(c, r));
  return out;
}

const overlaps = (a, b) => a.c1 <= b.c2 && b.c1 <= a.c2 && a.r1 <= b.r2 && b.r1 <= a.r2;

module.exports = { parseRange, parseCell, cellsOf, overlaps, addr, rangeText, colName, colNum, MAX_COL, MAX_ROW };
