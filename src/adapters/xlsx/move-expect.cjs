'use strict';

const formula = require('./formula.cjs');
const { isBlank } = require('./inventory.cjs');

/**
 * The move's CARRIED block expectation (owner decision 2026-09-30, owner channel). Pure.
 *
 *   { version: 1, fileId, source: { sheet, range }, target: { sheet, range }, cells }
 *   cells: rows x cols (the source's shape, row-major) of { t: 'blank' } | { t: 'number', v } | { t: 'text', v } |
 *          { t: 'formula', f } -- v as GetValue read it, f the formula TEXT
 *
 * Built from a fresh live read BEFORE the first write and kept in the repeatable task; bound to the document, both
 * sheets and both ranges. A cell matches by type (COUNT: 1 number / 0 not a number), blank, value as read, and for a
 * formula by its TEXT only -- a formula's computed result is not content. Terminal-state evidence, not proof of origin.
 */

const EXPECT_VERSION = 1;

/**
 * A text value the reader cannot tell from another type: GetValue gives a string and COUNT only separates numbers,
 * so 'TRUE' / 'FALSE' and the error literals may be a boolean / an error as well as text (not measured). Refused.
 */
const TYPE_AMBIGUOUS = /^(TRUE|FALSE|#NULL!|#DIV\/0!|#VALUE!|#REF!|#NAME\?|#NUM!|#N\/A|#GETTING_DATA)$/i;

const keysOf = c => Object.keys(c).sort().join(',');
const valued = c => keysOf(c) === 't,v' && typeof c.v === 'string' && c.v !== '';

/** Per type: the checks in order, each [holds, why]; the first that fails names the problem. */
const CELL_RULES = Object.freeze({
  blank: [[c => keysOf(c) === 't', 'blank takes no other field']],
  number: [[valued, 'number must be { t, v } with a non-empty string v (as read)'], [c => Number.isFinite(Number(c.v)), 'number v must read as a finite number']],
  text: [[valued, 'text must be { t, v } with a non-empty string v (as read)'], [c => !TYPE_AMBIGUOUS.test(c.v), 'type is not observable (boolean / error literal read as text)']],
  formula: [[c => keysOf(c) === 'f,t' && typeof c.f === 'string' && c.f.startsWith('='), 'formula must be { t, f } with f starting with =']],
});

function checkCell(c) {
  const rules = c && typeof c === 'object' && !Array.isArray(c) ? CELL_RULES[c.t] : null;
  if (!rules) return "must be an object whose t is 'blank' | 'number' | 'text' | 'formula'";
  const bad = rules.find(([ok]) => !ok(c));
  return bad ? bad[1] : null;
}

const sameBlock = (x, y) => !!x && typeof x === 'object' && JSON.stringify([keysOf(x), x.sheet, x.range]) === JSON.stringify(['range,sheet', y.sheet, y.range]);
const shaped = (cells, rows, cols) => Array.isArray(cells) && cells.length === rows && cells.every(r => Array.isArray(r) && r.length === cols);

function firstCellProblem(cells) {
  for (const [i, r] of cells.entries()) for (const [j, c] of r.entries()) { const why = checkCell(c); if (why) return `expect.cells[${i}][${j}]: ${why}`; }
  return cells.every(r => r.every(c => c.t === 'blank')) ? 'expect has only blank cells: nothing to prove after the move' : null;
}

/** The planner's check of a carried expectation against the step's own source / target blocks and shape. */
function checkExpect(e, source, target, rows, cols) {
  const isObj = !!e && typeof e === 'object' && !Array.isArray(e);
  const extra = isObj ? Object.keys(e).filter(k => !['version', 'fileId', 'source', 'target', 'cells'].includes(k)) : [];
  const problem = [
    [isObj, 'expect must be the carried block expectation { version, fileId, source, target, cells } (moveExpectation, from a fresh read before the first write)'],
    [() => extra.length === 0, `expect does not support field(s): ${extra.join(', ')}`],
    [() => e.version === EXPECT_VERSION, `expect.version must be ${EXPECT_VERSION}`],
    [() => typeof e.fileId === 'string' && e.fileId !== '', 'expect.fileId must name the document'],
    [() => sameBlock(e.source, source), `expect.source must be ${source.sheet}!${source.range}`],
    [() => sameBlock(e.target, target), `expect.target must be ${target.sheet}!${target.range}`],
    [() => shaped(e.cells, rows, cols), `expect.cells must be ${rows} x ${cols}`],
  ].find(([ok]) => !(typeof ok === 'function' ? ok() : ok));
  const why = problem ? problem[1] : firstCellProblem(e.cells);
  if (why) return { error: why };
  const clone = JSON.parse(JSON.stringify(e));
  return { expect: Object.freeze({ ...clone, source: Object.freeze(clone.source), target: Object.freeze(clone.target), cells: Object.freeze(clone.cells.map(r => Object.freeze(r.map(Object.freeze)))) }) };
}

const sameText = (want, got) => { const a = formula.canonical(want); const b = formula.canonical(got); return a !== null && a === b; };

/** Per expected type: does the read cell match (blank / value by type / formula by text). */
const MATCH = Object.freeze({
  blank: c => isBlank(c),
  formula: (c, e) => c.f !== '' && sameText(e.f, c.f),
  number: (c, e) => c.f === '' && !isBlank(c) && c.count === 1 && String(c.v) === e.v,
  text: (c, e) => c.f === '' && !isBlank(c) && c.count === 0 && String(c.v) === e.v,
});

const matches = (c, e) => !!c && MATCH[e.t](c, e);

/** A read cell -> its expectation cell, or null when its type is not observable. */
function cellOf(c) {
  if (c.f !== '') return { t: 'formula', f: c.f };
  if (isBlank(c)) return { t: 'blank' };
  if (c.count === 1) return { t: 'number', v: String(c.v) };
  return c.count === 0 && !TYPE_AMBIGUOUS.test(String(c.v)) ? { t: 'text', v: String(c.v) } : null;
}

module.exports = { EXPECT_VERSION, TYPE_AMBIGUOUS, checkCell, checkExpect, matches, cellOf };
