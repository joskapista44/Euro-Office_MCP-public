'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const format = require('../../src/adapters/xlsx/families/format.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * M7 backlog X-B05 (diagonal borders), per a maintainer's LIVE measurement (#ref, m7-b05-probe1):
 *   - the engine models DiagonalDown and DiagonalUp with ONE SHARED `border.d` style+colour, plus two
 *     INDEPENDENT boolean flags (`border.dd` / `border.du`) marking which direction(s) are drawn.
 *   - one direction alone, or both together with the SAME style+colour: precise, reopen-stable.
 *   - two DIFFERENT styles across the two directions: the LAST write wins for BOTH (the shared `d`) --
 *     not representable, must be REFUSED before dispatch.
 *   - 'None' does not clear a diagonal (same quirk as the other sides); diagonal removal is OUT OF
 *     SCOPE for this unit (no remove-mode support added here -- matches the backlog row and format.cjs's
 *     own existing scope-boundary note on diagonal restore).
 *   - a black diagonal (colour 0) reads back `c: null` after a reopen -- the SAME tolerance already used
 *     for every other border side (`holds()`'s generic branch), nothing diagonal-specific needed.
 */

const book = () => [
  { name: 'Data', cells: { A1: { v: 'Name' }, C3: { v: 1 } }, merges: [] },
];

const task = (range, spec, sheet = 'Data') => ({ id: 'm7-b05-diagonal', steps: [{ intent: 'format_range', sheet, range, format: spec }] });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}

const cell = (store, a) => store.durable.find(s => s.name === 'Data').cells[a];
const op = result => result.execution.appliedOperations[0];
const reason = result => op(result).evidence.evidence.reason;

const THIN_BLUE = { style: 'thin', color: '#0000FF' };
const MEDIUM_RED = { style: 'medium', color: '#FF0000' };

test('GREEN: DiagonalDown alone sets d + dd; du stays false', async () => {
  const store = createStore(book());
  const { result, m } = await run(store, task('C3', { border: { sides: ['diagonalDown'], ...THIN_BLUE } }));
  assert.equal(result.ok, true, JSON.stringify(result.execution.primitiveVerification));
  assert.equal(m.writes, 1);
  assert.deepEqual(cell(store, 'C3').border, { d: { s: 7, c: 255 }, dd: true });
});

test('GREEN: DiagonalUp alone sets d + du; dd stays false', async () => {
  const store = createStore(book());
  const { result } = await run(store, task('C3', { border: { sides: ['diagonalUp'], ...THIN_BLUE } }));
  assert.equal(result.ok, true, JSON.stringify(result.execution.primitiveVerification));
  assert.deepEqual(cell(store, 'C3').border, { d: { s: 7, c: 255 }, du: true });
});

test('GREEN: both directions together, SAME style, in one call -- precise and stable', async () => {
  const store = createStore(book());
  const { result } = await run(store, task('C3', { border: { sides: ['diagonalDown', 'diagonalUp'], ...THIN_BLUE } }));
  assert.equal(result.ok, true, JSON.stringify(result.execution.primitiveVerification));
  assert.deepEqual(cell(store, 'C3').border, { d: { s: 7, c: 255 }, dd: true, du: true });
  const again = await run(store, task('C3', { border: { sides: ['diagonalDown', 'diagonalUp'], ...THIN_BLUE } }));
  assert.equal(again.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
  assert.equal(again.m.commands, 0);
});

test('GREEN: requesting one direction when the OTHER already carries the SAME style is not a conflict', async () => {
  const b = book();
  b[0].cells.C3 = { ...b[0].cells.C3, border: { d: { s: 7, c: 255 }, du: true } };
  const store = createStore(b);
  const { result } = await run(store, task('C3', { border: { sides: ['diagonalDown'], ...THIN_BLUE } }));
  assert.equal(result.ok, true, JSON.stringify(result.execution.primitiveVerification));
  assert.deepEqual(cell(store, 'C3').border, { d: { s: 7, c: 255 }, du: true, dd: true });
});

test('REFUSED before dispatch: DiagonalDown requested while DiagonalUp already differs (the shared `d` would silently change it too)', async () => {
  const b = book();
  b[0].cells.C3 = { ...b[0].cells.C3, border: { d: { s: 12, c: 16711680 }, du: true } }; // medium/red already up
  const store = createStore(b);
  const { result, driver } = await run(store, task('C3', { border: { sides: ['diagonalDown'], ...THIN_BLUE } }));
  assert.equal(result.ok, false);
  assert.equal(reason(result), 'xlsx-format-diagonal-conflicting-style');
  assert.equal(driver.calls.dispatch, 0);
  assert.deepEqual(store.durable, b, 'nothing written');
});

test('REFUSED before dispatch, symmetric: DiagonalUp requested while DiagonalDown already differs', async () => {
  const b = book();
  b[0].cells.C3 = { ...b[0].cells.C3, border: { d: { s: 7, c: 255 }, dd: true } }; // thin/blue already down
  const store = createStore(b);
  const { result, driver } = await run(store, task('C3', { border: { sides: ['diagonalUp'], ...MEDIUM_RED } }));
  assert.equal(result.ok, false);
  assert.equal(reason(result), 'xlsx-format-diagonal-conflicting-style');
  assert.equal(driver.calls.dispatch, 0);
});

test('REFUSED before dispatch (a maintainer review): the OTHER direction flag unreadable -- not taken as "not drawn" (fail-closed)', async () => {
  const b = book();
  b[0].cells.C3 = { ...b[0].cells.C3, border: { d: { s: 12, c: 16711680 }, du: true } };
  const store = createStore(b);
  const { result, driver } = await run(store, task('C3', { border: { sides: ['diagonalDown'], ...THIN_BLUE } }), { diagonalFlagUnobservable: 'du' });
  assert.equal(result.ok, false);
  assert.equal(reason(result), 'xlsx-format-diagonal-unobservable');
  assert.equal(driver.calls.dispatch, 0);
  assert.deepEqual(store.durable, b, 'nothing written');
});

test('RED: a write that sets the shared style but silently fails to raise its own direction flag -> L1 FAIL', async () => {
  const { result, driver } = await run(createStore(book()), task('C3', { border: { sides: ['diagonalDown'], ...THIN_BLUE } }), { fault: 'diagonal-missing-flag' });
  assert.equal(result.ok, false);
  assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL');
  assert.equal(driver.calls.forceSave, 0);
});

test('planner: diagonal sides never mix with top/bottom/left/right, and \'none\' is refused the same way for diagonal as for every other side', () => {
  const refuse = (spec, re) => { const p = planOperations(task('C3', spec)); assert.equal(p.ok, false, JSON.stringify(spec)); assert.match(p.detail, re); };
  refuse({ border: { sides: ['diagonalDown', 'top'], ...THIN_BLUE } }, /border.sides must be a non-empty list of distinct top \/ bottom \/ left \/ right/);
  refuse({ border: { sides: ['diagonalDown'], style: 'none', color: '#000000' } }, /border\.style 'none' is not a style/);
  const ok = planOperations(task('C3', { border: { sides: ['diagonalDown'], ...THIN_BLUE } }));
  assert.equal(ok.ok, true);
  assert.deepEqual(ok.operations[0].diagonal, { codes: ['dd'], expect: { s: 7, c: 255 } });
});

test('the command body calls SetBorders(\'DiagonalDown\', ...) / SetBorders(\'DiagonalUp\', ...), not a regular side setter', () => {
  const opn = planOperations(task('C3', { border: { sides: ['diagonalDown', 'diagonalUp'], ...THIN_BLUE } })).operations[0];
  const cmd = format.lower(opn);
  const calls = [];
  const rec = a => new Proxy({}, { get: (_t, m) => (...args) => calls.push([a, m, ...args]) });
  const ws = { GetName: () => 'Data', GetRange: a => rec(a) };
  const Api = { GetSheet: () => ws, CreateColorFromRGB: (r, g, b) => `rgb(${r},${g},${b})` };
  assert.deepEqual(new Function('Api', cmd.body)(Api), { ok: true });
  assert.deepEqual(calls, [['C3', 'SetBorders', 'DiagonalDown', 'Thin', 'rgb(0,0,255)'], ['C3', 'SetBorders', 'DiagonalUp', 'Thin', 'rgb(0,0,255)']]);
});

test('a black diagonal (colour 0) reading c:null after a reopen is still SATISFIED (the existing border tolerance, unmodified for diagonal)', async () => {
  const store = createStore(book());
  const spec = { border: { sides: ['diagonalDown'], style: 'thin', color: '#000000' } };
  await run(store, task('C3', spec));
  const again = await run(store, task('C3', spec), { blackReadsNull: true });
  assert.equal(again.m.commands, 0);
  assert.equal(again.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});
