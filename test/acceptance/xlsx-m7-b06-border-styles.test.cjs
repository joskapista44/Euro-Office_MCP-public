'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');
const format = require('../../src/adapters/xlsx/families/format.cjs');

/**
 * M7 backlog X-B06 (additional border styles), per a maintainer's LIVE measurement (#ref, m7-b06-probe1):
 * all 9 candidate styles (the exact names from the audit, S05) each read back a DISTINCT numeric code,
 * none colliding with the four already-measured styles (thin=7/medium=12/thick=13/dashed=6) or with
 * each other, reopen-stable. A purely additive map extension -- no new refusal class, no new axis
 * (unlike X-B05): every write path (regular sides, outline/inside/all, and diagonal) already looks the
 * style name up in the SAME STYLE/API_STYLE maps, so all of them gain the 9 styles for free.
 */

const book = () => [{ name: 'Data', cells: { C3: { v: 1 } }, merges: [] }];
const task = (range, spec, sheet = 'Data') => ({ id: 'm7-b06-styles', steps: [{ intent: 'format_range', sheet, range, format: spec }] });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}

const cell = (store, a) => store.durable.find(s => s.name === 'Data').cells[a];

const STYLES = Object.freeze({
  double: 1, hair: 2, dashDotDot: 3, dashDot: 4, dotted: 5, mediumDashDotDot: 8, slantDashDot: 9, mediumDashDot: 10, mediumDashed: 11,
});

for (const [style, code] of Object.entries(STYLES)) {
  test(`GREEN: border.style '${style}' writes code ${code} on a regular side, round-trips SATISFIED`, async () => {
    const store = createStore(book());
    const { result, m } = await run(store, task('C3', { border: { sides: ['top'], style, color: '#0000FF' } }));
    assert.equal(result.ok, true, JSON.stringify(result.execution.primitiveVerification));
    assert.equal(m.writes, 1);
    assert.equal(cell(store, 'C3').border.t.s, code);
    const again = await run(store, task('C3', { border: { sides: ['top'], style, color: '#0000FF' } }));
    assert.equal(again.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
    assert.equal(again.m.commands, 0);
  });
}

test('GREEN: a new style also works through X-B05 diagonal (the same STYLE/API_STYLE maps)', async () => {
  const store = createStore(book());
  const { result } = await run(store, task('C3', { border: { sides: ['diagonalDown'], style: 'double', color: '#0000FF' } }));
  assert.equal(result.ok, true, JSON.stringify(result.execution.primitiveVerification));
  assert.deepEqual(cell(store, 'C3').border, { d: { s: 1, c: 255 }, dd: true });
});

test('GREEN: a new style also works through \'all\' mode (outline + inside)', async () => {
  const store = createStore([{ name: 'Data', cells: { A1: { v: 1 }, B1: { v: 1 } } }]);
  const { result } = await run(store, task('A1:B1', { border: { mode: 'all', style: 'hair', color: '#0000FF' } }));
  assert.equal(result.ok, true, JSON.stringify(result.execution.primitiveVerification));
});

test('an unknown style name is still refused, and the error lists all 13 measured styles now', () => {
  const p = planOperations(task('C3', { border: { sides: ['top'], style: 'double-thick', color: '#0000FF' } }));
  assert.equal(p.ok, false);
  for (const name of ['thin', 'medium', 'thick', 'dashed', ...Object.keys(STYLES)]) assert.match(p.detail, new RegExp(name));
});

test('no two of the 9 new styles, nor any new style against an original one, collide on the numeric code', () => {
  const all = { thin: 7, medium: 12, thick: 13, dashed: 6, ...STYLES };
  const codes = Object.values(all);
  assert.equal(new Set(codes).size, codes.length, JSON.stringify(all));
});

// a maintainer review: the fake reads the spec, not the command, so the API style NAME in the command body needs its own test
const API_NAMES = { double: 'Double', hair: 'Hair', dashDotDot: 'DashDotDot', dashDot: 'DashDot', dotted: 'Dotted', mediumDashDotDot: 'MediumDashDotDot', slantDashDot: 'SlantDashDot', mediumDashDot: 'MediumDashDot', mediumDashed: 'MediumDashed' };
test('the command body passes each new style by its EO API name to SetBorders', () => {
  for (const [style, api] of Object.entries(API_NAMES)) {
    const opn = planOperations(task('C3', { border: { sides: ['top'], style, color: '#0000FF' } })).operations[0];
    const calls = [];
    const rec = a => new Proxy({}, { get: (_t, m) => (...args) => calls.push([a, m, ...args]) });
    const ws = { GetName: () => 'Data', GetRange: a => rec(a) };
    const Api = { GetSheet: () => ws, CreateColorFromRGB: (r, g, b) => `rgb(${r},${g},${b})` };
    assert.deepEqual(new Function('Api', format.lower(opn).body)(Api), { ok: true }, style);
    assert.deepEqual(calls.filter(c => c[1] === 'SetBorders'), [['C3', 'SetBorders', 'Top', api, 'rgb(0,0,255)']], style);
  }
});
