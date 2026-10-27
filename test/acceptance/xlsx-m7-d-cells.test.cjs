'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * M7 S26 / S28 (owner M7 GO 2026-10-03; tier D probe docs/evidence/M7-D-PROBE.md): typed values in write_range --
 * { text } with '@' first, { date, format } / { time, format } with the measured masks. The number format is a
 * declared change on exactly those cells.
 */

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}
const book = () => [{ name: 'T', cells: { A1: { v: 'id' }, B9: { v: 5, nf: '0.00' } } }, { name: 'O', cells: { A1: { v: 'keep' } } }];
const write = values => ({ id: 'm7-d-s26', steps: [{ intent: 'write_range', sheet: 'T', range: `A2:A${values.length + 1}`, values: values.map(v => [v]) }] });
const T = store => store.durable.find(s => s.name === 'T').cells;
const op = result => result.execution.appliedOperations[0];

// the probe's six coercible texts, written out (m7-d-probe1-20261003.json R2.TX column B)
const TEXTS = ['001', '1/2', 'TRUE', ' a ', '2026-10-03']; // '=x' kept too, but reads back as a formula -> refused

test('S26 APPLIED: five coercible texts kept exactly with nf @ (COUNT 0); retry SATISFIED 0/0', async () => {
  const store = createStore(book());
  const t = write(TEXTS.map(text => ({ text })));
  const { result, m } = await run(store, t);
  assert.equal(result.ok, true, JSON.stringify([result.execution.primitiveVerification, result.execution.preservationVerification]));
  assert.equal(m.commands, 1);
  TEXTS.forEach((v, i) => { assert.equal(T(store)[`A${i + 2}`].v, v); assert.equal(T(store)[`A${i + 2}`].nf, '@'); });
  assert.equal(T(store).B9.nf, '0.00', 'another cell keeps its number format');
  const again = await run(store, t);
  assert.equal(again.m.commands, 0);
  assert.equal(op(again.result).status, OPERATION_STATUS.SATISFIED);
});

test('S26: the command sets @ BEFORE the value (the measured order); the plain path stays unchanged', () => {
  const p = planOperations(write([{ text: '001' }, 'plain']));
  assert.equal(p.ok, true);
  const body = require('../../src/adapters/xlsx/families/cells.cjs').lower(p.operations[0]).body;
  assert.match(body, /if \(p\.order === 'nf-first'\) \{ r\.SetNumberFormat\(p\.nf\); r\.SetValue\(p\.v\); \}/);
  assert.match(body, /"a":"A3","v":"plain"\}/);
});

test('S28 APPLIED: dates and a time as serial + measured mask; retry 0/0', async () => {
  const store = createStore(book());
  const t = write([{ date: '2026-10-03', format: 'yyyy-mm-dd' }, { date: '2026-10-03', format: 'dd/mm/yyyy' }, { date: '2026-10-03T18:00', format: 'yyyy-mm-dd hh:mm' }, { time: '12:00', format: 'hh:mm' }]);
  const { result } = await run(store, t);
  assert.equal(result.ok, true, JSON.stringify(result.execution.primitiveVerification));
  // the probe's serials, written out (m7-d-probe1-20261003.json R2.TX column G)
  assert.deepEqual([2, 3, 4, 5].map(r => T(store)[`A${r}`].v), [46298, 46298, 46298.75, 0.5]);
  assert.deepEqual([2, 3, 4, 5].map(r => T(store)[`A${r}`].nf), ['yyyy-mm-dd', 'dd/mm/yyyy', 'yyyy-mm-dd hh:mm', 'hh:mm']);
  assert.equal((await run(store, t)).m.commands, 0);
});

test('S26/S28 refused at planning: empty text, an unmeasured mask, an impossible date, a date with a time mask, unknown keys', () => {
  for (const [v, re] of [[{ text: '' }, /non-empty/], [{ date: '2026-10-03', format: 'mmm d' }, /measured masks/], [{ date: '2026-02-30', format: 'yyyy-mm-dd' }, /not a valid date/],
    [{ date: '2026-10-03', format: 'yyyy-mm-dd hh:mm' }, /measured masks/], [{ text: 'a', nf: '@' }, /typed value is/], [{ time: '25:00', format: 'hh:mm' }, /HH:MM/]]) {
    // X-T01 (LIVE xt01-r1): { text: '=x' } / { text: '+1' } are measured and SUPPORTED now (xt01-cell-value-typing.test.cjs)
    const p = planOperations(write([v]));
    assert.equal(p.ok, false, JSON.stringify(v));
    assert.match(p.detail, re);
  }
  assert.equal(planOperations(write(['001'])).ok, false, 'a bare coercible string stays refused');
});

test('S26/S28 RED: order swapped (coerced) / format not applied -> L1 FAIL; the format spills to a neighbour -> L3 FAIL', async () => {
  for (const fault of ['typed-order-swapped', 'typed-no-format']) {
    const { result } = await run(createStore(book()), write([{ text: '001' }]), { fault });
    assert.equal(result.ok, false, fault);
    assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL', fault);
  }
  const b = book();
  b[0].cells.A3 = { v: 'x' };
  const { result } = await run(createStore(b), write([{ text: '001' }]), { fault: 'typed-format-spills' });
  assert.equal(result.ok, false);
  assert.equal(result.execution.preservationVerification[0].verdict, 'FAIL', JSON.stringify(result.execution.preservationVerification));
});

test('S26/S28 RED: a numeric-looking text coerced to a number (the string still equal, COUNT 1) / a date one minute off -> L1 FAIL', async () => {
  const coerced = await run(createStore(book()), write([{ text: '5' }]), { fault: 'typed-order-swapped' });
  assert.equal(coerced.result.execution.primitiveVerification[0].verdict, 'FAIL');
  const minute = await run(createStore(book()), write([{ date: '2026-10-03T18:00', format: 'yyyy-mm-dd hh:mm' }]), { fault: 'typed-minute-off' });
  assert.equal(minute.result.execution.primitiveVerification[0].verdict, 'FAIL');
});

test('S26 LIVE regression (m7-d-acc1): empty targets OUTSIDE the used range get their number format read through the window', async () => {
  const store = createStore([{ name: 'T', cells: { A1: { v: 'id' } } }, { name: 'O', cells: { A1: { v: 'keep' } } }]);
  const { result } = await run(store, write([{ text: '001' }, { date: '2026-10-03', format: 'yyyy-mm-dd' }]));
  assert.equal(result.ok, true, JSON.stringify(op(result).evidence));
  assert.equal(T(store).A2.nf, '@');
});

/**
 * X-F11: three masks added to DATE_MASKS on the strength of the X-F09 S23 probe
 * (m7-s23-probe1, byte-exact through two reopens) -- but that probe wrote the mask onto an
 * ALREADY-NUMERIC cell; it never measured these masks through the TYPED write_range order
 * (SetValue(serial) then SetNumberFormat). The session-level test below proves that order and
 * the serials; the LIVE harness (m7-f11-date-masks) is what actually measures the real engine.
 */
const NEW_MASKS = [
  [{ date: '2026-10-03', format: '[$-409]mmmm d, yyyy' }, '[$-409]mmmm d, yyyy'],
  [{ date: '2026-10-03', format: '[$-40E]yyyy. mmmm d.' }, '[$-40E]yyyy. mmmm d.'],
  [{ date: '2026-10-03T18:00', format: 'dd/mm/yyyy hh:mm' }, 'dd/mm/yyyy hh:mm'],
];

test('X-F11: each of the three newly measured masks is accepted at planning, INDIVIDUALLY -- dropping any one mask fails exactly this check for it, not the others', () => {
  for (const [v, label] of NEW_MASKS) {
    const p = planOperations(write([v]));
    assert.equal(p.ok, true, `${label}: ${p.ok ? '' : p.detail}`);
  }
});

test('X-F11 APPLIED: the three new masks write as serial + mask (the measured value-first order), retry 0/0', async () => {
  const store = createStore(book());
  const t = write(NEW_MASKS.map(([v]) => v));
  const { result } = await run(store, t);
  assert.equal(result.ok, true, JSON.stringify(result.execution.primitiveVerification));
  assert.deepEqual([2, 3, 4].map(r => T(store)[`A${r}`].v), [46298, 46298, 46298.75]);
  assert.deepEqual([2, 3, 4].map(r => T(store)[`A${r}`].nf), NEW_MASKS.map(([, label]) => label));
  assert.equal((await run(store, t)).m.commands, 0);
});

test('X-F11: the excluded duration mask [h]:mm:ss stays refused -- out of scope on purpose (zero\'s own line), a separate backlog item', () => {
  const p = planOperations(write([{ time: '12:00', format: '[h]:mm:ss' }]));
  assert.equal(p.ok, false);
  assert.match(p.detail, /format for a time must be one of/);
});

test('X-F11: an unmeasured date mask still refused, and a plain date-looking string stays COERCIBLE-refused (no regression from the catalogue growing)', () => {
  const p1 = planOperations(write([{ date: '2026-10-03', format: 'mmm d' }]));
  assert.equal(p1.ok, false);
  assert.match(p1.detail, /measured masks/);
  const p2 = planOperations(write(['2026-10-05']));
  assert.equal(p2.ok, false, 'a bare date-looking string stays refused');
});
