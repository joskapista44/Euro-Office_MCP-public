'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * M7 backlog X-F07b (align.readingOrder) / X-F07c (align.indent, align.shrink), per a maintainer's own
 * LIVE measurement (#ref, m7-s31c-probe1-20261005): owner-approved internal dependency (the owner,
 * owner channel, 2026-10-05), on the SAME internal Range model align.h/v/border already
 * read -- read back since Wave B READ /4 through its OWN DS Bridge capability (range.alignmentExt.read/v1).
 *
 *   align.readingOrder 'context'|'ltr'|'rtl' -> setReadingOrder(0|1|2); unset reads back the
 *     semantic default 'context' (0), never UNOBSERVABLE for that reason alone.
 *   align.indent  an integer in [0, 15] -> setIndent(n); ONLY valid with align.horizontal 'left'
 *     or 'right' in the SAME step (every other alignment silently zeroes it); never on a
 *     whole-sheet range (setIndent there calls setAngle instead).
 *   align.shrink  boolean -> setShrinkToFit(bool).
 */

const book = () => [{ name: 'Data', cells: { A1: { v: 'Name' }, C3: { v: 1 } }, merges: [] }];
const task = (range, spec, sheet = 'Data') => ({ id: 'm7-f07bc', steps: [{ intent: 'format_range', sheet, range, format: spec }] });

// M11.5a.4: indent / shrink are a Bridge WRITE served only on /6 -- this family's runs default to that runtime
// (the runtime gate itself: xlsx-b05w-runtime-gate.test.cjs).
const RT6 = '9.3.4|euro-office-bridge-v1/9.3.4.60/6';
async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, { runtime: RT6, ...script });
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}

const cell = (store, a) => store.durable.find(s => s.name === 'Data').cells[a];
const op = result => result.execution.appliedOperations[0];
const reason = result => op(result).evidence.evidence.reason;

// ---- readingOrder ----

for (const [label, value, code] of [['rtl', 'rtl', 2], ['ltr', 'ltr', 1], ['context', 'context', 0]]) {
  test(`GREEN: align.readingOrder '${label}' APPLIED, byte-exact (${code}), retry SATISFIED 0/0`, async () => {
    const store = createStore(book());
    // 'context' is the untouched default -- seed a DIFFERENT value first so this is a genuine
    // transition, not an already-default SATISFIED (rtl/ltr need no seed: untouched is 0 already).
    if (value === 'context') await run(store, task('A1', { align: { readingOrder: 'rtl' } }));
    const { result, m } = await run(store, task('A1', { align: { readingOrder: value } }));
    assert.equal(result.ok, true, JSON.stringify(result.execution.primitiveVerification));
    assert.equal(m.commands, 1);
    assert.equal(cell(store, 'A1').readingOrder, code);
    const again = await run(store, task('A1', { align: { readingOrder: value } }));
    assert.equal(again.m.commands, 0);
    assert.equal(op(again.result).status, OPERATION_STATUS.SATISFIED);
  });
}

test('GREEN: align.readingOrder on an UNTOUCHED cell is already SATISFIED at \'context\' (the measured default)', async () => {
  const { result, m } = await run(createStore(book()), task('A1', { align: { readingOrder: 'context' } }));
  assert.equal(result.ok, true, JSON.stringify(result.execution.primitiveVerification));
  assert.equal(m.commands, 0);
  assert.equal(op(result).status, OPERATION_STATUS.SATISFIED);
});

test("refused BEFORE the write, 0/0: align.readingOrder an unsupported value", () => {
  const p = planOperations(task('A1', { align: { readingOrder: 'up' } }));
  assert.equal(p.ok, false);
  assert.match(p.detail, /align\.readingOrder must be one of context, ltr, rtl/);
});

// ---- indent ----

for (const h of ['left', 'right']) {
  test(`GREEN: align.indent 7 with align.horizontal '${h}' in the SAME step -- APPLIED, retry SATISFIED 0/0`, async () => {
    const store = createStore(book());
    const { result, m } = await run(store, task('A1', { align: { indent: 7, horizontal: h } }));
    assert.equal(result.ok, true, JSON.stringify(result.execution.primitiveVerification));
    assert.equal(m.commands, 1);
    assert.equal(cell(store, 'A1').indent, 7);
    const again = await run(store, task('A1', { align: { indent: 7, horizontal: h } }));
    assert.equal(again.m.commands, 0);
    assert.equal(op(again.result).status, OPERATION_STATUS.SATISFIED);
  });
}

test("refused BEFORE the write, 0/0: align.indent WITHOUT horizontal left/right in the SAME step", () => {
  const bare = planOperations(task('A1', { align: { indent: 2 } }));
  assert.equal(bare.ok, false);
  assert.match(bare.detail, /align\.indent requires align\.horizontal 'left' or 'right'/);
  const center = planOperations(task('A1', { align: { indent: 2, horizontal: 'center' } }));
  assert.equal(center.ok, false);
  assert.match(center.detail, /align\.indent requires align\.horizontal 'left' or 'right'/);
  const distributed = planOperations(task('A1', { align: { indent: 2, horizontal: 'distributed' } }));
  assert.equal(distributed.ok, false, 'distributed is not even a supported align.horizontal value');
  assert.match(distributed.detail, /align\.horizontal must be one of/);
});

for (const bad of [-1, 16, 2.5, '2']) {
  test(`refused BEFORE the write, 0/0: align.indent out of range or non-integer (${JSON.stringify(bad)})`, () => {
    const p = planOperations(task('A1', { align: { indent: bad, horizontal: 'left' } }));
    assert.equal(p.ok, false);
    assert.match(p.detail, /align\.indent must be an integer in \[0, 15\]/);
  });
}

test('refused BEFORE the write, 0/0: align.indent on a whole-sheet range (setIndent there calls setAngle instead)', () => {
  const p = planOperations(task('A1:XFD1048576', { align: { indent: 2, horizontal: 'left' } }));
  assert.equal(p.ok, false);
  assert.match(p.detail, /align\.indent is not supported on a whole-sheet range/);
});

// ---- shrink ----

for (const value of [true, false]) {
  test(`GREEN: align.shrink ${value} APPLIED, retry SATISFIED 0/0`, async () => {
    const store = createStore(book());
    if (!value) await run(store, task('A1', { align: { shrink: true } })); // seed a non-default state first
    const { result, m } = await run(store, task('A1', { align: { shrink: value } }));
    assert.equal(result.ok, true, JSON.stringify(result.execution.primitiveVerification));
    assert.equal(m.commands, 1);
    assert.equal(cell(store, 'A1').shrink, value);
    const again = await run(store, task('A1', { align: { shrink: value } }));
    assert.equal(again.m.commands, 0);
    assert.equal(op(again.result).status, OPERATION_STATUS.SATISFIED);
  });
}

test("refused BEFORE the write, 0/0: align.shrink not a boolean", () => {
  const p = planOperations(task('A1', { align: { shrink: 'yes' } }));
  assert.equal(p.ok, false);
  assert.match(p.detail, /align\.shrink must be true or false/);
});

// ---- the build gate, SEPARATE from align.h/v/border's own ----

test('refused BEFORE the write, 0/0: an unverified align-ext build -- the three new axes unobservable, align.h/v/border unaffected', async () => {
  const { result, driver } = await run(createStore(book()), task('A1', { align: { readingOrder: 'rtl' } }), { alignExtBuild: false });
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(reason(result), 'xlsx-format-readingOrder-unobservable');
  // the SAME unverified-align-ext build does NOT block align.h, a DIFFERENT, independently-gated axis
  const { result: hResult } = await run(createStore(book()), task('A1', { align: { horizontal: 'left' } }), { alignExtBuild: false });
  assert.equal(hResult.ok, true, JSON.stringify(hResult.execution.primitiveVerification));
});

test('refused BEFORE the write, 0/0: the engine throws for THIS cell\'s readingOrder / indent / shrink reader', async () => {
  const ro = await run(createStore(book()), task('A1', { align: { readingOrder: 'rtl' } }), { readingOrderUnobservable: ['A1'] });
  assert.equal(ro.result.ok, false);
  assert.equal(reason(ro.result), 'xlsx-format-readingOrder-unobservable');
  const ind = await run(createStore(book()), task('A1', { align: { indent: 2, horizontal: 'left' } }), { indentUnobservable: ['A1'] });
  assert.equal(ind.result.ok, false);
  assert.equal(reason(ind.result), 'xlsx-format-indent-unobservable');
  const sh = await run(createStore(book()), task('A1', { align: { shrink: true } }), { shrinkUnobservable: ['A1'] });
  assert.equal(sh.result.ok, false);
  assert.equal(reason(sh.result), 'xlsx-format-shrink-unobservable');
});

// ---- the command body: internal-model write with the api-missing safety net ----

test('the command body: the PUBLIC SetReadingOrder; indent / shrink through ONE Bridge WRITE (M11.5a.4) -- no internal setter in the command', () => {
  const format = require('../../src/adapters/xlsx/families/format.cjs');
  const opn = planOperations(task('A1', { align: { readingOrder: 'rtl', indent: 3, horizontal: 'left', shrink: true } })).operations[0];
  const cmd = format.lower(opn);
  assert.doesNotMatch(cmd.body, /setIndent|setShrinkToFit|c\.range\./);
  assert.match(cmd.body, /c\.SetReadingOrder\("rtl"\)/);
  assert.match(cmd.body, /__eoB05Gate\(name\)/);
  assert.match(cmd.body, /__eoB05Write\(a, \{"indent":3,"shrink":true\}, \["indent","shrink"\]\)/);
  // the full in-editor contract against the REAL /6 block: xlsx-b05w-bridge-write-cutover.test.cjs
});
