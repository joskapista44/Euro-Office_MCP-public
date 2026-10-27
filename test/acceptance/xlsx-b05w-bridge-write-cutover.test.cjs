'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS, UNCERTAINTY_REASON } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const format = require('../../src/adapters/xlsx/families/format.cjs');
const drv = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/**
 * M11.5a.4.1 (owner GO the owner, owner channel, internal tracker): B05-WRITE product cutover to the Bridge WRITE
 * capability. Part A runs the PRODUCT's lowered command + the driver's injected client against the REAL bridge block
 * (fake Api whose every touch is recorded). Part B runs the canonical executor over the fake driver, which returns what
 * the real driver returns for each body outcome (xlsx-driver b05DispatchResult, pinned in part A).
 *
 * DS9346 (owner D5=B, VALTAS, docs/DS9346-RUNTIME-RETIREMENT.md): the 9.3.4.60 family (/5 through /16) is retired --
 * ACCEPTED carries ONE entry now (9.3.5.0/16), so V6 is the 9.3.5.0 block (ALL capabilities in one file, unlike the
 * 9.3.4.60 family's per-revision rev6/rev7/.. candidate directories) and RT6 is that one accepted identity. V5 (the
 * 9.3.4.60 BASE block, never had B05 at all) is UNCHANGED and keeps serving its one role unaffected by the
 * retirement: a real, historical, genuinely-unaccepted-for-B05 block -- now simply unaccepted for EVERYTHING, which
 * is the SAME outcome (PATCH_VERSION_MISMATCH) its one test already asserted.
 */

const DIR935 = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.5.0', 'euro-office-bridge-v1');
const V935 = fs.readFileSync(path.join(DIR935, 'bridge-block.js'), 'utf8');
const V6 = V935;
const V5 = fs.readFileSync(path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.4.60', 'euro-office-bridge-v1', 'bridge-block.js'), 'utf8');
const RT6 = '9.3.5|euro-office-bridge-v1/9.3.5.0/16';
const RT5 = '9.3.4|euro-office-bridge-v1/9.3.4.60/5';
const task = (spec, range = 'A1') => ({ id: 'b05c', steps: [{ intent: 'format_range', sheet: 'Data', range, format: spec }] });
const INDENT = { align: { horizontal: 'left', indent: 2 } };
const SHRINK = { align: { shrink: true } };
const BOTH = { align: { horizontal: 'left', indent: 3, shrink: true } };
const lowered = (spec, range) => format.lower(planOperations(task(spec, range)).operations[0], { sheet: 'Data' });

// ---- Part A: the product command against the REAL 9.3.5.0 block (V6 alias) --------------------------------------------

/** fake workbook `Data`: per-cell inner models (indent / shrink state), public setters and every inner touch recorded. */
function fakeEditor(block = V6, { missing = [], throws = [], perCell = {} } = {}) {
  const calls = []; const state = {};
  const innerOf = a => {
    state[a] = state[a] || { indent: 0, shrink: false };
    const s = state[a]; const miss = [...missing, ...((perCell[a] || {}).missing || [])];
    const t = { getAlign: () => ({ getReadingOrder: () => 0, getIndent: () => s.indent, getShrinkToFit: () => s.shrink }) };
    if (!miss.includes('setIndent')) t.setIndent = v => { if (throws.includes('setIndent')) throw new Error('boom'); s.indent = v; };
    if (!miss.includes('setShrinkToFit')) t.setShrinkToFit = v => { if (throws.includes('setShrinkToFit')) throw new Error('boom'); s.shrink = v; };
    return new Proxy(t, { get(o, k) { const v = o[k]; return typeof v === 'function' ? (...x) => { calls.push(['inner', a, k, ...x]); return v(...x); } : v; } });
  };
  const cellObj = a => new Proxy({ range: innerOf(a) }, { get: (o, k) => (k in o ? o[k] : (...x) => calls.push(['public', a, k, ...x])) });
  const ws = { GetName: () => 'Data', GetRange: a => cellObj(a) };
  const api = { GetSheet: n => (n === 'Data' ? ws : null), CreateColorFromRGB: () => null };
  new Function('Api', `${block.replace(/,\s*$/, '')};`)(api);
  const bridge = api.GetEuroOfficeBridge;
  const writes = [];
  api.GetEuroOfficeBridge = function () { const b = bridge.apply(this, arguments); const w = b.write; if (typeof w === 'function') b.write = (id, args) => { writes.push([id, args]); return w(id, args); }; return b; };
  return { api, calls, state, writes, inner: () => calls.filter(c => c[0] === 'inner' && /^set/.test(c[2])), publicCalls: () => calls.filter(c => c[0] === 'public') };
}
const runBody = (env, l) => new Function('Api', drv.b05BridgeWritePrelude() + l.body)(env.api);
const W = 'range.alignmentExt.write/v1';

for (const [label, spec, args] of [['indent only', INDENT, { indent: 2 }], ['shrink only', SHRINK, { shrink: true }], ['both', BOTH, { indent: 3, shrink: true }]]) {
  test(`A ${label}: ONE Bridge WRITE per cell with ONLY the requested keys; no direct setter in the product command`, () => {
    const env = fakeEditor(); const l = lowered(spec);
    assert.equal(l.meta.bridgeAlignExtWrite, true);
    assert.doesNotMatch(l.body, /setIndent|setShrinkToFit/, 'B05_DIRECT_INTERNAL in the command = 0');
    assert.deepEqual(runBody(env, l), { ok: true });
    assert.deepEqual(env.writes, [[W, { sheet: 'Data', cell: 'A1', ...args }]]);
    const want = Object.entries(args).map(([k, v]) => ['inner', 'A1', k === 'indent' ? 'setIndent' : 'setShrinkToFit', v]);
    assert.deepEqual(env.inner(), want, 'unrequested property: no setter (PRESERVE)');
  });
}

test('A multi-cell range: one write per cell, in cell order', () => {
  const env = fakeEditor();
  assert.deepEqual(runBody(env, lowered(SHRINK, 'A1:B1')), { ok: true });
  assert.deepEqual(env.writes.map(([, a]) => a.cell), ['A1', 'B1']);
});

test('A no B05 key: the command is unmarked and carries no B05 client call', () => {
  const l = lowered({ align: { horizontal: 'left' } });
  assert.equal(l.meta.bridgeAlignExtWrite, false);
  assert.doesNotMatch(l.body, /__eoB05/);
});

const editManifest = (src, fn) => {
  const m = /var EO_BRIDGE_MANIFEST = (\{.*?\], \}); function argsOk/s.exec(src);
  assert.ok(m, 'manifest literal found');
  const obj = new Function(`return ${m[1]};`)();
  fn(obj);
  return src.replace(m[1], JSON.stringify(obj));
};
for (const [label, block, reason] of [
  // DS9346: V5 is documentServerVersion 9.3.4 -- now refused one check EARLIER than before
  // (RUNTIME_VERSION_MISMATCH, classifyHandshake's own first accepted-list check), since 9.3.4 is
  // no longer an accepted documentServerVersion at all, not merely one whose /5 patch lacks B05.
  ['/5 (no WRITE capability)', V5, 'RUNTIME_VERSION_MISMATCH'],
  // DS9346: with ACCEPTED carrying ONE entry, any patchVersion other than the accepted one mismatches --
  // no "safely unclaimed number" chase needed anymore (the 9.3.4.60-era comment chain this replaced had
  // to keep moving this example forward as the accepted list grew; a single-entry list has no such edge).
  ['9.3.5.0 claiming an unaccepted revision', editManifest(V6, m => { m.patchVersion = 'euro-office-bridge-v1/9.3.5.0/999'; }), 'PATCH_VERSION_MISMATCH'],
  ['/6 without the write entry', editManifest(V6, m => { m.capabilities = m.capabilities.filter(c => c.id !== W); }), 'CAPABILITY_MISSING'],
  ['/6, write entry version 2', editManifest(V6, m => { m.capabilities.find(c => c.id === W).version = 2; }), 'CAPABILITY_VERSION_MISMATCH'],
  ['/6, write entry kind READ', editManifest(V6, m => { m.capabilities.find(c => c.id === W).kind = 'READ'; }), 'CAPABILITY_KIND_MISMATCH'],
  ['/6, write entry mutation NO', editManifest(V6, m => { m.capabilities.find(c => c.id === W).mutation = 'NO'; }), 'CAPABILITY_KIND_MISMATCH'],
]) {
  test(`A gate (${label}): refused BEFORE any mutation of the command -- ${reason}, no public setter, no write, no fallback`, () => {
    const env = fakeEditor(block);
    assert.deepEqual(runBody(env, lowered(BOTH)), { ok: false, stage: 'bridge-write-gate', reason, mutation: 'NONE' });
    assert.deepEqual([env.writes.length, env.inner().length, env.publicCalls().length], [0, 0, 0]);
  });
}

test('A gate: a bridge without write() -> CAPABILITY_MISSING, 0 mutation', () => {
  const env = fakeEditor(); const orig = env.api.GetEuroOfficeBridge;
  env.api.GetEuroOfficeBridge = () => { const b = orig(); delete b.write; return b; };
  assert.deepEqual(runBody(env, lowered(INDENT)), { ok: false, stage: 'bridge-write-gate', reason: 'CAPABILITY_MISSING', mutation: 'NONE' });
  assert.deepEqual([env.inner().length, env.publicCalls().length], [0, 0]);
});

test('A setter missing: the bridge refuses that cell, the command stops, the observed state is reported', () => {
  const env = fakeEditor(V6, { missing: ['setIndent'] });
  const r = runBody(env, lowered(INDENT));
  assert.deepEqual([r.ok, r.stage, r.cell, r.reason, r.partial, r.applied, r.cellsDone], [false, 'bridge-write', 'A1', 'BRIDGE:setIndent-missing', false, [], []]);
  assert.deepEqual(r.observed, { A1: { indent: 0, shrink: false } });
  assert.deepEqual(env.inner(), [], 'no internal setter ran');
});

test('A partial write: indent applied, shrink threw -> partial true, applied [indent], ONE write, observed indent at target', () => {
  const env = fakeEditor(V6, { throws: ['setShrinkToFit'] });
  const r = runBody(env, lowered(BOTH));
  assert.deepEqual([r.ok, r.stage, r.reason, r.partial, r.applied], [false, 'bridge-write', 'BRIDGE:setShrinkToFit-threw', true, ['indent']]);
  assert.deepEqual(r.observed, { A1: { indent: 3, shrink: false } });
  assert.equal(env.writes.length, 1, 'no second Bridge write in the same invocation');
});

test('A setter throw on the first key: partial false, nothing applied', () => {
  const env = fakeEditor(V6, { throws: ['setIndent'] });
  const r = runBody(env, lowered(BOTH));
  assert.deepEqual([r.reason, r.partial, r.applied], ['BRIDGE:setIndent-threw', false, []]);
});

test('A a later cell fails: the earlier cell stays written (no rollback), cellsDone + observed name both, no further write', () => {
  const env = fakeEditor(V6, { perCell: { B1: { missing: ['setShrinkToFit'] } } });
  const r = runBody(env, lowered(SHRINK, 'A1:C1'));
  assert.deepEqual([r.cell, r.cellsDone], ['B1', ['A1']]);
  assert.deepEqual(r.observed, { A1: { indent: 0, shrink: true }, B1: { indent: 0, shrink: false } });
  assert.deepEqual(env.writes.map(([, a]) => a.cell), ['A1', 'B1']);
});

for (const [label, reply] of [['ok without value', { ok: true }], ['applied != requested', { ok: true, value: { applied: ['indent'] } }], ['extra key', { ok: true, value: { applied: ['indent', 'shrink'] }, x: 1 }], ['not an object', 'yes'], ['refusal without reason', { ok: false }]]) {
  test(`A malformed receipt (${label}) -> UNEXPECTED_SHAPE, uncertain, controlled (no crash, no fallback)`, () => {
    const env = fakeEditor(); const orig = env.api.GetEuroOfficeBridge;
    env.api.GetEuroOfficeBridge = () => { const b = orig(); b.write = () => reply; return b; };
    const r = runBody(env, lowered(BOTH));
    assert.deepEqual([r.ok, r.stage, r.reason, r.uncertain], [false, 'bridge-write', 'UNEXPECTED_SHAPE', true]);
    assert.deepEqual(env.inner(), []);
  });
}

test('A a throwing write() -> BRIDGE_THREW, uncertain', () => {
  const env = fakeEditor(); const orig = env.api.GetEuroOfficeBridge;
  env.api.GetEuroOfficeBridge = () => { const b = orig(); b.write = () => { throw new Error('x'); }; return b; };
  const r = runBody(env, lowered(INDENT));
  assert.deepEqual([r.reason, r.uncertain], ['BRIDGE_THREW', true]);
});

test('A without the injected client the command cannot write: ReferenceError caught -> not ok, 0 mutation', () => {
  const env = fakeEditor();
  const r = new Function('Api', lowered(INDENT).body)(env.api);
  assert.deepEqual([r.ok, r.stage], [false, 'threw']);
  assert.deepEqual([env.writes.length, env.inner().length, env.publicCalls().length], [0, 0, 0]);
});

test('A driver b05DispatchResult: ok ONLY when the command completed AND its body said ok; partial -> partiallyApplied', () => {
  assert.deepEqual(drv.b05DispatchResult({ ok: true, outcome: 'command-completed', result: { ok: true } }), { ok: true, dispatched: true, outcome: 'command-completed', error: null });
  const p = drv.b05DispatchResult({ ok: true, result: { ok: false, stage: 'bridge-write', partial: true } });
  assert.deepEqual([p.ok, p.partiallyApplied, p.outcome], [true, true, 'b05-bridge-write-partial']);
  for (const r of [{ ok: true, result: { ok: false, stage: 'bridge-write', partial: false } }, { ok: true, result: { ok: false, stage: 'bridge-write-gate' } }, { ok: true, result: null }, { ok: false, outcome: 'command-callback-timeout' }]) {
    const d = drv.b05DispatchResult(r);
    assert.deepEqual([d.ok, d.dispatched, d.partiallyApplied], [false, true, undefined], JSON.stringify(r));
  }
});

test('A driver second line: a B05 command is refused WITHOUT an editor on a runtime other than /6', async () => {
  const d = drv.createEuroOfficeXlsxDriver({ url: 'u', user: 'u', pass: 'p', loadPlaywright: () => ({ ok: false }) });
  const r = await d.dispatch(null, lowered(INDENT));
  assert.deepEqual([r.ok, r.dispatched, r.outcome], [false, false, 'b05-bridge-write-runtime-not-proven']);
});

// ---- Part B: the canonical executor over the fake driver ---------------------------------------------------------------

const book = () => [{ name: 'Data', cells: { A1: { v: 'x' }, B1: { v: 'y' } }, merges: [] }];
async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, { runtime: RT6, ...script });
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}
const op = r => r.execution.appliedOperations[0];
const cell = (store, a) => store.durable.find(s => s.name === 'Data').cells[a];

for (const [label, spec, check] of [['indent', INDENT, c => c.indent === 2], ['shrink', SHRINK, c => c.shrink === true], ['both', BOTH, c => c.indent === 3 && c.shrink === true]]) {
  test(`B ${label} on /6: APPLIED (LIVE_VERIFY on the fresh READ) + FORCESAVE_CONFIRMED; retry SATISFIED with 0 command / 0 Bridge write`, async () => {
    const store = createStore(book());
    const first = await run(store, task(spec));
    assert.equal(op(first.result).status, OPERATION_STATUS.APPLIED, JSON.stringify(first.result.execution));
    assert.equal(first.result.persistenceState, 'FORCESAVE_CONFIRMED');
    assert.equal(first.driver.calls.bridgeWrites, 1);
    assert.ok(check(cell(store, 'A1')), 'stored state');
    const again = await run(store, task(spec));
    assert.equal(op(again.result).status, OPERATION_STATUS.SATISFIED);
    assert.deepEqual([again.m.commands, again.driver.calls.dispatch, again.driver.calls.bridgeWrites], [0, 0, 0]);
  });
}

// GH #326/#327: the accepted list widened again (9.3.5.0/16 + /17, compatibility-first) -- /17 moved OFF this
// refused list onto the accepted one, so this example moves forward to /18 (the next unclaimed number), the
// SAME "keep moving this example forward" edge the DS9346-era comment this replaced already named.
for (const rt of [RT5, '9.3.4|euro-office-bridge-v1/9.3.4.60/4', '9.3.5|euro-office-bridge-v1/9.3.5.0/18', 'unknown']) {
  test(`B wrong runtime ${rt}: B05 FAIL_CLOSED after the prewrite READ, before any write -- dispatch 0, Bridge write 0, mutation 0`, async () => {
    const store = createStore(book()); const before = JSON.stringify(store.durable);
    const { result, driver } = await run(store, task(BOTH), { runtime: rt });
    assert.equal(op(result).status, OPERATION_STATUS.FAILED);
    assert.equal(op(result).evidence.evidence.reason, 'xlsx-align-ext-bridge-write-runtime-not-proven');
    assert.deepEqual([driver.calls.dispatch, driver.calls.bridgeWrites], [0, 0]);
    assert.equal(JSON.stringify(store.durable), before);
  });
}

test('B DS5_POLICY=B05_ONLY_FAIL_CLOSED: on a runtime without B05, a non-B05 format key still APPLIES (capability-specific refusal, no XLSX shutdown)', async () => {
  const { result, driver } = await run(createStore(book()), task({ align: { horizontal: 'right' } }), { runtime: RT5 });
  assert.equal(op(result).status, OPERATION_STATUS.APPLIED);
  assert.equal(driver.calls.dispatch, 1);
});

test('B already holding on a runtime without B05: SATISFIED with 0 write (the gate decides only a write)', async () => {
  const b = book(); b[0].cells.A1.shrink = true;
  const { result, driver } = await run(createStore(b), task(SHRINK), { runtime: RT5 });
  assert.equal(op(result).status, OPERATION_STATUS.SATISFIED);
  assert.equal(driver.calls.dispatch, 0);
});

test('B missing capability (in-editor gate refused): DISPATCHED_OUTCOME_UNKNOWN, nothing changed, no fallback', async () => {
  const store = createStore(book()); const before = JSON.stringify(store.durable);
  const { result } = await run(store, task(BOTH), { b05Write: 'gate-refused' });
  assert.equal(op(result).status, OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN);
  assert.equal(op(result).uncertainty.dispatchEvidence.evidence.stage, 'bridge-write-gate');
  assert.equal(JSON.stringify(store.durable), before);
});

test('B partial write: DISPATCHED_OUTCOME_UNKNOWN (ADAPTER_REPORTED_PARTIALLY_APPLIED), observed state in diagnostics, no rollback; a new invocation APPLIES the rest', async () => {
  const store = createStore(book());
  const first = await run(store, task(BOTH), { b05Write: 'partial' });
  assert.equal(op(first.result).status, OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN);
  assert.equal(op(first.result).uncertainty.reason, UNCERTAINTY_REASON.ADAPTER_REPORTED_PARTIALLY_APPLIED);
  assert.deepEqual(op(first.result).uncertainty.dispatchEvidence.evidence.applied, ['indent']);
  assert.equal(first.driver.calls.bridgeWrites, 1, 'no second write in the same invocation');
  const again = await run(store, task(BOTH));
  assert.equal(op(again.result).status, OPERATION_STATUS.APPLIED);
  assert.deepEqual([cell(store, 'A1').indent, cell(store, 'A1').shrink], [3, true]);
});

for (const mode of ['error-applied', 'malformed']) {
  test(`B write error with the TARGET OBSERVED (${mode}): DISPATCHED_OUTCOME_UNKNOWN, NOT APPLIED, no forcesave claimed`, async () => {
    const first = await run(createStore(book()), task(BOTH), { b05Write: mode });
    assert.equal(op(first.result).status, OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN);
    assert.equal(first.result.ok, false);
    assert.notEqual(first.result.persistenceState, 'FORCESAVE_CONFIRMED');
    assert.equal(first.driver.calls.forceSave, 0);
  });
}

test('B retry after UNKNOWN: a NEW invocation decides from its own fresh prewrite READ -- target observed -> SATISFIED 0 write; not observed -> APPLY', async () => {
  // whether the uncertain write survived the closed session is a LIVE question (the fake keeps only force-saved state);
  // the product semantics are the same either way: it re-reads, it never trusts the earlier receipt
  const held = createStore(book());
  await run(held, task(BOTH)); // the state an uncertain write that DID survive leaves behind (target present)
  const sat = await run(held, task(BOTH));
  assert.equal(op(sat.result).status, OPERATION_STATUS.SATISFIED);
  assert.deepEqual([sat.driver.calls.dispatch, sat.driver.calls.bridgeWrites], [0, 0]);
  const store = createStore(book());
  await run(store, task(BOTH), { b05Write: 'error-applied' });
  const again = await run(store, task(BOTH));
  assert.equal(op(again.result).status, OPERATION_STATUS.APPLIED);
});

test('B setter missing (nothing written): DISPATCHED_OUTCOME_UNKNOWN, store unchanged; retry APPLIED', async () => {
  const store = createStore(book()); const before = JSON.stringify(store.durable);
  const first = await run(store, task(INDENT), { b05Write: 'error-none' });
  assert.equal(op(first.result).status, OPERATION_STATUS.DISPATCHED_OUTCOME_UNKNOWN);
  assert.equal(JSON.stringify(store.durable), before);
  const again = await run(store, task(INDENT));
  assert.equal(op(again.result).status, OPERATION_STATUS.APPLIED);
});

test('B invalid input is refused by the plan, before any read or write (unchanged)', () => {
  for (const spec of [{ align: { horizontal: 'left', indent: 16 } }, { align: { horizontal: 'left', indent: 1.5 } }, { align: { indent: 2 } }, { align: { shrink: 'yes' } }]) {
    assert.equal(planOperations(task(spec)).ok, false, JSON.stringify(spec));
  }
  assert.equal(planOperations(task(INDENT, 'A1:XFD1048576')).ok, false, 'no indent on a whole-sheet range');
});
