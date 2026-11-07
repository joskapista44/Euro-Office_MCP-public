'use strict';

/**
 * GH #300 L2: set_hyperlink / remove_hyperlink -- CANDIDATE, offline-tested only (hyperlinks-write.cjs),
 * NOT in XLSX_WRITE_ALLOWLIST (write-tools.cjs). This test drives the family through the canonical
 * executor directly (the SAME harness gh300-k2-add-comment.test.cjs uses for add_comment), bypassing
 * the write-tools.cjs product surface entirely -- the product surface's own refusal of these two intents
 * is proven separately, below, through the real validateWriteRequest.
 *
 * Engine path / signature / what remains unmeasured: docs/DS-9.3.5-BRIDGE18-L2-HYPERLINK-WRITE-PLAN.md and
 * the family module's own header comment.
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const W = require('../../src/server/write-tools.cjs');
const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const fam = require('../../src/adapters/xlsx/families/hyperlinks-write.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');
const { req } = require('../fakes/write-wiring.cjs');

const OTHER = { ref: 'D4', address: 'https://older.example/', location: null, tooltip: null, kind: 'external' };
// SURVIVOR: a SECOND pre-existing link, never the target of any step here -- the "other link" a
// touches-other fault has something to corrupt even when the step's OWN target (D4, for remove_hyperlink) is the only
// other hyperlink otherwise present.
const SURVIVOR = { ref: 'E5', address: 'https://survivor.example/', location: null, tooltip: null, kind: 'external' };
const book = () => [{ name: 'A', cells: { A1: { v: 'x' }, B2: { v: 'linked value' }, C3: { v: 7 } }, hyperlinks: [OTHER, SURVIVOR] }, { name: 'Other', cells: { A1: { v: 'y' } } }];
const setStep = (extra = {}) => ({ intent: 'set_hyperlink', sheet: 'A', cell: 'B2', address: 'https://example.com/target', ...extra });
const removeStep = (extra = {}) => ({ intent: 'remove_hyperlink', sheet: 'A', cell: 'D4', ...extra });
const task = step => ({ id: 'gh300-l2', steps: [step] });

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, script);
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}
const linkAt = (store, cell) => (store.durable[0].hyperlinks || []).find(l => l.ref === cell) || null;

test('L2 schema (family-level, pre-release): sheet + cell + address accepted; location/tooltip optional non-empty; an unknown field, a multi-cell, an empty string refused', () => {
  const planOne = s => fam.plan(s, 0);
  assert.equal(planOne(setStep()).ok, true, JSON.stringify(planOne(setStep())));
  assert.equal(planOne(setStep({ tooltip: 'see here' })).ok, true);
  assert.equal(planOne(setStep({ location: 'Sheet2!A1' })).ok, false, 'address AND location refused (GH #352: the engine drops the location)');
  assert.equal(planOne(setStep({ kind: 'external' })).ok, false, 'kind is not a step field');
  assert.equal(planOne(setStep({ cell: 'B2:C3' })).ok, false, 'one cell only');
  assert.equal(planOne(setStep({ address: '' })).ok, false, 'empty address refused, not treated as omitted');
  assert.equal(planOne(setStep({ location: '' })).ok, false, 'empty location refused (omit it instead)');
  assert.equal(planOne({ intent: 'set_hyperlink', sheet: 'A', cell: 'B2' }).ok, false, 'an address or a location is required');
  assert.equal(planOne(setStep({ address: 'x'.repeat(fam.MAX_ADDRESS + 1) })).ok, false, 'over the ceiling');
  assert.equal(planOne(setStep({ address: 'bell \u0007' })).ok, false, 'a control character');
  assert.equal(planOne(removeStep()).ok, true, JSON.stringify(planOne(removeStep())));
  assert.equal(planOne({ intent: 'remove_hyperlink', sheet: 'A', cell: 'D4', address: 'x' }).ok, false, 'remove_hyperlink takes no address');
});

test('L2 set_hyperlink: APPLIED -- the cell holds exactly the written link, the older link and every cell unchanged; retry SATISFIED 0 commands', async () => {
  const store = createStore(book());
  const { result, m } = await run(store, task(setStep()));
  assert.equal(result.ok, true, JSON.stringify([result.execution.primitiveVerification, result.execution.appliedOperations[0].evidence]));
  assert.equal(m.commands, 1);
  assert.deepEqual(linkAt(store, 'B2'), { ref: 'B2', address: 'https://example.com/target', location: null, tooltip: null, kind: 'external' });
  assert.deepEqual(linkAt(store, 'D4'), OTHER);
  assert.equal(store.durable[0].cells.B2.v, 'linked value');
  const again = await run(store, task(setStep()));
  assert.equal(again.m.commands, 0);
  assert.equal(again.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('L2 set_hyperlink: a tooltip round-trips exactly; retry with a DIFFERENT tooltip is NOT satisfied by the old link', async () => {
  const store = createStore(book());
  const { result } = await run(store, task(setStep({ tooltip: 'a tip' })));
  assert.equal(result.ok, true, JSON.stringify(result.execution));
  assert.deepEqual(linkAt(store, 'B2'), { ref: 'B2', address: 'https://example.com/target', location: null, tooltip: 'a tip', kind: 'external' });
  const diff = await run(store, task(setStep({ tooltip: 'a different tip' })));
  assert.equal(diff.m.commands, 0, 'refused before dispatch -- the cell already holds a different hyperlink');
  assert.equal(diff.result.ok, false);
  assert.match(JSON.stringify(diff.result.execution), /xlsx-hyperlink-cell-occupied/);
});

test('L2 set_hyperlink: a cell already holding a DIFFERENT hyperlink is refused before the write (SetHyperlink overwrite is unmeasured), 0 commands', async () => {
  const store = createStore(book());
  const { result, m } = await run(store, task(setStep({ cell: 'D4' })));
  assert.equal(result.ok, false);
  assert.equal(m.commands, 0);
  assert.match(JSON.stringify(result.execution), /xlsx-hyperlink-cell-occupied/);
  assert.deepEqual(linkAt(store, 'D4'), OTHER, 'the existing (different) link stays, untouched');
});

test('L2 set_hyperlink: a blank target cell is refused before the write, named (xlsx-hyperlink-cell-empty), 0 commands (GH300L2 family decision 2026-10-28, #ref)', async () => {
  const store = createStore([{ name: 'A', cells: { A1: { v: 'x' } }, hyperlinks: [] }]);
  const { result, m } = await run(store, task(setStep({ cell: 'E7' })));
  assert.equal(result.ok, false);
  assert.equal(m.commands, 0);
  assert.match(JSON.stringify(result.execution), /xlsx-hyperlink-cell-empty/);
  assert.equal(linkAt(store, 'E7'), null, 'nothing was written');
});

test('L2 set_hyperlink: a formula target cell is refused before the write, named (xlsx-hyperlink-cell-formula-unmeasured), 0 commands (GH300L2 family decision 2026-10-28, #ref)', async () => {
  const store = createStore([{ name: 'A', cells: { A1: { v: 'x' }, F8: { f: '=A1+1', v: 2 } }, hyperlinks: [] }]);
  const { result, m } = await run(store, task(setStep({ cell: 'F8' })));
  assert.equal(result.ok, false);
  assert.equal(m.commands, 0);
  assert.match(JSON.stringify(result.execution), /xlsx-hyperlink-cell-formula-unmeasured/);
  assert.equal(linkAt(store, 'F8'), null, 'nothing was written');
});

test('L2 set_hyperlink: an unobservable hyperlink read fails closed before the write, 0 commands', async () => {
  const store = createStore(book());
  const { result, m } = await run(store, task(setStep()), { hyperlinksUnobservable: true });
  assert.equal(result.ok, false);
  assert.equal(m.commands, 0);
  assert.match(JSON.stringify(result.execution), /xlsx-hyperlinks-unobservable/);
});

for (const fault of ['hyperlink-wrong-address', 'hyperlink-touches-other', 'ignore', 'side-effect']) {
  test(`L2 set_hyperlink: a fake write with fault ${fault} fails the verify, not passes it`, async () => {
    const store = createStore(book());
    const { result, m } = await run(store, task(setStep()), { fault });
    assert.equal(m.commands, 1, 'the write was dispatched -- this is the VERIFY failing, not a refusal');
    assert.equal(result.ok, false);
  });
}

test('L2 remove_hyperlink: APPLIED -- the cell loses its link, every OTHER link and every cell unchanged; retry SATISFIED 0 commands (nothing to remove)', async () => {
  const store = createStore(book());
  const { result, m } = await run(store, task(removeStep()));
  assert.equal(result.ok, true, JSON.stringify([result.execution.primitiveVerification, result.execution.appliedOperations[0].evidence]));
  assert.equal(m.commands, 1);
  assert.equal(linkAt(store, 'D4'), null);
  assert.equal(store.durable[0].cells.D4, undefined, 'removing a hyperlink never touches the cell (there was no D4 cell value to begin with, and none appears)');
  const again = await run(store, task(removeStep()));
  assert.equal(again.m.commands, 0);
  assert.equal(again.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('L2 remove_hyperlink: a cell with NO hyperlink is already SATISFIED, 0 commands (idempotent target state)', async () => {
  const store = createStore(book());
  const { result, m } = await run(store, task(removeStep({ cell: 'C3' })));
  assert.equal(result.ok, true, JSON.stringify(result.execution));
  assert.equal(m.commands, 0);
  assert.equal(result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('L2 remove_hyperlink: an unobservable hyperlink read fails closed before the write, 0 commands', async () => {
  const store = createStore(book());
  const { result, m } = await run(store, task(removeStep()), { hyperlinksUnobservable: true });
  assert.equal(result.ok, false);
  assert.equal(m.commands, 0);
  assert.match(JSON.stringify(result.execution), /xlsx-hyperlinks-unobservable/);
});

for (const fault of ['hyperlink-touches-other', 'ignore', 'side-effect']) {
  test(`L2 remove_hyperlink: a fake write with fault ${fault} fails the verify, not passes it`, async () => {
    const store = createStore(book());
    const { result, m } = await run(store, task(removeStep()), { fault });
    assert.equal(m.commands, 1, 'the write was dispatched -- this is the VERIFY failing, not a refusal');
    assert.equal(result.ok, false);
  });
}

test('L2: set_hyperlink / remove_hyperlink must each be the only step of their task; liveScope asks only the hyperlinkSheets', () => {
  const p = planOperations({ id: 't', steps: [setStep(), setStep({ cell: 'C3' })] });
  assert.equal(p.ok, false);
  assert.match(p.detail, /only step of its task/);
  const p2 = planOperations({ id: 't', steps: [removeStep(), removeStep({ cell: 'A1' })] });
  assert.equal(p2.ok, false);
  assert.match(p2.detail, /only step of its task/);
  const s = liveScope(task(setStep()));
  assert.deepEqual(s.hyperlinkSheets, ['A']);
});

test('L2: the lowered set_hyperlink command names the public ApiWorksheet.SetHyperlink, sTextToDisplay is the resolved target\'s OWN text (never \'\', GH300L2 family decision 2026-10-28); remove_hyperlink calls ApiRange.ClearHyperlinks -- NEITHER names a Bridge call', () => {
  const setPlan = fam.plan(setStep({ tooltip: 'tip' }), 0);
  assert.equal(setPlan.ok, true, setPlan.detail);
  const setCmd = fam.lower(setPlan.operation, { sheet: 'A', cell: 'B2', text: 'linked value' });
  assert.equal(setCmd.kind, 'callCommand');
  assert.match(setCmd.body, /ws\.SetHyperlink\("B2", "https:\/\/example\.com\/target", "", "tip", "linked value"\)/);
  assert.doesNotMatch(setCmd.body, /GetEuroOfficeBridge|bridgeAxis/);
  assert.throws(() => fam.lower(setPlan.operation), /requires a resolved target with text/, 'no silent \'\' fallback');
  const removePlan = fam.plan(removeStep(), 0);
  assert.equal(removePlan.ok, true, removePlan.detail);
  const removeCmd = fam.lower(removePlan.operation);
  assert.match(removeCmd.body, /ws\.GetRange\("D4"\)\.ClearHyperlinks\(\)/);
  assert.doesNotMatch(removeCmd.body, /GetEuroOfficeBridge|bridgeAxis/);
});

test('L2 RELEASED (release-flip, after the LIVE probe/acceptance): validateWriteRequest accepts both intents on the product surface', () => {
  const v = step => W.validateWriteRequest(req([step]));
  const a = v(setStep());
  assert.equal(a.ok, true, a.reason || a.detail);
  const b = v(removeStep());
  assert.equal(b.ok, true, b.reason || b.detail);
  assert.ok(W.XLSX_WRITE_ALLOWLIST.includes('set_hyperlink') && W.XLSX_WRITE_ALLOWLIST.includes('remove_hyperlink'), 'released, no longer withheld');
});

test('L2 RELEASED: the schema itself still refuses the untested shapes, named -- a multi-cell range (CELL pattern); GH #352: a location-only (internal) write passes the schema', () => {
  const v = step => W.validateWriteRequest(req([step]));
  const multi = v({ ...setStep(), cell: 'A1:B2' });
  assert.equal(multi.ok, false);
  const internal = v({ intent: 'set_hyperlink', id: 's0', sheet: 'Sheet1', cell: 'E5', location: 'Sheet1!A1' });
  assert.equal(internal.ok, true, internal.reason || internal.detail);
});

test('L2 RELEASED: the address is an absolute http(s) URL or a mailto: address -- every other scheme is refused pre-dispatch (the docx add_link rule)', () => {
  const v = address => W.validateWriteRequest(req([{ ...setStep(), address }]));
  for (const ok of ['https://example.com/a', 'http://example.com', 'mailto:someone@example.com']) assert.equal(v(ok).ok, true, ok);
  for (const bad of ['javascript:alert(1)', 'file://server/share/x.xlsx', 'file:///etc/passwd', 'data:text/html,x', 'example.com', '/relative/path', 'HTTPS://EXAMPLE.COM', 'https://exa mple.com', 'https://example.com/\u0001']) {
    assert.equal(v(bad).ok, false, `must refuse: ${bad}`);
  }
});
