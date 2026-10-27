'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const print = require('../../src/adapters/xlsx/families/print.cjs');
const { readQuery } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { installBridge, withPatch } = require('../fakes/ds-bridge-runtime.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');
const { runViaBridge } = require('../fakes/print-bridge-runtime.cjs');

/** M6 wave 8: the print family on the sheet model (W8b-proven), on the fake. */

const book = () => [{ name: 'P', cells: { A1: { v: 1 } } }, { name: 'O', cells: { A1: { v: 'keep' } }, print: { hCenter: true } }];
const one = step => ({ id: 'm6-w8r', steps: [{ sheet: 'P', ...step }] });
const STEPS = {
  fit: { intent: 'set_print_fit', fitToWidth: 1, fitToHeight: 1 },
  header: { intent: 'set_header_footer', slot: 'oddHeader', value: '&CM6' },
  rowBreak: { intent: 'set_page_break', axis: 'row', at: 11 },
  firstPage: { intent: 'set_first_page_number', value: 3 },
  center: { intent: 'set_print_centering', horizontal: true },
};

async function run(store, t, script = {}) {
  let n = 0;
  // M11.5b.4 (DS9346: now 9.3.5.0/16, the sole accepted runtime): the print WRITE goes through the
  // Bridge WRITE -- the fake session runs on the accepted runtime unless a test says otherwise
  const driver = createFakeXlsxDriver(store, { runtime: '9.3.5|euro-office-bridge-v1/9.3.5.0/16', ...script });
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, driver, m: result.metrics };
}
const reason = result => result.execution.appliedOperations[0].evidence.evidence.reason;

for (const [k, step] of Object.entries(STEPS)) {
  test(`${k}: APPLIED (L1 + L3), the retry SATISFIED 0/0`, async () => {
    const store = createStore(book());
    const a = await run(store, one(step));
    assert.equal(a.result.ok, true, JSON.stringify([a.result.execution.primitiveVerification, a.result.execution.preservationVerification]));
    assert.equal(a.m.commands, 1);
    const r = await run(store, one(step));
    assert.equal(r.m.commands, 0);
    assert.equal(r.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
  });
}

test('fit: the numbers without the fitToPage FLAG are NOT satisfied (audit PP5) -> L1 FAIL', async () => {
  const { result } = await run(createStore(book()), one(STEPS.fit), { fault: 'fit-no-flag' });
  assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL');
  // and a sheet that ALREADY has 1x1 with the flag off is APPLY, not SATISFIED
  const b = book(); b[0].print = { fitW: 1, fitH: 1, fitFlag: false };
  const a = await run(createStore(b), one(STEPS.fit));
  assert.equal(a.m.commands, 1);
});

test('RED: ignored / a wrong value -> L1 FAIL (every intent); another print field / another sheet / a cell -> L3 FAIL', async () => {
  for (const [k, step] of Object.entries(STEPS)) {
    for (const fault of ['ignore', 'print-wrong-value']) {
      const { result, driver } = await run(createStore(book()), one(step), { fault });
      assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL', `${k} ${fault}`);
      assert.equal(driver.calls.forceSave, 0, `${k} ${fault}`);
    }
    for (const fault of ['print-side', 'side-effect']) {
      const { result } = await run(createStore(book()), one(step), { fault });
      assert.equal(result.execution.preservationVerification[0].verdict, 'FAIL', `${k} ${fault}`);
    }
  }
  for (const [k, fault] of [['fit', 'fit-wrong-height'], ['rowBreak', 'break-not-manual'], ['firstPage', 'first-page-unused']]) {
    const { result } = await run(createStore(book()), one(STEPS[k]), { fault });
    assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL', `${k} ${fault}`);
  }
  const b = book(); b[0].print = { rowBreaks: [{ id: 4, man: true }] };
  const drop = await run(createStore(b), one(STEPS.rowBreak), { fault: 'break-drops-other' });
  assert.equal(drop.result.execution.preservationVerification[0].verdict, 'FAIL', 'an add may not remove an existing break');
});

test('refused BEFORE dispatch: an unproven build', async () => {
  const { result, driver } = await run(createStore(book()), one(STEPS.center), { printBuild: false });
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.equal(reason(result), 'xlsx-print-unobservable');
});

test('RED (owner decision 2026-09-30): the WRITE is gated on its own full build list -- a readable sheet on another build is refused BEFORE dispatch', async () => {
  for (const step of Object.values(STEPS)) {
    const { result, driver } = await run(createStore(book()), one(step), { runtime: '9.3.5-0000000000000000000000000000000000' });
    assert.equal(result.ok, false, step.intent);
    assert.equal(driver.calls.dispatch, 0, step.intent);
    assert.equal(reason(result), 'xlsx-print-bridge-write-runtime-not-proven', step.intent);
  }
  // a setting that already holds needs no write: SATISFIED 0/0 is not a write, the read gate alone decides it
  const b = book(); b[0].print = { hCenter: true };
  const s = await run(createStore(b), one(STEPS.center), { runtime: '9.3.5-0000000000000000000000000000000000' });
  assert.equal(s.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
  assert.equal(s.driver.calls.dispatch, 0);
});

test('RED (owner decision 2026-09-30): a missing required API refuses BEFORE any setter or transaction call (M11.5b.4: inside the /7 Bridge WRITE)', () => {
  const full = () => {
    const log = [];
    const ps = { asc_setFitToWidth: x => log.push(['fitW', x]), asc_setFitToHeight: x => log.push(['fitH', x]), asc_setFirstPageNumber: x => log.push(['first', x]) };
    const m = { PagePrintOptions: { asc_getPageSetup: () => ps, asc_setHorizontalCentered: x => log.push(['hc', x]) }, setFitToPage: (w, h) => log.push(['flag', w, h]),
      headerFooter: { setOddHeader: x => log.push(['oh', x]) }, changeRowColBreaks: () => log.push(['brk']), getPrintAreaRangeByRowCol: () => 'r' };
    const H = { Create_NewPoint: () => log.push(['np']), StartTransaction: () => log.push(['st']), EndTransaction: () => log.push(['et']) };
    return { log, ps, m, H };
  };
  const cases = [
    ['fit', x => { delete x.ps.asc_setFitToHeight; }], ['fit', x => { delete x.m.setFitToPage; }],
    ['header', x => { delete x.m.headerFooter.setOddHeader; }], ['header', x => { delete x.H.EndTransaction; }],
    ['rowBreak', x => { delete x.m.getPrintAreaRangeByRowCol; }], ['firstPage', x => { delete x.ps.asc_setFirstPageNumber; }],
    ['center', x => { delete x.m.PagePrintOptions.asc_setHorizontalCentered; }],
  ];
  for (const [k, cut] of cases) {
    const x = full(); cut(x);
    const r = runViaBridge(planOperations(one(STEPS[k])).operations[0], { GetName: () => 'P', worksheet: x.m }, { History: x.H });
    assert.equal(r.ok, false, k);
    assert.equal(r.stage, 'bridge-write', `${k} ${JSON.stringify(r)}`);
    assert.match(r.reason, /^BRIDGE:(.+-missing|history-missing)$/, `${k} ${r.reason}`);
    assert.deepEqual(x.log, [], `${k}: nothing was called`);
  }
});

test('the print WRITE runtime list (M11.5b.4 Bridge WRITE; DS9346: now the single accepted 9.3.5.0/16 entry): the adapter and the driver gate on the SAME id list; the direct-internal list is gone; the driver refuses any other runtime untouched', async () => {
  const drv = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
  assert.deepEqual([...print.B10_BRIDGE_WRITE_RUNTIMES], [...drv.B10_BRIDGE_WRITE_RUNTIMES]);
  for (const k of ['EO_INTERNAL_PRINT_RUNTIMES', 'EO_INTERNAL_PRINT_WRITE_RUNTIMES']) { assert.equal(k in drv, false, k); assert.equal(k in print, false, k); }
  assert.deepEqual([...drv.B10_BRIDGE_WRITE_RUNTIMES], ['9.3.5|euro-office-bridge-v1/9.3.5.0/16', '9.3.5|euro-office-bridge-v1/9.3.5.0/17'], 'the explicit manifest runtime id, never a URL cache key');
  const d = drv.createEuroOfficeXlsxDriver({ url: 'u', user: 'u', pass: 'p', loadPlaywright: () => ({ ok: false }) });
  const r = await d.dispatch(null, print.lower(planOperations(one(STEPS.center)).operations[0]));
  assert.equal(r.ok, false);
  assert.equal(r.dispatched, false, 'no editor was touched');
  assert.equal(r.outcome, 'b10-bridge-write-runtime-not-proven');
});

test('planner: only the measured shapes', () => {
  const refuse = (t, re) => { const p = planOperations(t); assert.equal(p.ok, false, JSON.stringify(t)); assert.match(p.detail, re); };
  // X-P (m7-s28-fit-off-probe MEASURED): fit-to-page OFF (0/0) now plans, same as any
  // other fit value -- it is no longer refused.
  const fitOff = planOperations(one({ ...STEPS.fit, fitToWidth: 0, fitToHeight: 0 }));
  assert.equal(fitOff.ok, true, JSON.stringify(fitOff));
  assert.deepEqual(fitOff.operations[0].want, { fitW: 0, fitH: 0 });
  // M7 F2 (owner GO 2026-10-03, 31278 point 3) widened the measured set: footer, clears, column / removed breaks, vertical
  // centering, horizontal false (xlsx-m7-f2-print.test.cjs); what F2 did NOT measure stays refused
  // requirement change (X-P07, owner OD-A, owner channel): the first / even slots are admitted now -- an unknown
  // slot stays refused
  refuse(one({ ...STEPS.header, slot: 'centerHeader' }), /slot must be one of oddHeader, oddFooter, firstHeader/);
  refuse(one({ ...STEPS.header, value: '' }), /non-empty code, or null/);
  refuse(one({ ...STEPS.rowBreak, axis: 'diagonal' }), /axis must be 'row' or 'column'/);
  refuse(one({ ...STEPS.rowBreak, at: 1 }), /at must be a row number >= 2/);
  refuse(one({ ...STEPS.rowBreak, axis: 'column', at: 'A' }), /column letter after A/);
  refuse(one({ ...STEPS.rowBreak, remove: false }), /remove must be true/);
  refuse(one({ ...STEPS.firstPage, value: 0 }), /positive integer, or null/);
  refuse(one({ ...STEPS.center, horizontal: 'yes' }), /horizontal must be true or false/);
  // X-P: vertical false now plans, same as vertical true -- it is no longer refused.
  const vOff = planOperations(one({ intent: 'set_print_centering', vertical: false }));
  assert.equal(vOff.ok, true, JSON.stringify(vOff));
  assert.deepEqual(vOff.operations[0].want, { center: 'vCenter', value: false });
  refuse(one({ intent: 'set_print_centering', vertical: 'yes' }), /vertical must be true or false/);
  refuse(one({ ...STEPS.center, vertical: true }), /exactly one of horizontal \/ vertical/);
  // M7 X-P03 (m7-p57-print-titles-columns-probe MEASURED): column titles are now supported
  // too (names family, xlsx-m7-f-printnames.test.cjs) -- row titles are the names family's own concern.
  const two = { id: 'x', steps: [{ sheet: 'P', ...STEPS.fit }, { sheet: 'P', ...STEPS.fit, fitToHeight: 2 }] };
  refuse(two, /on print:P:fit$/);
  assert.equal(planOperations({ id: 'x', steps: [{ sheet: 'P', ...STEPS.fit }, { sheet: 'P', ...STEPS.header }] }).ok, true, 'two axes of one sheet compose');
});

test('the command bodies: the model setters, now made by sheet.printSettings.write/v1 (M11.5b.4) -- the SAME calls, in order', () => {
  const log = [];
  const ps = { asc_setFitToWidth: x => log.push(['fitW', x]), asc_setFitToHeight: x => log.push(['fitH', x]), asc_setFirstPageNumber: x => log.push(['first', x]) };
  const m = { PagePrintOptions: { asc_getPageSetup: () => ps, asc_setHorizontalCentered: x => log.push(['hc', x]) }, setFitToPage: (...a) => log.push(['flag', ...a]),
    headerFooter: { setOddHeader: x => log.push(['oh', x]) }, changeRowColBreaks: (a, b, c, d, e) => log.push(['brk', b, c, d, e]), getPrintAreaRangeByRowCol: r => `range@${r}` };
  const ws = { GetName: () => 'P', worksheet: m };
  const H = { Create_NewPoint: () => log.push(['np']), StartTransaction: () => log.push(['st']), EndTransaction: () => log.push(['et']) };
  for (const k of Object.keys(STEPS)) {
    const op = planOperations(one(STEPS[k])).operations[0];
    assert.doesNotMatch(print.lower(op).body, /PagePrintOptions|asc_set|setFitToPage|changeRowColBreaks|\.headerFooter|AscCommon|\.worksheet/, `${k}: no direct internal in the product command`);
    assert.deepEqual(runViaBridge(op, ws, { History: H }), { ok: true }, k);
  }
  assert.deepEqual(log, [['fitW', 1], ['fitH', 1], ['flag', true], ['np'], ['st'], ['oh', '&CM6'], ['et'], ['brk', 10, 'range@10', false, true], ['first', 3], ['hc', true]]);
});

test('RED (owner review 8527132): a throwing setOddHeader still closes the started transaction; the failure is kept, not reported as success (inside the /7 Bridge WRITE)', () => {
  const log = [];
  const m = { PagePrintOptions: { asc_getPageSetup: () => ({}) }, headerFooter: { setOddHeader: () => { log.push(['oh']); throw new Error('boom'); } } };
  const ws = { GetName: () => 'P', worksheet: m };
  const H = { Create_NewPoint: () => log.push(['np']), StartTransaction: () => log.push(['st']), EndTransaction: () => log.push(['et']) };
  const r = runViaBridge(planOperations(one(STEPS.header)).operations[0], ws, { History: H });
  assert.deepEqual(log, [['np'], ['st'], ['oh'], ['et']], 'EndTransaction after the throw');
  assert.equal(r.ok, false);
  assert.equal(r.stage, 'bridge-write');
  assert.equal(r.reason, 'BRIDGE:setOddHeader-threw');
  assert.deepEqual([r.partial, r.applied], [false, []]);
});

test('the driver reader (B09 through sheet.printSettings.read/v1): measured defaults (no sheetPr, null breaks, null centering) read as off / none; another patch is unobservable; breaks under the M7 S2a contract', () => {
  const ps = { asc_getFitToWidth: () => 0, asc_getFitToHeight: () => 0, asc_getScale: () => 100, asc_getWidth: () => 210, asc_getHeight: () => 297, useFirstPageNumber: false, firstPageNumber: null };
  const hf = odd => Object.fromEntries(['OddHeader', 'OddFooter', 'FirstHeader', 'FirstFooter', 'EvenHeader', 'EvenFooter', 'DifferentFirst', 'DifferentOddEven'].map(k => [`get${k}`, () => (k === 'OddHeader' ? odd : null)]));
  const model = over => ({ PagePrintOptions: { asc_getPageSetup: () => ps, asc_getHorizontalCentered: () => null, asc_getVerticalCentered: () => null }, headerFooter: hf(null), rowBreaks: null, sheetPr: null, ...over });
  const mk = (name, m) => ({ GetName: () => name, GetIndex: () => 0, GetUsedRange: () => null, worksheet: m });
  const brk = list => ({ getBreaks: () => list });
  const sheets = [mk('D', model({})), mk('F', model({ sheetPr: { FitToPage: 1 }, rowBreaks: brk([{ id: 10, man: true, min: null, max: 16383, pt: null }]), headerFooter: hf({ str: '&CW8' }) })), mk('X', model({ rowBreaks: {} })),
    mk('A', (() => { const m = model({}); delete m.rowBreaks; return m; })()), mk('L', model({ rowBreaks: brk(null) })),
    ...[{ id: '3', man: true }, { id: -1, man: true }, { id: 1048576, man: true }, { id: 3 }, { id: 3, man: 1 }, { id: 3, man: true, max: 16384 }, { id: 3, man: true, min: 5, max: 4 }, null]
      .map((b, i) => mk(`E${i}`, model({ rowBreaks: brk([b]) })))];
  const Api = { GetSheets: () => sheets, GetSheet: n => sheets.find(s => s.GetName() === n), WorksheetFunction: { COUNT: () => 0 } };
  installBridge(Api);
  const r = new Function('Api', readQuery([], { printSheets: ['*'] }))(Api).printReads;
  assert.deepEqual(r.D, { fitW: 0, fitH: 0, fitFlag: false, oddHeader: null, rowBreaks: [], useFirstPage: false, firstPage: null, hCenter: false, scale: 100, width: 210, height: 297 });
  assert.equal(r.F.fitFlag, true);
  assert.deepEqual(r.F.rowBreaks, [{ id: 10, man: true, min: null, max: 16383 }], 'the measured element, pt not read');
  assert.equal(r.F.oddHeader, '&CW8');
  assert.equal(r.X, 'UNOBSERVABLE:BRIDGE:rowBreaks-unexpected-shape');
  assert.equal(r.A, 'UNOBSERVABLE:BRIDGE:rowBreaks-absent', 'an absent key is never []');
  assert.equal(r.L, 'UNOBSERVABLE:BRIDGE:rowBreaks-unexpected-shape', 'getBreaks() not a list');
  for (let i = 0; i < 8; i += 1) assert.equal(r[`E${i}`], 'UNOBSERVABLE:BRIDGE:rowBreaks-element-invalid', `E${i}`);
  assert.equal(r.D.colBreaks, undefined, 'the column breaks are read only when asked (the structure family)');
  assert.equal(new Function('Api', readQuery([], { printSheets: ['*'] }))(withPatch(Api)).printReads.D, 'UNOBSERVABLE:PATCH_VERSION_MISMATCH', 'a /3 bridge refuses, no fallback');
  assert.equal(new Function('Api', readQuery([], { printSheets: ['*'] }))({ ...Api, GetEuroOfficeBridge: undefined }).printReads.D, 'UNOBSERVABLE:BRIDGE_UNAVAILABLE');
  assert.equal(new Function('Api', readQuery([], { printSheets: ['Nope'] }))(Api).printReads.Nope, 'UNOBSERVABLE:sheet-missing', 'the driver pre-check, before any bridge call');
});

// M11.5b.4 cutover: B10 print WRITE through the Bridge WRITE ONLY. DS9346 (owner D5=B, VALTAS,
// docs/DS9346-RUNTIME-RETIREMENT.md): the 9.3.4.60 family (/7 through /16, the entire list GH #138
// U4 widened) is retired -- ONE accepted entry now (9.3.5.0/16). /4 /5 /6 era examples (never
// carried this capability) stay refused before the write (no fallback) in the loop below.
for (const rt of ['9.3.5|euro-office-bridge-v1/9.3.5.0/16']) test(`M11.5b.4 print WRITE on ${rt}: every intent APPLIED via ONE Bridge WRITE, retry SATISFIED 0/0`, async () => {
  for (const [k, step] of Object.entries(STEPS)) {
    const store = createStore(book());
    const a = await run(store, one(step), { runtime: rt });
    assert.equal(a.result.ok, true, k); assert.equal(a.m.commands, 1, k); assert.equal(a.driver.calls.bridgeWrites, 1, k);
    const r = await run(store, one(step), { runtime: rt });
    assert.equal(r.m.commands, 0, k);
  }
});
for (const rt of ['4', '5', '6'].map(n => `9.3.4|euro-office-bridge-v1/9.3.4.60/${n}`)) test(`M11.5b.4 print WRITE on ${rt}: refused before the write`, async () => {
  for (const step of Object.values(STEPS)) {
    const { result, driver } = await run(createStore(book()), one(step), { runtime: rt });
    assert.equal(reason(result), 'xlsx-print-bridge-write-runtime-not-proven'); assert.deepEqual([driver.calls.dispatch, driver.calls.bridgeWrites], [0, 0]);
  }
});
