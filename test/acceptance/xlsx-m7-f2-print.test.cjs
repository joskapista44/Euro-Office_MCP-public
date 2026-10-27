'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const print = require('../../src/adapters/xlsx/families/print.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');
const { runViaBridge } = require('../fakes/print-bridge-runtime.cjs');

/**
 * M7 tier F2 (owner GO 2026-10-03, owner channel, point 3): the print WRITE widened to the shapes the F2 probe
 * measured (docs/evidence/M7-F2-PROBE.md) -- Q9 odd footer set / clear and odd header clear, Q11 a column break add,
 * Q12 a row / column break removal, Q14 the first page number clear (the useFirstPage flag stays), Q15 vertical
 * centering true and horizontal false. Every step: APPLIED (L1 + L3) then SATISFIED 0/0; a removal keeps the others.
 */

const book = () => [
  { name: 'P', cells: { A1: { v: 1 } }, print: { oddHeader: '&Chead0', oddFooter: '&Cfoot0', rowBreaks: [{ id: 10, man: true }, { id: 20, man: true }], colBreaks: [{ id: 4, man: true }, { id: 7, man: true }], useFirstPage: true, firstPage: 3, hCenter: true } },
  { name: 'O', cells: { A1: { v: 'keep' } }, print: { hCenter: true, oddFooter: '&Cother' } },
];
const one = step => ({ id: 'm7-f2', steps: [{ sheet: 'P', ...step }] });
const STEPS = {
  footerSet: { intent: 'set_header_footer', slot: 'oddFooter', value: '&Cfoot' },
  footerClear: { intent: 'set_header_footer', slot: 'oddFooter', value: null },
  headerClear: { intent: 'set_header_footer', slot: 'oddHeader', value: null },
  colBreakAdd: { intent: 'set_page_break', axis: 'column', at: 'J' },
  rowBreakRemove: { intent: 'set_page_break', axis: 'row', at: 11, remove: true },
  colBreakRemove: { intent: 'set_page_break', axis: 'column', at: 'E', remove: true },
  firstPageClear: { intent: 'set_first_page_number', value: null },
  vertical: { intent: 'set_print_centering', vertical: true },
  horizontalOff: { intent: 'set_print_centering', horizontal: false },
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
const sheetP = store => store.durable.find(s => s.name === 'P').print;

for (const [k, step] of Object.entries(STEPS)) {
  test(`${k}: APPLIED (L1 + L3), the retry SATISFIED 0/0`, async () => {
    const store = createStore(book());
    const a = await run(store, one(step));
    assert.equal(a.result.ok, true, JSON.stringify([a.result.execution.primitiveVerification, a.result.execution.preservationVerification, a.result.execution.appliedOperations]));
    assert.equal(a.m.commands, 1);
    const r = await run(store, one(step));
    assert.equal(r.m.commands, 0, k);
    assert.equal(r.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED, k);
  });
}

test('the stored effect: a removal takes exactly its break; a first-page clear keeps the useFirstPage flag', async () => {
  const s1 = createStore(book());
  await run(s1, one(STEPS.rowBreakRemove));
  assert.deepEqual(sheetP(s1).rowBreaks, [{ id: 20, man: true }]);
  assert.deepEqual(sheetP(s1).colBreaks, [{ id: 4, man: true }, { id: 7, man: true }]);
  const s2 = createStore(book());
  await run(s2, one(STEPS.colBreakAdd));
  assert.deepEqual(sheetP(s2).colBreaks.map(b => b.id), [4, 7, 9]);
  const s3 = createStore(book());
  await run(s3, one(STEPS.firstPageClear));
  assert.equal(sheetP(s3).firstPage, null);
  assert.equal(sheetP(s3).useFirstPage, true);
  // a removal of a break that is not there is SATISFIED 0/0 (the requested state already holds)
  const gone = await run(createStore(book()), one({ ...STEPS.rowBreakRemove, at: 31 }));
  assert.equal(gone.m.commands, 0);
  assert.equal(gone.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('RED: ignored / a wrong value -> L1 FAIL; another print field / another sheet / a cell -> L3 FAIL', async () => {
  for (const [k, step] of Object.entries(STEPS)) {
    for (const fault of ['ignore', 'print-wrong-value']) {
      if (fault === 'print-wrong-value' && step.intent === 'set_page_break') continue;
      const { result, driver } = await run(createStore(book()), one(step), { fault });
      assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL', `${k} ${fault}`);
      assert.equal(driver.calls.forceSave, 0, `${k} ${fault}`);
    }
    for (const fault of ['print-side', 'side-effect']) {
      const { result } = await run(createStore(book()), one(step), { fault });
      assert.equal(result.execution.preservationVerification[0].verdict, 'FAIL', `${k} ${fault}`);
    }
  }
  for (const k of ['rowBreakRemove', 'colBreakRemove', 'colBreakAdd']) {
    for (const fault of ['break-drops-other', 'break-adds-other', 'break-moves-other']) {
      const { result } = await run(createStore(book()), one(STEPS[k]), { fault });
      assert.equal(result.ok, false, `${k} ${fault}`);
      assert.equal(result.execution.preservationVerification[0].verdict, 'FAIL', `${k} ${fault}: exactly the requested break changes`);
    }
  }
  const ignored = await run(createStore(book()), one(STEPS.rowBreakRemove), { fault: 'break-remove-ignored' });
  assert.equal(ignored.result.execution.primitiveVerification[0].verdict, 'FAIL');
  const flag = await run(createStore(book()), one(STEPS.firstPageClear), { fault: 'first-page-flag' });
  assert.equal(flag.result.execution.preservationVerification[0].verdict, 'FAIL', 'the clear may not touch useFirstPage');
});

test('refused BEFORE dispatch: an unwidened read, an automatic break, an unproven build', async () => {
  for (const k of ['footerSet', 'colBreakAdd', 'vertical']) {
    const { result, driver } = await run(createStore(book()), one(STEPS[k]), { printWide: false });
    assert.equal(result.ok, false, k);
    assert.equal(driver.calls.dispatch, 0, k);
    assert.match(JSON.stringify(result), /widened-fields-not-read/, k);
  }
  // an unwidened read does not block a step that needs no widened field
  const plain = await run(createStore(book()), one({ intent: 'set_header_footer', slot: 'oddHeader', value: '&Cx' }), { printWide: false });
  assert.equal(plain.result.ok, true);
  const auto = book();
  auto[0].print.rowBreaks = [{ id: 10, man: false }];
  const a = await run(createStore(auto), one(STEPS.rowBreakRemove));
  assert.equal(a.result.ok, false);
  assert.equal(a.driver.calls.dispatch, 0);
  assert.match(JSON.stringify(a.result), /xlsx-print-break-not-manual/);
  const b = await run(createStore(book()), one(STEPS.vertical), { printBuild: false });
  assert.equal(b.result.ok, false);
  assert.equal(b.driver.calls.dispatch, 0);
});

// M11.5b.4: the F2-measured call shapes are now made by the /7 Bridge WRITE -- the product command runs through it here
test('the command bodies: the F2-measured call shapes (through sheet.printSettings.write/v1)', () => {
  const log = [];
  const ps = { asc_setFirstPageNumber: x => log.push(['first', x]) };
  const m = { PagePrintOptions: { asc_getPageSetup: () => ps, asc_setHorizontalCentered: x => log.push(['hc', x]), asc_setVerticalCentered: x => log.push(['vc', x]) },
    headerFooter: { setOddHeader: x => log.push(['oh', x]), setOddFooter: x => log.push(['of', x]) }, changeRowColBreaks: (a, b, c, d, e) => log.push(['brk', a, b, c, d, e]), getPrintAreaRangeByRowCol: (r, c) => `range@${r},${c}` };
  const ws = { GetName: () => 'P', worksheet: m };
  const H = { Create_NewPoint: () => {}, StartTransaction: () => {}, EndTransaction: () => {} };
  for (const k of Object.keys(STEPS)) assert.deepEqual(runViaBridge(planOperations(one(STEPS[k])).operations[0], ws, { History: H }), { ok: true }, k);
  assert.deepEqual(log, [['of', '&Cfoot'], ['of', null], ['oh', null], ['brk', null, 9, 'range@null,9', true, true], ['brk', 10, null, 'range@10,null', false, true],
    ['brk', 4, null, 'range@null,4', true, true], ['first', null], ['vc', true], ['hc', false]]);
});

/**
 * X-P (m7-s28-fit-off-probe MEASURED): fit-to-page OFF (0/0) and vertical
 * centering false. Both widen a shape F2 already had (set_print_fit / set_print_centering
 * vertical), so the generic STEPS loop above does not cover them -- on the loop's own book(),
 * sheet 'P' never sets fitW/fitH/fitFlag/vCenter, so PRINT_DEFAULT already reads 0/0/false/false:
 * the OFF shapes would be SATISFIED on their very FIRST call, not APPLIED-then-retry like every
 * other STEPS entry. Tested here on their own: a genuine APPLY from an already-ON, seeded sheet
 * (proving the write path fires), AND the already-satisfied-by-default case on the untouched
 * sheet 'P' (proving the shortcut is correct, not an accident of the fixture).
 */
const fitOn = () => { const b = book(); b[0].print = { ...b[0].print, fitW: 50, fitH: 60, fitFlag: true }; return b; };
const vOn = () => { const b = book(); b[0].print = { ...b[0].print, vCenter: true }; return b; };
const STEPS_P = {
  fitOff: { intent: 'set_print_fit', fitToWidth: 0, fitToHeight: 0 },
  verticalOff: { intent: 'set_print_centering', vertical: false },
};

test('fitOff: APPLIED from an already-ON sheet (the flag flips to off), the retry SATISFIED 0/0', async () => {
  const store = createStore(fitOn());
  const a = await run(store, one(STEPS_P.fitOff));
  assert.equal(a.result.ok, true, JSON.stringify(a.result.execution.primitiveVerification));
  assert.equal(a.m.commands, 1);
  assert.deepEqual({ fitW: sheetP(store).fitW, fitH: sheetP(store).fitH, fitFlag: sheetP(store).fitFlag }, { fitW: 0, fitH: 0, fitFlag: false });
  const r = await run(store, one(STEPS_P.fitOff));
  assert.equal(r.m.commands, 0);
  assert.equal(r.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('fitOff: SATISFIED at once on an untouched sheet (fitW=fitH=0, flag already off by default)', async () => {
  const { result, m } = await run(createStore(book()), one(STEPS_P.fitOff));
  assert.equal(result.ok, true, JSON.stringify(result.execution.primitiveVerification));
  assert.equal(m.commands, 0);
  assert.equal(result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('fitOff RED: the numbers land but the flag incorrectly stays on -> L1 FAIL', async () => {
  const { result, driver } = await run(createStore(fitOn()), one(STEPS_P.fitOff), { fault: 'fit-flag-stuck-on' });
  assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL');
  assert.equal(driver.calls.forceSave, 0);
});

test('verticalOff: APPLIED from an already-true sheet, the retry SATISFIED 0/0', async () => {
  const store = createStore(vOn());
  const a = await run(store, one(STEPS_P.verticalOff));
  assert.equal(a.result.ok, true, JSON.stringify(a.result.execution.primitiveVerification));
  assert.equal(a.m.commands, 1);
  assert.equal(sheetP(store).vCenter, false);
  const r = await run(store, one(STEPS_P.verticalOff));
  assert.equal(r.m.commands, 0);
  assert.equal(r.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('verticalOff: SATISFIED at once on an untouched sheet (vCenter already false by default)', async () => {
  const { result, m } = await run(createStore(book()), one(STEPS_P.verticalOff));
  assert.equal(result.ok, true, JSON.stringify(result.execution.primitiveVerification));
  assert.equal(m.commands, 0);
  assert.equal(result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('planning: each step claims its own axis -- a footer and a header of one sheet compose, two footers do not', () => {
  assert.equal(planOperations({ id: 'x', steps: [{ sheet: 'P', ...STEPS.footerSet }, { sheet: 'P', ...STEPS.headerClear }] }).ok, true);
  assert.equal(planOperations({ id: 'x', steps: [{ sheet: 'P', ...STEPS.footerSet }, { sheet: 'P', ...STEPS.footerClear }] }).ok, false);
  assert.equal(planOperations({ id: 'x', steps: [{ sheet: 'P', ...STEPS.vertical }, { sheet: 'P', ...STEPS.horizontalOff }] }).ok, true);
});

/**
 * zero (m7-p49-live1, 2026-10-05): the fitToPage FLAG takes ONE boolean -- Worksheet.setFitToPage(val) stores val as
 * sheetPr.FitToPage, and the editor's own fit path passes `!fitToHeightAuto || !fitToWidthAuto` (WorksheetView.js
 * _changeFitToPage). The lowering used to pass (fitW, fitH), so the flag became fitW: 50 for a 50x60 fit (the driver
 * reads `=== true || === 1` -> false, verify failed after the write) and 0 for a 0 x 60 fit (the flag OFF while a
 * height fit was asked). Only 1 x 1 was ever LIVE-tested, where fitW happened to be 1.
 */
test('set_print_fit: the fitToPage flag is the boolean "any fit dimension set", never the width', () => {
  const calls = [];
  const ps = { asc_setFitToWidth: x => calls.push(['w', x]), asc_setFitToHeight: x => calls.push(['h', x]), asc_setScale: x => calls.push(['scale', x]) };
  const m = { PagePrintOptions: { asc_getPageSetup: () => ps }, setFitToPage: (...a) => calls.push(['flag', ...a]) };
  const ws = { GetName: () => 'P', worksheet: m };
  const H = { Create_NewPoint: () => {}, StartTransaction: () => {}, EndTransaction: () => {} };
  const flagFor = (w, h) => {
    calls.length = 0;
    const op = planOperations(one({ intent: 'set_print_fit', fitToWidth: w, fitToHeight: h })).operations[0];
    runViaBridge(op, ws, { History: H });
    return calls.filter(c => c[0] === 'flag').map(c => c.slice(1));
  };
  assert.deepEqual(flagFor(50, 60), [[true]]);
  assert.deepEqual(flagFor(0, 60), [[true]]);
  assert.deepEqual(flagFor(1, 1), [[true]]);
  assert.deepEqual(flagFor(0, 0), [[false]]);
});

/**
 * Owner decision 1/a (owner channel, 2026-10-05): fit-to-page OFF also resets the print scale to 100 %, as the
 * editor's own UI does (Toolbar.js: asc_SetPrintScale(0, 0, 100)) -- p53 measured the editor's fit-on leaving scale 16,
 * which the old product fit-off kept. The scale is read on every print step (an unwanted scale change fails L3).
 */
const fitOnScaled = () => { const b = fitOn(); b[0].print = { ...b[0].print, scale: 16 }; return b; };

test('fitOff resets the scale: from fit 50 x 60 at scale 16 -> APPLIED, fit 0 / 0, flag off, scale 100; retry 0/0', async () => {
  const store = createStore(fitOnScaled());
  const a = await run(store, one(STEPS_P.fitOff));
  assert.equal(a.result.ok, true, JSON.stringify(a.result.execution.primitiveVerification));
  assert.deepEqual(['fitW', 'fitH', 'fitFlag', 'scale'].map(k => sheetP(store)[k]), [0, 0, false, 100]);
  const r = await run(store, one(STEPS_P.fitOff));
  assert.equal(r.m.commands, 0);
  assert.equal(r.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
});

test('fitOff RED: the scale stays at 16 -> L1 FAIL (a fit-off is not satisfied while the scale is not 100)', async () => {
  const { result } = await run(createStore(fitOnScaled()), one(STEPS_P.fitOff), { fault: 'fit-scale-stuck' });
  assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL');
});

test('fitOff is NOT satisfied by fit 0 / 0 + flag off alone while the scale is 16 -- it writes', async () => {
  const b = book(); b[0].print = { ...b[0].print, fitW: 0, fitH: 0, fitFlag: false, scale: 16 };
  const { result, m } = await run(createStore(b), one(STEPS_P.fitOff));
  assert.equal(result.ok, true, JSON.stringify(result.execution.primitiveVerification));
  assert.equal(m.commands, 1);
});

test('the scale is in every print read: a fit-ON step that also moves the scale fails L3', async () => {
  const { result } = await run(createStore(book()), one({ intent: 'set_print_fit', fitToWidth: 1, fitToHeight: 1 }), { fault: 'print-side-scale' });
  assert.equal(result.ok, false);
  assert.equal(result.execution.preservationVerification[0].verdict, 'FAIL');
});

test('the fit-off command writes asc_setScale(100); a fit-on command does not touch the scale (through the /7 Bridge WRITE)', () => {
  const calls = [];
  const env = (noScale = false) => { calls.length = 0; const ps = { asc_setFitToWidth: x => calls.push(['w', x]), asc_setFitToHeight: x => calls.push(['h', x]) }; if (!noScale) ps.asc_setScale = x => calls.push(['scale', x]);
    return { GetName: () => 'P', worksheet: { PagePrintOptions: { asc_getPageSetup: () => ps }, setFitToPage: x => calls.push(['flag', x]) } }; };
  const opOf = step => planOperations(one(step)).operations[0];
  assert.deepEqual(runViaBridge(opOf(STEPS_P.fitOff), env(), {}), { ok: true });
  assert.deepEqual(calls, [['w', 0], ['h', 0], ['flag', false], ['scale', 100]]);
  const r = runViaBridge(opOf(STEPS_P.fitOff), env(true), {});
  assert.deepEqual([r.ok, r.reason, calls.length], [false, 'BRIDGE:asc_setScale-missing', 0], 'checked before the first call, like every other setter');
  assert.deepEqual(runViaBridge(opOf({ intent: 'set_print_fit', fitToWidth: 1, fitToHeight: 1 }), env(), {}), { ok: true });
  assert.deepEqual(calls, [['w', 1], ['h', 1], ['flag', true]]);
});

/**
 * OD-A (% scale) and OD-C (paper size), owner channel: set_print_scale = the UI's asc_SetPrintScale(0, 0, n)
 * (fit OFF + the scale); set_paper_size = one of the editor's own menu sizes (exact mm; anything else is saved as A4).
 */
test('set_print_scale 75: from fit 50 x 60 -> APPLIED, fit 0 / 0, flag off, scale 75; retry 0/0', async () => {
  const store = createStore(fitOnScaled());
  const a = await run(store, one({ intent: 'set_print_scale', scale: 75 }));
  assert.equal(a.result.ok, true, JSON.stringify(a.result.execution.primitiveVerification));
  assert.deepEqual(['fitW', 'fitH', 'fitFlag', 'scale'].map(k => sheetP(store)[k]), [0, 0, false, 75]);
  const r = await run(store, one({ intent: 'set_print_scale', scale: 75 }));
  assert.equal(r.m.commands, 0);
});

test('set_print_scale RED: the scale lands one off -> L1 FAIL', async () => {
  const { result } = await run(createStore(book()), one({ intent: 'set_print_scale', scale: 75 }), { fault: 'print-wrong-value' });
  assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL');
});

test('set_paper_size US Letter -> APPLIED 215.9 x 279.4, retry 0/0; A4 on a default sheet is SATISFIED at once', async () => {
  const store = createStore(book());
  const a = await run(store, one({ intent: 'set_paper_size', paper: 'US Letter' }));
  assert.equal(a.result.ok, true, JSON.stringify(a.result.execution.primitiveVerification));
  assert.deepEqual([sheetP(store).width, sheetP(store).height], [215.9, 279.4]);
  assert.equal((await run(store, one({ intent: 'set_paper_size', paper: 'US Letter' }))).m.commands, 0);
  assert.equal((await run(createStore(book()), one({ intent: 'set_paper_size', paper: 'A4' }))).m.commands, 0);
});

test('set_paper_size RED: the width lands wrong -> L1 FAIL; a paper step moving the scale -> L3 FAIL', async () => {
  assert.equal((await run(createStore(book()), one({ intent: 'set_paper_size', paper: 'A3' }), { fault: 'print-wrong-value' })).result.execution.primitiveVerification[0].verdict, 'FAIL');
  const side = await run(createStore(book()), one({ intent: 'set_paper_size', paper: 'A3' }), { fault: 'paper-side-scale' });
  assert.equal(side.result.execution.preservationVerification[0].verdict, 'FAIL');
});

test('planning: a scale outside [10, 400], a non-integer scale, an unknown paper name or raw mm are refused; scale and fit share one axis', () => {
  for (const s of [{ intent: 'set_print_scale', scale: 9 }, { intent: 'set_print_scale', scale: 401 }, { intent: 'set_print_scale', scale: 75.5 },
    { intent: 'set_paper_size', paper: 'A6' }, { intent: 'set_paper_size', paper: [216, 279] }]) assert.equal(planOperations(one(s)).ok, false, JSON.stringify(s));
  assert.equal(planOperations({ id: 'x', steps: [{ sheet: 'P', intent: 'set_print_scale', scale: 75 }, { sheet: 'P', intent: 'set_print_fit', fitToWidth: 1, fitToHeight: 1 }] }).ok, false);
  assert.equal(planOperations({ id: 'x', steps: [{ sheet: 'P', intent: 'set_print_scale', scale: 75 }, { sheet: 'P', intent: 'set_paper_size', paper: 'A3' }] }).ok, true);
});

const pageEnv = (missing = []) => {
  const calls = []; const ps = {};
  for (const k of ['asc_setFitToWidth', 'asc_setFitToHeight', 'asc_setScale', 'asc_setWidth', 'asc_setHeight']) if (!missing.includes(k)) ps[k] = x => calls.push([k, x]);
  const m = { PagePrintOptions: { asc_getPageSetup: () => ps } }; if (!missing.includes('setFitToPage')) m.setFitToPage = x => calls.push(['setFitToPage', x]);
  return { calls, ws: { GetName: () => 'P', worksheet: m } };
};
test('the command bodies: scale = the UI shape, paper = width + height, every setter checked first (through the /7 Bridge WRITE)', () => {
  const opOf = step => planOperations(one(step)).operations[0];
  let e = pageEnv(); assert.deepEqual(runViaBridge(opOf({ intent: 'set_print_scale', scale: 75 }), e.ws, {}), { ok: true });
  assert.deepEqual(e.calls, [['asc_setFitToWidth', 0], ['asc_setFitToHeight', 0], ['setFitToPage', false], ['asc_setScale', 75]]);
  e = pageEnv(); assert.deepEqual(runViaBridge(opOf({ intent: 'set_paper_size', paper: 'US Letter' }), e.ws, {}), { ok: true });
  assert.deepEqual(e.calls, [['asc_setWidth', 215.9], ['asc_setHeight', 279.4]]);
  e = pageEnv(['asc_setHeight']); const r = runViaBridge(opOf({ intent: 'set_paper_size', paper: 'A4' }), e.ws, {});
  assert.deepEqual([r.ok, r.reason, e.calls.length], [false, 'BRIDGE:asc_setHeight-missing', 0]);
});

test('set_paper_size RED: only the height lands wrong -> L1 FAIL (both dimensions are proven)', async () => {
  const { result } = await run(createStore(book()), one({ intent: 'set_paper_size', paper: 'A3' }), { fault: 'paper-wrong-height' });
  assert.equal(result.execution.primitiveVerification[0].verdict, 'FAIL');
});

test('set_print_scale checks all four setters before the first call (through the /7 Bridge WRITE)', () => {
  const op = planOperations(one({ intent: 'set_print_scale', scale: 75 })).operations[0];
  for (const f of ['asc_setFitToWidth', 'asc_setFitToHeight', 'setFitToPage', 'asc_setScale']) {
    const e = pageEnv([f]); const r = runViaBridge(op, e.ws, {});
    assert.deepEqual([r.ok, r.reason, e.calls.length], [false, `BRIDGE:${f}-missing`, 0], f);
  }
});

/** X-P07 (OD-A, owner channel): first / even header-footer texts and the two enabling flags. */
test('X-P07: each first / even slot -> APPLIED (L1 + L3), retry 0/0; the odd slots untouched', async () => {
  for (const slot of ['firstHeader', 'firstFooter', 'evenHeader', 'evenFooter']) {
    const store = createStore(book());
    const t = one({ intent: 'set_header_footer', slot, value: `&C${slot}` });
    const a = await run(store, t);
    assert.equal(a.result.ok, true, `${slot} ${JSON.stringify(a.result.execution.primitiveVerification)}`);
    assert.equal(sheetP(store)[slot], `&C${slot}`);
    assert.deepEqual([sheetP(store).oddHeader, sheetP(store).oddFooter], ['&Chead0', '&Cfoot0']);
    assert.equal((await run(store, t)).m.commands, 0, slot);
  }
});

test('X-P07: the two flags -> APPLIED true, retry 0/0, then false APPLIED; a wrong flag value -> L1 FAIL', async () => {
  for (const flag of ['differentFirst', 'differentOddEven']) {
    const store = createStore(book());
    const on = one({ intent: 'set_header_footer_flag', flag, value: true });
    assert.equal((await run(store, on)).result.ok, true, flag);
    assert.equal((await run(store, on)).m.commands, 0, flag);
    assert.equal((await run(store, one({ intent: 'set_header_footer_flag', flag, value: false }))).result.ok, true, flag);
    const red = await run(createStore(book()), on, { fault: 'print-wrong-value' });
    assert.equal(red.result.execution.primitiveVerification[0].verdict, 'FAIL', flag);
  }
});

test('X-P07: an unwidened read refuses a first / even slot and a flag before the write (fail-closed)', async () => {
  for (const step of [{ intent: 'set_header_footer', slot: 'evenFooter', value: '&Cx' }, { intent: 'set_header_footer_flag', flag: 'differentFirst', value: true }]) {
    const { result, m } = await run(createStore(book()), one(step), { printWide: false });
    assert.equal(result.ok, false, JSON.stringify(step));
    assert.deepEqual([m.writes, m.commands], [0, 0]);
  }
});

test('X-P07: the command bodies -- the slot setter / the flag setter inside a closed transaction, checked first (through the /7 Bridge WRITE)', () => {
  const opOf = step => planOperations(one(step)).operations[0];
  const hfEnv = (missing = []) => { const log = []; const hf = {}; for (const k of ['setFirstFooter', 'setDifferentOddEven']) if (!missing.includes(k)) hf[k] = x => log.push([k, x]);
    const H = { Create_NewPoint: () => log.push(['np']), StartTransaction: () => log.push(['st']), EndTransaction: () => log.push(['et']) };
    return { log, ws: { GetName: () => 'P', worksheet: { PagePrintOptions: { asc_getPageSetup: () => ({}) }, headerFooter: hf } }, asc: { History: H } }; };
  let e = hfEnv(); assert.deepEqual(runViaBridge(opOf({ intent: 'set_header_footer', slot: 'firstFooter', value: '&Cf' }), e.ws, e.asc), { ok: true });
  assert.deepEqual(e.log, [['np'], ['st'], ['setFirstFooter', '&Cf'], ['et']]);
  e = hfEnv(['setFirstFooter']); const r = runViaBridge(opOf({ intent: 'set_header_footer', slot: 'firstFooter', value: '&Cf' }), e.ws, e.asc);
  assert.deepEqual([r.ok, r.reason, e.log], [false, 'BRIDGE:setFirstFooter-missing', []], 'checked before the transaction opens');
  e = hfEnv(); assert.deepEqual(runViaBridge(opOf({ intent: 'set_header_footer_flag', flag: 'differentOddEven', value: true }), e.ws, e.asc), { ok: true });
  assert.deepEqual(e.log, [['np'], ['st'], ['setDifferentOddEven', true], ['et']]);
  for (const bad of [{ intent: 'set_header_footer_flag', flag: 'scaleWithDoc', value: true }, { intent: 'set_header_footer_flag', flag: 'differentFirst', value: 1 }]) {
    assert.equal(planOperations(one(bad)).ok, false, JSON.stringify(bad));
  }
});
