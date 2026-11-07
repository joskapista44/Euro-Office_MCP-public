'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { xlsxAdapter, planOperations, liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const print = require('../../src/adapters/xlsx/families/print.cjs');
const { projectPageSetupRead } = require('../../src/adapters/xlsx/families/page-read.cjs');
const { readQuery, b26BridgeWritePrelude, B26_BRIDGE_WRITE_RUNTIMES } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { createStore, createFakeXlsxDriver } = require('../fakes/fake-xlsx-driver.cjs');

/** GH #354 / #355: the print LAYOUT intents (header/footer margin, page order) over the Bridge /21 sheet.printSettings.write/v2. */

const R19 = '9.3.5|euro-office-bridge-v1/9.3.5.0/19';
const R21 = '9.3.5|euro-office-bridge-v1/9.3.5.0/21';
const R22 = '9.3.5|euro-office-bridge-v1/9.3.5.0/22';
const R23 = '9.3.5|euro-office-bridge-v1/9.3.5.0/23';
const R24 = '9.3.5|euro-office-bridge-v1/9.3.5.0/24';
const R25 = '9.3.5|euro-office-bridge-v1/9.3.5.0/25'; // W2 regression fix (workbook.protection.read/v1 null branch) -- sheet.printSettings.write/v2 byte-identical to /24
const R26 = '9.3.5|euro-office-bridge-v1/9.3.5.0/26'; // review fix (f0451f47, the two guard change) -- sheet.printSettings.write/v2 untouched, byte-identical to /25
const R27 = '9.3.5|euro-office-bridge-v1/9.3.5.0/27'; // GH #346 AUTO-REPAIR 1. egyseg (chart.series.style.read/v1 dataPoint only) -- sheet.printSettings.write/v2 byte-identical to /26
const R28 = '9.3.5|euro-office-bridge-v1/9.3.5.0/28'; // GH #361/#363 pre-install readiness gate (internal tracker) -- byte-identical warrant, ds935-bridge28-table-autofilter-criteria-write-handler-identity.test.cjs
const R29 = '9.3.5|euro-office-bridge-v1/9.3.5.0/29'; // GH #365 option C U1 (internal tracker) -- byte-identical warrant, ds935-bridge29-veryhidden-write-handler-identity.test.cjs
const book = () => [{ name: 'P', cells: { A1: { v: 1 } } }, { name: 'O', cells: { A1: { v: 'keep' } }, layout: { headerMargin: 9, footerMargin: 9, pageOrder: 'overThenDown' } }];
const one = step => ({ id: 'gh354', steps: [{ sheet: 'P', ...step }] });
const STEPS = {
  header: { intent: 'set_header_footer_margin', slot: 'header', value: 12.5 },
  footer: { intent: 'set_header_footer_margin', slot: 'footer', value: 0 },
  order: { intent: 'set_page_order', order: 'overThenDown' },
};

async function run(store, t, script = {}) {
  let n = 0;
  const driver = createFakeXlsxDriver(store, { runtime: R21, ...script });
  const options = { driver, fileId: 'disposable', now: () => (n += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(t, { observation: { live: liveScope(t) } }));
  return { result, m: result.metrics };
}
const evidenceOf = result => result.execution.appliedOperations[0].evidence.evidence;

test('the runtime list: sheet.printSettings.write/v2 is /21 + /22 + /23 + /24 + /25 + /26 (the /22, /23, /24, /25, /26 and /27 copies byte-identical; adapter and driver agree)', () => {
  assert.deepEqual([...print.B26_BRIDGE_WRITE_RUNTIMES], [R21, R22, R23, R24, R25, R26, R27, R28, R29]);
  assert.deepEqual([...B26_BRIDGE_WRITE_RUNTIMES], [R21, R22, R23, R24, R25, R26, R27, R28, R29]);
});

for (const [k, step] of Object.entries(STEPS)) {
  test(`${k}: APPLIED on /21 (L1 on the v2 read + L3), the retry SATISFIED 0/0`, async () => {
    const store = createStore(book());
    const a = await run(store, one(step));
    assert.equal(a.result.ok, true, JSON.stringify([a.result.execution.primitiveVerification, a.result.execution.preservationVerification]));
    assert.equal(a.m.commands, 1);
    const r = await run(store, one(step));
    assert.equal(r.m.commands, 0);
    assert.equal(r.result.execution.appliedOperations[0].status, OPERATION_STATUS.SATISFIED);
  });
}

test('RED: a wrong value -> L1 FAIL; another layout field moves -> L3 FAIL', async () => {
  for (const step of Object.values(STEPS)) {
    const w = await run(createStore(book()), one(step), { fault: 'print-wrong-value' });
    assert.equal(w.result.execution.primitiveVerification[0].verdict, 'FAIL', JSON.stringify(step));
    const s = await run(createStore(book()), one(step), { fault: 'print-side' });
    assert.equal(s.result.ok, false, JSON.stringify(step));
    assert.equal(s.result.execution.preservationVerification[0].verdict, 'FAIL', JSON.stringify(step));
  }
});

test('NEG: a /19 runtime (no v2) -> FAIL_CLOSED, named, zero commands; the read gap is the layout field only', async () => {
  for (const step of Object.values(STEPS)) {
    const a = await run(createStore(book()), one(step), { runtime: R19, layoutBuild: false });
    assert.equal(a.m.commands, 0);
    assert.equal(a.result.ok, false);
    assert.equal(evidenceOf(a.result).reason, 'xlsx-print-layout-unobservable', JSON.stringify(evidenceOf(a.result)));
    assert.equal(evidenceOf(a.result).why, 'UNOBSERVABLE:CAPABILITY_MISSING');
    // the layout field readable but the runtime not /21 -> refused before the write, named
    const b = await run(createStore(book()), one(step), { runtime: R19 });
    assert.equal(b.m.commands, 0);
    assert.equal(evidenceOf(b.result).reason, 'xlsx-print-layout-bridge-write-runtime-not-proven');
  }
});

test('plan: the shapes the family refuses, named, before any dispatch', () => {
  const bad = [
    [{ intent: 'set_header_footer_margin', slot: 'left', value: 5 }, /slot must be 'header' or 'footer'/],
    [{ intent: 'set_header_footer_margin', slot: 'header', value: 200.5 }, /millimetres in \[0, 200\]/],
    [{ intent: 'set_header_footer_margin', slot: 'header', value: -1 }, /millimetres/],
    [{ intent: 'set_header_footer_margin', slot: 'header', value: '5' }, /millimetres/],
    [{ intent: 'set_page_order', order: 'acrossThenDown' }, /order must be one of downThenOver, overThenDown/],
    [{ intent: 'set_page_order', order: 'overThenDown', value: 1 }, /does not support field\(s\): value/],
  ];
  for (const [step, re] of bad) {
    const p = planOperations(one(step));
    assert.equal(p.ok, false, JSON.stringify(step));
    assert.match(JSON.stringify(p), re);
  }
});

test('the lowered command: one sheet.printSettings.write/v2 through the B26 client; meta marks it a layout write', () => {
  const p = planOperations(one(STEPS.order));
  const cmd = print.lower(p.operations[0]);
  assert.match(cmd.body, /__eoB26Gate\(name\)/);
  assert.match(cmd.body, /__eoB26Write\(\{"setting":"pageOrder","value":"overThenDown"\}, \["pageOrder"\]\)/);
  assert.doesNotMatch(cmd.body, /__eoB10/);
  assert.equal(cmd.meta.bridgePrintWrite, true);
  assert.equal(cmd.meta.bridgePrintLayoutWrite, true);
  // every other print intent stays on the v1 client, without the layout mark
  const q = planOperations(one({ intent: 'set_print_scale', scale: 55 }));
  const v1 = print.lower(q.operations[0]);
  assert.match(v1.body, /__eoB10Gate\(name\)/);
  assert.equal('bridgePrintLayoutWrite' in v1.meta, false);
});

test('liveScope asks for the layout field ONLY for a layout step', () => {
  assert.equal(liveScope(one(STEPS.header)).printLayout, true);
  assert.equal('printLayout' in liveScope(one({ intent: 'set_print_scale', scale: 55 })), false);
});

test('readQuery: the v2 read is generated only when asked, as its own axis', () => {
  assert.doesNotMatch(readQuery([], { printSheets: ['P'] }), /sheet\.printSettings\.read\/v2/);
  const q = readQuery([], { printSheets: ['P'], printLayout: true });
  assert.match(q, /sheet\.printSettings\.read\/v2/);
  assert.match(q, /layout: axis\(\(\) => \{ const x = bridgeAxis/);
});

/** the B26 client + the lowered body against a fake Bridge: the gate's runtime list, the args, the receipt check */
function runClient(step, { patchVersion = 'euro-office-bridge-v1/9.3.5.0/21', receipt } = {}) {
  const calls = [];
  const bridge = {
    manifest: () => ({ bridgeVersion: '1.0.0', documentServerVersion: '9.3.5', patchVersion,
      capabilities: patchVersion.endsWith('/21') ? [{ id: 'sheet.printSettings.write/v2', version: 2, kind: 'WRITE', mutation: 'TARGETED' }, { id: 'sheet.printSettings.read/v2', version: 2, kind: 'READ', mutation: 'NO' }] : [] }),
    write: (id, args) => { calls.push([id, args]); return receipt || { ok: true, value: { applied: [args.setting] } }; },
    read: () => ({ ok: true, value: { headerMargin: 1, footerMargin: 2, pageOrder: 'downThenOver' } }),
  };
  const ws = { GetName: () => 'P' };
  const api = { GetEuroOfficeBridge: () => bridge, GetSheet: n => (n === 'P' ? ws : null) };
  const op = planOperations(one(step)).operations[0];
  const out = new Function('Api', b26BridgeWritePrelude() + print.lower(op).body)(api);
  return { out, calls };
}

test('B26 client: /21 writes exactly once with the capability args; /19 refuses at the gate, nothing written', () => {
  const a = runClient(STEPS.header);
  assert.deepEqual(a.out, { ok: true });
  assert.deepEqual(a.calls, [['sheet.printSettings.write/v2', { sheet: 'P', setting: 'headerMargin', value: 12.5 }]]);
  const b = runClient(STEPS.header, { patchVersion: 'euro-office-bridge-v1/9.3.5.0/19' });
  assert.equal(b.out.ok, false);
  assert.equal(b.out.stage, 'bridge-write-gate');
  assert.equal(b.out.reason, 'PATCH_VERSION_MISMATCH');
  assert.equal(b.calls.length, 0);
  // a bridge reason comes back named, with the fresh v2 observation
  const c = runClient(STEPS.order, { receipt: { ok: false, reason: 'no-effect' } });
  assert.equal(c.out.ok, false);
  assert.equal(c.out.reason, 'BRIDGE:no-effect');
  assert.deepEqual(c.out.observed, { headerMargin: 1, footerMargin: 2, pageOrder: 'downThenOver' });
  // the satisfied-receipt shape is NOT a write receipt (the adapter never writes a value that already holds)
  const d = runClient(STEPS.order, { receipt: { ok: true, value: { applied: [], satisfied: true } } });
  assert.equal(d.out.ok, false);
  assert.equal(d.out.reason, 'UNEXPECTED_SHAPE');
});

test('driver source: the B26 runtime check sits before the editor call, on the meta mark', () => {
  const src = require('node:fs').readFileSync(require.resolve('../../src/drivers/eurooffice/xlsx-driver.cjs'), 'utf8');
  const gate = src.indexOf("const b26 = b10 && command.meta.bridgePrintLayoutWrite === true;");
  const refuse = src.indexOf("outcome: 'b26-bridge-write-runtime-not-proven'");
  const call = src.indexOf('(b26 ? b26BridgeWritePrelude() : b10BridgeWritePrelude()) + body');
  assert.ok(gate > 0 && refuse > gate && call > refuse);
});

test('read_page_setup: the layout fields mirror the write words; a /19 gap is these three only, named', () => {
  const sheets = [{ name: 'P' }];
  const base = { fitW: 0, fitH: 0, fitFlag: false, scale: 100, width: 210, height: 297, rowBreaks: [], colBreaks: [], oddHeader: null, oddFooter: null, firstHeader: null, firstFooter: null, evenHeader: null, evenFooter: null, diffFirst: false, diffOddEven: false, useFirstPage: false, firstPage: null, hCenter: false, vCenter: false };
  const ok = projectPageSetupRead({ sheets, printReads: { P: { ...base, layout: { headerMargin: 7.62, footerMargin: null, pageOrder: 'overThenDown' } } } }, { sheet: 'P' });
  assert.equal(ok.print.headerMargin, 7.62);
  assert.equal(ok.print.footerMargin, null);
  assert.equal(ok.print.pageOrder, 'overThenDown');
  assert.equal('unrepresentable' in ok.print, false);
  const gap = projectPageSetupRead({ sheets, printReads: { P: { ...base, layout: 'UNOBSERVABLE:CAPABILITY_MISSING' } } }, { sheet: 'P' });
  assert.equal(gap.print.layoutReason, 'UNOBSERVABLE:CAPABILITY_MISSING');
  assert.equal(gap.print.scale, 100, 'the v1 half stands');
  const odd = projectPageSetupRead({ sheets, printReads: { P: { ...base, layout: { headerMargin: -3, footerMargin: 5, pageOrder: 'sideways' } } } }, { sheet: 'P' });
  assert.deepEqual(odd.print.unrepresentable.map(x => x.field), ['headerMargin', 'pageOrder']);
  // an unasked read (no layout field) carries no layout keys at all
  const none = projectPageSetupRead({ sheets, printReads: { P: base } }, { sheet: 'P' });
  assert.equal('pageOrder' in none.print, false);
});
