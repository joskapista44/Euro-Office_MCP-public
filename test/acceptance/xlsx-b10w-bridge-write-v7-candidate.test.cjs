'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const print = require('../../src/adapters/xlsx/families/print.cjs');

/**
 * M11.5b.1 (owner: M11.5 continuous mandate, owner channel, internal tracker) -- OFFLINE acceptance of the /7
 * CANDIDATE block (documentserver-patches/9.3.4.60/euro-office-bridge-v1/rev7-candidate/bridge-block.js): the B10-WRITE
 * Bridge capability sheet.printSettings.write/v1. The ACTUAL candidate text runs against a fake Api whose model objects
 * record every call (the /6 candidate test pattern). PARITY: for every B10 intent the bridge makes exactly the same
 * mutation calls, in the same order, as the product's direct-internal lowering made -- recorded from the last source that
 * had it (9c8633b, test/fixtures/b10-direct-internal-call-parity-9c8633b.json; M11.5b.4 removed that path) against the
 * same fake model. The bridge is the same engine path, moved behind an explicit capability.
 */

const DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.4.60', 'euro-office-bridge-v1');
const read = p => fs.readFileSync(p, 'utf8');
const V6 = read(path.join(DIR, 'rev6-candidate', 'bridge-block.js'));
const V7 = read(path.join(DIR, 'rev7-candidate', 'bridge-block.js'));
const bridgeOf = (src, api, ascCommon) => new Function('Api', 'AscCommon', `${src.replace(/,\s*$/, '')}; return Api.GetEuroOfficeBridge();`)(api, ascCommon);
const W = 'sheet.printSettings.write/v1';

const SETTERS = {
  ps: ['asc_setFitToWidth', 'asc_setFitToHeight', 'asc_setScale', 'asc_setWidth', 'asc_setHeight', 'asc_setFirstPageNumber'],
  ppo: ['asc_setHorizontalCentered', 'asc_setVerticalCentered'],
  m: ['setFitToPage', 'changeRowColBreaks'],
  hf: ['setOddHeader', 'setOddFooter', 'setFirstHeader', 'setFirstFooter', 'setEvenHeader', 'setEvenFooter', 'setDifferentFirst', 'setDifferentOddEven'],
  H: ['Create_NewPoint', 'StartTransaction', 'EndTransaction'],
};

/** fake sheet `S`: every model object records its calls as [owner, method, ...args]. */
function fakeEnv({ missing = [], throws = [], noHistory = false, noModel = false } = {}) {
  const calls = [];
  const obj = owner => {
    const o = {};
    for (const k of SETTERS[owner]) {
      if (missing.includes(k)) continue;
      o[k] = (...a) => { calls.push([owner, k, ...a]); if (throws.includes(k)) throw new Error(`boom-${k}`); };
    }
    return o;
  };
  const ps = obj('ps'); const ppo = obj('ppo'); const hf = obj('hf'); const m = obj('m'); const H = obj('H');
  ppo.asc_getPageSetup = () => ps;
  m.PagePrintOptions = ppo; m.headerFooter = hf;
  m.getPrintAreaRangeByRowCol = (r, c) => ({ printArea: [r, c] });
  const ws = { GetName: () => 'S', get worksheet() { return noModel ? null : m; } };
  const api = { GetSheet: n => (n === 'S' ? ws : null) };
  const asc = noHistory ? {} : { History: H };
  return { api, asc, calls, mutations: () => calls.filter(c => c[0] !== 'H') };
}
const write = (env, args, id = W) => bridgeOf(V7, env.api, env.asc).write(id, args);
const PRE_CUTOVER = JSON.parse(fs.readFileSync(path.join(__dirname, '..', 'fixtures', 'b10-direct-internal-call-parity-9c8633b.json'), 'utf8'));

test('/7 manifest: patchVersion /7, /6 capabilities byte-identical, exactly one new WRITE entry', () => {
  const m6 = bridgeOf(V6, { GetSheet: () => null }).manifest(); const m7 = bridgeOf(V7, { GetSheet: () => null }).manifest();
  assert.equal(m6.patchVersion, 'euro-office-bridge-v1/9.3.4.60/6');
  assert.equal(m7.patchVersion, 'euro-office-bridge-v1/9.3.4.60/7');
  assert.equal(m7.capabilities.length, m6.capabilities.length + 1);
  assert.deepEqual(m7.capabilities.slice(0, m6.capabilities.length), m6.capabilities);
  const w = m7.capabilities.filter(c => c.id === W);
  assert.equal(w.length, 1);
  assert.deepEqual([w[0].version, w[0].kind, w[0].format, w[0].mutation], [1, 'WRITE', 'XLSX', 'TARGETED']);
  for (const k of ['bridgeVersion', 'documentServerVersion', 'documentServerBuild']) assert.equal(m7[k], m6[k]);
});

test('/7 = /6 + two inserted segments + the patchVersion bump (every /6 implementation byte-identical)', () => {
  let t = V7.replace('patchVersion: "euro-office-bridge-v1/9.3.4.60/7"', 'patchVersion: "euro-office-bridge-v1/9.3.4.60/6"');
  const cut = (s, startMark, endMark) => { const i = s.indexOf(startMark); const j = s.indexOf(endMark, i); assert.ok(i > 0 && j > i, startMark); return s.slice(0, i) + s.slice(j); };
  t = cut(t, '{ id: "sheet.printSettings.write/v1"', '], };');
  t = cut(t, '"sheet.printSettings.write/v1": function', '}; var EO_BRIDGE_HYDRATIONS = ');
  assert.equal(t, V6);
});

test('class separation: read()/hydrate() refuse the WRITE id; write() refuses READ ids and other versions', () => {
  const env = fakeEnv(); const b = bridgeOf(V7, env.api, env.asc);
  const args = { sheet: 'S', setting: 'scale', scale: 50 };
  assert.deepEqual(b.read(W, args), { ok: false, reason: 'capability-not-advertised' });
  assert.deepEqual(b.hydrate(W, args), { ok: false, reason: 'capability-not-advertised' });
  for (const id of ['sheet.printSettings.read/v1', 'sheet.printSettings.write/v2', 'nope']) assert.deepEqual(b.write(id, args), { ok: false, reason: 'capability-not-advertised' }, id);
  assert.deepEqual(env.calls, []);
  assert.equal(typeof bridgeOf(V6, env.api, env.asc).write('range.alignmentExt.write/v1', { sheet: 'X', cell: 'A1', indent: 1 }).reason, 'string', '/6 write() still serves B05');
  assert.deepEqual(bridgeOf(V6, env.api, env.asc).write(W, args), { ok: false, reason: 'capability-not-advertised' }, '/6: B10 WRITE missing is detectable');
});

// [product step (print.cjs intent), bridge args, expected receipt]
const PARITY = [
  [{ intent: 'set_print_fit', fitToWidth: 1, fitToHeight: 2 }, { setting: 'fit', fitToWidth: 1, fitToHeight: 2 }, ['fitW', 'fitH', 'fitFlag']],
  [{ intent: 'set_print_fit', fitToWidth: 0, fitToHeight: 0 }, { setting: 'fit', fitToWidth: 0, fitToHeight: 0 }, ['fitW', 'fitH', 'fitFlag', 'scale']],
  [{ intent: 'set_print_scale', scale: 75 }, { setting: 'scale', scale: 75 }, ['fitW', 'fitH', 'fitFlag', 'scale']],
  [{ intent: 'set_paper_size', paper: 'A4' }, { setting: 'paper', width: 210, height: 297 }, ['width', 'height']],
  [{ intent: 'set_paper_size', paper: 'US Letter' }, { setting: 'paper', width: 215.9, height: 279.4 }, ['width', 'height']],
  [{ intent: 'set_first_page_number', value: 3 }, { setting: 'firstPage', firstPage: 3 }, ['firstPage']],
  [{ intent: 'set_first_page_number', value: null }, { setting: 'firstPage', firstPage: null }, ['firstPage']],
  [{ intent: 'set_print_centering', horizontal: true }, { setting: 'center', axis: 'horizontal', value: true }, ['hCenter']],
  [{ intent: 'set_print_centering', vertical: false }, { setting: 'center', axis: 'vertical', value: false }, ['vCenter']],
  [{ intent: 'set_header_footer', slot: 'oddHeader', value: '&CTitle' }, { setting: 'headerFooter', slot: 'oddHeader', value: '&CTitle' }, ['oddHeader']],
  [{ intent: 'set_header_footer', slot: 'evenFooter', value: null }, { setting: 'headerFooter', slot: 'evenFooter', value: null }, ['evenFooter']],
  [{ intent: 'set_header_footer_flag', flag: 'differentFirst', value: true }, { setting: 'headerFooterFlag', flag: 'differentFirst', value: true }, ['diffFirst']],
  [{ intent: 'set_header_footer_flag', flag: 'differentOddEven', value: false }, { setting: 'headerFooterFlag', flag: 'differentOddEven', value: false }, ['diffOddEven']],
  [{ intent: 'set_page_break', axis: 'row', at: 11 }, { setting: 'pageBreak', axis: 'row', id: 10, remove: false }, ['rowBreaks']],
  [{ intent: 'set_page_break', axis: 'row', at: 11, remove: true }, { setting: 'pageBreak', axis: 'row', id: 10, remove: true }, ['rowBreaks']],
  [{ intent: 'set_page_break', axis: 'column', at: 'E' }, { setting: 'pageBreak', axis: 'column', id: 4, remove: false }, ['colBreaks']],
  [{ intent: 'set_page_break', axis: 'column', at: 'E', remove: true }, { setting: 'pageBreak', axis: 'column', id: 4, remove: true }, ['colBreaks']],
];

test('PARITY: for every B10 intent the bridge makes EXACTLY the former product direct-internal calls (incl. history), in order', () => {
  assert.equal(PRE_CUTOVER.source, '9c8633b920f3b2799418109d586cea475971559e');
  assert.equal(Object.keys(PRE_CUTOVER.calls).length, PARITY.length);
  assert.deepEqual([...new Set(PARITY.map(p => p[0].intent))].sort(), [...print.INTENTS].sort(), 'all 8 B10 intents covered');
  for (const [step, args, applied] of PARITY) {
    const pre = PRE_CUTOVER.calls[JSON.stringify(step)]; assert.ok(Array.isArray(pre) && pre.length >= 1, JSON.stringify(step));
    const be = fakeEnv(); assert.deepEqual(write(be, { sheet: 'S', ...args }), { ok: true, value: { applied } }, JSON.stringify(args));
    assert.deepEqual(JSON.parse(JSON.stringify(be.calls)), pre, JSON.stringify(step));
    assert.deepEqual(print.bridgeArgs(print.plan({ ...step, sheet: 'S' }, 0).operation), { args, requested: applied }, `${JSON.stringify(step)}: the product sends exactly these args`);
    assert.ok(be.mutations().length >= 1, JSON.stringify(step));
  }
});

test('header/footer settings run inside ONE history transaction that is always closed, also when the setter throws', () => {
  let e = fakeEnv();
  write(e, { sheet: 'S', setting: 'headerFooter', slot: 'firstHeader', value: 'x' });
  assert.deepEqual(e.calls.map(c => c[1]), ['Create_NewPoint', 'StartTransaction', 'setFirstHeader', 'EndTransaction']);
  e = fakeEnv({ throws: ['setDifferentFirst'] });
  assert.deepEqual(write(e, { sheet: 'S', setting: 'headerFooterFlag', flag: 'differentFirst', value: true }), { ok: false, reason: 'setDifferentFirst-threw', partial: false, applied: [] });
  assert.deepEqual(e.calls.map(c => c[1]), ['Create_NewPoint', 'StartTransaction', 'setDifferentFirst', 'EndTransaction']);
  e = fakeEnv({ noHistory: true });
  assert.deepEqual(write(e, { sheet: 'S', setting: 'headerFooter', slot: 'oddHeader', value: 'x' }), { ok: false, reason: 'history-missing' });
  assert.deepEqual(e.calls, []);
  e = fakeEnv({ noHistory: true });
  assert.deepEqual(write(e, { sheet: 'S', setting: 'scale', scale: 50 }), { ok: true, value: { applied: ['fitW', 'fitH', 'fitFlag', 'scale'] } }, 'non-transaction settings do not need History');
});

test('invalid inputs refuse BEFORE any mutation (zero calls)', () => {
  const cases = [
    [undefined, 'bad-args'], [{ setting: 'scale', scale: 50 }, 'bad-args'], [{ sheet: '', setting: 'scale', scale: 50 }, 'bad-args'],
    [{ sheet: 'S', scale: 50 }, 'bad-args'], [{ sheet: 'S', setting: 'scale' }, 'bad-args'],
    [{ sheet: 'S', setting: 'margins', top: 1 }, 'unknown-setting'], [{ sheet: 'S', setting: 'toString' }, 'unknown-setting'],
    [{ sheet: 'S', setting: 'scale', scale: 50, fitToWidth: 1 }, 'unsupported-property'], [{ sheet: 'S', setting: 'scale', scale: 50, requester: 'x' }, 'unsupported-property'],
    ...[9, 401, 50.5, '50', null, NaN].map(v => [{ sheet: 'S', setting: 'scale', scale: v }, 'invalid-value']),
    ...[[-1, 1], [101, 1], [1, 1.5], ['1', 1]].map(([w, h]) => [{ sheet: 'S', setting: 'fit', fitToWidth: w, fitToHeight: h }, 'invalid-value']),
    ...[[210, 296], [0, 0], ['210', 297]].map(([w, h]) => [{ sheet: 'S', setting: 'paper', width: w, height: h }, 'invalid-value']),
    ...[0, 32768, 1.5, '2'].map(v => [{ sheet: 'S', setting: 'firstPage', firstPage: v }, 'invalid-value']),
    [{ sheet: 'S', setting: 'center', axis: 'diagonal', value: true }, 'invalid-value'], [{ sheet: 'S', setting: 'center', axis: 'horizontal', value: 1 }, 'invalid-value'],
    [{ sheet: 'S', setting: 'headerFooter', slot: 'header', value: 'x' }, 'invalid-value'], [{ sheet: 'S', setting: 'headerFooter', slot: 'oddHeader', value: '' }, 'invalid-value'],
    [{ sheet: 'S', setting: 'headerFooter', slot: 'oddHeader', value: 'x'.repeat(256) }, 'invalid-value'],
    [{ sheet: 'S', setting: 'headerFooterFlag', flag: 'scaleWithDoc', value: true }, 'invalid-value'],
    [{ sheet: 'S', setting: 'pageBreak', axis: 'row', id: 0, remove: false }, 'invalid-value'], [{ sheet: 'S', setting: 'pageBreak', axis: 'column', id: 16384, remove: false }, 'invalid-value'],
    [{ sheet: 'S', setting: 'pageBreak', axis: 'row', id: 5, remove: 'yes' }, 'invalid-value'], [{ sheet: 'S', setting: 'pageBreak', axis: 'z', id: 5, remove: false }, 'invalid-value'],
    [{ sheet: 'X', setting: 'scale', scale: 50 }, 'sheet-not-found'],
  ];
  for (const [args, reason] of cases) { const e = fakeEnv(); assert.deepEqual(write(e, args), { ok: false, reason }, JSON.stringify(args)); assert.deepEqual(e.calls, [], JSON.stringify(args)); }
  const e = fakeEnv({ noModel: true }); assert.deepEqual(write(e, { sheet: 'S', setting: 'scale', scale: 50 }), { ok: false, reason: 'model-missing' }); assert.deepEqual(e.calls, []);
});

test('missing setter: refused before the FIRST mutation, also when the earlier setters exist', () => {
  let e = fakeEnv({ missing: ['asc_setScale'] });
  assert.deepEqual(write(e, { sheet: 'S', setting: 'scale', scale: 50 }), { ok: false, reason: 'asc_setScale-missing' }); assert.deepEqual(e.calls, []);
  e = fakeEnv({ missing: ['asc_setScale'] });
  assert.deepEqual(write(e, { sheet: 'S', setting: 'fit', fitToWidth: 0, fitToHeight: 0 }), { ok: false, reason: 'asc_setScale-missing' }, 'fit OFF needs the scale setter');
  assert.deepEqual(e.calls, []);
  e = fakeEnv({ missing: ['asc_setScale'] });
  assert.deepEqual(write(e, { sheet: 'S', setting: 'fit', fitToWidth: 1, fitToHeight: 1 }), { ok: true, value: { applied: ['fitW', 'fitH', 'fitFlag'] } }, 'an unneeded missing setter does not block');
  e = fakeEnv({ missing: ['changeRowColBreaks'] });
  assert.deepEqual(write(e, { sheet: 'S', setting: 'pageBreak', axis: 'row', id: 3, remove: false }), { ok: false, reason: 'changeRowColBreaks-missing' }); assert.deepEqual(e.mutations(), []);
  e = fakeEnv({ missing: ['setEvenHeader'] });
  assert.deepEqual(write(e, { sheet: 'S', setting: 'headerFooter', slot: 'evenHeader', value: 'x' }), { ok: false, reason: 'setEvenHeader-missing' }); assert.deepEqual(e.calls, [], 'no history point opened either');
});

test('setter throws: controlled failure; PARTIAL write reported with the applied fields, never rolled back, no fallback', () => {
  let e = fakeEnv({ throws: ['asc_setFitToWidth'] });
  assert.deepEqual(write(e, { sheet: 'S', setting: 'scale', scale: 50 }), { ok: false, reason: 'asc_setFitToWidth-threw', partial: false, applied: [] });
  e = fakeEnv({ throws: ['setFitToPage'] });
  assert.deepEqual(write(e, { sheet: 'S', setting: 'scale', scale: 50 }), { ok: false, reason: 'setFitToPage-threw', partial: true, applied: ['fitW', 'fitH'] });
  assert.deepEqual(e.mutations().map(c => c[1]), ['asc_setFitToWidth', 'asc_setFitToHeight', 'setFitToPage'], 'nothing after the throw, no compensating call');
  e = fakeEnv({ throws: ['asc_setHeight'] });
  assert.deepEqual(write(e, { sheet: 'S', setting: 'paper', width: 210, height: 297 }), { ok: false, reason: 'asc_setHeight-threw', partial: true, applied: ['width'] });
});

test('the receipt carries no observed value (verification is a fresh sheet.printSettings.read/v1)', () => {
  const r = write(fakeEnv(), { sheet: 'S', setting: 'firstPage', firstPage: 7 });
  assert.deepEqual(Object.keys(r.value), ['applied']);
});
