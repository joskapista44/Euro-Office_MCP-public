'use strict';

/**
 * GH #354 / #355 -- the euro-office-bridge-v1 /21 WRITE candidate for DS 9.3.5 (documentserver-patches/9.3.5.0/euro-office-bridge-v1/
 * rev21-candidate/): a LOAD-TIME installer (Layout history change types 16 / 17 / 18) + sheet.printSettings.write/v2.
 * Offline: the REAL candidate block is evaluated against a fake window / Api / History in the SHAPES read from the factory 9.3.5
 * sdk-all.js (UndoRedoRedoLayout.prototype.UndoRedo, UndoRedoData_Layout(from, to), asc_CPageMargins.asc_get/setHeader/Footer,
 * asc_CPageSetup.pageOrder, History.Create_NewPoint/StartTransaction/Add/EndTransaction). This proves the HANDLER and INSTALLER LOGIC --
 * NOT that a 16-18 change survives co-editing, the server-side save and a fresh reopen on the real engine (design gates G1-G3, NOT MEASURED).
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');
const { spawnSync } = require('node:child_process');

const DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.5.0', 'euro-office-bridge-v1');
const V20 = fs.readFileSync(path.join(DIR, 'rev20-candidate', 'bridge-block.js'), 'utf8');
const V21 = fs.readFileSync(path.join(DIR, 'rev21-candidate', 'bridge-block.js'), 'utf8');
const W2 = 'sheet.printSettings.write/v2';
const sha = t => crypto.createHash('sha256').update(t).digest('hex');
const reason = r => (r && r.ok === false ? r.reason : `OK:${JSON.stringify(r && r.value)}`);

// ---- the fake engine ----------------------------------------------------------------------------------

/** the factory Layout undo class: types 1-15 only, a switch with no default (an unknown type is a silent no-op) */
function layoutClass(log) {
  function UndoRedoRedoLayout(wb) { this.wb = wb; }
  UndoRedoRedoLayout.prototype.Undo = function (Type, Data, nSheetId) { this.UndoRedo(Type, Data, nSheetId, true); };
  UndoRedoRedoLayout.prototype.Redo = function (Type, Data, nSheetId) { this.UndoRedo(Type, Data, nSheetId, false); };
  UndoRedoRedoLayout.prototype.UndoRedo = function (Type, Data, nSheetId, bUndo) {
    if (Type >= 1 && Type <= 15) log.push(['orig', Type, bUndo ? Data.from : Data.to]);
  };
  return UndoRedoRedoLayout;
}

function engine({ header = 7.62, footer = 7.62, order = 0, setterIgnores = false, noSetHeader = false, addThrows = false, historyOn = true } = {}) {
  const log = [];
  const pm = { header, footer, asc_getHeader() { return this.header; }, asc_getFooter() { return this.footer; },
    asc_setFooter(v) { if (!setterIgnores) this.footer = v; } };
  if (!noSetHeader) pm.asc_setHeader = function (v) { if (!setterIgnores) this.header = v; };
  const ps = { pageOrder: order };
  const ppo = { asc_getPageMargins: () => pm, asc_getPageSetup: () => ps };
  const triggers = [];
  const m = { PagePrintOptions: ppo, getId: () => 'sheet-7', workbook: { handlers: { trigger: (...a) => triggers.push(a) } } };
  const wb = { getWorksheetById: id => (id === 'sheet-7' ? m : null), oApi: { _onUpdateLayoutMenu: id => log.push(['menu', id]) } };
  const UndoRedoRedoLayout = layoutClass(log);
  function UndoRedoData_Layout(from, to) { this.from = from; this.to = to; }
  const X = { UndoRedoRedoLayout, UndoRedoData_Layout, g_oUndoRedoLayout: new UndoRedoRedoLayout(wb) };
  const C = { historyitem_Layout_Orientation: 11 };
  const H = {
    calls: [],
    Create_NewPoint() { this.calls.push('point'); }, StartTransaction() { this.calls.push('start'); }, EndTransaction() { this.calls.push('end'); },
    Is_On: () => historyOn,
    Add(cls, type, sheetId, range, data) { if (addThrows) throw new Error('x'); this.calls.push(['add', cls === X.g_oUndoRedoLayout, type, sheetId, data.from, data.to]); },
  };
  return { log, pm, ps, m, wb, X, C, H, triggers, window: { AscCommonExcel: X, AscCH: C }, AscCommon: { History: H } };
}

/** evaluates the block the way it sits in sdk-all.js: the installer runs at load, then the bridge entrypoint is defined */
function load(e, block = V21, sheets = { S: { worksheet: e.m } }) {
  const Api = { GetSheet: n => (Object.prototype.hasOwnProperty.call(sheets, n) ? sheets[n] : null) };
  new Function('Api', 'Asc', 'window', 'AscCommon', block.replace(/\}\s*,\s*$/, '}'))(Api, {}, e.window, e.AscCommon);
  return Api.GetEuroOfficeBridge();
}

// ---- the build ----------------------------------------------------------------------------------------

test('the /21 block = installer + the /20 block with exactly the builder substitutions; the builder is deterministic', () => {
  const r = spawnSync('python3', ['-I', '-c', 'import importlib.util,json,sys;s=importlib.util.spec_from_file_location("b",sys.argv[1]);m=importlib.util.module_from_spec(s);s.loader.exec_module(m);v20,v21=m.build_block();print(json.dumps({"sha20":m.V20_BLOCK_SHA,"v21":v21.decode("utf-8"),"inst":m.one_line(m.INSTALLER_SRC),"impl":m.one_line(m.HANDLER_SRC),"entry":m.ENTRY}))',
    path.join(DIR, 'rev21-candidate', 'build-rev21.py')], { encoding: 'utf8', maxBuffer: 1 << 24 });
  assert.equal(r.status, 0, r.stderr);
  const b = JSON.parse(r.stdout);
  assert.equal(sha(V20), b.sha20, 'the builder pins the /20 candidate block');
  assert.equal(b.v21, V21, 'the committed block is exactly what the builder produces');
  const back = V21.slice(b.inst.length)
    .replace(b.impl, '')
    .replace(b.entry, '')
    .replace('patchVersion: "euro-office-bridge-v1/9.3.5.0/21"', 'patchVersion: "euro-office-bridge-v1/9.3.5.0/20"');
  assert.ok(V21.startsWith(b.inst), 'the installer is the very first expression of the block');
  assert.equal(back, V20, 'every /20 capability (manifest + handlers) is byte-for-byte unchanged');
});

test('RED control: the /20 block does not advertise write/v2 and installs no Layout type', () => {
  const e = engine();
  const br = load(e, V20);
  assert.equal(reason(br.write(W2, { sheet: 'S', setting: 'headerMargin', value: 10 })), 'capability-not-advertised');
  assert.equal(e.C.historyitem_Layout_HeaderMargin, undefined);
});

// ---- the installer -----------------------------------------------------------------------------------

test('installer: types 16/17/18 on AscCH, UndoRedo wrapped for exactly those types, every other type reaches the original', () => {
  const e = engine({ header: 7.62, footer: 7.62, order: 0 });
  load(e);
  assert.deepEqual([e.C.historyitem_Layout_HeaderMargin, e.C.historyitem_Layout_FooterMargin, e.C.historyitem_Layout_PageOrder], [16, 17, 18]);
  const L = e.X.g_oUndoRedoLayout;
  L.Redo(16, new e.X.UndoRedoData_Layout(7.62, 12.7), 'sheet-7');
  L.Redo(17, new e.X.UndoRedoData_Layout(7.62, 5), 'sheet-7');
  L.Redo(18, new e.X.UndoRedoData_Layout(0, 1), 'sheet-7');
  assert.deepEqual([e.pm.header, e.pm.footer, e.ps.pageOrder], [12.7, 5, 1], 'redo applies the "to" value');
  L.Undo(16, new e.X.UndoRedoData_Layout(7.62, 12.7), 'sheet-7');
  L.Undo(18, new e.X.UndoRedoData_Layout(0, 1), 'sheet-7');
  assert.deepEqual([e.pm.header, e.ps.pageOrder], [7.62, 0], 'undo applies the "from" value');
  L.Redo(11, new e.X.UndoRedoData_Layout(0, 1), 'sheet-7');
  assert.deepEqual(e.log.filter(x => x[0] === 'orig'), [['orig', 11, 1]], 'type 11 goes to the original function, 16-18 never do');
  L.Redo(16, new e.X.UndoRedoData_Layout(1, 2), 'no-such-sheet');
  assert.equal(e.pm.header, 7.62, 'an unknown sheet id is a no-op, like the original');
});

test('installer: idempotent (a second load does not wrap twice) and inert without the Layout class', () => {
  const e = engine();
  load(e);
  const wrapped = e.X.UndoRedoRedoLayout.prototype.UndoRedo;
  load(e);
  assert.equal(e.X.UndoRedoRedoLayout.prototype.UndoRedo, wrapped);
  const bare = engine();
  delete bare.X.UndoRedoRedoLayout;
  const br = load(bare);
  assert.equal(bare.C.historyitem_Layout_HeaderMargin, undefined, 'nothing installed');
  assert.equal(reason(br.write(W2, { sheet: 'S', setting: 'headerMargin', value: 10 })), 'layout-types-not-installed');
  delete bare.X.g_oUndoRedoLayout;
  assert.equal(reason(br.write(W2, { sheet: 'S', setting: 'headerMargin', value: 10 })), 'layout-history-missing');
});

// ---- sheet.printSettings.write/v2 -------------------------------------------------------------------

test('write/v2 headerMargin: one point + transaction, History.Add(Layout, 16, sheet, old -> new), receipt, page-setup event', () => {
  const e = engine();
  const r = load(e).write(W2, { sheet: 'S', setting: 'headerMargin', value: 12.7 });
  assert.deepEqual(r, { ok: true, value: { applied: ['headerMargin'] } });
  assert.equal(e.pm.header, 12.7);
  assert.deepEqual(e.H.calls, ['point', 'start', ['add', true, 16, 'sheet-7', 7.62, 12.7], 'end']);
  assert.deepEqual(e.triggers, [['onChangePageSetupProps', 'sheet-7']]);
});

test('write/v2 footerMargin -> type 17; pageOrder overThenDown -> stored 1, type 18; downThenOver -> 0', () => {
  const e = engine({ order: 0 });
  const br = load(e);
  assert.equal(reason(br.write(W2, { sheet: 'S', setting: 'footerMargin', value: 5 })), 'OK:{"applied":["footerMargin"]}');
  assert.equal(reason(br.write(W2, { sheet: 'S', setting: 'pageOrder', value: 'overThenDown' })), 'OK:{"applied":["pageOrder"]}');
  assert.equal(e.ps.pageOrder, 1);
  assert.equal(reason(br.write(W2, { sheet: 'S', setting: 'pageOrder', value: 'downThenOver' })), 'OK:{"applied":["pageOrder"]}');
  assert.equal(e.ps.pageOrder, 0);
  const adds = e.H.calls.filter(c => Array.isArray(c)).map(c => [c[2], c[4], c[5]]);
  assert.deepEqual(adds, [[17, 7.62, 5], [18, 0, 1], [18, 1, 0]]);
});

test('write/v2 retry: the target value already held is SATISFIED with NO history', () => {
  const e = engine({ header: 12.7, order: 1 });
  const br = load(e);
  assert.deepEqual(br.write(W2, { sheet: 'S', setting: 'headerMargin', value: 12.7 }), { ok: true, value: { applied: [], satisfied: true } });
  assert.deepEqual(br.write(W2, { sheet: 'S', setting: 'pageOrder', value: 'overThenDown' }), { ok: true, value: { applied: [], satisfied: true } });
  assert.deepEqual(e.H.calls, []);
});

test('write/v2 refusals before any mutation, named', () => {
  const e = engine();
  const br = load(e);
  const cases = [
    [null, 'bad-args'], [{ setting: 'headerMargin', value: 1 }, 'bad-args'], [{ sheet: 'S', value: 1 }, 'bad-args'],
    [{ sheet: 'S', setting: 'headerMargin' }, 'bad-args'],
    [{ sheet: 'S', setting: 'leftMargin', value: 1 }, 'unknown-setting'],
    [{ sheet: 'S', setting: 'headerMargin', value: -1 }, 'invalid-value'], [{ sheet: 'S', setting: 'headerMargin', value: 500.1 }, 'invalid-value'],
    [{ sheet: 'S', setting: 'headerMargin', value: NaN }, 'invalid-value'], [{ sheet: 'S', setting: 'footerMargin', value: '5' }, 'invalid-value'],
    [{ sheet: 'S', setting: 'pageOrder', value: 1 }, 'invalid-value'], [{ sheet: 'S', setting: 'pageOrder', value: 'overthendown' }, 'invalid-value'],
    [{ sheet: 'Nope', setting: 'headerMargin', value: 1 }, 'sheet-not-found'],
  ];
  for (const [args, want] of cases) assert.equal(reason(br.write(W2, args)), want, JSON.stringify(args));
  assert.deepEqual(e.H.calls, [], 'no history touched');
  assert.deepEqual([e.pm.header, e.pm.footer, e.ps.pageOrder], [7.62, 7.62, 0], 'nothing written');
});

test('write/v2 engine-side refusals: history off / setter missing / no-effect / Add throws (value restored)', () => {
  let e = engine({ historyOn: false });
  assert.equal(reason(load(e).write(W2, { sheet: 'S', setting: 'headerMargin', value: 9 })), 'history-off');
  assert.equal(e.pm.header, 7.62);
  e = engine({ noSetHeader: true });
  assert.equal(reason(load(e).write(W2, { sheet: 'S', setting: 'headerMargin', value: 9 })), 'asc_setHeader-missing');
  e = engine({ setterIgnores: true });
  assert.equal(reason(load(e).write(W2, { sheet: 'S', setting: 'footerMargin', value: 9 })), 'no-effect', 'an unchanged read-back is never success');
  e = engine({ addThrows: true });
  assert.equal(reason(load(e).write(W2, { sheet: 'S', setting: 'headerMargin', value: 9 })), 'history-add-threw');
  assert.equal(e.pm.header, 7.62, 'the value set before the failed Add is restored');
  assert.deepEqual(e.H.calls, ['point', 'start', 'end'], 'the transaction is always closed');
});

test('write/v2 refuses when the installer did not run (a /20 engine object with the /21 handler)', () => {
  const e = engine();
  const br = load(e);
  delete e.X.UndoRedoRedoLayout.prototype.__eoLayoutV21;
  assert.equal(reason(br.write(W2, { sheet: 'S', setting: 'pageOrder', value: 'overThenDown' })), 'layout-types-not-installed');
  assert.equal(e.ps.pageOrder, 0);
});
