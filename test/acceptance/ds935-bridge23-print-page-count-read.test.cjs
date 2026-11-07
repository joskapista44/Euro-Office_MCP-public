'use strict';

/**
 * GH #356 / #357 -- the euro-office-bridge-v1 /23 READ candidate for DS 9.3.5 (documentserver-patches/9.3.5.0/euro-office-bridge-v1/
 * rev23-candidate/): sheet.printPageCount.read/v1. Offline: the REAL candidate block is evaluated against a fake window / Api / History
 * in the SHAPES read from the served 9.3.5 cell sdk (WorkbookView.calcPagesPrint -> WorksheetView.calcPagesPrint -> recalcPrintScale
 * -> _setPrintScale on the model's CURRENT PagePrintOptions: History.Create_NewPoint (a no-op while TurnOffHistory > 0) +
 * asc_setScale; changeZoom(1) and back; asc_CPageOptions.prototype.clone(ws); c_kMaxPrintPages). This proves the HANDLER LOGIC -- the
 * same logic measured on the engine before the patch (docs/evidence/gh357-clone-pagecount-probe-live-r1-20261101.json), NOT the block
 * itself on the engine (design gates G1-G3, NOT MEASURED).
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');
const { spawnSync } = require('node:child_process');

const DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.5.0', 'euro-office-bridge-v1');
const V22 = fs.readFileSync(path.join(DIR, 'rev22-candidate', 'bridge-block.js'), 'utf8');
const V23 = fs.readFileSync(path.join(DIR, 'rev23-candidate', 'bridge-block.js'), 'utf8');
const PC = 'sheet.printPageCount.read/v1';
const sha = t => crypto.createHash('sha256').update(t).digest('hex');
const reason = r => (r && r.ok === false ? r.reason : `OK:${JSON.stringify(r && r.value)}`);

// ---- the fake engine ----------------------------------------------------------------------------------

/** pages: the count the layout gives; fit: the sheet is fit-to-page (the engine computes FIT_SCALE and writes it to the CURRENT options) */
function engine({ fit = true, pages = 1, hidden = false, cap = 1500, calcThrows = false, zoomStuck = false } = {}) {
  const H = {
    TurnOffHistory: 0, Points: [],
    TurnOff() { this.TurnOffHistory += 1; }, TurnOn() { this.TurnOffHistory = Math.max(0, this.TurnOffHistory - 1); },
    Is_On() { return this.TurnOffHistory === 0; },
    Create_NewPoint() { if (this.TurnOffHistory !== 0) return false; this.Points.push({ items: [] }); return true; },
  };
  const FIT_SCALE = 40;
  const options = ws => {
    const o = { ws, scale: 100, fitW: fit ? 1 : 0, fitH: fit ? 1 : 0 };
    o.asc_getPageSetup = () => ({
      asc_getScale: () => o.scale, asc_getFitToWidth: () => o.fitW, asc_getFitToHeight: () => o.fitH,
      asc_setScale: (v, notHist) => { const old = o.scale; o.scale = v; if (!notHist && o.ws && H.Is_On() && old !== v) H.Points[H.Points.length - 1].items.push(['scale', old, v]); },
    });
    o.clone = w => { const c = options(w); c.scale = o.scale; c.fitW = o.fitW; c.fitH = o.fitH; return c; };
    return o;
  };
  const m = { getIndex: () => 3, getHidden: () => hidden };
  m.PagePrintOptions = options(m);
  const real = m.PagePrintOptions;
  const wb = {
    zoom: 0.75, calls: [],
    getZoom() { return this.zoom; },
    changeZoom(z) { this.zoom = z; },
    calcPagesPrint(adj) {
      this.calls.push({ type: adj.printType, sheets: adj.activeSheetsArray, optionsSeen: m.PagePrintOptions, historyOff: H.TurnOffHistory });
      const viewZoom = this.zoom;
      this.changeZoom(1);
      const ps = m.PagePrintOptions.asc_getPageSetup();
      if (ps.asc_getFitToWidth() || ps.asc_getFitToHeight()) {
        if (ps.asc_getScale() !== FIT_SCALE) { H.Create_NewPoint(); ps.asc_setScale(FIT_SCALE, false); }
      }
      if (calcThrows) throw new Error('layout');
      if (!zoomStuck) this.changeZoom(viewZoom);
      return { arrPages: new Array(Math.min(pages, cap)).fill({}) };
    },
  };
  function asc_CAdjustPrint() { this.printType = 0; this.activeSheetsArray = null; }
  asc_CAdjustPrint.prototype.asc_setPrintType = function (v) { this.printType = v; };
  asc_CAdjustPrint.prototype.asc_setActiveSheetsArray = function (v) { this.activeSheetsArray = v; };
  const Asc = { editor: { wb }, asc_CAdjustPrint, c_oAscPrintType: { ActiveSheets: 0, EntireWorkbook: 1, Selection: 2 } };
  return { H, m, real, wb, FIT_SCALE, window: { Asc, AscCommonExcel: { c_kMaxPrintPages: cap } }, AscCommon: { History: H }, sheetObj: { worksheet: m } };
}

function load(e, block = V23) {
  const sheets = { S: e.sheetObj };
  const Api = { GetSheet: n => (Object.prototype.hasOwnProperty.call(sheets, n) ? sheets[n] : null) };
  new Function('Api', 'Asc', 'window', 'AscCommon', block.replace(/\}\s*,\s*$/, '}'))(Api, {}, e.window, e.AscCommon);
  return Api.GetEuroOfficeBridge();
}
/** the no-write invariant: the real options object, its scale, no History point, History back on, the zoom */
const untouched = e => ({ identity: e.m.PagePrintOptions === e.real, scale: e.real.scale, points: e.H.Points.length, off: e.H.TurnOffHistory, zoom: e.wb.zoom });
const CLEAN = { identity: true, scale: 100, points: 0, off: 0, zoom: 0.75 };

// ---- the build ----------------------------------------------------------------------------------------

test('the /23 block = the /22 block with exactly the builder substitutions; the builder is deterministic', () => {
  const r = spawnSync('python3', ['-I', '-c', 'import importlib.util,json,sys;s=importlib.util.spec_from_file_location("b",sys.argv[1]);m=importlib.util.module_from_spec(s);s.loader.exec_module(m);v22,v23=m.build_block();print(json.dumps({"sha22":m.V22_BLOCK_SHA,"v23":v23.decode("utf-8"),"impl":m.one_line(m.HANDLER_SRC),"entries":m.ENTRIES}))',
    path.join(DIR, 'rev23-candidate', 'build-rev23.py')], { encoding: 'utf8', maxBuffer: 1 << 24 });
  assert.equal(r.status, 0, r.stderr);
  const b = JSON.parse(r.stdout);
  assert.equal(sha(V22), b.sha22, 'the builder pins the /22 block');
  assert.equal(b.v23, V23, 'the committed block is exactly what the builder produces');
  const back = V23.replace(b.impl, '').replace(b.entries, '')
    .replace('patchVersion: "euro-office-bridge-v1/9.3.5.0/23"', 'patchVersion: "euro-office-bridge-v1/9.3.5.0/22"');
  assert.equal(back, V22, 'every /22 capability (manifest + handlers + the layout installer) is byte-for-byte unchanged');
});

test('RED control: the /22 block does not advertise the read', () => {
  assert.equal(reason(load(engine(), V22).read(PC, { sheet: 'S' })), 'capability-not-advertised');
});

test('manifest: /23, the read listed as READ / mutation NO, every /22 capability still listed', () => {
  const man = load(engine()).manifest();
  assert.equal(man.patchVersion, 'euro-office-bridge-v1/9.3.5.0/23');
  const c = man.capabilities.find(x => x.id === PC);
  assert.equal(c.kind, 'READ'); assert.equal(c.mutation, 'NO');
  const ids = man.capabilities.map(x => x.id);
  const ids22 = load(engine(), V22).manifest().capabilities.map(x => x.id);
  for (const id of ids22) assert.ok(ids.includes(id), id);
  assert.equal(ids.length, ids22.length + 1);
});

// ---- the read ----------------------------------------------------------------------------------------

test('RED control of the fake: a DIRECT calcPagesPrint on a fit sheet writes the scale and a History point (the M1/M2 r1 shape)', () => {
  const e = engine();
  e.wb.calcPagesPrint({ printType: 0, activeSheetsArray: [3] });
  assert.equal(e.real.scale, 40);
  assert.equal(e.H.Points.length, 1);
});

test('fit sheet: pages 1, the engine computed the fit scale on the CLONE, the live model untouched (no History point, scale, identity, zoom)', () => {
  const e = engine({ fit: true, pages: 1 });
  assert.deepEqual(load(e).read(PC, { sheet: 'S' }), { ok: true, value: { pages: 1 } });
  assert.deepEqual(untouched(e), CLEAN);
  const call = e.wb.calls[0];
  assert.deepEqual({ type: call.type, sheets: call.sheets, off: call.historyOff }, { type: 0, sheets: [3], off: 1 }, 'ActiveSheets [the sheet index], History off for the call');
  assert.notEqual(call.optionsSeen, e.real, 'the engine saw the clone, not the live options');
  assert.equal(call.optionsSeen.scale, e.FIT_SCALE, 'the computed fit scale landed on the clone');
});

test('the read is repeatable: three calls give the same count and leave the model clean', () => {
  const e = engine({ fit: true, pages: 1 });
  const br = load(e);
  for (let i = 0; i < 3; i += 1) assert.equal(reason(br.read(PC, { sheet: 'S' })), 'OK:{"pages":1}');
  assert.deepEqual(untouched(e), CLEAN);
});

test('no-fit sheet: pages 9; an empty layout is a legitimate 0, not a refusal', () => {
  let e = engine({ fit: false, pages: 9 });
  assert.equal(reason(load(e).read(PC, { sheet: 'S' })), 'OK:{"pages":9}');
  assert.deepEqual(untouched(e), CLEAN);
  e = engine({ fit: false, pages: 0 });
  assert.equal(reason(load(e).read(PC, { sheet: 'S' })), 'OK:{"pages":0}');
});

test('a hidden sheet is sheet-hidden (never 0) and the engine is not called', () => {
  const e = engine({ hidden: true });
  assert.equal(reason(load(e).read(PC, { sheet: 'S' })), 'sheet-hidden');
  assert.equal(e.wb.calls.length, 0);
});

test('a count at the engine cap is page-count-capped, never a truncated number', () => {
  const e = engine({ fit: false, pages: 5000, cap: 1500 });
  assert.equal(reason(load(e).read(PC, { sheet: 'S' })), 'page-count-capped');
  assert.deepEqual(untouched(e), CLEAN);
});

test('a throwing layout is print-calc-threw, and the options, History and zoom are restored anyway', () => {
  const e = engine({ fit: true, calcThrows: true });
  assert.equal(reason(load(e).read(PC, { sheet: 'S' })), 'print-calc-threw');
  assert.deepEqual(untouched(e), CLEAN);
});

test('a zoom the engine left moved is put back by the handler', () => {
  const e = engine({ fit: false, pages: 9, zoomStuck: true });
  assert.equal(reason(load(e).read(PC, { sheet: 'S' })), 'OK:{"pages":9}');
  assert.equal(e.wb.zoom, 0.75);
});

test('refusals before any engine call, named', () => {
  const cases = [
    [undefined, 'bad-args'], [{}, 'bad-args'], [{ sheet: '' }, 'bad-args'], [{ sheet: 7 }, 'bad-args'], [{ sheet: 'NOPE' }, 'sheet-not-found'],
  ];
  for (const [args, want] of cases) {
    const e = engine();
    assert.equal(reason(load(e).read(PC, args)), want, JSON.stringify(args));
    assert.equal(e.wb.calls.length, 0);
  }
  const missing = [
    [e => { delete e.window.Asc.editor.wb; }, 'print-calc-unavailable'],
    [e => { delete e.window.Asc.asc_CAdjustPrint; }, 'print-calc-unavailable'],
    [e => { delete e.window.AscCommonExcel.c_kMaxPrintPages; }, 'print-calc-unavailable'],
    [e => { delete e.AscCommon.History.TurnOff; }, 'history-missing'],
    [e => { e.real.clone = () => { throw new Error('x'); }; }, 'options-clone-threw'],
    [e => { e.real.clone = () => e.real; }, 'options-clone-threw'],
    [e => { delete e.m.getHidden; }, 'model-missing'],
  ];
  for (const [mutate, want] of missing) {
    const e = engine();
    mutate(e);
    assert.equal(reason(load(e).read(PC, { sheet: 'S' })), want);
    assert.equal(e.wb.calls.length, 0);
    assert.deepEqual(untouched(e), CLEAN);
  }
});

test('a non-array result is print-calc-unexpected-shape', () => {
  const e = engine();
  e.wb.calcPagesPrint = () => ({ arrPages: 3 });
  assert.equal(reason(load(e).read(PC, { sheet: 'S' })), 'print-calc-unexpected-shape');
});

test('sabotage: the same block WITHOUT the clone swap writes the live scale -- the no-write assertions turn red', () => {
  const impl = 'm.PagePrintOptions = copy;';
  assert.equal(V23.split(impl).length - 1, 1);
  const e = engine({ fit: true, pages: 1 });
  assert.equal(reason(load(e, V23.replace(impl, '')).read(PC, { sheet: 'S' })), 'OK:{"pages":1}');
  assert.notDeepEqual(untouched(e), CLEAN, 'without the swap the engine wrote the live options');
  assert.equal(e.real.scale, 40);
});

// ---- the plan ----------------------------------------------------------------------------------------

test('plan-rev23.py: refuses an unknown input and writes nothing', () => {
  const tmp = fs.mkdtempSync(path.join(require('node:os').tmpdir(), 'p23-'));
  const input = path.join(tmp, 'sdk-all.js'); fs.writeFileSync(input, 'not the unpatched sdk');
  const out = path.join(tmp, 'plan.json');
  const r = spawnSync('python3', ['-I', path.join(DIR, 'rev23-candidate', 'plan-rev23.py'), '--input', input, '--out', out], { encoding: 'utf8' });
  assert.equal(r.status, 2, r.stderr);
  assert.match(r.stderr, /UNKNOWN_INPUT/);
  assert.equal(fs.existsSync(out), false);
  fs.rmSync(tmp, { recursive: true, force: true });
});
