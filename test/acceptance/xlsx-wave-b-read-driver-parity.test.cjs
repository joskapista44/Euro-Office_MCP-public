'use strict';

/**
 * Wave B READ driver cutover (B05 READ, B09, B13 -> DS Bridge /4), MOCK parity old == bridge.
 *
 * OLD_* below are the pre-cutover in-command readers, copied verbatim from the driver at 5ea34bb3 (the direct internal
 * reads this cutover removes). Each fixture is read twice on the SAME fake model: by the OLD reader, and by today's
 * driver readQuery through the REAL /4 bridge block (test/fakes/ds-bridge-runtime.cjs). A proven value must be equal;
 * a gap must stay a gap. The known, stated deviations (design 3.1-3.3) are asserted as such, not hidden.
 * A /3 runtime (the shipped old.txt block, and the /4 block reporting /3) refuses every migrated axis -- no fallback.
 * LIVE parity is a separate, owner-gated step (test/live/xlsx-wave-b-read-live-acceptance.live.cjs).
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { readQuery } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { installBridge, withPatch } = require('../fakes/ds-bridge-runtime.cjs');

const OLD_BLOCK = fs.readFileSync(path.join(__dirname, '../../documentserver-patches/9.3.4.60/euro-office-bridge-v1/old.txt'), 'utf8').replace(/,\s*$/, '');
const installV3 = api => { new Function('Api', `${OLD_BLOCK};`)(api); return api; };

// ---- OLD readers (5ea34bb3, verbatim apart from being plain functions) ----
function OLD_alignExt(cell) {
  const out = {};
  try {
    const inner = cell.range;
    const al = inner && typeof inner.getAlign === 'function' ? inner.getAlign(false) : null;
    if (al) {
      const ro = typeof al.getReadingOrder === 'function' ? al.getReadingOrder() : undefined;
      out.readingOrder = ro === undefined ? 0 : (typeof ro === 'number' ? ro : 'UNOBSERVABLE:readingOrder-not-number');
      const ind = typeof al.getIndent === 'function' ? al.getIndent() : undefined;
      out.indent = typeof ind === 'number' ? ind : 'UNOBSERVABLE:indent-missing';
      const sh = typeof al.getShrinkToFit === 'function' ? al.getShrinkToFit() : undefined;
      out.shrink = typeof sh === 'boolean' ? sh : 'UNOBSERVABLE:shrink-missing';
    } else {
      out.readingOrder = 'UNOBSERVABLE:getAlign-missing'; out.indent = 'UNOBSERVABLE:getAlign-missing'; out.shrink = 'UNOBSERVABLE:getAlign-missing';
    }
  } catch (err) {
    out.readingOrder = 'UNOBSERVABLE:threw'; out.indent = 'UNOBSERVABLE:threw'; out.shrink = 'UNOBSERVABLE:threw';
  }
  return out;
}
const OLD_breaksOfModel = (m, key, idMax, spanMax) => {
  if (!m || !(key in m)) return 'UNOBSERVABLE:' + key + '-absent';
  const br = m[key];
  if (br === null) return [];
  if (!br || typeof br.getBreaks !== 'function') return 'UNOBSERVABLE:' + key + '-unexpected-shape';
  const list = br.getBreaks();
  if (!Array.isArray(list)) return 'UNOBSERVABLE:' + key + '-unexpected-shape';
  const span = v => v === null || v === undefined || (Number.isInteger(v) && v >= 0 && v <= spanMax);
  const out = [];
  for (const b of list) {
    if (!b || !Number.isInteger(b.id) || b.id < 0 || b.id > idMax || typeof b.man !== 'boolean' || !span(b.min) || !span(b.max)) return 'UNOBSERVABLE:' + key + '-element-invalid';
    const min = Number.isInteger(b.min) ? b.min : null;
    const max = Number.isInteger(b.max) ? b.max : null;
    if (min !== null && max !== null && min > max) return 'UNOBSERVABLE:' + key + '-element-invalid';
    out.push({ id: b.id, man: b.man, min, max });
  }
  return out;
};
const OLD_txt = x => (x && typeof x.str === 'string' ? x.str : null);
function OLD_print(m, wide) {
  try {
    if (!m || !m.PagePrintOptions || !m.headerFooter || typeof m.headerFooter.getOddHeader !== 'function') return 'UNOBSERVABLE:model-missing';
    const ps = m.PagePrintOptions.asc_getPageSetup();
    const breaks = OLD_breaksOfModel(m, 'rowBreaks', 1048575, 16383);
    if (typeof breaks === 'string') return breaks;
    const fp = m.sheetPr ? m.sheetPr.FitToPage : false;
    const oh = m.headerFooter.getOddHeader();
    const base = { fitW: ps.asc_getFitToWidth(), fitH: ps.asc_getFitToHeight(), fitFlag: fp === true || fp === 1, oddHeader: oh && typeof oh.str === 'string' ? oh.str : null,
      rowBreaks: breaks, useFirstPage: ps.useFirstPageNumber === true, firstPage: Number.isInteger(ps.firstPageNumber) ? ps.firstPageNumber : null, hCenter: m.PagePrintOptions.asc_getHorizontalCentered() === true, scale: ps.asc_getScale(), width: ps.asc_getWidth(), height: ps.asc_getHeight() };
    if (!wide) return base;
    return { ...base, oddFooter: OLD_txt(m.headerFooter.getOddFooter()), vCenter: m.PagePrintOptions.asc_getVerticalCentered() === true,
      firstHeader: OLD_txt(m.headerFooter.getFirstHeader()), firstFooter: OLD_txt(m.headerFooter.getFirstFooter()), evenHeader: OLD_txt(m.headerFooter.getEvenHeader()), evenFooter: OLD_txt(m.headerFooter.getEvenFooter()),
      diffFirst: (v => v === true || v === 1)(m.headerFooter.getDifferentFirst()), diffOddEven: (v => v === true || v === 1)(m.headerFooter.getDifferentOddEven()) };
  } catch (err) { return 'UNOBSERVABLE:' + String(err).slice(0, 80); }
}
const OLD_dataRef = p => { try { return String(p.pivot.asc_getDataRef()); } catch (err) { return 'UNOBSERVABLE:' + String(err).slice(0, 80); } };
const bridged = old => (typeof old === 'string' ? old.replace(/^UNOBSERVABLE:/, 'UNOBSERVABLE:BRIDGE:') : old);

// ---- B05 READ ----
function alignBook(alignOf) {
  const font = { GetBold: () => null, GetItalic: () => null, GetName: () => 'Calibri', GetSize: () => 11, GetColor: () => null, GetUnderline: () => null, GetStrikethrough: () => null, GetSubscript: () => null, GetSuperscript: () => null };
  const cells = {};
  const addrs = Object.keys(alignOf);
  for (const a of addrs) {
    cells[a] = { GetAddress: () => a, GetValue: () => 'x', GetFormula: () => 'x', MergeArea: { GetAddress: () => a }, GetNumberFormat: () => 'General', GetFillColor: () => null, GetWrapText: () => false, GetOrientation: () => 0,
      GetCharacters: () => ({ GetFont: () => font }),
      range: { getAlign: alignOf[a] === null ? undefined : () => alignOf[a], getBorderFull: () => null } };
  }
  const rng = a => (cells[a] ? { ...cells[a], ForEach: fn => fn(cells[a]) } : { GetAddress: () => a, ForEach: fn => addrs.forEach(x => fn(cells[x])) });
  const ws = { GetName: () => 'S', GetIndex: () => 0, GetUsedRange: () => null, GetRange: rng };
  return { cells, Api: { GetSheets: () => [ws], GetSheet: n => (n === 'S' ? ws : undefined), WorksheetFunction: { COUNT: () => 0 } } };
}
const AL = over => ({ getAlignHorizontal: () => 0, getAlignVertical: () => 0, getReadingOrder: () => 1, getIndent: () => 2, getShrinkToFit: () => false, ...over });
const KEYS = ['readingOrder', 'indent', 'shrink'];
const newAlign = (Api, a) => {
  const c = new Function('Api', readQuery([{ sheet: 'S', range: a, style: true }], { usedRanges: false }))(Api).cells.find(x => x.a === a);
  return Object.fromEntries(KEYS.map(k => [k, c[k]]));
};

test('B05 READ parity: every proven align-ext value equals the old direct read; the readingOrder default 0 survives', () => {
  const fixtures = { A1: AL(), A2: AL({ getReadingOrder: () => undefined }), A3: AL({ getReadingOrder: undefined }), A4: AL({ getReadingOrder: () => 2, getIndent: () => 0, getShrinkToFit: () => true }), A5: AL({ getIndent: () => 15 }) };
  const { cells, Api } = alignBook(fixtures);
  installBridge(Api);
  for (const a of Object.keys(fixtures)) assert.deepEqual(newAlign(Api, a), OLD_alignExt(cells[a]), a);
});

test('B05 READ gaps: a gap the old reader had stays a gap; the stated deviation -- ONE capability, so a one-key gap now voids all three keys', () => {
  const fixtures = { B1: null, B2: AL({ getIndent: () => 'x' }), B3: AL({ getShrinkToFit: () => 1 }), B4: AL({ getReadingOrder: () => 'r' }) };
  const { cells, Api } = alignBook(fixtures);
  installBridge(Api);
  for (const a of Object.keys(fixtures)) {
    const old = OLD_alignExt(cells[a]); const now = newAlign(Api, a);
    const oldGaps = KEYS.filter(k => typeof old[k] === 'string');
    assert.ok(oldGaps.length > 0, `${a}: the fixture is a gap for the old reader`);
    assert.ok(KEYS.every(k => now[k] === now.readingOrder && /^UNOBSERVABLE:BRIDGE:/.test(now[k])), `${a}: all three unobservable, one reason`);
    assert.ok(oldGaps.some(k => bridged(old[k]) === now[k]), `${a}: the reason is the old reader's own (${old[oldGaps[0]]} vs ${now.readingOrder})`);
  }
});

// ---- B09 ----
function printModel({ ps = {}, hf = {}, ppo = {}, ...over } = {}) {
  const t = s => (s === null || s === undefined ? s : { str: s });
  const P = { asc_getFitToWidth: () => 0, asc_getFitToHeight: () => 0, asc_getScale: () => 100, asc_getWidth: () => 210, asc_getHeight: () => 297, useFirstPageNumber: false, firstPageNumber: null, ...ps };
  const H = { getOddHeader: () => null, getOddFooter: () => null, getFirstHeader: () => null, getFirstFooter: () => null, getEvenHeader: () => null, getEvenFooter: () => null, getDifferentFirst: () => null, getDifferentOddEven: () => null, ...Object.fromEntries(Object.entries(hf).map(([k, v]) => [k, typeof v === 'function' ? v : () => t(v)])) };
  return { PagePrintOptions: { asc_getPageSetup: () => P, asc_getHorizontalCentered: () => null, asc_getVerticalCentered: () => null, ...ppo }, headerFooter: H, rowBreaks: null, sheetPr: null, colBreaks: null, ...over };
}
const brk = list => ({ getBreaks: () => list });
function printBook(models) {
  const sheets = Object.entries(models).map(([n, m]) => ({ GetName: () => n, GetIndex: () => 0, GetUsedRange: () => null, worksheet: m }));
  return { GetSheets: () => sheets, GetSheet: n => sheets.find(s => s.GetName() === n), WorksheetFunction: { COUNT: () => 0 } };
}
const newPrint = (Api, wide) => new Function('Api', readQuery([], { printSheets: ['*'], printWide: wide }))(Api).printReads;
const noCol = v => (typeof v === 'string' ? v : Object.fromEntries(Object.entries(v).filter(([k]) => k !== 'colBreaks')));

test('B09 parity: the narrow and the wide print read equal the old model read, value AND key order (W8b defaults, F2 / X-P07 fields)', () => {
  const models = {
    Def: printModel(),
    Full: printModel({ ps: { asc_getFitToWidth: () => 1, asc_getFitToHeight: () => 2, asc_getScale: () => 85, asc_getWidth: () => 297, asc_getHeight: () => 420, useFirstPageNumber: true, firstPageNumber: 5 },
      hf: { getOddHeader: '&CW8', getOddFooter: '&P', getFirstHeader: 'F', getFirstFooter: '', getEvenHeader: 'E', getEvenFooter: null, getDifferentFirst: () => 1, getDifferentOddEven: () => true },
      ppo: { asc_getHorizontalCentered: () => true, asc_getVerticalCentered: () => true }, sheetPr: { FitToPage: 1 },
      rowBreaks: brk([{ id: 10, man: true, min: null, max: 16383, pt: null }, { id: 20, man: false, min: 0, max: 5 }]) }),
    Odd: printModel({ ps: { firstPageNumber: 2.5 }, hf: { getDifferentFirst: () => 0, getDifferentOddEven: () => 'true' }, ppo: { asc_getHorizontalCentered: () => 1 }, sheetPr: { FitToPage: true } }),
  };
  const Api = printBook(models);
  installBridge(Api);
  for (const wide of [false, true]) {
    const r = newPrint(Api, wide);
    for (const n of Object.keys(models)) {
      const old = OLD_print(models[n], wide);
      assert.deepEqual(noCol(r[n]), old, `${n} wide=${wide}`);
      assert.deepEqual(Object.keys(noCol(r[n])), Object.keys(old), `${n} wide=${wide}: key order`);
      assert.equal('colBreaks' in r[n], wide, 'the column breaks ride with the wide read only (unchanged)');
    }
  }
});

test('B09 gaps: every old row-break gap is the same reason through the bridge; the stated deviations are stricter, never a value', () => {
  const same = {
    Abs: (() => { const m = printModel(); delete m.rowBreaks; return m; })(),
    Shape: printModel({ rowBreaks: {} }), NotList: printModel({ rowBreaks: brk(null) }),
    Bad: printModel({ rowBreaks: brk([{ id: 3 }]) }), Span: printModel({ rowBreaks: brk([{ id: 3, man: true, min: 5, max: 4 }]) }),
    NoModel: printModel({ PagePrintOptions: null }),
  };
  const Api = printBook(same);
  installBridge(Api);
  const r = newPrint(Api, true);
  for (const n of Object.keys(same)) assert.equal(r[n], bridged(OLD_print(same[n], true)), n);
  // deviation 1 (design 3.2): ANY missing getter the capability reads -> model-missing (the old reader checked getOddHeader
  // only, and a narrow read could still return a value without the wide getters)
  const thin = { Thin: printModel({ hf: { getEvenFooter: undefined } }) };
  delete thin.Thin.headerFooter.getEvenFooter;
  const t = printBook(thin); installBridge(t);
  assert.equal(typeof OLD_print(thin.Thin, false), 'object', 'the old narrow read answered');
  assert.equal(newPrint(t, false).Thin, 'UNOBSERVABLE:BRIDGE:model-missing', 'the bridge refuses the whole read');
  // deviation 2: a throwing getBreaks is the capability's reason, not the raw error text
  const thr = { Thr: printModel({ rowBreaks: { getBreaks: () => { throw new Error('boom'); } } }) };
  const a = printBook(thr); installBridge(a);
  assert.match(OLD_print(thr.Thr, true), /^UNOBSERVABLE:Error: boom/);
  assert.equal(newPrint(a, true).Thr, 'UNOBSERVABLE:BRIDGE:rowBreaks-unexpected-shape');
});

// ---- B13 ----
function pivotBook(pivotsBySheet) {
  const grid = v => ({ GetAddress: () => 'A1:B2', GetValue: () => v });
  const sheets = Object.entries(pivotsBySheet).map(([n, specs]) => {
    const pivots = specs.map(([name, ref, srcThrows]) => ({
      GetName: () => name,
      GetSource: () => { if (srcThrows) throw new Error('no source after reopen'); return { GetAddress: () => `${n}!$A$1:$B$2` }; },
      GetTableRange1: () => grid([['a', 'b']]), GetDataBodyRange: () => grid([['1']]),
      GetPivotFields: () => [{ GetName: () => 'a' }], GetRowFields: () => [], GetColumnFields: () => [], GetDataFields: () => [],
      pivot: ref === undefined ? undefined : { asc_getDataRef: () => ref } }));
    return { GetName: () => n, GetIndex: () => 0, GetUsedRange: () => null, GetAllPivotTables: () => pivots, GetRange: () => grid([['k', 'v']]) };
  });
  return { GetSheets: () => sheets, GetSheet: n => sheets.find(s => s.GetName() === n), WorksheetFunction: { COUNT: () => 0 } };
}
const newPivots = Api => new Function('Api', readQuery([], { objectSheets: ['*'] }))(Api).objects;

test('B13 parity: the pivot data reference (pivotRefs) and the content-proof input equal the old direct read', () => {
  const spec = { H: [['P1', 'Data!$A$1:$B$2'], ['P2', "'My Sheet'!$A$1:$B$2", true]], Data: [], 'My Sheet': [] };
  const Api = pivotBook(spec);
  installBridge(Api);
  const o = newPivots(Api);
  const pv = Api.GetSheet('H').GetAllPivotTables();
  assert.deepEqual(o.H.pivots.map(p => p.dataRef), pv.map(OLD_dataRef));
  const proof = o.H.pivots[1].proof;
  assert.equal(proof.dataRef, OLD_dataRef(pv[1]), 'the proof is built on the same reference');
  assert.deepEqual([proof.source.sheet, proof.source.range], ['My Sheet', 'A1:B2']);
  assert.equal(o.H.pivots[0].proof, undefined, 'a readable public source needs no proof (unchanged)');
});

test('B13 gaps and the stated deviations: model gap stays a gap; a non-string reference and a duplicate name refuse instead of coercing / picking', () => {
  const Api = pivotBook({ H: [['M', undefined], ['N', 7], ['D', 'X!$A$1:$B$2'], ['D', 'Y!$A$1:$B$2']] });
  installBridge(Api);
  const [m, n, d1, d2] = newPivots(Api).H.pivots;
  const pv = Api.GetSheet('H').GetAllPivotTables();
  assert.match(OLD_dataRef(pv[0]), /^UNOBSERVABLE:/);
  assert.equal(m.dataRef, 'UNOBSERVABLE:BRIDGE:pivot-model-missing');
  assert.equal(OLD_dataRef(pv[1]), '7', 'the old read coerced');
  assert.equal(n.dataRef, 'UNOBSERVABLE:BRIDGE:dataRef-not-string');
  assert.deepEqual([d1.dataRef, d2.dataRef], ['UNOBSERVABLE:BRIDGE:pivot-name-ambiguous', 'UNOBSERVABLE:BRIDGE:pivot-name-ambiguous'], 'never one picked for the other');
});

// ---- the /3 runtime ----
test('a /3 runtime refuses every migrated axis -- the shipped old.txt block AND the /4 block reporting /3; no direct-read fallback', () => {
  const { cells, Api: aApi } = alignBook({ A1: AL() });
  const pApi = printBook({ D: printModel() });
  const vApi = pivotBook({ H: [['P', 'Data!$A$1:$B$2', true]], Data: [] });
  for (const [label, wrap] of [['old.txt /3', api => installV3(api)], ['/4 reporting /3', api => { installBridge(api); return withPatch(api); }]]) {
    const a = newAlign(wrap({ ...aApi }), 'A1');
    for (const k of KEYS) assert.equal(a[k], 'UNOBSERVABLE:PATCH_VERSION_MISMATCH', `${label} ${k}`);
    assert.equal(newPrint(wrap({ ...pApi }), true).D, 'UNOBSERVABLE:PATCH_VERSION_MISMATCH', `${label} print`);
    const pv = newPivots(wrap({ ...vApi })).H.pivots[0];
    assert.equal(pv.dataRef, 'UNOBSERVABLE:PATCH_VERSION_MISMATCH', `${label} pivot`);
    assert.equal(pv.proof, 'UNOBSERVABLE:PATCH_VERSION_MISMATCH', `${label} proof`);
  }
  assert.ok(cells.A1, 'fixture sanity');
});
