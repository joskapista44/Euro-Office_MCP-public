'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const { readQuery } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');

/**
 * The XLSX driver's LIVE read (M6). A read must never mutate: the generated body may call only
 * getters, ForEach and WorksheetFunction.COUNT -- for any window list, including hostile names.
 */
const MUTATING = /\.(Add\w*|Delete|Remove\w*|Insert\w*|Set\w*|Merge|UnMerge|Split\w*|Paste\w*|Replace\w*|Move\w*|Apply\w*|Clear\w*|Copy\w*|AutoFit|Sort|SetValue|ShowAllData)\s*\(/;
const ALLOWED = /^(Get\w*|ForEach|COUNT)$/;
const QUERIES = [readQuery(), readQuery([{ sheet: 'Data', range: 'A1:B2' }]), readQuery([{ sheet: "x'); ws.GetRange('A1').SetValue(1); ('", range: 'A1' }]),
  readQuery([], { formatSheets: ['Data'], objectSheets: ['Data', "x'); Api.GetSheet('Data').Delete(); ('"] }),
  readQuery([], { layout: [{ sheet: "x'); Api.GetSheet('Data').Delete(); ('", maxCol: 3, maxRow: 2 }] }),
  readQuery([], { panes: true, internalPanes: true }),
  readQuery([], { validationRanges: [{ sheet: "x'); Api.GetSheet('Data').GetRange('A1').GetValidation().Delete(); ('", range: 'A1' }] }),
  readQuery([], { cfRanges: [{ sheet: "x'); Api.GetSheet('Data').GetRange('A1').GetFormatConditions().Delete(); ('", range: 'A1' }] }),
  readQuery([], { filterSheets: ["x'); Api.GetSheet('Data').GetAutoFilter().ShowAllData(); ('"], internalFilter: true }),
  readQuery([], { chartSheets: ['*', "x'); Api.GetSheet('Data').GetAllCharts()[0].SetName('z'); ('"] }),
  readQuery([], { pageSheets: ['*', "x'); Api.GetSheet('Data').SetPageOrientation('xlLandscape'); ('"], pagePerSheet: true }),
  readQuery([], { printSheets: ['*', "x'); Api.GetSheet('Data').worksheet.PagePrintOptions.asc_setHorizontalCentered(true); ('"] })];

test('the read query calls no mutating editor method, for every window shape', () => {
  for (const q of QUERIES) {
    const code = q.replace(/const (windows|formatSheets|objectSheetsAsked|layoutAsked|validationAsked|cfAsked|filterAsked|chartAsked|pageAsked|printAsked) = \[[\s\S]*?\];/g, 'const $1 = [];'); // the lists are DATA (JSON), not calls
    assert.doesNotMatch(code, MUTATING);
    const calls = [...code.matchAll(/\.([A-Z]\w*)\s*\(/g)].map(m => m[1]);
    assert.ok(calls.length >= 6, 'positive control: the extractor sees the query calls');
    for (const c of calls) assert.match(c, ALLOWED, `unexpected editor call in a read: ${c}`);
  }
});

test('a hostile sheet name stays inside the JSON window list (it is data, never code)', () => {
  const q = QUERIES[2];
  const list = q.match(/const windows = (\[[\s\S]*?\]);/)[1];
  assert.equal(JSON.parse(list)[0].sheet, "x'); ws.GetRange('A1').SetValue(1); ('");
});

test('NEGATIVE CONTROL: the pattern catches a mutation smuggled into a read', () => {
  assert.match("ws.GetRange('A1').SetValue(1);", MUTATING);
  assert.match('r.Merge(false);', MUTATING);
  assert.doesNotMatch('Delete', ALLOWED);
});

/**
 * The generated body is JavaScript the editor evaluates; a syntax error would surface only LIVE. It is
 * compiled here and run against a minimal in-memory stand-in for the spreadsheet Api (shape only --
 * the real getter behaviour is the wave-0 probe's job).
 */
function fakeApi() {
  const { parseRange, cellsOf } = require('../../src/adapters/xlsx/a1.cjs');
  const data = { Data: { A1: 'Name', B1: 2, C1: '=B1*2' } };
  const rangeOf = (sheet, a1) => {
    const g = parseRange(a1);
    return {
      GetAddress: () => a1,
      ForEach: fn => { for (const a of cellsOf(g)) fn(cellOf(sheet, a)); },
    };
  };
  const cellOf = (sheet, a) => {
    const raw = data[sheet][a];
    const f = typeof raw === 'string' && raw.startsWith('=') ? raw : '';
    // the MEASURED shapes (wave-0 probe, internal tracker): GetValue is a string for numbers too, and
    // GetFormula returns the value text for a non-formula cell
    const v = f ? '4' : raw === undefined ? '' : String(raw);
    return { GetAddress: () => a, GetValue: () => v, GetFormula: () => f || v, MergeArea: { GetAddress: () => a }, _num: typeof raw === 'number',
      GetNumberFormat: () => (a === 'B1' ? '0.00' : 'General'), GetFillColor: () => (a === 'B1' ? { GetRGB: () => 16776960 } : 'none') };
  };
  const ws = name => ({ GetName: () => name, GetIndex: () => 0, GetUsedRange: () => rangeOf(name, 'A1:C1'), GetRange: a1 => rangeOf(name, a1) });
  return {
    GetSheets: () => [ws('Data')], GetSheet: n => (data[n] ? ws(n) : null),
    WorksheetFunction: { COUNT: cell => (cell._num ? 1 : 0) },
  };
}

test('the read body compiles and returns the payload shape on a stand-in Api', () => {
  const run = new Function('Api', readQuery([{ sheet: 'Data', range: 'D1:D2' }, { sheet: 'Nope', range: 'A1' }]));
  const out = run(fakeApi());
  assert.equal(out.ok, true);
  assert.deepEqual(out.sheets, [{ name: 'Data', index: 0, used: 'A1:C1' }]);
  assert.deepEqual(out.cells.map(c => c.a), ['A1', 'B1', 'C1', 'D1', 'D2'], 'used range + window, a missing sheet window skipped');
  assert.deepEqual(out.cells.find(c => c.a === 'B1'), { sheet: 'Data', a: 'B1', v: '2', f: '2', count: 1, merge: 'B1' });
  const { normalizeInventory } = require('../../src/adapters/xlsx/inventory.cjs');
  const inv = normalizeInventory(out);
  assert.deepEqual(inv.cells['Data\u0000B1'], { v: '2', f: '', count: 1, merge: 'B1' }, 'normalization: a value-text GetFormula is not a formula');
  assert.equal(inv.cells['Data\u0000C1'].f, '=B1*2');
  assert.equal(inv.complete, true);
  assert.equal(out.cells.find(c => c.a === 'C1').f, '=B1*2');
  assert.equal(out.truncated, false);
  assert.deepEqual(out.unsupported, []);
});

test('format is read only for window cells that ask for it -- also when the cell was already read in the used range', () => {
  const out = new Function('Api', readQuery([{ sheet: 'Data', range: 'B1:D1', format: true }]))(fakeApi());
  assert.equal(out.ok, true);
  const by = a => out.cells.find(c => c.a === a);
  assert.equal(out.cells.filter(c => c.a === 'B1').length, 1, 'no duplicate: the used-range entry is augmented');
  assert.deepEqual([by('B1').nf, by('B1').fill], ['0.00', 16776960]);
  assert.deepEqual([by('D1').nf, by('D1').fill], ['General', 'S:none'], 'a window cell outside the used range');
  assert.equal('nf' in by('A1'), false, 'a used-range cell outside the window carries no format');
  const { normalizeInventory } = require('../../src/adapters/xlsx/inventory.cjs');
  const inv = normalizeInventory(out);
  assert.deepEqual(inv.cells['Data\u0000B1'], { v: '2', f: '', count: 1, merge: 'B1', nf: '0.00', fill: 16776960, bold: 'UNOBSERVABLE:BRIDGE_UNAVAILABLE', italic: 'UNOBSERVABLE:BRIDGE_UNAVAILABLE' }, 'no DS Bridge -> bold/italic unobservable, never false (B01 / B02 cutover)');
  assert.equal('nf' in inv.cells['Data\u0000A1'], false);
});

// B01 / B02 DS Bridge cutover (owner GO, internal tracker point 8; LIVE parity #ref / #ref): cell-level bold and italic
// are read ONLY through the DS Bridge (`textStyle.bold.read/v1` / `textStyle.italic.read/v1`). These EXECUTE the generated
// code against the REAL served bridge block (test/fakes/ds-bridge-runtime.cjs), so every gap is the bridge's own reason.
test('DS Bridge bold / italic: the value through the bridge, every gap a reason string -- never false, no direct fallback', () => {
  const { installBridge } = require('../fakes/ds-bridge-runtime.cjs');
  // the bridge reads ws.GetRange(a1).range; the fake gives every range (and every ForEach cell) the same model range;
  // `patchVersion` overrides the served manifest's patch identity (an unaccepted runtime, owner decision (B))
  const withRange = (inner, bridge = true, patchVersion = null) => { const api = fakeApi(); const orig = api.GetSheet; api.GetSheet = n => { const ws = orig(n); if (!ws) return ws; const gr = ws.GetRange; ws.GetRange = a1 => { const r = gr(a1); const fe = r.ForEach; r.ForEach = fn => fe(cell => fn({ ...cell, range: inner })); r.range = inner; return r; }; return ws; }; if (bridge) installBridge(api); if (patchVersion) { const real = api.GetEuroOfficeBridge; api.GetEuroOfficeBridge = () => { const b = real(); const m = b.manifest(); return { manifest: () => ({ ...m, patchVersion }), read: b.read }; }; } return api; };
  const read = (api, key) => new Function('Api', readQuery([{ sheet: 'Data', range: 'D1', format: true }], {}))(api).cells.find(c => c.a === 'D1')[key];
  const both = { getFont: () => ({ getBold: () => true, getItalic: () => false }) };
  assert.equal(read(withRange(both), 'bold'), true);
  assert.equal(read(withRange(both), 'italic'), false, 'italic reads its own capability, not bold');
  assert.equal(read(withRange({ getFont: () => ({ getBold: () => false, getItalic: () => true }) }), 'bold'), false);
  assert.equal(read(withRange({ getFont: () => ({ getBold: () => false, getItalic: () => true }) }), 'italic'), true);
  for (const key of ['bold', 'italic']) {
    const flag = key === 'bold' ? 'Bold' : 'Italic';
    assert.equal(read(withRange(both, false), key), 'UNOBSERVABLE:BRIDGE_UNAVAILABLE', `${key}: no bridge -> never read directly`);
    assert.equal(read(withRange(both, true, 'euro-office-bridge-v1/9.3.4.60/2'), key), 'UNOBSERVABLE:PATCH_VERSION_MISMATCH', `${key}: a served runtime that is not the accepted one`);
    assert.equal(read(withRange({}), key), 'UNOBSERVABLE:BRIDGE:font-unavailable');
    assert.equal(read(withRange({ getFont: () => ({}) }), key), `UNOBSERVABLE:BRIDGE:${key}-not-boolean`);
    assert.equal(read(withRange({ getFont: () => ({ [`get${flag}`]: () => null }) }), key), `UNOBSERVABLE:BRIDGE:${key}-not-boolean`, 'null is never false');
    assert.equal(read(withRange({ getFont: () => { throw new Error('x'); } }), key), 'UNOBSERVABLE:BRIDGE:internal-threw');
    assert.equal(read(withRange(undefined), key), 'UNOBSERVABLE:BRIDGE:range-missing');
  }
  const q = readQuery([{ sheet: 'Data', range: 'D1', format: true }], {});
  assert.doesNotMatch(q, /getBold|getItalic|getFont/, 'the direct internal font read is gone from the generated query');
  assert.equal((q.match(/bridgeFontFlag\(out, /g) || []).length, 2, 'exactly the two bridge reads');
});

test('the style read calls no direct internal getter or setter: W4c alignment / borders (B04) and align-ext (B05 READ) go through the DS Bridge', () => {
  const code = readQuery([{ sheet: 'Data', range: 'A1', format: true, style: true }]);
  const dotCalls = [...code.matchAll(/\.(get|set)([A-Z]\w*)\s*\(/g)].map(m => m[1] + m[2]);
  // X-F01 v2: bold/italic now dispatch through one shared readInternalFontFlag(cell,
  // enabled, getter, key) generated helper -- the getter NAME travels as a literal string argument
  // ('getBold' / 'getItalic'), not a `.getXxx(` dot-call, so it is picked up here too; still only
  // ever a get*, never a set* (dynamic or dot-call).
  const dynamicNames = [...code.matchAll(/'(get|set)([A-Z]\w*)'/g)].map(m => m[1] + m[2]);
  const internal = [...dotCalls, ...dynamicNames];
  // B05 READ cutover (Wave B READ /4): getAlign / getIndent / getReadingOrder / getShrinkToFit are gone -- the three axes
  // come from range.alignmentExt.read/v1; no lower-case internal call is left in the style read.
  assert.deepEqual([...new Set(internal)].sort(), []);
  assert.match(code, /"range\.alignmentExt\.read\/v1"/, 'positive control: the bridge align-ext read is in the query');
  assert.match(code, /"range\.appearance\.read\/v1"/);
  assert.doesNotMatch(code, /\.set[A-Z]\w*\s*\(/, 'no lower-case internal setter anywhere in a read');
  assert.doesNotMatch(code, /'set[A-Z]\w*'/, 'no dynamically-dispatched internal setter name either');
});

test('copy_sheet (M11.5c.4 Bridge WRITE): the adapter and the driver gate on the SAME /8-only id; the internal-copy kind is gone; the driver refuses any other runtime without touching the editor', async () => {
  const { B03_BRIDGE_WRITE_RUNTIMES, createEuroOfficeXlsxDriver } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
  const sheets = require('../../src/adapters/xlsx/families/sheets.cjs');
  assert.deepEqual([...sheets.B03_BRIDGE_WRITE_RUNTIMES], [...B03_BRIDGE_WRITE_RUNTIMES]);
  assert.deepEqual([...B03_BRIDGE_WRITE_RUNTIMES], ['9.3.4|euro-office-bridge-v1/9.3.4.60/8'], 'the explicit manifest runtime id (/8 only), never a cache key');
  const d = createEuroOfficeXlsxDriver({ url: 'u', user: 'u', pass: 'p', loadPlaywright: () => ({ ok: false }) });
  const lowered = sheets.lower({ intent: 'copy_sheet', selector: { sheet: 'A', name: 'B' } });
  assert.equal(lowered.kind, 'callCommand'); assert.equal(lowered.meta.bridgeCopyWrite, true);
  const r = await d.dispatch(null, lowered);
  assert.deepEqual([r.ok, r.dispatched, r.outcome], [false, false, 'b03-bridge-write-runtime-not-proven'], 'no editor was touched');
  const unmarked = await d.dispatch(null, { ...lowered, meta: { ...lowered.meta, bridgeCopyWrite: false } });
  assert.deepEqual([unmarked.ok, unmarked.dispatched, unmarked.outcome], [false, false, 'b03-bridge-write-runtime-not-proven'], 'a copy_sheet command without the Bridge mark is refused too');
  const old = await d.dispatch(null, { kind: 'internalCopySheet', sheet: 'A', name: 'B' });
  assert.deepEqual([old.ok, old.dispatched], [false, false], 'the former internal-copy kind is not served');
});

test('the own-pane read (freeze, wave 4) never activates a sheet and adds no internal getter: the pane comes through the DS Bridge (B06)', () => {
  const code = readQuery([], { panes: true });
  assert.doesNotMatch(code, /SetActive/, 'a read that activates a sheet changes the view state it is measuring');
  const internal = q => new Set([...q.matchAll(/\.(get|set)([A-Z]\w*)\s*\(/g)].map(m => m[1] + m[2]));
  const base = internal(readQuery());
  assert.deepEqual([...internal(code)].filter(x => !base.has(x)).sort(), [], 'the pane block adds no direct internal call');
  assert.doesNotMatch(code, /topLeftFrozenCell|getSheetView/, 'the direct pane read is gone');
  assert.match(code, /sheet\.freezeState\.read\/v1/, 'positive control: the bridge pane read is in the query');
  assert.doesNotMatch(readQuery(), /freezeState|GetActiveSheet/, 'nothing of it unless asked (a readLive stub keys on GetActiveSheet)');
});

test('the print read (wave 8) is generated only when asked and goes through the DS Bridge only (B09 cutover): no direct model getter, no setter', () => {
  assert.doesNotMatch(readQuery(), /"sheet\.printSettings\.read\/v1"|getOddHeader|getBreaks|asc_getFitToWidth/, 'no print read unless asked');
  const getters = code => new Set([...code.matchAll(/\.(get|asc_get)([A-Z]\w*)\s*\(/g)].map(m => m[1] + m[2]));
  for (const code of [readQuery([], { printSheets: ['*'] }), readQuery([], { printSheets: ['*'], printWide: true, colBreaks: true })]) {
    const base = getters(readQuery());
    assert.deepEqual([...getters(code)].filter(g => !base.has(g)).sort(), [], 'the print read adds no direct internal getter');
    assert.doesNotMatch(code, /PagePrintOptions|headerFooter|\.worksheet\b|rowBreaks\.getBreaks/, 'the direct model read is gone, no fallback');
    assert.match(code, /"sheet\.printSettings\.read\/v1"/, 'positive control: the bridge print read is in the query');
    assert.doesNotMatch(code, /\.(set|asc_set)[A-Z]\w*\s*\(/, 'no internal setter anywhere in a read');
  }
});
