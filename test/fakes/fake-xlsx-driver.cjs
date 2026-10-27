'use strict';

const { XLSX_RUNTIME_ID } = require('../../src/core/runtime-identity.cjs');
const { translateFormula } = require('../../src/adapters/xlsx/shift.cjs');
const { parseRange: parseA1Range, addr: a1Addr } = require('../../src/adapters/xlsx/a1.cjs');
/**
 * A fake XLSX editor transport (M6). Same DURABLE / LIVE split as the fake DOCX transport:
 *   open        LIVE := copy of DURABLE
 *   dispatch    changes LIVE only (it applies the command's structured `meta`; it cannot run `body`)
 *   forceSave   DURABLE := copy of LIVE   (canSave then reports false)
 *   close       LIVE is dropped; an unsaved change does NOT reach DURABLE
 *
 * readLive returns the payload shape the real XLSX driver produces: the sheet list with used ranges,
 * every cell of every used range plus the requested windows, v/f/count/merge per cell.
 *
 * Scripted faults act at the transport boundary: fault 'ignore' | 'partial' | 'side-effect' | 'throw' |
 * 'clear-resets-format' (the measured ApiRange.Clear() behaviour: value AND number format / fill / bold reset) |
 * 'clear-drops-bold' (only the cell bold is lost); 'side-erase' (a clear also erases script.eraseCells); boldUnobservable: 'always' | 'after-dispatch';
 * readTruncated, countUnreadable (count null on every cell). Sheet operations (wave 3) follow the measured
 * probe: a rename rewrites the references to the sheet, a delete turns them into #REF!; faults
 * 'create-default-name' (the measured silent 'Sheet2'), 'rename-no-follow' (references left behind),
 * 'create-reorders' (the other sheets' order changes), 'rename-dv-no-follow' / 'rename-name-no-follow' (a validation /
 * a defined name left on the old sheet name). copy_sheet (M11.5c.4: ONE worksheet.copy/v1 Bridge WRITE, meta.bridgeCopyWrite;
 * script.b03Write 'gate-refused' | 'error-applied' (the copy lands, the receipt says copy-error)):
 * a sheet's `objects` ({ charts, pivots, names }) and `rules` are copied with self-references rewritten;
 * faults 'copy-duplicate' (a second sheet of the name), 'copy-keeps-self-reference', 'copy-drops-object',
 * 'copy-drops-rule', 'copy-drops-cell', 'copy-timeout' (nothing happens, ok false); script.runtime (default the proven runtime identity), script.rules (a
 * complete rule observation is served only when true).
 *
 * Merge / unmerge (wave 4, meta.family 'merge') follow the W4b LIVE probe: a merge absorbs every merge it
 * intersects and SILENTLY empties every non-top-left cell; an unmerge dissolves every merge it intersects and
 * keeps the content. Faults 'ignore', 'partial' (only the first row / the first intersecting merge),
 * 'side-effect' (Z99 written), 'merge-drops-top-left' (the top-left emptied too), 'side-merge' (script.sideMerge
 * merged as well), 'merge-lands-elsewhere' (script.elsewhere merged instead of the target).
 *
 * Layout (wave 4, meta.family 'layout'): a sheet may carry layout { cols: { B: { width, hidden } }, rows: { '2': {
 * height, hidden } } }; defaults width 8.43 / height null (the measured unset row) / hidden false; a hidden column
 * reads width 0, a hidden row height null (W4b). Read only when the request names `layout`, for every sheet over
 * its used range and the requested extent. Faults 'ignore', 'partial' (the first column / row only),
 * 'side-effect', 'side-layout' (script.sideColumn gets width 33), 'layout-drifts' (the value lands 0.5 off);
 * script.layoutUnobservable (every sheet's axis a reason string).
 *
 * Style (wave 4, meta.family 'format'): a cell may carry wrap, font { bold, italic, name, size, color }, align { h, v },
 * border { t, b, l, r, d: { s, c }, dd, du } (its OWN sides). Read only for windows with `style: true` and the
 * `styleSheets` used ranges, in the W4a / W4a2 / W4c shapes: defaults wrap false, font { null, null, 'Calibri', 11,
 * 0 }, align { h: null, v: 0 }, border sides { s: 0, c: 0 }; an EMPTY cell's font reads the SAME as a non-empty
 * cell's (X-F02, m7-s31b-probe1/2 MEASURED: the public reader at (1,1) observes it too) --
 * `script.emptyFontUnobservable` [cells] opts a specific empty cell into 'UNOBSERVABLE:threw' for tests that
 * target that gap directly; a side with no own border reads the neighbour's facing side (the measured shared
 * edge). script.appearanceBuild false -> align / border 'UNOBSERVABLE:runtime-not-verified'; script.mixedRuns
 * [cells] -> font 'UNOBSERVABLE:mixed-runs' (non-empty cells only);
 * script.blackReadsNull -> a black border colour reads null (the measured after-reopen shape). Faults 'ignore',
 * 'partial' (the first cell only), 'side-effect', 'format-leaks' (script.leakCell gets the same spec),
 * 'format-extra-key' (the targets also get number format '0.00').
 *
 * Freeze (wave 4, meta.family 'freeze'): a sheet may carry pane { row0, col0 }; the store keeps `active` (default the
 * first sheet). Read only when the request names `panes`: every sheet's own pane ({ state: 'frozen', row0, col0 } |
 * { state: 'none' }) and the active sheet's name. The dispatch follows the command: activate, freeze, restore.
 * script.paneBuild false -> every pane 'UNOBSERVABLE:runtime-not-verified'. Faults 'ignore', 'wrong-sheet' (the
 * freeze lands on the ACTIVE sheet, the measured W4b2 trap), 'no-restore' (the target stays active),
 * 'side-effect' (another sheet's pane changes), 'freeze-off-by-one' (row0 + 1).
 *
 * Structure (wave 5, meta.family 'structure'): insert / delete rows / columns on the named sheet -- cells move,
 * formulas on EVERY sheet re-pointed (the reference-aware shift), merges moved, a deleted block's cells gone.
 * Faults 'ignore', 'partial' (one row / column fewer), 'no-reference-follow' (formulas keep their text),
 * 'side-effect', 'other-sheet-moves' (a second sheet's cells shift too).
 *
 * Range (wave 5, meta.family 'range'): copy_range copies each source cell (a formula translated by the delta,
 * the reference-aware translation) with its format; move_range moves it unchanged and clears the source. Faults
 * 'ignore', 'partial' (the first cell only), 'copy-no-translate', 'move-keeps-source', 'side-effect',
 * 'copy-corrupts-source' (a correct copy, then the source's first cell is overwritten); L-06 'copy-saves-literal' (the DS
 * Copy shows the TRANSLATED formula in session but the SAVED file carries the source's LITERAL text and stale value --
 * a maintainer's l064/l065 probe; 'cross-copy-delayed': an editor Copy onto another sheet is overwritten at the next save even after an explicit formula write (l068); 'copy-xsheet-shows-literal': onto another sheet the formula is shown untranslated, l066; a command whose meta.explicitFormulaWrite is true re-wrote the target formulas, so they save as shown); M7 S1 L3 controls
 * 'range-side-name' (a workbook name appears), 'range-side-rule' (the first rule of the first sheet with rules
 * changes), 'range-side-chart' (a chart series' values move a row), 'range-side-filter' (an AutoFilter appears/changes).
 *
 * Sort (wave 6, meta.family 'sort'): a stable sort of the data rows by the key (numbers numerically, text by code
 * point), rows moved whole, blank keys last, a formula re-pointed to its new row (M7 S5 R21). Faults 'ignore',
 * 'sort-key-only' (only the key column moves -- rows broken), 'sort-reverse' (the opposite order), 'side-effect',
 * 'sort-formula-stays' (a moved formula keeps its old row's references), 'sort-blanks-first' (blank keys first).
 *
 * Validation (wave 6, meta.family 'validation'): a validation rule is a sheet rule { kind: 'validation', range, formula,
 * formula2, api: { type, alertStyle, operator } }; the internal reader lists it, the public range read
 * (`validationRanges`) answers from the rule overlapping the range (absent when none). Faults 'ignore',
 * 'validation-wrong-bounds', 'validation-leaks' (a second rule on script.leakRange), 'side-effect';
 * script.validationReadUnobservable.
 *
 * Conditional formats (wave 6, meta.family 'cf'): a sheet rule { kind: 'conditional', type, operator, formula, fill,
 * range }; the range read (`cfRanges`) lists every rule overlapping the range. Faults 'ignore', 'cf-wrong-fill',
 * 'cf-removes-other' (every rule on the range goes), 'cf-alters-other' (another rule's formula changes), 'cf-duplicate' (the rule added twice), 'side-effect';
 * script.cfReadUnobservable.
 *
 * Charts (wave 7, meta.family 'charts'): objects.charts records may carry name and title; read for `chartSheets`
 * ('*' = every sheet) as { charts: [{ name, type, title, series: [{ name, values, cat }] }], pivots: [names] }.
 * script.chartsUnobservable -> every sheet 'UNOBSERVABLE:threw'. Faults 'ignore', 'chart-wrong-title',
 * 'chart-wrong-range' (the values one row short), 'chart-duplicate' (added twice), 'chart-alters-other' (another
 * chart's title changes), 'rename-drops-series' (the renamed chart loses its last series), 'chart-extra-unnamed' (a
 * second, unnamed chart too), 'chart-duplicate-drops-neighbour' (added twice, another chart of the sheet removed:
 * the count still +1), 'rename-copies' (the old name stays, a copy gets the new one), 'side-effect'.
 *
 * Page layout (wave 8, meta.family 'page'): a sheet may carry page { orientation, printGridlines, printHeadings }
 * (default xlPortrait / false / false); read for `pageSheets` ('*' = every sheet). script.pageBuild false ->
 * 'UNOBSERVABLE:per-sheet-read-not-verified' (the fake otherwise stands for a proven build); script.pageFieldMissing
 * (printHeadings read as undefined -- a getter that returned nothing), script.pageMarginMissing (leftMargin the same).
 * Faults 'ignore',
 * 'page-partial' (only the first requested field), 'page-side' (every unrequested field changes), 'page-margin-drift'
 * (a margin lands 0.5 mm off), 'page-other-sheet'
 * (another sheet's orientation changes), 'side-effect'.
 *
 * AutoFilter (wave 6, meta.family 'filter'): a sheet may carry autoFilter { range, filters: [{ field, on, operator,
 * criteria1, criteria2 }] }; read for `filterSheets`. script.filterBuild false -> 'UNOBSERVABLE:field-build-not-
 * verified' (the fake otherwise stands for a proven build). Faults 'ignore', 'filter-wrong-field', 'filter-wrong-value',
 * 'filter-extra-field' (a second active filter), 'filter-drops-other' (every other field's filter is dropped; M7
 * R19), 'filter-wrong-operator' (xlAnd <-> xlOr; M7 R19), 'side-effect'.
 *
 * A cell may carry nf (number format, default 'General') and fill (default 'S:No Fill', the measured read of an unfilled cell, m7-s15-probe3); they are read only
 * for window cells requested with `format: true`, as the real driver does.
 *
 * Array formulas (GH #301 D4, meta.family 'array-formula'): a member cell carries `array` (the array's A1 ref) and `af`
 * (its formula as GetFormulaArray returns it) next to its own f. Read only via `arrayCells` ({sheet, cell}), keyed
 * `sheet!cell`, with the LIVE-measured raw GetFormulaArray semantics (docs/evidence/gh301-d4-array-formula-probe-
 * 20261025.json): a member -> its array's formula, a plain formula -> null, a value -> the value, an empty cell -> ''.
 * The write overwrites every area cell (the measured SetFormulaArray does, silently). Faults 'ignore', 'side-effect',
 * 'array-formula-wrong-formula' (another formula is written -- L1 FAIL), 'array-formula-short' (the area's last cell is
 * left out -- L1 FAIL), 'array-formula-spill-ring' (the array grows one row past the area -- the ring guard's L1 FAIL);
 * script.arrayUnobservable -> every read 'UNOBSERVABLE:threw'.
 *
 * Per-run font (M7 backlog X-F04, meta.family 'format_run'): a cell may
 * carry `runs` ([{start, length, font: {bold, italic, name, size, color}}], 1-based, contiguous,
 * covering the whole text) -- absent means uniformly described by the cell's own `font` (the
 * SAME default tuple the 'format' family's font reader uses, minus the X-F05 effects, which are
 * out of this family's scope). Read only via `runCells` ({sheet, cell}), keyed `sheet!cell`, as
 * the character-by-character run-map the real driver's PUBLIC reader also builds; 'UNOBSERVABLE:
 * empty-cell' / 'UNOBSERVABLE:text-too-long' (256-char ceiling, TEMPORARY pending a LIVE timing
 * measurement) mirror the real driver's own named gaps. Faults 'ignore', 'side-effect',
 * 'format-run-wrong-axis' (the written bold value is flipped -- L1 FAIL), 'format-run-other-
 * char-changed' (an UNTOUCHED character's font flips, `script.faultCharIndex` names which one --
 * L3 FAIL), 'format-run-text-changed' (the cell's text grows -- L3 `run-map-length-changed`);
 * script.runMapUnobservable -> every run-map read 'UNOBSERVABLE:threw'.
 */

const { parseRange, cellsOf, addr, rangeText } = require('../../src/adapters/xlsx/a1.cjs');

const clone = x => JSON.parse(JSON.stringify(x));

// M7 P9: a pivot record with `proofSpec` ({ dataRef, fields, rows, cols, data, body: { range, values }, table:
// { range, values } }) stands for a REOPENED pivot: its public source reads 'UNOBSERVABLE:...(reading 'ws')' (W3g /
// P-A) and the driver's raw proof is built from the LIVE store -- the named range's values and every other sheet's
// values at the same address, as the real driver reads them. script.pivotProofUnobservable -> the proof is a reason.
const gridOf = (sheet, range) => { const g = parseRange(range); if (!sheet || !g) return 'UNOBSERVABLE:no-sheet'; const rows = []; for (let r = g.r1; r <= g.r2; r += 1) { const row = []; for (let c = g.c1; c <= g.c2; c += 1) { const v = (sheet.cells[addr(c, r)] || {}).v; row.push(v === undefined || v === null ? '' : String(v)); } rows.push(row); } return rows; };
const withPivotProofs = (live, script, o) => (!Array.isArray(o.pivots) ? o : { ...o, pivots: o.pivots.map(pv => {
  if (!pv.proofSpec) return pv;
  const spec = pv.proofSpec;
  if (script.pivotProofUnobservable) return { name: pv.name, source: "UNOBSERVABLE:TypeError: Cannot read properties of undefined (reading 'ws')", proof: script.pivotProofUnobservable };
  const i = spec.dataRef.lastIndexOf('!');
  const sheet = spec.dataRef.slice(0, i).replace(/^'(.*)'$/, '$1').replace(/''/g, "'");
  const range = spec.dataRef.slice(i + 1).replace(/\$/g, '');
  const src = live.find(x => x.name === sheet);
  return { name: pv.name, source: "UNOBSERVABLE:TypeError: Cannot read properties of undefined (reading 'ws')", dataRef: spec.dataRef, fields: spec.fields, ...(pv.table ? { table: pv.table, layout: pv.layout } : {}),
    proof: { ...spec, source: { sheet, range, values: src ? gridOf(src, range) : 'UNOBSERVABLE:pivot-source-sheet-missing' }, others: live.filter(x => x.name !== sheet).map(x => ({ sheet: x.name, values: gridOf(x, range) })) } };
}) });

// X-F05 (LIVE m7-s21-probe1/2): the word form the spec/setter uses -> the enum string the getter reads back.
const UNDERLINE_ENUM = { single: 'xlUnderlineStyleSingle', singleAccounting: 'xlUnderlineStyleSingleAccounting', double: 'xlUnderlineStyleDouble', doubleAccounting: 'xlUnderlineStyleDoubleAccounting', none: 'xlUnderlineStyleNone' };
/** X-F05: underline's word -> its enum; strikeout / subscript / superscript true -> their own field true, false -> null
 * (the same null/false ambiguity bold/italic's public reader has); setting one script true drops the other to null. */
function applyFontEffects(f, font) {
  if ('underline' in font) f.underline = UNDERLINE_ENUM[font.underline];
  if ('strikeout' in font) f.strike = font.strikeout === true ? true : null;
  for (const [k, own, other] of [['subscript', 'sub', 'sup'], ['superscript', 'sup', 'sub']]) {
    if (!(k in font)) continue;
    f[own] = font[k] === true ? true : null;
    if (font[k] === true) f[other] = null;
  }
}

/** sheets: [{ name, cells: { A1: { v, f?, nf?, fill? } }, merges?: ['C1:D1'] }] */
function createStore(sheets) {
  return { durable: clone(sheets), saves: 0 };
}

function usedOf(sheet) {
  const filled = Object.entries(sheet.cells).filter(([, c]) => (c.f && c.f !== '') || (c.v !== null && c.v !== '' && c.v !== undefined));
  if (!filled.length) return null;
  const gs = filled.map(([a]) => parseRange(a));
  const c1 = Math.min(...gs.map(g => g.c1));
  const r1 = Math.min(...gs.map(g => g.r1));
  const c2 = Math.max(...gs.map(g => g.c2));
  const r2 = Math.max(...gs.map(g => g.r2));
  return c1 === c2 && r1 === r2 ? addr(c1, r1) : `${addr(c1, r1)}:${addr(c2, r2)}`;
}

function mergeAreaOf(sheet, a) {
  const g = parseRange(a);
  for (const m of sheet.merges || []) {
    const mg = parseRange(m);
    if (g.c1 >= mg.c1 && g.c1 <= mg.c2 && g.r1 >= mg.r1 && g.r1 <= mg.r2) return m;
  }
  return a;
}

function createFakeXlsxDriver(store, script = {}) {
  const calls = { open: 0, close: 0, forceSave: 0, canSave: 0, readLive: 0, dispatch: 0, readPackage: 0, bridgeWrites: 0 };
  const reads = [];
  let live = null;
  let openedFileId = null;
  let dirty = false;
  let literalOnSave = {}; // L-06 'copy-saves-literal': sheet\u0000addr -> the source's literal { f, v }, until a formula write replaces it
  let confirmedAt = null; // the dispatch count at the last confirmed forcesave (M7 save contract)

  /**
   * The MEASURED getter shapes (M6 wave-0 LIVE probe, run m6-wave0-1790539082918-d3eb97, internal tracker):
   * GetValue returns a STRING for numbers too ('1', '2.5'); GetFormula returns the formula for a formula
   * cell ('= Monthly_Plan!B16') and the VALUE TEXT for any other non-blank cell ('1', 'x'); a blank cell
   * reads v '' / f '' / COUNT 0; COUNT(cell) is 1 for a number, 0 otherwise; MergeArea is the merged
   * area's address on every cell inside it, the cell itself otherwise.
   */
  const nb = (a, side) => { const g = parseRange(a); const [dc, dr] = { t: [0, -1], b: [0, 1], l: [-1, 0], r: [1, 0] }[side]; return g.c1 + dc < 1 || g.r1 + dr < 1 ? null : addr(g.c1 + dc, g.r1 + dr); };
  const OPP = { t: 'b', b: 't', l: 'r', r: 'l' };
  function styleOf(sheet, a, c, v) {
    // X-F07a (LIVE m7-s24-probe1/2): rotation's untouched default is 0 -- the SAME
    // default EVERY existing cell already has, so no regression for unrelated format ops.
    // `script.rotUnobservable` (an address list) opts specific cells into the MEASURED throw
    // (a never-formatted, truly blank cell) for tests that target that gap directly.
    const out = { wrap: c.wrap === true, rot: (script.rotUnobservable || []).includes(a) ? 'UNOBSERVABLE:threw' : (c.rot === undefined ? 0 : c.rot) };
    // M7 X-F05 (LIVE m7-s21-probe1/2): underline's untouched/off default is the enum
    // string 'xlUnderlineStyleNone' (not null -- it has no boolean off-state); strike/sub/sup are
    // null both when never touched and when explicitly set false (the same null/false ambiguity
    // bold/italic's public reader already has).
    const f = { bold: null, italic: null, name: 'Calibri', size: 11, color: 0, underline: 'xlUnderlineStyleNone', strike: null, sub: null, sup: null, ...(c.font || {}) };
    // X-F02: an empty cell's font reads the SAME shape as a non-empty cell's --
    // `emptyFontUnobservable` opts a specific empty cell into the real engine's throw case.
    out.font = v === '' ? ((script.emptyFontUnobservable || []).includes(a) ? 'UNOBSERVABLE:threw' : f) : ((script.mixedRuns || []).includes(a) ? 'UNOBSERVABLE:mixed-runs' : f);
    if (script.appearanceBuild === false) {
      out.align = 'UNOBSERVABLE:runtime-not-verified'; out.border = 'UNOBSERVABLE:runtime-not-verified';
      out.readingOrder = 'UNOBSERVABLE:runtime-not-verified'; out.indent = 'UNOBSERVABLE:runtime-not-verified'; out.shrink = 'UNOBSERVABLE:runtime-not-verified';
      return out;
    }
    out.align = { h: null, v: 0, ...(c.align || {}) };
    const own = c.border || {};
    const shade = x => (script.blackReadsNull && x.c === 0 && x.s > 0 ? { s: x.s, c: null } : x);
    const side = k => {
      if (own[k] && own[k].s > 0) return shade(own[k]);
      const n = nb(a, k);
      const o = n && sheet.cells[n] && sheet.cells[n].border && sheet.cells[n].border[OPP[k]];
      return o && o.s > 0 ? shade(o) : { s: 0, c: 0 };
    };
    out.border = { t: side('t'), b: side('b'), l: side('l'), r: side('r'), d: shade(own.d || { s: 0, c: 0 }), dd: script.diagonalFlagUnobservable === 'dd' ? 'UNOBSERVABLE:threw' : own.dd === true, du: script.diagonalFlagUnobservable === 'du' ? 'UNOBSERVABLE:threw' : own.du === true };
    // X-F07b/c (owner GO 2026-10-05): readingOrder/indent/shrink, gated by their OWN, SEPARATE
    // build flag (`alignExtBuild`) -- independent of `appearanceBuild` (align.h/v/border), same
    // as the real driver's two independent build lists. Untouched defaults: readingOrder 0
    // ('context'), indent 0, shrink false -- the SAME defaults every existing cell already has,
    // so no regression for unrelated format ops. Per-address opt-in throws (readingOrderUnobservable
    // / indentUnobservable / shrinkUnobservable) for tests that target that gap directly.
    if (script.alignExtBuild === false) {
      out.readingOrder = 'UNOBSERVABLE:runtime-not-verified'; out.indent = 'UNOBSERVABLE:runtime-not-verified'; out.shrink = 'UNOBSERVABLE:runtime-not-verified';
    } else {
      out.readingOrder = (script.readingOrderUnobservable || []).includes(a) ? 'UNOBSERVABLE:threw' : (c.readingOrder === undefined ? 0 : c.readingOrder);
      out.indent = (script.indentUnobservable || []).includes(a) ? 'UNOBSERVABLE:threw' : (c.indent === undefined ? 0 : c.indent);
      out.shrink = (script.shrinkUnobservable || []).includes(a) ? 'UNOBSERVABLE:threw' : (c.shrink === undefined ? false : c.shrink);
    }
    return out;
  }

  function readCell(sheet, a, fmt, sty) {
    const c = sheet.cells[a] || { v: null };
    const formula = c.f || '';
    const blank = c.v === null || c.v === undefined || c.v === '';
    const count = script.countUnreadable ? null : (!blank && typeof c.v === 'number' ? 1 : 0);
    const v = blank ? '' : String(c.v);
    const out = { sheet: sheet.name, a, v, f: formula || v, count, merge: mergeAreaOf(sheet, a) };
    if (fmt) {
      out.nf = c.nf || 'General'; out.fill = c.fill || 'S:No Fill';
      const hide = script.boldUnobservable === 'always' || (script.boldUnobservable === 'after-dispatch' && calls.dispatch > 0);
      out.bold = hide ? 'UNOBSERVABLE:threw' : c.bold === true;
      const hideItalic = script.italicUnobservable === 'always' || (script.italicUnobservable === 'after-dispatch' && calls.dispatch > 0);
      out.italic = hideItalic ? 'UNOBSERVABLE:threw' : c.italic === true;
    }
    if (sty) Object.assign(out, styleOf(sheet, a, c, v));
    // M7 S3 review negatives: after the first dispatch a style getter throws (alignment) / one style key is gone (fill)
    const later = calls.dispatch > 0;
    if (later && script.styleGetterThrows && sty) out.align = 'UNOBSERVABLE:threw';
    if (later && script.fillDropped && fmt) delete out.fill;
    return out;
  }

  function applyClear({ sheet, cells }) {
    const ws = live.find(s => s.name === sheet);
    if (!ws) return;
    const f = script.fault;
    if (f === 'ignore') return;
    const plan = f === 'partial' ? cells.slice(0, 1) : cells;
    for (const a of plan) {
      const c = ws.cells[a] || {};
      if (f === 'clear-resets-format') ws.cells[a] = { v: null };
      else if (f === 'clear-drops-bold') ws.cells[a] = { ...c, v: null, f: undefined, bold: false };
      else ws.cells[a] = { ...c, v: null, f: undefined };
    }
    if (f === 'side-effect') ws.cells.Z99 = { v: 'side effect' };
    if (f === 'side-erase') for (const k of script.eraseCells || []) ws.cells[k] = { v: null };
    dirty = true;
  }

  // L-14 / L-17 (LIVE-measured, zero): ClearFormats() puts every format axis back to base and KEEPS the value; Clear() also empties the value.
  // Faults: 'ignore' (nothing happens), 'clear-formats-erases-value' (the formats scope also drops the value),
  // 'clear-formats-keeps-axis' (script.keepAxis survives), 'side-effect' (Z99 appears).
  function applyClearScope({ sheet, cells, action }) {
    const ws = live.find(s => s.name === sheet);
    if (!ws) return;
    const f = script.fault;
    if (f === 'ignore') return;
    for (const a of cells) {
      const c = ws.cells[a] || {};
      const kept = f === 'clear-formats-keeps-axis' && script.keepAxis ? { [script.keepAxis]: c[script.keepAxis] } : {};
      const value = action === 'clear-all' || f === 'clear-formats-erases-value' ? { v: null } : { v: c.v, ...(c.f ? { f: c.f } : {}) };
      ws.cells[a] = { ...value, ...kept };
    }
    if (f === 'side-effect') ws.cells.Z99 = { v: 'side effect' };
    dirty = true;
  }

  function applyWrite({ sheet, cells }) {
    const ws = live.find(s => s.name === sheet);
    if (!ws) return;
    const f = script.fault;
    if (f === 'ignore') return;
    const plan = f === 'partial' ? cells.slice(0, 1) : cells;
    for (const { a, v, nf } of plan) {
      const prev = ws.cells[a] || {};
      if (nf !== undefined) {
        // M7 S26/S28 (tier D probe): '@' then the value keeps the text; a value then a mask stores the serial.
        // Faults: 'typed-order-swapped' (the value first: numeric-looking text coerced), 'typed-no-format' (nf not
        // applied), 'typed-format-spills' (the next cell of the column gets the format too), 'typed-minute-off'
        const coerce = f === 'typed-order-swapped' && typeof v === 'string' && /^-?\d+(\.\d+)?$/.test(v);
        const shifted = f === 'typed-minute-off' && typeof v === 'number' ? v + 1 / 1440 : v; // one minute off
        ws.cells[a] = { ...prev, v: coerce ? Number(v) : shifted, f: undefined, ...(f === 'typed-no-format' ? {} : { nf }) };
        if (f === 'typed-format-spills') { const g = parseRange(a); const n = addr(g.c1, g.r1 + 1); ws.cells[n] = { ...(ws.cells[n] || { v: null }), nf }; }
        continue;
      }
      if (typeof v === 'string' && v.startsWith('=')) {
        // the MEASURED re-rendering (wave-2 sets a/b): '= ' prefix, upper-case and space-free outside
        // string literals; a sheet name that differs from a live name only in case drops the formula
        const lost = [...v.matchAll(/([A-Za-z_][\w.]*)!/g)].some(m => !live.some(x => x.name === m[1]) && live.some(x => x.name.toLowerCase() === m[1].toLowerCase()));
        if (lost) { ws.cells[a] = { ...prev, v: null, f: undefined }; continue; }
        // string literals AND quoted sheet names stay verbatim (wave 2 set b; M7 R26c probe: = 'Bevétel x'!A1*2)
        let rendered = '= ' + v.slice(1).split(/("(?:[^"]|"")*"|'(?:[^']|'')*')/).map((part, i) => (i % 2 ? part : part.replace(/\s+/g, '').toUpperCase())).join('');
        for (const x of live) rendered = rendered.split(`${x.name.toUpperCase()}!`).join(`${x.name}!`); // sheet names keep their live case (measured)
        const value = (script.formulaValues && script.formulaValues[a] !== undefined) ? script.formulaValues[a] : 0;
        ws.cells[a] = { ...prev, v: value, f: rendered };
      } else ws.cells[a] = { ...prev, v, f: undefined };
    }
    // M7 S3 negative controls: the write ALSO changes a style key of its first cell
    if (f === 'write-changes-style') ws.cells[plan[0].a] = { ...ws.cells[plan[0].a], wrap: true };
    if (f === 'side-effect') ws.cells.Z99 = { v: 'side effect' };
    dirty = true;
  }

  const selfRe = n => new RegExp(`(^|[^\\w.])${n}!`, 'g');
  const rewrite = (text, from, to) => (typeof text === 'string' ? text.split(/("(?:[^"]|"")*")/).map((part, i) => (i % 2 ? part : part.replace(selfRe(from), `$1${to}!`))).join('') : text);
  function applyCopy(m) {
    const f = script.fault;
    const src = live.find(x => x.name === m.sheet);
    const keep = f === 'copy-keeps-self-reference';
    const re = t => (keep ? t : rewrite(t, m.sheet, m.name));
    const cells = {};
    for (const [a, c] of Object.entries(src.cells)) cells[a] = { ...c, ...(c.f ? { f: re(c.f) } : {}) };
    const o = src.objects || { charts: [], pivots: [], names: [] };
    const objects = {
      charts: o.charts.map(c => ({ ...c, series: c.series.map(x => ({ name: re(x.name), values: re(x.values), cat: re(x.cat), x: re(x.x) })) })),
      pivots: f === 'copy-drops-object' ? [] : o.pivots.map(pv => ({ ...pv, source: re(pv.source) })),
      names: o.names.map(n => ({ ...n, refersTo: re(n.refersTo) })),
    };
    const rules = (src.rules || []).map(r => ({ ...r, formula: re(r.formula) })).slice(0, f === 'copy-drops-rule' ? -1 : undefined);
    if (f === 'copy-drops-cell') delete cells.A1;
    live.push({ name: m.name, cells, objects, rules, merges: src.merges });
    if (f === 'copy-duplicate') live.push({ name: m.name, cells: {}, objects: { charts: [], pivots: [], names: [] }, rules: [] });
    if (f === 'side-effect') live.find(x => x.name !== m.sheet && x.name !== m.name).cells.Z99 = { v: 'side effect' }; // a THIRD sheet: L3's job
    dirty = true;
  }

  function applySheet(m) {
    const f = script.fault;
    if (f === 'ignore') return;
    const refRe = n => new RegExp(`(^|[^\\w.])${n}!`, 'g');
    if (m.action === 'create_sheet') {
      live.splice(1, 0, { name: f === 'create-default-name' ? 'Sheet2' : m.name, cells: {} });
      if (f === 'create-reorders') live.unshift(live.pop()); // L1 cannot see it: count +1, exact name, empty
    }
    if (m.action === 'rename_sheet') {
      live.find(x => x.name === m.sheet).name = m.name;
      // M7 R26c: a quoted reference follows too, written back quoted when the new name needs quotes
      const bare = n => /^[A-Za-z_\u00C0-\uFFFF][\w.\u00C0-\uFFFF]*$/.test(n) && !/^[A-Za-z]{1,3}\d+$/.test(n);
      const newRef = bare(m.name) ? m.name : `'${m.name.replace(/'/g, "''")}'`;
      const quotedOld = `'${m.sheet.replace(/'/g, "''")}'!`;
      if (f !== 'rename-no-follow') for (const ws of live) for (const c of Object.values(ws.cells)) if (c.f) c.f = c.f.split(/("(?:[^"]|"")*")/).map((part, i) => (i % 2 ? part : part.split(quotedOld).join(`${newRef}!`).replace(refRe(m.sheet), `$1${newRef}!`))).join('');
      // measured (W3f run 2): a validation follows the rename, a conditional format does NOT; names
      // follow here by default (UNMEASURED on the engine) -- faults leave one of them behind
      const rw = t => rewrite(t, m.sheet, m.name);
      for (const ws of live) {
        if (f !== 'rename-dv-no-follow') ws.rules = (ws.rules || []).map(r => (r.kind === 'validation' ? { ...r, formula: rw(r.formula), ...(r.formula2 ? { formula2: rw(r.formula2) } : {}) } : r));
        if (ws.objects && f !== 'rename-name-no-follow') ws.objects = { ...ws.objects, names: (ws.objects.names || []).map(n => ({ ...n, refersTo: rw(n.refersTo) })) };
        // measured (W3g run 1): chart series follow; a pivot source follows in the session
        if (ws.objects) ws.objects = { ...ws.objects, charts: (ws.objects.charts || []).map(c => ({ ...c, series: c.series.map(x => ({ ...x, name: rw(x.name), values: rw(x.values), cat: rw(x.cat), x: rw(x.x) })) })), pivots: (ws.objects.pivots || []).map(pv => ({ ...pv, source: rw(pv.source) })) };
      }
      if (f !== 'rename-name-no-follow') liveNames = liveNames.map(n => ({ ...n, ref: rw(n.ref) }));
    }
    if (m.action === 'delete_sheet') {
      live = live.filter(x => x.name !== m.sheet);
      for (const ws of live) for (const c of Object.values(ws.cells)) if (c.f && refRe(m.sheet).test(c.f)) { c.f = c.f.replace(new RegExp(`${m.sheet}![A-Z]+\\d+`, 'g'), '#REF!'); c.v = '#REF!'; }
    }
    if (m.action === 'move_sheet') {
      const ws = live.find(x => x.name === m.sheet);
      live = live.filter(x => x !== ws);
      const j = live.findIndex(x => x.name === m.reference);
      live.splice(m.where === 'before' ? j : j + 1, 0, ws);
    }
    if (f === 'side-effect') live[0].cells.Z99 = { v: 'side effect' };
    dirty = true;
  }

  const hit = (m, g) => { const mg = parseRange(m); return mg.c1 <= g.c2 && mg.c2 >= g.c1 && mg.r1 <= g.r2 && mg.r2 >= g.r1; };
  function applyMerge(m) {
    const ws = live.find(s => s.name === m.sheet);
    if (!ws) return;
    const f = script.fault;
    if (f === 'ignore') return;
    let range = f === 'merge-lands-elsewhere' ? script.elsewhere : m.range;
    let g = parseRange(range);
    if (f === 'partial' && m.action === 'merge_range') { range = `${addr(g.c1, g.r1)}:${addr(g.c2, g.r1)}`; g = parseRange(range); }
    const merges = ws.merges || [];
    if (m.action === 'merge_range' && m.across) {
      // X-L06: Merge(true) splits into ONE MERGE PER ROW (adc47f6 LIVE probe) -- the content guard
      // (classify) already refused any row with a non-left-most non-blank cell before dispatch, so
      // here every row's non-left-most cell is simply emptied, mirroring the measured engine.
      const rowAreas = [];
      for (let r = g.r1; r <= g.r2; r += 1) {
        const rowG = { c1: g.c1, r1: r, c2: g.c2, r2: r, count: g.cols };
        const tl = addr(g.c1, r);
        for (let c = g.c1; c <= g.c2; c += 1) { const a = addr(c, r); if (a !== tl) { const cell = ws.cells[a]; if (cell) ws.cells[a] = { ...cell, v: null, f: undefined }; } }
        rowAreas.push(rangeText(rowG));
      }
      ws.merges = [...merges.filter(x => !hit(x, g)), ...rowAreas];
    } else if (m.action === 'merge_range') {
      const tl = addr(g.c1, g.r1);
      for (const a of cellsOf(g)) if (a !== tl || f === 'merge-drops-top-left') { const c = ws.cells[a]; if (c) ws.cells[a] = { ...c, v: null, f: undefined }; }
      ws.merges = [...merges.filter(x => !hit(x, g)), range];
    } else {
      const drop = merges.filter(x => hit(x, g));
      const keep = f === 'partial' ? drop.slice(1) : [];
      ws.merges = merges.filter(x => !hit(x, g)).concat(keep);
    }
    if (f === 'side-merge') ws.merges = [...ws.merges, script.sideMerge];
    if (f === 'side-effect') ws.cells.Z99 = { v: 'side effect' };
    dirty = true;
  }

  function applyLayout(m) {
    const ws = live.find(s => s.name === m.sheet);
    if (!ws) return;
    const f = script.fault;
    if (f === 'ignore') return;
    ws.layout = ws.layout || { cols: {}, rows: {} };
    const book = m.kind === 'columns' ? (ws.layout.cols = ws.layout.cols || {}) : (ws.layout.rows = ws.layout.rows || {});
    if (m.prop === 'autofit') {
      // M7 X-L01/X-L02: simulate AutoFit's computed size (there is no real content to fit
      // against in the fake, so a fixed value stands in). Idempotency / non-cumulative-ness across a
      // retry is an owner-option-B requirement proven externally (the LIVE harness, two real-driver
      // calls) -- a fixed fitted value is enough for the offline family tests (plan/gate/lower/L1/L3).
      // 'autofit-no-size' leaves the size unread -- a RED L1 test (verifyPrimitive's own weaker, honest
      // check: readable size, not hidden).
      const size = m.kind === 'columns' ? 'width' : 'height';
      for (const i of f === 'partial' ? m.items.slice(0, 1) : m.items) {
        if (f === 'autofit-no-size') continue;
        book[i] = { ...(book[i] || {}), [size]: m.kind === 'columns' ? 18 : 22.5, hidden: false };
      }
    } else {
    // X-L04 (S18 LIVE m7-s18-probe2): the measured row-axis drift (15.37 set -> 15.35 after reopen,
    // 0.02 off) is WITHIN TOL (0.05) and must still hold -- the existing 'layout-drifts' (+0.5) proves
    // the OUTSIDE-tolerance FAIL; this proves the INSIDE-tolerance PASS, same mechanism, smaller offset.
    const value = f === 'layout-drifts' && typeof m.value === 'number' ? m.value + 0.5
      : f === 'layout-drifts-within-tol' && typeof m.value === 'number' ? m.value - 0.02 : m.value;
    const stored = m.prop === 'width' && typeof value === 'number' && f !== 'width-stores-requested' ? canonWidth(value) : value;
    for (const i of f === 'partial' ? m.items.slice(0, 1) : m.items) book[i] = { ...(book[i] || {}), [m.prop]: stored };
    }
    // M7 S3 negative control: the size step ALSO changes a style key (script.styleCell, default A1)
    if (f === 'layout-changes-style') { const at = script.styleCell || 'A1'; ws.cells[at] = { ...(ws.cells[at] || { v: null }), fill: hexInt('#FF0000') }; }
    if (f === 'side-layout') ws.layout.cols[script.sideColumn] = { ...(ws.layout.cols[script.sideColumn] || {}), width: 33 };
    // W4e LIVE: an unset (auto) row may READ as its single-line default after a write -- the same auto state
    if (f === 'auto-row-materialized' || f === 'side-row-height') ws.layout.rows = { ...(ws.layout.rows || {}), [script.sideRow]: { ...((ws.layout.rows || {})[script.sideRow] || {}), height: f === 'auto-row-materialized' ? 14.25 : 30 } };
    if (f === 'auto-row-hidden') ws.layout.rows = { ...(ws.layout.rows || {}), [script.sideRow]: { height: 14.25, hidden: true } };
    if (f === 'side-effect') ws.cells.Z99 = { v: 'side effect' };
    dirty = true;
  }

  // OD-2a / D-W: the engine's own column-width chain (Workbook.js 4994-5022), maxDigitWidth 7 / paddingPlusBorder 5 (the
  // measured default font): 20.5 -> 20.43, 0.25 -> 0.58, 10.25 -> 10.29 (p54). A WIDTH write stores the canonical value (the
  // engine stores the quantised model width); the read answers `canon` per requested width unless the build is not
  // verified (script.colWidthCanonBuild === false) or it throws (script.colWidthCanonThrows).
  const MDW = 7;
  const PAD = 5;
  const canonWidth = x => {
    if (x <= 0) return 0;
    const mcw = Math.floor((x * MDW + PAD) / MDW * 256) / 256;
    const px = Math.floor(((256 * mcw + Math.floor(128 / MDW)) / 256) * MDW);
    const one = MDW + PAD;
    return px < one ? (1 - Math.floor(100 * (one - px) / one + 0.49999) / 100) : Math.floor((px - PAD) / MDW * 100 + 0.5) / 100;
  };
  const canonRead = list => {
    if (script.colWidthCanonBuild === false) return 'UNOBSERVABLE:runtime-not-verified';
    if (script.colWidthCanonThrows) return 'UNOBSERVABLE:TypeError: threw';
    return Object.fromEntries(list.map(x => [String(x), canonWidth(x)]));
  };
  const colNameOf = n => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
  function layoutOf(asked) {
    const out = {};
    for (const sheet of live) {
      if (script.layoutUnobservable) { out[sheet.name] = 'UNOBSERVABLE:threw'; continue; }
      const u = usedOf(sheet);
      const g = u ? parseRange(u) : null;
      let maxCol = g ? g.c2 : 0;
      let maxRow = g ? g.r2 : 0;
      const canonAsked = [];
      for (const x of asked) if (x.sheet === sheet.name) { maxCol = Math.max(maxCol, x.maxCol || 0); maxRow = Math.max(maxRow, x.maxRow || 0); canonAsked.push(...(x.canon || [])); }
      const L = sheet.layout || { cols: {}, rows: {} };
      const cols = {};
      const rows = {};
      for (let c = 1; c <= maxCol; c += 1) { const e = (L.cols || {})[colNameOf(c)] || {}; const hidden = e.hidden === true; cols[colNameOf(c)] = { width: hidden ? 0 : (e.width === undefined ? 8.43 : e.width), hidden }; }
      for (let r = 1; r <= maxRow; r += 1) { const e = (L.rows || {})[String(r)] || {}; const hidden = e.hidden === true; rows[String(r)] = { height: hidden ? null : (e.height === undefined ? null : e.height), hidden }; }
      out[sheet.name] = canonAsked.length ? { cols, rows, canon: canonRead(canonAsked) } : { cols, rows };
    }
    return out;
  }

  const H = { left: 1, center: 2, right: 0, justify: 3 };
  const V = { bottom: 0, center: 1, distributed: 2, justify: 3, top: 4 };
  // X-B06 (MEASURED, m7-b06-probe1): 9 additional styles, distinct codes, no collision.
  const STY = { thin: 7, medium: 12, thick: 13, dashed: 6, double: 1, hair: 2, dashDotDot: 3, dashDot: 4, dotted: 5, mediumDashDotDot: 8, slantDashDot: 9, mediumDashDot: 10, mediumDashed: 11 };
  const READING_ORDER = { context: 0, ltr: 1, rtl: 2 };
  const hexInt = h => (h === 'auto' ? null : parseInt(h.slice(1), 16)); // X-B07: 'auto' reads as c: null (measured)
  function styleCell(ws, a, spec) {
    const c = { ...(ws.cells[a] || { v: null }) };
    if ('fill' in spec && spec.fill !== 'none') c.fill = hexInt(spec.fill);
    if ('wrap' in spec) c.wrap = spec.wrap;
    if ('numberFormat' in spec) c.nf = spec.numberFormat;
    if (spec.font) {
      const f = { ...(c.font || {}) };
      for (const [k, v] of Object.entries(spec.font)) { if (!['bold', 'italic', 'underline', 'strikeout', 'subscript', 'superscript'].includes(k)) f[k] = k === 'color' ? hexInt(v) : v; }
      if (spec.font.bold === true) { f.bold = true; c.bold = true; } // W4 LIVE 2026-09-29: font bold IS the cell bold
      else if (spec.font.bold === false) { f.bold = null; c.bold = false; } // X-F01/S07: the public reader never reports a hard false -- only the existing internal cell-level reader does
      // X-F01 italic: mirrors bold's own shadow-flip (MEASURED, m7-f01-live1 LIVE, cell D9,
      //) -- font.italic:true also flips the internal cell-level 'italic' key true.
      if (spec.font.italic === true) { f.italic = true; c.italic = true; }
      else if (spec.font.italic === false) { f.italic = null; c.italic = false; } // X-F01/S07: public reader never reports a hard false
      applyFontEffects(f, spec.font);
      c.font = f;
    }
    if (spec.align) c.align = { ...(c.align || {}), ...('horizontal' in spec.align ? { h: H[spec.align.horizontal] } : {}), ...('vertical' in spec.align ? { v: V[spec.align.vertical] } : {}) };
    // GH #93 (MEASURED, test/live/justify-text-cell-probe.live.cjs, session + reopen, text / number / empty cells):
    // SetAlignHorizontal('justify') also turns the cell's wrap ON -- the engine's own coupling
    if (spec.align && spec.align.horizontal === 'justify' && !('wrap' in spec)) c.wrap = true;
    if (spec.align && 'rotation' in spec.align) c.rot = spec.align.rotation;
    // X-F07b/c (owner GO 2026-10-05): readingOrder/indent/shrink, flat keys (same reasoning as `rot`).
    if (spec.align && 'readingOrder' in spec.align) c.readingOrder = READING_ORDER[spec.align.readingOrder];
    if (spec.align && 'indent' in spec.align) c.indent = spec.align.indent;
    if (spec.align && 'shrink' in spec.align) c.shrink = spec.align.shrink;
    // X-B05 (MEASURED, m7-b05-probe1): DiagonalDown/DiagonalUp share ONE `d` style+colour -- writing either
    // sets `d` and ONLY that direction's own flag (dd or du); the OTHER flag is left exactly as it was.
    if (spec.border && spec.border.sides && spec.border.sides.some(x => x === 'diagonalDown' || x === 'diagonalUp')) {
      const border = { ...(c.border || {}), d: { s: STY[spec.border.style], c: hexInt(spec.border.color) } };
      if (script.fault !== 'diagonal-missing-flag') for (const x of spec.border.sides) border[x === 'diagonalDown' ? 'dd' : 'du'] = true;
      c.border = border;
    } else if (spec.border && !['outline', 'inside', 'all', 'remove'].includes(spec.border.mode)) { const b = { ...(c.border || {}) }; for (const x of spec.border.sides) b[{ top: 't', bottom: 'b', left: 'l', right: 'r' }[x]] = { s: STY[spec.border.style], c: hexInt(spec.border.color) }; c.border = b; } // outline/inside/all write via applyOutline's edge list, remove via applyRemove
    if (script.fault === 'format-extra-key') c.nf = '0.00';
    if (script.fault === 'bold-cleared') c.bold = false; // the cell-level bold goes the WRONG way under a font.bold
    if (script.fault === 'bold-false-not-applied') c.bold = true; // X-F01: a bold:false write that silently no-ops
    if (script.fault === 'italic-false-not-applied') c.italic = true; // X-F01: an italic:false write that silently no-ops
    ws.cells[a] = c;
  }
  /**
   * M7 S4 S01: the outline's per-edge-cell SetBorders, as MEASURED (S4 probe): the edge cell's own side is set; a ring
   * neighbour's own facing side that DIFFERS is cleared (an equal one stays). L3 negative controls: 'outline-inner-side'
   * (an edge cell's inner side set too) | 'outline-miss-side' (one edge side not set) | 'outline-ring-not-cleared' (the
   * neighbour's differing facing side kept) | 'outline-ring-other-side' (a ring cell's non-facing side changed) |
   * 'outline-inner-edge-lost' (an inner cell's own side cleared) | 'outline-diagonal-lost'.
   */
  function applyOutline(ws, m, f) {
    const SD = { Top: 't', Bottom: 'b', Left: 'l', Right: 'r' };
    const OP = { t: 'b', b: 't', l: 'r', r: 'l' };
    const prop = { s: STY[m.spec.border.style], c: hexInt(m.spec.border.color) };
    const edges = f === 'outline-miss-side' ? m.edges.slice(1) : m.edges;
    for (const [a, sides] of edges) {
      for (const x of sides.map(y => SD[y])) {
        const c = { v: null, ...(ws.cells[a] || {}) };
        c.border = { ...(c.border || {}), [x]: prop };
        ws.cells[a] = c;
        const n = nb(a, x);
        const nc = n && ws.cells[n];
        const own = nc && nc.border && nc.border[OP[x]];
        if (f !== 'outline-ring-not-cleared' && own && own.s > 0 && !(own.s === prop.s && own.c === prop.c)) ws.cells[n] = { ...nc, border: { ...nc.border, [OP[x]]: { s: 0, c: 0 } } };
      }
    }
    const first = m.edges[0] && m.edges[0][0];
    if (f === 'outline-inner-side' && first) ws.cells[first] = { ...ws.cells[first], border: { ...ws.cells[first].border, r: prop, b: prop } };
    const at = script.faultCell;
    if (f === 'outline-ring-other-side' && at) ws.cells[at] = { v: null, ...(ws.cells[at] || {}), border: { ...((ws.cells[at] || {}).border || {}), l: { s: 13, c: 0 } } };
    if (f === 'outline-inner-edge-lost' && at && ws.cells[at] && ws.cells[at].border) ws.cells[at] = { ...ws.cells[at], border: {} };
    if (f === 'outline-diagonal-lost' && at && ws.cells[at] && ws.cells[at].border) ws.cells[at] = { ...ws.cells[at], border: { ...ws.cells[at].border, d: { s: 0, c: 0 }, dd: false } };
  }

  /**
   * M7 B04 (remove, a maintainer decision #ref): the composite's net observable effect -- every side in `m.removeClear[a]`
   * reads REMOVED on cell `a`, every other axis of `a` exactly as it was (the real composite reads every kept axis
   * before ClearFormats() and restores it unchanged; simulating that restore as a no-op is faithful, not a shortcut,
   * because that is precisely what a correct restore produces). Fault hooks: 'remove-one-cell-only' (only the first
   * named cell is cleared -- the shared edge's OTHER cell keeps its old side) | 'remove-kept-axis-lost' (the first
   * cell's fill is dropped, as if the restore silently skipped an axis) | 'remove-ring-other-side' (a window cell's
   * UNRELATED side changes). No diagonal-loss fault: gate() already refuses before the write whenever any window
   * cell carries a diagonal, so a diagonal can never reach this function in the first place -- there is no
   * correct/incorrect pair left to distinguish once that refusal has fired.
   */
  function applyRemove(ws, m, f) {
    const cells = f === 'remove-one-cell-only' ? m.removeCells.slice(0, 1) : m.removeCells;
    for (const a of cells) {
      const c = { v: null, ...(ws.cells[a] || {}) };
      const border = { ...(c.border || {}) };
      for (const x of m.removeClear[a]) border[x] = { s: 0, c: 0 };
      ws.cells[a] = { ...c, border };
    }
    const first = m.removeCells[0];
    if (f === 'remove-kept-axis-lost' && first) ws.cells[first] = { ...ws.cells[first], fill: undefined };
    const at = script.faultCell;
    if (f === 'remove-ring-other-side' && at) ws.cells[at] = { v: null, ...(ws.cells[at] || {}), border: { ...((ws.cells[at] || {}).border || {}), l: { s: 13, c: 0 } } };
  }

  /**
   * M7 X-F08 (fill-clear): the composite's net observable effect -- `fill` becomes
   * 'S:No Fill' (the MEASURED read: GetFillColor() returns the string 'No Fill' after ClearFormats(),
   * m7-s15-probe4/5, stored by xlsx-driver.cjs's fallback as 'S:No Fill'), every other axis of the
   * target cell exactly as it was (the real composite reads every kept axis before
   * ClearFormats() and restores it unchanged; simulating that restore as a no-op is faithful,
   * same reasoning as applyRemove's own comment above). Fault hooks: 'fillclear-fill-not-cleared'
   * (the fill stays, as if ClearFormats() never ran) | 'fillclear-kept-axis-lost' (the target's
   * wrap is dropped, as if the restore silently skipped an axis) | 'fillclear-neighbour-leak'
   * (script.leakCell's own border.b side flips to a different style -- simulates the shared-edge
   * write disturbing a neighbour, the regression the format-fill-clear.cjs shared-edge refusal
   * used to guard against before the LIVE measurement showed the restore cannot do this).
   */
  function applyFillClear(ws, m, f) {
    // X-B10 RED mutant 'fillclear-merge-destroyed': the engine's ClearFormats() DISSOLVED the merge LIVE (xb10 v2,
    // after save + reopen) -- every merge whose top-left is a cleared cell is gone, the members read as separate cells.
    if (f === 'fillclear-merge-destroyed') ws.merges = (ws.merges || []).filter(x => !m.fillClearCells.includes(x.split(':')[0]));
    for (const a of m.fillClearCells) {
      const c = { v: null, ...(ws.cells[a] || {}) };
      if (f !== 'fillclear-fill-not-cleared') c.fill = 'S:No Fill';
      if (f === 'fillclear-kept-axis-lost') c.wrap = false;
      ws.cells[a] = c;
    }
    if (f === 'fillclear-neighbour-leak' && script.leakCell) {
      const n = { v: null, ...(ws.cells[script.leakCell] || {}) };
      n.border = { ...(n.border || {}), b: { s: 13, c: 0 } }; // a DIFFERENT style than whatever it had
      ws.cells[script.leakCell] = n;
    }
  }

  // X-F04 (format_run): a cell may carry `runs` (an array of
  // {start, length, font: {bold, italic, name, size, color}}, 1-based, contiguous, covering
  // the WHOLE text) -- the per-character "ground truth" the run-map read/write mirrors. Absent
  // `runs` means the cell is uniformly described by its existing `font` (today's whole-cell
  // shape), matching the real engine's own state before any per-run write ever touched it.
  const RUN_FONT_DEFAULT = Object.freeze({ bold: null, italic: null, name: 'Calibri', size: 11, color: 0 });
  function runFontTupleAt(c, i) {
    if (Array.isArray(c.runs)) {
      const seg = c.runs.find(r => i >= r.start && i < r.start + r.length);
      if (seg) return { ...RUN_FONT_DEFAULT, ...seg.font };
    }
    return { ...RUN_FONT_DEFAULT, ...(c.font ? { bold: c.font.bold, italic: c.font.italic, name: c.font.name, size: c.font.size, color: c.font.color } : {}) };
  }
  /** The SAME character-by-character merge the real driver's PUBLIC run-map read does (zero
   * #ref's correction) -- 'UNOBSERVABLE:empty-cell' / 'UNOBSERVABLE:text-too-long' mirror the
   * real driver's own named gaps (`RUN_MAP_CHAR_CEILING` there is 256, matched here). */
  function runMapOf(ws, a) {
    const c = ws.cells[a] || { v: '' };
    const text = c.v === null || c.v === undefined ? '' : String(c.v);
    if (text === '') return 'UNOBSERVABLE:empty-cell';
    if (text.length > 256) return 'UNOBSERVABLE:text-too-long';
    const runs = [];
    let cur = null;
    for (let i = 1; i <= text.length; i += 1) {
      const ft = runFontTupleAt(c, i);
      const k = JSON.stringify(ft);
      if (cur && cur.key === k) cur.length += 1; else { cur = { start: i, length: 1, font: ft, key: k }; runs.push(cur); }
    }
    return { text, runs: runs.map(({ start, length, font }) => ({ start, length, font })) };
  }
  function arrayReadsOf(asked) {
    const out = {};
    for (const w of asked) {
      if (script.arrayUnobservable) { out[`${w.sheet}!${w.cell}`] = 'UNOBSERVABLE:threw'; continue; }
      const ws = live.find(x => x.name === w.sheet);
      if (!ws) { out[`${w.sheet}!${w.cell}`] = 'UNOBSERVABLE:sheet-missing'; continue; }
      const c = ws.cells[w.cell];
      out[`${w.sheet}!${w.cell}`] = !c ? '' : (c.array ? c.af : (c.f ? null : (c.v === null || c.v === undefined ? '' : c.v)));
    }
    return out;
  }
  // GH #300 K2: comments -- a sheet's `comments` list of { cell, text, author, userId, timeUtc, solved, replies }, read only
  // via `commentSheets` in row-major cell order (the driver's K1 read shape); script.commentsUnobservable -> 'UNOBSERVABLE:threw'.
  // The session author is script.sessionUser (default { id: 'oc_fake_tester', name: 'Fake Tester' }); null = the session has
  // none (the read says so, the dispatch refuses, as the real driver does). The write is AddComment + SetUserId; a second
  // comment on a commented cell is a silent no-op (measured). Faults 'ignore', 'side-effect', 'comment-wrong-author' (another
  // author name), 'comment-no-user-id' (SetUserId lost), 'comment-touches-other' (another comment's text changes).
  const sessionUserOf = () => (script.sessionUser === undefined ? { id: 'oc_fake_tester', name: 'Fake Tester' } : script.sessionUser);
  const rowMajor = (x, y) => { const a = parseRange(x.cell); const b = parseRange(y.cell); return a.r1 - b.r1 || a.c1 - b.c1; };
  function commentReadsOf(asked) {
    const out = {};
    for (const n of asked) {
      if (script.commentsUnobservable) { out[n] = 'UNOBSERVABLE:threw'; continue; }
      const ws = live.find(x => x.name === n);
      out[n] = ws ? { comments: clone(ws.comments || []).sort(rowMajor) } : 'UNOBSERVABLE:sheet-missing';
    }
    return out;
  }
  function applyComment(m) {
    const f = script.fault;
    const author = sessionUserOf();
    if (f === 'ignore' || !author) return;
    const ws = live.find(x => x.name === m.sheet);
    if (!ws) return;
    ws.comments = ws.comments || [];
    if (ws.comments.some(c => c.cell === m.cell)) return; // the measured silent no-op
    ws.comments.push({ cell: m.cell, text: m.text, author: f === 'comment-wrong-author' ? 'Someone Else' : author.name, userId: f === 'comment-no-user-id' ? '' : author.id, timeUtc: 0, solved: false, replies: [] });
    if (f === 'comment-touches-other') { const o = ws.comments.find(c => c.cell !== m.cell); if (o) o.text += ' (changed)'; }
    if (f === 'side-effect') live.find(x => x !== ws).cells.Z99 = { v: 'side effect' };
  }
  function applyArrayFormula(m) {
    const f = script.fault;
    if (f === 'ignore') return;
    const ws = live.find(x => x.name === m.sheet);
    if (!ws) return;
    const g = parseRange(m.range);
    const r2 = f === 'array-formula-spill-ring' ? g.r2 + 1 : g.r2;
    const ref = rangeText({ ...g, r2 });
    const text = f === 'array-formula-wrong-formula' ? `${m.formula}+1` : m.formula;
    for (let r = g.r1; r <= r2; r += 1) {
      for (let c = g.c1; c <= g.c2; c += 1) {
        if (f === 'array-formula-short' && r === g.r2 && c === g.c2) continue;
        ws.cells[addr(c, r)] = { v: '', f: text, array: ref, af: text };
      }
    }
    if (f === 'side-effect') live.find(x => x !== ws).cells.Z99 = { v: 'side effect' };
  }
  function runReadsOf(asked) {
    const out = {};
    for (const w of asked) {
      if (script.runMapUnobservable) { out[`${w.sheet}!${w.cell}`] = 'UNOBSERVABLE:threw'; continue; }
      const ws = live.find(x => x.name === w.sheet);
      out[`${w.sheet}!${w.cell}`] = ws ? runMapOf(ws, w.cell) : 'UNOBSERVABLE:sheet-missing';
    }
    return out;
  }
  /** Faults: 'ignore', 'side-effect' (Z99 written), 'format-run-wrong-axis' (the WRONG bold value
   * applied -- L1 FAIL), 'format-run-other-char-changed' (an untouched character's font flips --
   * L3 FAIL, `script.faultCharIndex` names which 1-based position), 'format-run-text-changed'
   * (the cell's text itself grows -- L3 `run-map-length-changed`). */
  function applyFormatRun(m) {
    const ws = live.find(x => x.name === m.sheet);
    if (!ws) return;
    const f = script.fault;
    if (f === 'ignore') return;
    const c = { ...(ws.cells[m.cell] || { v: '' }) };
    const text = c.v === null || c.v === undefined ? '' : String(c.v);
    const before = runMapOf(ws, m.cell);
    const existingRuns = typeof before === 'string' ? [{ start: 1, length: Math.max(text.length, 1), font: RUN_FONT_DEFAULT }] : before.runs;
    const flat = [];
    for (const r of existingRuns) for (let i = 0; i < r.length; i += 1) flat.push(r.font);
    const runsToApply = f === 'format-run-wrong-axis' ? m.runs.map(r => ('bold' in r.want ? { ...r, want: { ...r.want, bold: !r.want.bold } } : r)) : m.runs;
    for (const run of runsToApply) for (let i = run.start; i < run.start + run.length; i += 1) flat[i - 1] = { ...flat[i - 1], ...run.want };
    if (f === 'format-run-other-char-changed') { const idx = (script.faultCharIndex || 1) - 1; if (flat[idx]) flat[idx] = { ...flat[idx], bold: !flat[idx].bold }; }
    if (f === 'format-run-text-changed') c.v = `${text}X`;
    const runs = [];
    let cur = null;
    for (let i = 0; i < flat.length; i += 1) {
      const k = JSON.stringify(flat[i]);
      if (cur && cur.key === k) cur.length += 1; else { cur = { start: i + 1, length: 1, font: flat[i], key: k }; runs.push(cur); }
    }
    c.runs = runs.map(({ start, length, font }) => ({ start, length, font }));
    ws.cells[m.cell] = c;
    if (f === 'side-effect') ws.cells.Z99 = { v: 'side effect' };
    dirty = true;
  }

  /**
   * GH #83: border.mode 'cells', merge-aware -- `m.cellsSides` (address -> API side names, the SAME map
   * lower() embeds as CELLS_SIDES) replaces the uniform `spec.border.sides` write: each cell gets ONLY its
   * own (possibly perimeter-reduced) sides, and a cell absent from the map (fully interior to a merge) is
   * left untouched entirely -- no call, same as the real per-cell SetBorders loop never reaching it.
   */
  function applyCellsSidesBorder(ws, m, f) {
    const SD = { Top: 't', Bottom: 'b', Left: 'l', Right: 'r' };
    const prop = { s: STY[m.spec.border.style], c: hexInt(m.spec.border.color) };
    // GH #83 RED mutant 'gh83-interior-leak': every target cell gets EVERY requested side, ignoring the
    // per-cell perimeter filter -- proves L1/driftOf catches a write that bleeds onto a merge's interior edge.
    const sidesOf = a => (f === 'gh83-interior-leak' ? m.spec.border.sides.map(x => ({ top: 'Top', bottom: 'Bottom', left: 'Left', right: 'Right' }[x])) : m.cellsSides[a]);
    for (const a of f === 'gh83-interior-leak' ? m.cells : Object.keys(m.cellsSides)) {
      const sides = sidesOf(a);
      if (!sides || !sides.length) continue;
      const c = { v: null, ...(ws.cells[a] || {}) };
      c.border = { ...(c.border || {}) };
      for (const x of sides.map(y => SD[y])) c.border[x] = prop;
      ws.cells[a] = c;
    }
  }

  function applyFormat(m) {
    const ws = live.find(s => s.name === m.sheet);
    if (!ws) return;
    const f = script.fault;
    if (f === 'ignore') return;
    if (m.cellsSides) applyCellsSidesBorder(ws, m, f);
    else for (const a of f === 'partial' ? m.cells.slice(0, 1) : m.cells) styleCell(ws, a, m.spec);
    if (f === 'format-leaks') styleCell(ws, script.leakCell, m.spec);
    if (m.edges) applyOutline(ws, m, f);
    if (m.removeCells) applyRemove(ws, m, f);
    if (m.fillClearCells) applyFillClear(ws, m, f);
    // M7 S3 negative control: the format ALSO changes the value of its first cell
    if (f === 'format-changes-value') ws.cells[m.cells[0]] = { ...ws.cells[m.cells[0]], v: 'changed by format', f: undefined };
    if (f === 'side-effect') ws.cells.Z99 = { v: 'side effect' };
    dirty = true;
  }

  /**
   * M11.5a.4 B05 Bridge WRITE: what the REAL driver returns (xlsx-driver b05DispatchResult) for the body's outcome.
   * script.b05Write: undefined (every write accepted) | 'gate-refused' (the client gate, before any mutation) |
   * 'error-none' (setter missing on the first cell: nothing written) | 'partial' (indent written, shrink threw) |
   * 'error-applied' (everything written, then an internal throw) | 'malformed' (written, the receipt shape is wrong).
   */
  function b05Write(m) {
    const ws = live.find(s => s.name === m.sheet);
    const mode = script.b05Write;
    const fail = (evidence, partiallyApplied = false) => (partiallyApplied
      ? { ok: true, dispatched: true, partiallyApplied: true, outcome: 'b05-bridge-write-partial', evidence, error: null }
      : { ok: false, dispatched: true, outcome: 'b05-bridge-write-failed', evidence, error: null });
    if (mode === 'gate-refused') return fail({ ok: false, stage: 'bridge-write-gate', reason: script.b05GateReason || 'CAPABILITY_MISSING', mutation: 'NONE' });
    if (mode === 'error-none') return fail({ ok: false, stage: 'bridge-write', cell: m.cells[0], reason: 'BRIDGE:setIndent-missing', uncertain: false, partial: false, applied: [], cellsDone: [] });
    if (mode === 'partial') {
      const al = { ...m.spec.align }; delete al.shrink;
      styleCell(ws, m.cells[0], { ...m.spec, align: al }); calls.bridgeWrites += 1; dirty = true;
      return fail({ ok: false, stage: 'bridge-write', cell: m.cells[0], reason: 'BRIDGE:setShrinkToFit-threw', uncertain: false, partial: true, applied: ['indent'], cellsDone: [] }, true);
    }
    applyFormat(m); calls.bridgeWrites += m.cells.length;
    if (mode === 'error-applied') return fail({ ok: false, stage: 'bridge-write', cell: m.cells[0], reason: 'BRIDGE:internal-threw', uncertain: false, partial: false, applied: [], cellsDone: [] });
    if (mode === 'malformed') return fail({ ok: false, stage: 'bridge-write', cell: m.cells[0], reason: 'UNEXPECTED_SHAPE', uncertain: true, partial: false, applied: [], cellsDone: [] });
    drawingFault();
    return { ok: true, dispatched: true, outcome: 'command-completed', error: null };
  }

  let active = null;
  let liveNames = [];
  /** Workbook-level names (wave 6, meta.family 'names'): Api.AddDefName / Delete; faults 'ignore', 'name-wrong-ref', 'side-effect',
   * 'name-wrong-scope' (M7 R16). */
  function applyNames(m) {
    const f = script.fault;
    if (f === 'ignore') return;
    // stored VERBATIM, as the engine does (W5-W9 LIVE 2026-09-29: AddDefName(n, '=S!$A$1') kept the '=' and the name was broken)
    const ref = m.refersTo && f === 'name-wrong-ref' ? m.refersTo.replace(/\d+$/, d => String(Number(d) + 1)) : m.refersTo;
    // M7 R16: a sheet-scope name lives on its sheet (ws.AddDefName / ws.GetDefName(n).Delete); fault 'name-wrong-scope'
    // puts it at workbook level instead
    const ws = m.scope && f !== 'name-wrong-scope' ? live.find(x => x.name === m.scope) : null;
    if (ws) {
      const own = (ws.objects && ws.objects.names) || [];
      const kept = own.filter(n => n.name.toLowerCase() !== m.name.toLowerCase());
      // X-P01 RED mutant 'name-repoint-duplicates': the engine KEEPS the old name and adds the new one (the LIVE probe measured the opposite: ONE Print_Area, replaced)
      ws.objects = { charts: [], pivots: [], ...ws.objects, names: m.action.startsWith('set_') ? [...(f === 'name-repoint-duplicates' ? own : kept), { name: m.name, refersTo: `=${ref}` }] : kept };
    } else if (m.action.startsWith('set_')) liveNames = [...liveNames, { name: m.name, ref }];
    else liveNames = liveNames.filter(n => n.name.toLowerCase() !== m.name.toLowerCase());
    if (f === 'side-effect') live[0].cells.Z99 = { v: 'side effect' };
    dirty = true;
  }
  function applyFreeze(m) {
    const f = script.fault;
    if (f === 'ignore') return;
    const origin = active;
    const target = f === 'wrong-sheet' ? live.find(x => x.name === origin) : live.find(x => x.name === m.sheet);
    if (!target) return;
    if (m.action === 'unfreeze_panes') delete target.pane;
    else target.pane = { row0: m.rows + (f === 'freeze-off-by-one' ? 1 : 0), col0: m.columns };
    if (f === 'side-effect') { const other = live.find(x => x.name !== m.sheet); if (other) other.pane = { row0: 9, col0: 0 }; }
    active = f === 'no-restore' ? m.sheet : origin;
    dirty = true;
  }
  const panesOf = () => Object.fromEntries(live.map(x => [x.name, script.paneBuild === false ? 'UNOBSERVABLE:runtime-not-verified' : (x.pane && x.pane.split ? 'UNOBSERVABLE:pane-split-or-unknown' : (x.pane ? { state: 'frozen', row0: x.pane.row0, col0: x.pane.col0 } : { state: 'none' }))]));

  function applyStructure(m) {
    const f = script.fault;
    if (f === 'ignore') return;
    const { shiftFormula, moveIndex, moveSpan } = require('../../src/adapters/xlsx/shift.cjs');
    const count = f === 'partial' ? m.count - Math.sign(m.count) : m.count;
    const op = { sheet: m.sheet, axis: m.axis, at: m.at, count };
    const rows = m.axis === 'rows';
    const move = ws => {
      const cells = {};
      for (const [a, c] of Object.entries(ws.cells)) {
        const g = parseRange(a);
        const n = moveIndex(rows ? g.r1 : g.c1, op);
        if (n === null) continue;
        cells[rows ? addr(g.c1, n) : addr(n, g.r1)] = c;
      }
      ws.cells = cells;
      // M7 S5 R5 negative control 'merge-keeps-size': a merge the cut crosses moves but neither grows nor shrinks
      const span = (lo, hi) => { const sp = moveSpan(lo, hi, op); return sp && f === 'merge-keeps-size' ? [sp[0], sp[0] + hi - lo] : sp; };
      ws.merges = (ws.merges || []).map(x => { const g = parseRange(x); const sp = rows ? span(g.r1, g.r2) : span(g.c1, g.c2); if (!sp) return null; return rows ? `${addr(g.c1, sp[0])}:${addr(g.c2, sp[1])}` : `${addr(sp[0], g.r1)}:${addr(sp[1], g.r2)}`; }).filter(Boolean);
    };
    const target = live.find(x => x.name === m.sheet);
    if (!target) return;
    move(target);
    // M7 E R2 / R3a: a rule stays where it is (measured for rules before the cut); fault 'rule-moves' shifts every rule of
    // the target by the step (an L3 negative control)
    if (f === 'rule-moves') target.rules = (target.rules || []).map(r => { const g = parseRange(r.range); return { ...r, range: rows ? `${addr(g.c1, g.r1 + count)}:${addr(g.c2, g.r2 + count)}` : `${addr(g.c1 + count, g.r1)}:${addr(g.c2 + count, g.r2)}` }; });
    if (f === 'other-sheet-moves') { const o = live.find(x => x !== target); if (o) move(o); }
    if (f !== 'no-reference-follow') for (const ws of live) for (const c of Object.values(ws.cells)) if (c.f) c.f = shiftFormula(c.f, ws.name, op).formula;
    // M7 S4 R1: defined names follow the shift as a formula does (W5 + S4 probe); a reference into a deleted block
    // reads '#REF!'. L3 negative controls: 'name-no-follow' | 'name-wrong-shift' (one row / column too far) |
    // 'name-scope-lost' (a sheet-level name becomes workbook level) | 'name-dropped' | 'name-appears' | 'name-other-changes'
    // (an untouched name is re-pointed) -- on the first matching name
    const shiftRef = ref => { const s = shiftFormula(`=${ref}`, null, op); return s.broken ? '#REF!' : s.formula.slice(1); };
    const wrong = ref => { const s = shiftFormula(`=${ref}`, null, { ...op, count: op.count + Math.sign(op.count) }); return s.broken ? '#REF!' : s.formula.slice(1); };
    const follow = ref => (f === 'name-no-follow' ? ref : f === 'name-wrong-shift' && shiftRef(ref) !== ref ? wrong(ref) : shiftRef(ref));
    liveNames = liveNames.map(n => ({ ...n, ref: follow(n.ref) }));
    for (const ws of live) if (ws.objects && Array.isArray(ws.objects.names)) ws.objects = { ...ws.objects, names: ws.objects.names.map(n => ({ ...n, refersTo: `=${follow(String(n.refersTo).replace(/^=/, ''))}` })) };
    if (f === 'name-scope-lost') { const ws = live.find(x => x.objects && x.objects.names && x.objects.names.length); if (ws) { const [n] = ws.objects.names; ws.objects = { ...ws.objects, names: ws.objects.names.slice(1) }; liveNames.push({ name: n.name, ref: String(n.refersTo).replace(/^=/, '') }); } }
    if (f === 'name-dropped') liveNames = liveNames.slice(1);
    if (f === 'name-appears') liveNames = [...liveNames, { name: 'Appeared', ref: `${m.sheet}!$A$1` }];
    if (f === 'name-other-changes') { const i = liveNames.findIndex(n => !n.ref.toUpperCase().startsWith(`${m.sheet.toUpperCase()}!`)); if (i >= 0) liveNames[i] = { ...liveNames[i], ref: liveNames[i].ref.replace(/\$(\d+)$/, (_, r) => `$${Number(r) + 1}`) }; }
    if (f === 'side-effect') live.find(x => x !== target).cells.Z99 = { v: 'side effect' };
    // M7 S2a L3 negative controls: the step ALSO moves a layout object (on script.layoutSheet, default the target)
    const lw = live.find(x => x.name === (script.layoutSheet || m.sheet));
    if (f === 'layout-moves-filter' && lw && lw.autoFilter) lw.autoFilter = { ...lw.autoFilter, range: 'A1:A99' };
    if (f === 'layout-moves-row-break' && lw) lw.print = { ...(lw.print || {}), rowBreaks: ((lw.print || {}).rowBreaks || []).map(b => ({ ...b, id: b.id + 1 })) };
    if (f === 'layout-moves-col-break' && lw) lw.print = { ...(lw.print || {}), colBreaks: ((lw.print || {}).colBreaks || []).map(b => ({ ...b, id: b.id + 1 })) };
    if (f === 'layout-moves-pane' && lw && lw.pane) lw.pane = { ...lw.pane, row0: lw.pane.row0 + 1 };
    dirty = true;
  }

  function applyRange(m) {
    const f = script.fault;
    if (f === 'ignore') return;
    const { translateFormula, moveFormula } = require('../../src/adapters/xlsx/shift.cjs');
    const src = live.find(x => x.name === m.sheet);
    const dst = live.find(x => x.name === m.targetSheet);
    if (!src || !dst) return;
    const pairs = f === 'partial' ? m.pairs.slice(0, 1) : m.pairs;
    const snap = Object.fromEntries(pairs.map(({ from }) => [from, src.cells[from] ? { ...src.cells[from] } : undefined]));
    if (m.action === 'move_range' && f !== 'move-keeps-source') for (const { from } of pairs) delete src.cells[from];
    // M7 E R7: the Cut re-points every reference ENTIRELY inside the block, on every sheet and in the names, as the tier
    // E probe measured (shift.cjs moveFormula); fault 'move-no-repoint' leaves the OUTSIDE references and names as they were
    const g = parseRange(m.range);
    const mv = { from: { sheet: m.sheet, r1: g.r1, c1: g.c1, r2: g.r2, c2: g.c2 }, to: { sheet: m.targetSheet }, dr: m.delta.dr, dc: m.delta.dc };
    const repoint = m.action === 'move_range';
    if (repoint && f !== 'move-no-repoint') {
      for (const ws of live) for (const c of Object.values(ws.cells)) if (c.f) c.f = moveFormula(c.f, ws.name, mv).formula;
      liveNames = liveNames.map(n => ({ ...n, ref: moveFormula(`=${n.ref}`, '', mv).formula.slice(1) }));
    }
    for (const { from, to } of pairs) {
      const c = snap[from];
      if (!c) { delete dst.cells[to]; continue; }
      // L-06 r8: a CROSS-SHEET copy dispatched WITHOUT the editor's Copy (meta.crossSheetWrite) writes values / formulas only -- the cell FORMAT is not carried
      const moved = m.crossSheetWrite === true ? (c.f ? { v: c.v, f: c.f } : { v: c.v }) : { ...c };
      if (m.action === 'copy_range' && c.f && f !== 'copy-no-translate') moved.f = translateFormula(c.f, m.delta.dr, m.delta.dc).formula;
      if (repoint && c.f) moved.f = moveFormula(c.f, m.sheet, mv).formula;
      // L-06 r6 'copy-xsheet-shows-literal': onto ANOTHER sheet the editor shows (and saves) the source text untranslated (zero l066: B4 '= B2+1')
      if (m.action === 'copy_range' && c.f && f === 'copy-xsheet-shows-literal' && m.targetSheet !== m.sheet) moved.f = c.f;
      dst.cells[to] = moved;
      // an explicit formula write after the Copy (meta.formulaWrites: the planner's translation) sets the cell's formula text as written
      if (m.action === 'copy_range' && m.explicitFormulaWrite === true && m.formulaWrites && typeof m.formulaWrites[to] === 'string') dst.cells[to] = { ...dst.cells[to], f: m.formulaWrites[to] };
      // L-06 r8 'cross-copy-delayed' (zero l068): an editor Copy onto ANOTHER sheet leaves a delayed state that overwrites a formula written after it at the next save
      if (m.action === 'copy_range' && c.f && f === 'cross-copy-delayed' && m.targetSheet !== m.sheet && m.crossSheetWrite !== true) literalOnSave[`${dst.name}\u0000${to}`] = { f: c.f, v: c.v, shown: dst.cells[to].f };
      if (m.action === 'copy_range' && c.f && f === 'copy-saves-literal') {
        if (m.explicitFormulaWrite === true) delete literalOnSave[`${dst.name}\u0000${to}`];
        else literalOnSave[`${dst.name}\u0000${to}`] = { f: c.f, v: c.v, shown: moved.f };
      }
    }
    if (m.action === 'copy_range' && f === 'copy-corrupts-source') src.cells[pairs[0].from] = { v: 'CORRUPTED' };
    if (f === 'side-effect') { const o = live.find(x => x !== src) || src; o.cells.Z99 = { v: 'side effect' }; }
    // M7 S1 L3 negative controls: a dependent changes as a side effect of the copy / move (on a sheet not in the op)
    if (f === 'range-side-name') liveNames = [...liveNames, { name: 'SideName', ref: `${src.name}!$Z$1` }];
    if (f === 'range-side-rule') { const o = live.find(x => (x.rules || []).length) || src; o.rules = (o.rules || []).map((r, i) => (i === 0 ? { ...r, formula: `${r.formula}+0` } : r)); }
    if (f === 'range-side-chart') { const o = live.find(x => x.objects && (x.objects.charts || []).length); if (o) o.objects = { ...o.objects, charts: o.objects.charts.map(c => ({ ...c, series: c.series.map(x => ({ ...x, values: `${x.values}` === '' ? x.values : x.values.replace(/\$(\d+)$/, (_, n) => `$${Number(n) + 1}`) })) })) }; }
    if (f === 'range-side-filter') { const o = live.find(x => x.autoFilter) || src; o.autoFilter = { range: 'A1:A9', filters: [] }; }
    dirty = true;
  }

  function applySort(m) {
    const f = script.fault;
    if (f === 'ignore') return;
    const ws = live.find(x => x.name === m.sheet);
    if (!ws) return;
    const { c1, c2, first, r2 } = m.grid;
    const rows = [];
    for (let r = first; r <= r2; r += 1) { const t = []; for (let c = c1; c <= c2; c += 1) t.push(ws.cells[addr(c, r)]); rows.push(t); }
    const k = m.keyColumn - c1;
    const valAt = (t, kk) => (t[kk] && t[kk].v !== null ? t[kk].v : '');
    const val = t => valAt(t, k);
    const cmpAt = kk => (a, b) => { const x = valAt(a, kk); const y = valAt(b, kk); if (typeof x === 'number' && typeof y === 'number') return x - y; return String(x) < String(y) ? -1 : String(x) > String(y) ? 1 : 0; };
    // M7 S5 R21 (W6 probe): blank keys go last; a formula moves with its row and points at its new row
    // (fault 'sort-formula-stays': it keeps its old row's references)
    const blank = t => val(t) === '';
    const desc = (m.order === 'desc') !== (f === 'sort-reverse');
    // GH #301 D3: the secondary keys (m.keys[1..]) break the primary key's ties, in order (fault 'sort-primary-only':
    // they are ignored -- ties stay in source order)
    const rest = (m.keys || []).slice(1).filter(() => f !== 'sort-primary-only').map(x => ({ c: cmpAt(x.column - c1), desc: x.order === 'desc' }));
    const tie = (a, b) => { for (const r of rest) { const c = r.desc ? r.c(b, a) : r.c(a, b); if (c) return c; } return 0; };
    const sorted = rows.map((t, i) => [t, i, first + i]).sort((a, b) => ((blank(a[0]) - blank(b[0])) * (f === 'sort-blanks-first' ? -1 : 1)) || (desc ? cmpAt(k)(b[0], a[0]) : cmpAt(k)(a[0], b[0])) || tie(a[0], b[0]) || a[1] - b[1]);
    sorted.forEach(([t, , from], i) => {
      for (let c = c1; c <= c2; c += 1) {
        if (f === 'sort-key-only' && c - c1 !== k) continue;
        const a = addr(c, first + i);
        const cell = t[c - c1];
        const moved = cell && cell.f && f !== 'sort-formula-stays' ? { ...cell, f: cell.f.replace(new RegExp(`([A-Z]{1,3})${from}(?![0-9])`, 'g'), `$1${first + i}`) } : cell;
        if (moved) ws.cells[a] = moved; else delete ws.cells[a];
      }
    });
    if (f === 'side-effect') live.find(x => x !== ws).cells.Z99 = { v: 'side effect' };
    dirty = true;
  }

  const vOverlap = (a, b) => { const x = parseRange(a); const y = parseRange(b); return !!x && !!y && x.c1 <= y.c2 && y.c1 <= x.c2 && x.r1 <= y.r2 && y.r1 <= x.r2; };
  // L-22a: xlValidateTime's measured quirk (the probe, test/live/xlsx-l22-validation-types-probe.live.cjs) --
  // the editor echoes formula1/formula2 back as LOCALE AM/PM TEXT, never the day fraction it was written with.
  // The fake simulates exactly this one quirk (every other type's formula echoes back literally, unchanged).
  const fractionToAmPmText = frac => {
    const totalSec = Math.round(Number(frac) * 86400);
    const h = Math.floor(totalSec / 3600) % 24;
    const min = Math.floor((totalSec % 3600) / 60);
    const s = totalSec % 60;
    const h12 = h % 12 === 0 ? 12 : h % 12;
    const pad = n => String(n).padStart(2, '0');
    return `${h12}:${pad(min)}:${pad(s)} ${h < 12 ? 'AM' : 'PM'}`;
  };
  // L-22d: a date serial echoed back as the editor's LOCALE text N/N/YYYY (script.dateLocale 'MD' = the measured US shape, default; 'DM' = a day-first locale)
  const serialToDateText = v => {
    const d = new Date(Date.UTC(1899, 11, 30) + Number(v) * 86400000);
    const mo = d.getUTCMonth() + 1; const da = d.getUTCDate(); const y = d.getUTCFullYear();
    return script.dateLocale === 'DM' ? `${da}/${mo}/${y}` : `${mo}/${da}/${y}`;
  };
  function applyValidation(m) {
    const f = script.fault;
    if (f === 'ignore') return;
    const ws = live.find(x => x.name === m.sheet);
    if (!ws) return;
    ws.rules = (ws.rules || []).filter(r => !(r.kind === 'validation' && vOverlap(r.range, m.range)));
    if (m.action === 'set_validation') {
      const isTime = m.want.type === 'xlValidateTime';
      const isDate = m.want.type === 'xlValidateDate';
      const encode = v => (isTime && v !== '' ? fractionToAmPmText(v) : (isDate && v !== '' ? serialToDateText(v) : v));
      const formula2Want = f === 'validation-wrong-bounds' ? String(Number(m.want.formula2) + (isTime ? 1 / 86400 : 1)) : m.want.formula2;
      // 'validation-date-garbled-echo': the editor's locale text is NOT the measured N/N/YYYY shape -- dateReadSerials() must name this a gap.
      // 'validation-time-garbled-echo': the editor's locale text is NOT the measured AM/PM shape (e.g. a
      // different locale) -- timeReadFraction() must name this a gap, never guess a match or a difference.
      // L-22b 'validation-custom-formula-wrong': the stored formula1 differs from what was requested (a
      // different absolute formula lands) -- sameFormula() must catch this via L1, same as every other type's
      // wrong-value RED case.
      // L-22d v2: date carries TWO independent encodings -- the public locale text (formula/formula2, what
      // the real editor's GetFormula1/2 returns) and the bridge's own raw serial (bridgeFormula/bridgeFormula2,
      // what the real DS Bridge range.validations.read/v1 row carries, dependents() below projects it). They
      // are read from DIFFERENT fault knobs so a test can make either one wrong independently:
      //   'validation-date-bridge-wrong'  the BRIDGE serial is off by one, the public text still matches the request
      //   'validation-date-text-conflict' the PUBLIC text names a different date, the bridge serial still matches
      //   'validation-date-garbled-echo'  (unchanged) only the public text is not the measured N/N/YYYY shape at all
      const formula1 = f === 'validation-time-garbled-echo' ? 'not-a-time'
        : f === 'validation-date-garbled-echo' && isDate ? 'not-a-date'
        : f === 'validation-date-text-conflict' && isDate ? serialToDateText(Number(m.want.formula1) + 10)
        : f === 'validation-custom-formula-wrong' && m.want.type === 'xlValidateCustom' ? '=$Z$99>0'
        : encode(m.want.formula1);
      const formula2 = f === 'validation-time-garbled-echo' && formula2Want !== '' ? 'not-a-time-2'
        : f === 'validation-date-garbled-echo' && isDate && formula2Want !== '' ? 'not-a-date-2'
        : f === 'validation-date-text-conflict' && isDate && formula2Want !== '' ? serialToDateText(Number(formula2Want) + 10)
        : encode(formula2Want);
      // the bridge mirrors whatever was ACTUALLY requested to be written, including 'validation-wrong-bounds'
      // (formula2Want, not m.want.formula2) -- only 'validation-date-bridge-wrong' deliberately mis-stores it.
      const bridgeFormula1 = f === 'validation-date-bridge-wrong' && isDate ? String(Number(m.want.formula1) + 1) : m.want.formula1;
      const bridgeFormula2 = f === 'validation-date-bridge-wrong' && isDate && formula2Want !== '' ? String(Number(formula2Want) + 1) : formula2Want;
      const add = range => ({ kind: 'validation', type: 1, range, formula: formula1, formula2, api: { type: m.want.type, alertStyle: m.want.alertStyle, operator: m.want.operator },
        ...(isDate ? { bridgeFormula: bridgeFormula1, bridgeFormula2: bridgeFormula2 } : {}) });
      ws.rules.push(add(m.range));
      if (f === 'validation-leaks') ws.rules.push(add(script.leakRange));
    }
    if (f === 'side-effect') live.find(x => x !== ws).cells.Z99 = { v: 'side effect' };
    dirty = true;
  }
  const firstCellOf = range => { const g = parseA1Range(range); return g ? a1Addr(g.c1, g.r1) : range; };
  function validationReadsOf(asked) {
    const out = {};
    for (const w of asked) {
      if (script.validationReadUnobservable) { out[`${w.sheet}!${w.range}`] = 'UNOBSERVABLE:threw'; continue; }
      const ws = live.find(x => x.name === w.sheet);
      const r = ws && (ws.rules || []).find(x => x.kind === 'validation' && vOverlap(x.range, w.range));
      if (!r) { out[`${w.sheet}!${w.range}`] = { absent: true }; continue; }
      // L-22e: the editor shows a custom formula AS IF WRITTEN AT the rule's own sheet's CURRENT active
      // cell (measured, xlsx-l22e-custom-relative-ref-probe.live.cjs r1-r3) -- the fake's own default is
      // "no quirk" (activeCell happens to equal the rule's own anchor, so the echo is the exact stored
      // text, unchanged from every other type). 'validation-custom-read-shifted-from-elsewhere' simulates
      // the measured defect for real: the active cell is a FIXED, unrelated cell, and formula1 is shifted
      // from the rule's own anchor to THAT cell, exactly as the real editor was measured to do -- so the
      // RED/GREEN this exercises proves the adapter's reAnchoredFormula() un-shifts it correctly, not
      // just that the fake happens to already agree with the stored text.
      let formula1 = r.formula;
      let activeCell = firstCellOf(r.range);
      if (script.fault === 'validation-custom-read-shifted-from-elsewhere' && r.api && r.api.type === 'xlValidateCustom') {
        const elsewhere = 'A1';
        const from = parseA1Range(activeCell); const to = parseA1Range(elsewhere);
        formula1 = translateFormula(formula1, to.r1 - from.r1, to.c1 - from.c1).formula;
        activeCell = elsewhere;
      }
      // L-22e: GetActiveCell() itself unobservable (the driver's own UNOBSERVABLE: sentinel) -- the formula
      // text is unaffected, only the reference cell needed to re-anchor it is missing.
      if (script.fault === 'validation-custom-active-cell-unobservable' && r.api && r.api.type === 'xlValidateCustom') activeCell = 'UNOBSERVABLE:active-cell-threw';
      out[`${w.sheet}!${w.range}`] = { ...(r.api || { type: 'xlValidateList', alertStyle: 'xlValidAlertStop', operator: 'xlBetween' }), formula1, formula2: r.formula2 || '', activeCell };
    }
    return out;
  }

  // M7 X-O01: the three new types' own identity fields (no operator/formula/fill) -- mirrors cf.cjs's
  // own CF_IDENTITY_FIELDS table, so the fake's write/read/identity logic stays in lock-step with the real one.
  const CF_OWN_FIELDS = {
    xlAboveAverageCondition: ['aboveBelow', 'numStdDev'],
    xlTop10: ['rank', 'percent', 'topBottom'],
    xlUniqueValues: [],
    // M7 X-O01: no type-level identity field; X-O01 appearance keys join per rule (rule.idKeys)
    xlColorScale: [],
    xlDatabar: [],
    xlIconSets: [],
  };
  // X-O01: the appearance fields the real reader returns for these types (CF_SCALAR_FIELDS + stopColors), with the
  // engine defaults the LIVE phase-1 probe read; a rule's identity is the appearance keys it asked for (rule.idKeys)
  const CF_APPEARANCE_DEFAULTS = {
    xlDatabar: { barColor: 6524614, barFillType: 'xlDataBarFillGradient', direction: 'xlContext', showValue: true, axisPosition: 'xlDataBarAxisAutomatic' },
    xlIconSets: { iconSet: 'xl3TrafficLights1', showIconOnly: false, reverseOrder: false },
    xlColorScale: { stopColors: JSON.stringify([16279915, 16771972, 6536827]) },
  };
  function applyCf(m) {
    const f = script.fault;
    if (f === 'ignore') return;
    const ws = live.find(x => x.name === m.sheet);
    if (!ws) return;
    const r = m.rule;
    const own = CF_OWN_FIELDS[r.type] && [...CF_OWN_FIELDS[r.type], ...(r.idKeys || [])];
    const mine = x => x.kind === 'conditional' && x.range === m.range && x.type === r.type
      && (own ? own.every(k => x[k] === r[k]) : x.operator === r.operator && String(x.formula) === r.formula1 && (x.fill === undefined ? null : x.fill) === r.fill);
    ws.rules = ws.rules || [];
    if (f === 'cf-removes-other') ws.rules = ws.rules.filter(x => !(x.kind === 'conditional' && vOverlap(x.range, m.range)));
    if (f === 'cf-alters-other') ws.rules = ws.rules.map(x => {
      if (x.kind !== 'conditional' || mine(x)) return x;
      if (x.formula !== undefined) return { ...x, formula: `${x.formula}+0` };
      if (x.type === 'xlAboveAverageCondition') return { ...x, aboveBelow: !x.aboveBelow };
      if (x.type === 'xlTop10') return { ...x, rank: x.rank + 1 };
      return x; // xlUniqueValues has no scalar to perturb; other faults cover it
    });
    // 'cf-wrong-value' corrupts EVERY own field; 'cf-wrong-one-field' (script.cfWrongField) corrupts
    // just the NAMED field -- surgical enough to prove each identity field individually matters.
    const wrongFields = f === 'cf-wrong-value' ? own : (f === 'cf-wrong-one-field' ? [script.cfWrongField] : []);
    const flip = v => (typeof v === 'boolean' ? !v : (typeof v === 'number' ? v + 1 : `${v}-wrong`));
    if (m.action === 'add_conditional_format') {
      const add = own
        ? { kind: 'conditional', type: r.type, range: m.range, ...(CF_APPEARANCE_DEFAULTS[r.type] || {}), ...Object.fromEntries(own.map(k => [k, wrongFields.includes(k) ? flip(r[k]) : r[k]])) }
        : { kind: 'conditional', type: r.type, operator: r.operator, formula: r.formula1, ...(r.formula2 ? { formula2: f === 'cf-wrong-formula2' ? `${r.formula2}0` : r.formula2 } : {}), fill: f === 'cf-wrong-fill' ? (r.fill === null ? 1 : r.fill + 1) : (r.fill === null ? undefined : r.fill), range: m.range };
      ws.rules.push(add);
      if (f === 'cf-duplicate') ws.rules.push({ ...add });
      if (f === 'cf-unique-unidentifiable' && r.type === 'xlUniqueValues') ws.rules[ws.rules.length - 1] = { kind: 'conditional', type: '', range: m.range };
    } else {
      const i = ws.rules.findIndex(mine);
      if (i >= 0) ws.rules.splice(i, 1);
    }
    if (f === 'side-effect') live.find(x => x !== ws).cells.Z99 = { v: 'side effect' };
    dirty = true;
  }
  function cfReadsOf(asked) {
    const out = {};
    for (const w of asked) {
      const ws = live.find(x => x.name === w.sheet);
      if (script.cfReadUnobservable) { out[`${w.sheet}!${w.range}`] = 'UNOBSERVABLE:threw'; continue; }
      const matching = (ws.rules || []).filter(x => x.kind === 'conditional' && vOverlap(x.range, w.range));
      if (matching.some(x => x.type === '')) { out[`${w.sheet}!${w.range}`] = 'UNOBSERVABLE:cf-type-empty'; continue; }
      out[`${w.sheet}!${w.range}`] = matching.map(x => {
        const own = CF_OWN_FIELDS[x.type];
        const look = CF_APPEARANCE_DEFAULTS[x.type];
        if (own) return { type: x.type, appliesTo: x.range, ...Object.fromEntries(own.map(k => [k, x[k]])), ...(look ? Object.fromEntries(Object.keys(look).map(k => [k, x[k] === undefined ? look[k] : x[k]])) : {}) };
        return { type: x.type || 'xlExpression', operator: x.operator || null, formula1: x.formula, formula2: x.formula2 || '', fill: x.fill === undefined ? null : x.fill, appliesTo: x.range };
      });
    }
    return out;
  }

  function applyFilter(m) {
    const f = script.fault;
    if (f === 'ignore') return;
    const ws = live.find(x => x.name === m.sheet);
    if (!ws) return;
    if (m.action === 'filter_range') {
      const af = ws.autoFilter && ws.autoFilter.range === m.range ? ws.autoFilter : { range: m.range, filters: [] };
      // M7 R19: values -> xlOr =a / =b; and -> xlAnd c1 / c2 (the probe's read-back)
      const wrong = f === 'filter-wrong-value';
      // GH #301 D2 (AF-02): 3+ values -> xlFilterValues + the array; top -> xlTop10Items n; dynamic -> xlFilterDynamic constant
      const DYN = { aboveAverage: 'xlFilterAboveAverage', belowAverage: 'xlFilterBelowAverage' };
      const crit = m.values && m.values.length > 2 ? { operator: 'xlFilterValues', criteria1: wrong ? m.values.slice(0, -1) : [...m.values], criteria2: null }
        : m.top ? { operator: 'xlTop10Items', criteria1: wrong ? m.top + 1 : m.top, criteria2: null }
          : m.dynamic ? { operator: 'xlFilterDynamic', criteria1: wrong ? 'xlFilterToday' : DYN[m.dynamic], criteria2: null }
            : m.values ? { operator: 'xlOr', criteria1: `=${m.values[0]}`, criteria2: `=${m.values[1]}` } : (m.and ? { operator: 'xlAnd', criteria1: m.and[0], criteria2: m.and[1] } : { operator: 'xlOr', criteria1: `=${wrong ? `${m.value}!` : m.value}`, criteria2: null });
      if (wrong && m.values && m.values.length === 2) crit.criteria2 = `=${m.values[1]}!`;
      if (f === 'filter-wrong-operator') crit.operator = crit.operator === 'xlAnd' ? 'xlOr' : 'xlAnd';
      af.filters = af.filters.filter(x => x.field !== m.field && f !== 'filter-drops-other').concat([{ field: f === 'filter-wrong-field' ? m.field + 1 : m.field, on: true, ...crit }]);
      if (f === 'filter-extra-field') af.filters.push({ field: m.field + 1, on: true, operator: 'xlOr', criteria1: '=extra', criteria2: null });
      ws.autoFilter = af;
    } else if (ws.autoFilter) ws.autoFilter = { ...ws.autoFilter, filters: ws.autoFilter.filters.map(x => ({ ...x, on: false })) };
    if (f === 'side-effect') live.find(x => x !== ws).cells.Z99 = { v: 'side effect' };
    dirty = true;
  }
  function applyCharts(m) {
    const f = script.fault;
    if (f === 'ignore') return;
    const ws = live.find(x => x.name === m.sheet);
    if (!ws) return;
    ws.objects = ws.objects || { charts: [], pivots: [], names: [] };
    const list = ws.objects.charts;
    if (m.action === 'set_chart') {
      const series = m.series.map(x => ({ name: `=${x.name}`, values: `=${f === 'chart-wrong-range' ? x.values.replace(/\d+$/, d => String(Number(d) - 1)) : x.values}`, cat: `=${x.cat}`, x: '' }));
      // M7 S5 C1 L1 negative: 'chart-wrong-series' -- the editor builds a different series set (the last one missing)
      // X-O03: the requested type, read from meta (set by charts.cjs's lower()) -- previously
      // hardcoded 'bar', which silently hid any other type ever reaching this far.
      // X-O04: a newly created chart's legend defaults to 'bottom' (measured: an unsupported
      // write leaves it at this default, i.e. this IS the editor's own untouched default).
      const chart = { name: m.name, type: m.type, title: f === 'chart-wrong-title' ? `${m.title}!` : m.title, legend: 'bottom', series: f === 'chart-wrong-series' ? series.slice(0, -1) : series };
      list.push(chart);
      if (f === 'chart-duplicate' || f === 'chart-duplicate-drops-neighbour') list.push(clone(chart));
      if (f === 'chart-duplicate-drops-neighbour') list.splice(list.findIndex(c => c.name !== m.name), 1);
      if (f === 'chart-extra-unnamed') list.push({ ...clone(chart), name: null });
    } else if (m.action === 'set_chart_title') {
      // M7 C3 (tier D probe): the title reads back with a trailing CR LF; fault 'title-drops-series'
      const hit = list.find(c => c.name === m.name);
      if (hit) { hit.title = `${m.title}\r\n`; if (f === 'title-drops-series') hit.series = hit.series.slice(0, -1); }
    } else if (m.action === 'set_chart_axis_title') {
      // X-O05 (LIVE MEASURED): same CRLF shape as set_chart_title, keyed by the requested axis.
      // Faults: 'axis-title-wrong' (a different text lands), 'axis-title-drops-other-axis' (the OTHER axis clears).
      const hit = list.find(c => c.name === m.name);
      if (hit) {
        const key = m.axis === 'horizontal' ? 'horAxisTitle' : 'verAxisTitle';
        const other = m.axis === 'horizontal' ? 'verAxisTitle' : 'horAxisTitle';
        hit[key] = `${f === 'axis-title-wrong' ? `${m.title}!` : m.title}\r\n`;
        if (f === 'axis-title-drops-other-axis') hit[other] = undefined;
      }
    } else if (m.action === 'set_chart_legend') {
      // X-O04: measured, SetLegendPos rejects an unsupported position by returning false, without
      // changing it -- fault 'legend-rejected' simulates that (the position stays as it was).
      const hit = list.find(c => c.name === m.name);
      if (hit && f !== 'legend-rejected') { hit.legend = m.legend; if (f === 'legend-drops-series') hit.series = hit.series.slice(0, -1); }
    } else if (m.action === 'move_chart') {
      // X-O06-6: moves the chart's pos (read back through drawingReadsOf's chart mapper, which
      // defaults to CHART_GEOM.pos when unset -- setting it explicitly here is what makes the
      // move observable). Fault 'move-chart-wrong-position' lands one column over.
      // Fault 'move-chart-size-drifts' ALSO perturbs the SAME chart's own width by 1 EMU,
      // simulating the real ±1 EMU drift X-O08 measured -- proves `sizeBlind` absorbs it (L3
      // must stay GREEN; a self-check in the test file flips `sizeBlind` off to prove it would
      // otherwise FAIL).
      const hit = list.find(c => c.name === m.name);
      if (hit) {
        hit.pos = { fromCol: f === 'move-chart-wrong-position' ? m.at.col + 1 : m.at.col, colOffset: 0, fromRow: m.at.row, rowOffset: 0 };
        if (f === 'move-chart-size-drifts') hit.w = (hit.w === undefined ? CHART_GEOM.w : hit.w) - 1;
      }
    } else if (m.action === 'set_chart_style') {
      // X-O04/1+2: the six ORIGINAL targets are WRITE-ONLY, no readable field for the style itself --
      // nothing to mutate on the fake chart record normally. The one observable fault
      // ('style-setter-rejected') is handled at the dispatch layer above, BEFORE applyCharts is even
      // called. 'style-alters-legend' is a SABOTAGE fault (an unrelated field changes as a side
      // effect of the style write) -- this IS observable (the family's own PRIMITIVE.set_chart_style
      // "kept" check reads it back), proving that check actually fires, not just that it exists.
      const hit = list.find(c => c.name === m.name);
      if (hit && f === 'style-alters-legend') hit.legend = hit.legend === 'bottom' ? 'top' : 'bottom';
      // X-O04/11: marker/gridlines ARE readable (through the chartReadsOf markerStyle/gridlinesColor
      // fields above) -- a successful dispatch records the written colour, the SAME round-trip the
      // real X-O04/10 bridge read proves end-to-end. 'style-bridge-stale' is a SABOTAGE fault: the
      // dispatch succeeds but the bridge-read value does NOT change, proving classifyStyle/PRIMITIVE
      // actually read-and-compare rather than trusting the dispatch outcome.
      if (hit && m.target === 'marker' && f !== 'style-bridge-stale') {
        const s = hit.series[m.seriesIndex];
        if (s) {
          if (m.styleFill !== null) s.markerFill = m.styleFill;
          if (m.styleOutline !== null) s.markerOutline = m.styleOutline.color;
        }
      }
      if (hit && m.target === 'gridlines' && f !== 'style-bridge-stale') {
        hit.gridlinesColor = { ...(hit.gridlinesColor || {}), [`${m.axis}Gridlines`]: m.styleOutline.color };
      }
      // X-O05/3: gapOverlap is chart-level (no seriesIndex) -- only the requested field(s) move,
      // same "only what was asked" convention markerFill/markerOutline above already follow.
      if (hit && m.target === 'gapOverlap' && f !== 'style-bridge-stale') {
        hit.gapOverlap = { ...(hit.gapOverlap || {}), ...(m.styleGapWidth !== null ? { gapWidth: m.styleGapWidth } : {}), ...(m.styleOverlap !== null ? { overlap: m.styleOverlap } : {}) };
      }
      // X-O05/3: seriesOrder writes the WHOLE permutation in one call -- every series' own order
      // moves together, mirroring the real handler's one-setOrder()-call-per-series loop.
      if (hit && m.target === 'seriesOrder' && f !== 'style-bridge-stale') {
        m.styleOrder.forEach((v, i) => { if (hit.series[i]) hit.series[i].order = v; });
      }
    } else if (m.action === 'delete_chart') {
      // X-O07/3: removes the WHOLE chart record -- the generic `f === 'ignore'` fault above
      // (dispatch succeeds, nothing actually happens) already covers the RED case, same
      // convention every other intent in this family shares; no delete-specific fault needed.
      const idx = list.findIndex(c => c.name === m.name);
      if (idx >= 0) list.splice(idx, 1);
    } else if (m.action === 'copy_chart') {
      // X-O07/4b: v1 copies type/series/title/legend/axis-titles only (never the style targets --
      // charts.cjs's own COPY_CHART_ALWAYS_OMITTED names them, unconditionally, in the receipt).
      // Series are the SOURCE's own records, verbatim (same underlying data, not re-derived from
      // meta.series which stays null for this action -- see charts.cjs's lower() comment on why).
      const srcWs = live.find(x => x.name === m.sourceSheet);
      const srcList = (srcWs && srcWs.objects && srcWs.objects.charts) || [];
      const src = srcList.find(c => c.name === m.sourceName);
      if (!src) return;
      // Fault 'copy-wrong-series' (M7-style negative): the recreated chart's series drop its last entry.
      const series = f === 'copy-wrong-series' ? src.series.slice(0, -1) : src.series.map(clone);
      const chart = { name: m.name, type: m.type, title: m.title, legend: m.legend, series, pos: { fromCol: m.at.col, colOffset: 0, fromRow: m.at.row, rowOffset: 0 } };
      if (m.copyHorAxisTitle) chart.horAxisTitle = `${m.copyHorAxisTitle}\r\n`;
      if (m.copyVerAxisTitle) chart.verAxisTitle = `${m.copyVerAxisTitle}\r\n`;
      // GH #302 G6: the copy carries the chart-level flags and every POINT label (a fresh chart: every point SERIES, all off);
      // fault 'copy-drops-point-labels' -- the point calls do nothing (the G6 RED for the point half)
      if (m.copyLabels) chart.dataLabels = { ...m.copyLabels };
      if (src.effectiveLabels && typeof src.effectiveLabels === 'object') {
        chart.effectiveLabels = Object.fromEntries(Object.keys(src.effectiveLabels).map(k => [k, { level: 'SERIES', showSerName: false, showCatName: false, showVal: false, showPercent: false }]));
        if (f !== 'copy-drops-point-labels') for (const x of m.copyPoints || []) chart.effectiveLabels[x.key] = { level: 'POINT', ...x.flags };
      }
      // Fault 'copy-alters-source': a SABOTAGE fault -- the source chart changes as a side effect
      // of copying it, proving verifyPreservation's "every OTHER chart kept" check actually fires.
      if (f === 'copy-alters-source') src.title = `${src.title}!`;
      list.push(chart);
    } else {
      const hit = list.find(c => c.name === m.name);
      if (hit && f === 'rename-copies') list.push({ ...clone(hit), name: m.newName });
      else if (hit) { hit.name = m.newName; if (f === 'rename-drops-series') hit.series = hit.series.slice(0, -1); }
    }
    if (f === 'chart-alters-other') { const o = live.flatMap(x => (x.objects ? x.objects.charts : [])).find(c => c.name !== m.name && c.name !== m.newName); if (o) o.title = `${o.title}~`; }
    if (f === 'side-effect') live.find(x => x !== ws).cells.Z99 = { v: 'side effect' };
    dirty = true;
  }
  /**
   * M11.5b.4 B10 Bridge WRITE: what the REAL driver returns (xlsx-driver b10DispatchResult) for the body's outcome.
   * script.b10Write: undefined (accepted) | 'gate-refused' (the client gate, before any mutation) | 'error-none' (setter
   * missing: nothing written) | 'partial' (the first setter landed, a later one threw -- fit / scale: fitW only) |
   * 'error-applied' (everything written, then an internal throw) | 'malformed' (written, the receipt shape is wrong).
   */
  function b10Write(m) {
    const mode = script.b10Write;
    const fail = (evidence, partiallyApplied = false) => (partiallyApplied
      ? { ok: true, dispatched: true, partiallyApplied: true, outcome: 'b10-bridge-write-partial', evidence, error: null }
      : { ok: false, dispatched: true, outcome: 'b10-bridge-write-failed', evidence, error: null });
    if (mode === 'gate-refused') return fail({ ok: false, stage: 'bridge-write-gate', reason: script.b10GateReason || 'CAPABILITY_MISSING', mutation: 'NONE' });
    if (mode === 'error-none') return fail({ ok: false, stage: 'bridge-write', reason: 'BRIDGE:asc_setFitToWidth-missing', uncertain: false, partial: false, applied: [] });
    if (mode === 'partial') {
      const ws = live.find(x => x.name === m.sheet);
      ws.print = { ...PRINT_DEFAULT, ...clone(ws.print || {}), fitW: 0 }; calls.bridgeWrites += 1; dirty = true;
      return fail({ ok: false, stage: 'bridge-write', reason: 'BRIDGE:asc_setFitToHeight-threw', uncertain: false, partial: true, applied: ['fitW'] }, true);
    }
    applyPrint(m); calls.bridgeWrites += 1;
    if (mode === 'error-applied') return fail({ ok: false, stage: 'bridge-write', reason: 'BRIDGE:internal-threw', uncertain: false, partial: false, applied: [] });
    if (mode === 'malformed') return fail({ ok: false, stage: 'bridge-write', reason: 'UNEXPECTED_SHAPE', uncertain: true, partial: false, applied: [] });
    drawingFault();
    return { ok: true, dispatched: true, outcome: 'command-completed', error: null };
  }
  // Print settings (wave 8, meta.family 'print'): a sheet may carry print { fitW, fitH, fitFlag, oddHeader, rowBreaks,
  // useFirstPage, firstPage, hCenter }; read for `printSheets`; script.printBuild false -> 'UNOBSERVABLE:build-not-
  // verified'. Faults 'ignore', 'fit-no-flag' (the numbers land, the flag stays off), 'fit-flag-stuck-on' (X-P,
  //: the 0/0 OFF shape lands but the flag incorrectly stays on), 'print-wrong-value',
  // 'print-side' (another print field of the sheet changes too), 'break-drops-other' (an existing break goes),
  // 'fit-wrong-height', 'break-not-manual' (man false), 'first-page-unused' (the number without its use flag), 'side-effect'.
  const PRINT_DEFAULT = { fitW: 0, fitH: 0, fitFlag: false, oddHeader: null, rowBreaks: [], useFirstPage: false, firstPage: null, hCenter: false, oddFooter: null, vCenter: false, colBreaks: [], scale: 100, width: 210, height: 297, firstHeader: null, firstFooter: null, evenHeader: null, evenFooter: null, diffFirst: false, diffOddEven: false };
  function applyPrint(m) {
    const f = script.fault;
    if (f === 'ignore') return;
    const ws = live.find(x => x.name === m.sheet);
    if (!ws) return;
    const p = { ...PRINT_DEFAULT, ...clone(ws.print || {}) };
    const w = m.want;
    const bump = x => (f === 'print-wrong-value' ? x + 1 : x);
    if (m.action === 'set_print_fit') {
      p.fitW = bump(w.fitW); p.fitH = f === 'fit-wrong-height' ? w.fitH + 1 : w.fitH;
      // X-P: 0/0 expects the flag OFF, any other value expects it ON.
      const expectOn = w.fitW !== 0 || w.fitH !== 0;
      p.fitFlag = f === 'fit-no-flag' ? false : (f === 'fit-flag-stuck-on' ? true : expectOn);
      // owner 1/a: OFF resets the scale to 100 ('fit-scale-stuck' keeps it); 'print-side-scale' moves it on any fit step
      if (!expectOn && f !== 'fit-scale-stuck') p.scale = 100;
      if (f === 'print-side-scale') p.scale = 55;
    }
    // OD-A / OD-C: a scale step = fit OFF + the scale (the UI's asc_SetPrintScale(0, 0, n)); a paper step = width / height
    if (m.action === 'set_print_scale') { p.fitW = 0; p.fitH = 0; p.fitFlag = false; p.scale = bump(w.scale); }
    if (m.action === 'set_paper_size') { p.width = bump(w.width); p.height = f === 'paper-wrong-height' ? w.height + 1 : w.height; if (f === 'paper-side-scale') p.scale = 55; }
    if (m.action === 'set_header_footer') p[w.slot] = f === 'print-wrong-value' ? `${w.value}!` : w.value;
    // X-P07: a flag step sets exactly its flag ('print-wrong-value' flips it the wrong way)
    if (m.action === 'set_header_footer_flag') p[w.flag] = f === 'print-wrong-value' ? !w.value : w.value;
    if (m.action === 'set_page_break') {
      const others = p[w.breaks].filter(b => b.id !== w.breakId);
      // M7 F2: a removal takes exactly its break ('break-remove-ignored' leaves it); an add keeps the others
      p[w.breaks] = w.remove ? (f === 'break-remove-ignored' ? p[w.breaks] : (f === 'break-drops-other' ? [] : others))
        : [...(f === 'break-drops-other' ? [] : p[w.breaks]), { id: bump(w.breakId), man: f !== 'break-not-manual' }];
      // 'break-adds-other': one more break appears; 'break-moves-other': another break's id shifts (the count stays)
      if (f === 'break-adds-other') p[w.breaks] = [...p[w.breaks], { id: 99, man: true }];
      if (f === 'break-moves-other') p[w.breaks] = p[w.breaks].map((b, i) => (i === 0 && b.id !== w.breakId ? { ...b, id: b.id + 1 } : b));
    }
    if (m.action === 'set_first_page_number') {
      if (w.firstPage === null) { p.firstPage = f === 'print-wrong-value' ? 1 : null; if (f === 'first-page-flag') p.useFirstPage = !p.useFirstPage; }
      else { p.useFirstPage = f !== 'first-page-unused'; p.firstPage = bump(w.firstPage); }
    }
    if (m.action === 'set_print_centering') p[w.center] = f === 'print-wrong-value' ? !w.value : w.value;
    if (f === 'print-side') { if (m.action === 'set_print_centering') p.oddHeader = 'side'; else p.hCenter = !p.hCenter; }
    ws.print = p;
    if (f === 'side-effect') live.find(x => x !== ws).cells.Z99 = { v: 'side effect' };
    dirty = true;
  }
  // colBreaks only when asked (`colBreaks: true`, the structure family) -- the D2 READ; script.colBreakBuild false ->
  // 'UNOBSERVABLE:runtime-not-verified' on that field
  /** The measured table (P-A3) of `grid` at `at` on `host` -> its address. */
  function writePivotTable(host, at, grid, fl, f) {
    const [ri, ci, di] = [fl.rows, fl.cols, fl.sum].map(x => grid[0].indexOf(x));
    const sums = {};
    for (const row of grid.slice(1)) { const k = `${row[ri]}\u0000${row[ci]}`; sums[k] = (sums[k] || 0) + Number(row[di]); }
    const rk = [...new Set(grid.slice(1).map(r => String(r[ri])))].sort();
    const ck = [...new Set(grid.slice(1).map(r => String(r[ci])))].sort();
    const put = (c, r, v) => { host.cells[addr(at.c1 + c, at.r1 + r)] = { v }; };
    put(0, 0, `Sum of ${fl.sum}`); put(1, 0, 'Column labels');
    put(0, 1, 'Row Labels'); ck.forEach((k, j) => put(1 + j, 1, k)); put(1 + ck.length, 1, 'Grand Total');
    const colTot = ck.map(() => 0);
    rk.forEach((r, i) => {
      put(0, 2 + i, r);
      let tot = 0;
      ck.forEach((c, j) => { const v = sums[`${r}\u0000${c}`]; if (v !== undefined) { put(1 + j, 2 + i, String(f === 'pivot-wrong-sum' && i === 0 && j === 0 ? v + 1 : v)); tot += v; colTot[j] += v; } });
      put(1 + ck.length, 2 + i, String(tot));
    });
    put(0, 2 + rk.length, 'Grand Total'); colTot.forEach((v, j) => put(1 + j, 2 + rk.length, String(v))); put(1 + ck.length, 2 + rk.length, String(colTot.reduce((a, b) => a + b, 0)));
    return { table: `${addr(at.c1, at.r1)}:${addr(at.c1 + ck.length + 1, at.r1 + rk.length + 2)}`, width: ck.length };
  }

  const gridAt = (sheet, range) => { const g = parseRange(range); const out = []; for (let r = g.r1; r <= g.r2; r += 1) { const row = []; for (let c = g.c1; c <= g.c2; c += 1) row.push((sheet.cells[addr(c, r)] || {}).v); out.push(row); } return out; };
  const refOf = (sheet, range) => `${sheet}!$${range.replace(':', ':$').replace(/([A-Z]+)(\d+)/g, '$1$$$2')}`;

  // X-O08 drawings-insert (insert_shape / insert_image): pushes a new raw drawing onto ws.drawings,
  // in the SAME shape drawingReadsOf's `own` mapper already expects (kind/name/pos/width/height,
  // +text for a shape) -- reusing the EXISTING generic drawing machinery, not a parallel system.
  // Faults: 'ignore' (no-op write); 'drawing-insert-wrong-position' (lands one column over);
  // 'drawing-insert-wrong-kind' (an image is recorded as a shape, or vice versa);
  // 'drawing-insert-wrong-stroke-width' (X-O08-7 v2: the SetLine write lands a width +1 EMU off
  // the requested one -- an L1 FAIL on the stroke axis, the SAME shape the position fault uses).
  // A shape's default (no `stroke` requested) strokeWidth is 0 -- the product's own hardcoded
  // `Api.CreateStroke(0, Api.CreateNoFill())`, never undefined/unobserved, mirroring the real driver.
  // GH #139 D4/D6: fillLineColor, the SAME /14+ driver read (drawings.cjs's normalize() now passes
  // it through) -- the written value by default (an accepted runtime, matching today's real /15);
  // script.colourUnobservable simulates a pre-/14 accepted runtime (CAPABILITY_MISSING, the colour
  // axis stays unprovable, same as before this capability existed); 'drawing-insert-wrong-colour' /
  // 'drawing-insert-wrong-stroke-colour' flip the leading hex digit of the written fill/stroke
  // colour -- an L1 FAIL on that ONE axis, the SAME shape the position/stroke-width faults use.
  const flipHex = hex => (hex[0] === 'F' ? `0${hex.slice(1)}` : `F${hex.slice(1)}`);
  function applyDrawingsInsert(m) {
    const f = script.fault;
    if (f === 'ignore') return;
    const ws = live.find(x => x.name === m.sheet);
    if (!ws) return;
    ws.drawings = ws.drawings || [];
    const col = f === 'drawing-insert-wrong-position' ? m.at.col + 1 : m.at.col;
    const kind = f === 'drawing-insert-wrong-kind' ? (m.kind === 'shape' ? 'image' : 'shape') : m.kind;
    const strokeWidth = m.stroke ? Math.round(m.stroke.width * 12700) + (f === 'drawing-insert-wrong-stroke-width' ? 1 : 0) : 0;
    const fillLineColor = kind !== 'shape' ? undefined : script.colourUnobservable ? 'UNOBSERVABLE:CAPABILITY_MISSING'
      : { fill: f === 'drawing-insert-wrong-colour' ? flipHex(m.fill) : m.fill, line: m.stroke ? (f === 'drawing-insert-wrong-stroke-colour' ? flipHex(m.stroke.color) : m.stroke.color) : 'noLine' };
    ws.drawings.push({ kind, name: m.name, pos: { fromCol: col, colOffset: 0, fromRow: m.at.row, rowOffset: 0 }, width: m.sizeEmu.width, height: m.sizeEmu.height, ...(kind === 'shape' ? { text: '', strokeWidth, fillLineColor } : {}) });
  }

  // M7 P1: create_pivot as the probe measured it (P-A3): the table at `at`, (2 + row keys + 1) x (1 + column keys + 1),
  // keys ascending, a title row, a label row, the sums, the grand totals; the pivot record on the host with its public
  // (in-session) source, the raw identity (dataRef, fields), the table address and the layout. Faults: 'pivot-wrong-sum',
  // 'pivot-overwrites-neighbour', 'pivot-not-named', 'pivot-touches-other-object', and 'ignore'.
  function applyPivot(m) {
    const f = script.fault;
    if (f === 'ignore') return;
    if (m.action === 'refresh_pivot') { applyPivotRefresh(m, f); return; }
    const src = live.find(x => x.name === m.sheet);
    const host = live.find(x => x.name === m.pivotSheet);
    if (!src || !host) return;
    const grid = gridAt(src, m.range);
    const at = parseRange(m.at);
    const { table, width } = writePivotTable(host, at, grid, m.fields, f);
    if (f === 'pivot-overwrites-neighbour') host.cells[addr(at.c1 + width + 2, at.r1)] = { v: 'x' };
    host.objects = host.objects || { charts: [], pivots: [], names: [] };
    const ref = refOf(m.sheet, m.range);
    host.objects = { ...host.objects, pivots: [...(host.objects.pivots || []), { name: f === 'pivot-not-named' ? 'PivotTable1' : m.name, source: ref, dataRef: ref, fields: grid[0].map(String), table, layout: { rows: [m.fields.rows], cols: [m.fields.cols], data: [`Sum of ${m.fields.sum}`] } }] };
    if (f === 'pivot-touches-other-object') host.objects = { ...host.objects, names: [...(host.objects.names || []), { name: 'Stray', refersTo: '=A1' }] };
    if (f === 'side-effect') live.find(x => x !== host).cells.Z99 = { v: 'side effect' };
    dirty = true;
  }

  // M7 P3: refresh_pivot as the public SetSource + RefreshTable measured it (m7-p3c-probe4): the pivot rebound to the NAMED
  // range, the old table cleared, the new one at the same anchor, the record an in-session one. Faults: 'refresh-no-rebind'
  // (the setter a silent no-op: the old reference refreshed), 'refresh-leaves-stale' (the old table not cleared),
  // 'pivot-wrong-sum', 'pivot-touches-other-object', 'side-effect'.
  function applyPivotRefresh(m, f) {
    const host = live.find(x => x.name === m.host);
    const pv = ((host && host.objects && host.objects.pivots) || []).find(p => p.name === m.name);
    const ref = pv && (f === 'refresh-no-rebind' ? pv.dataRef || pv.proofSpec.dataRef : refOf(m.sheet, m.range));
    const i = ref ? ref.lastIndexOf('!') : -1;
    const src = ref && live.find(x => x.name === ref.slice(0, i));
    if (!src) return;
    const old = parseRange(pv.table);
    if (f !== 'refresh-leaves-stale') for (let r = old.r1; r <= old.r2; r += 1) for (let c = old.c1; c <= old.c2; c += 1) delete host.cells[addr(c, r)];
    const sum = /^Sum of (.+)$/.exec(pv.layout.data[0])[1];
    const grid = gridAt(src, ref.slice(i + 1).replace(/\$/g, ''));
    const { table } = writePivotTable(host, { c1: old.c1, r1: old.r1 }, grid, { rows: pv.layout.rows[0], cols: pv.layout.cols[0], sum }, f);
    const fields = pv.fields || (pv.proofSpec || {}).fields;
    host.objects.pivots = host.objects.pivots.map(x => (x === pv ? { name: pv.name, source: ref, dataRef: ref, fields, table, layout: pv.layout } : x));
    if (f === 'pivot-touches-other-object') host.objects = { ...host.objects, names: [...(host.objects.names || []), { name: 'Stray', refersTo: '=A1' }] };
    if (f === 'side-effect') live.find(x => x !== host).cells.Z99 = { v: 'side effect' };
    dirty = true;
  }
  const printReadsOf = (asked, cols) => Object.fromEntries((asked.includes('*') ? live.map(x => x.name) : asked).map(n => { const ws = live.find(x => x.name === n);
    if (script.printBuild === false) return [n, 'UNOBSERVABLE:runtime-not-verified'];
    // M7 F2: the widened fields (oddFooter, vCenter, and with them colBreaks) as the proven build serves them;
    // script.printWide false -> an unwidened read
    const { colBreaks, oddFooter, vCenter, firstHeader, firstFooter, evenHeader, evenFooter, diffFirst, diffOddEven, ...pr } = { ...PRINT_DEFAULT, ...clone((ws && ws.print) || {}) };
    const wide = script.printWide !== false;
    return [n, { ...pr, ...(cols || wide ? { colBreaks: script.colBreakBuild === false ? 'UNOBSERVABLE:runtime-not-verified' : colBreaks } : {}), ...(wide ? { oddFooter, vCenter, firstHeader, firstFooter, evenHeader, evenFooter, diffFirst, diffOddEven } : {}) }]; }));
  const PAGE_DEFAULT = { orientation: 'xlPortrait', printGridlines: false, printHeadings: false, topMargin: 19.1, bottomMargin: 19.1, leftMargin: 17.8, rightMargin: 17.8 };
  function applyPage(m) {
    const f = script.fault;
    if (f === 'ignore') return;
    const ws = live.find(x => x.name === m.sheet);
    if (!ws) return;
    const keys = Object.keys(m.want);
    const pg = { ...PAGE_DEFAULT, ...(ws.page || {}) };
    for (const k of f === 'page-partial' ? keys.slice(0, 1) : keys) pg[k] = m.want[k];
    if (f === 'page-side') for (const k of Object.keys(PAGE_DEFAULT).filter(x => !keys.includes(x))) pg[k] = k === 'orientation' ? (pg[k] === 'xlPortrait' ? 'xlLandscape' : 'xlPortrait') : typeof pg[k] === 'number' ? pg[k] + 1 : !pg[k];
    if (f === 'page-margin-drift') for (const k of keys.filter(x => /Margin$/.test(x))) pg[k] += 0.5;
    ws.page = pg;
    if (f === 'page-other-sheet') { const o = live.find(x => x !== ws); o.page = { ...PAGE_DEFAULT, ...(o.page || {}), orientation: 'xlLandscape' }; }
    if (f === 'side-effect') live.find(x => x !== ws).cells.Z99 = { v: 'side effect' };
    dirty = true;
  }
  const pageReadsOf = asked => Object.fromEntries((asked.includes('*') ? live.map(x => x.name) : asked).map(n => { const ws = live.find(x => x.name === n);
    const st = { ...PAGE_DEFAULT, ...((ws && ws.page) || {}) };
    if (script.pageFieldMissing) delete st.printHeadings;
    if (script.pageMarginMissing) delete st.leftMargin;
    return [n, script.pageBuild === false ? 'UNOBSERVABLE:per-sheet-read-not-verified' : st]; }));
  // X-O04/10/11: MARKER_WRITE_TYPES mirrors the real driver's own scope (xlsx-driver.cjs chartReads:
  // type==='scatter'||'lineNormal') -- a series only carries `markerStyle` on those two types,
  // matching the X-O04/6 LIVE fixture (an untouched series' field is null, never absent/guessed).
  // script.chartStyleUnobservable simulates a pre-/11 bridge (both fields read 'UNOBSERVABLE:...').
  const MARKER_WRITE_TYPES = ['scatter', 'lineNormal'];
  const chartReadsOf = asked => Object.fromEntries((asked.includes('*') ? live.map(x => x.name) : asked).map(n => { const ws = live.find(x => x.name === n); const o = ws && ws.objects ? ws.objects : { charts: [], pivots: [] };
    return [n, script.chartsUnobservable ? 'UNOBSERVABLE:threw' : { charts: (o.charts || []).map(c => ({ name: c.name === undefined ? null : c.name,
      // X-O03/3 (LIVE MEASURED): GetChartType() reads a radar chart back as "unknown", never
      // "radar" -- every other type (stock included) reads back as its own name.
      type: c.type === 'radar' ? 'unknown' : c.type, title: c.title === undefined ? null : c.title, legend: c.legend === undefined ? 'bottom' : c.legend,
      series: c.series.map((x, si) => ({ name: x.name, values: x.values, cat: x.cat,
        // X-O05/2 (LIVE VERIFIED): order is read on EVERY chart type, not scoped like markerStyle
        // above -- an untouched series' baseline is its own index (the measured default, sequential).
        order: script.chartStyleUnobservable ? 'UNOBSERVABLE:CAPABILITY_MISSING' : (script.orderReadAsObject ? { order: x.order === undefined ? si : x.order } : (x.order === undefined ? si : x.order)),
        ...(MARKER_WRITE_TYPES.includes(c.type) ? { markerStyle: script.chartStyleUnobservable ? 'UNOBSERVABLE:CAPABILITY_MISSING' : { fill: x.markerFill === undefined ? null : x.markerFill, outline: x.markerOutline === undefined ? null : x.markerOutline } } : {}) })),
      // X-O05: an untouched chart's axis titles / data labels are absent in the fake's own object -- normalized to the
      // baseline the real engine MEASURED (null / null / every flag false); script.chartLabelsUnobservable
      // makes the labels read return a non-object (the unobservable branch)
      horAxisTitle: c.horAxisTitle === undefined ? null : c.horAxisTitle, verAxisTitle: c.verAxisTitle === undefined ? null : c.verAxisTitle,
      dataLabels: script.chartLabelsUnobservable ? undefined : (c.dataLabels || { showSerName: false, showCatName: false, showVal: false, showPercent: false }),
      // GH #302 G6: the per-point effective labels only when the chart record carries them (absent = not read, the
      // unobservable branch -- the fake's charts predate this field)
      ...(c.effectiveLabels !== undefined ? { effectiveLabels: clone(c.effectiveLabels) } : {}),
      gridlinesColor: script.chartStyleUnobservable ? 'UNOBSERVABLE:CAPABILITY_MISSING' : { verticalGridlines: (c.gridlinesColor && c.gridlinesColor.verticalGridlines !== undefined) ? c.gridlinesColor.verticalGridlines : null, horizontalGridlines: (c.gridlinesColor && c.gridlinesColor.horizontalGridlines !== undefined) ? c.gridlinesColor.horizontalGridlines : null },
      // X-O05/2 (LIVE VERIFIED): chart-level, read on EVERY chart type the fake carries (the real
      // bridge handler itself gates to bar-only; classifyStyle's own bar-only check runs BEFORE
      // this field is ever read, so the fake does not need to repeat that gate here).
      gapOverlap: script.chartStyleUnobservable ? 'UNOBSERVABLE:CAPABILITY_MISSING' : { gapWidth: (c.gapOverlap && c.gapOverlap.gapWidth !== undefined) ? c.gapOverlap.gapWidth : null, overlap: (c.gapOverlap && c.gapOverlap.overlap !== undefined) ? c.gapOverlap.overlap : null } })), pivots: (o.pivots || []).map(p => p.name) }]; }));
  // Drawings (M7 S2b): a sheet's charts (objects.charts; pos / w / h default to the probe's chart geometry, w: 0 = the
  // measured 0 x 0) plus its explicit `drawings` (raw images / shapes in the driver's read shape); read for
  // `drawingSheets` ('*' = every sheet). script.drawingsUnobservable -> every sheet 'UNOBSERVABLE:threw';
  // script.drawingField: [field, value] overwrites one field of every image / shape (a missing / bad getter).
  const CHART_GEOM = { pos: { fromCol: 2, colOffset: 0, fromRow: 12, rowOffset: 0 }, w: 3600000, h: 2160000 };
  let anchorAsked = false;
  const anchorOf = d => (d.pos && typeof d.pos === 'object' && Number.isFinite(d.width) && Number.isFinite(d.height)
    ? { type: 1, to: { col: d.pos.fromCol + Math.floor(d.width / 609600), colOff: d.width % 609600, row: d.pos.fromRow + Math.floor(d.height / 181000), rowOff: d.height % 181000 } }
    : 'UNOBSERVABLE:no-boundsFromTo-to');
  const drawingReadsOf = asked => Object.fromEntries((asked.includes('*') ? live.map(x => x.name) : asked).map(n => {
    const ws = live.find(x => x.name === n);
    if (script.drawingsUnobservable) return [n, 'UNOBSERVABLE:threw'];
    // an unnamed chart: EO names every drawing ('Chart <n>'); an explicit null (fault 'chart-extra-unnamed') reads 'null'
    const charts = ((ws && ws.objects && ws.objects.charts) || []).map((c, i) => ({ kind: 'chart', name: c.name === undefined ? `Chart ${i + 1}` : (c.name === null ? 'null' : c.name), parent: n,
      pos: clone(c.pos || CHART_GEOM.pos), width: c.w === undefined ? CHART_GEOM.w : c.w, height: c.h === undefined ? CHART_GEOM.h : c.h, rotation: 0, flipH: null, flipV: null, text: 'UNOBSERVABLE:no-public-content-getter' }));
    const own = ((ws && ws.drawings) || []).map(d => ({ parent: n, rotation: 0, flipH: d.kind === 'shape' ? false : null, flipV: d.kind === 'shape' ? false : null, ...clone(d) }));
    if (script.drawingField) for (const d of own) d[script.drawingField[0]] = script.drawingField[1];
    const all = [...own, ...charts];
    // M7 D-B: the internal anchor read, only when asked on a proven build (script.anchorBuild === false = not proven);
    // an explicit `anchor` on a fixture drawing wins, else the probe's cell metrics (col 609600 / row 181000 EMU)
    if (anchorAsked) for (const d of all) if (d.anchor === undefined) d.anchor = anchorOf(d);
    // M7 D-B hydration (B15 /5: sheet.model.ensureReady/v1 through the bridge, before the anchor read): script.hydration 'missing' -> the sheet's whole
    // drawing read is the reason; 'not-hydrated' -> every anchor of the sheet reads the init did not reach it
    if (anchorAsked && script.hydration === 'missing') return [n, 'UNOBSERVABLE:hydration-BRIDGE:view-creator-missing'];
    if (anchorAsked && script.hydration === 'not-hydrated') for (const d of all) d.anchor = 'UNOBSERVABLE:geometry-not-hydrated';
    return [n, all];
  }));
  /** L3 negative controls on the drawings, after the product's dispatch: faults 'drawing-moves' / 'drawing-renames' /
   *  'drawing-resizes' (+1 EMU) / 'drawing-text' / 'drawing-swaps' (same count: one replaced by another) / 'chart-moves' /
   *  'chart-renames-other' / 'chart-series-other' / 'chart-size-lost' (a chart reads 0 x 0 after the write), on script.drawingSheet */
  const DRAWING_FAULTS = {
    'drawing-moves': (d) => { d.pos = { ...d.pos, fromRow: d.pos.fromRow + 1 }; },
    'drawing-renames': (d) => { d.name = `${d.name}_X`; },
    'drawing-resizes': (d) => { d.width += 1; },
    'drawing-text': (d) => { d.text = `${d.text}~`; },
    'drawing-size-lost': (d) => { d.width = 0; d.height = 0; },
    'drawing-swaps': (d, c, ws) => { ws.drawings[0] = { ...clone(d), kind: d.kind === 'shape' ? 'image' : 'shape', ...(d.kind === 'shape' ? {} : { text: 'swapped' }) }; },
    // M7 D-B: only the anchor read changes (pos / size untouched) -- the anchor L3 alone must catch it
    'drawing-anchor': (d) => { d.anchor = { type: 1, to: { col: 30, colOff: 0, row: 30, rowOff: 0 } }; },
    'chart-moves': (d, c) => { const p = c.pos || CHART_GEOM.pos; c.pos = { ...p, fromCol: p.fromCol + 1 }; },
    'chart-renames-other': (d, c) => { c.name = `${c.name}_X`; },
    'chart-series-other': (d, c) => { c.series = c.series.slice(0, -1); },
    'chart-size-lost': (d, c) => { c.w = 0; c.h = 0; },
  };
  function drawingFault() {
    const ws = live.find(x => x.name === script.drawingSheet);
    const fn = DRAWING_FAULTS[script.fault];
    if (!ws || !fn) return;
    const d = (ws.drawings || [])[0];
    const c = ((ws.objects && ws.objects.charts) || [])[0];
    if (script.fault.startsWith('chart-') ? c : d) fn(d, c, ws);
  }
  const filterReadsOf = asked => Object.fromEntries((asked.includes('*') ? live.map(x => x.name) : asked).map(n => { const ws = live.find(x => x.name === n); return [n, script.filterBuild === false ? 'UNOBSERVABLE:field-build-not-verified' : (ws && ws.autoFilter ? clone(ws.autoFilter) : { range: null, filters: [] })]; }));
  // GH #326/#327: a sheet is unprotected unless the fixture opts in with `ws.protection = { forbids: {<flag
  // overrides, default false>}, unlocked: [<A1 addresses, EXPLICITLY unlocked>] }` -- the SAME shape
  // sheet.protection.read/v1 reads on the real DS 9.3.5 bridge (docs/DS-9.3.5-BRIDGE17-PROTECTION-READ.md): a
  // NON-EMPTY cell NOT named in `unlocked` is the Excel default, LOCKED (an empty cell is never listed either
  // way -- the SAME fail-closed direction the real handler's own range walk uses). An unprotected sheet needs
  // no cell-level walk (`cells: null`). `script.protectionUnread` names a sheet whose read degrades, for the
  // fail-closed branch's own test.
  const FLAGS = ['objects', 'scenarios', 'formatCells', 'formatColumns', 'formatRows', 'insertColumns', 'insertRows', 'insertHyperlinks', 'deleteColumns', 'deleteRows', 'selectLockedCells', 'sort', 'autoFilter', 'pivotTables', 'selectUnlockedCells'];
  const protectionValue = (n, range) => {
    const ws = live.find(x => x.name === n);
    if (!ws) return 'UNOBSERVABLE:sheet-missing';
    if ((script.protectionUnread || []).includes(n)) return 'UNOBSERVABLE:scripted';
    const p = ws.protection;
    if (!p) return { protected: false, forbids: null, cells: range !== undefined ? rangeCells(ws, range, null) : null };
    const forbids = Object.fromEntries(FLAGS.map(f => [f, !!(p.forbids && p.forbids[f])]));
    return { protected: true, forbids, cells: range !== undefined ? rangeCells(ws, range, p.unlocked || []) : null };
  };
  // `unlocked === null` (the sheet is NOT protected): every non-empty cell reads as unlocked (no restriction
  // applies either way); otherwise ONLY the named addresses do -- the Excel locked-by-default direction.
  const rangeCells = (ws, range, unlocked) => {
    const g = parseRange(range);
    const nonEmpty = cellsOf(g).filter(a => ws.cells && ws.cells[a] !== undefined);
    return { range, nonEmpty: nonEmpty.length, unlocked: unlocked === null ? nonEmpty : nonEmpty.filter(a => unlocked.includes(a)), hidden: [] };
  };
  const protectionReadsOf = asked => Object.fromEntries((asked.includes('*') ? live.map(x => x.name) : asked).map(n => [n, protectionValue(n)]));
  const protectionRangeReadsOf = asked => Object.fromEntries(asked.map(w => [`${w.sheet}!${w.range}`, protectionValue(w.sheet, w.range)]));

  const sideNames = [];
  /** faults 'side-rule' / 'side-name': the dispatch also adds a rule on another sheet / a workbook name (L3's job) */
  function sideWorkbook(target) {
    if (script.fault === 'side-rule') { const ws = live.find(x => x.name === 'Other') || live.find(x => x.name !== target); ws.rules = [...(ws.rules || []), { kind: 'conditional', range: 'Z1', formula: '=A1>0' }]; }
    if (script.fault === 'side-name') sideNames.push({ name: 'Sneaky', ref: `${live[0].name}!$A$1` });
  }

  // '*' = every sheet, as the real driver reads it (M7 P1 asks for every sheet's pivots without the dependents read)
  const objectSheets = request => (request && (request.dependents || ((request.objectSheets || []).includes('*'))) ? live.map(x => x.name) : ((request && request.objectSheets) || []).filter(n => live.some(x => x.name === n)));

  /**
   * The dependency readers in the REAL driver's raw shapes (W3d/W3e): internal names [{ name, ref,
   * scopeSheet }] (workbook names: the store's `names`, else script.workbookNames; sheet-level ones from each sheet's
   * objects.names, ref without its '='), internal validations { sheet: [{ type, ranges, formula1,
   * formula2 }] } and public conditional formats { sheet: [{ type, appliesTo, formula1, formula2 }] }
   * from each sheet's `rules`. Faults: script.dependencyBuild false ('UNOBSERVABLE:PATCH_VERSION_MISMATCH', the bridge's own identity refusal),
   * script.namesFault / validationsFault / conditionalFault (that axis a reason string).
   */
  function dependents() {
    const bare = t => String(t || '').replace(/^=/, '');
    const rect = a => { const g = parseRange(a); return { r1: g.r1 - 1, c1: g.c1 - 1, r2: g.r2 - 1, c2: g.c2 - 1 }; };
    const names = [...liveNames.concat(sideNames).map(n => ({ name: n.name, ref: n.ref, scopeSheet: null })),
      ...live.flatMap(x => ((x.objects && x.objects.names) || []).map(n => ({ name: n.name, ref: bare(n.refersTo), scopeSheet: x.name })))];
    // L-22d v2: the bridge projects the rule's OWN raw serial (r.bridgeFormula/2) when the rule carries one
    // (date) -- never the public locale text (r.formula/2), which is what every OTHER type still projects.
    const validations = Object.fromEntries(live.map(x => [x.name, (x.rules || []).filter(r => r.kind === 'validation').map(r => ({ type: 4, ranges: [rect(r.range)], formula1: bare('bridgeFormula' in r ? r.bridgeFormula : r.formula), formula2: ('bridgeFormula2' in r ? r.bridgeFormula2 : r.formula2) ? bare('bridgeFormula2' in r ? r.bridgeFormula2 : r.formula2) : null }))]));
    const conditional = Object.fromEntries(live.map(x => [x.name, (x.rules || []).filter(r => r.kind === 'conditional').map(r => ({ type: r.type || 'xlExpression', ...(r.type === 'xlCellValue' ? { operator: r.operator } : {}), appliesTo: `${x.name}!$${r.range.replace(/(\d)/, '$$$1')}`, formula1: bare(r.formula), formula2: r.formula2 ? bare(r.formula2) : '' }))]));
    if (script.conditionalFault) conditional[live[0].name] = script.conditionalFault;
    if (script.dependencyBuild === false) return { dependents: 'UNOBSERVABLE:PATCH_VERSION_MISMATCH', conditional };
    return { dependents: { names: script.namesFault || names, validations: script.validationsFault || validations }, conditional };
  }

  return {
    calls,
    store,
    reads,
    liveState: () => clone(live || []),
    // M8 cross-format: the engine's open / close failure branches through the adapter (fake-session shape)
    async open(args) { calls.open += 1; if (script.openFails) return { ok: false, detail: 'fake xlsx open refused' }; openedFileId = args && typeof args.fileId === 'string' ? args.fileId : null; live = clone(store.durable); liveNames = clone(store.names || script.workbookNames || []); active = store.active || live[0].name; dirty = false; return { ok: true, id: `fake-xlsx-${calls.open}` }; },
    // M7 save contract: like the real XLSX driver, forceSave returns its own paired confirmation, and close
    // re-evaluates it -- a dispatch after the confirmation supersedes it. script.forceSaveUnconfirmed models a refused one.
    async close() {
      calls.close += 1; live = null;
      const confirmation = confirmedAt === null ? null
        : (calls.dispatch > confirmedAt ? { state: 'unconfirmed', reason: 'superseded-by-a-newer-save' } : { state: 'confirmed', reason: 'forcesave-confirmed' });
      if (script.closeFails) return { ok: false, detail: 'fake xlsx close refused', confirmation: script.closeConfirmation || confirmation };
      return { ok: true, confirmation: script.closeConfirmation || confirmation };
    },
    async forceSave() {
      calls.forceSave += 1; store.durable = clone(live); for (const [k, lit] of Object.entries(literalOnSave)) { const [sh, a] = k.split('\u0000'); const ws = store.durable.find(x => x.name === sh); if (ws && ws.cells[a] && ws.cells[a].f === lit.shown) ws.cells[a] = { ...ws.cells[a], f: lit.f, v: lit.v }; } store.names = clone(liveNames); store.active = active; store.saves += 1; dirty = false;
      if (script.forceSaveUnconfirmed) return { ok: true, confirmed: false, forceSave: { state: 'unconfirmed', reason: script.forceSaveUnconfirmed } };
      confirmedAt = calls.dispatch;
      return { ok: true, confirmed: true, forceSave: { state: 'confirmed', reason: 'forcesave-confirmed' } };
    },
    async canSave() { calls.canSave += 1; return { ok: true, canSave: dirty }; },
    async readLive(_handle, request = {}) {
      calls.readLive += 1;
      reads.push(clone(request || {}));
      const out = [];
            const idx = new Map();
      const push = (sheet, a, fmt, sty) => {
        const k = `${sheet.name}!${a}`;
        if (idx.has(k)) { if (fmt || sty) Object.assign(out[idx.get(k)], readCell(sheet, a, fmt || sty, sty)); return; }
        idx.set(k, out.length); out.push(readCell(sheet, a, fmt || sty, sty));
      };
      const styled = n => ((request && request.styleSheets) || []).includes(n);
      for (const sheet of live) {
        const used = usedOf(sheet);
        if (used) for (const a of cellsOf(parseRange(used))) push(sheet, a, ((request && request.formatSheets) || []).includes(sheet.name), styled(sheet.name));
      }
      for (const w of (request && request.windows) || []) {
        const sheet = live.find(s => s.name === w.sheet);
        const g = parseRange(w.range);
        if (sheet && g) for (const a of cellsOf(g)) {
          push(sheet, a, w.format === true, w.style === true);
          // X-T01: the `nf` window flag reads the number format ALONE (the driver's PUBLIC GetNumberFormat)
          if (w.nf === true && w.format !== true) { const o = out[idx.get(`${sheet.name}!${a}`)]; if (!('nf' in o)) o.nf = readCell(sheet, a, true, false).nf; }
        }
      }
      return {
        ok: true,
        authority: 'LIVE_READ',
        payload: { sheets: live.map((s, index) => ({ name: s.name, index, used: usedOf(s) })), cells: out, truncated: script.readTruncated === true, unsupported: [],
          objects: Object.fromEntries(objectSheets(request).map(n => [n, script.objectsUnobservable ? { charts: [], pivots: [], names: 'UNOBSERVABLE:threw' } : withPivotProofs(live, script, JSON.parse(JSON.stringify(live.find(x => x.name === n).objects || { charts: [], pivots: [], names: [] })))])),
          ...(request && request.dependents ? dependents() : {}),
          ...(request && request.layout ? { layout: layoutOf(request.layout) } : {}),
          ...(request && request.panes ? { panes: panesOf(), active } : {}),
          ...(request && request.sheetVisibility ? { visibility: Object.fromEntries(live.map(x => [x.name, script.visibilityUnreadable ? 'UNOBSERVABLE:scripted' : x.hidden !== true])) } : {}),
          ...(request && request.validationRanges ? { validationReads: validationReadsOf(request.validationRanges) } : {}),
          ...(request && request.cfRanges ? { cfReads: cfReadsOf(request.cfRanges) } : {}),
          ...(request && request.runCells ? { runReads: runReadsOf(request.runCells) } : {}),
          ...(request && request.arrayCells ? { arrayReads: arrayReadsOf(request.arrayCells) } : {}),
          ...(request && request.commentSheets ? { commentReads: commentReadsOf(request.commentSheets) } : {}),
          ...(request && request.sessionUser ? { sessionUser: sessionUserOf() || 'UNOBSERVABLE:no-session-user' } : {}),
          ...(request && request.filterSheets ? { filterReads: filterReadsOf(request.filterSheets) } : {}),
          ...(request && request.chartSheets ? { chartReads: chartReadsOf(request.chartSheets) } : {}),
          ...(request && request.drawingSheets ? { drawingReads: (anchorAsked = !!request.drawingAnchor && script.anchorBuild !== false, drawingReadsOf(request.drawingSheets)) } : {}),
          ...(request && request.pageSheets ? { pageReads: pageReadsOf(request.pageSheets) } : {}),
          ...(request && request.printSheets ? { printReads: printReadsOf(request.printSheets, request.colBreaks === true) } : {}),
          ...(request && request.protectionSheets ? { protectionReads: protectionReadsOf(request.protectionSheets) } : {}),
          ...(request && request.protectionRanges ? { protectionRangeReads: protectionRangeReadsOf(request.protectionRanges) } : {}),
          runtime: { identity: script.runtime === undefined ? XLSX_RUNTIME_ID : script.runtime, cacheKey: '9.3.4-fakecachekey00', ...(script.fileIdUnobservable ? {} : { fileId: openedFileId }) } },
      };
    },
    async readPackage() { calls.readPackage += 1; return { ok: true, parts: {} }; },
    async dispatch(_handle, command) {
      calls.dispatch += 1;
      if (script.fault === 'throw') throw new Error('transport lost');
      if (!command || command.kind !== 'callCommand' || !command.meta) return { ok: false, dispatched: true };
      // GH #300 K2: the real driver's author prelude -- no session author, no dispatch
      if (command.meta.sessionAuthor === true && !sessionUserOf()) return { ok: false, dispatched: false, outcome: 'comment-author-identity-unavailable' };
      // M11.5c.4: copy_sheet is ONE worksheet.copy/v1 Bridge WRITE (meta.bridgeCopyWrite); the receipt rule of the driver's B03 client
      if (command.meta.family === 'sheets' && command.meta.bridgeCopyWrite === true) {
        const failed = evidence => ({ ok: false, dispatched: true, outcome: 'b03-bridge-write-failed', evidence, error: null });
        if (script.b03Write === 'gate-refused') return failed({ ok: false, stage: 'bridge-write-gate', reason: script.b03GateReason || 'CAPABILITY_MISSING', mutation: 'NONE' });
        if (script.fault === 'copy-timeout' || script.fault === 'ignore') { calls.bridgeWrites += 1; return failed({ ok: false, stage: 'bridge-write', reason: 'BRIDGE:copy-not-observed', uncertain: true, mutation: 'UNKNOWN' }); }
        applyCopy(command.meta); calls.bridgeWrites += 1;
        sideWorkbook(command.meta.name);
        drawingFault();
        const hits = live.filter(x => x.name === command.meta.name).length;
        if (script.b03Write === 'error-applied') return failed({ ok: false, stage: 'bridge-write', reason: 'BRIDGE:copy-error', uncertain: true, mutation: 'UNKNOWN' });
        return hits === 1 ? { ok: true, dispatched: true, outcome: 'command-completed', error: null } : failed({ ok: false, stage: 'bridge-write', reason: 'BRIDGE:copy-unexpected', uncertain: true, mutation: 'UNKNOWN' });
      }
      if (command.meta.action === 'write') applyWrite(command.meta);
      if (command.meta.action === 'clear-content') applyClear(command.meta);
      if (command.meta.action === 'clear-formats' || command.meta.action === 'clear-all') applyClearScope(command.meta);
      if (command.meta.family === 'sheets') { applySheet(command.meta); sideWorkbook(command.meta.sheet); }
      if (command.meta.family === 'merge') applyMerge(command.meta);
      if (command.meta.family === 'layout') applyLayout(command.meta);
      if (command.meta.family === 'format' && command.meta.bridgeAlignExtWrite === true) return b05Write(command.meta);
      if (command.meta.family === 'format') applyFormat(command.meta);
      if (command.meta.family === 'freeze') applyFreeze(command.meta);
      if (command.meta.family === 'structure') applyStructure(command.meta);
      if (command.meta.family === 'range') applyRange(command.meta);
      if (command.meta.family === 'names') applyNames(command.meta);
      if (command.meta.family === 'sort') applySort(command.meta);
      if (command.meta.family === 'validation') applyValidation(command.meta);
      if (command.meta.family === 'cf') applyCf(command.meta);
      if (command.meta.family === 'format_run') applyFormatRun(command.meta);
      if (command.meta.family === 'filter') applyFilter(command.meta);
      if (command.meta.family === 'array-formula') applyArrayFormula(command.meta);
      if (command.meta.family === 'comments') applyComment(command.meta);
      // X-O04/1+2 (1bafc2ca): set_chart_style has NO getter -- the real engine's own setter
      // return value (checked inside lower()'s callCommand body, fail-closed) is the ONLY
      // failure signal; the fake has no JS body to execute, so a rejection is simulated at THIS
      // dispatch layer instead (the generic `f === 'ignore'` no-op-but-ok=true path cannot stand
      // in for it, since nothing downstream would ever notice a style write that silently no-oped).
      if (command.meta.family === 'charts' && command.meta.action === 'set_chart_style' && script.fault === 'style-setter-rejected') {
        return { ok: false, dispatched: true, outcome: 'style-setter-rejected', evidence: { stage: 'style-fill-rejected' } };
      }
      if (command.meta.family === 'charts') applyCharts(command.meta);
      if (command.meta.family === 'drawings-insert') applyDrawingsInsert(command.meta);
      if (command.meta.family === 'page') applyPage(command.meta);
      if (command.meta.family === 'print' && command.meta.bridgePrintWrite === true) return b10Write(command.meta);
      if (command.meta.family === 'print') applyPrint(command.meta);
      if (command.meta.family === 'pivots') applyPivot(command.meta);
      drawingFault();
      return { ok: true, dispatched: true };
    },
  };
}

module.exports = { createStore, createFakeXlsxDriver };
