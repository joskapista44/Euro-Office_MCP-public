'use strict';

// L-10: the read_freeze_panes / read_page_setup / read_filter projections (src/adapters/xlsx/families/{panes,page,filter}-read.cjs). The ENGINE side is simulated in the shape the driver's own
// reads return (xlsx-driver.cjs: panes via sheet.freezeState.read/v1, pageReads via the public page getters, printReads via sheet.printSettings.read/v1, filterReads via the public AutoFilter
// getters + the field identity) from each WRITER's own lowered operation, so a round trip request -> engine entry -> projection must reproduce the request. The LIVE acceptance of the same
// is test/live/release-l10-sheet-settings-product.live.cjs (through the product, EO_BIN).

const test = require('node:test');
const assert = require('node:assert/strict');

const freeze = require('../../src/adapters/xlsx/families/freeze.cjs');
const page = require('../../src/adapters/xlsx/families/page.cjs');
const print = require('../../src/adapters/xlsx/families/print.cjs');
const filter = require('../../src/adapters/xlsx/families/filter.cjs');
const { projectPanesRead } = require('../../src/adapters/xlsx/families/panes-read.cjs');
const { projectPageSetupRead, PAPER_OF } = require('../../src/adapters/xlsx/families/page-read.cjs');
const { projectFilterRead } = require('../../src/adapters/xlsx/families/filter-read.cjs');
const rt = require('../../src/server/read-tools.cjs');

const SHEETS = [{ name: 'S', index: 0, used: 'A1:C5' }];
const A = { sheet: 'S' };

// ---------------------------------------------------------------------------------------------------- freeze panes
const panes = p => projectPanesRead({ sheets: SHEETS, panes: p === undefined ? null : { S: p } }, A);

test('PANES ROUND TRIP: freeze_panes rows / columns as the writer lowered them read back as the same numbers; unfreeze reads frozen:false', () => {
  for (const spec of [{ rows: 2 }, { columns: 3 }, { rows: 1, columns: 2 }]) {
    const e = freeze.plan({ intent: 'freeze_panes', sheet: 'S', ...spec }, 0).operation.expect;
    const out = panes(e);
    assert.equal(out.frozen, true);
    assert.deepEqual({ rows: out.rows, columns: out.columns }, { rows: spec.rows, columns: spec.columns });
  }
  const none = freeze.plan({ intent: 'unfreeze_panes', sheet: 'S' }, 0).operation.expect;
  assert.deepEqual(panes(none), { sheet: 'S', frozen: false });
});

test('PANES: a gap is NAMED, never "not frozen": not read, the bridge reason, a split / unknown pane, an impossible frozen 0 / 0', () => {
  assert.deepEqual(panes(undefined), { sheet: 'S', frozen: null, reason: 'UNOBSERVABLE:not-read' });
  assert.equal(panes('UNOBSERVABLE:BRIDGE:worksheet-missing').reason, 'UNOBSERVABLE:BRIDGE:worksheet-missing');
  assert.equal(panes({ state: 'split' }).reason, 'UNOBSERVABLE:split-or-unknown-pane');
  assert.equal(panes({ state: 'frozen', row0: 0, col0: 0 }).reason, 'UNOBSERVABLE:split-or-unknown-pane');
  assert.equal(panes({ state: 'frozen', row0: -1, col0: 2 }).reason, 'UNOBSERVABLE:split-or-unknown-pane');
  assert.equal(panes(7).reason, 'UNOBSERVABLE:unexpected-shape');
  assert.deepEqual(projectPanesRead({ sheets: SHEETS, panes: {} }, A), { sheet: 'S', frozen: null, reason: 'UNOBSERVABLE:not-read' });
});

test('PANES: a sheet that does not exist is SHEET_NOT_FOUND; a pane beyond the writable bounds keeps its numbers and is named', () => {
  assert.deepEqual(projectPanesRead({ sheets: SHEETS, panes: {} }, { sheet: 'Nope' }), { sheet: 'Nope', error: 'SHEET_NOT_FOUND' });
  // X-L09 (owner GO 36696): the writable bound is 100000 rows / 1000 columns now (was 1000 / 100) -- 1500 x 120 is representable, 100001 x 1001 is not
  const ok = panes({ state: 'frozen', row0: 1500, col0: 120 });
  assert.deepEqual([ok.rows, ok.columns, ok.unrepresentable], [1500, 120, undefined]);
  const big = panes({ state: 'frozen', row0: 100001, col0: 1001 });
  assert.deepEqual([big.rows, big.columns], [100001, 1001]);
  assert.deepEqual(big.unrepresentable.map(u => u.field), ['rows', 'columns']);
});

// ---------------------------------------------------------------------------------------------------- filter
const flt = (r, name = 'S') => projectFilterRead({ sheets: SHEETS, filterReads: r === undefined ? null : { [name]: r } }, A);
const eng = (field, operator, criteria1, criteria2 = null, on = true) => ({ field, on, operator, criteria1, criteria2 });

test('FILTER ROUND TRIP: filter_range value / values / and read back as the same request fields (the read mirrors the writer)', () => {
  const one = filter.plan({ intent: 'filter_range', sheet: 'S', range: 'A1:C5', field: 2, value: 3 }, 0).operation;
  assert.equal(one.value, '3');
  assert.deepEqual(flt({ range: '$A$1:$C$5', filters: [eng(2, 'xlOr', '=3')] }).filters, [{ field: 2, value: '3' }]);
  assert.deepEqual(flt({ range: 'A1:C5', filters: [eng(2, 'xlAnd', '3')] }).filters, [{ field: 2, value: '3' }], 'a bare criteria1 reads the same (filter.cjs bare())');
  const two = filter.plan({ intent: 'filter_range', sheet: 'S', range: 'A1:C5', field: 1, values: ['a', 'b'] }, 0).operation;
  assert.deepEqual([...two.values], ['a', 'b']);
  assert.deepEqual(flt({ range: 'A1:C5', filters: [eng(1, 'xlOr', '=a', '=b')] }).filters, [{ field: 1, values: ['a', 'b'] }]);
  const and = filter.plan({ intent: 'filter_range', sheet: 'S', range: 'A1:C5', field: 3, and: ['>1', '<5'] }, 0).operation;
  assert.deepEqual([...and.and], ['>1', '<5']);
  assert.deepEqual(flt({ range: 'A1:C5', filters: [eng(3, 'xlAnd', '>1', '<5')] }).filters, [{ field: 3, and: ['>1', '<5'] }]);
  assert.equal(flt({ range: '$A$1:$C$5', filters: [] }).range, 'A1:C5', 'the $ is stripped');
});

test('FILTER (L-10b r1): a value LIST is read by its RAW SHAPE -- 1 element -> value, 2 different plain -> values, 3+ or non-plain -> unsupported; NEVER joined into one value', () => {
  const one = flt({ range: 'A1:C5', filters: [eng(1, 'xlFilterValues', ['3'])] }).filters[0];
  assert.deepEqual(one, { field: 1, value: '3' });
  const two = flt({ range: 'A1:C5', filters: [eng(1, 'xlFilterValues', ['2', '4'])] }).filters[0];
  assert.deepEqual(two, { field: 1, values: ['2', '4'] });
  assert.deepEqual(flt({ range: 'A1:C5', filters: [eng(1, 'xlFilterValues', [2, 4])] }).filters[0], { field: 1, values: ['2', '4'] }, 'numbers are the writer\'s String() of the value');
  // GH #301 D2: 3..10 different plain values under xlFilterValues ARE a writer shape now -> values; 11, or 3 under another operator, stay a gap
  assert.deepEqual(flt({ range: 'A1:C5', filters: [eng(1, 'xlFilterValues', ['2', '3', '4'])] }).filters[0], { field: 1, values: ['2', '3', '4'] });
  assert.equal(flt({ range: 'A1:C5', filters: [eng(1, 'xlOr', ['2', '3', '4'])] }).filters[0].unsupported, true, 'a 3-list only as xlFilterValues');
  for (const list of [['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11'], ['2', '3', '2'], ['2', '2'], ['2', '>4'], ['=2', '4'], [''], [], ['2', null]]) {
    const out = flt({ range: 'A1:C5', filters: [eng(1, 'xlFilterValues', list)] }).filters[0];
    assert.equal(out.unsupported, true, JSON.stringify(list));
    assert.match(out.reason, /^NO_WRITE_PATH:/, JSON.stringify(list));
    assert.deepEqual(out.engine.criteria1, list, 'the raw list is echoed, not joined');
    assert.equal(out.value, undefined, JSON.stringify(list));
    assert.equal(out.values, undefined, JSON.stringify(list));
  }
  // the operator the engine labels a list with: xlFilterValues (how it is made) and xlOr (how filter.cjs documents its read-back) are lists; anything else is not a shape the writer can create
  for (const op of ['xlOr', '']) assert.deepEqual(flt({ range: 'A1:C5', filters: [eng(1, op, ['2', '4'])] }).filters[0], { field: 1, values: ['2', '4'] }, `operator ${JSON.stringify(op)}`);
  for (const op of ['xlAnd', 'xlTop10Items', 'xlFilterCellColor']) assert.equal(flt({ range: 'A1:C5', filters: [eng(1, op, ['2', '4'])] }).filters[0].unsupported, true, `operator ${op}`);
  // a list with a second criterion, or on a pair operator, is not a shape the writer can create
  assert.equal(flt({ range: 'A1:C5', filters: [eng(1, 'xlFilterValues', ['2', '4'], '=5')] }).filters[0].unsupported, true);
  assert.equal(flt({ range: 'A1:C5', filters: [eng(1, 'xlOr', ['2', '4'], '=5')] }).filters[0].unsupported, true);
  assert.equal(flt({ range: 'A1:C5', filters: [eng(1, 'xlAnd', ['>1', '<5'], '<9')] }).filters[0].unsupported, true);
});

test('FILTER (L-10b r1): the raw engine shape is a DIAGNOSTIC next to the filters (engineShapes: operator, criteria1 / criteria2 type + length), in the same order, so a LIVE run sees it without going through the projection', () => {
  const out = flt({ range: 'A1:C5', filters: [eng(2, 'xlOr', '=3'), eng(1, 'xlFilterValues', ['2', '4']), eng(3, 'xlAnd', '>1', '<5')] });
  assert.deepEqual(out.engineShapes, [
    { field: 2, operator: 'xlOr', criteria1: 'string', criteria2: 'null' },
    { field: 1, operator: 'xlFilterValues', criteria1: 'array', criteria1Length: 2, criteria2: 'null' },
    { field: 3, operator: 'xlAnd', criteria1: 'string', criteria2: 'string' },
  ]);
  assert.deepEqual(flt({ range: null, filters: [] }).engineShapes, [], 'no AutoFilter: an empty list');
  assert.equal(flt(undefined).engineShapes, undefined, 'a whole-sheet gap has no shapes (filters is null)');
});

test('FILTER: no AutoFilter is an EMPTY read (range null, filters []); a cleared filter keeps its controls as on:false; several fields stay in order', () => {
  assert.deepEqual(flt({ range: null, filters: [] }), { sheet: 'S', range: null, filters: [], engineShapes: [] });
  assert.deepEqual(flt({ range: 'A1:C5', filters: [eng(2, 'xlOr', '=3', null, false)] }).filters, [{ field: 2, on: false }]);
  assert.deepEqual(flt({ range: 'A1:C5', filters: [eng(1, 'xlOr', '=x'), eng(3, 'xlAnd', '>1', '<5')] }).filters.map(f => f.field), [1, 3]);
});

test('FILTER: what the writer cannot create is NAMED (NO_WRITE_PATH), never turned into a value', () => {
  for (const f of [eng(1, 'xlAnd', '>5'), eng(1, 'xlAnd', '>=5', '<9'), eng(1, 'xlOr', '=a', '<9'), eng(1, 'xlOr', '=a', '=a'), eng(1, 'xlTop10Items', '10', null), eng(1, 'xlOr', '', '=b'), eng(1, 'xlFilterCellColor', 'red', 'blue')]) {
    const out = flt({ range: 'A1:C5', filters: [f] }).filters[0];
    assert.equal(out.unsupported, true, JSON.stringify(f));
    assert.match(out.reason, /^NO_WRITE_PATH:/);
    assert.deepEqual(out.engine, { operator: f.operator, criteria1: f.criteria1, criteria2: f.criteria2 });
    assert.equal(out.value, undefined);
  }
  assert.match(flt({ range: 'A1:C5', filters: [{ field: 0, on: true, operator: 'xlOr', criteria1: '=a', criteria2: null }] }).filters[0].reason, /^UNOBSERVABLE:field-identity/);
});

test('FILTER: whole-sheet gaps are strings on the result (filters null), never "no filter"; a missing sheet is SHEET_NOT_FOUND', () => {
  assert.deepEqual(flt(undefined), { sheet: 'S', range: null, filters: null, reason: 'UNOBSERVABLE:not-read' });
  assert.equal(flt('UNOBSERVABLE:field-identity-mismatch').reason, 'UNOBSERVABLE:field-identity-mismatch');
  assert.equal(flt({ range: 'A1', filters: 'x' }).reason, 'UNOBSERVABLE:unexpected-shape');
  assert.deepEqual(projectFilterRead({ sheets: SHEETS, filterReads: {} }, { sheet: 'Nope' }), { sheet: 'Nope', error: 'SHEET_NOT_FOUND' });
});

// ---------------------------------------------------------------------------------------------------- page setup
const PAGE_ENGINE = { orientation: 'xlLandscape', printGridlines: true, printHeadings: false, topMargin: 25, bottomMargin: 19.1, leftMargin: 17.8, rightMargin: 17.8 };
const PRINT_ENGINE = { fitW: 1, fitH: 0, fitFlag: true, oddHeader: '&CReport', oddFooter: null, firstHeader: null, firstFooter: null, evenHeader: null, evenFooter: null, diffFirst: false, diffOddEven: false,
  hCenter: true, vCenter: false, useFirstPage: true, firstPage: 5, scale: 100, width: 297, height: 420, rowBreaks: [{ id: 10, man: true }, { id: 30, man: false }], colBreaks: [{ id: 4, man: true }] };
const setup = (pageReads, printReads) => projectPageSetupRead({ sheets: SHEETS, pageReads: pageReads === undefined ? null : { S: pageReads }, printReads: printReads === undefined ? null : { S: printReads } }, A);

test('PAGE ROUND TRIP: set_page_layout fields (orientation, gridlines, headings, margins in mm) read back as the request', () => {
  const want = page.plan({ intent: 'set_page_layout', sheet: 'S', orientation: 'landscape', printGridlines: true, printHeadings: false, topMargin: 25 }, 0).operation.want;
  const out = setup({ ...PAGE_ENGINE, orientation: want.orientation, printGridlines: want.printGridlines, printHeadings: want.printHeadings, topMargin: want.topMargin }, undefined).page;
  assert.deepEqual([out.orientation, out.printGridlines, out.printHeadings, out.topMargin], ['landscape', true, false, 25]);
  assert.equal(out.unrepresentable, undefined);
});

test('PRINT ROUND TRIP: every set_print_* writer shape reads back as its request (fit, header, break, first page number, centering, scale, paper, flag)', () => {
  const w = (step) => print.plan({ sheet: 'S', ...step }, 0).operation.want;
  const fit = w({ intent: 'set_print_fit', fitToWidth: 1, fitToHeight: 0 });
  const hdr = w({ intent: 'set_header_footer', slot: 'oddHeader', value: '&CReport' });
  const row = w({ intent: 'set_page_break', axis: 'row', at: 11 });
  const col = w({ intent: 'set_page_break', axis: 'column', at: 'E' });
  const fp = w({ intent: 'set_first_page_number', value: 5 });
  const hc = w({ intent: 'set_print_centering', horizontal: true });
  const sc = w({ intent: 'set_print_scale', scale: 80 });
  const pp = w({ intent: 'set_paper_size', paper: 'A3' });
  const fl = w({ intent: 'set_header_footer_flag', flag: 'differentFirst', value: true });
  const eng = { ...PRINT_ENGINE, fitW: fit.fitW, fitH: fit.fitH, oddHeader: hdr.value, rowBreaks: [{ id: row.breakId, man: true }], colBreaks: [{ id: col.breakId, man: true }],
    firstPage: fp.firstPage, useFirstPage: true, hCenter: hc.value, scale: sc.scale, fitFlag: false, width: pp.width, height: pp.height, diffFirst: fl.value };
  const out = setup(undefined, eng).print;
  assert.deepEqual([out.fitToWidth, out.fitToHeight, out.oddHeader, out.rowBreaks, out.columnBreaks, out.firstPageNumber, out.horizontal, out.scale, out.paper, out.differentFirst],
    [1, 0, '&CReport', [11], ['E'], 5, true, 80, 'A3', true]);
  assert.equal(out.unrepresentable, undefined);
});

test('PRINT: every paper of the editor\'s own table round-trips by name; an unknown size is NAMED (paper null); an automatic break is counted, never listed', () => {
  assert.equal(Object.keys(PAPER_OF).length, Object.keys(print.PAPER).length);
  for (const [name, [width, height]] of Object.entries(print.PAPER)) assert.equal(setup(undefined, { ...PRINT_ENGINE, width, height }).print.paper, name);
  const odd = setup(undefined, { ...PRINT_ENGINE, width: 200, height: 300 }).print;
  assert.equal(odd.paper, null);
  assert.deepEqual(odd.unrepresentable, [{ field: 'paper', engine: { width: 200, height: 300 }, reason: "not one of the editor's paper sizes" }]);
  const out = setup(undefined, PRINT_ENGINE).print;
  assert.deepEqual(out.rowBreaks, [11]);
  assert.deepEqual(out.automaticBreaks, { rows: 1, columns: 0 });
});

test('PRINT: no first page number reads null; out-of-schema values are NAMED and null (scale, orientation, a margin, a header)', () => {
  assert.equal(setup(undefined, { ...PRINT_ENGINE, useFirstPage: false, firstPage: 1 }).print.firstPageNumber, null);
  const p = setup(undefined, { ...PRINT_ENGINE, scale: 9, oddHeader: 5 }).print;
  assert.deepEqual([p.scale, p.oddHeader], [null, null]);
  assert.deepEqual(p.unrepresentable.map(u => u.field), ['scale', 'oddHeader']);
  const g = setup({ ...PAGE_ENGINE, orientation: 'xlOther', leftMargin: 999 }, undefined).page;
  assert.deepEqual([g.orientation, g.leftMargin], [null, null]);
  assert.deepEqual(g.unrepresentable.map(u => u.field), ['orientation', 'leftMargin']);
});

test('PAGE SETUP: the two halves fail INDEPENDENTLY -- a named reason on one half, the other still read; strings, not-read and a bad shape are all named', () => {
  const both = setup('UNOBSERVABLE:per-sheet-read-not-verified', PRINT_ENGINE);
  assert.deepEqual(both.page, { reason: 'UNOBSERVABLE:per-sheet-read-not-verified' });
  assert.equal(both.print.fitToWidth, 1);
  const other = setup(PAGE_ENGINE, 'UNOBSERVABLE:BRIDGE:model-missing');
  assert.equal(other.page.orientation, 'landscape');
  assert.deepEqual(other.print, { reason: 'UNOBSERVABLE:BRIDGE:model-missing' });
  const none = setup(undefined, undefined);
  assert.deepEqual([none.page.reason, none.print.reason], ['UNOBSERVABLE:not-read', 'UNOBSERVABLE:not-read']);
  assert.equal(setup(undefined, { fitW: 1 }).print.reason, 'UNOBSERVABLE:unexpected-shape');
  assert.deepEqual(projectPageSetupRead({ sheets: SHEETS, pageReads: {}, printReads: {} }, { sheet: 'Nope' }), { sheet: 'Nope', error: 'SHEET_NOT_FOUND' });
});

test('PAGE SETUP: print area / print titles are said to be OUTSIDE this read (defined names), on every result', () => {
  assert.ok(setup(PAGE_ENGINE, PRINT_ENGINE).caveats.some(c => /print area and print titles are defined names/.test(c)));
});

// ---------------------------------------------------------------------------------------------------- the tools
test('TOOL SURFACE: the three reads are ACCEPTED (sheet only), each request asks the driver for exactly its own read, no unknown field, sheet required', () => {
  for (const op of ['read_freeze_panes', 'read_page_setup', 'read_filter']) {
    assert.equal(rt.OPERATIONS.xlsx[op].status, rt.STATUS.ACCEPTED, op);
    assert.ok(/L-10/.test(rt.OPERATIONS.xlsx[op].evidence), op);
    assert.equal(rt.validateReadRequest('xlsx', { fileId: '1', operation: op, sheet: 'S' }).ok, true, op);
    assert.equal(rt.validateReadRequest('xlsx', { fileId: '1', operation: op }).ok, false, `${op}: sheet is required`);
    assert.equal(rt.validateReadRequest('xlsx', { fileId: '1', operation: op, sheet: 'S', range: 'A1' }).ok, false, `${op}: no range`);
  }
  assert.deepEqual(rt.xlsxRequest('read_freeze_panes', { sheet: 'S' }), { windows: [], panes: true });
  assert.deepEqual(rt.xlsxRequest('read_page_setup', { sheet: 'S' }), { windows: [], pageSheets: ['S'], printSheets: ['S'] });
  assert.deepEqual(rt.xlsxRequest('read_filter', { sheet: 'S' }), { windows: [], filterSheets: ['S'] });
  assert.equal(rt.validateReadRequest('xlsx', { fileId: '1', operation: 'read_print_rendering' }).code, 'REFUSED_NOT_MEASURED', 'the RENDERED print stays refused');
});
