'use strict';

/**
 * The M11.4a.1 WRITE refusal matrix (shared by the M11.4a.1 handler test and the M11.4a.2 wiring test): every case
 * must be refused PRE_DISPATCH, for the whole task, before a broker / browser / driver / executor exists.
 */

const FILE = '1303626';

const VALID = Object.freeze({
  write_range: { intent: 'write_range', sheet: 'Budget', range: 'A1:B2', values: [['Name', 3], ['pear', 7.5]] },
  insert_rows: { intent: 'insert_rows', sheet: 'Budget', at: 3, count: 2, anchor: { column: 'A', value: 'pear' } },
  delete_rows: { intent: 'delete_rows', sheet: 'Budget', at: 2, count: 1, anchor: { column: 'A', value: 7 } },
  rename_sheet: { intent: 'rename_sheet', sheet: 'Budget', name: 'Budget_2026' },
  unmerge_range: { intent: 'unmerge_range', sheet: 'Budget', range: 'C1:D1' },
  set_chart_title: { intent: 'set_chart_title', sheet: 'Budget', name: 'Chart 1', title: 'Sales 2026' },
  // W1a (owner decision 2026-10-13; census ACCEPTED_WRITE)
  merge_range: { intent: 'merge_range', sheet: 'Budget', range: 'E1:F1' },
  sort_range: { intent: 'sort_range', sheet: 'Budget', range: 'A1:B3', keyColumn: 'A', order: 'asc', hasHeaders: false },
  add_conditional_format: { intent: 'add_conditional_format', sheet: 'Budget', range: 'A1:A3', rule: { type: 'cellValue', operator: 'greater', formula1: 5 } },
  set_print_area: { intent: 'set_print_area', sheet: 'Budget', range: 'A1:D10' },
  layout_range: { intent: 'layout_range', sheet: 'Budget', columns: 'B', width: 20 },
  // W1d (owner decision 2026-10-13; census ACCEPTED_WRITE)
  freeze_panes: { intent: 'freeze_panes', sheet: 'Budget', rows: 1 },
  unfreeze_panes: { intent: 'unfreeze_panes', sheet: 'Budget' },
  set_validation: { intent: 'set_validation', sheet: 'Budget', range: 'A1:A3', type: 'wholeNumber', operator: 'between', alertStyle: 'stop', min: 1, max: 10 },
  clear_validation: { intent: 'clear_validation', sheet: 'Budget', range: 'A1:A3' },
  delete_conditional_format: { intent: 'delete_conditional_format', sheet: 'Budget', range: 'A1:A3', rule: { type: 'cellValue', operator: 'greater', formula1: 5 } },
  clear_filter: { intent: 'clear_filter', sheet: 'Budget', range: 'A1:B3' },
  clear_print_area: { intent: 'clear_print_area', sheet: 'Budget' },
  create_sheet: { intent: 'create_sheet', name: 'NewSheet' },
  rename_chart: { intent: 'rename_chart', sheet: 'Budget', name: 'Chart 1', newName: 'Chart 2' },
  set_chart_axis_title: { intent: 'set_chart_axis_title', sheet: 'Budget', name: 'Chart 1', axis: 'horizontal', title: 'Months' },
  // X-O04/11 (owner decision "B", internal tracker): marker/gridlines only -- the other six
  // internal set_chart_style targets stay unreleased (REFUSALS below has the named rejections)
  set_chart_style: { intent: 'set_chart_style', sheet: 'Budget', name: 'Chart 1', target: 'marker', seriesIndex: 0, fill: '#808080' },
  // GH #140 U3 (census ACCEPTED_WRITE)
  create_pivot: { intent: 'create_pivot', name: 'PV1', sheet: 'Budget', range: 'A1:C5', rows: 'Name', columns: 'Qty', sum: 'Total', pivotSheet: 'K', at: 'H1' },
});
const req = (steps, extra = {}) => ({ fileId: FILE, task: { id: 't1', steps }, ...extra });
const wr = (range, values, id) => ({ intent: 'write_range', ...(id ? { id } : {}), sheet: 'Budget', range, values });

const deep = (o, k, v) => ({ ...o, [k]: v });
const REFUSALS = [
  // [label, args, code]
  ['unknown intent', req([{ intent: 'make_coffee', sheet: 'Budget' }]), 'REFUSED_UNKNOWN_INTENT'],
  // GH #106: clear_range is allowlisted now; L-06/b released move_range too; GH #138 U1 / U3 / U4 released set_defined_name,
  // set_page_layout and the rest of the print family (B10 now empty, no XLSX intent depends on it); GH #140 U3 / U4 released
  // create_pivot and refresh_pivot -- since then NO routed XLSX intent is outside the allowlist (measured: xlsxRoutedIntents()
  // minus XLSX_WRITE_ALLOWLIST is empty), so REFUSED_NOT_ALLOWLISTED has no live XLSX example left. A refresh_pivot with a layout
  // key is refused by the strict schema (UNKNOWN_FIELD).
  // A clear_range without its range is still refused on its shape
  ['refresh_pivot refuses a layout key', req([{ intent: 'refresh_pivot', pivot: 'PE', sourceSheet: 'Src', sourceRange: 'A1:C7', rows: 'reg' }]), 'UNKNOWN_FIELD'],
  ['clear_range requires a range', req([{ intent: 'clear_range', sheet: 'Budget' }]), 'REFUSED_SHAPE'],
  // W1d: freeze_panes is allowlisted now, but still refuses a shape neither rows nor columns names (CONSTRAINTS.freeze_panes)
  ['freeze_panes requires rows and/or columns', req([{ intent: 'freeze_panes', sheet: 'Budget' }]), 'REFUSED_SHAPE'],
  // W1d: create_sheet is allowlisted now, but a cell-reference-shaped name is still refused (bareNameProblem)
  ['create_sheet name shaped like a cell reference', req([{ intent: 'create_sheet', name: 'A1' }]), 'REFUSED_SHAPE'],
  // X-O04/11: the six internal set_chart_style targets are refused at the SCHEMA level (target enum
  // admits only marker/gridlines) -- never reaching charts.cjs's own, wider STYLE_TARGETS at all.
  ['set_chart_style plotArea target (internal-only, not released)', req([{ intent: 'set_chart_style', sheet: 'Budget', name: 'Chart 1', target: 'plotArea', fill: '#808080' }]), 'REFUSED_SHAPE'],
  ['set_chart_style gridlines with fill (gridlines has no fill)', req([{ intent: 'set_chart_style', sheet: 'Budget', name: 'Chart 1', target: 'gridlines', axis: 'vertical', fill: '#808080' }]), 'REFUSED_SHAPE'],
  ['set_chart_style gridlines without axis', req([{ intent: 'set_chart_style', sheet: 'Budget', name: 'Chart 1', target: 'gridlines', outline: { width: 1, color: '#000000' } }]), 'REFUSED_SHAPE'],
  ['set_chart_style gridlines without outline', req([{ intent: 'set_chart_style', sheet: 'Budget', name: 'Chart 1', target: 'gridlines', axis: 'vertical' }]), 'REFUSED_SHAPE'],
  ['set_chart_style marker without seriesIndex', req([{ intent: 'set_chart_style', sheet: 'Budget', name: 'Chart 1', target: 'marker', fill: '#808080' }]), 'REFUSED_SHAPE'],
  ['set_chart_style marker with axis (axis is gridlines-only)', req([{ intent: 'set_chart_style', sheet: 'Budget', name: 'Chart 1', target: 'marker', seriesIndex: 0, axis: 'vertical', fill: '#808080' }]), 'REFUSED_SHAPE'],
  ['set_chart_style marker with neither fill nor outline', req([{ intent: 'set_chart_style', sheet: 'Budget', name: 'Chart 1', target: 'marker', seriesIndex: 0 }]), 'REFUSED_SHAPE'],
  // L-05/3: copy_sheet is released (the B03 boundary lifted); its malformed shapes are refused before dispatch
  ['copy_sheet onto its own name (case-insensitively)', req([{ intent: 'copy_sheet', sheet: 'Budget', name: 'budget' }]), 'REFUSED_SHAPE'],
  // X-N01: format_range is released with format.numberFormat only; any other format key stays behind B05-WRITE
  // GitHub #88/#89/#93 released fill, font and align, L-01 wrap -- every OTHER key is still held behind the boundary
  ['boundary B05W format_range (a non-released key)', req([{ intent: 'format_range', sheet: 'Budget', range: 'A1', format: { textDirection: 'rtl' } }]), 'REFUSED_TEMPORARY_BOUNDARY'],
  // GitHub #89: font must carry at least one of color / bold / italic -- an empty object is refused, named
  ['format_range font empty object', req([{ intent: 'format_range', sheet: 'Budget', range: 'A1', format: { font: {} } }]), 'REFUSED_SHAPE'],
  // GitHub #89 (+ L-02/1 name and size, L-02/2 underline and strikeout): an unsupported key in the nested object is still refused by the strict schema
  ['format_range font unsupported key (subscript)', req([{ intent: 'format_range', sheet: 'Budget', range: 'A1', format: { font: { subscript: true } } }]), 'UNKNOWN_FIELD'],
  // GitHub #93: align must carry at least one of horizontal / vertical -- an empty object is refused, named
  ['format_range align empty object', req([{ intent: 'format_range', sheet: 'Budget', range: 'A1', format: { align: {} } }]), 'REFUSED_SHAPE'],
  // GitHub #93 (+ L-03/2 released rotation / readingOrder / indent / shrink): an unknown key in the nested align object is still refused by the strict schema
  ['format_range align unsupported key (direction)', req([{ intent: 'format_range', sheet: 'Budget', range: 'A1', format: { align: { direction: 'rtl' } } }]), 'UNKNOWN_FIELD'],
  // GH #138 U4: set_print_fit is released now (B10 lifted) -- a missing field is REFUSED_SHAPE, not
  // REFUSED_TEMPORARY_BOUNDARY (covered in the new-schema cases below); B05W/B18/B21 still cover the
  // boundary-refusal MECHANISM itself.
  ['boundary B18 remove_list', req([{ intent: 'remove_list' }]), 'REFUSED_TEMPORARY_BOUNDARY'],
  ['boundary B21 edit_header', req([{ intent: 'edit_header' }]), 'REFUSED_TEMPORARY_BOUNDARY'],
  ['DOCX intent rewrite_paragraph', req([{ intent: 'rewrite_paragraph', text: 'x' }]), 'REFUSED_CROSS_FORMAT'],
  ['DOCX intent insert_table_row_after', req([{ intent: 'insert_table_row_after' }]), 'REFUSED_CROSS_FORMAT'],
  ['READ operation read_range', req([{ intent: 'read_range', sheet: 'Budget', range: 'A1' }]), 'REFUSED_READ_OPERATION'],
  // unsupported shapes (REFUSED_SHAPE, never PARTIAL)
  // W1a HATOKOR +1: formulas is now a recognized field (was UNKNOWN_FIELD pre-W1a) -- this is the accepted shape,
  // moved to VALID_SINGLE_STEP coverage via a dedicated multi-step composition check below instead
  // X-T01: { text } is released on the product surface; a date-typed value is not
  ['write_range date-typed value (not released)', req([wr('A1', [[{ date: '2026-10-03', format: 'yyyy-mm-dd' }]])]), 'REFUSED_SHAPE'],
  ['write_range null cell', req([wr('A1', [[null]])]), 'REFUSED_SHAPE'],
  ['write_range boolean cell', req([wr('A1', [[true]])]), 'REFUSED_SHAPE'],
  ['write_range empty string', req([wr('A1', [['']])]), 'REFUSED_SHAPE'],
  ['write_range leading space', req([wr('A1', [[' x']])]), 'REFUSED_SHAPE'],
  ['write_range formula-like text', req([wr('A1', [['=SUM(A2)']])]), 'REFUSED_SHAPE'],
  ['write_range number-like text', req([wr('A1', [['001']])]), 'REFUSED_SHAPE'],
  ['write_range $A$1', req([wr('$A$1', [['x']])]), 'REFUSED_SHAPE'],
  ['write_range sheet-qualified range', req([wr('Budget!A1', [['x']])]), 'REFUSED_SHAPE'],
  ['write_range lower-case range', req([wr('a1', [['x']])]), 'REFUSED_SHAPE'],
  ['write_range matrix mismatch', req([wr('A1:B2', [['x']])]), 'REFUSED_SHAPE'],
  ['write_range reversed range', req([wr('B2:A1', [['x', 'y'], ['z', 'w']])]), 'REFUSED_SHAPE'],
  ['write_range over 1000 cells', req([wr('A1:A1001', Array.from({ length: 1001 }, (_, i) => [`v${i}`]))]), 'REFUSED_SHAPE'],
  ['insert_rows count 101', req([{ ...VALID.insert_rows, count: 101 }]), 'REFUSED_SHAPE'],
  ['insert_rows at 0', req([{ ...VALID.insert_rows, at: 0 }]), 'REFUSED_SHAPE'],
  ['insert_rows apostrophe sheet (L-21/2: sheet-name-apostrophe-unsupported)', req([{ ...VALID.insert_rows, sheet: "O'Brien" }]), 'REFUSED_SHAPE'],
  ['insert_rows missing anchor', req([{ ...VALID.insert_rows, anchor: undefined }]), 'REFUSED_SHAPE'],
  ['delete_rows anchor.row (columns lane)', req([{ ...VALID.delete_rows, anchor: { row: 1, value: 'x' } }]), 'UNKNOWN_FIELD'],
  ['rename_sheet apostrophe new name (L-21/2)', req([{ ...VALID.rename_sheet, name: "D'Arcy" }]), 'REFUSED_SHAPE'],
  ['rename_sheet cell-ref name', req([{ ...VALID.rename_sheet, name: 'AB12' }]), 'REFUSED_SHAPE'],
  ['rename_sheet same name', req([{ ...VALID.rename_sheet, name: 'BUDGET' }]), 'REFUSED_SHAPE'],
  ['unmerge_range across', req([{ ...VALID.unmerge_range, across: true }]), 'UNKNOWN_FIELD'],
  ['set_chart_title two lines', req([{ ...VALID.set_chart_title, title: 'a\nb' }]), 'REFUSED_SHAPE'],
  ['set_chart_title apostrophe sheet (L-21/2)', req([{ ...VALID.set_chart_title, sheet: "O'Brien" }]), 'REFUSED_SHAPE'],
  // W1a: merge_range (unknown field, wrong shape)
  ['merge_range unknown field', req([{ ...VALID.merge_range, bogus: 1 }]), 'UNKNOWN_FIELD'],
  ['merge_range single cell', req([{ ...VALID.merge_range, range: 'A1' }]), 'REFUSED_SHAPE'],
  ['merge_range across on a single column', req([{ ...VALID.merge_range, range: 'A1:A2', across: true }]), 'REFUSED_SHAPE'],
  // W1a: sort_range (unknown field, wrong shape)
  ['sort_range unknown field', req([{ ...VALID.sort_range, bogus: 1 }]), 'UNKNOWN_FIELD'],
  ['sort_range keyColumn outside range', req([{ ...VALID.sort_range, keyColumn: 'Z' }]), 'REFUSED_SHAPE'],
  ['sort_range bad order', req([{ ...VALID.sort_range, order: 'ascending' }]), 'REFUSED_SHAPE'],
  // W1a: add_conditional_format (unknown field, wrong shape)
  ['add_conditional_format unknown field', req([{ ...VALID.add_conditional_format, bogus: 1 }]), 'UNKNOWN_FIELD'],
  ['add_conditional_format unknown rule type', req([{ ...VALID.add_conditional_format, rule: { type: 'zebra' } }]), 'REFUSED_SHAPE'],
  ['add_conditional_format bad hex fill', req([{ ...VALID.add_conditional_format, rule: { type: 'cellValue', operator: 'greater', formula1: 5, fill: 'red' } }]), 'REFUSED_SHAPE'],
  // W1a: set_print_area (unknown field, wrong shape)
  ['set_print_area unknown field', req([{ ...VALID.set_print_area, bogus: 1 }]), 'UNKNOWN_FIELD'],
  ['set_print_area range and areas together', req([{ ...VALID.set_print_area, areas: ['A1:B1', 'C1:D1'] }]), 'REFUSED_SHAPE'],
  ['set_print_area neither range nor areas', req([{ intent: 'set_print_area', sheet: 'Budget' }]), 'REFUSED_SHAPE'],
  ['set_print_area one-area areas array', req([{ intent: 'set_print_area', sheet: 'Budget', areas: ['A1:B1'] }]), 'REFUSED_SHAPE'],
  // W1a: layout_range (unknown field, wrong shape)
  ['layout_range unknown field', req([{ ...VALID.layout_range, bogus: 1 }]), 'UNKNOWN_FIELD'],
  ['layout_range both columns and rows', req([{ ...VALID.layout_range, rows: '2:5' }]), 'REFUSED_SHAPE'],
  ['layout_range width and height together', req([{ ...VALID.layout_range, height: 20 }]), 'REFUSED_SHAPE'],
  ['layout_range width out of range', req([{ ...VALID.layout_range, width: 300 }]), 'REFUSED_SHAPE'],
  ['layout_range width on rows', req([{ intent: 'layout_range', sheet: 'Budget', rows: '2', width: 20 }]), 'REFUSED_SHAPE'],
  // W1a scope +1 (GitHub #74): write_range's `formulas` field
  ['write_range formulas non-equals prefix', req([{ intent: 'write_range', sheet: 'Budget', range: 'A1', formulas: [['SUM(A2)']] }]), 'REFUSED_SHAPE'],
  ['write_range formulas external workbook ref', req([{ intent: 'write_range', sheet: 'Budget', range: 'A1', formulas: [['=[Book1.xlsx]Sheet1!A1']] }]), 'REFUSED_SHAPE'],
  ['write_range formulas over the length ceiling', req([{ intent: 'write_range', sheet: 'Budget', range: 'A1', formulas: [[`=${'A'.repeat(8192)}`]] }]), 'REFUSED_SHAPE'],
  ['write_range neither values nor formulas', req([{ intent: 'write_range', sheet: 'Budget', range: 'A1' }]), 'REFUSED_SHAPE'],
  // malformed / empty / limit
  ['task is an array', { fileId: FILE, task: [VALID.write_range] }, 'MALFORMED_TASK'],
  ['task missing', { fileId: FILE }, 'MALFORMED_TASK'],
  ['steps not an array', { fileId: FILE, task: { steps: VALID.write_range } }, 'MALFORMED_TASK'],
  ['step not an object', req(['write_range']), 'MALFORMED_TASK'],
  ['intent missing', req([{ sheet: 'Budget', range: 'A1', values: [['x']] }]), 'MALFORMED_TASK'],
  ['empty task', req([]), 'EMPTY_TASK'],
  ['51 steps', req(Array.from({ length: 51 }, (_, i) => wr(`A${i + 1}`, [['x']], `w${i}`))), 'LIMIT'],
  ['bad fileId', { ...req([VALID.write_range]), fileId: '../etc' }, 'INVALID_ARGS'],
  ['bad task id', { fileId: FILE, task: { id: 'a b', steps: [VALID.write_range] } }, 'MALFORMED_TASK'],
  // credential injection, any depth
  ['top-level token', req([VALID.write_range], { token: 'x' }), 'CREDENTIAL_FIELD'],
  ['task-level password', { fileId: FILE, task: { steps: [VALID.write_range], password: 'x' } }, 'CREDENTIAL_FIELD'],
  ['step-level keyName', req([deep(VALID.write_range, 'keyName', 'k')]), 'CREDENTIAL_FIELD'],
  ['anchor-level credential', req([{ ...VALID.insert_rows, anchor: { column: 'A', value: 'x', credential: 'c' } }]), 'CREDENTIAL_FIELD'],
  ['cell-level token object', req([wr('A1', [[{ token: 'secret' }]])]), 'CREDENTIAL_FIELD'],
  ['app_password variant', req([deep(VALID.rename_sheet, 'app_password', 'x')]), 'CREDENTIAL_FIELD'],
  // requester spoof
  ['top-level requester', req([VALID.write_range], { requester: 'maintainer' }), 'REQUESTER_OVERRIDE'],
  ['step-level uid', req([deep(VALID.unmerge_range, 'uid', 'maintainer')]), 'REQUESTER_OVERRIDE'],
  ['task-level fileOwner', { fileId: FILE, task: { steps: [VALID.write_range], fileOwner: 'maintainer' } }, 'REQUESTER_OVERRIDE'],
  ['step-level user', req([deep(VALID.set_chart_title, 'user', 'maintainer')]), 'CREDENTIAL_FIELD'],
  // broker / NC route override
  ['top-level brokerUrl', req([VALID.write_range], { brokerUrl: 'http://evil' }), 'ROUTE_OVERRIDE'],
  ['task-level url', { fileId: FILE, task: { steps: [VALID.write_range], url: 'http://evil' } }, 'ROUTE_OVERRIDE'],
  ['step-level nc_identity', req([deep(VALID.write_range, 'nc_identity', 'x')]), 'ROUTE_OVERRIDE'],
  // unknown fields
  ['top-level strategy', req([VALID.write_range], { strategy: 'transaction' }), 'UNKNOWN_FIELD'],
  ['task-level preconditions', { fileId: FILE, task: { steps: [VALID.write_range], preconditions: [] } }, 'UNKNOWN_FIELD'],
  ['step-level unknown', req([deep(VALID.write_range, 'style', true)]), 'UNKNOWN_FIELD'],
  // mixed and composition: the WHOLE task is refused, the allowlisted part does not run
  // GH #138 U4: set_print_fit is released now (B10 lifted) -- format_range's B05-WRITE boundary
  // takes over as the mixed-task boundary example.
  ['mixed write_range + B05W format_range', req([VALID.write_range, { intent: 'format_range', sheet: 'Budget', range: 'A1', format: { textDirection: 'rtl' } }]), 'REFUSED_TEMPORARY_BOUNDARY'],
  // W1d: freeze_panes is allowlisted now, but its selector is a fixed 'A1' placeholder with no declared footprint --
  // ANY other step touching the same sheet clashes at composition (verified: xlsxAdapter.planOperations ->
  // 'overlapping-writes-refused: s0 and s1 on Budget (axes not declared)')
  ['mixed write_range + freeze_panes clash at composition', req([VALID.write_range, { intent: 'freeze_panes', sheet: 'Budget', rows: 1 }]), 'REFUSED_COMPOSITION'],
  ['mixed write_range + DOCX', req([VALID.write_range, { intent: 'rewrite_paragraph' }]), 'REFUSED_CROSS_FORMAT'],
  ['mixed write_range + bad second shape', req([VALID.write_range, wr('E1', [['=x']])]), 'REFUSED_SHAPE'],
  ['rename_sheet + write_range', req([VALID.rename_sheet, wr('E1', [['x']])]), 'REFUSED_COMPOSITION'],
  ['write_range + insert_rows', req([wr('E1', [['x']]), VALID.insert_rows]), 'REFUSED_COMPOSITION'],
  ['two overlapping write_range', req([wr('A1:B1', [['a', 'b']], 'w1'), wr('B1', [['c']], 'w2')]), 'REFUSED_COMPOSITION'],
  ['duplicate step ids', req([wr('A1', [['a']], 'w'), wr('C1', [['c']], 'w')]), 'REFUSED_COMPOSITION'],
  ['two titles on one chart', req([{ ...VALID.set_chart_title, id: 'c1' }, { ...VALID.set_chart_title, id: 'c2', title: 'Other' }]), 'REFUSED_COMPOSITION'],
];

module.exports = { FILE, VALID, REFUSALS, req, wr };
