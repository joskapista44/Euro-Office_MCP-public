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
});
const req = (steps, extra = {}) => ({ fileId: FILE, task: { id: 't1', steps }, ...extra });
const wr = (range, values, id) => ({ intent: 'write_range', ...(id ? { id } : {}), sheet: 'Budget', range, values });

const deep = (o, k, v) => ({ ...o, [k]: v });
const REFUSALS = [
  // [label, args, code]
  ['unknown intent', req([{ intent: 'make_coffee', sheet: 'Budget' }]), 'REFUSED_UNKNOWN_INTENT'],
  ['non-allowlisted merge_range', req([{ intent: 'merge_range', sheet: 'Budget', range: 'A1:B1' }]), 'REFUSED_NOT_ALLOWLISTED'],
  ['non-allowlisted clear_range', req([{ intent: 'clear_range', sheet: 'Budget', range: 'A1' }]), 'REFUSED_NOT_ALLOWLISTED'],
  ['non-allowlisted create_sheet', req([{ intent: 'create_sheet', name: 'New' }]), 'REFUSED_NOT_ALLOWLISTED'],
  ['boundary B03 copy_sheet', req([{ intent: 'copy_sheet', sheet: 'Budget', name: 'Copy' }]), 'REFUSED_TEMPORARY_BOUNDARY'],
  ['boundary B05W format_range', req([{ intent: 'format_range', sheet: 'Budget', range: 'A1' }]), 'REFUSED_TEMPORARY_BOUNDARY'],
  ['boundary B10 set_print_fit', req([{ intent: 'set_print_fit', sheet: 'Budget' }]), 'REFUSED_TEMPORARY_BOUNDARY'],
  ['boundary B18 remove_list', req([{ intent: 'remove_list' }]), 'REFUSED_TEMPORARY_BOUNDARY'],
  ['boundary B21 edit_header', req([{ intent: 'edit_header' }]), 'REFUSED_TEMPORARY_BOUNDARY'],
  ['DOCX intent rewrite_paragraph', req([{ intent: 'rewrite_paragraph', text: 'x' }]), 'REFUSED_CROSS_FORMAT'],
  ['DOCX intent insert_table_row_after', req([{ intent: 'insert_table_row_after' }]), 'REFUSED_CROSS_FORMAT'],
  ['READ operation read_range', req([{ intent: 'read_range', sheet: 'Budget', range: 'A1' }]), 'REFUSED_READ_OPERATION'],
  // unsupported shapes (REFUSED_SHAPE, never PARTIAL)
  ['write_range formulas', req([{ ...VALID.write_range, values: undefined, formulas: [['=1', '=2'], ['=3', '=4']] }]), 'UNKNOWN_FIELD'],
  ['write_range typed value', req([wr('A1', [[{ text: '001' }]])]), 'REFUSED_SHAPE'],
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
  ['insert_rows quoted sheet', req([{ ...VALID.insert_rows, sheet: 'My Sheet' }]), 'REFUSED_SHAPE'],
  ['insert_rows missing anchor', req([{ ...VALID.insert_rows, anchor: undefined }]), 'REFUSED_SHAPE'],
  ['delete_rows anchor.row (columns lane)', req([{ ...VALID.delete_rows, anchor: { row: 1, value: 'x' } }]), 'UNKNOWN_FIELD'],
  ['rename_sheet quoted new name', req([{ ...VALID.rename_sheet, name: 'Budget 2026' }]), 'REFUSED_SHAPE'],
  ['rename_sheet cell-ref name', req([{ ...VALID.rename_sheet, name: 'AB12' }]), 'REFUSED_SHAPE'],
  ['rename_sheet same name', req([{ ...VALID.rename_sheet, name: 'BUDGET' }]), 'REFUSED_SHAPE'],
  ['unmerge_range across', req([{ ...VALID.unmerge_range, across: true }]), 'UNKNOWN_FIELD'],
  ['set_chart_title two lines', req([{ ...VALID.set_chart_title, title: 'a\nb' }]), 'REFUSED_SHAPE'],
  ['set_chart_title quoted sheet', req([{ ...VALID.set_chart_title, sheet: 'My Sheet' }]), 'REFUSED_SHAPE'],
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
  ['mixed write_range + B10 print', req([VALID.write_range, { intent: 'set_print_fit', sheet: 'Budget', wide: 1, tall: 1 }]), 'REFUSED_TEMPORARY_BOUNDARY'],
  ['mixed write_range + merge_range', req([VALID.write_range, { intent: 'merge_range', sheet: 'Budget', range: 'E1:F1' }]), 'REFUSED_NOT_ALLOWLISTED'],
  ['mixed write_range + DOCX', req([VALID.write_range, { intent: 'rewrite_paragraph' }]), 'REFUSED_CROSS_FORMAT'],
  ['mixed write_range + bad second shape', req([VALID.write_range, wr('E1', [['=x']])]), 'REFUSED_SHAPE'],
  ['rename_sheet + write_range', req([VALID.rename_sheet, wr('E1', [['x']])]), 'REFUSED_COMPOSITION'],
  ['write_range + insert_rows', req([wr('E1', [['x']]), VALID.insert_rows]), 'REFUSED_COMPOSITION'],
  ['two overlapping write_range', req([wr('A1:B1', [['a', 'b']], 'w1'), wr('B1', [['c']], 'w2')]), 'REFUSED_COMPOSITION'],
  ['duplicate step ids', req([wr('A1', [['a']], 'w'), wr('C1', [['c']], 'w')]), 'REFUSED_COMPOSITION'],
  ['two titles on one chart', req([{ ...VALID.set_chart_title, id: 'c1' }, { ...VALID.set_chart_title, id: 'c2', title: 'Other' }]), 'REFUSED_COMPOSITION'],
];

module.exports = { FILE, VALID, REFUSALS, req, wr };
