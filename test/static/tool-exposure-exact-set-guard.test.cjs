'use strict';

/**
 * Static gate: the product MCP exposes EXACTLY { product_status, xlsx_read, docx_read, xlsx_execute_task, docx_execute_task }
 * (owner decisions M11.3b.2 and M11.4b.1, owner channel; docx_execute_task = GH #228, PROPOSED STATE, owner decision
 * pending -- BEFORE the owner's GO this set is the four tools without it and EXPECTED_WITHHELD is ['docx_execute_task']). Nothing
 * else may appear and no other routed write capability is released. Each WRITE tool is served only while ITS OWN FORMAT's half of
 * the WRITE release gate holds (M11.4a.1/2/3 CLOSED + the exact accepted allowlist of that format). Changing the exposed set is a
 * release decision: this gate must be changed in the same diff, on purpose.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const { EXPOSED_CAPABILITY_TOOLS, TOOL_ROWS, exposedTools, capabilityInventory, writeReleaseGate, WRITE_RELEASE_PRECONDITIONS, ACCEPTED_XLSX_WRITE_ALLOWLIST, ACCEPTED_DOCX_WRITE_ALLOWLIST } = require('../../src/server/tools.cjs');
const { buildServerTools } = require('../../src/server/server-tools.cjs');
const W = require('../../src/server/write-tools.cjs');

const EXPECTED_EXPOSED = Object.freeze(['product_status', 'xlsx_read', 'docx_read', 'xlsx_execute_task', 'docx_execute_task']);
const EXPECTED_WITHHELD = Object.freeze([]);
const EXPECTED_CAPABILITY_TOOLS = Object.freeze(['xlsx_read', 'docx_read', 'xlsx_execute_task', 'docx_execute_task']);
// GH #228 (PROPOSED): the one DOCX WRITE intent the owner is asked to release
const DOCX_ALLOWLIST = Object.freeze(['rewrite_paragraph', 'insert_paragraph_before', 'insert_paragraph_after', 'delete_paragraph', 'move_paragraph_before', 'move_paragraph_after', 'apply_list', 'set_list_level', 'indent_list_item', 'outdent_list_item', 'remove_list', 'replace_table_cell', 'format_text', 'insert_table_row_before', 'insert_table_row_after', 'delete_table_row', 'insert_table_column_before', 'insert_table_column_after', 'delete_table_column', 'merge_table_cells', 'split_table_cell', 'format_table', 'set_page_setup', 'edit_header', 'edit_footer', 'define_style', 'modify_style', 'set_numbering', 'set_list_type', 'add_footnote', 'add_endnote', 'add_link', 'add_bookmark', 'resolve_comment', 'reply_comment', 'add_comment_on_match', 'set_image_size', 'set_image_wrap']);
// the owner-accepted XLSX WRITE set (M11.4a.0 six + X-N01 format_range, numberFormat-only, owner decision 2026-10-12;
// + W1a's 5 census-ACCEPTED_WRITE intents, owner decision 2026-10-13; + W1d's 10 more, same decision, 2026-10-13)
const ALLOWLIST = Object.freeze(['write_range', 'insert_rows', 'delete_rows', 'rename_sheet', 'unmerge_range', 'set_chart_title', 'format_range',
  'merge_range', 'sort_range', 'add_conditional_format', 'set_print_area', 'layout_range',
  'freeze_panes', 'unfreeze_panes', 'set_validation', 'clear_validation', 'delete_conditional_format',
  'clear_filter', 'clear_print_area', 'create_sheet', 'rename_chart', 'set_chart_axis_title', 'clear_range',
  // X-O04/11 (owner decision "B", internal tracker): marker/gridlines, verified via the X-O04/10 bridge read
  'set_chart_style',
  // L-04 (owner decision 2026-10-14): the column twins of insert_rows / delete_rows
  'insert_columns', 'delete_columns',
  // L-05 (owner decision 2026-10-15): delete_sheet
  'delete_sheet',
  // L-05/2: move_sheet
  'move_sheet',
  // L-07: set_chart, set_chart_legend
  'set_chart', 'set_chart_legend',
  // L-06: copy_range
  'copy_range',
  // X-O07/4b (internal tracker): copy_chart, v1 scope (type/series/title/legend/axis-titles/position only)
  'copy_chart',
  // L-05/3: copy_sheet (the B03 boundary lifted, owner decision)
  'copy_sheet',
  // L-06/b (owner decision, zero 2026-10-15): move_range -- the caller never supplies `expect`
  'move_range',
  // X-O07/3 (internal tracker; X-O07/2 LIVE MEASURED and reviewed): verified via the X-O07/2 DS
  // Bridge /13 WRITE round-trip, same release convention as set_chart_style above.
  'delete_chart',
  // L-11/2 (owner decision "A", the owner, owner channel, the answer to 36637): filter_range, SMALL workbooks only
  'filter_range',
  // GH #138 U1 (owner GO on the #141 roadmap): set_defined_name / delete_defined_name, verified by read_defined_names
  'set_defined_name', 'delete_defined_name',
  // L-18/1 (internal tracker): delete_cells -- block-shift up|left, scoped to the LIVE-measured probe shape
  'delete_cells',
  // GH #138 U2 (same roadmap GO): set_print_titles / clear_print_titles, verified by read_defined_names's `titles` projection
  'set_print_titles', 'clear_print_titles',
  // GH #138 U3 (same roadmap GO): set_page_layout, verified by read_page_setup
  'set_page_layout',
  // GH #139 U2 (same #141 roadmap GO as GH #138; the GH #139 U1 matrix's own C16 gap): move_chart,
  // verified by read_drawings's own pos fields
  'move_chart',
  // GH #139 U3 (same #141 roadmap GO): insert_shape / insert_image, verified by read_drawings
  'insert_shape', 'insert_image',
  // GH #137 U2 (same #141 roadmap GO; READ mirror released in U1): format_run, verified by read_runs
  'format_run',
  // GH #138 U4 (same #141 roadmap GO as GH #138/#139): the remaining 8 print.cjs intents -- fit,
  // header/footer text, manual page breaks, first page number, print centering, scale, paper size,
  // header/footer flags. The write path migrated off the B10 boundary onto the DS Bridge WRITE
  // sheet.printSettings.write/v1 (/7+) in M11.5b.4, verified by read_page_setup's `print` field.
  'set_print_fit', 'set_header_footer', 'set_page_break', 'set_first_page_number', 'set_print_centering',
  'set_print_scale', 'set_paper_size', 'set_header_footer_flag',
  // GH #140 U4 (owner GO on the #141 roadmap): refresh_pivot, verified by read_pivots
  'refresh_pivot',
  // GH #140 U3 (same #141 roadmap GO; docs/GH140-PIVOT-CONTRACT.md section 4): create_pivot -- census
  // ACCEPTED_WRITE, verified by read_pivots (GH #140 U2, already released)
  'create_pivot',
  // GH #301 D4: write_array_formula -- SetFormulaArray over an empty, unbordered area
  'write_array_formula',
  // GH #300 K2: add_comment -- the author is the authenticated session user, verified by read_comments (K1)
  'add_comment',
  // GH #302 G4: set_shape_text -- a shape's paragraphs through GetContent, verified by read_drawings' text
  'set_shape_text',
  // GH #302 G5: set_chart_point_data_label -- SetShowPointDataLabel, verified by read_charts' effectiveLabels
  'set_chart_point_data_label',
  // GH #302 G7: set_shape_rotation -- SetRotation, verified by read_drawings' rotation
  'set_shape_rotation',
  // GH #303 R1: set_sheet_visibility -- SetVisible, verified by read_sheet_state's visibility
  'set_sheet_visibility']);

/** judge({ exposed, capabilityList, rows }) -> list of violations (empty = exact). Pure, so the NEG controls can drive it. */
function judge({ exposed, capabilityList, rows }) {
  const v = [];
  const extra = exposed.filter(t => !EXPECTED_EXPOSED.includes(t));
  const missing = EXPECTED_EXPOSED.filter(t => !exposed.includes(t));
  if (extra.length) v.push(`extra exposed: ${extra.join(',')}`);
  if (missing.length) v.push(`missing exposed: ${missing.join(',')}`);
  if (new Set(exposed).size !== exposed.length) v.push('duplicate exposed tool');
  for (const t of capabilityList) if (!EXPECTED_CAPABILITY_TOOLS.includes(t)) v.push(`capability list names ${t}`);
  for (const r of rows) {
    if (r.READ_WRITE !== 'WRITE') continue;
    const released = r.EXPOSE_IN_PRODUCT_MCP !== 'NO' || capabilityList.includes(r.TOOL);
    if (released && !['xlsx_execute_task', 'docx_execute_task'].includes(r.TOOL)) v.push(`WRITE tool released: ${r.TOOL}`);
  }
  return v;
}

const live = () => ({ exposed: exposedTools(), capabilityList: [...EXPOSED_CAPABILITY_TOOLS], rows: TOOL_ROWS });

test('EXACT_TOOLSET: exposedTools() = product_status, xlsx_read, docx_read, xlsx_execute_task, docx_execute_task -- nothing else', () => {
  assert.deepEqual(exposedTools(), [...EXPECTED_EXPOSED]);
  assert.deepEqual(judge(live()), []);
});

test('DOCX_WRITE_RELEASED_ONLY_AS_PROPOSED: docx_execute_task serves exactly the accepted DOCX intents; every other routed write capability is not exposed', () => {
  assert.deepEqual(TOOL_ROWS.filter(r => !exposedTools().includes(r.TOOL)).map(r => r.TOOL), [...EXPECTED_WITHHELD]);
  assert.deepEqual(capabilityInventory().filter(r => r.EXPOSE_IN_PRODUCT_MCP !== 'NO').map(r => r.TOOL).sort(), ['docx:add_bookmark', 'docx:add_comment_on_match', 'docx:add_endnote', 'docx:add_footnote', 'docx:add_link', 'docx:apply_list', 'docx:define_style', 'docx:delete_paragraph', 'docx:delete_table_column', 'docx:delete_table_row', 'docx:edit_footer', 'docx:edit_header', 'docx:format_table', 'docx:format_text', 'docx:indent_list_item', 'docx:insert_paragraph_after', 'docx:insert_paragraph_before', 'docx:insert_table_column_after', 'docx:insert_table_column_before', 'docx:insert_table_row_after', 'docx:insert_table_row_before', 'docx:merge_table_cells', 'docx:modify_style', 'docx:move_paragraph_after', 'docx:move_paragraph_before', 'docx:outdent_list_item', 'docx:remove_list', 'docx:replace_table_cell', 'docx:reply_comment', 'docx:resolve_comment', 'docx:rewrite_paragraph', 'docx:set_list_level', 'docx:set_list_type', 'docx:set_numbering', 'docx:set_page_setup', 'docx:split_table_cell', 'docx:set_image_size', 'docx:set_image_wrap'].sort(), 'no other routed write capability is released');
  assert.deepEqual(capabilityInventory().filter(r => r.EXPOSE_IN_PRODUCT_MCP !== 'NO').map(r => r.TOOL).sort(), ['docx:add_bookmark', 'docx:add_comment_on_match', 'docx:add_endnote', 'docx:add_footnote', 'docx:add_link', 'docx:apply_list', 'docx:define_style', 'docx:delete_paragraph', 'docx:delete_table_column', 'docx:delete_table_row', 'docx:edit_footer', 'docx:edit_header', 'docx:format_table', 'docx:format_text', 'docx:indent_list_item', 'docx:insert_paragraph_after', 'docx:insert_paragraph_before', 'docx:insert_table_column_after', 'docx:insert_table_column_before', 'docx:insert_table_row_after', 'docx:insert_table_row_before', 'docx:merge_table_cells', 'docx:modify_style', 'docx:move_paragraph_after', 'docx:move_paragraph_before', 'docx:outdent_list_item', 'docx:remove_list', 'docx:replace_table_cell', 'docx:reply_comment', 'docx:resolve_comment', 'docx:rewrite_paragraph', 'docx:set_list_level', 'docx:set_list_type', 'docx:set_numbering', 'docx:set_page_setup', 'docx:split_table_cell', 'docx:set_image_size', 'docx:set_image_wrap'].sort(), 'no other routed write capability is released');
  assert.deepEqual(capabilityInventory().filter(r => r.EXPOSE_IN_PRODUCT_MCP !== 'NO').map(r => r.TOOL).sort(), ['docx:add_bookmark', 'docx:add_comment_on_match', 'docx:add_endnote', 'docx:add_footnote', 'docx:add_link', 'docx:apply_list', 'docx:define_style', 'docx:delete_paragraph', 'docx:delete_table_column', 'docx:delete_table_row', 'docx:edit_footer', 'docx:edit_header', 'docx:format_table', 'docx:format_text', 'docx:indent_list_item', 'docx:insert_paragraph_after', 'docx:insert_paragraph_before', 'docx:insert_table_column_after', 'docx:insert_table_column_before', 'docx:insert_table_row_after', 'docx:insert_table_row_before', 'docx:merge_table_cells', 'docx:modify_style', 'docx:move_paragraph_after', 'docx:move_paragraph_before', 'docx:outdent_list_item', 'docx:remove_list', 'docx:replace_table_cell', 'docx:reply_comment', 'docx:resolve_comment', 'docx:rewrite_paragraph', 'docx:set_list_level', 'docx:set_list_type', 'docx:set_numbering', 'docx:set_page_setup', 'docx:split_table_cell', 'docx:set_image_size', 'docx:set_image_wrap'].sort(), 'no other routed write capability is released');
  assert.deepEqual(capabilityInventory().filter(r => r.EXPOSE_IN_PRODUCT_MCP !== 'NO').map(r => r.TOOL).sort(), ['docx:add_bookmark', 'docx:add_comment_on_match', 'docx:add_endnote', 'docx:add_footnote', 'docx:add_link', 'docx:apply_list', 'docx:define_style', 'docx:delete_paragraph', 'docx:delete_table_column', 'docx:delete_table_row', 'docx:edit_footer', 'docx:edit_header', 'docx:format_table', 'docx:format_text', 'docx:indent_list_item', 'docx:insert_paragraph_after', 'docx:insert_paragraph_before', 'docx:insert_table_column_after', 'docx:insert_table_column_before', 'docx:insert_table_row_after', 'docx:insert_table_row_before', 'docx:merge_table_cells', 'docx:modify_style', 'docx:move_paragraph_after', 'docx:move_paragraph_before', 'docx:outdent_list_item', 'docx:remove_list', 'docx:replace_table_cell', 'docx:reply_comment', 'docx:resolve_comment', 'docx:rewrite_paragraph', 'docx:set_list_level', 'docx:set_list_type', 'docx:set_numbering', 'docx:set_page_setup', 'docx:split_table_cell', 'docx:set_image_size', 'docx:set_image_wrap'].sort(), 'no other routed write capability is released');
  assert.deepEqual(capabilityInventory().filter(r => r.EXPOSE_IN_PRODUCT_MCP !== 'NO').map(r => r.TOOL).sort(), ['docx:add_bookmark', 'docx:add_comment_on_match', 'docx:add_endnote', 'docx:add_footnote', 'docx:add_link', 'docx:apply_list', 'docx:define_style', 'docx:delete_paragraph', 'docx:delete_table_column', 'docx:delete_table_row', 'docx:edit_footer', 'docx:edit_header', 'docx:format_table', 'docx:format_text', 'docx:indent_list_item', 'docx:insert_paragraph_after', 'docx:insert_paragraph_before', 'docx:insert_table_column_after', 'docx:insert_table_column_before', 'docx:insert_table_row_after', 'docx:insert_table_row_before', 'docx:merge_table_cells', 'docx:modify_style', 'docx:move_paragraph_after', 'docx:move_paragraph_before', 'docx:outdent_list_item', 'docx:remove_list', 'docx:replace_table_cell', 'docx:reply_comment', 'docx:resolve_comment', 'docx:rewrite_paragraph', 'docx:set_list_level', 'docx:set_list_type', 'docx:set_numbering', 'docx:set_page_setup', 'docx:split_table_cell', 'docx:set_image_size', 'docx:set_image_wrap'].sort(), 'no other routed write capability is released');
  assert.deepEqual(capabilityInventory().filter(r => r.EXPOSE_IN_PRODUCT_MCP !== 'NO').map(r => r.TOOL).sort(), ['docx:add_bookmark', 'docx:add_comment_on_match', 'docx:add_endnote', 'docx:add_footnote', 'docx:add_link', 'docx:apply_list', 'docx:define_style', 'docx:delete_paragraph', 'docx:delete_table_column', 'docx:delete_table_row', 'docx:edit_footer', 'docx:edit_header', 'docx:format_table', 'docx:format_text', 'docx:indent_list_item', 'docx:insert_paragraph_after', 'docx:insert_paragraph_before', 'docx:insert_table_column_after', 'docx:insert_table_column_before', 'docx:insert_table_row_after', 'docx:insert_table_row_before', 'docx:merge_table_cells', 'docx:modify_style', 'docx:move_paragraph_after', 'docx:move_paragraph_before', 'docx:outdent_list_item', 'docx:remove_list', 'docx:replace_table_cell', 'docx:reply_comment', 'docx:resolve_comment', 'docx:rewrite_paragraph', 'docx:set_list_level', 'docx:set_list_type', 'docx:set_numbering', 'docx:set_page_setup', 'docx:split_table_cell', 'docx:set_image_size', 'docx:set_image_wrap'].sort(), 'no other routed write capability is released');
  assert.deepEqual(capabilityInventory().filter(r => r.EXPOSE_IN_PRODUCT_MCP !== 'NO').map(r => r.TOOL).sort(), ['docx:add_bookmark', 'docx:add_comment_on_match', 'docx:add_endnote', 'docx:add_footnote', 'docx:add_link', 'docx:apply_list', 'docx:define_style', 'docx:delete_paragraph', 'docx:delete_table_column', 'docx:delete_table_row', 'docx:edit_footer', 'docx:edit_header', 'docx:format_table', 'docx:format_text', 'docx:indent_list_item', 'docx:insert_paragraph_after', 'docx:insert_paragraph_before', 'docx:insert_table_column_after', 'docx:insert_table_column_before', 'docx:insert_table_row_after', 'docx:insert_table_row_before', 'docx:merge_table_cells', 'docx:modify_style', 'docx:move_paragraph_after', 'docx:move_paragraph_before', 'docx:outdent_list_item', 'docx:remove_list', 'docx:replace_table_cell', 'docx:reply_comment', 'docx:resolve_comment', 'docx:rewrite_paragraph', 'docx:set_list_level', 'docx:set_list_type', 'docx:set_numbering', 'docx:set_page_setup', 'docx:split_table_cell', 'docx:set_image_size', 'docx:set_image_wrap'].sort(), 'no other routed write capability is released');
  assert.deepEqual([...EXPOSED_CAPABILITY_TOOLS], [...EXPECTED_CAPABILITY_TOOLS]);
  assert.equal(Object.isFrozen(EXPOSED_CAPABILITY_TOOLS), true);
  assert.deepEqual([...W.DOCX_WRITE_ALLOWLIST], [...DOCX_ALLOWLIST]);
  assert.deepEqual([...ACCEPTED_DOCX_WRITE_ALLOWLIST], [...DOCX_ALLOWLIST]);
});

test('WRITE_RELEASE_GATE: M11.4a.1/2/3 CLOSED, each format\'s handler allowlist is EXACTLY its accepted set', () => {
  assert.deepEqual({ ...WRITE_RELEASE_PRECONDITIONS }, { 'M11.4a.1': 'CLOSED', 'M11.4a.2': 'CLOSED', 'M11.4a.3': 'CLOSED' });
  assert.deepEqual([...ACCEPTED_XLSX_WRITE_ALLOWLIST], [...ALLOWLIST]);
  assert.deepEqual([...W.XLSX_WRITE_ALLOWLIST].sort(), [...ALLOWLIST].sort());
  assert.deepEqual([...W.DOCX_WRITE_ALLOWLIST].sort(), [...DOCX_ALLOWLIST].sort());
  assert.deepEqual(writeReleaseGate(), { ok: true, violations: [] });
  assert.deepEqual(writeReleaseGate({ format: 'xlsx' }), { ok: true, violations: [] });
  assert.deepEqual(writeReleaseGate({ format: 'docx' }), { ok: true, violations: [] });
});

test('the served registry is the exact set (what tools/list carries)', () => {
  const config = { schemaVersion: 2, nextcloud: { url: 'https://nc.example' }, broker: { url: 'http://127.0.0.1:9/api/euro-broker' } };
  assert.deepEqual(Object.keys(buildServerTools({ config, configIdentity: 'fixture' }).tools), [...EXPECTED_EXPOSED]);
});

test('NEG controls: an extra tool, a missing tool, a rogue WRITE tool or a write capability turns the gate red', () => {
  const base = live();
  assert.notDeepEqual(judge({ ...base, exposed: [...base.exposed, 'docx_execute_task'] }), []);
  assert.notDeepEqual(judge({ ...base, exposed: base.exposed.filter(t => t !== 'docx_read') }), []);
  assert.notDeepEqual(judge({ ...base, exposed: base.exposed.filter(t => t !== 'xlsx_execute_task') }), []);
  assert.notDeepEqual(judge({ ...base, exposed: base.exposed.filter(t => t !== 'docx_execute_task') }), []);
  assert.notDeepEqual(judge({ ...base, capabilityList: [...base.capabilityList, 'docx_rogue_tool'] }), []);
  assert.notDeepEqual(judge({ ...base, rows: [...base.rows, { TOOL: 'docx_rogue_write', READ_WRITE: 'WRITE', EXPOSE_IN_PRODUCT_MCP: 'YES' }] }), []);
  assert.notDeepEqual(judge({ ...base, exposed: [...base.exposed, 'xlsx:write_range'] }), []);
});

test('NEG controls: the WRITE release gate fails closed PER FORMAT -- an open milestone withholds both WRITE tools; a bad allowlist withholds only its own format\'s tool', () => {
  const DOCX = [...DOCX_ALLOWLIST];
  const ALL_WRITE_OFF = ['product_status', 'xlsx_read', 'docx_read'];
  const badAll = [
    { preconditions: { ...WRITE_RELEASE_PRECONDITIONS, 'M11.4a.3': 'ACTIVE' }, xlsxAllowlist: ALLOWLIST, docxAllowlist: DOCX },
    { preconditions: { 'M11.4a.1': 'CLOSED', 'M11.4a.2': 'CLOSED' }, xlsxAllowlist: ALLOWLIST, docxAllowlist: DOCX },
  ];
  const badXlsx = [
    { xlsxAllowlist: [...ALLOWLIST, 'clear_print_titles'], docxAllowlist: DOCX }, // clear_print_titles: PARTIAL in the census, not yet accepted
    { xlsxAllowlist: ALLOWLIST.slice(1), docxAllowlist: DOCX },
    { xlsxAllowlist: [...ALLOWLIST.slice(0, ALLOWLIST.length - 1), 'insert_rows'], docxAllowlist: DOCX }, // same size, one duplicate, one missing
  ];
  const badDocx = [
    { xlsxAllowlist: ALLOWLIST, docxAllowlist: [] }, // the BEFORE state: the DOCX tool is withheld, the XLSX tool is NOT
    { xlsxAllowlist: ALLOWLIST, docxAllowlist: [...DOCX, 'insert_paragraph_before'] }, // a widened allowlist
    { xlsxAllowlist: ALLOWLIST, docxAllowlist: ['insert_paragraph_before'] }, // a different intent
    { xlsxAllowlist: ALLOWLIST, docxAllowlist: [DOCX[0], DOCX[0]] }, // a duplicate
  ];
  for (const b of [...badAll, ...badXlsx, ...badDocx]) assert.equal(writeReleaseGate(b).ok, false, JSON.stringify(b));
  assert.equal(writeReleaseGate({ xlsxAllowlist: [...ALLOWLIST].reverse(), docxAllowlist: DOCX }).ok, true, 'order does not matter, membership does');
  // and the served surface obeys it
  for (const b of badAll) assert.deepEqual(exposedTools({ writeGate: b }), ALL_WRITE_OFF, JSON.stringify(b));
  for (const b of badXlsx) assert.deepEqual(exposedTools({ writeGate: b }), [...ALL_WRITE_OFF, 'docx_execute_task'], `a bad XLSX allowlist must not withhold the DOCX tool: ${JSON.stringify(b)}`);
  for (const b of badDocx) assert.deepEqual(exposedTools({ writeGate: b }), [...ALL_WRITE_OFF, 'xlsx_execute_task'], `a bad DOCX allowlist must not withhold the XLSX tool: ${JSON.stringify(b)}`);
});

module.exports = { judge, EXPECTED_EXPOSED, EXPECTED_WITHHELD };
