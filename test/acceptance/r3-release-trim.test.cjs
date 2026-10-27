'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const D = require('../../src/server/docx-write-tools.cjs');
const W = require('../../src/server/write-tools.cjs');
const T = require('../../src/server/tools.cjs');

/**
 * Release 3 trim (a maintainer, 2026-10-19): the accepted DOCX WRITE allowlist is the 25 LIVE-PASS intents. The eight others stay routed and coded (the _ALL tables) but are not exposed.
 * GH #268d FLIP (this unit, rebased onto the release-4 line after the comments FLIP ad35d011): set_image_size/set_image_wrap move from
 * WITHHELD into RELEASED_EXTRA, their own LIVE PASS already proven (g268dv1 withheld-gate PASS, g268dv2 handler-acceptance MEASURED).
 * Nothing stays withheld on this line.
 */

const RELEASED_EXTRA = ['insert_table_row_before', 'insert_table_row_after', 'delete_table_row', 'insert_table_column_before', 'insert_table_column_after', 'delete_table_column', 'merge_table_cells', 'split_table_cell', 'format_table', 'set_page_setup', 'edit_header', 'edit_footer', 'define_style', 'modify_style', 'set_numbering', 'set_list_type', 'add_footnote', 'add_endnote', 'add_link', 'add_bookmark', 'resolve_comment', 'reply_comment', 'add_comment_on_match', 'set_image_size', 'set_image_wrap'];
const WITHHELD = [];

test('TRIM: the allowlist is exactly the trimmed set and the four tables (both allowlists, published schemas, constraints) are the same ordered list', () => {
  const a = [...W.DOCX_WRITE_ALLOWLIST];
  assert.equal(a.length, 38);
  for (const i of RELEASED_EXTRA) assert.ok(a.includes(i), i);
  assert.deepEqual([...T.ACCEPTED_DOCX_WRITE_ALLOWLIST], a);
  assert.deepEqual(Object.keys(D.DOCX_STEP_SCHEMAS), a);
  assert.deepEqual(Object.keys(D.DOCX_CONSTRAINTS), a);
});

test('TRIM: no intent stays withheld on this line (the _ALL tables still code everything routed, but WITHHELD above is empty)', () => {
  const routed = D.docxRoutedIntents();
  for (const i of WITHHELD) {
    assert.ok(routed.has(i), `${i} routed`);
    assert.ok(D.DOCX_STEP_SCHEMAS_ALL[i] && D.DOCX_CONSTRAINTS_ALL[i], `${i} still coded`);
    assert.ok(!W.DOCX_WRITE_ALLOWLIST.includes(i) && !T.ACCEPTED_DOCX_WRITE_ALLOWLIST.includes(i), `${i} not allowlisted`);
    assert.equal(D.DOCX_STEP_SCHEMAS[i], undefined, `${i} not published`);
  }
  assert.equal(Object.keys(D.DOCX_STEP_SCHEMAS_ALL).length, 38);
});

test('TRIM (RED/GREEN): every withheld intent is REFUSED_NOT_ALLOWLISTED pre-dispatch, dispatched:false, and no tool/handler path opens a driver', async () => {
  for (const intent of WITHHELD) {
    const v = D.validateDocxWriteRequest({ fileId: '1', task: { steps: [{ intent }] } });
    assert.equal(v.code, 'REFUSED_NOT_ALLOWLISTED', intent);
    assert.equal(v.dispatched, false, intent);
  }
  // POS. CONTROL: a released intent of the same shape is NOT refused for the allowlist
  const ok = D.validateDocxWriteRequest({ fileId: '1', task: { steps: [{ intent: 'set_page_setup', size: { width: 12240, height: 15840 } }] } });
  assert.notEqual(ok.code, 'REFUSED_NOT_ALLOWLISTED');
});

test('TRIM: the exposure gate passes on the trimmed state and the tool inventory exposes exactly the allowlist', () => {
  assert.deepEqual(T.writeReleaseGate(), { ok: true, violations: [] });
  const exposed = T.capabilityInventory().filter(r => r.EXPOSE_IN_PRODUCT_MCP !== 'NO' && r.TOOL.startsWith('docx:')).map(r => r.TOOL).sort();
  assert.deepEqual(exposed, W.DOCX_WRITE_ALLOWLIST.map(i => `docx:${i}`).sort());
});
