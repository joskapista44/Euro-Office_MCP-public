'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const snapshot = require('../fixtures/docx-source-registry.json');
const { INTENTS, CLASSES, PRESERVATION_AUDIT } = require('../../src/adapters/docx/intents.cjs');
const { planOperations } = require('../../src/adapters/docx/adapter.cjs');

/**
 * M4 28/28 accounting gate (brief point 4). NOT parity: it proves only that no source intent
 * disappeared and none was invented. The source side is the pinned registry snapshot
 * (scripts/docx-source-registry-snapshot.cjs, joskapista44/euro-mcp f8c6d6cf).
 */

function account(sourceIntents, targetIntents) {
  const src = new Set(sourceIntents);
  const tgt = targetIntents.map(i => i.intent);
  const tgtSet = new Set(tgt);
  return {
    registered: src.size,
    accounted: [...src].filter(i => tgtSet.has(i)).length,
    missing: [...src].filter(i => !tgtSet.has(i)),
    extra: [...tgtSet].filter(i => !src.has(i)),
    duplicates: tgt.filter((i, k) => tgt.indexOf(i) !== k),
  };
}

test('the snapshot is the pinned source registry', () => {
  assert.equal(snapshot.source, 'joskapista44/euro-mcp');
  assert.match(snapshot.commit, /^f8c6d6cf/);
  assert.equal(snapshot.count, snapshot.intents.length);
  assert.match(snapshot.registrySha256, /^[0-9a-f]{64}$/);
});

test('DOCX_REGISTERED_INTENTS=28, DOCX_MIGRATION_ACCOUNTED=28/28, MISSING=[], EXTRA=[]', () => {
  const a = account(snapshot.intents.map(i => i.intent), INTENTS);
  assert.equal(a.registered, 28);
  assert.equal(a.accounted, 28);
  assert.deepEqual(a.missing, []);
  assert.deepEqual(a.extra, []);
  assert.deepEqual(a.duplicates, []);
});

test('every row carries an allowed migration class, a family, a target module and a status', () => {
  for (const row of INTENTS) {
    assert.ok(CLASSES.includes(row.migrationClass), `${row.intent}: ${row.migrationClass}`);
    assert.notEqual(row.migrationClass, 'COPY_SOURCE_FILE');
    assert.ok(row.family && row.targetModule && row.status, row.intent);
    assert.equal(row.batchSafe, 'NONE', `${row.intent}: no batch proof in M4`);
    assert.ok(row.source && row.source.evidence, `${row.intent}: source evidence`);
  }
});

test('the family split follows the registry primitive, not memory', () => {
  const byIntent = new Map(snapshot.intents.map(i => [i.intent, i.primitive]));
  const allowed = {
    replace: ['text'], format: ['formatting'], paragraph: ['paragraph-structure'], list: ['list'],
    table: ['table-cell', 'table-row', 'table-column', 'table-merge-split'],
    rich_structure: ['rich-inline', 'notes', 'header-footer'],
  };
  for (const row of INTENTS) {
    assert.ok(allowed[byIntent.get(row.intent)].includes(row.family), `${row.intent} -> ${row.family}`);
  }
});

test('NEGATIVE CONTROL: the gate detects a dropped, an invented and a duplicated intent', () => {
  const src = snapshot.intents.map(i => i.intent);
  const dropped = account(src, INTENTS.filter(i => i.intent !== 'split_table_cell'));
  assert.deepEqual(dropped.missing, ['split_table_cell']);
  const invented = account(src, [...INTENTS, { intent: 'make_coffee' }]);
  assert.deepEqual(invented.extra, ['make_coffee']);
  const dup = account(src, [...INTENTS, INTENTS[0]]);
  assert.deepEqual(dup.duplicates, ['rewrite_paragraph']);
});

test('DOCX_PRESERVATION_AUDIT=28/28: every intent has a placement and a source reference, no invented P2', () => {
  const modes = ['NONE', 'PRIMITIVE_INTEGRAL', 'POST_PRIMITIVE'];
  for (const row of INTENTS) {
    assert.ok(modes.includes(row.preservationMode), `${row.intent}: ${row.preservationMode}`);
    assert.match(row.preservationEvidence, /\.cjs:\d/, `${row.intent}: the placement cites the source verifier`);
  }
  assert.deepEqual(Object.keys(PRESERVATION_AUDIT).sort(), snapshot.intents.map(i => i.intent).sort(), 'audit rows == registry');
  // The audit found no POST_PRIMITIVE intent; one appearing later needs its own source evidence.
  assert.deepEqual(INTENTS.filter(r => r.preservationMode === 'POST_PRIMITIVE').map(r => r.intent), []);
});

test('every migrated intent plans exactly the audited placement', () => {
  const stepFor = {
    rewrite_paragraph: { selector: { text: 'x' }, text: 'y' },
    format_text: { selector: { text: 'x' }, format: { bold: true } },
    insert_paragraph_before: { selector: { text: 'x' }, text: 'y' },
    insert_paragraph_after: { selector: { text: 'x' }, text: 'y' },
    delete_paragraph: { selector: { text: 'x' } },
    move_paragraph_before: { selector: { text: 'x' }, destinationSelector: { text: 'y' } },
    move_paragraph_after: { selector: { text: 'x' }, destinationSelector: { text: 'y' } },
    apply_list: { selector: { text: 'x' }, type: 'bullet' },
    set_list_level: { selector: { text: 'x' }, level: 1 },
    set_list_type: { selector: { text: 'x' }, type: 'numbered' },
    indent_list_item: { selector: { text: 'x' } },
    outdent_list_item: { selector: { text: 'x' } },
    remove_list: { selector: { text: 'x' } },
    add_bookmark: { selector: { text: 'x' }, name: 'b_1' },
    add_link: { selector: { text: 'x' }, url: 'https://example.invalid/x' },
    add_footnote: { selector: { text: 'x' }, text: 'n' },
    add_endnote: { selector: { text: 'x' }, text: 'n' },
    replace_table_cell: { selector: { text: 'x' }, text: 'y' },
    insert_table_row_before: { selector: { text: 'x' }, terminalRowCount: 4 },
    insert_table_row_after: { selector: { text: 'x' }, terminalRowCount: 4 },
    delete_table_row: { selector: { tableCell: { text: 'x', tableAnchor: 'a' } }, terminalRowCount: 2 },
    insert_table_column_before: { selector: { text: 'x' }, terminalColumnCount: 4 },
    insert_table_column_after: { selector: { text: 'x' }, terminalColumnCount: 4 },
    delete_table_column: { selector: { tableCell: { text: 'x', tableAnchor: 'a' } }, terminalColumnCount: 2 },
    merge_table_cells: { selector: { tableCell: { text: 'x', tableAnchor: 'a' } }, endSelector: { text: 'y' }, terminalMergedText: 'x y', terminalMergedSpan: { rows: 1, cols: 2 } },
    split_table_cell: { selector: { tableCell: { text: 'x', tableAnchor: 'a' } }, splitRows: 1, splitCols: 2 },
    edit_header: { section: { index: 0, of: 1 }, text: 'h' },
    edit_footer: { section: { index: 0, of: 1 }, text: 'f' },
  };
  const migrated = INTENTS.filter(r => r.status !== 'NOT_MIGRATED');
  assert.deepEqual(migrated.map(r => r.intent).sort(), Object.keys(stepFor).sort(), 'the migrated set (28/28 FULL)');
  // Owner accounting: W1.1 closed format_text matchText (LIVE PASS) -> 7/28 FULL, no PARTIAL left.
  assert.deepEqual(migrated.filter(r => r.partial).map(r => r.intent), [], 'no PARTIAL intent');
  // Owner decision after 28596: the list family is FULL (LIVE target 6/6; 3 SOURCE_DEFECT_FIXED) -> 13/28 FULL.
  // Wave 3 (rich-inline + notes): LIVE PASS 2026-09-24 (docs/evidence/M4-W3-RICH-LIVE-ACCEPTANCE.md) -> 17/28 FULL.
  // Wave 4a (table-cell): LIVE target 4/4 2026-09-24; owner decision (owner channel, after 28674):
  // FULL with SOURCE_DEFECT_UNPERFORMED=1 (docs/evidence/M4-W4A-TABLE-CELL-LIVE-ACCEPTANCE.md) -> 18/28 FULL.
  // Wave 4b (table-row): LIVE target 10/10 + regression 9/9 + rewrite PASS (2026-09-24,
  // docs/evidence/M4-W4B-TABLE-ROW-LIVE-ACCEPTANCE.md); source 3 PARITY + 1 SOURCE_DEFECT_UNPERFORMED,
  // which by the owner policy (after 28674, point 11) does not lower FULL -> 21/28 FULL.
  // Wave 4c (table-column): LIVE target 16/16 (2 GUARD refusals as measured) + regression 11/11 + rewrite
  // PASS (2026-09-24, docs/evidence/M4-W4C-TABLE-COLUMN-LIVE-ACCEPTANCE.md); source 3 PARITY + 1
  // SOURCE_DEFECT_UNPERFORMED (col_multi_table), which does not lower FULL -> 24/28 FULL.
  // Wave 4d (table-merge-split): LIVE target 18/18 (3 GUARD refusals as measured) + regression 10/10 +
  // rewrite PASS (2026-09-24, docs/evidence/M4-W4D-TABLE-MERGE-SPLIT-LIVE-ACCEPTANCE.md); source 1 PARITY +
  // 3 SOURCE_DEFECT_UNPERFORMED (merge_multi, split_multi, split_h), which does not lower FULL -> 26/28 FULL.
  // Wave 5 (header-footer): LIVE target 14/14 (4 GUARD refusals as measured) + regression 17/17 + rewrite
  // PASS (2026-09-25, docs/evidence/M4-W5-HEADER-FOOTER-LIVE-ACCEPTANCE.md); source 4 PARITY + 1
  // TARGET_SEMANTICS_STRONGER (hf_multi_header: the source refuses more than one section) -> 28/28 FULL.
  assert.deepEqual(migrated.filter(r => r.pendingLive).map(r => r.intent), [], 'M4_MIGRATED_INTENTS=28/28 FULL, nothing pending LIVE');
  assert.equal(migrated.filter(r => !r.pendingLive).length, 28);
  assert.equal(migrated.length, 28);
  for (const row of migrated) {
    const planned = planOperations({ steps: [{ intent: row.intent, ...stepFor[row.intent] }] });
    assert.equal(planned.ok, true, `${row.intent}: ${planned.detail}`);
    assert.equal(planned.operations[0].preservationMode, row.preservationMode, row.intent);
  }
});
