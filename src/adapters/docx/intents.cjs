'use strict';

/**
 * The 28/28 DOCX intent migration inventory (M4 brief points 3-5).
 *
 * The canonical list is the SOURCE registry (`docx-agent-capabilities.cjs` AGENT_INTENTS at
 * joskapista44/euro-mcp f8c6d6cf), pinned in test/fixtures/docx-source-registry.json and compared
 * by the accounting gate (test/static/docx-intent-accounting.test.cjs). This file is the TARGET
 * side: every source intent must appear here exactly once, with its family, target module,
 * migration class and status.
 *
 * Families follow the registry's own primitive/action split, not a list written from memory:
 *   replace -> text | format -> formatting | paragraph -> paragraph-structure | list -> list
 *   table   -> table-cell / table-row / table-column / table-merge-split   (by tableAction)
 *   rich_structure -> rich-inline (bookmark, link) / notes (footnote, endnote) /
 *                     header-footer (section-global target, NOT a paragraph selector)
 *
 * MIGRATION_CLASS (only these four; COPY_SOURCE_FILE does not exist):
 *   ADAPT                   the source code path fits the contract with local changes
 *   EXTRACT_SEMANTICS       the Word semantics are re-expressed against the contract; the source
 *                           file is an oracle, not a module to carry over
 *   REWRITE_FROM_CONTRACT   the source path conflicts with the contract (e.g. writes that bypass
 *                           the single dispatch surface, or a weaker verifier than the contract
 *                           requires) and is rebuilt from the contract, using the source as oracle
 *   EO_LIMIT                a measured Euro-Office limitation bounds the capability
 *
 * Evidence references are to the source tree at f8c6d6cf (read-only).
 */

const CLASSES = Object.freeze(['ADAPT', 'EXTRACT_SEMANTICS', 'REWRITE_FROM_CONTRACT', 'EO_LIMIT']);
const STATUS = Object.freeze({ SLICE: 'M4_VERTICAL_SLICE', WAVE1: 'M4_WAVE1', WAVE2: 'M4_WAVE2', WAVE3: 'M4_WAVE3', WAVE4: 'M4_WAVE4', WAVE5: 'M4_WAVE5', NOT_MIGRATED: 'NOT_MIGRATED' });


// Why a family is REWRITE_FROM_CONTRACT, stated once.
const BYPASS = 'source writes via session.frame.evaluate, bypassing callCommand -- the target must route every mutation through the common dispatch capability';

/**
 * PRESERVATION PLACEMENT AUDIT (owner review R1 of the M4 first checkpoint), 28/28, measured on
 * the SOURCE verifier contract, not inferred from intent names. The question per intent: does the
 * source make preservation part of the PRIMITIVE's success (a mismatch fails the primitive, so the
 * source never recorded it APPLIED), or is it absent?
 *
 *   PRIMITIVE_INTEGRAL (P1)  the source primitive returns ok:false on a preservation mismatch
 *   NONE (P0)                the source primitive success checks no preservation invariant
 *   POST_PRIMITIVE (P2)      no intent: nothing in the source checks an INTENT's preservation after
 *                            it was recorded APPLIED. (The source's TASK-level constraints --
 *                            doNotTouchTables, section, preserveFormattingExceptRequested,
 *                            docx-agent-whole-verify.cjs:46-52 -- are post-primitive, but they
 *                            belong to the task, not to any intent; they are not migrated in M4.)
 *
 * Source references are to f8c6d6cf (read-only). The agent path reaches every primitive through
 * docx-session-transaction.cjs applyOne; a primitive ok:false there becomes
 * 'agent-executor-primitive-failed' and no APPLIED receipt entry (docx-agent-executor.cjs:62).
 */
const ROUND_TRIP = 'verifyLiveRoundTrip (unaffected paragraphs/tables) + target text unchanged inside the success branch; else round-trip-safety-failed / target-text-changed';
const TABLE = 'docx-tables.cjs:83 verifyTableObservation: preservation mismatches are in the same list as the primitive checks, ok = no mismatch';
const PRESERVATION_AUDIT = Object.freeze({
  rewrite_paragraph: ['PRIMITIVE_INTEGRAL', 'docx-session-basic-edit.cjs:70-72: seven run properties compared after L1; mismatch -> ok:false rewrite-run-properties-not-preserved; source acceptance test_docx_rewrite_property_preservation_live.cjs requires runPropertiesPreserved===true'],
  format_text: ['PRIMITIVE_INTEGRAL', `docx-session-formatting.cjs:103-108: ${ROUND_TRIP}`],
  insert_paragraph_before: ['NONE', 'docx-session-paragraph-operations.cjs:33: success = observeRelationInSession only'],
  insert_paragraph_after: ['NONE', 'docx-session-paragraph-operations.cjs:33: success = observeRelationInSession only'],
  delete_paragraph: ['NONE', 'docx-session-paragraph-operations.cjs:33: success = observeRelationInSession only'],
  move_paragraph_before: ['NONE', 'docx-session-paragraph-operations.cjs:33: success = observeRelationInSession only'],
  move_paragraph_after: ['NONE', 'docx-session-paragraph-operations.cjs:33: success = observeRelationInSession only'],
  apply_list: ['PRIMITIVE_INTEGRAL', `docx-session-lists.cjs:59-66: ${ROUND_TRIP}`],
  set_list_level: ['PRIMITIVE_INTEGRAL', `docx-session-lists.cjs:59-66: ${ROUND_TRIP}`],
  indent_list_item: ['PRIMITIVE_INTEGRAL', `docx-session-lists.cjs:59-66: ${ROUND_TRIP}`],
  set_list_type: ['PRIMITIVE_INTEGRAL', `docx-session-lists.cjs:59-66: ${ROUND_TRIP}`],
  outdent_list_item: ['PRIMITIVE_INTEGRAL', `docx-session-lists.cjs:59-66 (same path, removal via :49-52): ${ROUND_TRIP}`],
  remove_list: ['PRIMITIVE_INTEGRAL', `docx-session-lists.cjs:59-66 (same path, removal via :49-52): ${ROUND_TRIP}`],
  replace_table_cell: ['PRIMITIVE_INTEGRAL', `${TABLE}; set_cell_text :107, table-shape-preserved :163, unaffected-cell-preserved :185`],
  insert_table_row_before: ['PRIMITIVE_INTEGRAL', `${TABLE}; add_row_before :214, original-cell-preserved :300`],
  insert_table_row_after: ['PRIMITIVE_INTEGRAL', `${TABLE}; add_row_after :317, original-cell-preserved :403`],
  delete_table_row: ['PRIMITIVE_INTEGRAL', `${TABLE}; delete_row :420, original-cell-preserved :491`],
  insert_table_column_before: ['PRIMITIVE_INTEGRAL', `${TABLE}; add_column_before :531, original-cell-preserved :592`],
  insert_table_column_after: ['PRIMITIVE_INTEGRAL', `${TABLE}; add_column_after :609, original-cell-preserved :670`],
  delete_table_column: ['PRIMITIVE_INTEGRAL', `${TABLE}; delete_column :687, original-cell-preserved :750`],
  merge_table_cells: ['PRIMITIVE_INTEGRAL', `${TABLE}; merge_cells :785, unaffected-cell-preserved :855/893/924/947`],
  split_table_cell: ['PRIMITIVE_INTEGRAL', `${TABLE}; split_cell :965, unaffected-cell-preserved :1025/1057/1103/1175, split-content-preserved :1071`],
  add_bookmark: ['NONE', 'docx-rich-structures.cjs:80-84: success = exactly one bookmark with the name'],
  add_link: ['PRIMITIVE_INTEGRAL', 'docx-session-rich-structures.cjs:129: run properties compared after the hyperlink verify; mismatch -> ok:false hyperlink-run-properties-not-preserved'],
  add_footnote: ['NONE', 'docx-rich-structures.cjs:89-93: success = exactly one note with the text'],
  add_endnote: ['NONE', 'docx-rich-structures.cjs:89-93: success = exactly one note with the text'],
  // Owner W5 point 19 (re-audit): the SOURCE placement was P0 (docx-rich-structures.cjs:94-101 checks
  // the part's text only) and stays true of the source. The TARGET is P1 by the owner brief (W5
  // points 11, 21, 22: the body and every other part/section must not change, as part of success).
  // No formatting of the replaced part is claimed: the edit replaces the whole content (probe 2).
  edit_header: ['PRIMITIVE_INTEGRAL', 'source docx-rich-structures.cjs:94-101 (P0: header text + target + variant); target P1 by owner W5 points 11/21/22: body + every other part + section count unchanged'],
  edit_footer: ['PRIMITIVE_INTEGRAL', 'source docx-rich-structures.cjs:94-101 (P0: footer text + target + variant); target P1 by owner W5 points 11/21/22: body + every other part + section count unchanged'],
});

function audited(intent) {
  const a = PRESERVATION_AUDIT[intent];
  if (!a) throw new Error(`no preservation audit row for ${intent}`);
  return { preservationMode: a[0], preservationEvidence: a[1] };
}

const I = (intent, family, cls, fields) => Object.freeze({ intent, family, migrationClass: cls, batchSafe: 'NONE', ...audited(intent), ...fields });

const INTENTS = Object.freeze([
  I('rewrite_paragraph', 'text', 'EXTRACT_SEMANTICS', {
    targetModule: 'src/adapters/docx/families/text.cjs',
    status: STATUS.SLICE,
    enrichment: 'NONE',
    preservation: 'REQUIRED: run properties bold, italic, underline, strike, fontFamily, fontSize, language (source BE:20)',
    refs: 'producer (semantic result = new text); not a consumer',
    source: {
      planner: 'docx-agent-planner.cjs:38', classifier: 'docx-agent-idempotency.cjs:38-72 (W9.2-R)',
      resolver: 'docx-targeting.cjs:resolveStableTarget:118', lowering: 'coedit-docx-target-text.cjs:commandBody:32',
      dispatch: 'docx-session-basic-edit.cjs:54 callCommand', primitiveVerify: 'docx-session-basic-edit.cjs:62-72',
      wholeVerify: 'docx-agent-whole-verify.cjs:68 (present-once only)', preservation: 'docx-session-basic-edit.cjs:20',
      evidence: 'test_docx_agent_rewrite_terminal_live.cjs, test_docx_agent_rewrite_terminal_identity.cjs',
    },
    note: 'target L2 is STRONGER than source: old text absent + terminal context, not present-once only',
  }),
  I('format_text', 'formatting', 'EXTRACT_SEMANTICS', {
    targetModule: 'src/adapters/docx/families/formatting.cjs', status: STATUS.WAVE1, enrichment: 'NONE',
    preservation: 'REQUIRED: target text unchanged + round-trip (other paragraphs, table shapes, images)',
    refs: 'producer; consumer', source: { classifier: 'docx-agent-idempotency.cjs:92-101', lowering: 'coedit-docx-formatting.cjs:43', dispatch: 'docx-session-formatting.cjs:95 callCommand', primitiveVerify: 'docx-session-formatting.cjs:98-108', wholeVerify: 'docx-agent-whole-verify.cjs:69', evidence: 'test_docx_proofing_language_live_acceptance.cjs (no *_terminal_live)' },
    note: 'source classifier: ambiguity falls to APPLY, not fail-closed -- must be fail-closed in the target',
    // W1.1 closed the substring (matchText) target: LIVE PASS 2026-09-24, docs/evidence/M4-W11-*.
  }),
  ...['insert_paragraph_before', 'insert_paragraph_after'].map(intent => I(intent, 'paragraph-structure', 'REWRITE_FROM_CONTRACT', {
    targetModule: 'src/adapters/docx/families/paragraph.cjs', status: STATUS.WAVE1, enrichment: 'NONE',
    preservation: 'NOT DECLARED in source', refs: 'producer; consumer',
    source: { classifier: 'docx-agent-idempotency.cjs:86-90 (position-blind)', lowering: 'coedit-docx-paragraph-operations.cjs:14', dispatch: 'docx-session-paragraph-operations.cjs:29/30 callCommand', primitiveVerify: 'coedit-docx-paragraph-operations.cjs:23 (directional)', wholeVerify: 'docx-agent-whole-verify.cjs:68 (DIRECTION-BLIND)', evidence: 'test_docx_agent_rich_terminal_live.cjs (EURO_RICH_TERMINAL_INTENT)' },
    note: 'whole-task verify must be rebuilt directional (the known direction-blind insert finding)',
  })),
  I('delete_paragraph', 'paragraph-structure', 'REWRITE_FROM_CONTRACT', {
    targetModule: 'src/adapters/docx/families/paragraph.cjs', status: STATUS.WAVE1, enrichment: 'NONE',
    preservation: 'NOT DECLARED in source', refs: 'not a producer',
    source: { classifier: 'docx-agent-idempotency.cjs:83-84', primitiveVerify: 'coedit-docx-paragraph-operations.cjs:23', wholeVerify: 'docx-agent-whole-verify.cjs:67 (no surviving anchor)', evidence: 'test_docx_agent_delete_terminal_live.cjs' },
    note: 'target must add a surviving-anchor check',
  }),
  ...['move_paragraph_before', 'move_paragraph_after'].map(intent => I(intent, 'paragraph-structure', 'EXTRACT_SEMANTICS', {
    targetModule: 'src/adapters/docx/families/paragraph.cjs', status: STATUS.WAVE1, enrichment: 'NONE',
    preservation: 'NOT DECLARED in source', refs: 'producer; destinationRef consumer',
    source: { classifier: 'docx-agent-idempotency.cjs:74-81', primitiveVerify: 'coedit-docx-paragraph-operations.cjs:23', wholeVerify: 'docx-agent-whole-verify.cjs:66 (directional adjacency)', evidence: 'test_docx_agent_move_terminal_live.cjs' },
  })),
  ...['apply_list', 'set_list_level', 'indent_list_item', 'set_list_type'].map(intent => I(intent, 'list', 'EXTRACT_SEMANTICS', {
    targetModule: 'src/adapters/docx/families/list.cjs', status: STATUS.WAVE2, enrichment: 'NONE',
    preservation: 'REQUIRED: round-trip (docx-session-lists.cjs:64)', refs: 'producer',
    source: { classifier: 'docx-agent-idempotency.cjs:103-116', lowering: 'coedit-docx-lists.cjs:7', dispatch: 'docx-session-lists.cjs:47 callCommand', primitiveVerify: 'docx-session-lists.cjs:55-66', wholeVerify: 'docx-agent-whole-verify.cjs:70/80/81', evidence: 'test_docx_agent_list_relative_terminal_live.cjs' },
  })),
  ...['outdent_list_item', 'remove_list'].map(intent => I(intent, 'list', 'REWRITE_FROM_CONTRACT', {
    targetModule: 'src/adapters/docx/families/list.cjs', status: STATUS.WAVE2, enrichment: 'NONE',
    preservation: 'REQUIRED: round-trip', refs: intent === 'remove_list' ? 'not a producer' : 'producer',
    source: { dispatch: 'docx-session-lists.cjs:47 callCommand + :51 frame.evaluate put_ListType', primitiveVerify: 'docx-session-lists.cjs + docx-lists.cjs:99-105', wholeVerify: 'docx-agent-whole-verify.cjs:70', evidence: intent === 'remove_list' ? 'test_docx_agent_remove_list_live_acceptance.cjs' : 'test_docx_agent_list_relative_terminal_live.cjs' },
    note: `${BYPASS}. Target: ONE callCommand; the removal is ONE word /2 Bridge WRITE docx.paragraph.numbering.remove/v1 inside it (M11.5d.4; no public Builder call removes numbering -- measured live 2026-09-24 and 2026-10-11; no direct internal path)`,
  })),
  I('replace_table_cell', 'table-cell', 'REWRITE_FROM_CONTRACT', {
    targetModule: 'src/adapters/docx/families/table.cjs', status: STATUS.WAVE4, enrichment: 'NONE',
    preservation: 'REQUIRED: table shape + every unaffected cell (docx-tables.cjs:163/185)', refs: 'none',
    source: { classifier: 'docx-agent-idempotency.cjs:165-173', dispatch: 'docx-session-tables.cjs:33 frame.evaluate asc_AddText', primitiveVerify: 'docx-tables.cjs:107-200', wholeVerify: 'docx-agent-whole-verify.cjs:75', evidence: 'test_docx_agent_table_cell_terminal_live.cjs' },
    note: `${BYPASS}. Target: ONE callCommand on the cell's own public content API (table probe 2, 2026-09-24); the cell is identified by its full text through the public table read, not by Search`,
  }),
  ...['insert_table_row_before', 'insert_table_row_after', 'delete_table_row'].map(intent => I(intent, 'table-row', 'REWRITE_FROM_CONTRACT', {
    targetModule: 'src/adapters/docx/families/table-row.cjs', status: STATUS.WAVE4, enrichment: 'NONE',
    preservation: 'REQUIRED: original cells preserved', refs: 'none',
    source: { classifier: 'docx-agent-idempotency.cjs:118-125', dispatch: 'docx-session-tables.cjs:34-36 frame.evaluate', primitiveVerify: intent === 'delete_table_row' ? 'docx-tables.cjs:420/518 (target absent)' : 'docx-tables.cjs:214/317', wholeVerify: 'docx-agent-whole-verify.cjs:71', evidence: 'test_docx_agent_table_row_terminal_live.cjs' },
    note: `${BYPASS}. Target: ONE callCommand, public ApiTable.AddRow / RemoveRow on the cell found by table + cell identity (row probes 3-4, 2026-09-24); the RemoveRow return value is not evidence`,
  })),
  ...['insert_table_column_before', 'insert_table_column_after', 'delete_table_column'].map(intent => I(intent, 'table-column', 'REWRITE_FROM_CONTRACT', {
    targetModule: 'src/adapters/docx/families/table-column.cjs', status: STATUS.WAVE4, enrichment: 'NONE',
    preservation: 'REQUIRED: original cells preserved', refs: 'none',
    source: { classifier: 'docx-agent-idempotency.cjs:155-163', dispatch: 'docx-session-tables.cjs:37-39 frame.evaluate', primitiveVerify: 'docx-tables.cjs:531/609/687/772', wholeVerify: 'docx-agent-whole-verify.cjs:74', evidence: 'test_docx_agent_table_column_terminal_live.cjs' },
    note: `${BYPASS}. Target: ONE callCommand, public ApiTable.AddColumn / RemoveColumn on the cell found by table + cell identity; the logical column from the HTML colspan grid; measured-unsafe colspan shapes refused before dispatch (column probe 1, 2026-09-24); the RemoveColumn return value is not evidence`,
  })),
  ...['merge_table_cells', 'split_table_cell'].map(intent => I(intent, 'table-merge-split', 'REWRITE_FROM_CONTRACT', {
    targetModule: 'src/adapters/docx/families/table-merge-split.cjs', status: STATUS.WAVE4, enrichment: 'NONE',
    preservation: 'REQUIRED: unaffected cells; split content preserved', refs: 'none',
    source: { classifier: intent === 'merge_table_cells' ? 'docx-agent-idempotency.cjs:127-138' : 'docx-agent-idempotency.cjs:139-154', dispatch: intent === 'merge_table_cells' ? 'docx-session-tables.cjs:40 frame.evaluate MergeCells' : 'docx-session-tables.cjs:23 splitMutation', primitiveVerify: intent === 'merge_table_cells' ? 'docx-tables.cjs:785' : 'docx-tables.cjs:965', wholeVerify: intent === 'merge_table_cells' ? 'docx-agent-whole-verify.cjs:72 (row count recorded, not asserted)' : 'docx-agent-whole-verify.cjs:73', evidence: 'test_docx_agent_table_merge_split_terminal_live.cjs' },
    note: `${BYPASS}; merge L2 must assert structure (row/column counts, cell identity), not text only. Target: ONE callCommand, public ApiTable.MergeCells / Split on the anchored table's cells; the logical grid (colspan + rowspan) reconciled from the HTML projection and the public table read (merge/split probe 1, 2026-09-24); no internal dependency`,
  })),
  I('add_bookmark', 'rich-inline', 'EXTRACT_SEMANTICS', {
    targetModule: 'src/adapters/docx/families/rich.cjs', status: STATUS.WAVE3,
    // The source took the duplicate guard from the PACKAGE (docx-rich-structures.cjs:64-65). The
    // public live read (GetAllBookmarksNames + GetBookmarkRange, rich probes 1-3, 2026-09-24) is the
    // authority here; the package lags the save and the in-session package URL answered 403.
    enrichment: 'NONE',
    preservation: 'NOT DECLARED in source', refs: 'producer',
    source: { classifier: 'docx-agent-idempotency.cjs:181-187', dispatch: 'docx-session-rich-structures.cjs:117 callCommand', primitiveVerify: 'docx-session-rich-structures.cjs:119-130', wholeVerify: 'docx-agent-whole-verify.cjs:77', evidence: 'test_docx_agent_rich_terminal_live.cjs' },
  }),
  I('add_link', 'rich-inline', 'EXTRACT_SEMANTICS', {
    targetModule: 'src/adapters/docx/families/rich.cjs', status: STATUS.WAVE3, enrichment: 'NONE',
    preservation: 'REQUIRED: run properties before/after (docx-session-rich-structures.cjs:111,129)', refs: 'producer',
    source: { classifier: 'docx-agent-idempotency.cjs:181-187', dispatch: 'docx-session-rich-structures.cjs:115 callCommand', primitiveVerify: 'docx-rich-structures.cjs:85-87', wholeVerify: 'docx-agent-whole-verify.cjs:77', evidence: 'test_docx_agent_rich_terminal_live.cjs' },
  }),
  ...['add_footnote', 'add_endnote'].map(intent => I(intent, 'notes', 'EXTRACT_SEMANTICS', {
    targetModule: 'src/adapters/docx/families/notes.cjs', status: STATUS.WAVE3, enrichment: 'NONE',
    preservation: 'NOT DECLARED in source', refs: 'producer (marker); consumer after the marker',
    source: { classifier: intent === 'add_footnote' ? 'docx-agent-idempotency.cjs:189-200' : 'docx-agent-idempotency.cjs:181-187', dispatch: 'docx-session-rich-structures.cjs:117 callCommand', primitiveVerify: 'docx-session-rich-structures.cjs:80-91', wholeVerify: intent === 'add_footnote' ? 'docx-agent-whole-verify.cjs:78' : 'docx-agent-whole-verify.cjs:77', evidence: 'test_docx_agent_rich_terminal_live.cjs' },
  })),
  ...['edit_header', 'edit_footer'].map(intent => I(intent, 'header-footer', 'REWRITE_FROM_CONTRACT', {
    targetModule: 'src/adapters/docx/families/header-footer.cjs', status: STATUS.WAVE5, enrichment: 'NONE',
    preservation: 'REQUIRED (target): body + other parts + section count', refs: 'none',
    source: { planner: 'docx-agent-planner.cjs:30 (section-global)', classifier: 'docx-agent-idempotency.cjs:175-179', dispatch: 'docx-session-rich-structures.cjs:10 frame.evaluate', primitiveVerify: 'docx-session-rich-structures.cjs:62-74', wholeVerify: 'docx-agent-whole-verify.cjs:76', evidence: 'test_docx_agent_header_footer_terminal_live.cjs' },
    note: `${BYPASS}; section-global target modelled as such (no fake paragraph selector); source accepts only single-section documents -- NOT an Euro-Office limit: the public section API reads and writes every section (header/footer probes 1-2, 2026-09-25). Target: ONE callCommand, public GetSections + GetHeader/GetFooter(v, create) on section + kind + variant; the whole content replaced by one paragraph; default variant only (the title/even flags are not publicly observable); unsafe inheritance refused before dispatch`,
  })),
]);

module.exports = { INTENTS, CLASSES, STATUS, PRESERVATION_AUDIT };
