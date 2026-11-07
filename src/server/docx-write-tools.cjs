'use strict';

/**
 * GH #219 2/2 (DOCX P0, Lane B) -- the product WRITE contract for `docx_execute_task` (owner decisions,
 * docs/GH219-DOCX-WRITE-CONTRACT-DESIGN.md section 9, decisions 2-4). GH #228 (RELEASED, owner GO 2026-10-19,
 * docs/GH228-DOCX-WRITE-RELEASE-GATE-PROPOSAL.md): DOCX_WRITE_ALLOWLIST = [rewrite_paragraph] and the tool is named in
 * tools.cjs EXPOSED_CAPABILITY_TOOLS; before the owner's GO the list is empty and the tool is not registered. This module owns the docx-specific pieces only; every format-neutral piece (schema
 * validator, the recursive security walk, the executor, the plan validator, the receipt projection) is
 * IMPORTED from `write-tools.cjs` / `read-tools.cjs` / `core/execution/*` -- `write-tools.cjs` itself gained
 * only export lines + one additive parameter (`mapExecutionResult`'s `toolName`) for this; its own behaviour
 * on the xlsx path is unchanged (test/acceptance/m11-4a1-write-handler.test.cjs stays green, unmodified).
 *
 * Intent inventory (GH #219 2/2, decision 1 -- the "33 vs 28" question, now measured): `docx/intents.cjs`'s
 * `INTENTS` (the M4 28/28 migration list) plus `docx/capabilities.cjs`'s `NEW_INTENTS` (5 post-M4 additions:
 * define_style, modify_style, set_page_setup, format_table, set_numbering -- the SAME two lists `adapter.cjs`'s
 * own `BY_INTENT` merges) give 28 + 5 = 33, matching the card's "33 adapter semantics". `docxRoutedIntents()`
 * below is the UNION of both lists' intent names -- the full docx-routed set, not just the 28.
 *
 * DOCX_WRITE_ALLOWLIST (write-tools.cjs) was EMPTY in GH #219 (every step was therefore refused REFUSED_NOT_ALLOWLISTED);
 * with the GH #228 proposal it is [rewrite_paragraph] and every OTHER routed docx intent is refused REFUSED_NOT_ALLOWLISTED (a known docx intent, not yet released) rather than
 * REFUSED_CROSS_FORMAT (the xlsx-side classification a docx intent gets today when offered to
 * xlsx_execute_task) -- this is the one behavioural difference this unit adds, and it is the GH #219
 * "unsupported intent" required-test category.
 */

const { docxAdapter } = require('../adapters/docx/adapter.cjs');
const { INTENTS } = require('../adapters/docx/intents.cjs');
const { NEW_INTENTS } = require('../adapters/docx/capabilities.cjs');
const { textRunDiffs, productParagraphOf } = require('../adapters/docx/stored-content.cjs');
// GH #324 follow-up: the schema `pattern` string for every `language` field is DERIVED from this
// ONE regex (families/style-definition.cjs's own guard), the same way write-tools.cjs already
// reuses COMMENT_MAX_TEXT/PIVOT_FUNCTIONS from their owning family -- not a second, hand-kept copy.
const LANGUAGE_PATTERN_SOURCE = require('../adapters/docx/families/style-definition.cjs').LANGUAGE_PATTERN.source;
// GH #314: the border `space` allowlist is DERIVED from formatting.cjs's own write-side table (not
// a second, hand-kept copy) -- the same reuse convention LANGUAGE_PATTERN_SOURCE above already has.
const { BORDER_SPACE_POINTS } = require('../adapters/docx/families/formatting.cjs');
const { validatePlan } = require('../core/execution/plan.cjs');
const { withSession } = require('../core/session/lifecycle.cjs');
const { createExecutor } = require('../core/execution/executor.cjs');
const { createBrowserRuntime } = require('./browser-runtime.cjs');
const { createSessionGate } = require('./session-gate.cjs');
const { productDriver, readOnce, FORMATS, validateReadRequest } = require('./read-tools.cjs');
const W = require('./write-tools.cjs');

const { check, refuse, checkNamesDeep, toExecutorTask, mapExecutionResult, productWriteDriver,
  DOCX_WRITE_ALLOWLIST, MAX_STEPS, FILE_ID, ID, TASK_SCHEMA_KEYS, warmSessionOptions,
  awaitKeyRotation, applyStoredContentVerify } = W;

// ---- intent classification (the docx-side mirror of write-tools.cjs's intentRefusal) -------------------

const docxRoutedIntents = () => new Set([...INTENTS, ...NEW_INTENTS].map(i => i.intent));
const xlsxRoutedIntents = () => new Set(require('./tools.cjs').xlsxRoutedIntents());
const readOperations = () => { const { OPERATIONS } = require('./read-tools.cjs'); return new Set([...Object.keys(OPERATIONS.xlsx), ...Object.keys(OPERATIONS.docx)]); };

/** The class of one step's intent in the DOCX WRITE tool: null = allowlisted, else the refusal. */
function docxIntentRefusal(intent, index) {
  const at = { stepIndex: index, intent };
  if (typeof intent !== 'string' || intent === '') return refuse('MALFORMED_TASK', `step ${index}: intent must be a non-empty string`, { stepIndex: index });
  if (DOCX_WRITE_ALLOWLIST.includes(intent)) return null;
  if (docxRoutedIntents().has(intent)) return refuse('REFUSED_NOT_ALLOWLISTED', `step ${index}: ${intent} is a routed DOCX intent outside the WRITE release allowlist (${DOCX_WRITE_ALLOWLIST.join(', ') || 'empty'})`, at);
  if (xlsxRoutedIntents().has(intent)) return refuse('REFUSED_CROSS_FORMAT', `step ${index}: ${intent} is an XLSX intent (this is the DOCX WRITE tool)`, at);
  if (readOperations().has(intent)) return refuse('REFUSED_READ_OPERATION', `step ${index}: ${intent} is a READ operation (use the READ tool)`, at);
  return refuse('REFUSED_UNKNOWN_INTENT', `step ${index}: unknown intent ${intent}`, at);
}

// ---- schema (one step schema per allowlisted intent; additive, bounded by DOCX_WRITE_ALLOWLIST) -------

// GH #228 (RELEASED, owner GO 2026-10-19): rewrite_paragraph. The selector is the STABLE SEMANTIC one the shared resolver
// reads (resolver.cjs resolveStableTarget / field-contract.cjs SELECTOR_KEYS) -- physical locators (index / position / offset /
// paragraphIndex), bookmark / tableCell and `ref` are NOT in the schema, so they are refused pre-dispatch as UNKNOWN_FIELD
// (exact targeting; the planner would refuse them later with its own named reason). expectedPreservation is the OPTIONAL
// caller contract (families/text.cjs PRESERVED_PROPERTIES), nothing else.
const SEL_TEXT = Object.freeze({ type: 'string', minLength: 1, maxLength: 4096 });
// GH #320: the fingerprint / sectionHeading+paragraphOrdinal selector shapes are shared verbatim by
// rewrite_paragraph, insert_paragraph_before/after and delete_paragraph -- one definition, reused.
const FINGERPRINT = Object.freeze({ type: 'string', minLength: 1, maxLength: 256 });
const PARAGRAPH_ORDINAL = Object.freeze({ type: 'integer', minimum: 1, maximum: 100000 });
const PRECONDITION = Object.freeze({ type: 'object', properties: Object.freeze({ text: SEL_TEXT }), required: ['text'], additionalProperties: false });
// GH #307: a document reduced to its single, empty first paragraph has no text to name it by (every
// branch above keys off a TEXT); `emptyDocument` is a boolean sentinel, not a text anchor. ONLY
// `true` is accepted (`false` would just be a confusing way of saying "nothing"). Schema-shared with
// format_text (REWRITE_SELECTOR below), but scoped to rewrite_paragraph ONLY by field-contract.cjs
// SELECTOR_KEYS -- a format_text step naming it is refused there, named, pre-dispatch, 0 writes.
const EMPTY_DOCUMENT = Object.freeze({ const: true });
const REWRITE_SELECTOR = Object.freeze({
  type: 'object',
  properties: Object.freeze({
    fingerprint: FINGERPRINT,
    heading: SEL_TEXT, level: Object.freeze({ type: 'integer', minimum: 1, maximum: 9 }),
    text: SEL_TEXT, contains: Object.freeze({ type: 'boolean' }), kind: Object.freeze({ enum: ['paragraph', 'heading'] }),
    headingPath: Object.freeze({ type: 'array', minItems: 1, maxItems: 9, items: Object.freeze({ type: 'object',
      properties: Object.freeze({ level: Object.freeze({ type: 'integer', minimum: 1, maximum: 9 }), text: SEL_TEXT }), required: ['level', 'text'], additionalProperties: false }) }),
    before: SEL_TEXT, after: SEL_TEXT,
    sectionHeading: SEL_TEXT, paragraphOrdinal: PARAGRAPH_ORDINAL,
    emptyDocument: EMPTY_DOCUMENT,
    precondition: PRECONDITION,
  }),
  additionalProperties: false,
});
// GH #233: insert_paragraph_before / _after. The anchor selector is NARROWER than rewrite's, each omission measured (docs/GH233-PARAGRAPH-
// INSERT-PRODUCT-DECISION.md section 2): a heading / headingPath / contains anchor APPLIES and then reports ok:false -- a failing receipt
// over a persisted change, so they stay out; `before` / `after` would be silently dropped by the insert resolver (the duplicate stays
// ambiguous); `ref` needs dependsOn.
// GH #320: `fingerprint` and `sectionHeading` + `paragraphOrdinal` are released here -- the SAME two extra branches rewrite_paragraph
// already has, alongside the existing `text` anchor (paragraphSelectorBranch below requires exactly one of the three). Both were
// ALREADY migrated at the field-contract layer (field-contract.cjs SELECTOR_KEYS) and already resolve through the shared resolver
// (resolver.cjs) and paragraph.cjs's own ordinalTerminal -- this is schema widening onto an existing, reused path, not a new resolver.
const INSERT_SELECTOR = Object.freeze({
  type: 'object',
  properties: Object.freeze({
    text: SEL_TEXT, kind: Object.freeze({ enum: ['paragraph', 'heading'] }),
    fingerprint: FINGERPRINT, sectionHeading: SEL_TEXT, paragraphOrdinal: PARAGRAPH_ORDINAL,
    precondition: PRECONDITION,
  }),
  additionalProperties: false,
});
const insertStepSchema = intent => Object.freeze({
  type: 'object',
  properties: Object.freeze({
    intent: Object.freeze({ const: intent }), id: ID, selector: INSERT_SELECTOR,
    text: Object.freeze({ type: 'string', minLength: 1, maxLength: 32767 }),
  }),
  required: Object.freeze(['intent', 'selector', 'text']),
  additionalProperties: false,
});
// GH #234: delete_paragraph. A deletion is proven by what SURVIVES it, so the TEXT-anchored form names the target by its exact text
// AND both neighbours: `before` / `after`, at most 80 characters (a PREFIX of the neighbour's text -- the adapter compares a
// neighbour's first 80 characters; a longer value finds nothing: fail-closed, 0 writes, measured). Without both, an already-deleted
// target cannot be told from a typo (the adapter answers FAILED delete-terminal-identity-insufficient-selector, so a lost reply could
// never be retried). GH #321: an empty `before` / `after` ('' = the document edge) names the first / last paragraph; the adapter
// proves the edge BEFORE dispatch (families/paragraph.cjs classifyDelete), so a wrong '' is refused with 0 writes -- and both empty
// (the document's only paragraph) stays refused.
// GH #320: `fingerprint` and `sectionHeading` + `paragraphOrdinal` are released as ALTERNATIVE branches to `text` -- NOT alongside it:
// `before` / `after` narrow the resolver's OWN text-branch matching (resolver.cjs) and are otherwise ignored by it, so field-contract.cjs
// refuses them beside a fingerprint/ordinal selector (the SAME exclusivity rewrite_paragraph already has -- "text would be ignored beside
// the fingerprint selector"); they are therefore required ONLY on the text branch (paragraphConstraintProblem below), never on the other
// two. `fingerprint` resolves and retries exactly like `text` (NOT_FOUND on a retry falls back to field-contract's own surviving-context
// proof when before/after are given on the text branch; a fingerprint-only retry is never SATISFIED -- safe, not idempotent, documented in
// paragraph.cjs). `sectionHeading` + `paragraphOrdinal` is accepted at THIS layer and at field-contract.cjs, and is ALWAYS refused one
// layer further in, by adapter.cjs's own ordinalAfterStructuralStep (named `ordinal-selector-on-count-changing-step`): a delete changes
// how many paragraphs its section holds, so an ordinal would not name the same paragraph on a retry -- the SAME reason insert_paragraph_*
// is the ONE count-changing intent exempted from that guard, and delete_paragraph deliberately is not. Widening the schema here still
// matters: the caller now gets that one, named, correct-layer reason instead of a pre-dispatch UNKNOWN_FIELD that looks like a typo.
const NEIGHBOUR_TEXT = Object.freeze({ type: 'string', minLength: 0, maxLength: 80 });
const DELETE_SELECTOR = Object.freeze({
  type: 'object',
  properties: Object.freeze({
    text: SEL_TEXT, before: NEIGHBOUR_TEXT, after: NEIGHBOUR_TEXT, kind: Object.freeze({ enum: ['paragraph', 'heading'] }),
    fingerprint: FINGERPRINT, sectionHeading: SEL_TEXT, paragraphOrdinal: PARAGRAPH_ORDINAL,
    precondition: PRECONDITION,
  }),
  additionalProperties: false,
});
// GH #235: move_paragraph_before / _after. The moved paragraph is selected like the PRE-#320 insert anchor
// (text required, kind, precondition -- the same L2 defect for a heading / fingerprint / contains selector,
// measured on BOTH selectors of a move) and the destination is { text (required), kind? } (the contract has
// no precondition on a destination; before / after / headingPath are not honoured by the move resolver
// either). Kept as its OWN schema object, separate from insert_paragraph_*'s INSERT_SELECTOR: GH #320 is
// scoped to insert_paragraph_before/after and delete_paragraph only -- move was never asked for a fingerprint
// / sectionHeading+paragraphOrdinal source selector, and adapter.cjs's ordinalAfterStructuralStep refuses an
// ordinal on move_paragraph_* anyway (it is COUNT_CHANGING, not in ORDINAL_INSERT), so widening would only
// ever reach the SAME "refused, named" outcome delete_paragraph's ordinal branch does -- not worth sharing a
// mutable schema object across two intents outside this release's own stated scope.
const MOVE_SOURCE_SELECTOR = Object.freeze({
  type: 'object',
  properties: Object.freeze({
    text: SEL_TEXT, kind: Object.freeze({ enum: ['paragraph', 'heading'] }),
    precondition: PRECONDITION,
  }),
  required: Object.freeze(['text']),
  additionalProperties: false,
});
const MOVE_DESTINATION = Object.freeze({
  type: 'object',
  properties: Object.freeze({ text: SEL_TEXT, kind: Object.freeze({ enum: ['paragraph', 'heading'] }) }),
  required: Object.freeze(['text']),
  additionalProperties: false,
});
const moveStepSchema = intent => Object.freeze({
  type: 'object',
  properties: Object.freeze({ intent: Object.freeze({ const: intent }), id: ID, selector: MOVE_SOURCE_SELECTOR, destinationSelector: MOVE_DESTINATION }),
  required: Object.freeze(['intent', 'selector', 'destinationSelector']),
  additionalProperties: false,
});
// GH #236: apply_list / set_list_level / set_list_type. The list family's own plan() accepts a selector of EXACTLY { text } (or { ref }), so the
// target is that and nothing else; type is bullet | numbered, level 0-8 (apply_list defaults to 0). These are the shapes the Wave 2 LIVE
// acceptance measured (docs/evidence/M4-W2-LIST-LIVE-ACCEPTANCE.md) -- a numbering's type is PER LEVEL and shared by every item on that
// numbering level. set_list_type was WITHHELD from this release: when it re-templates a level another item sits on, the family ends
// DISPATCHED_OUTCOME_UNKNOWN and the DS persists the unsaved edit on its own (LIVE g236v1 ROUND6 + guncprobe, 2026-10-19). GH #236b releases it WITH a pre-dispatch guard:
// the product always sends sharedDefinition: 'refuse' (toAdapterStep; a caller-supplied sharedDefinition is UNKNOWN_FIELD), so a level other items sit on fails closed BEFORE any dispatch.
const LIST_TARGET = Object.freeze({ type: 'object', properties: Object.freeze({ text: SEL_TEXT }), required: Object.freeze(['text']), additionalProperties: false });
const LIST_TYPE = Object.freeze({ enum: ['bullet', 'numbered'] });
const LIST_LEVEL = Object.freeze({ type: 'integer', minimum: 0, maximum: 8 });
const listStepSchema = (intent, properties, required) => Object.freeze({
  type: 'object',
  properties: Object.freeze({ intent: Object.freeze({ const: intent }), id: ID, selector: LIST_TARGET, ...properties }),
  required: Object.freeze(['intent', 'selector', ...required]),
  additionalProperties: false,
});
// GH #237: indent_list_item / outdent_list_item / remove_list. Same { text } target as the other list intents, no other fields. A RELATIVE step
// (indent / outdent) has a provable terminal level only after an ABSOLUTE step for the same target earlier in the SAME task (list.cjs header:
// the source silently assumed level 0; here a standalone relative step is FAILED list-relative-level-unprovable) -- so the product releases
// them only as the second step of [apply_list | set_list_level, indent | outdent] on one text (the Wave 2 LIVE shapes: apply + indent, apply + outdent).
// remove_list is the Bridge-backed removal (docx.paragraph.numbering.remove/v1, word /3 runtime only, no direct-internal fallback).
const listRelativeSchema = intent => listStepSchema(intent, {}, []);
// GH #249: replace_table_cell. The cell is named by TABLE identity + CELL identity (table.cjs cellSelector, owner W4a closure point 2): the cell's
// FULL normalized text (what read_tables returns as `text`; a merged multi-paragraph cell reads its paragraphs joined by a space), optionally with
// `tableAnchor` = the full text of ANOTHER cell that names exactly one table. Exactly that one canonical form is released: the bare { text } form
// (the rewrite_paragraph shape), the combined { tableCell, text } form, kind / precondition / heading / fingerprint / ref and physical row / column
// hints are NOT in the schema (UNKNOWN_FIELD / REFUSED_SHAPE pre-dispatch; the adapter refuses them as well -- measured, docs/GH249-...).
const CELL_TARGET = Object.freeze({
  type: 'object',
  properties: Object.freeze({
    tableCell: Object.freeze({
      type: 'object',
      properties: Object.freeze({ text: SEL_TEXT, tableAnchor: SEL_TEXT }),
      required: Object.freeze(['text']),
      additionalProperties: false,
    }),
  }),
  required: Object.freeze(['tableCell']),
  additionalProperties: false,
});
// GH #250: insert_table_row_before / insert_table_row_after. The target ROW is named by one of its cells with the #249 identity (CELL_TARGET: the full normalized cell
// text, optionally with tableAnchor = another cell that names exactly one table) -- the row the cell sits in. `terminalRowCount` (the target table's row count AFTER this
// operation, i.e. the read_tables count + 1) is REQUIRED: it is the retry proof (the adapter's SATISFIED needs the empty row on the declared side AND that count). The source's
// `terminalTableAnchor` alias, the bare { text } form and a position are not in the schema (one way to say each thing).
const ROW_COUNT = Object.freeze({ type: 'integer', minimum: 2, maximum: 1000 });
// GH #291 T-05 (owner decision): `extendMerge: true` is the EXPLICIT caller mode for an insert whose new row lands inside a
// vertical merge -- that merge grows by one row (LIVE, docs/evidence/gh291t5-row-insert-merge-probe-20261022.json). Without it the
// insert keeps refusing (the default does not change); with it an insert that cuts no merge is refused too (the flag would mean nothing).
const rowInsertSchema = intent => Object.freeze({
  type: 'object',
  properties: Object.freeze({ intent: Object.freeze({ const: intent }), id: ID, selector: CELL_TARGET, terminalRowCount: ROW_COUNT, extendMerge: Object.freeze({ const: true }) }),
  required: Object.freeze(['intent', 'selector', 'terminalRowCount']),
  additionalProperties: false,
});
// GH #251: delete_table_row. The row is named like in #250 (a cell of the row, the #249 identity) but the surviving TABLE ANCHOR is REQUIRED: the delete's proof is "the target is absent from the anchor table
// and the row count is terminal" (table-row.cjs plan(): a delete without an anchor is refused), so the anchor must be a cell of a row that SURVIVES (an anchor in the target row is table-anchor-in-target-row).
// terminalRowCount = the table's row count AFTER the delete (read_tables count - 1), at least 1. The source's `terminalTableAnchor` alias is not in the schema.
const ANCHORED_CELL_TARGET = Object.freeze({
  type: 'object',
  properties: Object.freeze({
    tableCell: Object.freeze({
      type: 'object',
      properties: Object.freeze({ text: SEL_TEXT, tableAnchor: SEL_TEXT }),
      required: Object.freeze(['text', 'tableAnchor']),
      additionalProperties: false,
    }),
  }),
  required: Object.freeze(['tableCell']),
  additionalProperties: false,
});
const ROW_COUNT_AFTER_DELETE = Object.freeze({ type: 'integer', minimum: 1, maximum: 1000 });
// GH #252: insert_table_column_before / insert_table_column_after. The target COLUMN is the logical column of the named cell (the #249 identity, CELL_TARGET): the cell must sit in exactly ONE logical
// column; a cell that spans several columns (also several rows, GH #291 H3 / H3v) names the span's own edge: after = its right, before = its left. `terminalColumnCount` = the table's LOGICAL width after the insert (>= 2) -- the retry proof; the logical width is the
// cell count of any row WITHOUT a colspan (read_tables exposes no colspan; a short row reads as a trailing grid gap, GH #291 H4 / H5 -- other unequal widths are fail-closed uneven-logical-width). The source's `terminalTableAnchor` alias, the bare
// { text } form and a position are not in the schema.
const COLUMN_COUNT = Object.freeze({ type: 'integer', minimum: 2, maximum: 1000 });
// GH #253: delete_table_column. The column is named like in #252 (a cell of the column, the #249 identity) but the surviving TABLE ANCHOR is REQUIRED (ANCHORED_CELL_TARGET, as #251): the delete's proof is
// "the target is absent from the anchor table and the logical width is terminal" (table-column.cjs plan(): a delete without an anchor is refused), so the anchor must be a cell OUTSIDE the deleted column
// (an anchor in the target column is table-anchor-in-target-column). terminalColumnCount = the table's LOGICAL width AFTER the delete (read_tables width - 1), at least 1.
const COLUMN_COUNT_AFTER_DELETE = Object.freeze({ type: 'integer', minimum: 1, maximum: 1000 });
// GH #254: merge_table_cells. `selector` is the START corner (ANCHORED_CELL_TARGET: the #249 identity with the surviving TABLE ANCHOR REQUIRED -- a cell OUTSIDE the region, table-anchor-in-merge-region otherwise);
// `endSelector` is the region's other corner and carries NO anchor of its own (the table is the start corner's: the adapter's cross-table endSelector branch is not released). The region is the logical bounding
// rectangle of the two corners. `terminalMergedText` = the region's paragraphs in row-major order joined by a space (what read_tables returns for the merged cell) and `terminalMergedSpan` = the region's logical
// size { rows, cols } (rows * cols >= 2): the adapter derives both from the fresh read and fails closed (merge-terminal-contract-mismatch) when they differ; they are also the retry proof.
const END_CELL_TARGET = Object.freeze({
  type: 'object',
  properties: Object.freeze({
    tableCell: Object.freeze({
      type: 'object',
      properties: Object.freeze({ text: SEL_TEXT }),
      required: Object.freeze(['text']),
      additionalProperties: false,
    }),
  }),
  required: Object.freeze(['tableCell']),
  additionalProperties: false,
});
// GH #255: split_table_cell. The target is named like in #253/#254 (the #249 identity, the surviving TABLE ANCHOR REQUIRED: a cell OTHER than the target). splitRows / splitCols are the adapter's MEASURED
// geometry sets (table-merge-split.cjs SPLIT_ROWS_MEASURED [1..4] / SPLIT_COLS_MEASURED [1..6], W05 + GH #291 S2 / S2x): the published schema says what is released, and any other value is REFUSED_SHAPE pre-dispatch
// (the acceptance pins these maxima to the adapter's constants). A RESTORE must have rows x cols == the merged owner's span (GH #291 S3 / S3x: a merged owner may also split into fewer equal pieces of a measured shape); a 1 x 1 split is degenerate (the planner refuses it). The width of the new cells is out of the contract (W05-2).
// GH #258: format_table -- the `style` field ONLY. The table is named by ONE of its cells under the adapter's own `table` key (table-format.cjs STEP_KEYS), in the one released #249 form (CELL_TARGET:
// tableCell { text, tableAnchor? }; the cell itself is never touched). `style` is a table style NAME that must exist in the document's own table-style list (exact match, resolved by the adapter; unknown /
// ambiguous / unreadable -> fail closed, no style is created). borders / shading / cellMargins / indent / width are not in the schema (UNKNOWN_FIELD): the adapter supports them, the product does not release them here.
const SPLIT_ROWS = Object.freeze({ type: 'integer', minimum: 1, maximum: 4 });
const SPLIT_COLS = Object.freeze({ type: 'integer', minimum: 1, maximum: 6 });
const MERGED_SPAN = Object.freeze({
  type: 'object',
  properties: Object.freeze({ rows: Object.freeze({ type: 'integer', minimum: 1, maximum: 1000 }), cols: Object.freeze({ type: 'integer', minimum: 1, maximum: 1000 }) }),
  required: Object.freeze(['rows', 'cols']),
  additionalProperties: false,
});
const columnInsertSchema = intent => Object.freeze({
  type: 'object',
  properties: Object.freeze({ intent: Object.freeze({ const: intent }), id: ID, selector: CELL_TARGET, terminalColumnCount: COLUMN_COUNT }),
  required: Object.freeze(['intent', 'selector', 'terminalColumnCount']),
  additionalProperties: false,
});
const PRESERVED_VALUE = Object.freeze({ anyOf: [Object.freeze({ type: 'boolean' }), Object.freeze({ type: 'number' }), Object.freeze({ type: 'string', maxLength: 256 })] });
// GH #229 (format_text core run-formatting, PREPARATION, owner decision pending alongside #228 --
// docs/GH229-FORMAT-TEXT-PRODUCT-DECISION.md): the selector is the SAME stable-semantic shape as
// rewrite_paragraph (REWRITE_SELECTOR, reused verbatim) -- whole-paragraph targeting only. `format`
// accepts EXACTLY the seven run properties families/text.cjs's rewrite_paragraph already preserves
// (PRESERVED_PROPERTIES), plus (GH #230) color/highlight/vertAlign, plus (GH #231, below)
// alignment/spaceBefore/spaceAfter/indentLeft/indentRight/indentFirstLine/shading, plus (GH #232, below)
// tabs/borderTop/borderBottom/borderLeft/borderRight/keepNext/keepLines/pageBreakBefore/widowControl -- `style`
// (and lineSpacing, see #231) remain UNKNOWN_FIELD, a deliberate product-schema narrowing: the adapter itself
// can do more than this release allows.
// fontSize is in POINTS (test/fakes/fake-docx-driver.cjs's own pointSize()/*2 convention, matching
// the Builder's SetFontSize(points) call in formatting.cjs's commandBody); the upper bound (1638pt)
// mirrors the OOXML half-point ceiling (3276 half-points = 1638pt), not independently measured.
//
// GH #229 second slice (PREPARATION, owner decision pending alongside #228/#229): `matchText`
// (the adapter's OPTIONAL substring/range targeting -- families/formatting.cjs's `isRange` branch,
// docs/GH229-FORMAT-TEXT-PRODUCT-DECISION.md section 12) is now IN the schema -- a non-empty string
// naming the literal substring inside the resolved parent paragraph. The ambiguity/preservation
// contract is NOT re-implemented here: it is the adapter's own `plan()`/`rangeState`/`rangeTerminal`/
// `parentAddressable` (families/formatting.cjs:260-373), reused unchanged. Pre-dispatch this adds only
// a non-empty-string check (`plan()`'s own, plus this schema's `minLength:1`); the substring's
// uniqueness WITHIN the parent and the parent's own uniqueness in the WHOLE document are EXECUTION-time
// fail-closed classifications (`range-ambiguous-in-paragraph` / `format-parent-not-addressable`), the
// same way a `rewrite_paragraph`/whole-paragraph `format_text` selector's ambiguity is -- never a
// pre-dispatch REFUSED_SHAPE.
//
// GH #230 (PREPARATION, owner decision pending alongside #228/#229): color/highlight/vertAlign --
// EXACTLY the MEASURED W-08 value shapes (families/formatting.cjs's own `normalizeW08`/`HEX_RE`/
// `HIGHLIGHT_VALUES`/`VERT_ALIGN_VALUES`, :166-184; the write calls `range.SetColor(Api.HexColor(...))`/
// `SetHighlight`/`SetVertAlign`, :535-537), already LIVE-proven through this SAME product intent
// (docs/evidence/w08-color-highlight-vertalign-acc3-PASS-20261006.json, before this unit existed --
// this is productization of an already-working adapter capability, not new adapter behaviour).
// `color` accepts ONLY `#RRGGBB` -- no 'auto'/reset value is accepted, because no SetColor call shape
// for resetting to auto was ever measured (the header comment's `Api.AutoColor` constructor exists
// but was not the one the W-08 probe picked as working); READ can still observe 'auto' on a run that
// was never explicitly coloured (src/drivers/eurooffice/docx-driver.cjs's `colorHex`, :186-195:
// `null` -> 'auto', `IsAutoColor()` -> 'auto', else '#RRGGBB') -- this is an asymmetry, not an
// oversight: see docs/GH230-FORMAT-TEXT-COLOR-HIGHLIGHT-VERTALIGN-DECISION.md section 2.
//
// GH #231 (PREPARATION, owner decision pending alongside #228/#229/#230): the paragraph-level keys,
// shared by BOTH the whole-paragraph and the matchText/range write path (the SAME `paraPr` block,
// formatting.cjs:538-549 / :618-629) -- a paragraph-level key is NEVER refused for being combined
// with `matchText`; it writes the PARENT paragraph regardless, exactly as docs/W09-TABS-BORDERS-
// PAGINATION-DESIGN-20261007.md section 1 already documents for this family ("a PARAGRAFUS-szintu
// mezok a RANGE celzas eseten IS a SZULO bekezdesre irodnak"). `indentLeft`/`indentRight`/
// `indentFirstLine`/`shading` are LIVE-PROVEN through this SAME product intent
// (docs/evidence/w09-indent-shading-acc3-PASS-20261006.json, including the hanging-indent case --
// a NEGATIVE indentFirstLine, -12pt, measured APPLIED). `alignment`/`spaceBefore`/`spaceAfter` have
// the SAME simple points<->twips write/read convention (docx-driver.cjs:327-333) but NO equivalent
// LIVE evidence file through this product path -- shipped here as code-measured, flagged in
// docs/GH231-FORMAT-TEXT-PARAGRAPH-PROPS-DECISION.md section 4 for this unit's OWN LIVE harness to
// be the first to verify. `lineSpacing` is DELIBERATELY EXCLUDED: its READ conversion
// (docx-driver.cjs:334, `n>10 ? (n>=200 ? n/240 : (n+1)/100) : n`) is NOT a clean inverse of the
// WRITE's `*240` (formatting.cjs:543/623) for any value below 200/240≈0.833 -- a code-verified
// round-trip defect (write 0.5 -> read back 1.21), not merely unmeasured; see decision doc section 5.
// `shading` accepts ONLY a BARE 6-hex-digit string, NO leading '#' -- UNLIKE `color`/`FORMAT_TEXT_VALUE`'s
// own `#RRGGBB` convention: `normalizeShading` stores `hex.toUpperCase()` without a '#'
// (families/formatting.cjs:112-118) and the REAL driver's read side returns the SAME bare form
// (`h.toUpperCase()`, no '#', docx-driver.cjs:352-354) -- this schema matches the MEASURED shape so
// the written value round-trips character-for-character through read_formatting, not a cosmetic
// `#`-prefixed contract that would silently differ from what comes back. Same 'auto' asymmetry as
// `color`: no 'auto' write value (`SHADING_HEX_RE` is hex-only); READ can still observe 'auto'
// (docx-driver.cjs:348-350).
//
// GH #232 (PREPARATION, owner decision pending alongside #228..#231 -- docs/GH232-FORMAT-TEXT-TABS-BORDERS-PAGINATION-DECISION.md):
// tabs / per-side borders / pagination flags, the MEASURED W-09 axes (docs/W09-TABS-BORDERS-PAGINATION-DESIGN-20261007.md,
// docs/evidence/w09-3-live-acceptance-PASS-20261007.json), through the SAME paragraph-level `paraPr` block as #231. Every
// limit the adapter itself enforces is mirrored here so the refusal is PRE-DISPATCH (families/formatting.cjs:120-145,
// :218-226):
//   tabs        array of { pos (points, >= 0), align 'left' | 'right' | 'center' (GH #290) } -- NO leader, no decimal; [] removes
//               the paragraph's own tabs (GH #290 F-02, a style's tabs untouched)
//               (measured: decimal is silently stored as left, a leader argument is ignored).
//               The write REPLACES the whole list (W09-1e); the READ returns { pos, align } in the same points.
//   border*     { style 'single', size (points), color '#RRGGBB' } OR { style 'none' } (GH #290: removes the side) per side,
//               four independent sides (W09-1e); other styles refused (measured: silently stored as none). ASYMMETRY:
//               the adapter accepts the colour WITH the '#' (HEX_RE) but the driver READ returns it BARE ('FF0000', or
//               'auto') -- the OPPOSITE of `shading`; a caller comparing strings must strip the '#' (decision doc section 4).
//   pagination  keepNext/keepLines/pageBreakBefore/widowControl true or false (GH #290: false writes an explicit override, proven
//               by the read's paginationDirect; widowControl false measured 2026-10-23).
// The bounds on pos (<= 1000pt, the #231 indent bound), size (0.125..12pt) and the 32-entry tab cap are SCHEMA GUARDS, not
// measurements; what IS derived from the measured quantisation is the GRID check in DOCX_CONSTRAINTS.format_text below.
const TAB_STOP = Object.freeze({
  type: 'object',
  properties: Object.freeze({ pos: Object.freeze({ type: 'number', minimum: 0, maximum: 1000 }), align: Object.freeze({ enum: ['left', 'right', 'center'] }) }),
  required: Object.freeze(['pos', 'align']),
  additionalProperties: false,
});
// GH #290: a side is EITHER { style 'single', size, color } OR { style 'none' } (removes it: an explicit w:val="none").
const BORDER_SIDE = Object.freeze({ anyOf: Object.freeze([
  Object.freeze({
    type: 'object',
    properties: Object.freeze({
      style: Object.freeze({ const: 'single' }),
      size: Object.freeze({ type: 'number', minimum: 0.125, maximum: 12 }),
      color: Object.freeze({ type: 'string', pattern: '^#[0-9a-fA-F]{6}$' }),
      // GH #314 (LIVE-measured): optional, points -- deliberately NO
      // enum/range HERE -- measured: a bound at this schema layer is caught FIRST, with a generic
      // "unsupported value" that never reaches formatGridProblem's own, enumerated message below.
      // formatGridProblem is the SOLE authority on which numbers are acceptable (0..31 would be
      // just as wrong a bound as an enum: it would steal -1/32 from the same enumerated message).
      // The SAME split size/tabs already use (GH #232): a loose schema, a precise constraint fn.
      space: Object.freeze({ type: 'number' }),
    }),
    required: Object.freeze(['style', 'size', 'color']),
    additionalProperties: false,
  }),
  Object.freeze({ type: 'object', properties: Object.freeze({ style: Object.freeze({ const: 'none' }) }), required: Object.freeze(['style']), additionalProperties: false }),
]) });
// GH #290: keepNext / keepLines / pageBreakBefore / widowControl true OR false (false = an explicit override of the style).
const PAGINATION_FLAG = Object.freeze({ type: 'boolean' });
const FORMAT_TEXT_VALUE = Object.freeze({
  bold: Object.freeze({ type: 'boolean' }), italic: Object.freeze({ type: 'boolean' }),
  underline: Object.freeze({ type: 'boolean' }), strike: Object.freeze({ type: 'boolean' }),
  fontFamily: Object.freeze({ type: 'string', minLength: 1, maxLength: 128 }),
  fontSize: Object.freeze({ type: 'number', minimum: 1, maximum: 1638 }),
  // GH #324 follow-up: `SetLanguage` on an invalid-looking tag is a silent no-op at the engine
  // level (measured at the style layer, families/style-definition.cjs -- the run-level setter is
  // the SAME ApiTextPr call, families/formatting.cjs), so this field was previously unvalidated
  // and could look APPLIED while writing nothing. Pattern = LANGUAGE_PATTERN_SOURCE (imported
  // above), the SAME shape STYLE_TEXT.language uses, not a second copy.
  language: Object.freeze({ type: 'string', minLength: 1, maxLength: 35, pattern: LANGUAGE_PATTERN_SOURCE }),
  color: Object.freeze({ type: 'string', pattern: '^#[0-9a-fA-F]{6}$' }),
  highlight: Object.freeze({ enum: ['none', 'black', 'blue', 'cyan', 'darkBlue', 'darkCyan', 'darkGray', 'darkGreen', 'darkMagenta', 'darkRed', 'darkYellow', 'green', 'lightGray', 'magenta', 'red', 'white', 'yellow'] }),
  vertAlign: Object.freeze({ enum: ['baseline', 'subscript', 'superscript'] }),
  alignment: Object.freeze({ enum: ['left', 'right', 'center', 'justify'] }),
  spaceBefore: Object.freeze({ type: 'number', minimum: 0, maximum: 1000 }),
  spaceAfter: Object.freeze({ type: 'number', minimum: 0, maximum: 1000 }),
  indentLeft: Object.freeze({ type: 'number', minimum: -1000, maximum: 1000 }),
  indentRight: Object.freeze({ type: 'number', minimum: -1000, maximum: 1000 }),
  indentFirstLine: Object.freeze({ type: 'number', minimum: -1000, maximum: 1000 }),
  shading: Object.freeze({ type: 'string', pattern: '^[0-9a-fA-F]{6}$' }),
  tabs: Object.freeze({ type: 'array', minItems: 0, maxItems: 32, items: TAB_STOP }),
  borderTop: BORDER_SIDE, borderBottom: BORDER_SIDE, borderLeft: BORDER_SIDE, borderRight: BORDER_SIDE,
  keepNext: PAGINATION_FLAG, keepLines: PAGINATION_FLAG, pageBreakBefore: PAGINATION_FLAG, widowControl: PAGINATION_FLAG,
  // GH #231b (docs/GH231-FORMAT-TEXT-PARAGRAPH-PROPS-DECISION.md section 8.7): previously excluded
  // from GH231 due to a code-verified READ round-trip defect (docx-driver.cjs's GetSpacingLineValue
  // conversion), now fixed. Bounds are NOT independently LIVE-measured -- same, not-yet-measured-
  // bounds convention as fontSize's 1638pt ceiling above.
  lineSpacing: Object.freeze({ type: 'number', minimum: 0.1, maximum: 10 }),
  // GH #315 (LIVE-MEASURED -- the twips hypothesis AND the GetSpacingLineRule() getter both
  // confirmed, see docs/GH315-GH318-LINE-HEIGHT-CHARACTER-SPACING-PLAN.md section 3). lineSpacingRule +
  // lineHeight are a PAIR -- the family's own normalizeWanted() (families/formatting.cjs) enforces
  // "both or neither" and the mutual exclusion with `lineSpacing` above (same underlying property,
  // different rule); this schema only gates each field's own shape. Bounds NOT independently
  // LIVE-measured, same not-yet-measured-bounds convention as `lineSpacing`'s own ceiling above.
  lineSpacingRule: Object.freeze({ enum: ['exact', 'atLeast'] }),
  lineHeight: Object.freeze({ type: 'number', minimum: 1, maximum: 1000 }),
  // GH #318 (LIVE-MEASURED, same evidence as above): run-level, signed (OOXML allows condensed
  // text), no pairing requirement.
  characterSpacing: Object.freeze({ type: 'number', minimum: -1000, maximum: 1000 }),
});
const FORMAT_TEXT_FORMAT = Object.freeze({ type: 'object', properties: FORMAT_TEXT_VALUE, additionalProperties: false });
// GH #256 (PREPARATION, owner decision pending alongside #228..#232 -- docs/GH256-FORMAT-TABLE-BORDERS-SHADING-DECISION.md): format_table
// (families/table-format.cjs) released for the TABLE-LEVEL borders and shading axes -- plus (GH #257, below) cellMargins / indent / width; `style`
// stays UNKNOWN_FIELD. The table is named exactly like replace_table_cell's cell (CELL_TARGET, table.cjs cellSelector): the step key is `table` (the
// adapter's own, families/table-format.cjs STEP_KEYS), and only the one canonical { tableCell: { text [, tableAnchor] } } form is released.
// borders: any of the six edges (top/bottom/left/right/insideH/insideV), each EITHER { style:'single', size, color } -- size is an INTEGER in
// EIGHTHS OF A POINT (the adapter's and the editor's own unit, NOT points as in format_text's borders; 4 = 0.5pt, MEASURED) and color '#RRGGBB' --
// OR { style:'none' } alone (MEASURED: a 'none' write ignores size/color, so they are refused, not accepted-and-dropped). READ (read_tables
// table.format.borders) returns the SAME size in eighths under the name `sz`, the style under `value`, and the colour as { auto, r, g, b }.
// shading: '#RRGGBB'; the fixed 'clear' pattern is the adapter's constant (READ: { val:'clear', fill:{r,g,b}, color }). The 1..96 size range is a
// SCHEMA GUARD (96 eighths = 12pt), not a measurement; only 4 eighths (0.5pt) has LIVE evidence on the product-adapter path.
const TABLE_BORDER_EDGE = Object.freeze({ anyOf: Object.freeze([
  Object.freeze({
    type: 'object',
    properties: Object.freeze({ style: Object.freeze({ const: 'single' }), size: Object.freeze({ type: 'integer', minimum: 1, maximum: 96 }), color: Object.freeze({ type: 'string', pattern: '^#[0-9a-fA-F]{6}$' }) }),
    required: Object.freeze(['style', 'size', 'color']),
    additionalProperties: false,
  }),
  Object.freeze({ type: 'object', properties: Object.freeze({ style: Object.freeze({ const: 'none' }) }), required: Object.freeze(['style']), additionalProperties: false }),
]) });
// GH #257 (PREPARATION, owner decision pending alongside #228..#256 -- docs/GH257-FORMAT-TABLE-MARGINS-INDENT-WIDTH-DECISION.md): cellMargins / indent /
// width, the three TWIPS axes of table-format.cjs (validateCellMargins :83 / validateIndent :95 / validateWidth :98). The UNIT is TWIPS (1/20 pt) on WRITE
// (SetTableCellMargin{Top,Right,Bottom,Left}(n), SetTableInd(n), SetWidth('twips', n) -- the unit string is the adapter's fixed constant) and the READ returns
// the SAME twips ({ type:'dxa', w } for the margins and the width, a bare number for the indent) -- unlike format_text's indent*/space* keys, which are POINTS.
// cellMargins needs ALL FOUR edges (the adapter refuses fewer); each margin edge and the indent is a NON-NEGATIVE integer, the width a POSITIVE one.
// Percent / point units and a negative indent are NOT measured and NOT in the schema. The 0..31680 twip range (31680 = 22 inches) is a SCHEMA GUARD, not a
// measurement; LIVE evidence exists for margins 50/60/70/80, indent 100 and width 2000 (docs/evidence/docx-w04-table-format-live-acceptance-rebased-PASS-20261012.json),
// and (GH #257b) for an explicit ZERO on all four margin edges and on the indent through THIS product path (docs/GH257B-FORMAT-TABLE-ZERO-MARGIN-INDENT-DECISION.md).
// GH #257 refused zero pre-dispatch because the editor might drop an explicit 0 (read back UNSET = a failed L1 AFTER the write); the raw probe (g257v1 ROUND 6)
// measured it stored (indent 0, margin top { dxa, 0 }) and GH #257b lifted the refusal. A zero WIDTH stays refused: it was never probed or requested.
const TWIPS = Object.freeze({ type: 'integer', minimum: 0, maximum: 31680 });
const TABLE_CELL_MARGINS = Object.freeze({
  type: 'object',
  properties: Object.freeze({ top: TWIPS, right: TWIPS, bottom: TWIPS, left: TWIPS }),
  required: Object.freeze(['top', 'right', 'bottom', 'left']),
  additionalProperties: false,
});
const TABLE_BORDERS = Object.freeze({
  type: 'object',
  properties: Object.freeze(Object.fromEntries(['top', 'bottom', 'left', 'right', 'insideH', 'insideV'].map(e => [e, TABLE_BORDER_EDGE]))),
  additionalProperties: false,
});
// GH #265 (PREPARATION, owner decision pending alongside #228..#264 -- docs/GH265-DOCX-MODIFY-STYLE-DECISION.md): modify_style takes the SAME step shape and the SAME six fields as define_style
// (STYLE_NAME / STYLE_TEXT / STYLE_PARAGRAPH below) on an EXISTING style: a missing name fails closed (modify_style-name-does-not-exist), it never creates one; the retry IS idempotent
// (SATISFIED, 0 writes -- measured MODIFY_RETRY). The adapter's L3 proves every OTHER style and the body unchanged; it does NOT compare the modified style's own UNNAMED fields (the LIVE
// harness does, from the saved styles.xml). There is no way to RESET a field (no 'auto' colour, no unset): only a value can be written.
// GH #264 (PREPARATION, owner decision pending alongside #228..#257b -- docs/GH264-DOCX-DEFINE-STYLE-DECISION.md): define_style (families/style-definition.cjs) creates a
// NEW paragraph style by name -- the adapter's own intent name (the GitHub issue calls it define_paragraph_style). EXACTLY the six fields the family measured end to end
// (docs/evidence/docx-w10-style-definition-live-rebased-PASS-20261012.json): text { bold, size (points), color } and paragraph { align, indentLeft (points), spacingAfter (points) }.
// NOT the whole #229..#232 property set: highlight / spaceBefore / indentRight / tabs / borders / pagination have NO setter in the
// family and no measurement (italic / underline / fontFamily joined via GH #292, language via GH #324 -- see STYLE_TEXT.language below) -- they, `basedOn`
// (accepted-then-refused by the adapter: UNMEASURED) and a theme colour are UNKNOWN_FIELD. `color` is a BARE 6-hex string (no '#',
// the OPPOSITE of format_text's color, like shading). A name that already exists is decided by its CONTENT (GH #264): every requested field reads back as requested -> SATISFIED,
// nothing written (so a retry is our own retry, like every other intent's); a DIFFERENT field -> FAIL_CLOSED, OPERATION_FAILED, reason define_style-name-already-exists-with-different-fields:<name>:
// <differing group.key list>, never an overwrite (the adapter's classify, families/style-definition.cjs). READ-BACK: GH #224b gave read_styles a `names` argument -- the name-targeted
// doc.GetStyle read -- so a defined custom style, which the plain GetAllStyles enumeration is blind to (CUSTOM_STYLE_LISTED=false on DS 9.3.5, GH #224), IS readable through the product by name;
// the adapter also verifies the write through its own name-targeted read (adapter.cjs styleNamesOf).
// GUARDS, not measurements: name 1..128 printable characters; size on the 0.5-point grid and indentLeft / spacingAfter on the 0.25-point grid (float-safe: the L1 compares with
// strict equality after the editor's half-point / twip rounding); size 0.5..1638 (the format_text fontSize bound), indentLeft / spacingAfter 0..1000. Measured values: size 16, bold,
// color 2255AA, align center, indentLeft 18, spacingAfter 12.
const STYLE_NAME = Object.freeze({ type: 'string', minLength: 1, maxLength: 128, pattern: '^[^\\u0000-\\u001f\\u007f]+$' });
const STYLE_TEXT = Object.freeze({
  type: 'object',
  properties: Object.freeze({
    bold: Object.freeze({ type: 'boolean' }),
    size: Object.freeze({ type: 'number', minimum: 0.5, maximum: 1638, description: 'POINTS, a multiple of 0.5' }),
    color: Object.freeze({ type: 'string', pattern: '^([0-9a-fA-F]{6}|auto)$', description: 'BARE 6-hex-digit RGB, e.g. 2255AA -- NO leading #, UNLIKE format_text color (#RRGGBB); or "auto" (GH #292 S8)' }),
    // GH #292 S1 (LIVE, docs/evidence/gh292-style-definition-probe-20261021.json): stored on the style; the flags only as true (false not measured on a style)
    // GH #292 S8 (LIVE): false is an explicit off on the style (not a removal back to inherited)
    italic: Object.freeze({ type: 'boolean' }),
    underline: Object.freeze({ type: 'boolean' }),
    strike: Object.freeze({ type: 'boolean' }),
    fontFamily: Object.freeze({ type: 'string', minLength: 1, maxLength: 64, pattern: '^[^\\u0000-\\u001f\\u007f]+$' }),
    // GH #324 (LIVE, probe2.json): ApiTextPr.SetLanguage persists on a style (Normal + a fresh
    // custom one), read back via ToJSON() rPr.lang.val. The engine itself does NOT validate the
    // string (an unrecognized one silently writes nothing) -- this pattern is OUR OWN guard,
    // refusing pre-dispatch rather than letting a requested write silently do nothing. The pattern
    // itself is LANGUAGE_PATTERN_SOURCE (imported above from families/style-definition.cjs) -- the
    // SAME simplified BCP-47 shape format_text's own `language` field uses below.
    language: Object.freeze({ type: 'string', minLength: 1, maxLength: 35, pattern: LANGUAGE_PATTERN_SOURCE }),
  }),
  additionalProperties: false,
});
const STYLE_PARAGRAPH = Object.freeze({
  type: 'object',
  properties: Object.freeze({
    align: Object.freeze({ enum: ['left', 'right', 'center', 'justify'] }),
    indentLeft: Object.freeze({ type: 'number', minimum: 0, maximum: 1000, description: 'POINTS, a multiple of 0.25' }),
    spacingAfter: Object.freeze({ type: 'number', minimum: 0, maximum: 1000, description: 'POINTS, a multiple of 0.25' }),
    // GH #292 S2 (LIVE, same evidence): stored on the style; the flags only as true
    spacingBefore: Object.freeze({ type: 'number', minimum: 0, maximum: 1000, description: 'POINTS, a multiple of 0.25' }),
    indentRight: Object.freeze({ type: 'number', minimum: 0, maximum: 1000, description: 'POINTS, a multiple of 0.25' }),
    indentFirstLine: Object.freeze({ type: 'number', minimum: -1000, maximum: 1000, description: 'POINTS, a multiple of 0.25; negative = a hanging indent' }),
    keepNext: Object.freeze({ type: 'boolean' }),
    keepLines: Object.freeze({ type: 'boolean' }),
    pageBreakBefore: Object.freeze({ type: 'boolean' }),
  }),
  additionalProperties: false,
});
// GH #266: add_link / add_bookmark. The link target field is `href`, NOT `url`: the shared security walk (write-tools.cjs ROUTE_FIELDS) refuses a field named `url` at any depth as a broker / Nextcloud route
// override, and that gate is not relaxed; validateDocxWriteRequest maps href -> the adapter's `url` (toAdapterStep) before the planner and the executor see the step. The selector is the PARENT unit's exact text ({ text } only: no ref, no physical locator); the optional matchText is the piece of that unit the link / bookmark
// covers and must occur exactly once in it (the format_text pattern). The link target is an absolute http(s) URL or a mailto: address -- every other scheme (javascript:, file:, data:, a
// relative path, a bare host) is refused pre-dispatch, because the editor stores whatever string it is given. The bookmark name is the editor's identifier shape (the adapter's BOOKMARK_NAME).
const LINK_TARGET_TEXT = Object.freeze({ type: 'object', properties: Object.freeze({ text: SEL_TEXT }), required: Object.freeze(['text']), additionalProperties: false });
const LINK_URL = Object.freeze({ type: 'string', minLength: 8, maxLength: 2048,
  pattern: W.LINK_URL_PATTERN,
  description: 'absolute http:// or https:// URL, or a mailto: address (lower-case scheme); no whitespace or control characters; every other scheme is refused' });
const BOOKMARK_NAME_SCHEMA = Object.freeze({ type: 'string', pattern: '^[A-Za-z_][A-Za-z0-9_]{0,39}$', description: 'letters, digits, underscore; must not start with a digit; at most 40 characters' });
// GH #248: set_numbering fine-tunes the numbering a list item ALREADY sits on (it never makes a list). Released: `code` (the six templates), `start`, `suffix`; `custom` and `symbol` stay
// UNKNOWN_FIELD (open-ended format strings, not measured). The product step carries a POLICY the adapter's own default does not: a definition other items sit on (same numbering, same level)
// is NEVER written -- the step fails closed before dispatch with a named reason (toAdapterStep sets sharedDefinition: 'refuse'). `start` is bounded by the schema (1..9999), not by a measurement.
// GH #319: `bulletColor`/`bulletFont` released additively -- the bullet glyph's own colour/typeface, CODE:'bullet' required in the SAME step (families/numbering.cjs's plan(), same reasoning `symbol`'s own gate already applies).
const NUMBERING_START = Object.freeze({ type: 'integer', minimum: 1, maximum: 9999 });
// GH #293 L7 (owner GO B, 2026-10-22): `scope` -- 'item' (the default: the released semantics above, unchanged) or 'list' (EXPLICIT: every
// item of the target's list at its level changes, and the result names them in affectedItems). Both scopes take the affected set from the
// editor's own numbering collection through the word /5 Bridge (a table-cell item included); off /5, or when the set / its isolation is
// not provable, the step fails closed before any write. A definition another list shares is detached first (scope 'list' only, only
// when no count can flow between the lists).
const NUMBERING_SCOPE = Object.freeze({ enum: ['item', 'list'] });
// GH #319 (LIVE-measured, gh319-bulletcolorfont-probe-20261028): the bullet GLYPH's own colour
// and typeface -- through `level.GetTextPr()` (no direct setter on the level object). Schema-level
// this is type/shape only (the HEX_RE-equivalent pattern families/numbering.cjs's own plan() also
// re-checks); the CODE:'bullet'-in-the-same-step requirement is a cross-field rule this `anyOf`-
// free object schema cannot express, so it is enforced where `symbol`'s own identical requirement
// already is -- families/numbering.cjs's plan(), which this product schema dispatches into
// unchanged (a non-bullet code + either field is refused there, named, before any write).
const NUMBERING_BULLET_COLOR = Object.freeze({ type: 'string', pattern: '^#[0-9A-Fa-f]{6}$', description: '#RRGGBB -- the bullet glyph\'s own colour, CODE: \'bullet\' required in the same step' });
const NUMBERING_BULLET_FONT = Object.freeze({ type: 'string', minLength: 1, maxLength: 128, description: 'a font family name -- the bullet glyph\'s own typeface, CODE: \'bullet\' required in the same step. The glyph is a character code in THIS font: a symbol font (Symbol, Wingdings) draws a different shape for the same character, so set `symbol` together with `bulletFont`' });
// GH #262 (page setup margins + portrait size, product release, docs/GH262-PAGE-SETUP-WRITE-DECISION.md):
// the ONE-step product narrowing of families/page-setup.cjs's set_page_setup. The adapter's own `plan()`
// (page-setup.cjs) already accepts ANY positive-integer portrait size and ANY non-negative-integer margin
// (measured only the ARGUMENT ORDER and the write/read/reopen mechanics, with one arbitrary tuple chosen
// to be mutually distinguishable: size {11000,14000}, margins {top:1102,right:1203,bottom:1304,left:1001},
// docs/evidence/docx-w11-page-setup-live-acceptance-rebased-PASS-20261012.json -- NOT a claim that this
// specific tuple, or any OTHER specific tuple, is independently measured). This product schema narrows
// the mechanism-level acceptance to a bounded, named set, pre-dispatch, rather than pass an unmeasured
// shape through: `size` to a GRID of the two real-world portrait paper sizes a caller would plausibly
// want (US Letter / ISO A4 -- SetPageSize is a plain 2-int call, value-agnostic by the adapter's own
// measurement, so neither size is independently riskier than the other); `margins` to a per-side RANGE
// (0..2880 twips = 0..2in), not a full-tuple enum -- the measured tuple proved ARGUMENT ORDER with four
// distinct values, not "these four values are the only allowed ones", and the mechanism itself takes the
// four sides as independent arguments.
// GH #263 (columns + startPageNumber, docs/GH263-PAGE-SETUP-COLUMNS-STARTPAGE-DECISION.md): the same
// RANGE treatment as margins, not an enum -- neither field has a small, named, real-world set the way
// paper sizes do (any positive column count / any positive start page number is a legitimate caller
// intent), and the measured tuple (columns: {count:2, space:720}, startPageNumber:7, same evidence file
// above) proved argument order with ONE distinguishable tuple, not an exhaustive allowed set. Both caps
// below are this release's own choice, not a measured limit -- flagged for review in the decision doc.
const PAGE_SETUP_SIZES = Object.freeze([
  Object.freeze({ width: 12240, height: 15840 }), // US Letter, portrait (8.5in x 11in)
  Object.freeze({ width: 11906, height: 16838 }), // ISO A4, portrait (210mm x 297mm)
]);
// check() (write-tools.cjs) supports `anyOf`, not `oneOf` -- the two sizes are mutually exclusive by
// their own distinct const pairs, so "at least one matches" and "exactly one matches" coincide here.
const PAGE_SETUP_SIZE = Object.freeze({
  anyOf: PAGE_SETUP_SIZES.map(s => Object.freeze({ type: 'object',
    properties: Object.freeze({ width: Object.freeze({ const: s.width }), height: Object.freeze({ const: s.height }) }),
    required: Object.freeze(['width', 'height']), additionalProperties: false })),
});
const PAGE_SETUP_MARGIN_VALUE = Object.freeze({ type: 'integer', minimum: 0, maximum: 2880 });
const PAGE_SETUP_MARGINS = Object.freeze({
  type: 'object',
  properties: Object.freeze({ top: PAGE_SETUP_MARGIN_VALUE, right: PAGE_SETUP_MARGIN_VALUE, bottom: PAGE_SETUP_MARGIN_VALUE, left: PAGE_SETUP_MARGIN_VALUE }),
  required: Object.freeze(['top', 'right', 'bottom', 'left']),
  additionalProperties: false,
});
// GH #263: columns.count 1..12 (a practical cap -- the adapter's own plan() only requires >=1), columns.space
// 0..2880 twips (same convention as margins, since it is the same unit and the same "gap" semantics).
const PAGE_SETUP_COLUMNS = Object.freeze({
  type: 'object',
  properties: Object.freeze({
    count: Object.freeze({ type: 'integer', minimum: 1, maximum: 12 }),
    space: Object.freeze({ type: 'integer', minimum: 0, maximum: 2880 }),
  }),
  required: Object.freeze(['count', 'space']),
  additionalProperties: false,
});
// GH #263: startPageNumber 1..9999 -- a generous practical cap (four digits covers any plausible printed
// page number), not a measured limit; the adapter's own plan() only requires a positive integer.
const PAGE_SETUP_START_PAGE_NUMBER = Object.freeze({ type: 'integer', minimum: 1, maximum: 9999 });
// section is REQUIRED, explicit, no implicit section -- same discipline page-setup.cjs's own plan()
// already enforces (its own `s.of < 1 || s.index < 0 || s.index >= s.of` check); this schema's own
// minimum bounds are pre-dispatch, the adapter's own range check at execution is the second, authoritative
// layer (defense in depth, same pattern as REWRITE_SELECTOR's rewriteSelectorBranch below).
const PAGE_SETUP_SECTION = Object.freeze({
  type: 'object',
  properties: Object.freeze({ index: Object.freeze({ type: 'integer', minimum: 0 }), of: Object.freeze({ type: 'integer', minimum: 1 }) }),
  required: Object.freeze(['index', 'of']),
  additionalProperties: false,
});
// GH #289 paragraph mode (docs/GH289-HF-DOCUMENTED-LIMITS-DECISION.md, H-01): the paragraph of an EXISTING part whose
// text changes, its formatting kept; `of` = the part's paragraph count the caller read (read_headers_footers).
const HF_PARAGRAPH = Object.freeze({
  type: 'object',
  properties: Object.freeze({ index: Object.freeze({ type: 'integer', minimum: 0 }), of: Object.freeze({ type: 'integer', minimum: 1 }) }),
  required: Object.freeze(['index', 'of']),
  additionalProperties: false,
});
// GH #289 H-07a: the later sections that INHERIT the written part (they show it without one of their own), named by the
// caller exactly as the refusal / read lists them -- a write that changes what other sections show is never implicit.
const HF_INHERITED_BY = Object.freeze({ type: 'array', items: Object.freeze({ type: 'integer', minimum: 1 }), uniqueItems: true, maxItems: 1000 });
// GH #317 "lines" (step 3): an alternative to `text` -- a full-part write of N paragraphs.
// maxItems 200 is a SCHEMA bound (a header/footer realistically holds a handful of lines), not a
// measured limit -- header-footer.cjs's own plan() is the authority for the per-line text rules
// (non-empty, no tab/line break).
const HF_LINES = Object.freeze({ type: 'array', minItems: 1, maxItems: 200, items: Object.freeze({ type: 'string', minLength: 1, maxLength: 32767 }) });
const notesStepSchema = intent => Object.freeze({
  type: 'object',
  properties: Object.freeze({ intent: Object.freeze({ const: intent }), id: ID, selector: LIST_TARGET, text: Object.freeze({ type: 'string', minLength: 1, maxLength: 32767 }) }),
  required: Object.freeze(['intent', 'selector', 'text']),
  additionalProperties: false,
});
// GH #269e: the comment string shapes (the adapter's plan() refuses a blank text / author after this).
const COMMENT_TEXT = Object.freeze({ type: 'string', minLength: 1, maxLength: 32767 });
const COMMENT_AUTHOR = Object.freeze({ type: 'string', minLength: 1, maxLength: 512 });
const COMMENT_QUOTE = Object.freeze({ type: 'string', maxLength: 32767 });
// Release 3 trim: the step schemas / constraints of EVERY routed DOCX intent stay defined here (_ALL); the PUBLISHED sets below are exactly the allowlist (a re-release is a one-line
// allowlist change in write-tools.cjs + tools.cjs, nothing is re-written).
// GH #268d: the target is named EXACTLY one of two ways, never by docPr.id/name (dropped from
// read_images entirely -- GH #268b, editor-session-local, measured LIVE). { paragraph, ordinal? } --
// the ordinal-th (1-based, default 1) image whose parent paragraph text equals `paragraph`, counted
// in document order; or { index } -- the document-level index read_images itself returns, valid for
// one read only. The family's own plan() (image.cjs) re-enforces the XOR and re-resolves against a
// FRESH read_images observation at classify/resolve time -- never a cached one.
const IMAGE_TARGET = Object.freeze({
  type: 'object',
  properties: Object.freeze({
    paragraph: SEL_TEXT,
    ordinal: Object.freeze({ type: 'integer', minimum: 1, maximum: 1000 }),
    index: Object.freeze({ type: 'integer', minimum: 0, maximum: 100000 }),
  }),
  additionalProperties: false,
});
// GH #296: delete_image / replace_image -- IMAGE_TARGET plus an optional `part` assertion (the
// read_images row's own part label); the family's own plan() (image.cjs) re-checks the shape and
// fails closed on a mismatch, never re-targeting.
const IMAGE_EDIT_TARGET = Object.freeze({
  type: 'object',
  properties: Object.freeze({
    ...IMAGE_TARGET.properties,
    part: Object.freeze({ type: 'string', pattern: '^(body|(header|footer):[0-9]+:(default|title|even))$', maxLength: 64 }),
  }),
  additionalProperties: false,
});
// GH #308: insert_table. The anchor is the unique body-level paragraph/heading the new table goes
// right after (table-insert.cjs's own resolveAnchor -- a table/list, or a text that only occurs
// inside a table cell, is refused the same way every other family's resolveStableTarget call
// already refuses it). rows/cols 1..50 mirrors the family's own MAX_DIM (a schema sanity cap, not a
// measured limit -- Api.CreateTable(0 or negative) returns null, measured). `cells`, when given,
// must be an array of exactly `rows` rows of exactly `cols` strings (the family's own plan() is the
// authority for the exact-size cross-check; this schema states only the per-dimension bound).
const TABLE_INSERT_TARGET = Object.freeze({
  type: 'object',
  properties: Object.freeze({ afterParagraph: SEL_TEXT }),
  required: Object.freeze(['afterParagraph']),
  additionalProperties: false,
});
const TABLE_INSERT_DIM = Object.freeze({ type: 'integer', minimum: 1, maximum: 50 });
const TABLE_INSERT_CELLS = Object.freeze({
  type: 'array', minItems: 1, maxItems: 50,
  items: Object.freeze({ type: 'array', minItems: 1, maxItems: 50, items: Object.freeze({ type: 'string', maxLength: 32767 }) }),
});
// GH #313: insert_section_break. The anchor is the unique body-level paragraph/heading the new
// section boundary closes (section-break-insert.cjs's own resolveAnchor -- same contract as
// TABLE_INSERT_TARGET above). type is required: only nextPage/continuous were LIVE-measured
// (test/live/docx-gh313-section-break-probe.live.cjs) -- evenPage/oddPage/nextColumn are refused
// here rather than guessed, the same NEM-MÉRT discipline set_image_wrap's `tight` value uses.
const SECTION_BREAK_TARGET = Object.freeze({
  type: 'object',
  properties: Object.freeze({ afterParagraph: SEL_TEXT }),
  required: Object.freeze(['afterParagraph']),
  additionalProperties: false,
});
const SECTION_BREAK_TYPE = Object.freeze({ enum: ['nextPage', 'continuous'] });
// GH #306: insert_image. The anchor is the unique body-level paragraph/heading the new image goes
// relative to (image-insert.cjs's own resolveAnchor -- same contract as TABLE_INSERT_TARGET
// above). `source` is `data:image/png;base64,...` OR `data:image/jpeg;base64,...` ONLY -- an
// external URL is refused by the pattern itself, never fetched; maxLength 1,400,000 is the SAME
// 1 MiB margin the XLSX insert_image schema already uses (a 1 MiB PNG base64-encodes to exactly
// 1,398,126 characters with this prefix). widthEmu/heightEmu mirror set_image_size's own positive-
// integer convention. `position` ('after' default | 'end') is image-insert.cjs's own field --
// see that family's header comment for the measured difference and the RETRY IDENTITY limit.
//
// GH #306b 2. resz: the header/footer alternative shape, mutually exclusive with the body shape
// above by `additionalProperties: false` on both (check() supports `anyOf`, not `oneOf` -- same
// discipline PAGE_SETUP_SIZE above already uses, and "at least one matches" = "exactly one
// matches" here for the SAME reason: the two shapes require disjoint key sets). `section` reuses
// PAGE_SETUP_SECTION (the SAME { index, of } contract edit_header/edit_footer already use);
// `part` is header | footer; `variant` is default | title | even -- GH #296 part 3: title/even
// only on an ALREADY-EXISTING, ALREADY-ACTIVE part, never activated (image-insert.cjs's own header
// comment, LIVE-measured via docx-gh296-hf-title-even-image-insert-probe.live.cjs). `position` (below) picks
// 'end' (the part's own last/named paragraph's existing run flow, append -- the ORIGINAL, still the
// default) or 'start' (a fresh run inserted at index 0 of that SAME paragraph).
// GH #317 (step 3): optional `paragraph` -- the SAME HF_PARAGRAPH descriptor
// edit_header/edit_footer's own paragraph mode already uses, reused as-is; absent keeps the
// pre-existing last-paragraph behaviour.
// GH #317 (step 4): optional `position` -- `'start'|'end'` (default `'end'` = the pre-existing,
// unchanged behaviour), LIVE-measured via docx-gh317-image-start-of-paragraph-probe.live.cjs.
// Its OWN named enum, a DIFFERENT schema node from the body target's step-level `position`
// (IMAGE_INSERT_POSITION, 'after'|'end' -- a different contract entirely) -- the two never share a
// property, so this never widens the body shape.
const IMAGE_INSERT_HF_POSITION = Object.freeze({ enum: ['start', 'end'] });
const IMAGE_INSERT_HEADER_FOOTER_TARGET = Object.freeze({
  type: 'object',
  properties: Object.freeze({ section: PAGE_SETUP_SECTION, part: Object.freeze({ enum: ['header', 'footer'] }), variant: Object.freeze({ enum: ['default', 'title', 'even'] }), paragraph: HF_PARAGRAPH, position: IMAGE_INSERT_HF_POSITION }),
  required: Object.freeze(['section', 'part']),
  additionalProperties: false,
});
const IMAGE_INSERT_TARGET = Object.freeze({
  anyOf: Object.freeze([
    Object.freeze({ type: 'object', properties: Object.freeze({ afterParagraph: SEL_TEXT }), required: Object.freeze(['afterParagraph']), additionalProperties: false }),
    IMAGE_INSERT_HEADER_FOOTER_TARGET,
    // GH #306 part 4: the table-cell target -- replace_table_cell's own { tableCell: { text, tableAnchor? } } cell identity (image-insert.cjs planCell)
    Object.freeze({ type: 'object', properties: Object.freeze({ tableCell: Object.freeze({ type: 'object', properties: Object.freeze({ text: SEL_TEXT, tableAnchor: SEL_TEXT }), required: Object.freeze(['text']), additionalProperties: false }) }), required: Object.freeze(['tableCell']), additionalProperties: false }),
  ]),
});
const IMAGE_INSERT_SOURCE = Object.freeze({ type: 'string', pattern: '^data:image/(png|jpeg);base64,[A-Za-z0-9+/]+=*$', maxLength: 1400000 });
const IMAGE_INSERT_EMU = Object.freeze({ type: 'integer', minimum: 1 });
const IMAGE_INSERT_POSITION = Object.freeze({ enum: ['after', 'end'] });
// GH #322: insert_paragraphs. The anchor is the SAME { afterParagraph } contract as
// IMAGE_INSERT_TARGET/TABLE_INSERT_TARGET above (paragraphs-insert.cjs's own resolveAnchor).
// `paragraphs` is 1..200 items (the LIVE-measured ceiling, docx-gh322-bulk-insert-probe.live.cjs:
// 200 in one callCommand still lands in 36 ms); `format` reuses FORMAT_TEXT_FORMAT as-is (ONE
// constant, not a second copy) -- paragraphs-insert.cjs's own plan() is the authority that narrows
// the accepted keys to its subset (PARAGRAPHS_FORMAT_KEYS), the SAME split edit_header/edit_footer
// already use for their own `format` field above.
const PARAGRAPHS_INSERT_TARGET = Object.freeze({
  type: 'object',
  properties: Object.freeze({ afterParagraph: SEL_TEXT }),
  required: Object.freeze(['afterParagraph']),
  additionalProperties: false,
});
// GH #309: insert_shape. The anchor is the unique body-level paragraph/heading the new shape goes after (shape-insert.cjs's own resolveAnchor -- same contract as IMAGE_INSERT_TARGET above).
// `shape.type` is the two LIVE-measured presets only; widthMm/heightMm are integers (the measured sizes were whole millimetres, EMU = mm * 36000); fill/stroke are 6-digit hex; text is one plain
// line (no control characters) -- shape-insert.cjs plan() re-checks every one of these, the schema refuses the same shapes up front.
const SHAPE_INSERT_TARGET = Object.freeze({
  type: 'object',
  properties: Object.freeze({ afterParagraph: SEL_TEXT }),
  required: Object.freeze(['afterParagraph']),
  additionalProperties: false,
});
// GH #323: insert_chart. The anchor is the unique body-level paragraph/heading the new chart goes after (chart-insert.cjs REUSES insert_shape's own resolveAnchor). `chart.type` is the
// LIVE-measured set only; series values are finite numbers, one per category; chart-insert.cjs plan() re-checks every one of these (and the pie/doughnut single-series rule, the one-value-per-
// category rule and unique series names, which a JSON schema cannot express), the schema refuses the same shapes up front.
const CHART_INSERT_TARGET = Object.freeze({
  type: 'object',
  properties: Object.freeze({ afterParagraph: SEL_TEXT }),
  required: Object.freeze(['afterParagraph']),
  additionalProperties: false,
});
const CHART_LABEL = Object.freeze({ type: 'string', minLength: 1, maxLength: 100, pattern: '^[^\\u0000-\\u001f\\u007f]+$' });
const CHART_INSERT_CHART = Object.freeze({
  type: 'object',
  properties: Object.freeze({
    type: Object.freeze({ enum: ['bar', 'barStacked', 'barStackedPercent', 'horizontalBar', 'horizontalBarStacked', 'area', 'areaStacked', 'pie', 'doughnut'] }),
    categories: Object.freeze({ type: 'array', minItems: 1, maxItems: 100, items: CHART_LABEL }),
    series: Object.freeze({
      type: 'array', minItems: 1, maxItems: 10,
      items: Object.freeze({
        type: 'object',
        properties: Object.freeze({ name: CHART_LABEL, values: Object.freeze({ type: 'array', minItems: 1, maxItems: 100, items: Object.freeze({ type: 'number' }) }) }),
        required: Object.freeze(['name', 'values']),
        additionalProperties: false,
      }),
    }),
    widthMm: Object.freeze({ type: 'integer', minimum: 10, maximum: 400 }),
    heightMm: Object.freeze({ type: 'integer', minimum: 10, maximum: 400 }),
    title: Object.freeze({ type: 'string', minLength: 1, maxLength: 200, pattern: '^[^\\u0000-\\u001f\\u007f]+$' }),
    legend: Object.freeze({ enum: ['bottom'] }),
  }),
  required: Object.freeze(['type', 'categories', 'series', 'widthMm', 'heightMm']),
  additionalProperties: false,
});
const PARAGRAPHS_INSERT_ITEM = Object.freeze({
  type: 'object',
  properties: Object.freeze({
    text: Object.freeze({ type: 'string', minLength: 1, maxLength: 32767 }),
    style: Object.freeze({ type: 'string', minLength: 1, maxLength: 255 }),
    format: FORMAT_TEXT_FORMAT,
  }),
  required: Object.freeze(['text']),
  additionalProperties: false,
});
const PARAGRAPHS_INSERT_LIST = Object.freeze({ type: 'array', minItems: 1, maxItems: 200, items: PARAGRAPHS_INSERT_ITEM });
// GH #309 part 3: the optional FLOATING placement -- exactly the measured combinations (shape-insert.cjs planPlacement re-checks every one; 'tight' is not offered, see there)
const SHAPE_OFFSET_MM = Object.freeze({ type: 'integer', minimum: 0, maximum: 400 });
const SHAPE_INSERT_PLACEMENT = Object.freeze({
  type: 'object',
  properties: Object.freeze({
    wrap: Object.freeze({ enum: ['square', 'topAndBottom', 'inFront', 'behind'] }),
    horizontal: Object.freeze({ anyOf: Object.freeze([
      Object.freeze({ type: 'object', properties: Object.freeze({ relativeTo: Object.freeze({ const: 'margin' }), align: Object.freeze({ enum: ['left', 'center', 'right'] }) }), required: Object.freeze(['relativeTo', 'align']), additionalProperties: false }),
      Object.freeze({ type: 'object', properties: Object.freeze({ relativeTo: Object.freeze({ const: 'page' }), offsetMm: SHAPE_OFFSET_MM }), required: Object.freeze(['relativeTo', 'offsetMm']), additionalProperties: false }),
    ]) }),
    vertical: Object.freeze({ type: 'object', properties: Object.freeze({ relativeTo: Object.freeze({ enum: ['paragraph', 'page'] }), offsetMm: SHAPE_OFFSET_MM }), required: Object.freeze(['relativeTo', 'offsetMm']), additionalProperties: false }),
  }),
  required: Object.freeze(['wrap', 'horizontal', 'vertical']),
  additionalProperties: false,
});
const SHAPE_INSERT_SHAPE = Object.freeze({
  type: 'object',
  properties: Object.freeze({
    type: Object.freeze({ enum: ['rect', 'flowChartProcess', 'roundRect', 'ellipse', 'triangle', 'trapezoid', 'flowChartDecision', 'flowChartTerminator', 'rightArrow', 'chevron', 'homePlate'] }),
    widthMm: Object.freeze({ type: 'integer', minimum: 1, maximum: 400 }),
    heightMm: Object.freeze({ type: 'integer', minimum: 1, maximum: 400 }),
    fill: Object.freeze({ type: 'string', pattern: '^[0-9A-Fa-f]{6}$' }),
    stroke: Object.freeze({ type: 'string', pattern: '^[0-9A-Fa-f]{6}$' }),
    text: Object.freeze({ type: 'string', minLength: 1, maxLength: 500, pattern: '^[^\\u0000-\\u001f\\u007f]+$' }),
    placement: SHAPE_INSERT_PLACEMENT,
  }),
  required: Object.freeze(['type', 'widthMm', 'heightMm', 'fill', 'stroke', 'text']),
  additionalProperties: false,
});
const DOCX_STEP_SCHEMAS_ALL = Object.freeze({
  rewrite_paragraph: Object.freeze({
    type: 'object',
    properties: Object.freeze({
      intent: Object.freeze({ const: 'rewrite_paragraph' }), id: ID, selector: REWRITE_SELECTOR,
      text: Object.freeze({ type: 'string', minLength: 1, maxLength: 32767 }),
      expectedPreservation: Object.freeze({ type: 'object', additionalProperties: false,
        properties: Object.freeze(Object.fromEntries(['bold', 'italic', 'underline', 'strike', 'fontFamily', 'fontSize', 'language'].map(k => [k, PRESERVED_VALUE]))) }),
    }),
    required: Object.freeze(['intent', 'selector', 'text']),
    additionalProperties: false,
  }),
  insert_paragraph_before: insertStepSchema('insert_paragraph_before'),
  insert_paragraph_after: insertStepSchema('insert_paragraph_after'),
  delete_paragraph: Object.freeze({
    type: 'object',
    properties: Object.freeze({ intent: Object.freeze({ const: 'delete_paragraph' }), id: ID, selector: DELETE_SELECTOR }),
    required: Object.freeze(['intent', 'selector']),
    additionalProperties: false,
  }),
  move_paragraph_before: moveStepSchema('move_paragraph_before'),
  move_paragraph_after: moveStepSchema('move_paragraph_after'),
  apply_list: listStepSchema('apply_list', { type: LIST_TYPE, level: LIST_LEVEL }, ['type']),
  set_list_level: listStepSchema('set_list_level', { level: LIST_LEVEL }, ['level']),
  indent_list_item: listRelativeSchema('indent_list_item'),
  outdent_list_item: listRelativeSchema('outdent_list_item'),
  remove_list: listRelativeSchema('remove_list'),
  replace_table_cell: Object.freeze({
    type: 'object',
    properties: Object.freeze({
      intent: Object.freeze({ const: 'replace_table_cell' }), id: ID, selector: CELL_TARGET,
      text: Object.freeze({ type: 'string', minLength: 1, maxLength: 32767 }),
    }),
    required: Object.freeze(['intent', 'selector', 'text']),
    additionalProperties: false,
  }),
  format_text: Object.freeze({
    type: 'object',
    properties: Object.freeze({
      intent: Object.freeze({ const: 'format_text' }), id: ID, selector: REWRITE_SELECTOR, format: FORMAT_TEXT_FORMAT,
      // the parent's own literal substring (second slice, see the FORMAT_TEXT_VALUE comment above) --
      // reuses SEL_TEXT's bound, since matchText is itself a piece of paragraph text.
      matchText: SEL_TEXT,
    }),
    required: Object.freeze(['intent', 'selector', 'format']),
    additionalProperties: false,
  }),
  insert_table_row_before: rowInsertSchema('insert_table_row_before'),
  insert_table_row_after: rowInsertSchema('insert_table_row_after'),
  delete_table_row: Object.freeze({
    type: 'object',
    properties: Object.freeze({ intent: Object.freeze({ const: 'delete_table_row' }), id: ID, selector: ANCHORED_CELL_TARGET, terminalRowCount: ROW_COUNT_AFTER_DELETE }),
    required: Object.freeze(['intent', 'selector', 'terminalRowCount']),
    additionalProperties: false,
  }),
  insert_table_column_before: columnInsertSchema('insert_table_column_before'),
  insert_table_column_after: columnInsertSchema('insert_table_column_after'),
  delete_table_column: Object.freeze({
    type: 'object',
    properties: Object.freeze({ intent: Object.freeze({ const: 'delete_table_column' }), id: ID, selector: ANCHORED_CELL_TARGET, terminalColumnCount: COLUMN_COUNT_AFTER_DELETE }),
    required: Object.freeze(['intent', 'selector', 'terminalColumnCount']),
    additionalProperties: false,
  }),
  merge_table_cells: Object.freeze({
    type: 'object',
    properties: Object.freeze({ intent: Object.freeze({ const: 'merge_table_cells' }), id: ID, selector: ANCHORED_CELL_TARGET, endSelector: END_CELL_TARGET, terminalMergedText: SEL_TEXT, terminalMergedSpan: MERGED_SPAN }),
    required: Object.freeze(['intent', 'selector', 'endSelector', 'terminalMergedText', 'terminalMergedSpan']),
    additionalProperties: false,
  }),
  split_table_cell: Object.freeze({
    type: 'object',
    properties: Object.freeze({ intent: Object.freeze({ const: 'split_table_cell' }), id: ID, selector: ANCHORED_CELL_TARGET, splitRows: SPLIT_ROWS, splitCols: SPLIT_COLS }),
    required: Object.freeze(['intent', 'selector', 'splitRows', 'splitCols']),
    additionalProperties: false,
  }),
  format_table: Object.freeze({
    type: 'object',
    properties: Object.freeze({
      intent: Object.freeze({ const: 'format_table' }), id: ID, table: CELL_TARGET, borders: TABLE_BORDERS,
      shading: Object.freeze({ type: 'string', pattern: '^#[0-9a-fA-F]{6}$' }),
      style: SEL_TEXT,
      cellMargins: TABLE_CELL_MARGINS, indent: TWIPS, width: Object.freeze({ type: 'integer', minimum: 1, maximum: 31680 }),
    }),
    required: Object.freeze(['intent', 'table']),
    additionalProperties: false,
  }),
  define_style: Object.freeze({
    type: 'object',
    // GH #292 S3 (LIVE): basedOn = the NAME of an existing style of the same type (ApiStyle.SetBasedOn); define_style only
    // GH #292 S5 (LIVE): type 'character' creates a character style (text fields only); absent = 'paragraph'
    properties: Object.freeze({ intent: Object.freeze({ const: 'define_style' }), id: ID, name: STYLE_NAME, type: Object.freeze({ enum: ['paragraph', 'character'] }), basedOn: STYLE_NAME, text: STYLE_TEXT, paragraph: STYLE_PARAGRAPH }),
    required: Object.freeze(['intent', 'name']),
    additionalProperties: false,
  }),
  modify_style: Object.freeze({
    type: 'object',
    // GH #292 S4 (word /4 Bridge WRITE): next / link = the NAME of another style; next paragraph -> paragraph, link a paragraph and a
    // character style (both sides set); modify_style only, no clearing
    properties: Object.freeze({ intent: Object.freeze({ const: 'modify_style' }), id: ID, name: STYLE_NAME, next: STYLE_NAME, link: STYLE_NAME, text: STYLE_TEXT, paragraph: STYLE_PARAGRAPH }),
    required: Object.freeze(['intent', 'name']),
    additionalProperties: false,
  }),
  add_link: Object.freeze({
    type: 'object',
    properties: Object.freeze({ intent: Object.freeze({ const: 'add_link' }), id: ID, selector: LINK_TARGET_TEXT, matchText: SEL_TEXT, href: LINK_URL, tooltip: Object.freeze({ type: 'string', minLength: 1, maxLength: 255 }) }),
    required: Object.freeze(['intent', 'selector', 'href']),
    additionalProperties: false,
  }),
  add_bookmark: Object.freeze({
    type: 'object',
    properties: Object.freeze({ intent: Object.freeze({ const: 'add_bookmark' }), id: ID, selector: LINK_TARGET_TEXT, matchText: SEL_TEXT, name: BOOKMARK_NAME_SCHEMA }),
    required: Object.freeze(['intent', 'selector', 'name']),
    additionalProperties: false,
  }),
  set_numbering: Object.freeze({
    type: 'object',
    properties: Object.freeze({ intent: Object.freeze({ const: 'set_numbering' }), id: ID, selector: LINK_TARGET_TEXT,
      code: Object.freeze({ enum: ['1.', '1)', 'a.', 'I.', 'bullet', 'none'] }), start: NUMBERING_START, suffix: Object.freeze({ enum: ['tab', 'space', 'none'] }), scope: NUMBERING_SCOPE,
      bulletColor: NUMBERING_BULLET_COLOR, bulletFont: NUMBERING_BULLET_FONT }),
    required: Object.freeze(['intent', 'selector']),
    additionalProperties: false,
  }),
  set_list_type: listStepSchema('set_list_type', { type: LIST_TYPE }, ['type']),
  set_page_setup: Object.freeze({
    type: 'object',
    properties: Object.freeze({
      intent: Object.freeze({ const: 'set_page_setup' }), id: ID,
      section: PAGE_SETUP_SECTION, size: PAGE_SETUP_SIZE, margins: PAGE_SETUP_MARGINS,
      columns: PAGE_SETUP_COLUMNS, startPageNumber: PAGE_SETUP_START_PAGE_NUMBER,
    }),
    required: Object.freeze(['intent', 'section']),
    additionalProperties: false,
  }),
  // GH #259 (common/default header text, docs/GH259-DOCX-EDIT-HEADER-COMMON-DECISION.md): the adapter's
  // own plan() (header-footer.cjs) accepts a `variant` field (default/title/even) and treats an ABSENT
  // variant as 'default'. This product schema deliberately does NOT expose `variant` at all --
  // additionalProperties:false therefore refuses it as UNKNOWN_FIELD, which pins every released step to
  // the default variant, the ONLY one with no B21-BRIDGE dependency (CORRECTED, per review: flagOn()
  // returns true unconditionally for 'default', but shapeGuard()'s variantStateGuard() still requires
  // every section's titlePage/evenAndOdd to be known for EVERY variant including 'default' -- that read
  // is the SAME inventory.cjs variantFlag() pass GH #227's own read_section_flags already exposes, not
  // a new or hidden one; only the __eoB21Variant runtime bridge call stays gated on variant !== 'default'.
  // The census (docs/evidence/m11-4a0/census.json) still lists edit_header under B21 for its title/even
  // bridge dependency). Same key-restricted-release shape as format_range's X-N01 precedent: the intent
  // stays in TEMPORARY_BOUNDARY (its other shape, title/even, #261's own scope), this release narrows to
  // the one sub-shape with no BRIDGE dependency. No `selector` either -- the target is section + part +
  // variant, not a text match (header-footer.cjs's own plan() already refuses a `selector` on this intent).
  // GH #261 1/2: `variant` is now exposed (title | even), additive to GH #259's default-only
  // shape above -- omitted still means 'default' (the adapter's own plan(), unchanged). GH #261c:
  // writing to an INACTIVE title/even variant now ACTIVATES it (the adapter's own public-API
  // SetTitlePage/SetEvenAndOddHdrFtr path, stored LIVE evidence #261b) -- not gated here either
  // way: the schema only names the two shapes this release knows about.
  // GH #316: `format` reuses the format_text schema object as-is (ONE constant, not a second copy)
  // -- header-footer.cjs's own plan() is the authority that narrows the accepted keys to its subset.
  edit_header: Object.freeze({
    type: 'object',
    properties: Object.freeze({
      intent: Object.freeze({ const: 'edit_header' }), id: ID,
      section: PAGE_SETUP_SECTION,
      text: Object.freeze({ type: 'string', minLength: 1, maxLength: 32767 }),
      lines: HF_LINES,
      variant: Object.freeze({ enum: ['title', 'even'] }),
      paragraph: HF_PARAGRAPH,
      inheritedBy: HF_INHERITED_BY,
      pageNumber: Object.freeze({ type: 'boolean' }),
      format: FORMAT_TEXT_FORMAT,
    }),
    // GH #317: neither `text` nor `lines` is schema-required -- exactly one of the two is
    // header-footer.cjs's own plan() rule (a schema anyOf would need to duplicate every OTHER
    // property into two branches; the cross-field check belongs with the family's other
    // mutual-exclusion rules, e.g. lines+paragraph, the SAME place they already live).
    required: Object.freeze(['intent', 'section']),
    additionalProperties: false,
  }),
  // GH #260: the SAME shape and the SAME reasoning as edit_header above (docs/GH259-DOCX-EDIT-
  // HEADER-COMMON-DECISION.md applies verbatim, mirrored onto `footer` -- header-footer.cjs's own
  // flagOn() treats header and footer identically, keyed only by `op.variant`, never by `op.part`).
  // GH #261 1/2: same `variant` addition as edit_header above.
  edit_footer: Object.freeze({
    type: 'object',
    properties: Object.freeze({
      intent: Object.freeze({ const: 'edit_footer' }), id: ID,
      section: PAGE_SETUP_SECTION,
      text: Object.freeze({ type: 'string', minLength: 1, maxLength: 32767 }),
      lines: HF_LINES,
      variant: Object.freeze({ enum: ['title', 'even'] }),
      paragraph: HF_PARAGRAPH,
      inheritedBy: HF_INHERITED_BY,
      pageNumber: Object.freeze({ type: 'boolean' }),
      format: FORMAT_TEXT_FORMAT,
    }),
    required: Object.freeze(['intent', 'section']),
    additionalProperties: false,
  }),
  // GH #267 B (docs/GH267-DOCX-NOTES-WRITE-DECISION.md; LIVE g267v1 PASS R1-R3): add_footnote / add_endnote. The target is the LIVE-measured { text } shape (ONE paragraph that
  // must resolve uniquely -- the adapter's notes.cjs inlineTarget; no { ref }, no fingerprint, no physical locator), the note is ONE paragraph of text (a multi-paragraph note WRITE and
  // REPLACING an existing note's text are NOT released: no routed replace intent, and a different text on the same target APPENDS a second note -- g267v1 ROUND4).
  add_footnote: notesStepSchema('add_footnote'),
  add_endnote: notesStepSchema('add_endnote'),
  // GH #268d (routed + coded, WITHHELD -- docs/GH268C-DOCX-IMAGE-SIZE-WRAP-WRITE-DECISION.md; LIVE
  // g268cv1 PASS through withSession+docxAdapter+createExecutor, NOT yet through this handler):
  // set_image_size / set_image_wrap. The target is EXACTLY one of { paragraph, ordinal? } or
  // { index } -- never docPr.id/name (editor-session-local, GH #268b's own drop-id-name commit).
  // The family's own plan() (image.cjs) is the authority for the XOR and the re-resolution against
  // a fresh read_images observation; this schema's job is only the field shapes and the wrap enum.
  set_image_size: Object.freeze({
    type: 'object',
    properties: Object.freeze({
      intent: Object.freeze({ const: 'set_image_size' }), id: ID, selector: IMAGE_TARGET,
      widthEmu: Object.freeze({ type: 'integer', minimum: 1, maximum: 100000000 }),
      heightEmu: Object.freeze({ type: 'integer', minimum: 1, maximum: 100000000 }),
    }),
    required: Object.freeze(['intent', 'selector', 'widthEmu', 'heightEmu']),
    additionalProperties: false,
  }),
  // GH #268d: SetWrappingStyle. ONLY 'square' -- the one transition R4 LIVE-measured (inline ->
  // anchor via SetWrappingStyle('square')); every other value refused here AND again inside the
  // family's own commandBody (NEM-MÉRT, never assumed).
  set_image_wrap: Object.freeze({
    type: 'object',
    properties: Object.freeze({
      intent: Object.freeze({ const: 'set_image_wrap' }), id: ID, selector: IMAGE_TARGET,
      wrap: Object.freeze({ enum: ['square'] }),
    }),
    required: Object.freeze(['intent', 'selector', 'wrap']),
    additionalProperties: false,
  }),
  // GH #269e (routed + coded, WITHHELD -- docs/GH269e-DOCX-COMMENTS-PRODUCT-SCHEMAS-DECISION.md; the adapter side is #269c/#269d, offline-green, LIVE through the executor): the
  // three comment intents. The shapes are the ADAPTER's own contract (families/comments.cjs, comment-write.cjs plan()), which stays the authority for the rest. A comment is identified
  // by text + author + quote TOGETHER (all three REQUIRED: `quote` may be '' for a comment with no quoted range). resolve_comment's flag is `resolved` (default true), NOT `solved`.
  // reply_comment / add_comment_on_match have NO `author` field at all (the author is the editor session identity -- an owner decision): additionalProperties:false makes it
  // UNKNOWN_FIELD; `target.author` on a reply IDENTIFIES the existing comment, it is not the author of the reply.
  resolve_comment: Object.freeze({
    type: 'object',
    properties: Object.freeze({
      intent: Object.freeze({ const: 'resolve_comment' }), id: ID,
      text: COMMENT_TEXT, author: COMMENT_AUTHOR, quote: COMMENT_QUOTE,
      resolved: Object.freeze({ type: 'boolean' }),
    }),
    required: Object.freeze(['intent', 'text', 'author', 'quote']),
    additionalProperties: false,
  }),
  reply_comment: Object.freeze({
    type: 'object',
    properties: Object.freeze({
      intent: Object.freeze({ const: 'reply_comment' }), id: ID,
      target: Object.freeze({
        type: 'object',
        properties: Object.freeze({ text: COMMENT_TEXT, author: COMMENT_AUTHOR, quote: COMMENT_QUOTE }),
        required: Object.freeze(['text', 'author', 'quote']),
        additionalProperties: false,
      }),
      text: COMMENT_TEXT,
    }),
    required: Object.freeze(['intent', 'target', 'text']),
    additionalProperties: false,
  }),
  add_comment_on_match: Object.freeze({
    type: 'object',
    properties: Object.freeze({
      intent: Object.freeze({ const: 'add_comment_on_match' }), id: ID,
      matchText: COMMENT_TEXT, text: COMMENT_TEXT,
      occurrence: Object.freeze({ type: 'integer', minimum: 1, maximum: 100000 }),
    }),
    required: Object.freeze(['intent', 'matchText', 'text']),
    additionalProperties: false,
  }),
  // GH #308 (routed + coded; LIVE-measured, test/live/docx-gh308-create-table-probe.live.cjs;
  // table-insert.cjs's own plan() is the authority for the rows/cols vs cells exact-size cross-check): insert_table.
  insert_table: Object.freeze({
    type: 'object',
    description: 'Inserts a new table right after afterParagraph. If afterParagraph is the document\'s last body paragraph, the engine cannot end the body on a table and writes one empty paragraph after the new table (measured; this is not an error).',
    properties: Object.freeze({
      intent: Object.freeze({ const: 'insert_table' }), id: ID, selector: TABLE_INSERT_TARGET,
      rows: TABLE_INSERT_DIM, cols: TABLE_INSERT_DIM, cells: TABLE_INSERT_CELLS,
    }),
    required: Object.freeze(['intent', 'selector', 'rows', 'cols']),
    additionalProperties: false,
  }),
  // GH #313 (routed + coded; LIVE-measured, test/live/docx-gh313-section-break-probe.live.cjs):
  // insert_section_break. Columns are NOT a field here -- set_page_setup addresses the new
  // section by {index, of} once the break exists.
  insert_section_break: Object.freeze({
    type: 'object',
    description: 'Inserts a new section break right after afterParagraph (nextPage starts the following content on a new page; continuous does not). To set columns or other page-setup fields for the new section, use set_page_setup afterward, addressing it by its new section index.',
    properties: Object.freeze({
      intent: Object.freeze({ const: 'insert_section_break' }), id: ID, selector: SECTION_BREAK_TARGET,
      type: SECTION_BREAK_TYPE,
    }),
    required: Object.freeze(['intent', 'selector', 'type']),
    additionalProperties: false,
  }),
  // GH #306 (routed + coded; LIVE-measured, test/live/docx-gh306-insert-image-probe.live.cjs,
  // commit 749a4f607): insert_image. GH #306b 2. resz adds the header/footer selector shape below.
  insert_image: Object.freeze({
    type: 'object',
    description: 'Inserts a new image. selector is { afterParagraph } (position \'after\' default puts it in a new paragraph right after afterParagraph; position \'end\' appends it to the end of afterParagraph\'s own paragraph) OR { section: {index, of}, part: \'header\'|\'footer\', variant?: \'default\'|\'title\'|\'even\', paragraph?: {index, of}, position?: \'start\'|\'end\' } (targets the part\'s own last paragraph, or the named one -- the part must already exist; a title/even part must also already be active (titlePage / evenAndOdd on), it is never activated; selector.position \'end\' [default] appends to that paragraph\'s existing runs, \'start\' inserts a fresh run with the image BEFORE them) OR { tableCell: { text, tableAnchor? } } (the top-level table cell whose full text is text, unique -- inside the one table that has a cell with exactly tableAnchor, if given; the image is appended inline to the cell\'s last paragraph; only a cell of plain paragraphs). source is a data:image/png or data:image/jpeg base64 URI (no external URL). widthEmu/heightEmu may round-trip ±1 EMU off the requested value (measured engine rounding). Optional placement (body target with position \'after\' only) makes the image floating: wrap square | topAndBottom | inFront | behind, horizontal { relativeTo: margin, align: left | center | right } or { relativeTo: page, offsetMm }, vertical { relativeTo: paragraph | page, offsetMm } (whole millimetres 0..400); without it the image is inline.',
    properties: Object.freeze({
      intent: Object.freeze({ const: 'insert_image' }), id: ID, selector: IMAGE_INSERT_TARGET,
      source: IMAGE_INSERT_SOURCE, widthEmu: IMAGE_INSERT_EMU, heightEmu: IMAGE_INSERT_EMU, position: IMAGE_INSERT_POSITION,
      // GH #306 part 4: the SAME floating placement insert_shape takes (image-insert.cjs reuses shape-insert.cjs planPlacement); the planner allows it only with the body target, position 'after'
      placement: SHAPE_INSERT_PLACEMENT,
    }),
    required: Object.freeze(['intent', 'selector', 'source', 'widthEmu', 'heightEmu']),
    additionalProperties: false,
  }),
  // GH #322 (routed + coded; LIVE-measured, test/live/docx-gh322-bulk-insert-probe.live.cjs,
  // commit 48afe5074): insert_paragraphs.
  insert_paragraphs: Object.freeze({
    type: 'object',
    description: 'Inserts N paragraphs (1..200) right after afterParagraph, in ONE call, in the given order. Each item may name an existing paragraph style and/or a format (a subset of format_text\'s own keys: bold, italic, underline, strike, fontFamily, fontSize, color, alignment, spaceBefore, spaceAfter). All-or-nothing: an unknown style name refuses the whole batch, nothing written.',
    properties: Object.freeze({
      intent: Object.freeze({ const: 'insert_paragraphs' }), id: ID, selector: PARAGRAPHS_INSERT_TARGET,
      paragraphs: PARAGRAPHS_INSERT_LIST,
    }),
    required: Object.freeze(['intent', 'selector', 'paragraphs']),
    additionalProperties: false,
  }),
  // GH #309 (routed + coded; LIVE-measured, test/live/docx-gh309-shape-probe.live.cjs + docx-gh309-shape-read-probe.live.cjs): insert_shape.
  insert_shape: Object.freeze({
    type: 'object',
    description: 'Inserts a new inline shape with its own text in a new paragraph right after afterParagraph. shape.type is one of rect, flowChartProcess, roundRect, ellipse, triangle, trapezoid, flowChartDecision, flowChartTerminator, rightArrow, chevron, homePlate (the measured presets); widthMm / heightMm are whole millimetres (1..400); fill / stroke are 6-digit hex colours (RRGGBB); text is one plain line of at most 500 characters. The stroke width is fixed (1.5 pt). Optional shape.placement makes it floating: wrap square | topAndBottom | inFront | behind, horizontal { relativeTo: margin, align: left | center | right } or { relativeTo: page, offsetMm }, vertical { relativeTo: paragraph | page, offsetMm } (whole millimetres 0..400); without it the shape is inline. Not offered: tight / through wrap, header / footer / table-cell targets, multi-paragraph shape text.',
    properties: Object.freeze({ intent: Object.freeze({ const: 'insert_shape' }), id: ID, selector: SHAPE_INSERT_TARGET, shape: SHAPE_INSERT_SHAPE }),
    required: Object.freeze(['intent', 'selector', 'shape']),
    additionalProperties: false,
  }),
  // GH #323 (routed + coded; LIVE-measured, test/live/docx-gh323-native-chart-probe.live.cjs + -chart-readback-and-types-probe + -chart-tojson-series-probe): insert_chart.
  insert_chart: Object.freeze({
    type: 'object',
    description: 'Inserts a new inline native chart, carrying its own data, in a new paragraph right after afterParagraph. chart.type is bar, barStacked, barStackedPercent, horizontalBar, horizontalBarStacked, area, areaStacked, pie or doughnut (the measured set; pie and doughnut take exactly one series); categories are the category labels; series is a list of { name, values } with one number per category; widthMm / heightMm are whole millimetres (10..400); title is optional; legend may be bottom. Not offered: line, scatter, stock and 3D charts, editing an existing chart, table-cell / header / footer targets, floating placement, number formats, colours.',
    properties: Object.freeze({ intent: Object.freeze({ const: 'insert_chart' }), id: ID, selector: CHART_INSERT_TARGET, chart: CHART_INSERT_CHART }),
    required: Object.freeze(['intent', 'selector', 'chart']),
    additionalProperties: false,
  }),
  // GH #296 (routed + coded; LIVE-measured, test/live/docx-gh296-image-replace-delete-probe.live.cjs, run p296-r1 6/6): delete_image / replace_image.
  delete_image: Object.freeze({
    type: 'object',
    description: 'Deletes one existing image -- body, header or footer (default / title / even). selector is { paragraph, ordinal? } or { index } (the index read_images returns, valid for one read only), plus an optional part (\'body\' or \'header|footer:<section>:<default|title|even>\', the read_images row\'s own part label) that must match or the step is refused. The paragraph text around the image stays.',
    properties: Object.freeze({ intent: Object.freeze({ const: 'delete_image' }), id: ID, selector: IMAGE_EDIT_TARGET }),
    required: Object.freeze(['intent', 'selector']),
    additionalProperties: false,
  }),
  replace_image: Object.freeze({
    type: 'object',
    description: 'Replaces the picture of one existing image -- body, header or footer (default / title / even) -- keeping its position and placement (inline, or floating with its wrap). selector as delete_image. source is a data:image/png or data:image/jpeg base64 URI (no external URL). widthEmu / heightEmu together set a new size; omit both to keep the current size (a requested size may round-trip ±1 EMU).',
    properties: Object.freeze({ intent: Object.freeze({ const: 'replace_image' }), id: ID, selector: IMAGE_EDIT_TARGET, source: IMAGE_INSERT_SOURCE, widthEmu: IMAGE_INSERT_EMU, heightEmu: IMAGE_INSERT_EMU }),
    required: Object.freeze(['intent', 'selector', 'source']),
    additionalProperties: false,
  }),
});
const releasedOnly = (all, what) => Object.freeze(Object.fromEntries(DOCX_WRITE_ALLOWLIST.map(i => {
  if (!(i in all)) throw new TypeError(`${what}: no entry for the allowlisted intent ${i}`);
  return [i, all[i]];
})));
const DOCX_STEP_SCHEMAS = releasedOnly(DOCX_STEP_SCHEMAS_ALL, 'DOCX_STEP_SCHEMAS');
// the selector must name ONE resolvable branch (fingerprint | heading | text | sectionHeading + paragraphOrdinal | GH #307 emptyDocument);
// a selector with only qualifiers (kind / before / precondition ...) would resolve to UNSUPPORTED_SELECTOR at execution -- refused here
// instead, pre-dispatch. Shared with format_text (below); field-contract.cjs's SELECTOR_KEYS is what actually keeps emptyDocument out of
// format_text's reach (its own selector names there would be refused field-unsupported, one step later, still 0 writes).
const rewriteSelectorBranch = sel => (sel.fingerprint !== undefined || sel.heading !== undefined || sel.text !== undefined
  || sel.emptyDocument === true || (sel.sectionHeading !== undefined && sel.paragraphOrdinal !== undefined));
// GH #320: insert_paragraph_before/after / delete_paragraph -- the selector must name ONE of fingerprint, text or
// sectionHeading + paragraphOrdinal (the same "at least one" contract rewriteSelectorBranch enforces, minus `heading`:
// out of scope for this release, GH #233's own measured anchor omissions above still apply to it here).
const paragraphSelectorBranch = sel => (sel.fingerprint !== undefined || sel.text !== undefined
  || (sel.sectionHeading !== undefined && sel.paragraphOrdinal !== undefined));
// GH #320: delete_paragraph's `before` / `after` are the surviving-context proof (GH #234) -- meaningful, and therefore
// required, ONLY on the `text` branch: the resolver ignores them on a fingerprint / sectionHeading+paragraphOrdinal
// selector (field-contract.cjs refuses them there as "would be ignored beside" that branch), so requiring them
// unconditionally would make those two branches impossible to ever satisfy (required by this schema, then refused
// by field-contract for being present).
const usesTextAnchor = sel => sel.fingerprint === undefined && !(sel.sectionHeading !== undefined && sel.paragraphOrdinal !== undefined);
function deleteSelectorProblem(s) {
  if (!paragraphSelectorBranch(s.selector)) return 'selector must name fingerprint, text or sectionHeading + paragraphOrdinal';
  if (usesTextAnchor(s.selector) && (typeof s.selector.before !== 'string' || typeof s.selector.after !== 'string')) {
    return 'selector requires before and after (the surviving-context proof) when it names text';
  }
  return null;
}
// GH #233: the inserted text is ONE paragraph with something visible in it -- a blank text is refused by the planner anyway; a line break is
// NOT measured on the real editor (AddText of '\n' may split, drop or keep it), so it is refused here rather than guessed.
const insertTextProblem = s => (!paragraphSelectorBranch(s.selector) ? 'selector must name fingerprint, text or sectionHeading + paragraphOrdinal'
  : (!/\S/.test(s.text) ? 'text must contain a non-whitespace character'
  : (/[\r\n]/.test(s.text) ? 'text must be a single paragraph (no line break)' : null)));
const listTargetProblem = s => (/\S/.test(s.selector.text) ? null : 'selector.text must contain a non-whitespace character');
const moveSameParagraph = s => (s.selector.text === s.destinationSelector.text ? 'the moved paragraph and its destination are the same paragraph' : null);
// GH #232: a tab position and a border size are written on an integer grid -- twips (pos * 20) and eighth-points (size * 8), each
// ROUNDED by the adapter (formatting.cjs commandBody) -- and read back divided by the same factor, then compared within 0.01pt
// (formatting.cjs tabsEqual / borderSideEqual). A value that does not land on its grid (within that tolerance) would be written
// rounded, read back different, and end as a failed L1 AFTER the write was dispatched; refusing it here keeps it pre-dispatch.
// Tab positions must also be STRICTLY ASCENDING: tabsEqual compares index by index and whether the editor re-sorts or de-duplicates
// the list was NEVER measured, so the one order the read-back can prove regardless is sorted + unique.
const gridProblem = (v, factor, tolerancePt) => Math.abs(v * factor - Math.round(v * factor)) > tolerancePt * factor + 1e-9;
function formatGridProblem(format) {
  if (format.tabs !== undefined) {
    for (const [i, t] of format.tabs.entries()) {
      if (gridProblem(t.pos, 20, 0.01)) return `format.tabs[${i}].pos ${t.pos} is not on the 1/20-point (twip) grid the editor stores (it would be written rounded and read back different)`;
      if (i > 0 && !(t.pos > format.tabs[i - 1].pos)) return `format.tabs[${i}].pos must be strictly greater than the previous stop (ascending, no duplicates -- whether the editor re-orders the list is not measured)`;
    }
  }
  for (const k of ['borderTop', 'borderBottom', 'borderLeft', 'borderRight']) {
    if (format[k] !== undefined && gridProblem(format[k].size, 8, 0.01)) return `format.${k}.size ${format[k].size} is not on the 1/8-point grid the editor stores (it would be written rounded and read back different)`;
    // GH #314: space is NOT a grid (no tolerance makes sense) -- only 12 discrete point-values are
    // exactly reachable (the DS's own silent mm-rounding, formatting.cjs's own header comment); the
    // write side already translates through BORDER_SPACE_WRITE_INPUT, so anything outside this list
    // is refused here, enumerated, rather than silently landing on an unrequested value.
    if (format[k] !== undefined && format[k].space !== undefined && !BORDER_SPACE_POINTS.includes(format[k].space)) {
      return `format.${k}.space ${format[k].space} is not one of the only point-values the editor can store exactly: ${BORDER_SPACE_POINTS.join(', ')}`;
    }
  }
  return null;
}
// GH #264: the adapter's plan() (style-definition.cjs) refuses a blank name and an empty text / paragraph object; stated here pre-dispatch with the reason named, together with
// the float-safe grid guards (size 0.5 pt, indentLeft / spacingAfter 0.25 pt -- binary-exact, so the strict-equality L1 cannot disagree with the editor's rounding).
function styleProblem(s) {
  if (!/\S/.test(s.name)) return 'name must contain a non-whitespace character';
  if (s.basedOn !== undefined && !/\S/.test(s.basedOn)) return 'basedOn must contain a non-whitespace character';
  if (s.basedOn !== undefined && s.basedOn === s.name) return 'basedOn names the style itself';
  if (s.type === 'character' && s.paragraph !== undefined) return 'paragraph is not allowed on a character style (type character: text fields only)';
  for (const f of ['next', 'link']) if (s[f] !== undefined && !/\S/.test(s[f])) return `${f} must contain a non-whitespace character`;
  if (s.link !== undefined && s.link === s.name) return 'link names the style itself';
  if (s.text === undefined && s.paragraph === undefined && s.basedOn === undefined && s.next === undefined && s.link === undefined) return s.intent === 'define_style' ? 'define_style must name at least one of text, paragraph, basedOn' : 'modify_style must name at least one of text, paragraph, next, link';
  if (s.text !== undefined && Object.keys(s.text).length === 0) return 'text must name at least one of bold, size, color, italic, underline, strike, fontFamily, language';
  if (s.paragraph !== undefined && Object.keys(s.paragraph).length === 0) return 'paragraph must name at least one of align, indentLeft, spacingAfter, spacingBefore, indentRight, indentFirstLine, keepNext, keepLines, pageBreakBefore';
  if (s.text && s.text.fontFamily !== undefined && !/\S/.test(s.text.fontFamily)) return 'text.fontFamily must contain a non-whitespace character';
  if (s.text && s.text.size !== undefined && !Number.isInteger(s.text.size * 2)) return `text.size ${s.text.size} is not on the 0.5-point grid`;
  for (const k of ['indentLeft', 'spacingAfter', 'spacingBefore', 'indentRight', 'indentFirstLine']) {
    if (s.paragraph && s.paragraph[k] !== undefined && !Number.isInteger(s.paragraph[k] * 4)) return `paragraph.${k} ${s.paragraph[k]} is not on the 0.25-point grid`;
  }
  return null;
}
/** add_link / add_bookmark: the parent text and a matchText must carry something visible. */
const richProblem = s => (!s.selector.text.trim() ? 'selector.text must not be blank' : (s.matchText !== undefined && !s.matchText.trim() ? 'matchText must not be blank' : null));
const DOCX_CONSTRAINTS_ALL = Object.freeze({
  rewrite_paragraph: s => (!rewriteSelectorBranch(s.selector) ? 'selector must name fingerprint, heading, text, sectionHeading + paragraphOrdinal or emptyDocument'
    : (s.expectedPreservation !== undefined && Object.keys(s.expectedPreservation).length === 0 ? 'expectedPreservation must name at least one property' : null)),
  insert_paragraph_before: insertTextProblem,
  insert_paragraph_after: insertTextProblem,
  delete_paragraph: deleteSelectorProblem,
  move_paragraph_before: moveSameParagraph,
  move_paragraph_after: moveSameParagraph,
  apply_list: listTargetProblem,
  set_list_level: listTargetProblem,
  indent_list_item: listTargetProblem,
  outdent_list_item: listTargetProblem,
  remove_list: listTargetProblem,
  // the planner (table.cjs plan(), run on the step alone right after this) is the authority: a blank / tab / line-break text, a text equal to the cell
  // text, a tableAnchor equal to the old or the new text -- each REFUSED_SHAPE with its own detail (measured, docs/GH249-...)
  replace_table_cell: () => null,
  format_text: s => (!rewriteSelectorBranch(s.selector) ? 'selector must name fingerprint, heading, text or sectionHeading + paragraphOrdinal'
    : (Object.keys(s.format).length === 0 ? 'format must name at least one property' : formatGridProblem(s.format))),
  // the planner (table-row.cjs plan()) is the authority: an anchor equal to the target cell's text, a terminalRowCount below 2 -- REFUSED_SHAPE with its own detail
  insert_table_row_before: () => null,
  insert_table_row_after: () => null,
  // the planner (table-row.cjs plan()) is the authority: an anchor equal to the target cell's text, a terminalRowCount below 1 -- REFUSED_SHAPE with its own detail
  delete_table_row: () => null,
  // the planner (table-column.cjs plan()) is the authority: an anchor equal to the target cell's text, a terminalColumnCount below 2 -- REFUSED_SHAPE with its own detail
  insert_table_column_before: () => null,
  insert_table_column_after: () => null,
  // the planner (table-column.cjs plan()) is the authority: an anchor equal to the target cell's text, a terminalColumnCount below 1 -- REFUSED_SHAPE with its own detail
  delete_table_column: () => null,
  // the planner (table-merge-split.cjs plan()) is the authority: equal corners, an anchor equal to a corner or to the merged text, a merged text equal to a corner, rows * cols < 2 -- REFUSED_SHAPE with its own detail
  merge_table_cells: () => null,
  // the planner (table-merge-split.cjs plan()) is the authority: a 1 x 1 split, an anchor equal to the target -- REFUSED_SHAPE with its own detail
  split_table_cell: () => null,
  // the planner (table-format.cjs plan()) is the authority for the rest; the style NAME is resolved against the document's table-style list at execution (table-style-not-found / -ambiguous / -list-unobservable).
  // A BLANK tableAnchor is refused here: the planner would carry it as '' (not as "no anchor") and the table scope would then match any table that holds an EMPTY cell.
  // At least one axis must be named (the adapter's plan() refuses the same case; stated here so the refusal is pre-dispatch with the reason named); an empty borders object names no edge.
  format_table: s => {
    const a = s.table && s.table.tableCell && s.table.tableCell.tableAnchor;
    if (typeof a === 'string' && a.trim() === '') return 'a tableAnchor must be a non-empty cell text (a blank one would name any table that has an empty cell)';
    if (['borders', 'shading', 'cellMargins', 'indent', 'width', 'style'].every(k => s[k] === undefined)) return 'format_table must name at least one of borders, shading, cellMargins, indent, width, style';
    return s.borders !== undefined && Object.keys(s.borders).length === 0 ? 'borders must name at least one edge' : null;
  },
  define_style: s => styleProblem(s),
  modify_style: s => styleProblem(s),
  add_link: s => richProblem(s),
  add_bookmark: s => richProblem(s),
  set_numbering: s => (!s.selector.text.trim() ? 'selector.text must not be blank' : (s.code === undefined && s.start === undefined && s.suffix === undefined ? 'set_numbering must name at least one of code, start, suffix' : null)),
  set_list_type: listTargetProblem,
  // GH #262/#263: same "at least one field" discipline as format_text above -- page-setup.cjs's own
  // plan() requires at least one of its FOUR fields; this mirrors that rule now that the schema exposes
  // all four.
  set_page_setup: s => (s.size === undefined && s.margins === undefined && s.columns === undefined && s.startPageNumber === undefined
    ? 'set_page_setup requires at least one of size, margins, columns, startPageNumber' : null),
  // GH #259: the planner (header-footer.cjs plan(), run on the step alone right after this) is the
  // authority for the tab/line-break shape check (measured, docs/GH259-DOCX-EDIT-HEADER-COMMON-DECISION.md)
  // -- same discipline as replace_table_cell above. GH #316: the same planner is also the authority
  // for `format`'s accepted-key subset and its validation (normalizeWanted, reused from formatting.cjs).
  edit_header: () => null,
  // GH #260: same discipline, mirrored onto footer.
  edit_footer: () => null,
  // GH #267 B: a non-blank target, and ONE paragraph of non-blank note text (a line break in AddText is not measured on the real editor, as insertTextProblem).
  add_footnote: s => listTargetProblem(s) || insertTextProblem(s),
  add_endnote: s => listTargetProblem(s) || insertTextProblem(s),
  // GH #268d: the planner (image.cjs plan(), run on the step alone right after this) is the
  // authority for the selector-shape (exactly one of index/paragraph) and the wrap-value (NEM-MÉRT
  // values refused) checks -- same discipline as replace_table_cell/edit_header above.
  set_image_size: () => null,
  set_image_wrap: () => null,
  // GH #269e: a blank text / author / matchText would identify (or write) nothing visible; the planners (comments.cjs, comment-write.cjs plan(), run on the step alone right after
  // this) stay the authority for the rest (an unknown key inside `target` is the schema's job, a malformed target object the planner's).
  resolve_comment: s => (!s.text.trim() ? 'text must not be blank' : (!s.author.trim() ? 'author must not be blank' : null)),
  reply_comment: s => (!s.text.trim() ? 'text must not be blank' : (!s.target.text.trim() ? 'target.text must not be blank' : (!s.target.author.trim() ? 'target.author must not be blank' : null))),
  add_comment_on_match: s => (!s.matchText.trim() ? 'matchText must not be blank' : (!s.text.trim() ? 'text must not be blank' : null)),
  // GH #308: the planner (table-insert.cjs plan(), run on the step alone right after this) is the
  // authority for the blank-anchor and the cells-exact-size-match checks -- same discipline as
  // replace_table_cell/edit_header above.
  insert_table: () => null,
  // GH #313: the planner (section-break-insert.cjs plan(), run on the step alone right after
  // this) is the authority for the blank-anchor check; the schema's enum already refuses an
  // unmeasured type -- same discipline as insert_table above.
  insert_section_break: () => null,
  // GH #306: the planner (image-insert.cjs plan(), run on the step alone right after this) is the
  // authority for the blank-anchor check; the schema's pattern/enum already refuse a malformed
  // source and an unmeasured position -- same discipline as insert_table/insert_section_break above.
  insert_image: () => null,
  insert_paragraphs: () => null,
  // GH #309: the planner (shape-insert.cjs plan(), run on the step alone right after this) is the authority for the blank-anchor / text checks; the schema already refuses an unmeasured preset,
  // a non-integer size and a malformed colour -- same discipline as insert_image above.
  insert_shape: () => null,
  // GH #323: the planner (chart-insert.cjs plan()) is the authority for the blank-anchor, one-value-per-category, single-series and unique-name checks; the schema already refuses an
  // unmeasured type, a non-integer size and a non-number value -- same discipline as insert_shape above.
  insert_chart: () => null,
  // GH #296: the planner (image.cjs plan()) is the authority for the index/paragraph XOR and the widthEmu+heightEmu pairing; the schema already refuses a malformed part and a non-data: source.
  delete_image: () => null,
  replace_image: () => null,
});
const DOCX_CONSTRAINTS = releasedOnly(DOCX_CONSTRAINTS_ALL, 'DOCX_CONSTRAINTS');
const INSERT_INTENTS = Object.freeze(['insert_paragraph_before', 'insert_paragraph_after']);
/** A delete_paragraph / move_paragraph_* / list / replace_table_cell / insert_table_row_* / delete_table_row / insert_table_column_* / delete_table_column / merge_table_cells / split_table_cell / format_table task is that ONE step. Its proof is a property of the FINAL document (a delete: the named
 * neighbours adjacent; a move: the moved paragraph directly on its side of the destination; a list step: the item's type and level, on a
 * numbering whose levels other items share), so any other step in the task can move what it names: measured -- a delete whose neighbours a
 * later insert separates, two moves onto the same destination, a destination moved by a later step are each APPLIED and saved, then the task
 * reports ok:false; and a list step the adapter CANNOT prove (a set_list_type that would re-template a level another item sits on) is
 * refused as uncertain on its own, but next to a verified step the save of that verified step carries the unproven change with it.
 * Several such changes are several calls, each of them retryable on its own. */
const SOLO_INTENTS = Object.freeze(['delete_paragraph', 'move_paragraph_before', 'move_paragraph_after', 'apply_list', 'set_list_level', 'set_list_type', 'remove_list', 'replace_table_cell', 'insert_table_row_before', 'insert_table_row_after', 'delete_table_row', 'insert_table_column_before', 'insert_table_column_after', 'delete_table_column', 'merge_table_cells', 'split_table_cell', 'format_table', 'define_style', 'modify_style', 'add_link', 'add_bookmark', 'set_numbering']);
const LIST_ANCHORS = Object.freeze(['apply_list', 'set_list_level']);
const LIST_RELATIVES = Object.freeze(['indent_list_item', 'outdent_list_item']);
/** The one released multi-step list shape: [absolute anchor, relative step] on the SAME text. null = not that shape. */
const anchorThenRelative = steps => (steps.length === 2 && LIST_ANCHORS.includes(steps[0].intent) && LIST_RELATIVES.includes(steps[1].intent)
  && steps[0].selector.text === steps[1].selector.text ? { anchor: steps[0], relative: steps[1] } : null);
function deleteNotAlone(steps) {
  const rel = steps.findIndex(s => LIST_RELATIVES.includes(s.intent));
  if (rel >= 0) {
    const pair = anchorThenRelative(steps);
    if (!pair) return `step ${rel}: ${steps[rel].intent} needs an absolute anchor -- its task must be exactly [apply_list | set_list_level, ${steps[rel].intent}] on the SAME text (a standalone relative step has no provable terminal level)`;
    const level = pair.anchor.level === undefined ? 0 : pair.anchor.level;
    return pair.relative.intent === 'indent_list_item' && level >= 8 ? 'step 1: indent_list_item beyond level 8 (the anchor already puts the item on level 8)' : null;
  }
  const i = steps.findIndex(s => SOLO_INTENTS.includes(s.intent));
  return i >= 0 && steps.length > 1 ? `step ${i}: ${steps[i].intent} must be the only step of its task (its proof is the final position, which any other step can change)` : null;
}
/** Two inserts on the same anchor and the SAME side in one task: the second pushes the first off its anchor, so the first's whole-task
 * adjacency check fails AFTER both were applied and saved (measured: ok:false over a persisted change). Refused pre-dispatch. */
function duplicateInsertSlot(steps) {
  const seen = new Set();
  for (const [i, s] of steps.entries()) {
    if (!INSERT_INTENTS.includes(s.intent)) continue;
    const slot = `${s.intent}\u0000${s.selector.text}`;
    if (seen.has(slot)) return `step ${i}: ${s.intent} on the same anchor as an earlier step -- the earlier insert would no longer sit directly on its anchor`;
    seen.add(slot);
  }
  return null;
}

/** The MCP inputSchema shape -- an intentionally UNSATISFIABLE `oneOf: []` while the allowlist is empty
 * (no step can ever validate against it); this tool is not registered on the MCP surface in this unit
 * (decision 3), so no caller ever sees this schema yet -- it exists for the same reason the xlsx one does:
 * a single source of truth for `createDocxWriteTool`'s `description`/`inputSchema`.
 * GH #76 MODEL 1 DOCX port (owner decision "A", the owner channel, 2026-10-19): `verifyStoredContent` is the
 * SAME field, SAME description shape as write-tools.cjs's own `inputSchema()` -- one source of truth for the
 * flag's wire shape, two tools. */
function inputSchema() {
  return {
    type: 'object',
    properties: {
      fileId: { type: 'string', pattern: FILE_ID.source, description: 'Nextcloud fileId of the target DOCX file' },
      task: { type: 'object', properties: { id: ID, steps: { type: 'array', minItems: 1, maxItems: MAX_STEPS,
        items: { oneOf: DOCX_WRITE_ALLOWLIST.map(i => DOCX_STEP_SCHEMAS[i]) } },
        verifyStoredContent: { type: 'boolean', description: 'after a confirmed save, verify the stored content via a fresh, independent read before returning (default false; adds ~11-13.5s)' } },
        required: ['steps'], additionalProperties: false },
    },
    required: ['fileId', 'task'],
    additionalProperties: false,
  };
}

/** GH #266: the product add_link carries `href`; the adapter's field is `url` (see the LINK_URL comment). GH #248: set_numbering gets the refuse-shared policy. Every other step passes through untouched. */
const toAdapterStep = s => {
  if (s.intent === 'set_numbering' || s.intent === 'set_list_type') return { ...s, sharedDefinition: 'refuse' }; // GH #248 / #236b: the product never writes a definition other items sit on
  if (s.intent !== 'add_link' || s.href === undefined) return s;
  const { href, ...rest } = s;
  return { ...rest, url: href };
};

// ---- preflight: whole-task validation before any dispatch (the docx mirror of validateWriteRequest) -----

/**
 * validateDocxWriteRequest(args) -> { ok:true, fileId, task } | refusal (never throws). Structurally
 * identical to write-tools.cjs's validateWriteRequest: TASK_SCHEMA_KEYS is shared, now UNFILTERED (GH #76
 * MODEL 1 DOCX port, owner decision "A", the owner channel, 2026-10-19) -- `verifyStoredContent` is a
 * recognised docx task field, the SAME boolean-type check write-tools.cjs's own validateWriteRequest uses.
 */
function validateDocxWriteRequest(args) {
  if (!args || typeof args !== 'object' || Array.isArray(args)) return refuse('INVALID_ARGS', 'arguments must be an object');
  const names = checkNamesDeep(args, 'args');
  if (names) return names;
  const extra = Object.keys(args).find(k => !['fileId', 'task'].includes(k));
  if (extra !== undefined) return refuse('UNKNOWN_FIELD', `unknown field: ${extra}`, { path: `args.${extra}` });
  if (typeof args.fileId !== 'string' || !FILE_ID.test(args.fileId)) return refuse('INVALID_ARGS', 'fileId must be a numeric string');
  const task = args.task;
  if (!task || typeof task !== 'object' || Array.isArray(task)) return refuse('MALFORMED_TASK', 'task must be an object { id?, steps }');
  const taskExtra = Object.keys(task).find(k => !TASK_SCHEMA_KEYS.includes(k));
  if (taskExtra !== undefined) return refuse('UNKNOWN_FIELD', `unknown task field: ${taskExtra}`, { path: `args.task.${taskExtra}` });
  if (task.verifyStoredContent !== undefined && typeof task.verifyStoredContent !== 'boolean') return refuse('MALFORMED_TASK', 'task.verifyStoredContent must be a boolean');
  if (task.id !== undefined && check(ID, task.id, 'args.task.id')) return refuse('MALFORMED_TASK', 'task.id must match ^[A-Za-z0-9_.-]{1,64}$');
  if (!Array.isArray(task.steps)) return refuse('MALFORMED_TASK', 'task.steps must be an array');
  if (task.steps.length === 0) return refuse('EMPTY_TASK', 'task.steps is empty');
  if (task.steps.length > MAX_STEPS) return refuse('LIMIT', `task has ${task.steps.length} steps, over the ${MAX_STEPS} release bound`);
  for (const [i, s] of task.steps.entries()) {
    if (!s || typeof s !== 'object' || Array.isArray(s)) return refuse('MALFORMED_TASK', `step ${i} is not an object`, { stepIndex: i });
    const cls = docxIntentRefusal(s.intent, i);
    if (cls) return cls;
    // unreachable while DOCX_WRITE_ALLOWLIST is empty (docxIntentRefusal above always refuses first) --
    // guarded anyway so a future child-issue that forgets to add a DOCX_STEP_SCHEMAS/DOCX_CONSTRAINTS
    // entry alongside its allowlist addition fails loudly in its own tests, not here with a TypeError.
    const schema = DOCX_STEP_SCHEMAS[s.intent];
    if (schema) { const bad = check(schema, s, `args.task.steps[${i}]`); if (bad) return refuse(bad.code, `step ${i}: ${s.intent} -- ${bad.reason}`, { stepIndex: i, intent: s.intent, path: bad.path }); }
    const constraintFn = DOCX_CONSTRAINTS[s.intent];
    const why = constraintFn ? constraintFn(s) : null;
    if (why) return refuse('REFUSED_SHAPE', `step ${i}: ${s.intent} -- ${why}`, { stepIndex: i, intent: s.intent });
    // the canonical planner on the step ALONE (pure) -- the same discipline write-tools.cjs's
    // validateWriteRequest uses for xlsx; unreachable today for the same reason as above.
    const alone = docxAdapter.planOperations({ steps: [toAdapterStep(s)] });
    if (!alone || !alone.ok) return refuse('REFUSED_SHAPE', `step ${i}: ${s.intent} -- ${(alone && alone.detail) || 'the planner refused the step'}`, { stepIndex: i, intent: s.intent });
  }
  const lone = deleteNotAlone(task.steps);
  if (lone) return refuse('REFUSED_SHAPE', lone);
  const slot = duplicateInsertSlot(task.steps);
  if (slot) return refuse('REFUSED_SHAPE', slot);
  // the canonical planner, pre-dispatch (pure: no session) -- the same call the executor makes in-session
  const adapterTask = { ...task, steps: task.steps.map(toAdapterStep) };
  const planned = docxAdapter.planOperations(adapterTask);
  if (!planned || !planned.ok) return refuse('REFUSED_COMPOSITION', (planned && planned.detail) || 'the planner refused the task', planned && planned.index !== undefined ? { stepIndex: planned.index } : {});
  const validity = validatePlan(planned.operations);
  if (!validity.ok) return refuse('REFUSED_COMPOSITION', `plan invalid: ${validity.problems.join('; ')}`);
  return Object.freeze({ ok: true, fileId: args.fileId, task: toExecutorTask(adapterTask), plannedOperations: planned.operations.length });
}

// ---- stored-content verify (GH #76 MODEL 1 DOCX port -- owner decision "A", the owner channel, 2026-10-19;
// docs/UNCERTAIN-READBACK-DESIGN.md) -----------------------------------------------------------------------

/**
 * expectedParagraphsFromSteps(task) -> { targets: [{text, bold?}, ...], notProven: [...] }. The docx mirror of
 * write-tools.cjs's `expectedCellsFromSteps`: the ONE comparable intent today is `rewrite_paragraph`
 * (DOCX_WRITE_ALLOWLIST) -- its own requested post-state is `{text: step.text}`, plus `bold` when the caller
 * named it in `expectedPreservation` (the ONE field `stored-content.cjs`'s comparator covers, per decision doc
 * section 4; every OTHER `expectedPreservation` key is simply not comparable here and is never claimed to be).
 * Any OTHER intent is named in `notProven`, never guessed at -- unreachable today while DOCX_WRITE_ALLOWLIST
 * holds only rewrite_paragraph, the SAME "unreachable, guarded anyway" discipline this file already uses for
 * DOCX_STEP_SCHEMAS/DOCX_CONSTRAINTS.
 */
function expectedParagraphsFromSteps(task) {
  const targets = []; const notProven = [];
  for (const step of (task && task.steps) || []) {
    if (step.intent !== 'rewrite_paragraph') { notProven.push(`${step.intent}-verify-not-implemented`); continue; }
    targets.push({ text: step.text, ...(step.expectedPreservation && step.expectedPreservation.bold !== undefined ? { bold: step.expectedPreservation.bold } : {}) });
  }
  return { targets, notProven };
}

/**
 * runDocxStoredContentVerify({ v, r, preKey, configSource, createDriver, loadPlaywright, gate, now, keyRotation })
 * -> the docx mirror of write-tools.cjs's `runStoredContentVerify`, byte-identical in SHAPE and STATUS SET
 * (VERIFIED | MISMATCH | UNCERTAIN | NOT_APPLICABLE -- decision doc section 3: the four-outcome model this
 * unit's own first draft proposed is WITHDRAWN; GH #76's status set is the one in force). Only the comparator
 * differs: `stored-content.cjs`'s `textRunDiffs`/`productParagraphOf` (paragraph text + the single run-level
 * `bold` field) instead of xlsx's cell-level `cellDiffs`/`productCells`, and the read is `read_formatting`
 * (docx) instead of `read_range` (xlsx). Same rule, same bound (`awaitKeyRotation`, imported unchanged), same
 * "never while held" sequencing -- see `createDocxWriteExecution`'s own call site for why.
 */
async function runDocxStoredContentVerify({ v, r, preKey, configSource, createDriver, loadPlaywright, gate, now = Date.now, keyRotation = {}, timeoutMs, commandTimeoutMs }) {
  if (v.task.verifyStoredContent !== true) return null;
  const preservationScope = (r.preservation && r.preservation.scope) || null;
  const inheritedNotProven = (r.preservation && r.preservation.notProven) || [];
  const unresolved = (status, failureReason, overheadMs, extraNotProven = []) =>
    ({ status, scope: 'requested-post-state', preservationScope, notProven: [...inheritedNotProven, ...extraNotProven], overheadMs, failureReason });
  // requested, but nothing was saved under a confirmed forceSave: there is no new stored version to verify.
  if (r.persistence.state !== 'FORCESAVE_CONFIRMED') return unresolved('NOT_APPLICABLE', `no-confirmed-save:${r.persistence.state || 'none'}`, 0);
  const t0 = now();
  const { targets, notProven: skipped } = expectedParagraphsFromSteps(v.task);
  if (targets.length === 0) return unresolved('UNCERTAIN', 'read-incomplete', now() - t0, skipped);
  const rotation = await awaitKeyRotation(configSource, v.fileId, preKey, keyRotation);
  if (!rotation.rotated) return unresolved('UNCERTAIN', 'key-not-rotated', now() - t0, skipped);
  const lease = gate.tryEnter();
  if (!lease) return unresolved('UNCERTAIN', 'open-failed', now() - t0, skipped);
  let units = null; let incomplete = false;
  try {
    const driver = createDriver('docx', { configSource, loadPlaywright, timeoutMs, commandTimeoutMs });
    lease.track(driver);
    const opened = (await driver.open({ fileId: v.fileId, operation: 'READ' })) || {};
    if (!opened.ok) return unresolved('UNCERTAIN', 'open-failed', now() - t0, skipped);
    try {
      const vr = validateReadRequest('docx', { fileId: v.fileId, operation: 'read_formatting' });
      if (!vr.ok) { incomplete = true; }
      else {
        const out = await readOnce(FORMATS.docx, driver, vr, 'docx');
        if (!out.ok) incomplete = true; else units = out.result.units;
      }
    } finally {
      await driver.close().catch(() => {});
    }
  } finally {
    lease.release();
  }
  const overheadMs = now() - t0;
  if (incomplete || units === null) return unresolved('UNCERTAIN', 'read-incomplete', overheadMs, skipped);
  const diffs = [];
  for (const t of targets) diffs.push(...textRunDiffs(t, productParagraphOf(units, t.text), 'product'));
  if (diffs.length) return { status: 'MISMATCH', scope: 'requested-post-state', preservationScope, notProven: [...inheritedNotProven, ...skipped, ...diffs.map(x => `stored-poststate-mismatch:${x}`)], overheadMs, failureReason: null };
  return { status: 'VERIFIED', scope: 'requested-post-state', preservationScope, notProven: [...inheritedNotProven, ...skipped], overheadMs, failureReason: null };
}

// ---- execution (the docx mirror of createWriteExecution) -------------------------------------------------

// GH #261 2/2 (RELEASED by #261c): the header-footer family's own resolve() (via
// docxAdapter.resolveTarget) already classifies a flag-off title/even write as safely
// activatable -- shapeGuard()'s activationGuard() (families/header-footer.cjs:146-159) proves
// every UNREQUESTED slot it would also switch is provably unchanged first, and the activation
// write itself (SetTitlePage/SetEvenAndOddHdrFtr, header-footer.cjs:306-308) goes through the
// SAME __eoB21Variant bridge call (commandBody, header-footer.cjs:279-286) the write-side
// re-check already uses for the already-active case. #261 (this file's prior state) refused the
// OFF -> ON transition here with a resolveTarget wrapper (withActivationRefused) pending a
// stored LIVE PASS for that bridge path -- #261b supplied it
// (test/live/docx-hf-activation-bridge-readback-probe.live.cjs, g261bv2: the bridge reports the
// new state immediately after activation, persists across save+reopen, and matches the raw
// internal getter exactly, on both the activated section and an untouched one). The wrapper is
// therefore removed: `docxAdapter` is passed to `withSession` unwrapped, and the adapter's own,
// owner-approved (M5) activation rule is the only gate left -- unchanged, since header-footer.cjs
// itself was never modified by #261 or #261c.

/**
 * createDocxWriteExecution({ configSource, browsersPath, gate, runtime, loadPlaywright, createDriver, reuse,
 * keyRotation }) -> execute(v). GH #220 1/2: the GH #100 warm-session path is wired the SAME way
 * write-tools.cjs's createWriteExecution wires it for xlsx -- `warmSessionOptions` (write-tools.cjs, exported
 * unchanged for this reuse, GH #220 design doc section 3) builds the `reuse` object `withSession` consumes;
 * absent (the default, `reuse: null`) or with a disabled registry, nothing below changes from the pre-#220
 * shape: no health/documentKey/release call, no extra config fetch, every call opens and closes its own
 * session exactly as before -- a caller that never passes `reuse` sees NO behaviour change.
 * GH #76 MODEL 1 DOCX port (owner decision "A", docs/UNCERTAIN-READBACK-DESIGN.md): `verifyStoredContent`
 * wired the SAME way write-tools.cjs's createWriteExecution wires it for xlsx -- `preKey` is fetched ONLY on a
 * verifying call (never an unconditional extra config fetch: `test/acceptance/gh220-docx-write-lifecycle.
 * test.cjs`'s own config-fetch-count assertions depend on this), a verifying call never offers or takes a warm
 * session (the SAME race xlsx's own comment names), and `runDocxStoredContentVerify` runs strictly AFTER the
 * write's own lease is released (never two sessions on the same file at once).
 * `gate`/`runtime` are shared, single process-wide limits with the xlsx path when the caller passes the
 * SAME objects (design doc section 5: "egyetlen folyamat-szintű korlát mindkét formátumra"); `reuse.registry`
 * is a THIRD object a future wiring unit may or may not share across formats -- see the design doc's own
 * section on the reuse-key safety of doing so.
 */
function createDocxWriteExecution({ configSource, browsersPath, gate = createSessionGate(), runtime = null, loadPlaywright = null, createDriver = productDriver, reuse = null, keyRotation = {}, timeoutMs, commandTimeoutMs } = {}) {
  const rt = runtime || (loadPlaywright ? { check: () => ({ ok: true }), load: loadPlaywright } : createBrowserRuntime({ browsersPath }));
  return async function execute(v) {
    const verifying = v.task.verifyStoredContent === true;
    const lease = gate.tryEnter();
    if (!lease) return Object.freeze({ ok: false, stage: 'ADMISSION', code: 'BUSY', reason: `an editor session is active (MAX_ACTIVE_SESSIONS_PER_PROCESS=${gate.max}); retry after it closes`, dispatched: false });
    let r; let preKey = null;
    try {
      const ready = rt.check();
      if (!ready.ok) return { ok: false, stage: 'RUNTIME', code: ready.code || 'BROWSER_RUNTIME_UNAVAILABLE', reason: ready.detail, dispatched: false };
      if (verifying && configSource && typeof configSource.getEditorConfig === 'function') {
        try { const cfg = await configSource.getEditorConfig({ fileId: v.fileId, operation: 'WRITE' }); preKey = cfg && cfg.document ? cfg.document.key : null; } catch { preKey = null; }
      }
      const driver = productWriteDriver(createDriver('docx', { configSource, loadPlaywright: rt.load, timeoutMs, commandTimeoutMs }));
      lease.track(driver);
      const warm = verifying ? null : warmSessionOptions(reuse, configSource, v.fileId, lease);
      const options = { driver, fileId: v.fileId, ...(warm ? { reuse: warm } : {}) };
      const final = await withSession(options, docxAdapter, createExecutor(v.task, { strategy: 'batch' }));
      r = mapExecutionResult(final, v, 'docx_execute_task');
    } finally {
      lease.release();
    }
    const sv = await runDocxStoredContentVerify({ v, r, preKey, configSource, createDriver, loadPlaywright: rt.load, gate, keyRotation, timeoutMs, commandTimeoutMs });
    return applyStoredContentVerify(r, sv);
  };
}

/** createDocxWriteTool({ execute, ...createDocxWriteExecution options }) -> { description, inputSchema, handler(args) }.
 * NOT wired into server-tools.cjs / tools.cjs EXPOSED_CAPABILITY_TOOLS in this unit (decision 3) -- calling
 * this function builds a working handler object, but nothing in the product today constructs or exposes it. */
function createDocxWriteTool({ execute = null, ...wiring } = {}) {
  const run = execute || createDocxWriteExecution(wiring);
  return {
    description: `DOCX WRITE task on a Nextcloud file as the calling requester (own credential via the broker): the canonical task { id?, steps }, run by the canonical executor. Supported intents: ${DOCX_WRITE_ALLOWLIST.join(', ') || '(none yet -- the initial release allowlist is empty)'}; any other intent or shape is refused before dispatch.`,
    inputSchema: inputSchema(),
    async handler(args) {
      const v = validateDocxWriteRequest(args);
      if (!v.ok) return v;
      return run(v);
    },
  };
}

module.exports = {
  docxRoutedIntents, docxIntentRefusal, DOCX_STEP_SCHEMAS, DOCX_CONSTRAINTS, DOCX_STEP_SCHEMAS_ALL, DOCX_CONSTRAINTS_ALL, inputSchema,
  validateDocxWriteRequest, createDocxWriteExecution, createDocxWriteTool,
  // GH #76 MODEL 1 DOCX port: exported for direct, pure/narrow testing separately from the full end-to-end
  // tool handler, the same reason write-tools.cjs exports its own xlsx equivalents.
  expectedParagraphsFromSteps, runDocxStoredContentVerify,
};
