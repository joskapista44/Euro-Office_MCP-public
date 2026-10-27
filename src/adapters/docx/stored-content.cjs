'use strict';

/**
 * GH #221 1/2 (DOCX P0) -- the DOCX port of GH #76 MODEL 1 (src/adapters/xlsx/stored-content.cjs):
 * the pure, shared comparison `test/live/stored-content-gate.cjs`'s format-neutral `classifyGate()`
 * calls through its `diffsOf` parameter. PURE: no I/O, never throws.
 *
 * `classifyGate` itself needs NO docx-specific port -- write-tools.cjs's `mapExecutionResult` (which
 * src/server/docx-write-tools.cjs's `createDocxWriteExecution` already reuses verbatim, `toolName:
 * 'docx_execute_task'`) produces the SAME receipt shape for docx as for xlsx: `{ok, operations:
 * [{id,status,...}], persistence:{state,...}, preservation:{scope,notProven}}`. `classifyGate`'s
 * write/key-rotation/verify-timeout/preservation-scope checks read exactly those fields, with no
 * format awareness anywhere in that logic -- confirmed by reading both call sites, not assumed.
 * The design doc (docs/GH221-DOCX-STORED-CONTENT-GATE-DESIGN.md) maps every one of the canonical
 * 10 steps; only the WRITE/VERIFY product calls and this module's own comparator are docx-specific.
 *
 * Scope of THIS unit: paragraph text + the single run-level `bold` field (the simplest, most
 * universal docx write target, and the one the required negative-regression case needs). Table-cell
 * and section/header-footer post-state comparators are NAMED, NOT IMPLEMENTED here -- see the design
 * doc's "what this unit does not cover" section.
 */

/** The SAME single-authority-label convention cellDiffs() uses: 'product' (docx_read) or 'stored'
 * (the saved file's own bytes, parsed independently of any editor session). */

/**
 * textRunDiffs(expected, actual, authority) -> [named difference, ...]
 *   expected  { text?, bold? } -- only the fields present are checked. `text` is also the ANCHOR
 *             the caller uses to find `actual` in the first place (the SAME convention every docx
 *             WRITE intent already uses to target a paragraph/cell -- `table: { text }`,
 *             `cellSelector`, etc.) -- this comparator never re-derives the anchor, it only
 *             compares the fields of an ALREADY-FOUND entity.
 *   actual    a UNIFORM shape on EITHER authority -- `{ text, bold }` | undefined (not found at
 *             all). The CALLER builds this uniform shape once per authority:
 *               'product' (docx_read read_formatting unit, or a read_tables cell + its format):
 *                 { text: unit.text, bold: unit.runs && unit.runs[0] && unit.runs[0].format ?
 *                 unit.runs[0].format.bold : null }
 *               'stored' (the saved file's word/document.xml, parsed by the harness's own
 *                 storedParagraphs()/storedTableCellText(), never a product read):
 *                 { text, bold }
 *             Collapsing both authorities to the SAME shape here (unlike cellDiffs, which keeps
 *             'product'/'stored' in different raw shapes and un-shapes them per-authority inside
 *             the comparator) is possible because docx's two authorities already agree on field
 *             names once the caller has done its own authority-specific extraction -- there is no
 *             XLSX-style sharedStrings/numFmt-id translation needed on the docx side for this one
 *             axis (bold is `<w:b/>` on `stored`, `format.bold` on `product` -- both already plain
 *             booleans by the time they reach here).
 *   NOT found: `actual === undefined` always means a GENUINE gap (the anchor text was not present
 *   on that authority at all) -- never guessed, never conflated with "found, but a field mismatch".
 */
function textRunDiffs(expected, actual, authority) {
  if (!actual) return [`${expected.text}:${authority}:missing`];
  const out = [];
  if ('text' in expected && actual.text !== expected.text) out.push(`${expected.text}:${authority}:text`);
  if ('bold' in expected && actual.bold !== expected.bold) out.push(`${expected.text}:${authority}:bold`);
  return out;
}

/** productParagraphOf(units, text) -> the uniform { text, bold } shape (see textRunDiffs's own doc)
 * for the FIRST read_formatting unit whose own `.text` exactly equals `text`, or undefined if none
 * does. `units` is docx_read's `read_formatting` result's own `units` array (paragraph/heading/
 * list_item entries, each already carrying `.text` and `.runs` -- read-tools.cjs's `unitsOf()`);
 * never re-derived from `read_text`'s `blocks` (a second read this comparator does not need: a
 * paragraph's OWN text is already on the `read_formatting` unit, same field name). A paragraph with
 * more than one run (a MIXED-format paragraph) has no single `bold` value -- `runs.length !== 1`
 * reports `bold: null` rather than guessing the first run's value for the whole paragraph; a gate
 * round that needs this axis therefore writes exactly one run covering the whole text (the SAME
 * discipline xlsx's own `runBoldDiffs` already documents for its one-run assumption). */
function productParagraphOf(units, text) {
  const u = (units || []).find(x => x && x.text === text);
  if (!u) return undefined;
  // an empty trailing run (text '') carries no text to format -- DS 9.3.5 leaves one after rewrite_paragraph (GH #221 LIVE)
  const runs = Array.isArray(u.runs) ? u.runs.filter(r => r && r.text !== '') : [];
  const bold = runs.length === 1 && runs[0].format ? runs[0].format.bold : null;
  return { text: u.text, bold };
}

module.exports = { textRunDiffs, productParagraphOf };
