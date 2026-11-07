'use strict';

const { norm } = require('../inventory.cjs');
const { resolveStableTarget } = require('../resolver.cjs');

/**
 * DOCX `table-insert` family (GH #308, NEW_CAPABILITY): insert_table.
 *
 * No source intent: the table families (table.cjs / table-row.cjs / table-column.cjs /
 * table-format.cjs / table-merge-split.cjs) all operate on an EXISTING table; none creates one
 * (CreateTable/InsertContent: 0 hits in the pinned source registry, measured against main
 * e1e8c39a938a).
 *
 * LIVE-measured (test/live/docx-gh308-create-table-probe.live.cjs, run
 * gh308-tableprobe-1793158850683-882e64):
 *   - `doc.Push(table)` (the known-working call, docx-row-probe.live.cjs's own SETUP) lands the
 *     table at the document's END -- wrong for a position-preserving insert. When the anchor was
 *     the LAST body element the engine additionally inserted a TRAILING EMPTY PARAGRAPH after the
 *     pushed table (measured: elementCount 10 before -> 13 after; element 12 is an empty
 *     paragraph following the pushed table at 11).
 *   - `doc.AddElement(anchorIndex + 1, table)` lands the table EXACTLY between the anchor
 *     paragraph and its existing next sibling; `doc.AddElement` exists on this build and returns a
 *     boolean. The anchor's own body INDEX cannot be found by an identity walk
 *     (`doc.GetElement(i) === anchor`): the wrapper is a NEW object per call and the identity walk
 *     never matched (measured). The index comes from the SAME text-matching walk every other DOCX
 *     write family already uses (paragraph-lookup.cjs's own convention), re-checked for uniqueness
 *     inside the sandbox.
 *   - `Api.CreateTable(cols, rows)` returns `null` (no throw) for `cols` or `rows` 0 or negative;
 *     1x1 works. The write therefore validates rows/cols at PLAN time, never at dispatch.
 *   - A fresh reopen's saved document.xml carries exactly the written `<w:tbl>` count, in the
 *     written position, and the product's own `read_tables` (docx_read) reports the new table's
 *     cell texts/formats/style -- the write is visible on the product's read surface, not only in
 *     the raw XML.
 *
 * Preservation placement: PRIMITIVE_INTEGRAL. Every OTHER body block stays unchanged, in sequence
 * and content, after the new table's own position; the ONLY tolerated extra, and only when the
 * anchor was the LAST body element before the write, is EXACTLY ONE empty paragraph immediately
 * after the new table (the measured engine behaviour named above -- the body cannot end on a
 * table). Any other shape, including that same trailing paragraph when the anchor was NOT last,
 * is L3 FAIL.
 *
 * Scope (the card's own MI NEM TARTOZIK BELE): no table-style/border write (table-format.cjs
 * already covers that AFTER a table exists), no image-in-cell (#306), no nested-table anchor -- an
 * anchor "inside a table cell" is not a body-level paragraph/heading and is refused the same way
 * resolveStableTarget already refuses it for every other family (`primary.blocks` is body-level
 * only; a cell's own text is never one of its entries, only the whole table's joined text is).
 */

const MAX_DIM = 50; // a schema sanity cap, not a measured limit -- the probe measured only the 0/negative boundary
const TARGET_KINDS = Object.freeze(['paragraph', 'heading']);

function plan(step, index) {
  if (step.intent !== 'insert_table') return { ok: false, detail: `step ${index}: ${step.intent} is not a table-insert intent` };
  const sel = step.selector;
  if (!sel || typeof sel !== 'object' || Object.keys(sel).sort().join(',') !== 'afterParagraph') {
    return { ok: false, detail: `step ${index}: insert_table requires a selector of exactly { afterParagraph }` };
  }
  if (typeof sel.afterParagraph !== 'string' || !norm(sel.afterParagraph)) {
    return { ok: false, detail: `step ${index}: insert_table selector.afterParagraph must be a non-empty string` };
  }
  if (!Number.isInteger(step.rows) || step.rows < 1 || step.rows > MAX_DIM) {
    return { ok: false, detail: `step ${index}: insert_table requires an integer rows between 1 and ${MAX_DIM}` };
  }
  if (!Number.isInteger(step.cols) || step.cols < 1 || step.cols > MAX_DIM) {
    return { ok: false, detail: `step ${index}: insert_table requires an integer cols between 1 and ${MAX_DIM}` };
  }
  let cells = null;
  if (step.cells !== undefined) {
    if (!Array.isArray(step.cells) || step.cells.length !== step.rows) {
      return { ok: false, detail: `step ${index}: insert_table cells, when given, must be an array of exactly rows (${step.rows}) rows` };
    }
    for (const row of step.cells) {
      if (!Array.isArray(row) || row.length !== step.cols || row.some(c => typeof c !== 'string' || /[\t\r\n]/.test(c))) {
        return { ok: false, detail: `step ${index}: insert_table cells, when given, must be rows of exactly cols (${step.cols}) strings, none holding a tab or line break` };
      }
    }
    cells = step.cells.map(row => row.map(c => c));
  }
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`,
      intent: step.intent,
      selector: Object.freeze({ afterParagraph: sel.afterParagraph }),
      preservationMode: 'PRIMITIVE_INTEGRAL',
      rows: step.rows,
      cols: step.cols,
      cells,
      ...(Array.isArray(step.dependsOn) ? { dependsOn: step.dependsOn } : {}),
    },
  };
}

// ---- identity ------------------------------------------------------------------------------------

const blocksOf = primary => (primary && primary.blocks) || [];

/** The unique body-level paragraph/heading named by `text`, never a table or list. */
function resolveAnchor(primary, text) {
  const r = resolveStableTarget(primary, { text });
  if (!r.ok) return r;
  if (!TARGET_KINDS.includes(r.target.kind)) {
    return { ok: false, target: null, selector: r.selector, ambiguity: 'UNSUPPORTED_TARGET_KIND', kind: r.target.kind };
  }
  return r;
}

/** The normalized cell texts this operation expects right after the anchor: its own `cells`, or every cell empty. */
function expectedCells(op) {
  if (op.cells) return op.cells.map(row => row.map(c => norm(c)));
  return Array.from({ length: op.rows }, () => Array.from({ length: op.cols }, () => ''));
}

/** Does the block right after `anchorIndex` already hold a table of this operation's own shape and cell texts? */
function matchingTableAfter(primary, anchorIndex, op) {
  const blocks = blocksOf(primary);
  const next = blocks[anchorIndex + 1];
  if (!next || next.kind !== 'table') return false;
  const rows = next.rows || [];
  if (rows.length !== op.rows || rows.some(r => r.length !== op.cols)) return false;
  const want = expectedCells(op);
  return rows.every((row, r) => row.every((cell, k) => norm(cell.text) === want[r][k]));
}

// ---- classify / resolve ----------------------------------------------------------------------------

function classify(primary, op) {
  const anchor = resolveAnchor(primary, op.selector.afterParagraph);
  if (!anchor.ok) {
    return { state: 'FAIL_CLOSED', evidence: { reason: `insert-table-anchor-${String(anchor.ambiguity).toLowerCase().replace(/_/g, '-')}` } };
  }
  const anchorIndex = blocksOf(primary).indexOf(anchor.target);
  if (matchingTableAfter(primary, anchorIndex, op)) {
    return { state: 'SATISFIED', evidence: { anchor: anchor.target.text, rows: op.rows, cols: op.cols, proof: 'fresh-live-table-immediately-after-anchor' } };
  }
  return { state: 'APPLY', evidence: { anchor: anchor.target.text } };
}

function resolve(primary, selector) {
  const sel = selector || {};
  const anchor = resolveAnchor(primary, sel.afterParagraph);
  if (!anchor.ok) return { ok: false, target: null, selector, ambiguity: anchor.ambiguity };
  return { ok: true, target: { kind: 'insert_table_anchor', anchorText: norm(anchor.target.text) }, selector, ambiguity: null };
}

function preDispatchCheck(op, target) {
  return target && target.kind === 'insert_table_anchor' && target.anchorText === norm(op.selector.afterParagraph)
    ? { ok: true }
    : { ok: false, dispatched: false, detail: 'insert-table-anchor-target-unresolved' };
}

// ---- lower --------------------------------------------------------------------------------------

function lower(op, target) {
  const meta = Object.freeze({ family: 'table-insert', action: 'insert', anchorText: target.anchorText, rows: op.rows, cols: op.cols, cells: op.cells });
  return { kind: 'callCommand', body: commandBody(meta), meta };
}

/**
 * One callCommand (GH #308 probe candidate B, LIVE PASS). Re-locates the anchor paragraph by the
 * same text-matching walk every other DOCX write family uses, re-checked for uniqueness inside the
 * sandbox (the fresh-read-before-write guarantee re-stated inside it, not only outside); the
 * anchor's body INDEX comes from that SAME loop -- an identity walk never matches a new-per-call
 * wrapper (measured). Then `Api.CreateTable(cols, rows)` (null on the 0/negative boundary, guarded)
 * + `doc.AddElement(index + 1, table)`. Any guard miss returns unchanged.
 */
function commandBody(meta) {
  return `
try {
  const m = ${JSON.stringify(meta)};
  const doc = Api.GetDocument();
  if (typeof doc.AddElement !== 'function') return;
  const textOpts = { Numbering: false };
  const lookupTextOf = p => String(p.GetText(textOpts) || '').replace(/[\\r\\n]+$/, '');
  const n = doc.GetElementsCount();
  let anchorIndex = -1;
  for (let i = 0; i < n; i += 1) {
    const e = doc.GetElement(i);
    if (e && typeof e.GetClassType === 'function' && e.GetClassType() === 'paragraph' && lookupTextOf(e) === m.anchorText) {
      if (anchorIndex >= 0) return; // ambiguous at dispatch time: refuse rather than guess
      anchorIndex = i;
    }
  }
  if (anchorIndex < 0) return;
  const t = Api.CreateTable(m.cols, m.rows);
  if (!t) return;
  if (m.cells) {
    for (let r = 0; r < m.rows; r += 1) {
      for (let c = 0; c < m.cols; c += 1) {
        const text = m.cells[r][c];
        if (!text) continue;
        const p = t.GetRow(r).GetCell(c).GetContent().GetElement(0);
        const run = Api.CreateRun();
        run.AddText(text);
        p.AddElement(run);
      }
    }
  }
  doc.AddElement(anchorIndex + 1, t);
} catch (_) {}
`;
}

// ---- verify -------------------------------------------------------------------------------------

function verifyPrimitive(op, before, after) {
  const anchorBefore = resolveAnchor(before && before.primary, op.selector.afterParagraph);
  if (!anchorBefore.ok) return { ok: false, evidence: { reason: 'anchor not resolvable in the before-state' } };
  const anchorAfter = resolveAnchor(after && after.primary, op.selector.afterParagraph);
  if (!anchorAfter.ok) return { ok: false, evidence: { reason: 'anchor not resolvable in the after-state' } };
  const anchorIndexAfter = blocksOf(after.primary).indexOf(anchorAfter.target);
  const ok = matchingTableAfter(after.primary, anchorIndexAfter, op);
  return { ok, evidence: { anchor: anchorBefore.target.text, rows: op.rows, cols: op.cols, anchorIndexAfter, tableImmediatelyAfterAnchor: ok } };
}

/**
 * Equivalent HTML-projection spellings of "no border" that the engine's own re-save flips the
 * ANCHOR paragraph's style between (LIVE-measured, test/live/docx-gh308-insert-table-product.live.cjs
 * ROUND1):
 *   mso-border-<side>-alt:none   <->   border-<side>:none
 * Same meaning, different CSS-alike spelling. This is scoped to the single anchor-style comparison
 * in verifyPreservation below -- normalizing it everywhere would also hide a REAL border change
 * (e.g. border-left:1px solid) on any other block.
 */
function normalizeAnchorBorderNoneAlt(style) {
  if (typeof style !== 'string') return style;
  return style.replace(/\bmso-border-(left|top|right|bottom)-alt\s*:\s*none\b/gi, 'border-$1:none');
}

/**
 * Canonical, order-sensitive key for a block (same fields constraints.cjs's own canonicalBlock
 * compares). `anchorStyleNormalized` applies the mso-border-*-alt:none equivalence above -- pass it
 * true ONLY for the single anchor block; every other block, and every other field of the anchor
 * block, still compares byte-exact.
 */
const canonicalBlock = (b, anchorStyleNormalized = false) => (b.kind === 'table'
  ? ['table', (b.rows || []).map(row => row.map(c => [norm(c.text), c.span, c.rowspan]))]
  : [b.kind, b.level || null, norm(b.text), anchorStyleNormalized ? normalizeAnchorBorderNoneAlt(b.style || null) : (b.style || null)]);

/**
 * L3: every block before the anchor (inclusive) and every block after the new table is unchanged,
 * in sequence and content. The ONLY tolerated extra is the measured trailing empty paragraph, and
 * only when the anchor was the last body element before the write -- any other shape (missing,
 * extra, or present when the anchor was NOT last) fails.
 */
function verifyPreservation(op, before, after) {
  const anchorBefore = resolveAnchor(before && before.primary, op.selector.afterParagraph);
  if (!anchorBefore.ok) return { ok: false, evidence: { reason: 'anchor not resolvable in the before-state' } };
  const beforeBlocks = blocksOf(before.primary);
  const afterBlocks = blocksOf(after.primary);
  const anchorIndex = beforeBlocks.indexOf(anchorBefore.target);
  const wasLast = anchorIndex === beforeBlocks.length - 1;
  const tableAt = anchorIndex + 1;
  const newTable = afterBlocks[tableAt];
  const tableOk = !!newTable && newTable.kind === 'table';
  const beforePrefix = beforeBlocks.slice(0, anchorIndex + 1);
  const afterPrefix = afterBlocks.slice(0, anchorIndex + 1);
  // the anchor is the LAST entry of each prefix slice -- only that one gets the border-alt equivalence
  const prefixKey = blocks => blocks.map((b, i) => canonicalBlock(b, i === blocks.length - 1));
  const prefixSame = JSON.stringify(prefixKey(beforePrefix)) === JSON.stringify(prefixKey(afterPrefix));
  const restBefore = beforeBlocks.slice(anchorIndex + 1);
  const restAfter = afterBlocks.slice(tableAt + 1, tableAt + 1 + restBefore.length);
  const restSame = JSON.stringify(restBefore.map(b => canonicalBlock(b))) === JSON.stringify(restAfter.map(b => canonicalBlock(b)));
  const tail = afterBlocks.slice(tableAt + 1 + restBefore.length);
  const tailOk = wasLast
    ? tail.length === 1 && tail[0].kind === 'paragraph' && norm(tail[0].text) === ''
    : tail.length === 0;
  const lengthOk = afterBlocks.length === beforeBlocks.length + (wasLast ? 2 : 1);
  const ok = tableOk && prefixSame && restSame && tailOk && lengthOk;
  return { ok, evidence: { wasLast, anchorIndex, tableOk, prefixSame, restSame, tailOk, tailLength: tail.length, lengthOk } };
}

function wholeTaskChecks(op, finalObservation) {
  const anchor = resolveAnchor(finalObservation.primary, op.selector.afterParagraph);
  const anchorIndex = anchor.ok ? blocksOf(finalObservation.primary).indexOf(anchor.target) : -1;
  const ok = anchor.ok && matchingTableAfter(finalObservation.primary, anchorIndex, op);
  return [{ name: `${op.id}: a ${op.rows}x${op.cols} table sits immediately after "${op.selector.afterParagraph}"`, ok }];
}

// refs: none -- a table is not a text unit a later step can name
function producedText() {
  return null;
}

module.exports = {
  plan, classify, resolve, preDispatchCheck, lower, commandBody, verifyPrimitive, verifyPreservation, wholeTaskChecks, producedText,
  matchingTableAfter, expectedCells, resolveAnchor, canonicalBlock, normalizeAnchorBorderNoneAlt, MAX_DIM,
};
