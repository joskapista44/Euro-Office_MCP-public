'use strict';

const { norm, blockText } = require('../inventory.cjs');
const { resolveStableTarget } = require('../resolver.cjs');
const { canonicalBlock } = require('./table-insert.cjs');
const { normalizeWanted, compare } = require('./formatting.cjs');
const { EXACT_PARAGRAPH_LOOKUP } = require('./paragraph-lookup.cjs');

/**
 * DOCX `paragraphs-insert` family (GH #322, NEW_CAPABILITY): insert_paragraphs.
 *
 * A 6-page document needs ~230 steps (insert_paragraph_after + format_text per paragraph) through
 * the already-released per-paragraph intents, hitting MAX_STEPS=50 (write-tools.cjs) across 10+
 * task calls. Owner-side tervezesi dontes: NOT raising MAX_STEPS (that only lowers the CALL count,
 * not the per-step observe/verify cost) -- a new COMPOSITE intent instead: one anchor, N paragraphs,
 * ONE callCommand, verified as ONE unit.
 *
 * LIVE-measured (test/live/docx-gh322-bulk-insert-probe.live.cjs, commit 48afe5074):
 *   - `doc.AddElement(index, p)` at STRICTLY INCREASING indices (anchorIndex+1, anchorIndex+2, ...)
 *     lands N paragraphs in the requested order in ONE callCommand (the same mechanism GH #308/
 *     #306 already measured for a table/image, here for a run of plain body blocks); 50 paragraphs:
 *     24 ms JS-side; order and text survive a fresh reopen; the control paragraph is unaffected.
 *   - the SAME 50 paragraphs through the EXISTING, released product path (50 insert_paragraph_after
 *     steps, one task call at the MAX_STEPS ceiling): 38 773 ms total wall time -- three orders of
 *     magnitude slower than the single-callCommand route, which is the whole reason this intent
 *     exists.
 *   - 200 paragraphs in ONE callCommand: 36 ms, the session survives, the save lands.
 *
 * Anchor lookup: TEXT-based (`EXACT_PARAGRAPH_LOOKUP` + a text-matching index walk), never identity
 * -- the wrapper-identity class already measured three times (#308/#313/#306).
 *
 * Per-paragraph `format` is a SUBSET of format_text's own keys (`PARAGRAPHS_FORMAT_KEYS` below),
 * validated by the SAME `normalizeWanted` formatting.cjs's own format_text and header-footer.cjs's
 * edit_header/edit_footer already reuse (one constant, not a third copy). `style` is a SIBLING
 * field (an existing paragraph style name), not nested inside `format` -- the SAME split
 * insert_table's own `cells` vs. a future per-cell style would use.
 *
 * ALL-OR-NOTHING (the card's own explicit requirement): every requested `style` name is resolved
 * BEFORE any paragraph is inserted (`doc.GetStyle(name)` + the W-10-9 GetName()===name guard every
 * other family already uses) -- an unknown style refuses the WHOLE batch, nothing written. The
 * insertion loop itself is additionally wrapped so that an unexpected runtime failure mid-batch
 * rolls back every paragraph THIS call inserted (highest index first, so lower indices are never
 * shifted while removing) before returning unchanged -- a defensive second layer, since the
 * pre-check above already removes the only failure mode this build's API surface is known to have.
 *
 * Preservation placement: PRIMITIVE_INTEGRAL. The N new blocks (text + requested style/format, in
 * order, immediately after the anchor) are the ONLY new content; the anchor and everything before it
 * is unchanged; everything after the N new blocks is unchanged, in the SAME sequence (the SAME
 * block-sequence discipline `table-insert.cjs`'s own verifyPreservation already uses for its table,
 * `canonicalBlock` REUSED from there unchanged).
 *
 * Retry identity: SATISFIED when the N blocks immediately after the anchor already match every
 * requested paragraph's text AND (when requested) style/format, in order -- compared via
 * formatting.cjs's own exported `compare()` against each block's own ALREADY-ATTACHED `.format`
 * (inventory.cjs's `attachProperties`, the same evidence format_text's own verify already uses).
 * A same-length run with different text, or a different length, is APPLY, never a guessed match.
 *
 * Scope (the card's own MI NEM TARTOZIK BELE): no table/image in the batch (#308/#306 cover those
 * as their own intents), no MAX_STEPS change, no main-merge.
 */

const TARGET_KINDS = Object.freeze(['paragraph', 'heading']);
const MIN_PARAGRAPHS = 1;
const MAX_PARAGRAPHS = 200;
const PARAGRAPHS_FORMAT_KEYS = Object.freeze(['bold', 'italic', 'underline', 'strike', 'fontFamily', 'fontSize', 'color', 'alignment', 'spaceBefore', 'spaceAfter']);

function planParagraphItem(item, stepIndex, itemIndex) {
  const where = `step ${stepIndex}: paragraphs[${itemIndex}]`;
  if (!item || typeof item !== 'object' || Array.isArray(item)) return { ok: false, detail: `${where} must be an object` };
  if (typeof item.text !== 'string' || !norm(item.text)) return { ok: false, detail: `${where} requires a non-empty text` };
  if (/[\t\r\n]/.test(item.text)) return { ok: false, detail: `${where} text must not contain a tab or line break (one paragraph)` };
  let style;
  if (item.style !== undefined) {
    if (typeof item.style !== 'string' || !norm(item.style)) return { ok: false, detail: `${where} style must be a non-empty string` };
    style = item.style;
  }
  let format;
  if (item.format !== undefined) {
    if (!item.format || typeof item.format !== 'object' || Array.isArray(item.format)) return { ok: false, detail: `${where} format must be an object` };
    const unknown = Object.keys(item.format).filter(k => !PARAGRAPHS_FORMAT_KEYS.includes(k));
    if (unknown.length) return { ok: false, detail: `${where} format properties are limited to ${PARAGRAPHS_FORMAT_KEYS.join(', ')} -- unknown: ${unknown.join(', ')}` };
    const w = normalizeWanted(item.format);
    if (!w.ok) return { ok: false, detail: `${where}: ${w.detail}` };
    format = w.wanted;
  }
  return { ok: true, value: Object.freeze({ text: item.text, ...(style !== undefined ? { style } : {}), ...(format ? { format } : {}) }) };
}

function plan(step, index) {
  if (step.intent !== 'insert_paragraphs') return { ok: false, detail: `step ${index}: ${step.intent} is not an insert-paragraphs intent` };
  const sel = step.selector;
  if (!sel || typeof sel !== 'object' || Object.keys(sel).sort().join(',') !== 'afterParagraph') {
    return { ok: false, detail: `step ${index}: insert_paragraphs requires a selector of exactly { afterParagraph }` };
  }
  if (typeof sel.afterParagraph !== 'string' || !norm(sel.afterParagraph)) {
    return { ok: false, detail: `step ${index}: insert_paragraphs selector.afterParagraph must be a non-empty string` };
  }
  if (!Array.isArray(step.paragraphs) || step.paragraphs.length < MIN_PARAGRAPHS || step.paragraphs.length > MAX_PARAGRAPHS) {
    return { ok: false, detail: `step ${index}: insert_paragraphs requires paragraphs: an array of ${MIN_PARAGRAPHS}..${MAX_PARAGRAPHS} items` };
  }
  const paragraphs = [];
  for (let i = 0; i < step.paragraphs.length; i += 1) {
    const r = planParagraphItem(step.paragraphs[i], index, i);
    if (!r.ok) return r;
    paragraphs.push(r.value);
  }
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`, intent: step.intent, selector: Object.freeze({ afterParagraph: sel.afterParagraph }),
      preservationMode: 'PRIMITIVE_INTEGRAL', paragraphs: Object.freeze(paragraphs),
      ...(Array.isArray(step.dependsOn) ? { dependsOn: step.dependsOn } : {}),
    },
  };
}

// ---- the observation ------------------------------------------------------------------------------

const blocksOf = primary => (primary && primary.blocks) || [];

function resolveAnchor(primary, text) {
  const r = resolveStableTarget(primary, { text });
  if (!r.ok) return r;
  if (!TARGET_KINDS.includes(r.target.kind)) {
    return { ok: false, target: null, selector: r.selector, ambiguity: 'UNSUPPORTED_TARGET_KIND', kind: r.target.kind };
  }
  return r;
}

/** wantedOf(item) -- the merged compare() input: item.format plus item.style folded in under the SAME key `compare()` reads from a block's own `.format.style` (a NAME_KEY). */
const wantedOf = item => ({ ...(item.format || {}), ...(item.style !== undefined ? { style: item.style } : {}) });

/** A candidate block matches a requested paragraph when its text matches exactly, and -- only when style/format was requested -- its OWN attached `.format` already satisfies every requested key (compare(), never a guessed match). */
function blockMatches(block, item) {
  if (!block || block.kind === 'table' || blockText(block) !== norm(item.text)) return false;
  const wanted = wantedOf(item);
  if (Object.keys(wanted).length === 0) return true;
  const m = compare(block.format, wanted);
  return m.ok;
}

/** The N blocks immediately after anchorIndex match the requested sequence, in order, length-exact. */
function sequenceMatches(blocks, anchorIndex, paragraphs) {
  const n = paragraphs.length;
  for (let k = 0; k < n; k += 1) if (!blockMatches(blocks[anchorIndex + 1 + k], paragraphs[k])) return false;
  return true;
}

// ---- classify / resolve ----------------------------------------------------------------------------

function classify(primary, op) {
  const anchor = resolveAnchor(primary, op.selector.afterParagraph);
  if (!anchor.ok) {
    return { state: 'FAIL_CLOSED', evidence: { reason: `insert-paragraphs-anchor-${String(anchor.ambiguity).toLowerCase().replace(/_/g, '-')}` } };
  }
  const blocks = blocksOf(primary);
  const anchorIndex = blocks.indexOf(anchor.target);
  if (sequenceMatches(blocks, anchorIndex, op.paragraphs)) {
    return { state: 'SATISFIED', evidence: { anchor: anchor.target.text, count: op.paragraphs.length, proof: 'fresh-live-sequence-already-present' } };
  }
  return { state: 'APPLY', evidence: { anchor: anchor.target.text, count: op.paragraphs.length } };
}

function resolve(primary, selector) {
  const sel = selector || {};
  const anchor = resolveAnchor(primary, sel.afterParagraph);
  if (!anchor.ok) return { ok: false, target: null, selector, ambiguity: anchor.ambiguity };
  return { ok: true, target: { kind: 'insert_paragraphs_anchor', anchorText: norm(anchor.target.text) }, selector, ambiguity: null };
}

function preDispatchCheck(op, target) {
  return target && target.kind === 'insert_paragraphs_anchor' && target.anchorText === norm(op.selector.afterParagraph)
    ? { ok: true }
    : { ok: false, dispatched: false, detail: 'insert-paragraphs-anchor-target-unresolved' };
}

// ---- lower --------------------------------------------------------------------------------------

function lower(op, target) {
  const meta = Object.freeze({ family: 'paragraphs-insert', action: 'insert', anchorText: target.anchorText, paragraphs: op.paragraphs });
  return { kind: 'callCommand', body: commandBody(meta), meta };
}

/**
 * One callCommand (GH #322 probe, commit 48afe5074 -- route A/C LIVE PASS). Re-locates the anchor
 * by the SAME text-matching walk every other DOCX write family uses. ALL-OR-NOTHING: every
 * requested style name is resolved BEFORE any paragraph is inserted (the WHOLE batch refuses if one
 * does not resolve); the insertion loop is wrapped so an unexpected runtime failure mid-batch rolls
 * back every paragraph THIS call inserted (highest index first) before returning unchanged. Text
 * properties go through the new paragraph's OWN range (GetRange(0, text.length), never the bare
 * run -- the measured half-points-vs-points difference formatting.cjs's own header already states);
 * paragraph properties through GetParaPr(), the SAME split format_text's own commandBody uses.
 */
function commandBody(meta) {
  return `
try {
  const m = ${JSON.stringify(meta)};
  const doc = Api.GetDocument();
${EXACT_PARAGRAPH_LOOKUP}
  const anchor = unitOf(m.anchorText);
  if (!anchor) return;
  let anchorIndex = -1;
  const n0 = doc.GetElementsCount();
  for (let i = 0; i < n0; i += 1) {
    const e = doc.GetElement(i);
    if (e && typeof e.GetClassType === 'function' && e.GetClassType() === 'paragraph' && String(e.GetText({ Numbering: false }) || '').replace(/[\\r\\n]+$/, '') === m.anchorText) {
      if (anchorIndex >= 0) return; // ambiguous at dispatch time: refuse rather than guess
      anchorIndex = i;
    }
  }
  if (anchorIndex < 0) return;
  // ALL-OR-NOTHING pre-check: every requested style must resolve BEFORE any insertion.
  const styles = [];
  for (const item of m.paragraphs) {
    if (item.style === undefined) { styles.push(null); continue; }
    const s = doc.GetStyle(item.style);
    if (!s || typeof s.GetName !== 'function' || s.GetName() !== item.style) return; // refuse the whole batch
    styles.push(s);
  }
  const inserted = [];
  try {
    for (let k = 0; k < m.paragraphs.length; k += 1) {
      const item = m.paragraphs[k];
      const p = Api.CreateParagraph();
      p.AddText(item.text);
      doc.AddElement(anchorIndex + 1 + k, p);
      inserted.push(anchorIndex + 1 + k);
      const f = item.format || {};
      if (Object.keys(f).length) {
        const range = typeof p.GetRange === 'function' ? p.GetRange(0, item.text.length) : null;
        if (range) {
          if ('bold' in f) range.SetBold(f.bold);
          if ('italic' in f) range.SetItalic(f.italic);
          if ('underline' in f) range.SetUnderline(f.underline);
          if ('strike' in f) range.SetStrikeout(f.strike);
          if ('fontFamily' in f) range.SetFontFamily(f.fontFamily);
          if ('fontSize' in f) range.SetFontSize(f.fontSize);
          if ('color' in f) range.SetColor(Api.HexColor(f.color.replace('#', '')));
        }
        const paraPr = p.GetParaPr();
        if (paraPr) {
          if ('alignment' in f) paraPr.SetJc(f.alignment === 'justify' ? 'both' : f.alignment);
          if ('spaceBefore' in f) paraPr.SetSpacingBefore(f.spaceBefore * 20, false);
          if ('spaceAfter' in f) paraPr.SetSpacingAfter(f.spaceAfter * 20, false);
        }
      }
      if (styles[k]) p.SetStyle(styles[k]);
    }
  } catch (innerErr) {
    for (let j = inserted.length - 1; j >= 0; j -= 1) { try { doc.RemoveElement(inserted[j]); } catch (_) {} }
  }
} catch (_) {}
`;
}

// ---- verify -------------------------------------------------------------------------------------

function verifyPrimitive(op, before, after) {
  const anchorAfter = resolveAnchor(after && after.primary, op.selector.afterParagraph);
  if (!anchorAfter.ok) return { ok: false, evidence: { reason: 'anchor not resolvable in the after-state' } };
  const blocks = blocksOf(after.primary);
  const anchorIndex = blocks.indexOf(anchorAfter.target);
  const ok = sequenceMatches(blocks, anchorIndex, op.paragraphs);
  return { ok, evidence: { anchor: anchorAfter.target.text, count: op.paragraphs.length, anchorIndex } };
}

/**
 * PRIMITIVE_INTEGRAL: the N new blocks are exactly the requested sequence (above) immediately after
 * the anchor; the prefix (through the anchor) is byte-for-byte unchanged -- except the ANCHOR's own
 * style, compared under table-insert.cjs's measured mso-border-<side>-alt:none <-> border-<side>:none
 * equivalence (the SAME re-save flip, LIVE-measured again here on the gh271 representative document:
 * the only prefix difference was that spelling on the anchor); everything after the N new blocks is
 * unchanged, in the SAME sequence (canonicalBlock, reused from table-insert.cjs); the total block
 * count grows by exactly N.
 */
function verifyPreservation(op, before, after) {
  const anchorBefore = resolveAnchor(before && before.primary, op.selector.afterParagraph);
  if (!anchorBefore.ok) return { ok: false, evidence: { reason: 'anchor not resolvable in the before-state' } };
  const beforeBlocks = blocksOf(before.primary);
  const afterBlocks = blocksOf(after.primary);
  const anchorIndex = beforeBlocks.indexOf(anchorBefore.target);
  const n = op.paragraphs.length;

  const newBlocksOk = sequenceMatches(afterBlocks, anchorIndex, op.paragraphs);
  // the anchor is the LAST entry of each prefix slice -- only that one gets the border-alt equivalence
  const prefixKey = blocks => blocks.map((b, i) => canonicalBlock(b, i === blocks.length - 1));
  const beforePrefix = prefixKey(beforeBlocks.slice(0, anchorIndex + 1));
  const afterPrefix = prefixKey(afterBlocks.slice(0, anchorIndex + 1));
  const prefixSame = JSON.stringify(beforePrefix) === JSON.stringify(afterPrefix);
  const restBefore = beforeBlocks.slice(anchorIndex + 1).map(b => canonicalBlock(b));
  const restAfter = afterBlocks.slice(anchorIndex + 1 + n).map(b => canonicalBlock(b));
  const restSame = JSON.stringify(restBefore) === JSON.stringify(restAfter);
  const lengthOk = afterBlocks.length === beforeBlocks.length + n;
  const ok = newBlocksOk && prefixSame && restSame && lengthOk;
  return { ok, evidence: { newBlocksOk, prefixSame, restSame, lengthOk, count: n } };
}

function wholeTaskChecks(op, finalObservation) {
  const anchor = resolveAnchor(finalObservation.primary, op.selector.afterParagraph);
  let ok = false;
  if (anchor.ok) {
    const blocks = blocksOf(finalObservation.primary);
    ok = sequenceMatches(blocks, blocks.indexOf(anchor.target), op.paragraphs);
  }
  return [{ name: `${op.id}: ${op.paragraphs.length} paragraph(s) inserted immediately after "${op.selector.afterParagraph}"`, ok }];
}

// refs: none -- a batch of N paragraphs has no single semantic result a later step could ref
// unambiguously (the SAME choice table-insert.cjs/image-insert.cjs already make for their own
// multi-block/non-text writes).
function producedText() {
  return null;
}

module.exports = {
  plan, classify, resolve, preDispatchCheck, lower, commandBody, verifyPrimitive, verifyPreservation, wholeTaskChecks, producedText,
  resolveAnchor, blockMatches, sequenceMatches, wantedOf, PARAGRAPHS_FORMAT_KEYS, MIN_PARAGRAPHS, MAX_PARAGRAPHS,
};
