'use strict';

const { norm, blockText } = require('./inventory.cjs');

/**
 * Stable semantic target resolution — EXTRACT_SEMANTICS from source docx-targeting.cjs:118-169.
 *
 * Always against the observation it is handed (the executor hands a FRESH one). A selector is a
 * semantic descriptor; it is never a stored position. Ambiguity and absence fail closed.
 *
 * Supported in the M4 slice: `fingerprint`, `heading` (+`level`), `text` (+`contains`, `kind`,
 * `headingPath`, `before`, `after`). `bookmark` and `tableCell` belong to families not migrated yet
 * and are refused explicitly rather than half-implemented.
 *
 * GH #307: `emptyDocument` (rewrite_paragraph only -- field-contract.cjs scopes it, this function does
 * not know which intent is calling it) resolves ONLY when the whole document is exactly one paragraph
 * whose NORMALIZED text is empty -- measured live (GH #307 probe): the editor's own live view reads
 * an "empty" paragraph as "\r\n", the product's read_text as " " (one space), the saved document.xml
 * as "" -- three different raw strings, the same normalized identity. A literal `=== ''` check would
 * therefore reject the SAME state most of the time; `norm()` is the only check that is stable across
 * all three. Any other shape (more than one block, a non-paragraph, or a non-empty paragraph) is
 * refused by name (`EMPTY_DOCUMENT_NOT_EMPTY`) rather than falling through to UNSUPPORTED_SELECTOR.
 *
 * `before` / `after` keep the source meaning exactly: a substring match against the 80-character
 * normalized neighbour snippet (source TG:146-147).
 */

const PHYSICAL_KEYS = Object.freeze(['index', 'position', 'offset', 'paragraphIndex']);

function sameHeadingPath(a, b) {
  const aa = Array.isArray(a) ? a : [];
  const bb = Array.isArray(b) ? b : [];
  return aa.length === bb.length && aa.every((x, i) => x.level === bb[i].level && norm(x.text) === norm(bb[i].text));
}

/**
 * GH #359: `before`/`after` are captured from the task's pre-task observation, but an EARLIER
 * operation of the same task may already have rewritten that neighbouring block by the time this
 * selector is resolved -- the live context then holds the neighbour's NEW text, not the one the
 * caller's anchor still names. `selector.beforeAlt`/`afterAlt` (planner metadata, adapter.cjs's
 * attachAnchorAlternates) carries that earlier operation's own produced text; a context match
 * against either value is accepted. The caller's own anchor stays the primary match (also accepted
 * when the earlier step turns out not to have run), so an unrelated anchor is still refused.
 */
function matchesAnchor(block, side, wanted, alt) {
  const context = norm(block.locator && block.locator.context[side]);
  return context.includes(norm(wanted)) || (typeof alt === 'string' && context.includes(norm(alt)));
}

function unique(hits, selector) {
  if (hits.length === 1) return { ok: true, target: hits[0], selector, ambiguity: null };
  return { ok: false, target: null, selector, ambiguity: hits.length === 0 ? 'NOT_FOUND' : 'AMBIGUOUS', candidates: hits.length };
}

function resolveStableTarget(primary, selector) {
  const blocks = (primary && primary.blocks) || [];
  if (!selector || typeof selector !== 'object') return { ok: false, target: null, selector, ambiguity: 'INVALID_SELECTOR' };
  // A physical position is never an identity (M3: ref != locator). A selector that carries one is
  // refused, not silently narrowed to its semantic fields.
  if (PHYSICAL_KEYS.some(k => selector[k] !== undefined)) {
    return { ok: false, target: null, selector, ambiguity: 'PHYSICAL_LOCATOR_REFUSED' };
  }
  if (selector.bookmark !== undefined || selector.tableCell !== undefined) {
    return { ok: false, target: null, selector, ambiguity: 'SELECTOR_FAMILY_NOT_MIGRATED' };
  }
  if (selector.emptyDocument !== undefined) return resolveEmptyDocument(blocks, selector);
  if (selector.sectionHeading !== undefined || selector.paragraphOrdinal !== undefined) return resolveSectionParagraph(blocks, selector);
  if (selector.fingerprint) return unique(blocks.filter(b => b.locator && b.locator.fingerprint === selector.fingerprint), selector);
  if (selector.heading) {
    const level = selector.level == null ? null : Number(selector.level);
    return unique(blocks.filter(b => b.kind === 'heading' && norm(b.text) === norm(selector.heading) && (level === null || b.level === level)), selector);
  }
  if (selector.text !== undefined) {
    const wanted = norm(selector.text);
    let hits = blocks.filter(b => blockText(b) === wanted || (selector.contains === true && blockText(b).includes(wanted)));
    if (selector.kind) hits = hits.filter(b => b.kind === selector.kind);
    if (selector.headingPath) hits = hits.filter(b => sameHeadingPath(b.locator && b.locator.headingPath, selector.headingPath));
    if (selector.before) hits = hits.filter(b => matchesAnchor(b, 'before', selector.before, selector.beforeAlt));
    if (selector.after) hits = hits.filter(b => matchesAnchor(b, 'after', selector.after, selector.afterAlt));
    return unique(hits, selector);
  }
  return { ok: false, target: null, selector, ambiguity: 'UNSUPPORTED_SELECTOR' };
}

/**
 * M5: `sectionHeading` + `paragraphOrdinal` (source docx-agent-preflight.cjs:52-64). The heading's
 * normalized text must be unique (no level filter); its section runs to the next heading of the same
 * or a higher level; the ordinal is 1-based over the PARAGRAPH blocks of that range (nested
 * subsections included; headings, lists and tables not counted; empty paragraphs counted). Both keys
 * are required. Stricter than the source: the ordinal must be an integer NUMBER (the source's JS API
 * also took "2").
 */
function resolveSectionParagraph(blocks, selector) {
  const h = selector.sectionHeading;
  const n = selector.paragraphOrdinal;
  if (typeof h !== 'string' || !h.trim() || n === undefined) return { ok: false, target: null, selector, ambiguity: 'SECTION_PARAGRAPH_SELECTOR_INCOMPLETE' };
  const hits = [];
  blocks.forEach((b, i) => { if (b.kind === 'heading' && norm(b.text) === norm(h)) hits.push(i); });
  if (hits.length !== 1) return { ok: false, target: null, selector, ambiguity: hits.length ? 'SECTION_AMBIGUOUS' : 'SECTION_NOT_FOUND', candidates: hits.length };
  if (!Number.isInteger(n) || n < 1) return { ok: false, target: null, selector, ambiguity: 'PARAGRAPH_ORDINAL_INVALID' };
  const start = hits[0];
  let end = blocks.length;
  for (let i = start + 1; i < blocks.length; i += 1) if (blocks[i].kind === 'heading' && blocks[i].level <= blocks[start].level) { end = i; break; }
  const paragraphs = blocks.slice(start + 1, end).filter(b => b.kind === 'paragraph');
  if (n > paragraphs.length) return { ok: false, target: null, selector, ambiguity: 'PARAGRAPH_NOT_FOUND', available: paragraphs.length, requested: n };
  return { ok: true, target: paragraphs[n - 1], selector, ambiguity: null };
}

/**
 * GH #307: `{ emptyDocument: true }`. Resolves only when the body is exactly one paragraph and its
 * NORMALIZED text is empty (see the header comment for why `norm()`, not `===''`). Any other shape
 * is refused by name, never UNSUPPORTED_SELECTOR -- the caller asked a real question ("is the
 * document empty?") and deserves a real answer, not "selector not understood".
 */
function resolveEmptyDocument(blocks, selector) {
  if (selector.emptyDocument !== true) return { ok: false, target: null, selector, ambiguity: 'EMPTY_DOCUMENT_SELECTOR_INVALID' };
  if (blocks.length === 1 && blocks[0].kind === 'paragraph' && norm(blocks[0].text) === '') {
    return { ok: true, target: blocks[0], selector, ambiguity: null };
  }
  return { ok: false, target: null, selector, ambiguity: 'EMPTY_DOCUMENT_NOT_EMPTY' };
}

module.exports = { resolveStableTarget };
