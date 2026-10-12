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
 * `before` / `after` keep the source meaning exactly: a substring match against the 80-character
 * normalized neighbour snippet (source TG:146-147).
 */

const PHYSICAL_KEYS = Object.freeze(['index', 'position', 'offset', 'paragraphIndex']);

function sameHeadingPath(a, b) {
  const aa = Array.isArray(a) ? a : [];
  const bb = Array.isArray(b) ? b : [];
  return aa.length === bb.length && aa.every((x, i) => x.level === bb[i].level && norm(x.text) === norm(bb[i].text));
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
    if (selector.before) hits = hits.filter(b => norm(b.locator && b.locator.context.before).includes(norm(selector.before)));
    if (selector.after) hits = hits.filter(b => norm(b.locator && b.locator.context.after).includes(norm(selector.after)));
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

module.exports = { resolveStableTarget };
