'use strict';

/**
 * GH #302 G4: a shape's text as the write takes it and as the read gives it back -- shared by insert_shape's optional
 * `text` and set_shape_text. The G4 LIVE design probe (docs/evidence/gh302-g4-shape-text-probe-20261026.json):
 *   - the text is written into the shape's GetContent(): the first paragraph is the content's own element 0 (AddText),
 *     every further one Api.CreateParagraph() + Push; an EMPTY paragraph stays (one, two, three with an empty middle: exact,
 *     session + reopen + the saved <a:p> runs)
 *   - RemoveAllElements() leaves ONE empty paragraph (the replace measured ['', 'new1', 'new2']), so a replacement writes
 *     into element 0 after it exactly like a new shape does
 *   - the product's read (drawing.text.read/v1, read_drawings `text`) gives every paragraph with '\r\n' and joins them
 *     with '|': ['a', 'b'] -> "a\r\n|b\r\n" -- unambiguous because a paragraph may hold no line break (refused here)
 */

const MAX_PARAGRAPHS = 10;
const MAX_CHARS = 1000;
const CONTROL = /[\u0000-\u001F\u007F]/;

/** why `text` is not a writable paragraph list, or null */
function textProblem(text) {
  if (!Array.isArray(text) || text.length < 1 || text.length > MAX_PARAGRAPHS) return `text must be a list of 1..${MAX_PARAGRAPHS} paragraphs`;
  if (!text.every(p => typeof p === 'string' && p.length <= MAX_CHARS)) return `every paragraph must be a string of at most ${MAX_CHARS} characters`;
  if (text.some(p => CONTROL.test(p))) return 'a paragraph may hold no line break or control character (one list item per paragraph)';
  if (text.every(p => p === '')) return 'text must hold at least one non-empty paragraph';
  return null;
}

/** the read's string for a paragraph list (the measured drawing.text.read/v1 shape) */
const encode = paras => paras.map(p => `${p}\r\n`).join('|');

/** the in-editor JS that writes `paras` into the content object named `c` (its element 0 must be the one empty paragraph) */
const writeJs = paras => `{ const __p = ${JSON.stringify(paras)}; c.GetElement(0).AddText(__p[0]); for (const x of __p.slice(1)) { const q = Api.CreateParagraph(); if (x !== '') q.AddText(x); c.Push(q); } }`;

module.exports = { MAX_PARAGRAPHS, MAX_CHARS, textProblem, encode, writeJs };
