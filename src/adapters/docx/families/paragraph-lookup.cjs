'use strict';

/**
 * GH #311 / #312: the in-editor lookup of a text unit (a whole paragraph, or one exact text inside a paragraph) in
 * the document body, shared by the DOCX write families.
 *
 * The families used to find the unit with doc.Search(text) and require exactly one hit. Measured LIVE
 * (test/live/docx-gh311-312-paragraph-lookup-probe.live.cjs, docs/evidence/gh311-312-paragraph-lookup-probe-20261027.json):
 * doc.Search finds a text that also occurs inside another paragraph twice (#312) and a text holding a tab never
 * (#311), so neither could be written. Iterating the body paragraphs finds each exactly once:
 *   1. the ONE paragraph whose full text equals the unit -- even when other paragraphs contain it (#312);
 *   2. otherwise the ONE paragraph that contains the unit exactly once (a unit that is part of a paragraph, the
 *      case doc.Search used to cover); a unit found twice, or in two paragraphs, is not resolved.
 * The text is read with NewLineSeparator '\n' so a position counts one character per line break -- the convention
 * the rewrite already used for paragraph.GetRange(start, end).
 *
 * `unitRangeOf(t)` is the unit's own range: GetRange(start, start + length) -- the end is exclusive, and a whole
 * paragraph's range then covers exactly its text, not the paragraph mark (measured; GetRange() alone also formats
 * the mark). The returned range's text is checked against the unit.
 *
 * A hyperlink takes two range positions, one on each side of its text, that GetText does not count (measured LIVE,
 * test/live/docx-paragraph-lookup-hyperlink-shift-probe.live.cjs, docs/evidence/paragraph-lookup-hyperlink-shift-probe-20261030.json):
 * a unit after one link reads at text index + 2, after two at + 4, inside a link at + 1. `rangeAt(p, start, text)` maps a
 * text index to its range position over the paragraph's runs and hyperlinks (a hyperlink's text is GetDisplayedText -- it
 * has no GetText). A paragraph without a hyperlink keeps the plain index; one holding any other element kind gets no range.
 *
 * A unit that spans a run AND a hyperlink (text around a link -- every product intent passes a whole paragraph) gets a range
 * only when the caller asks for it, `unitRangeOf(t, true)`, and only when it covers each link it touches whole: measured LIVE
 * (W1 of the probe above), bold over such a range writes every run and keeps the link; a rewrite over it deletes the link (W2),
 * so the rewrite never asks. A unit ending or starting inside a link's text (a partial cover) is not measured: no range.
 *
 * Inlined into a command body: it expects `doc` in scope.
 */
const EXACT_PARAGRAPH_LOOKUP = `
  const lookupOpts = { Numbering: false, NewLineSeparator: '\\n' };
  const lookupTextOf = p => String(p.GetText(lookupOpts) || '').replace(/[\\r\\n]+$/, '');
  const unitOf = t => {
    if (!t) return null;
    const paragraphs = [];
    const n = doc.GetElementsCount();
    for (let i = 0; i < n; i += 1) {
      const e = doc.GetElement(i);
      if (e && typeof e.GetClassType === 'function' && e.GetClassType() === 'paragraph') paragraphs.push(e);
    }
    const exact = paragraphs.filter(p => lookupTextOf(p) === t);
    if (exact.length === 1) return { paragraph: exact[0], start: 0 };
    if (exact.length > 1) return null;
    let found = null;
    for (const p of paragraphs) {
      const s = lookupTextOf(p);
      const at = s.indexOf(t);
      if (at < 0) continue;
      if (found || s.indexOf(t, at + t.length) !== -1) return null;
      found = { paragraph: p, start: at };
    }
    return found;
  };
  const exactParagraphOf = t => { const u = unitOf(t); return u ? u.paragraph : null; };
  // [from, to] range positions of text[start, start + length), undefined when the paragraph holds no hyperlink, null when not mapped
  const linkedSpanOf = (p, start, length, crossLinks) => { try { return linkedSpanUnsafe(p, start, length, crossLinks); } catch (_) { return null; } };
  const elementsOf = p => {
    const els = [];
    if (typeof p.GetElementsCount === 'function') for (let j = 0; j < p.GetElementsCount(); j += 1) { const e = p.GetElement(j); els.push({ e, kind: e && typeof e.GetClassType === 'function' ? e.GetClassType() : null }); }
    return els;
  };
  const holdsHyperlink = p => elementsOf(p).some(x => x.kind === 'hyperlink');
  const linkedSpanUnsafe = (p, start, length, crossLinks) => {
    const els = elementsOf(p);
    if (!els.some(x => x.kind === 'hyperlink')) return undefined;
    const last = start + length - 1;
    let text = 0; let range = 0; let from = null; let first = -1; let partialStart = false;
    for (let j = 0; j < els.length; j += 1) {
      const { e, kind } = els[j];
      const read = kind === 'run' ? e.GetText : kind === 'hyperlink' ? e.GetDisplayedText : null;
      if (typeof read !== 'function') return null;
      const len = String(read.call(e, lookupOpts) || '').length;
      const pad = kind === 'hyperlink' ? 1 : 0;
      if (from === null && start < text + len) { from = range + pad + (start - text); first = j; partialStart = pad === 1 && start !== text; }
      if (from !== null && last < text + len) {
        const span = [from, range + pad + (last - text) + 1];
        if (first === j || !els.slice(first, j + 1).some(x => x.kind === 'hyperlink')) return span;
        return crossLinks && !partialStart && !(pad === 1 && last !== text + len - 1) ? span : null;
      }
      text += len; range += len + 2 * pad;
    }
    return null;
  };
  // GH #359 (2/2): measured LIVE (two plain runs, test/live/docx-5f96caad-rewrite8x-uncertain-lock-probe.live.cjs
  // MULTIRUN trap) that a run boundary the unit spans ALSO claims one GetRange position that GetText
  // does not count -- two runs needed +1 to the end. linkedSpanOf already prices a hyperlink's own
  // two-position padding; a paragraph with no hyperlink (span === undefined) can still cross plain-run
  // boundaries, so the end is widened by up to (run count - 1) until the slice matches exactly. Bounded
  // by the paragraph's own run count, and still fail-closed: a genuine text divergence matches at no
  // end and returns null, same as before this widening existed.
  const plainRunPadBudget = p => {
    const els = elementsOf(p);
    if (els.some(x => x.kind === 'hyperlink')) return 0;
    const runs = els.filter(x => x.kind === 'run').length;
    return runs > 1 ? runs - 1 : 0;
  };
  const rangeAt = (p, start, t, crossLinks) => {
    try {
      const span = linkedSpanOf(p, start, t.length, crossLinks);
      if (span === null) return null;
      if (span !== undefined) {
        const r = p.GetRange(span[0], span[1]);
        if (r && String(r.GetText(lookupOpts) || '') === t) return r;
        return null;
      }
      const baseEnd = start + t.length;
      const maxPad = plainRunPadBudget(p);
      for (let pad = 0; pad <= maxPad; pad += 1) {
        try { const r = p.GetRange(start, baseEnd + pad); if (r && String(r.GetText(lookupOpts) || '') === t) return r; } catch (_) {}
      }
    } catch (_) {}
    return null;
  };
  const unitRangeOf = (t, crossLinks = false) => {
    const u = unitOf(t);
    if (!u || typeof u.paragraph.GetRange !== 'function') return null;
    return rangeAt(u.paragraph, u.start, t, crossLinks);
  };`;

module.exports = { EXACT_PARAGRAPH_LOOKUP };
