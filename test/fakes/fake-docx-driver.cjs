'use strict';

const { internalGateOf, internalWriteRefusal, B21_BRIDGE_READ_RUNTIMES } = require('../../src/drivers/eurooffice/docx-driver.cjs');
// M8 runtime gate: the REAL driver's gate for script.runtime (default: word /3, where every DOCX path incl. the B18
// Bridge WRITE and the B21 Bridge READ run -- M11.5d.4 / M11.5e.4); script.gateOverride isolates ONE dependency (e.g.
// { numberingRemoveBridge: false }) so a test can reach that dependency's own refusal on the READ side, WITHOUT
// changing the EFFECTIVE runtime the second line (dispatch) sees -- see effectiveRuntime below.
const effectiveRuntime = script => (script.runtime === undefined ? B21_BRIDGE_READ_RUNTIMES[0] : script.runtime);
const gateOf = script => ({ ...internalGateOf(effectiveRuntime(script)), ...(script.gateOverride || {}) });

/**
 * A fake DOCX editor transport for the M4 rewrite slice.
 *
 * It separates DURABLE and LIVE state the way the W3.0 live probe measured the real editor:
 *   open        LIVE := copy of DURABLE
 *   dispatch    changes LIVE only
 *   forceSave   DURABLE := copy of LIVE   (canSave then reports false)
 *   close       LIVE is dropped; an unsaved change does NOT reach DURABLE
 *
 * It cannot run the editor API, so it applies the command's structured `meta` (the adapter's own
 * description of the same rewrite). The real driver executes `body` instead. Scripted faults act at
 * the transport boundary only.
 */

function htmlOf(paragraphs, images = [], doc = null) {
  let n = 0;
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  // Wave 2: consecutive list paragraphs of one numbering render as one <ul>/<ol> (measured live:
  // the editor projection splits a list where an item leaves it).
  let body = '';
  for (let i = 0; i < paragraphs.length; i += 1) {
    const p = paragraphs[i];
    if (p.list) {
      const tag = p.list.type === 'numbered' ? 'ol' : 'ul';
      let items = '';
      let j = i;
      // M5 gate 7: items of one numbering at levels of DIFFERENT types render as separate ul/ol runs
      // (list-level probe: P numbered@1 and its companion Q bullet@0, same numId, read as two lists)
      while (j < paragraphs.length && paragraphs[j].list && paragraphs[j].list.numId === p.list.numId && paragraphs[j].list.type === p.list.type) { items += `<li>${esc(paragraphs[j].text)}</li>`; j += 1; }
      body += `<${tag}>${items}</${tag}>`;
      i = j - 1;
    } else body += p.level ? `<h${p.level}>${inline(p)}</h${p.level}>` : `<p>${inline(p)}</p>`;
  }
  // Wave 4a: tables render after the body paragraphs; a cell's paragraphs are <p>s, a merged cell
  // carries colspan (table probe 1: the HTML joins a merged cell's paragraphs without a separator)
  // M5 (row 14): a grid-gap cell (`gap`, w:gridAfter) renders as the live projection does -- an EMPTY
  // `mso-cell-special:placeholder` td (gridAfter probe, run ...fcd2df); the driver read does not have it.
  // Its body is an &nbsp; paragraph -- the first LIVE run refused every case on a strictly-empty check (...ab094e)
  for (const t of (doc && doc.tables) || []) {
    body += `<table>${t.rows.map(r => `<tr>${r.filter(c => !c.vContinue).map(c => (c.gap ? '<td style="mso-cell-special:placeholder;border:none"><p><span>&nbsp;</span></p></td>' : `<td${c.span > 1 ? ` colspan="${c.span}"` : ''}${c.rowspan > 1 ? ` rowspan="${c.rowspan}"` : ''}>${c.paras.map(x => `<p><span>${esc(x.text)}</span></p>`).join('')}</td>`)).join('')}</tr>`).join('')}</table>`;
  }
  const imgs = images.map(i => `<p><img src="${esc(i.src)}" alt="${esc(i.alt || '')}"></p>`).join('');
  // Wave 3: footnote bodies render in a trailing list, as the live projection does (rich probe 3);
  // endnotes are not projected at all.
  const notes = (doc && doc.footnotes) || [];
  const list = notes.length ? `<div style="mso-element:footnote-list"><br clear="all"/><hr/>${notes.map((t, i) => `<div style="mso-element:footnote" id="ftn${i + 1}"><p class="MsoFootnoteText"><a style="mso-footnote-id:ftn${i + 1}" href="#_ftnref${i + 1}" name="_ftn${i + 1}"><span>[${i + 1}]</span></a><span>&nbsp;</span><span>${esc(t)}</span></p></div>`).join('')}</div>` : '';
  return `<html><body>${body}${imgs}${list}</body></html>`;

  // the paragraph's inline projection: a link wraps the text, a bookmark is a named anchor, each
  // footnote reference is an anchor carrying "[N]" (live shape, rich probe 3)
  function inline(p) {
    let t = esc(p.text);
    if (p.links && p.links.length) t = `<a href="${esc(p.links[0].url)}" title="${esc(p.links[0].tip)}">${t}</a>`;
    for (const b of (doc && doc.bookmarks) || []) if (b.text === p.text) t = `<a name="${esc(b.name)}">${t}</a>`;
    for (let i = 0; i < ((p.refs && p.refs.footnote) || 0); i += 1) { n += 1; t += `<a style="mso-footnote-id:ftn${n}" href="#_ftn${n}" name="_ftnref${n}"><span class="MsoFootnoteReference"><span>[${n}]</span></span></a>`; }
    return t;
  }
}

const TEXT_KEYS = ['bold', 'italic', 'underline', 'strike', 'fontFamily', 'fontSize', 'language', 'color', 'highlight', 'vertAlign'];
const pointSize = v => (Number(v) > 16 ? Number(v) / 2 : Number(v));

/** A paragraph's runs; a paragraph without explicit runs is one run carrying its `props`. */
const runsOf = p => (Array.isArray(p.runs) ? p.runs : [{ text: p.text, props: p.props || {} }]);

/** The effective run format the real driver reads from run.GetRange().GetTextPr(); unset -> absent. */
function runFormat(props) {
  const pr = props || {};
  const f = {};
  for (const k of ['bold', 'italic', 'underline', 'strike']) if (typeof pr[k] === 'boolean') f[k] = pr[k];
  if (pr.fontFamily) f.fontFamily = pr.fontFamily;
  if (pr.fontSize !== undefined && pr.fontSize !== null && Number.isFinite(Number(pr.fontSize))) f.fontSize = pointSize(pr.fontSize);
  if (pr.language) f.language = pr.language;
  if (pr.color) f.color = pr.color;
  if (pr.highlight) f.highlight = pr.highlight;
  if (pr.vertAlign) f.vertAlign = pr.vertAlign;
  return f;
}

/**
 * The whole-paragraph formatting the real driver reads (points, alignment name, style). A text
 * property that differs between runs is ABSENT, as the live Search-range GetTextPr returns
 * undefined for a mixed property (measured, W1.1 probe 2).
 */
function formatOf(p) {
  const pr = p.props || {};
  const base = {
    bold: !!pr.bold, italic: !!pr.italic, underline: !!pr.underline, strike: !!pr.strike,
    fontFamily: pr.fontFamily, fontSize: pointSize(pr.fontSize),
    language: pr.language, alignment: 'left', spaceBefore: 0, spaceAfter: 0, lineSpacing: 1, style: 'Normal',
    color: pr.color, highlight: pr.highlight, vertAlign: pr.vertAlign,
    // W-09: numeric getters (like spaceBefore/spaceAfter above) have a real baseline (no
    // indentation) even on an untouched paragraph. shading's "no colour set" sentinel is 'auto'
    // (the Builder's own IsAutoColor() terminology, measured) -- NOT absent/null: an absent key
    // would make `shading` unobservable (compare() treats undefined/null that way) and an
    // AUTO-shaded paragraph would then fail-closed on every first-time shading request instead of
    // correctly classifying as APPLY.
    indentLeft: 0, indentRight: 0, indentFirstLine: 0, shading: 'auto',
    // W-09-3: OOXML presence-only semantics (docx-driver.cjs's own new ParaPr.ToJSON() block) --
    // absence of <w:tabs>/<w:pBdr>/<w:keepNext> etc. on an untouched paragraph is a DEFINITE
    // "none"/"off", not unobservable (unlike style-definition.cjs's 'inherited'). `applyFormat`
    // below sets these as independent top-level keys, so SetTabs's measured REPLACE semantics and
    // each border side's measured ADD/independent semantics fall out of the existing per-key merge
    // with no extra fake-driver logic.
    tabs: [], borderTop: null, borderBottom: null, borderLeft: null, borderRight: null,
    keepNext: false, keepLines: false, pageBreakBefore: false, widowControl: false,
  };
  if (Array.isArray(p.runs)) {
    const fs = p.runs.map(r => runFormat(r.props));
    for (const k of TEXT_KEYS) {
      const vals = fs.map(f => f[k]);
      if (vals.every(v => v !== undefined && JSON.stringify(v) === JSON.stringify(vals[0]))) base[k] = vals[0];
      else delete base[k];
    }
  }
  return { ...base, ...(p.format || {}) };
}

function runPayload(p) {
  let at = 0;
  const runs = runsOf(p).map(r => { const x = { text: r.text, start: at, end: at + r.text.length, format: runFormat(r.props) }; at += r.text.length; return x; });
  return { runs, nonRunChildren: p.nonRunChildren || 0, runsCoverText: runs.map(r => r.text).join('') === p.text };
}

/** Set run properties on the characters [start, end) of `p`, re-merging equal neighbours. */
function setRangeProps(p, start, end, props) {
  const chars = [];
  for (const r of runsOf(p)) for (const ch of r.text) chars.push({ ch, props: { ...r.props } });
  for (let i = Math.max(0, start); i < Math.min(end, chars.length); i += 1) chars[i].props = { ...chars[i].props, ...props };
  const runs = [];
  for (const c of chars) {
    const last = runs[runs.length - 1];
    if (last && JSON.stringify(last.props) === JSON.stringify(c.props)) last.text += c.ch;
    else runs.push({ text: c.ch, props: c.props });
  }
  p.runs = runs;
}

const clone = paragraphs => paragraphs.map(p => ({ ...p, props: p.props ? { ...p.props } : p.props, format: p.format ? { ...p.format } : p.format, runs: Array.isArray(p.runs) ? p.runs.map(r => ({ text: r.text, props: { ...r.props } })) : p.runs, list: p.list ? { ...p.list } : p.list, links: p.links ? p.links.map(l => ({ ...l })) : p.links, refs: p.refs ? { ...p.refs } : p.refs }));
// Wave 3: the document-level state (bookmarks with their range text, note texts).
const cloneDoc = d => ({ bookmarks: (d.bookmarks || []).map(b => ({ ...b })), footnotes: [...(d.footnotes || [])], endnotes: [...(d.endnotes || [])],
  tables: (d.tables || []).map(t => ({ rows: t.rows.map(r => r.map(c => ({ ...c, paras: c.paras.map(x => ({ ...x, props: { ...(x.props || {}) } })) }))), ...(Array.isArray(t.gridCalc) ? { gridCalc: [...t.gridCalc] } : {}) })),
  // M5 gate 7: per numbering, the list type of each of its 9 levels (EO re-templates ONE level)
  numberings: Object.fromEntries(Object.entries(d.numberings || {}).map(([k, v]) => [k, [...v]])),
  // Wave 5: one entry per section; per kind (header/footer) and variant: null = no part of its own,
  // else { paras: [{ text, props, alignment }], nonPlain }. Absent -> one section with no parts.
  sections: (d.sections || [{}]).map(sec => ({ ...Object.fromEntries(['header', 'footer'].map(k => [k, Object.fromEntries(HF_VARIANTS.map(v => {
    const part = sec[k] && sec[k][v];
    return [v, part ? { nonPlain: !!part.nonPlain, paras: part.paras.map(x => ({ ...x, props: { ...(x.props || {}) } })) } : null];
  }))])), titlePage: sec.titlePage === true })),
  // M5: the document-wide even/odd flag (the section-level internal reader reports it per section)
  evenAndOdd: d.evenAndOdd === true });
const HF_VARIANTS = ['default', 'even', 'title'];

/** Wave 5: the driver's public read of one header/footer part (paragraph text ends in '\r\n', probes 1-2). */
function partRead(part) {
  if (!part) return null;
  const fmt = x => ({ ...runFormat(x.props), alignment: x.alignment || 'left', ...(x.style ? { style: x.style } : {}) });
  // M5 gate 6: the effective properties of the paragraph's one run, in the RAW getter units (the LIVE
  // read: bold/italic/underline/strike booleans, font family, size as GetFontSize returns it)
  const runs = x => [{ text: x.text, bold: !!(x.props || {}).bold, italic: !!(x.props || {}).italic, underline: !!(x.props || {}).underline, strike: !!(x.props || {}).strike,
    fontFamily: (x.props || {}).fontFamily, fontSize: (x.props || {}).fontSize }];
  return { paragraphs: part.paras.length, simple: !part.nonPlain, paras: part.paras.map(x => ({ text: `${x.text}\r\n`, format: fmt(x), runs: runs(x) })) };
}

/**
 * Wave 4a: the driver's public table read of one cell (table probes 1-2): the content text ends in
 * '\t' (a merged cell joins its paragraphs with '\r\n'); `simple` = only paragraphs made only of
 * runs; the effective properties of the FIRST paragraph (style name, alignment as read).
 */
function cellRead(c) {
  const p0 = c.paras[0] || { props: {} };
  const fmt = x => ({ ...runFormat(x.props), style: x.style || 'Normal', alignment: x.alignment || '1' });
  return { text: `${c.paras.map(x => x.text).join('\r\n')}\t`, paragraphs: c.paras.length, simple: !c.nonPlain,
    format: fmt(p0), paraFormats: c.paras.map(x => ({ text: x.text, format: fmt(x) })) };
}

function createStore(paragraphs, { images = [], doc = {} } = {}) {
  return { durable: clone(paragraphs), durableDoc: cloneDoc(doc), saves: 0, images };
}

function createFakeDocxDriver(store, script = {}) {
  const calls = { open: 0, close: 0, forceSave: 0, canSave: 0, readLive: 0, dispatch: 0, readPackage: 0 };
  let docxConfirmedAt = null; // M8: the dispatch count at the last confirmed forcesave (script.saveObserver only)
  let live = null;
  let liveDoc = null;
  let inNote = false;
  let dirty = false;

  // A fault acts on every dispatch, or only on the dispatch numbers listed in `faultOn` (1-based).
  const faultActive = () => !!script.fault && (!script.faultOn || script.faultOn.includes(calls.dispatch));
  const exact = t => live.filter(p => p.text === t);
  const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'en-US' };

  function applyRewrite({ oldText, newText }) {
    const hits = exact(oldText);
    if (hits.length !== 1) return; // the real command also returns without mutating
    if (script.fault === 'insert-only') { hits[0].text = `${newText}${oldText}`; dirty = true; return; }
    if (script.fault === 'ignore') return;
    hits[0].text = newText;
    if (script.fault === 'rewrite-into-list') hits[0].list = { type: 'bullet', level: 0, numId: 'undeclared' };
    // `dropPropsOn` limits the fault to the listed dispatch numbers (1-based); absent = every rewrite.
    if (script.fault === 'drop-props' && (!script.dropPropsOn || script.dropPropsOn.includes(calls.dispatch))) hits[0].props = { ...hits[0].props, bold: false, italic: false, language: 'en-US' };
    dirty = true;
  }

  /** The substring write (meta.matchText): the literal matchText exactly once inside the parent. */
  function applyRangeFormat({ targetText, matchText, format }) {
    const hits = exact(targetText);
    if (hits.length !== 1) return;
    const p = hits[0];
    if (p.text.split(matchText).length - 1 !== 1) return; // the real command also returns
    const f = faultActive() ? script.fault : null;
    if (f === 'ignore') return;
    const keys = f === 'format-partial' ? Object.keys(format).slice(0, 1) : Object.keys(format);
    const textProps = {};
    for (const k of keys) if (TEXT_KEYS.includes(k)) textProps[k] = k === 'fontSize' ? format[k] * 2 : format[k];
    const paraKeys = keys.filter(k => !TEXT_KEYS.includes(k));
    const start = p.text.indexOf(matchText);
    const other = live.find(x => x !== p && x.text.includes(matchText));
    if (f === 'range-paragraph-only') { p.format = { ...(p.format || {}), ...Object.fromEntries(keys.map(k => [k, format[k]])) }; dirty = true; return; }
    if (f === 'range-other-occurrence') { if (other) setRangeProps(other, other.text.indexOf(matchText), other.text.indexOf(matchText) + matchText.length, textProps); dirty = true; return; }
    setRangeProps(p, start, start + matchText.length + (f === 'range-spill' ? 1 : 0), textProps);
    if (f === 'range-also-other' && other) setRangeProps(other, other.text.indexOf(matchText), other.text.indexOf(matchText) + matchText.length, textProps);
    if (f === 'range-inside-other-prop') setRangeProps(p, start, start + matchText.length, { underline: !runFormat(runsOf(p)[0].props).underline });
    if (paraKeys.length) p.format = { ...(p.format || {}), ...Object.fromEntries(paraKeys.map(k => [k, format[k]])) };
    if (f === 'range-change-text') { p.text = `${p.text} (edited)`; p.runs = [...runsOf(p).slice(0, -1), { ...runsOf(p).slice(-1)[0], text: `${runsOf(p).slice(-1)[0].text} (edited)` }]; }
    dirty = true;
  }

  function applyFormat({ targetText, format }) {
    const hits = exact(targetText);
    if (hits.length !== 1) return;
    const f = faultActive() ? script.fault : null;
    if (f === 'ignore') return;
    const keys = Object.keys(format);
    const apply = f === 'format-partial' ? keys.slice(0, 1) : keys;
    const p = hits[0];
    const fmt = formatOf(p);
    for (const k of apply) fmt[k] = format[k];
    p.format = fmt;
    p.props = { ...p.props, ...Object.fromEntries(['bold', 'italic', 'underline', 'strike', 'fontFamily', 'language', 'color', 'highlight', 'vertAlign'].filter(k => k in fmt).map(k => [k, fmt[k]])) };
    if ('fontSize' in format && !(f === 'format-partial' && !apply.includes('fontSize'))) p.props.fontSize = format.fontSize * 2;
    if (Array.isArray(p.runs)) setRangeProps(p, 0, p.text.length, Object.fromEntries(apply.filter(k => TEXT_KEYS.includes(k)).map(k => [k, k === 'fontSize' ? format[k] * 2 : format[k]])));
    if (f === 'format-drop-other') { const i = live.findIndex(x => x !== p); if (i >= 0) live.splice(i, 1); }
    if (f === 'format-change-text') p.text = `${p.text} (edited)`;
    // W-09-3: a RED-proof fixture for the NEW sibling-format L3 check (otherParagraphsNewFieldsUnchanged)
    // -- another paragraph keeps its TEXT (the pre-existing L3 still passes) but its tabs are mutated as
    // an (unintended) side effect, the way a bug in the real driver's write could leak across paragraphs.
    if (f === 'format-other-new-field-side-effect') { const other = live.find(x => x !== p); if (other) other.format = { ...formatOf(other), tabs: [{ pos: 99, align: 'left' }] }; }
    dirty = true;
  }

  let numSeq = 100;
  /**
   * M5 gate 7: a numbering's type is PER LEVEL (list-level probe 2026-09-25). A fixture item seeds its
   * numbering (every level = its type, then its own level); an item's type is always its level's type.
   */
  function seedNumberings() {
    for (const p of live) if (p.list) { if (!liveDoc.numberings[p.list.numId]) liveDoc.numberings[p.list.numId] = Array(9).fill(p.list.type); }
    for (const p of live) if (p.list && p.list.type) liveDoc.numberings[p.list.numId][p.list.level] = liveDoc.numberings[p.list.numId][p.list.level] || p.list.type;
    syncTypes();
  }
  function syncTypes() { for (const p of live) if (p.list && liveDoc.numberings[p.list.numId] && p.list.level >= 0 && p.list.level <= 8) p.list.type = liveDoc.numberings[p.list.numId][p.list.level]; }
  function newNumbering(type) { let id; do { id = `n${numSeq += 1}`; } while (liveDoc.numberings[id]); liveDoc.numberings[id] = Array(9).fill(type); return id; }
  /** Wave 2: the list write (meta mirrors the callCommand's own inputs and precondition). */
  function applyList(m) {
    const hits = exact(m.targetText);
    if (hits.length !== 1) return;
    const p = hits[0];
    if (m.beforeType === 'normal' ? !!p.list : !p.list || p.list.level !== m.beforeLevel) return; // the real command also returns
    const f = faultActive() ? script.fault : null;
    if (f === 'ignore') return;
    const bump = f === 'list-wrong-level' ? 1 : 0;
    const a = m.listAction;
    if (a === 'apply_list') {
      if (m.continueFromText && f !== 'list-new-numbering') {
        const prev = exact(m.continueFromText);
        if (prev.length !== 1 || !prev[0].list) return;
        p.list = { type: prev[0].list.type, level: m.level + bump, numId: prev[0].list.numId };
      } else p.list = { type: m.type, level: m.level + bump, numId: newNumbering(m.type) };
    } else if (a === 'remove_list' || (a === 'outdent_list_item' && m.beforeLevel === 0)) {
      delete p.list;
    } else if (a === 'set_list_type') {
      // a numbering LEVEL template is shared: every item of this numbering at this level changes
      liveDoc.numberings[p.list.numId][m.beforeLevel] = m.type;
    } else {
      const next = a === 'set_list_level' ? m.level : a === 'indent_list_item' ? m.beforeLevel + 1 : m.beforeLevel - 1;
      p.list = { ...p.list, level: next + bump };
    }
    if (f === 'list-touch-other') { const o = live.find(x => x !== p && x.list); if (o) o.list = { ...o.list, level: o.list.level + 1 }; }
    // M5 gate 7: a write that re-templates ANOTHER, item-less level of the same numbering (collateral level change
    // that no list item can reveal -- only the per-level format read sees it)
    if (f === 'list-retemplate-other-level' && p.list) { const lv = liveDoc.numberings[p.list.numId]; const k = p.list.level === 5 ? 6 : 5; lv[k] = lv[k] === 'bullet' ? 'numbered' : 'bullet'; }
    syncTypes();
    // W2.1: the conversion also resets the target's run formatting (a lost earlier format / rewrite property)
    if (f === 'list-drop-format') { p.props = { ...p.props, bold: false, italic: false, language: 'en-US' }; if (Array.isArray(p.runs)) p.runs = [{ text: p.text, props: { ...p.props } }]; }
    if (f === 'list-change-text') p.text = `${p.text} (edited)`;
    dirty = true;
  }

  /** Wave 4a: the cell write (meta mirrors the callCommand: full-cell-text identity, plain-cell guard). */
  function applyCell({ targetText, tableAnchor, text }) {
    const n = x => String(x).replace(/\s+/g, ' ').trim();
    const cellText = c => n(c.paras.map(x => x.text).join('\r\n'));
    let tables = liveDoc.tables;
    if (tableAnchor !== null && tableAnchor !== undefined) {
      tables = tables.filter(t => t.rows.flat().some(c => cellText(c) === tableAnchor));
      if (tables.length !== 1) return;
    }
    const all = tables.flatMap(t => t.rows.flat());
    const hits = all.filter(c => cellText(c) === targetText);
    if (hits.length !== 1 || hits[0].nonPlain) return; // the real command also returns
    const f = faultActive() ? script.fault : null;
    if (f === 'ignore') return;
    const other = (script.otherTable ? liveDoc.tables.filter(t => !tables.includes(t)).flatMap(t => t.rows.flat()) : all).find(c => c !== hits[0]);
    const cell = f === 'cell-wrong-cell' && other ? other : hits[0];
    const p0 = cell.paras[0];
    const kept = { text: f === 'cell-append' ? `${p0.text}${text}` : text, props: { ...p0.props }, style: p0.style, alignment: p0.alignment };
    if (f === 'cell-drop-props') kept.props = { ...kept.props, bold: false, fontFamily: 'Arial' };
    if (f === 'cell-drop-style') kept.style = 'Table Text';
    cell.paras = f === 'cell-keep-rest' ? [kept, ...cell.paras.slice(1)] : [kept];
    if (f === 'cell-extra-empty-paragraph') cell.paras.push({ text: '', props: { ...kept.props } });
    if (f === 'cell-touch-other' && other) other.paras = [{ ...other.paras[0], text: `${other.paras[0].text} (edited)` }];
    if (f === 'cell-touch-other-format' && other) other.paras = [{ ...other.paras[0], props: { ...other.paras[0].props, italic: !other.paras[0].props.italic } }];
    if (f === 'cell-unmerge') cell.span = 1;
    if (f === 'cell-add-row') { const t = liveDoc.tables.find(x => x.rows.flat().includes(cell)); t.rows.push(t.rows[0].map(() => ({ span: 1, paras: [{ text: '', props: { ...kept.props } }] }))); }
    if (f === 'cell-change-outside' && live.length) live[0].text = `${live[0].text} (edited)`;
    dirty = true;
  }

  /** Wave 4b: a row write (meta mirrors the callCommand: anchor table, full-cell-text identity). */
  function applyRow({ rowAction, before, targetText, tableAnchor }) {
    const n = x => String(x).replace(/\s+/g, ' ').trim();
    const cellText = c => n(c.paras.map(x => x.text).join('\r\n'));
    let tables = liveDoc.tables;
    if (tableAnchor !== null && tableAnchor !== undefined) {
      tables = tables.filter(t => t.rows.flat().some(c => cellText(c) === tableAnchor));
      if (tables.length !== 1) return;
    }
    const hits = [];
    for (const t of tables) t.rows.forEach((row, ri) => row.forEach(c => { if (cellText(c) === targetText) hits.push({ t, ri }); }));
    if (hits.length !== 1) return;
    const f = faultActive() ? script.fault : null;
    if (f === 'ignore') return;
    const { t } = hits[0];
    let ri = hits[0].ri;
    if (f === 'row-stale-index') ri = Math.max(0, ri - 1); // a position from a stale state
    const blank = () => t.rows[ri].map(() => ({ span: 1, paras: [{ text: '', props: { ...PLAIN } }] }));
    if (rowAction === 'add') {
      const side = f === 'row-wrong-side' ? !before : before;
      const row = blank();
      if (f === 'row-new-shape') row.pop();
      if (f === 'row-new-not-empty') row[0].paras[0].text = 'x';
      t.rows.splice(side ? ri : ri + 1, 0, row);
      if (f === 'row-twice') t.rows.splice(side ? ri : ri + 1, 0, blank());
    } else {
      t.rows.splice(f === 'row-delete-other' ? (ri + 1 < t.rows.length ? ri + 1 : ri - 1) : ri, 1);
    }
    if (f === 'row-reorder' && t.rows.length > 2) { const x = t.rows[t.rows.length - 1]; t.rows[t.rows.length - 1] = t.rows[t.rows.length - 2]; t.rows[t.rows.length - 2] = x; }
    if (f === 'row-touch-cell') { const c = t.rows.flat().find(x => cellText(x) !== ''); if (c) c.paras = [{ ...c.paras[0], props: { ...c.paras[0].props, bold: !c.paras[0].props.bold } }]; }
    if (f === 'row-touch-other-table') { const o = liveDoc.tables.find(x => x !== t); if (o) o.rows[0][0].paras = [{ ...o.rows[0][0].paras[0], text: `${o.rows[0][0].paras[0].text} (edited)` }]; }
    if (f === 'row-change-outside' && live.length) live[0].text = `${live[0].text} (edited)`;
    dirty = true;
  }

  /**
   * Wave 4d: merge / split (meta mirrors the callCommand: anchor table, physical coordinates re-checked
   * by text). The model follows merge/split probe 1: MergeCells of a proper rectangle -> ONE owner with
   * the region's paragraphs in row-major order (each keeps its props), the covered rows below keep one
   * empty continuation cell of the owner's width; a non-rectangle / partial cover -> no change. Split
   * with rows x cols == the span -> 1x1 cells, all the content top-left, the new cells empty.
   */
  function tableByAnchor(tableAnchor) {
    const n = x => String(x).replace(/\s+/g, ' ').trim();
    const cellText = c => n(c.paras.map(x => x.text).join('\r\n'));
    const ts = liveDoc.tables.filter(t => t.rows.flat().some(c => cellText(c) === tableAnchor));
  return { t: ts.length === 1 ? ts[0] : null, cellText };
  }
  const startsOf = row => { let at = 0; return row.map(c => { const s0 = at; at += c.span || 1; return s0; }); };
  const blankLike = src => ({ span: 1, paras: [{ text: '', props: { ...PLAIN }, alignment: src && src.alignment }] });

  function applyMerge({ tableAnchor, startText, endText, start, end, cells }) {
    const { t, cellText } = tableByAnchor(tableAnchor);
    if (!t) return;
    const at = ([r, k]) => (t.rows[r] ? t.rows[r][k] : null);
    if (!at(start) || !at(end) || cellText(at(start)) !== startText || cellText(at(end)) !== endText) return;
    const f = faultActive() ? script.fault : null;
    if (f === 'ignore') return;
    const picked = cells.map(([r, k]) => ({ r, k, cell: at([r, k]), c: startsOf(t.rows[r])[k] }));
    if (picked.some(x => !x.cell || x.cell.vContinue || (x.cell.rowspan || 1) > 1)) return;
    const r1 = Math.min(...picked.map(x => x.r)); const r2 = Math.max(...picked.map(x => x.r));
    const c1 = Math.min(...picked.map(x => x.c)); const c2 = Math.max(...picked.map(x => x.c + (x.cell.span || 1) - 1));
    // a proper rectangle: every logical cell of the box covered by exactly the picked cells, none reaching out
    for (let r = r1; r <= r2; r += 1) {
      const st = startsOf(t.rows[r]);
      const inRow = t.rows[r].map((c, k) => ({ c, k, s: st[k], e: st[k] + (c.span || 1) - 1 })).filter(x => !(x.e < c1 || x.s > c2));
      if (inRow.some(x => x.s < c1 || x.e > c2 || !picked.some(p => p.r === r && p.k === x.k))) return;
    }
    const ordered = picked.slice().sort((a, b) => a.r - b.r || a.c - b.c);
    let paras = ordered.flatMap(x => x.cell.paras.map(p => ({ ...p, props: { ...p.props } })));
    if (f === 'merge-lose-content') paras = paras.slice(0, 1);
    if (f === 'merge-join-paragraphs') paras = [{ ...paras[0], text: paras.map(p => p.text).join(' ') }];
    if (f === 'merge-reformat' && paras[1]) paras[1] = { ...paras[1], props: { ...paras[1].props, bold: !paras[1].props.bold } };
    const cols = c2 - c1 + 1 - (f === 'merge-wrong-span' ? 1 : 0);
    const rows = r2 - r1 + 1;
    for (let r = r2; r >= r1; r -= 1) {
      const idx = picked.filter(x => x.r === r).map(x => x.k).sort((a, b) => a - b);
      const pos = idx[0];
      for (let i = idx.length - 1; i >= 0; i -= 1) t.rows[r].splice(idx[i], 1);
      t.rows[r].splice(pos, 0, r === r1 ? { span: cols, ...(rows > 1 ? { rowspan: rows } : {}), paras } : { span: cols, vContinue: true, paras: [{ text: '', props: { ...PLAIN } }] });
    }
    if (f === 'merge-touch-cell') { const c = t.rows.flat().find(x => !x.vContinue && cellText(x) !== '' && x.paras !== paras); if (c) c.paras = [{ ...c.paras[0], props: { ...c.paras[0].props, italic: !c.paras[0].props.italic } }]; }
    if (f === 'merge-touch-other-table') { const o = liveDoc.tables.find(x => x !== t); if (o) o.rows[0][0].paras = [{ ...o.rows[0][0].paras[0], text: `${o.rows[0][0].paras[0].text} (edited)` }]; }
    dirty = true;
  }

  function applySplit({ tableAnchor, targetText, cell, rows, cols }) {
    const { t, cellText } = tableByAnchor(tableAnchor);
    if (!t) return;
    const [r, k] = cell;
    const owner = t.rows[r] && t.rows[r][k];
    if (!owner || cellText(owner) !== targetText) return;
    const f = faultActive() ? script.fault : null;
    if (f === 'ignore') return;
    const rs = owner.rowspan || 1; const cs = owner.span || 1;
    if (rs !== rows || cs !== cols) return; // the refined-grid cases are refused before dispatch; not modelled
    const c0 = startsOf(t.rows[r])[k];
    const src = owner.paras[0];
    const newCells = n => Array.from({ length: n }, () => blankLike(src));
    const kept = { span: 1, paras: owner.paras.map(p => ({ ...p, props: { ...p.props } })) };
    if (f === 'split-join-paragraphs') kept.paras = [{ ...kept.paras[0], text: kept.paras.map(p => p.text).join(' ') }];
    if (f === 'split-reformat') kept.paras[0] = { ...kept.paras[0], props: { ...kept.paras[0].props, bold: !kept.paras[0].props.bold } };
    let first = [kept, ...newCells(cols - 1)];
    if (f === 'split-content-wrong-cell' && first.length > 1) first = [first[1], first[0], ...first.slice(2)];
    if (f === 'split-wrong-grid') first = first.slice(0, -1).concat([{ ...blankLike(src), span: 2 }]);
    t.rows[r].splice(k, 1, ...first);
    for (let i = 1; i < rs; i += 1) {
      const row = t.rows[r + i];
      const kk = startsOf(row).findIndex((s0, j) => s0 === c0 && row[j].vContinue);
      if (kk >= 0) row.splice(kk, 1, ...newCells(cols));
    }
    dirty = true;
  }

  /**
   * Wave 5: a header/footer write (meta mirrors the callCommand, which re-checks the plan's guards;
   * header/footer probes 1-2): the whole content becomes ONE plain paragraph (default run props,
   * left alignment); a created part exists only on that section. Faults model the negative controls.
   */
  function applyHeaderFooter({ part, variant, sectionIndex, sectionCount, create, text, activate = false }) {
    const secs = liveDoc.sections;
    if (secs.length !== sectionCount) return;
    const sec = secs[sectionIndex];
    // M5: the command re-reads the variant's flag (internal, read only) and refuses a changed state
    if (variant === 'title' && activate !== (sec && sec.titlePage !== true)) return;
    if (variant === 'even' && activate !== (liveDoc.evenAndOdd !== true)) return;
    if (!sec || !!sec[part][variant] === create) return;
    if (secs.length > 1) {
      for (let j = sectionIndex + 1; j < secs.length; j += 1) if (!secs[j][part][variant]) return;
      if (create && sectionIndex > 0) return;
    }
    if (sec[part][variant] && sec[part][variant].nonPlain) return;
    const f = faultActive() ? script.fault : null;
    if (f === 'ignore') return;
    const other = part === 'header' ? 'footer' : 'header';
    const fresh = t => ({ nonPlain: false, paras: [{ text: t, props: { ...PLAIN }, alignment: 'left' }] });
    if (f === 'hf-into-other-kind') sec[other][variant] = fresh(text);
    else if (f === 'hf-wrong-section') { const j = sectionIndex + 1 < secs.length ? sectionIndex + 1 : sectionIndex - 1; if (secs[j]) secs[j][part][variant] = fresh(text); }
    else if (f === 'hf-wrong-variant') sec[part].even = fresh(text);
    else if (f === 'hf-leftover-paragraph') sec[part][variant] = { nonPlain: false, paras: [{ text: '', props: { ...PLAIN }, alignment: 'left' }, { text, props: { ...PLAIN }, alignment: 'left' }] };
    else if (f === 'hf-append') sec[part][variant] = { nonPlain: false, paras: [{ text: `${(sec[part][variant] || fresh('')).paras.map(x => x.text).join(' ')} ${text}`.trim(), props: { ...PLAIN }, alignment: 'left' }] };
    else sec[part][variant] = fresh(text);
    if (f === 'hf-body-change' && live.length) live[0].text = `${live[0].text} (edited)`;
    if (f === 'hf-other-section' && secs.length > 1) { const j = sectionIndex === 0 ? 1 : 0; secs[j][other][variant] = fresh('CHANGED ELSEWHERE'); }
    if (f === 'hf-other-kind-changed') sec[other][variant] = fresh('CHANGED OTHER KIND');
    if (f === 'hf-cursor-stuck') inNote = true;
    // M5: activation AFTER the part write (the command's order); faults model a missed / a stray flag
    if (activate && f !== 'hf-flag-not-set') { if (variant === 'title') sec.titlePage = true; else if (variant === 'even') liveDoc.evenAndOdd = true; }
    if (f === 'hf-other-title-flag' && secs.length > 1) secs[sectionIndex === 0 ? 1 : 0].titlePage = true;
    dirty = true;
  }

  /**
   * Wave 4c: a column write (meta mirrors the callCommand). The model follows column probe 1: the
   * logical grid from the colspans; AddColumn puts one empty span-1 cell at the boundary in every row,
   * and where a colspan STRADDLES the boundary it lands before that cell (the measured drift);
   * RemoveColumn removes the span-1 cell of the column, and a colspan covering it is LOST, an empty
   * cell of span-1 appended to its row (the measured content loss). The adapter's guard refuses both.
   */
  function applyColumn({ columnAction, before, targetText, tableAnchor }) {
    const n = x => String(x).replace(/\s+/g, ' ').trim();
    const cellText = c => n(c.paras.map(x => x.text).join('\r\n'));
    let tables = liveDoc.tables;
    if (tableAnchor !== null && tableAnchor !== undefined) {
      tables = tables.filter(t => t.rows.flat().some(c => cellText(c) === tableAnchor));
      if (tables.length !== 1) return;
    }
    const hits = [];
    for (const t of tables) t.rows.forEach(row => { let at = 0; row.forEach((c, k) => { if (cellText(c) === targetText) hits.push({ t, col: at, span: c.span || 1, phys: k }); at += c.span || 1; }); });
    if (hits.length !== 1) return;
    const f = faultActive() ? script.fault : null;
    if (f === 'ignore') return;
    const { t } = hits[0];
    let col = hits[0].col;
    if (f === 'col-stale-index') col = Math.max(0, col - 1); // a position from a stale state
    const blank = () => ({ span: f === 'col-new-span2' ? 2 : 1, paras: [{ text: f === 'col-new-not-empty' ? 'x' : '', props: { ...PLAIN } }] });
    const starts = row => { let at = 0; return row.map(c => { const s0 = at; at += c.span || 1; return s0; }); };
    if (columnAction === 'add') {
      const side = f === 'col-wrong-side' ? !before : before;
      const b = side ? col : col + hits[0].span;
      t.rows.forEach(row => {
        const st = starts(row);
        let p = row.findIndex((c, k) => st[k] >= b || (st[k] < b && b < st[k] + (c.span || 1)));
        if (p < 0) p = row.length;
        row.splice(p, 0, blank());
        if (f === 'col-twice') row.splice(p, 0, blank());
      });
    } else {
      const width = t.rows[0].reduce((w, c) => w + (c.span || 1), 0);
      const del = f === 'col-delete-other' ? (col + 1 < width ? col + 1 : col - 1) : col;
      t.rows.forEach(row => {
        const st = starts(row);
        // the source's model: the PHYSICAL index in every row (a logical/physical confusion)
        if (f === 'col-physical-index') { row.splice(hits[0].phys, 1); return; }
        const k = row.findIndex((c, i) => st[i] <= del && del < st[i] + (c.span || 1));
        if (k < 0) return;
        const span = row[k].span || 1;
        row.splice(k, 1);
        if (span > 1) row.push({ span: span - 1, paras: [{ text: '', props: { ...PLAIN } }] });
      });
    }
    if (f === 'col-reorder') { const r0 = t.rows[0]; if (r0.length > 2) { const x = r0[r0.length - 1]; r0[r0.length - 1] = r0[r0.length - 2]; r0[r0.length - 2] = x; } }
    if (f === 'col-touch-cell') { const c = t.rows.flat().find(x => cellText(x) !== ''); if (c) c.paras = [{ ...c.paras[0], props: { ...c.paras[0].props, bold: !c.paras[0].props.bold } }]; }
    if (f === 'col-span-change') { const c = t.rows.flat().find(x => (x.span || 1) > 1); if (c) c.span = 1; }
    if (f === 'col-touch-other-table') { const o = liveDoc.tables.find(x => x !== t); if (o) o.rows[0][0].paras = [{ ...o.rows[0][0].paras[0], text: `${o.rows[0][0].paras[0].text} (edited)` }]; }
    if (f === 'col-change-outside' && live.length) live[0].text = `${live[0].text} (edited)`;
    dirty = true;
  }

  /** Wave 3: a bookmark on the exact target (meta mirrors the callCommand's inputs and guards). */
  function applyBookmark({ targetText, name }) {
    if (inNote) return;
    const hits = exact(targetText);
    if (hits.length !== 1 || liveDoc.bookmarks.some(b => b.name === name)) return;
    const f = faultActive() ? script.fault : null;
    if (f === 'ignore') return;
    const other = live.find(x => x !== hits[0]);
    liveDoc.bookmarks.push({ name, text: f === 'bookmark-wrong-target' && other ? other.text : targetText });
    if (f === 'bookmark-twice') liveDoc.bookmarks.push({ name, text: targetText });
    dirty = true;
  }

  /** Wave 3: a hyperlink over the target paragraph, its text properties restored (source parity). */
  function applyLink({ targetText, url, tooltip }) {
    if (inNote) return;
    const hits = exact(targetText);
    if (hits.length !== 1) return;
    // ApiRange.AddHyperlink with an empty screen tip writes nothing (W3 LIVE run 2)
    if (!tooltip) return;
    const f = faultActive() ? script.fault : null;
    if (f === 'ignore') return;
    const p = hits[0];
    p.links = [{ url: f === 'link-wrong-url' ? `${url}/wrong` : url, tip: tooltip }];
    // the "Hyperlink" style the restore should undo (probe 1: effective underline false -> true)
    if (f === 'link-no-restore') p.props = { ...p.props, underline: true };
    // the source's paragraph-wide AddHyperlink: an existing reference moves inside and loses its style (probe 4)
    if (f === 'link-swallows-reference' && p.refs) p.refs = { footnote: 0, endnote: 0 };
    // collateral: another paragraph also gets a link / an unrelated footnote disappears / a bookmark moves
    if (f === 'link-touch-other') { const o = live.find(x => x !== p); if (o) o.links = [{ url, tip: tooltip }]; }
    if (f === 'link-drop-other-footnote') liveDoc.footnotes.pop();
    if (f === 'link-move-bookmark' && liveDoc.bookmarks.length) liveDoc.bookmarks[0] = { ...liveDoc.bookmarks[0], text: p.text };
    dirty = true;
  }

  /** Wave 3: a footnote/endnote on the target; the command returns the cursor to the main body. */
  function applyNote({ kind, targetText, text }) {
    if (inNote) return;
    const hits = exact(targetText);
    if (hits.length !== 1) return;
    const f = faultActive() ? script.fault : null;
    if (f === 'ignore') return;
    const p = f === 'note-wrong-paragraph' ? live.find(x => x !== hits[0]) : hits[0];
    p.refs = { footnote: 0, endnote: 0, ...(p.refs || {}) };
    p.refs[kind] += 1;
    liveDoc[`${kind}s`].push(text);
    if (f === 'note-twice') liveDoc[`${kind}s`].push(text);
    if (f === 'note-stuck') inNote = true; // the cursor is left inside the note (the source's SetDocPosType need)
    dirty = true;
  }

  function applyInsert({ sourceText, side, text }) {
    const hits = exact(sourceText);
    if (hits.length !== 1) return;
    const f = faultActive() ? script.fault : null;
    if (f === 'ignore') return;
    const at = live.indexOf(hits[0]);
    const effective = f === 'wrong-side' ? (side === 'before' ? 'after' : 'before') : side;
    const item = { text, props: { ...PLAIN } };
    live.splice(effective === 'before' ? at : at + 1, 0, item);
    if (f === 'insert-twice') live.push({ text, props: { ...PLAIN } });
    dirty = true;
  }

  function applyDelete({ sourceText }) {
    const hits = exact(sourceText);
    if (hits.length !== 1) return;
    const f = faultActive() ? script.fault : null;
    if (f === 'ignore') return;
    const at = live.indexOf(hits[0]);
    if (f === 'delete-neighbour') { live.splice(at + 1 < live.length ? at + 1 : at - 1, 1); dirty = true; return; }
    live.splice(at, 1);
    dirty = true;
  }

  function applyMove({ sourceText, side, destinationText }) {
    const src = exact(sourceText);
    const dst = exact(destinationText);
    if (src.length !== 1 || dst.length !== 1) return;
    const f = faultActive() ? script.fault : null;
    if (f === 'ignore') return;
    const copy = { ...src[0], props: { ...src[0].props } };
    const effective = f === 'wrong-side' ? (side === 'before' ? 'after' : 'before') : side;
    let at = live.indexOf(dst[0]) + (effective === 'before' ? 0 : 1);
    if (f === 'move-not-adjacent') at = Math.min(at + 1, live.length);
    live.splice(at, 0, copy);
    if (f !== 'no-delete') live.splice(live.indexOf(src[0]), 1);
    dirty = true;
  }

  return {
    calls,
    store,
    liveState: () => clone(live || []),
    // M8 cross-format: the engine's open / close failure branches through the adapter (fake-session shape)
    async open() { calls.open += 1; if (script.openFails) return { ok: false, detail: 'fake docx open refused' }; live = clone(store.durable); liveDoc = cloneDoc(store.durableDoc || {}); seedNumberings(); inNote = !!script.startInNote; dirty = false; return { ok: true, id: `docx-session-${calls.open}` }; },
    // script.saveObserver (M8, opt-in): the real DOCX driver's save-ack pairing (M7, PR #17) -- forceSave returns its own
    // confirmation and close re-evaluates it (a dispatch after it supersedes it), as the XLSX fake does. Without the knob
    // this fake keeps its historical no-observer shape (EDITOR_CLEAN), so the existing DOCX tests are unchanged.
    async close() {
      calls.close += 1; live = null; liveDoc = null;
      const confirmation = !script.saveObserver || docxConfirmedAt === null ? null
        : (calls.dispatch > docxConfirmedAt ? { state: 'unconfirmed', reason: 'superseded-by-a-newer-save' } : { state: 'confirmed', reason: 'forcesave-confirmed' });
      if (script.closeFails) return { ok: false, detail: 'fake docx close refused', ...(confirmation ? { confirmation } : {}) };
      return confirmation ? { ok: true, confirmation } : { ok: true };
    },
    async forceSave() {
      calls.forceSave += 1; store.durable = clone(live); store.durableDoc = cloneDoc(liveDoc); store.saves += 1; dirty = false;
      if (!script.saveObserver) return { ok: true };
      docxConfirmedAt = calls.dispatch;
      return { ok: true, confirmed: true, forceSave: { state: 'confirmed', reason: 'forcesave-confirmed' } };
    },
    async canSave() { calls.canSave += 1; return { ok: true, canSave: dirty }; },
    async readLive() {
      calls.readLive += 1;
      // M5 gate 7: the internal per-level format read (bullet 5, numbered 13 -- measured values); scripted
      // faults: unreadable (reader absent), a disagreeing own level, an unmeasured format on one level
      const FMT = { bullet: 5, numbered: 13 };
      const levelFormatsOf = p => {
        if (script.levelFormatsUnreadable) return null;
        const f = liveDoc.numberings[p.list.numId].map(t => FMT[t]);
        if (script.levelFormatsDisagree) f[p.list.level] = f[p.list.level] === 5 ? 13 : 5;
        if (Number.isInteger(script.levelFormatUnknownAt)) f[script.levelFormatUnknownAt] = 99;
        return f;
      };
      const numberingOf = p => (script.numberingUnreadable ? {} : { numbering: p.list ? { present: true, level: p.list.level, numId: p.list.numId, levelFormats: levelFormatsOf(p) } : { present: false } });
      // M5 gate 8: the internal calculated grid (default: equal widths over the logical width)
      const gridOfTable = t => (script.gridUnreadable ? undefined : Array.isArray(t.gridCalc) ? [...t.gridCalc] : Array((t.rows[0] || []).reduce((a, c) => a + (c.span || 1), 0)).fill(20));
      // Wave 3: a link and each bookmark marker are non-run children (live: 'hyperlink' / 'unsupported')
      const markers = p => (p.links ? p.links.length : 0) + 2 * liveDoc.bookmarks.filter(b => b.text === p.text).length;
      const richOf = p => (script.richUnreadable ? {} : { hyperlinks: (p.links || []).map(l => ({ url: l.url, tip: l.tip, text: p.text, style: 'Hyperlink' })), noteRefs: { footnote: 0, endnote: 0, ...(p.refs || {}) } });
      const paragraphs = live.map(p => {
        const rp = script.propsUnreadable ? {} : runPayload(p);
        if (!script.propsUnreadable && markers(p)) { rp.nonRunChildren += markers(p); }
        return { text: p.text, props: script.propsUnreadable ? null : p.props, format: script.propsUnreadable ? undefined : formatOf(p), ...rp, ...numberingOf(p), ...richOf(p) };
      });
      const document = script.docUnreadable ? null : { bookmarks: liveDoc.bookmarks.map(b => ({ ...b })), footnotes: [...liveDoc.footnotes], endnotes: [...liveDoc.endnotes], inMainBody: !inNote,
        ...(script.tablesUnreadable ? {} : { tables: liveDoc.tables.map(t => ({ rows: t.rows.map(r => r.filter(c => !c.gap).map(cellRead)), gridCalc: gridOfTable(t) })).filter((_, i) => !(script.tableReadMissesOne && i === 0)) }),
        // M5 gate 6: the document default text properties (the replacement's run properties = PLAIN)
        defaultText: script.defaultTextUnreadable ? null : { fontFamily: 'Arial', fontSize: 22 },
        ...(script.sectionsUnreadable ? {} : { sections: liveDoc.sections.map(sec => ({ header: Object.fromEntries(HF_VARIANTS.map(v => [v, partRead(sec.header[v])])), footer: Object.fromEntries(HF_VARIANTS.map(v => [v, partRead(sec.footer[v])])),
          ...(script.flagsUnreadable ? {} : { titlePage: sec.titlePage === true, evenAndOdd: liveDoc.evenAndOdd === true }) })) }) };
      // M8 runtime gate: the REAL driver's gate (default: the verified runtime; script.runtime simulates another one)
      const gate = gateOf(script);
      return { ok: true, authority: 'LIVE_READ', payload: { html: htmlOf(live, store.images || [], liveDoc), paragraphs, document, runtime: { internalGate: gate } } };
    },
    async readPackage() { calls.readPackage += 1; return { ok: true, parts: {} }; },
    async dispatch(_handle, command) {
      // M8 runtime gate, the REAL driver's second line (internalWriteRefusal) on this fake's EFFECTIVE runtime --
      // gateOverride tweaks the READ-side view only, never this (the second line has no gate-object input, only a runtime id)
      const refused = internalWriteRefusal(command, effectiveRuntime(script));
      if (refused) { calls.refusedInternal = (calls.refusedInternal || 0) + 1; return refused; }
      calls.dispatch += 1;
      if (script.fault === 'throw') throw new Error('transport lost');
      if (!command || command.kind !== 'callCommand' || !command.meta) return { ok: false, dispatched: true };
      const a = command.meta.action;
      if (a === 'rewrite') applyRewrite(command.meta);
      else if (a === 'format') (command.meta.matchText !== undefined ? applyRangeFormat : applyFormat)(command.meta);
      else if (a === 'insert') applyInsert(command.meta);
      else if (a === 'delete') applyDelete(command.meta);
      else if (a === 'move') applyMove(command.meta);
      else if (a === 'list') applyList(command.meta);
      else if (a === 'bookmark') applyBookmark(command.meta);
      else if (a === 'link') applyLink(command.meta);
      else if (a === 'note') applyNote(command.meta);
      else if (a === 'cell') applyCell(command.meta);
      else if (a === 'row') applyRow(command.meta);
      else if (a === 'column') applyColumn(command.meta);
      else if (a === 'merge') applyMerge(command.meta);
      else if (a === 'split') applySplit(command.meta);
      else if (a === 'headerFooter') applyHeaderFooter(command.meta);
      // M5 constraints: a scripted SIDE EFFECT of the write on the live state (a change the command
      // should not have made -- another paragraph, a table, formatting), for the post-write L3 controls
      if (typeof script.sideEffect === 'function') { script.sideEffect({ live, liveDoc, dispatch: calls.dispatch }); dirty = true; }
      return { ok: true, dispatched: true };
    },
  };
}

module.exports = { createStore, createFakeDocxDriver, htmlOf, formatOf, runPayload };
