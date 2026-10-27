'use strict';

const { internalGateOf, internalWriteRefusal, B21_BRIDGE_READ_RUNTIMES, L7_BRIDGE_RUNTIMES } = require('../../src/drivers/eurooffice/docx-driver.cjs');
// M8 runtime gate: the REAL driver's gate for script.runtime (default: the accepted word runtime -- /5 since the GH #292 S4
// D-state -- where every DOCX path incl. the B18
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
    body += `<table>${t.rows.map(r => `<tr>${r.filter(c => !c.vContinue).map(c => (c.gap ? `<td${c.span > 1 ? ` colspan="${c.span}"` : ''} style="mso-cell-special:placeholder;border:none"><p><span>&nbsp;</span></p></td>` : `<td${c.span > 1 ? ` colspan="${c.span}"` : ''}${c.rowspan > 1 ? ` rowspan="${c.rowspan}"` : ''}>${c.paras.map(x => `<p><span>${esc(x.text)}</span></p>`).join('')}</td>`)).join('')}</tr>`).join('')}</table>`;
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
    // GH #290: the explicit (direct) flag state; null = absent (a fixture may model a style-set flag by leaving it null)
    paginationDirect: { keepNext: null, keepLines: null, pageBreakBefore: null, widowControl: null },
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
// GH #269d: the identity the fake editor stamps on a comment/reply written without an explicit author
const SESSION = Object.freeze({ currentName: 'EO Session Agent', documentUserId: 'eo-session-1' });
const cloneDoc = d => ({ ...(d.abstractOf ? { abstractOf: { ...d.abstractOf } } : {}),
  // GH #293 L7: list items the INVENTORY does not see (a table cell's) -- { text, numId, level, pos } -- and per numId the levels its
  // startOverride covers; both read only by the word /5 members model below
  ...(d.hiddenMembers ? { hiddenMembers: d.hiddenMembers.map(h => ({ ...h })) } : {}), ...(d.startOverrides ? { startOverrides: JSON.parse(JSON.stringify(d.startOverrides)) } : {}),
  comments: (d.comments || []).map(c => ({ ...c, replies: (c.replies || []).map(r => ({ ...r })) })), bookmarks: (d.bookmarks || []).map(b => ({ ...b })), footnotes: [...(d.footnotes || [])], endnotes: [...(d.endnotes || [])],
  // W04-4: `format` is FLAT on the table object (same dual-context reasoning as the
  // page-setup fields below -- cloneDoc() runs fixture -> live AND live -> durable; a write
  // (applyTableFormat()) sets `t.format` directly, never nested under anything else, so a plain
  // JSON round-trip here carries a write through BOTH directions without silently reverting it).
  // GH #291 N3: a cell may hold nested tables (`nested`), cloned the same way
  tables: (d.tables || []).map(t => ({ rows: t.rows.map(r => r.map(cloneCell)), ...(Array.isArray(t.gridCalc) ? { gridCalc: [...t.gridCalc] } : {}), format: t.format ? JSON.parse(JSON.stringify(t.format)) : {} })),
  // W04-4: the table-style name/id list `format`'s `style` field resolves against at write time
  // (same name -> real measured id "701" default, so a fixture needs nothing extra to use it).
  tableStyles: (d.tableStyles || DEFAULT_TABLE_STYLES).map(s => ({ ...s })),
  // GH #268b: the images, in the DRIVER's page-side row shape (docx-images-read.cjs): index / id / name / descr / title / cx / cy / drawingType / wrapType / paragraph / json
  images: (d.images || []).map(i => ({ ...i })),
  // M5 gate 7: per numbering, the list type of each of its 9 levels (EO re-templates ONE level)
  numberings: Object.fromEntries(Object.entries(d.numberings || {}).map(([k, v]) => [k, [...v]])),
  // W-07: per numbering, the SAME per-level ToJSON-derived fine-grained format (numFmt/lvlText/
  // suff/lvlJc/start) -- a SHARED definition, like `numberings` above: a write at one level
  // changes every item sitting on that numId+level.
  numberingFormats: Object.fromEntries(Object.entries(d.numberingFormats || {}).map(([k, v]) => [k, v.map(x => (x ? { ...x } : x))])),
  // Wave 5: one entry per section; per kind (header/footer) and variant: null = no part of its own,
  // else { paras: [{ text, props, alignment }], nonPlain }. Absent -> one section with no parts.
  sections: (d.sections || [{}]).map(sec => ({ ...Object.fromEntries(['header', 'footer'].map(k => [k, Object.fromEntries(HF_VARIANTS.map(v => {
    const part = sec[k] && sec[k][v];
    return [v, part ? { nonPlain: !!part.nonPlain, paras: part.paras.map(x => ({ ...x, props: { ...(x.props || {}) } })), ...(Array.isArray(part.tables) ? { tables: part.tables.map(t => ({ at: t.at, rows: t.rows.map(r => [...r]) })) } : {}) } : null];
  }))])), titlePage: sec.titlePage === true,
    // W11: page setup, US Letter portrait defaults. cloneDoc() runs in TWO contexts -- fixture ->
    // live (open()) AND live -> durable (forceSave()/close()) -- and by the second context the
    // fields are already FLAT on `sec` (applyPageSetup() writes them directly, same as titlePage
    // above), never nested. Reading flat here (not a nested sec.pageSetup) is what makes a write
    // survive a save; a nested read would silently re-apply the defaults on every forceSave,
    // discarding whatever was just written.
    pageWidth: sec.pageWidth ?? 12240,
    pageHeight: sec.pageHeight ?? 15840,
    margins: { top: 1440, right: 1440, bottom: 1440, left: 1440, ...sec.margins },
    columnsCount: sec.columnsCount ?? 1,
    // W11-4: GetColumnsSpaces() is independently re-readable (count equal columns have
    // count-1 gaps). Flat, same dual-context reasoning as the other page-setup fields above --
    // `sec.columnsSpaces` wins if already present (post-write), else derived from count/space.
    columnsSpaces: sec.columnsSpaces ?? Array.from({ length: Math.max(0, (sec.columnsCount ?? 1) - 1) }, () => sec.columnsSpace ?? 0),
    startPageNumber: sec.startPageNumber ?? 1,
  })),
  // M5: the document-wide even/odd flag (the section-level internal reader reports it per section)
  evenAndOdd: d.evenAndOdd === true,
  // W-10 (style DEFINITIONS): keyed by name, NEVER styleId (the real driver's own reason: ToJSON
  // carries no styleId, measured live #ref). Fixture seed: { name: { text: {...}, paragraph: {...} } }.
  styles: Object.fromEntries(Object.entries(d.styles || {}).map(([name, st]) => [name, { text: { ...(st.text || {}) }, paragraph: { ...(st.paragraph || {}) }, ...(st.basedOn ? { basedOn: st.basedOn } : {}), ...(st.type ? { type: st.type } : {}), ...(st.next ? { next: st.next } : {}), ...(st.link ? { link: st.link } : {}) }])), // GH #292 S3 / S5 / S4: base, type, next and link survive the save copy
  //: names created THIS session via define_style -- GetAllStyles()-blind by
  // measured fact. The FIXTURE's own seeded styles are never in here by DEFAULT (they model
  // built-ins) -- W-10-10: a fixture MAY seed `doc.customStyleNames` (an array of
  // names) to model a style that is ALREADY custom/GetAllStyles-blind from a PRIOR session (a maintainer's
  // own measured fact: the blindness is not limited to the SAME task that created the style).
  customStyleNames: new Set(d.customStyleNames || []) });
const HF_VARIANTS = ['default', 'even', 'title'];

// W04-4 (#ref, MEASURED default): the real document's own 'Table Grid' style, read live.
const DEFAULT_TABLE_STYLES = Object.freeze([Object.freeze({ name: 'Table Grid', styleId: 701 })]);
// Border edges whose ToJSON()/XML key names differ (MEASURED #ref): left->start (ToJSON only,
// XML keeps `left`), right->end (ToJSON only, XML keeps `right`). The fake driver stores borders
// keyed by the TASK-FACING edge name (top/bottom/left/right/insideH/insideV) -- same convention
// as table-format.cjs's own fields -- so this map is NOT needed inside the fake; it exists only in
// the real driver's ToJSON extraction and in the LIVE acceptance test's raw-XML/ToJSON cross-check.

/** Wave 5: the driver's public read of one header/footer part (paragraph text ends in '\r\n', probes 1-2). */
function partRead(part) {
  if (!part) return null;
  const fmt = x => ({ ...runFormat(x.props), alignment: x.alignment || 'left', ...(x.style ? { style: x.style } : {}) });
  // M5 gate 6: the effective properties of the paragraph's one run, in the RAW getter units (the LIVE
  // read: bold/italic/underline/strike booleans, font family, size as GetFontSize returns it)
  const runs = x => [{ text: x.text, bold: !!(x.props || {}).bold, italic: !!(x.props || {}).italic, underline: !!(x.props || {}).underline, strike: !!(x.props || {}).strike,
    fontFamily: (x.props || {}).fontFamily, fontSize: (x.props || {}).fontSize }];
  // GH #289: `x.runs` ([{ text, props? }]) models a paragraph of several runs (a field / an image is a run with no text)
  const runsOf = x => (Array.isArray(x.runs) ? x.runs.map(r => runs({ text: r.text, props: r.props || x.props })[0]) : runs(x));
  const textOf = x => (Array.isArray(x.runs) ? x.runs.map(r => r.text).join('') : x.text);
  // GH #289 H-05 (LIVE): `part.tables` ([{ at, rows: [[cell text]] }]) -- the element count counts them, the part is not simple,
  // and a part of paragraphs of runs + tables reads withTables
  const tables = Array.isArray(part.tables) ? part.tables : [];
  return { paragraphs: part.paras.length + tables.length, simple: !part.nonPlain && tables.length === 0, paras: part.paras.map(x => ({ text: `${textOf(x)}\r\n`, format: fmt(x), runs: runsOf(x), inline: x.inline === undefined ? [] : x.inline })),
    tables: tables.map(t => ({ at: t.at, rows: t.rows.map(r => r.map(x => `${x}\r\n`)) })), withTables: tables.length > 0 && !part.nonPlain };
}

/**
 * Wave 4a: the driver's public table read of one cell (table probes 1-2): the content text ends in
 * '\t' (a merged cell joins its paragraphs with '\r\n'); `simple` = only paragraphs made only of
 * runs; the effective properties of the FIRST paragraph (style name, alignment as read).
 */
function cellRead(c) {
  const p0 = c.paras[0] || { props: {} };
  const fmt = x => ({ ...runFormat(x.props), style: x.style || 'Normal', alignment: x.alignment || '1' });
  // GH #291 N3 (LIVE, gh291-n3-nested-cell-probe-20261022.json): a HOLDING cell's text runs on into its nested tables' text, its
  // element count counts the table, it is not simple; the nested tables come with it (nestedTables), each cell read the same way
  const nested = Array.isArray(c.nested) ? c.nested : [];
  // GH #291 N3x (LIVE, gh291-nested-remaining-probe-20261023.json): every level below runs on into the holder's text
  const fullText = x => [...x.paras.map(p => p.text), ...(Array.isArray(x.nested) ? x.nested.flatMap(nt => nt.rows.flat().map(fullText)) : [])].join('\r\n');
  const nestedTexts = nested.flatMap(nt => nt.rows.flat().map(fullText));
  return { text: `${[...c.paras.map(x => x.text), ...nestedTexts].join('\r\n')}\t`, paragraphs: c.paras.length + nested.length, simple: !c.nonPlain && nested.length === 0,
    format: fmt(p0), paraFormats: c.paras.map(x => ({ text: x.text, format: fmt(x) })), inline: c.inline === undefined ? [] : c.inline,
    ...(nested.length ? { nestedTables: nested.map(nt => ({ rows: nt.rows.map(r => r.map(cellRead)) })) } : {}) };
}
const cloneCell = c => ({ ...c, paras: c.paras.map(x => ({ ...x, props: { ...(x.props || {}) } })), ...(Array.isArray(c.nested) ? { nested: c.nested.map(nt => ({ rows: nt.rows.map(r => r.map(cloneCell)) })) } : {}) });

function createStore(paragraphs, { images = [], doc = {} } = {}) {
  return { durable: clone(paragraphs), durableDoc: cloneDoc(doc), saves: 0, images };
}

function createFakeDocxDriver(store, script = {}) {
  const calls = { open: 0, close: 0, forceSave: 0, canSave: 0, readLive: 0, dispatch: 0, readPackage: 0, release: 0, reset: 0, health: 0, documentKey: 0 };
  let docxConfirmedAt = null; // M8: the dispatch count at the last confirmed forcesave (script.saveObserver only)
  let live = null;
  let liveDoc = null;
  let inNote = false;
  let dirty = false;

  // A fault acts on every dispatch, or only on the dispatch numbers listed in `faultOn` (1-based).
  const faultActive = () => !!script.fault && (!script.faultOn || script.faultOn.includes(calls.dispatch));
  const exact = t => live.filter(p => p.text === t);
  const occurrences = (hay, needle) => { let n = 0; for (let i = hay.indexOf(needle); i >= 0; i = hay.indexOf(needle, i + 1)) n += 1; return n; };
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
    // GH #290: a written pagination flag is an EXPLICIT direct value
    for (const k of apply.filter(x => ['keepNext', 'keepLines', 'pageBreakBefore', 'widowControl'].includes(x))) fmt.paginationDirect = { ...(fmt.paginationDirect || {}), [k]: f === 'format-flag-not-explicit' ? null : format[k] };
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

  /**
   * GH #293 L7: the word /5 docx.numbering.listMembers.read/v1 model -- every item of each numbering in document order (a live paragraph
   * at pos 10*i, a hidden one at its own pos), the other numIds on the same definition (liveDoc.abstractOf, default its own), their
   * startOverride levels and before/after flags, a per-level signature from numberingFormats (lvlText '%1.%2' -> '2:0|1:.|2:1|1:.').
   * Off /5 (the effective runtime) every text is { unread: 'PATCH_VERSION_MISMATCH' }; script.membersUnreadable -> { unread }.
   */
  function allMembers() {
    const out = live.map((p, i) => (p.list ? { text: p.text, numId: p.list.numId, level: p.list.level, pos: i * 10 } : null)).filter(Boolean);
    for (const h of liveDoc.hiddenMembers || []) out.push({ ...h });
    return out.sort((a, b) => a.pos - b.pos);
  }
  const absOfNum = id => (liveDoc.abstractOf && liveDoc.abstractOf[id]) || `abs-${id}`;
  const sigOf = id => Array.from({ length: 9 }, (_, l) => {
    const f = liveDoc.numberingFormats[id] && liveDoc.numberingFormats[id][l];
    const lt = f && typeof f.lvlText === 'string' ? f.lvlText : `%${l + 1}.`;
    const tok = []; for (let k = 0; k < lt.length; k += 1) { if (lt[k] === '%' && /\d/.test(lt[k + 1] || '')) { tok.push(`2:${Number(lt[k + 1]) - 1}`); k += 1; } else tok.push(`1:${lt[k]}`); }
    if (script.membersSigUnreadableLevel === l) return null;
    return [f && f.numFmt ? f.numFmt : 'decimal', tok.join('|'), f && f.start !== undefined ? f.start : 1, f && f.suff ? f.suff : 'tab', f && f.lvlJc ? f.lvlJc : 'left'];
  });
  function membersOf(text) {
    if (!L7_BRIDGE_RUNTIMES.includes(effectiveRuntime(script))) return { unread: 'PATCH_VERSION_MISMATCH' };
    if (script.membersUnreadable) return { unread: 'BRIDGE:paragraphs-unreadable' };
    const all = allMembers();
    const hits = all.filter(m => m.text === text);
    const plain = live.filter(p => p.text === text);
    if (!hits.length) return plain.length === 1 ? { present: false } : { unread: plain.length ? 'target-ambiguous' : 'target-not-found' };
    if (hits.length !== 1) return { unread: 'target-ambiguous' };
    const t = hits[0];
    const abs = absOfNum(t.numId);
    const own = all.filter(m => m.numId === t.numId);
    const first = all.indexOf(own[0]); const last = all.indexOf(own[own.length - 1]);
    const contiguous = all.slice(first, last + 1).every(m => m.numId === t.numId);
    const otherIds = [...new Set(all.filter(m => m.numId !== t.numId && absOfNum(m.numId) === abs).map(m => m.numId))];
    const others = otherIds.map(id => { const ms = all.filter(m => m.numId === id); return { numId: id, starts: [...((liveDoc.startOverrides || {})[id] || [])], members: ms.map(m => ({ text: m.text, level: m.level })), sig: sigOf(id),
      before: ms.some(m => all.indexOf(m) < first), after: ms.some(m => all.indexOf(m) > last) }; });
    return { present: true, level: t.level, numId: t.numId, abstractId: abs, starts: [...((liveDoc.startOverrides || {})[t.numId] || [])], members: own.map(m => ({ text: m.text, level: m.level })),
      sig: sigOf(t.numId), others, order: script.membersOrderUnreadable ? null : { targetFirst: !others.some(o => o.before), contiguous } };
  }
  /** GH #293 L7: the docx.numbering.listDetach.write/v1 model (the handler's own re-read / compare / isolation), then the level write on the copy */
  function applyListDetach(m) {
    const v = membersOf(m.detach.text);
    if (!v || v.present !== true) return { ok: false, stage: 'list-detach', reason: 'BRIDGE:not-numbered', mutation: 'NONE' };
    if (JSON.stringify(v.members) !== JSON.stringify(m.detach.members)) return { ok: false, stage: 'list-detach', reason: 'BRIDGE:members-mismatch', mutation: 'NONE' };
    if (v.others.length !== m.detach.others) return { ok: false, stage: 'list-detach', reason: 'BRIDGE:others-mismatch', mutation: 'NONE' };
    if (v.others.length === 0) return { ok: false, stage: 'list-detach', reason: 'BRIDGE:nothing-to-detach', mutation: 'NONE' };
    const covers = (st, ms) => ms.every(x => st.includes(x.level));
    const iso = v.order && v.order.contiguous && (v.order.targetFirst || covers(v.starts, v.members)) && v.others.every(o => !(o.before && o.after) && (!o.after || covers(o.starts, o.members)));
    if (!iso) return { ok: false, stage: 'list-detach', reason: 'BRIDGE:not-isolatable', mutation: 'NONE' };
    if (script.fault === 'detach-not-observed') return { ok: false, stage: 'list-detach', reason: 'BRIDGE:write-not-observed', mutation: 'UNKNOWN' };
    numSeq += 1;
    const newId = `d${numSeq}`;
    liveDoc.numberingFormats[newId] = (liveDoc.numberingFormats[v.numId] || Array(9).fill(null)).map(x => (x ? { ...x } : x));
    liveDoc.numberings[newId] = [...(liveDoc.numberings[v.numId] || Array(9).fill('numbered'))];
    liveDoc.abstractOf = { ...(liveDoc.abstractOf || {}), [newId]: `abs-${newId}` };
    if (liveDoc.startOverrides && liveDoc.startOverrides[v.numId]) liveDoc.startOverrides[newId] = [...liveDoc.startOverrides[v.numId]];
    for (const p of live) if (p.list && p.list.numId === v.numId) p.list = { ...p.list, numId: newId };
    for (const h of liveDoc.hiddenMembers || []) if (h.numId === v.numId) h.numId = newId;
    dirty = true;
    applySetNumbering({ ...m, numId: newId });
    return { ok: true, detached: v.members.length };
  }

  /** W-07: the set_numbering write -- SHARED per numId+level, mirroring applyList's own
   * set_list_type branch (liveDoc.numberings[numId][level] = type, line ~297 above): a write
   * here changes liveDoc.numberingFormats[numId][level], visible to EVERY item sitting there. */
  function applySetNumbering(m) {
    const hits = exact(m.targetText);
    if (hits.length !== 1) return;
    const p = hits[0];
    if (!p.list || p.list.level !== m.ownLevel) return; // moved since planning -- the real command also returns
    const f = faultActive() ? script.fault : null;
    if (f === 'ignore') return;
    const slot = liveDoc.numberingFormats[m.numId] || (liveDoc.numberingFormats[m.numId] = Array(9).fill(null));
    const current = slot[m.ownLevel] || {};
    const next = { ...current };
    if (m.fields.format) { next.numFmt = m.fields.format.numFmt; next.lvlText = m.fields.format.lvlText; next.lvlJc = m.fields.format.align; }
    if (m.fields.start !== undefined) next.start = f === 'numbering-start-not-written' ? current.start : m.fields.start;
    if (m.fields.suffix !== undefined) next.suff = f === 'numbering-suffix-nothing-spelling' && m.fields.suffix === 'none' ? 'nothing' : m.fields.suffix;
    slot[m.ownLevel] = next;
    if (f === 'numbering-touch-other-numid') { const otherId = Object.keys(liveDoc.numberingFormats).find(k => k !== m.numId); if (otherId) liveDoc.numberingFormats[otherId][0] = { ...(liveDoc.numberingFormats[otherId][0] || {}), lvlJc: 'right' }; }
    // GH #248: a write at one level that also rewrites ANOTHER level of the same numbering
    if (f === 'numbering-touch-other-level') { const other = m.ownLevel === 0 ? 1 : 0; slot[other] = { ...(slot[other] || {}), numFmt: 'upperRoman' }; }
    if (f === 'numbering-change-text') p.text = `${p.text} (edited)`;
    dirty = true;
  }

  /** Wave 4a: the cell write (meta mirrors the callCommand: full-cell-text identity, plain-cell guard). */
  function applyCell({ targetText, tableAnchor, text }) {
    const n = x => String(x).replace(/\s+/g, ' ').trim();
    const cellText = c => n(cellRead(c).text);
    // GH #291 N3: GetAllTables() lists the nested tables too (a holding cell's text includes theirs)
    const allTables = ts => ts.flatMap(t => [t, ...allTables(t.rows.flat().flatMap(c => (Array.isArray(c.nested) ? c.nested : [])))]);
    let tables = allTables(liveDoc.tables);
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
    // GH #291: a write that breaks the table's vertical merge (owner rowspan gone, the continuation a plain cell) -- L3's HTML-span comparison must catch it
    if (f === 'cell-break-vmerge') breakVMerge(liveDoc.tables.find(x => x.rows.flat().includes(cell)));
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
      const at = side ? ri : ri + 1;
      // GH #291 T-05 (LIVE): a new row INSIDE a merge's span extends it -- the owner's rowspan grows, the new row's cell in that column is a
      // continuation; 'row-insert-merge-not-extended' models an editor that left the merge as it was (a plain cell inside the span)
      if (f !== 'row-insert-merge-not-extended') t.rows.forEach((r0, r) => r0.forEach((c, k) => { if (c.rowspan > 1 && r < at && at <= r + c.rowspan - 1) { c.rowspan += 1; row[k] = { span: 1, vContinue: true, paras: [{ text: '', props: { ...PLAIN } }] }; } }));
      t.rows.splice(at, 0, row);
      if (f === 'row-twice') t.rows.splice(side ? ri : ri + 1, 0, blank());
    } else {
      const gone = f === 'row-delete-other' ? (ri + 1 < t.rows.length ? ri + 1 : ri - 1) : ri;
      // GH #291 T-06 (LIVE): a deleted continuation row shortens every merge covering it (a two-row merge ends);
      // 'row-delete-merge-not-shortened' models an editor that left the owner's rowspan as it was
      if (f !== 'row-delete-merge-not-shortened') t.rows.forEach((row, r) => row.forEach(c => { if (c.rowspan > 1 && r < gone && gone <= r + c.rowspan - 1) { c.rowspan -= 1; if (c.rowspan === 1) delete c.rowspan; } }));
      t.rows.splice(gone, 1);
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
  // GH #291 H5m: the driver's physical cell index skips a grid gap (it is no cell in the editor); the fake keeps gaps as cells
  const physIndex = (row, k) => { let n = -1; for (let i = 0; i < (row || []).length; i += 1) if (!row[i].gap) { n += 1; if (n === k) return i; } return -1; };
  const blankLike = src => ({ span: 1, paras: [{ text: '', props: { ...PLAIN }, alignment: src && src.alignment }] });

  function applyMerge({ tableAnchor, startText, endText, start, end, cells }) {
    const { t, cellText } = tableByAnchor(tableAnchor);
    if (!t) return;
    const at = ([r, k]) => (t.rows[r] ? t.rows[r][physIndex(t.rows[r], k)] : null);
    if (!at(start) || !at(end) || cellText(at(start)) !== startText || cellText(at(end)) !== endText) return;
    const f = faultActive() ? script.fault : null;
    if (f === 'ignore') return;
    const picked = cells.map(([r, k0]) => { const k = physIndex(t.rows[r], k0); return { r, k, cell: at([r, k0]), c: startsOf(t.rows[r])[k] }; });
    if (picked.some(x => !x.cell || x.cell.vContinue)) return;
    // GH #291 T-07 (LIVE): a picked vMerge owner covers its continuation cells below -- they go with the merge
    const r1 = Math.min(...picked.map(x => x.r)); const r2 = Math.max(...picked.map(x => x.r + (x.cell.rowspan || 1) - 1));
    const c1 = Math.min(...picked.map(x => x.c)); const c2 = Math.max(...picked.map(x => x.c + (x.cell.span || 1) - 1));
    const coveredBy = (r, k, s0) => picked.some(p => p.r === r && p.k === k) || (t.rows[r][k].vContinue && picked.some(p => p.c === s0 && p.r < r && r <= p.r + (p.cell.rowspan || 1) - 1));
    // a proper rectangle: every logical cell of the box covered by exactly the picked cells (or their continuations), none reaching out
    for (let r = r1; r <= r2; r += 1) {
      const st = startsOf(t.rows[r]);
      const inRow = t.rows[r].map((c, k) => ({ c, k, s: st[k], e: st[k] + (c.span || 1) - 1 })).filter(x => !(x.e < c1 || x.s > c2));
      if (inRow.some(x => x.s < c1 || x.e > c2 || !coveredBy(r, x.k, x.s))) return;
    }
    const ordered = picked.slice().sort((a, b) => a.r - b.r || a.c - b.c);
    let paras = ordered.flatMap(x => x.cell.paras.map(p => ({ ...p, props: { ...p.props } })));
    if (f === 'merge-lose-content') paras = paras.slice(0, 1);
    if (f === 'merge-join-paragraphs') paras = [{ ...paras[0], text: paras.map(p => p.text).join(' ') }];
    if (f === 'merge-reformat' && paras[1]) paras[1] = { ...paras[1], props: { ...paras[1].props, bold: !paras[1].props.bold } };
    const cols = c2 - c1 + 1 - (f === 'merge-wrong-span' ? 1 : 0);
    const rows = r2 - r1 + 1;
    for (let r = r2; r >= r1; r -= 1) {
      const st = startsOf(t.rows[r]);
      const idx = t.rows[r].map((_, k) => k).filter(k => st[k] >= c1 && st[k] <= c2 && coveredBy(r, k, st[k])).sort((a, b) => a - b);
      const pos = idx[0];
      for (let i = idx.length - 1; i >= 0; i -= 1) t.rows[r].splice(idx[i], 1);
      t.rows[r].splice(pos, 0, r === r1 ? { span: cols, ...(rows > 1 ? { rowspan: rows } : {}), paras } : { span: cols, vContinue: true, paras: [{ text: '', props: { ...PLAIN } }] });
    }
    if (f === 'merge-touch-cell') { const c = t.rows.flat().find(x => !x.vContinue && cellText(x) !== '' && x.paras !== paras); if (c) c.paras = [{ ...c.paras[0], props: { ...c.paras[0].props, italic: !c.paras[0].props.italic } }]; }
    if (f === 'merge-touch-other-table') { const o = liveDoc.tables.find(x => x !== t); if (o) o.rows[0][0].paras = [{ ...o.rows[0][0].paras[0], text: `${o.rows[0][0].paras[0].text} (edited)` }]; }
    dirty = true;
  }

  /**
   * W05-3: TWO cases, matching table-merge-split.cjs's own two classifySplit branches.
   *   RESTORE (rs===rows && cs===cols, rs*cs>1): UNCHANGED logic, below.
   *   NEW SPLIT (rs===cs===1, a never-merged cell): MEASURED (W05-1a/1b) -- the grid GROWS; the
   *     target's row gets `cols` new 1x1 cells (content kept in the first); if rows>1, every
   *     OTHER cell in the target's own row extends via a `rowspan` (vMerge owner) across the
   *     newly-inserted rows, each of which gets a `vContinue` placeholder in that same column
   *     slot -- the exact shape W05-1b measured ("a SZOMSZED oszlop vMerge restart/continue-val
   *     egyesul"). Width/tcW is OUT of scope (point 2 of docs/W05-SPLIT-GEOMETRY-DESIGN-
   *     20261007.md v2) -- this model only tracks logical cell/row/span structure, never widths.
   */
  function applySplit({ tableAnchor, targetText, cell, rows, cols }) {
    const { t, cellText } = tableByAnchor(tableAnchor);
    if (!t) return;
    const [r, k0] = cell;
    const k = physIndex(t.rows[r], k0);
    const owner = t.rows[r] && t.rows[r][k];
    if (!owner || cellText(owner) !== targetText) return;
    const f = faultActive() ? script.fault : null;
    if (f === 'ignore') return;
    const rs = owner.rowspan || 1; const cs = owner.span || 1;
    const c0 = startsOf(t.rows[r])[k];
    const src = owner.paras[0];
    const newCells = n => Array.from({ length: n }, () => blankLike(src));
    if (rs === rows && cs === cols) {
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
      return;
    }
    // GH #291 S3 (LIVE): a merged cell split into fewer pieces than its span -- each piece an equal share of it, the content top-left
    if (rs * cs > 1 && rs % rows === 0 && cs % cols === 0 && (rs !== rows || cs !== cols)) {
      // fault 'split-partial-1x1-pieces': the pieces come out 1x1 (a wrong-size partial split the L1 piece check must catch)
      const pr = f === 'split-partial-1x1-pieces' ? 1 : rs / rows; const pc = f === 'split-partial-1x1-pieces' ? 1 : cs / cols;
      const piece = (keep, top) => ({ span: pc, ...(pr > 1 && top ? { rowspan: pr } : {}), ...(top ? {} : { vContinue: true }), paras: keep ? owner.paras.map(p => ({ ...p, props: { ...p.props } })) : [{ text: '', props: { ...PLAIN } }] });
      t.rows[r].splice(k, 1, ...Array.from({ length: cols }, (_, j) => piece(j === 0, true)));
      for (let i = 1; i < rs; i += 1) {
        const row = t.rows[r + i];
        const kk = startsOf(row).findIndex((s0, j) => s0 === c0 && row[j].vContinue);
        if (kk >= 0) row.splice(kk, 1, ...Array.from({ length: cols }, () => piece(false, i % pr === 0)));
      }
      dirty = true;
      return;
    }
    if (rs === 1 && cs === 1 && rows * cols >= 2) {
      const kept = { span: 1, paras: owner.paras.map(p => ({ ...p, props: { ...p.props } })) };
      if (f === 'split-join-paragraphs') kept.paras = [{ ...kept.paras[0], text: kept.paras.map(p => p.text).join(' ') }];
      if (f === 'split-reformat') kept.paras[0] = { ...kept.paras[0], props: { ...kept.paras[0].props, bold: !kept.paras[0].props.bold } };
      let firstRowNew = [kept, ...newCells(cols - 1)];
      if (f === 'split-content-wrong-cell' && firstRowNew.length > 1) firstRowNew = [firstRowNew[1], firstRowNew[0], ...firstRowNew.slice(2)];
      t.rows[r].splice(k, 1, ...firstRowNew);
      // MEASURED (W05-1a): splitting r0c0 into 1x2 on a 2x2 table left row 1's FIRST cell at
      // gridSpan=2 -- the OTHER rows' cell at this same LOGICAL column grows by (cols-1) to
      // cover the widened grid. Found by logical start position (startsOf), not physical index
      // (which may differ between rows of different spans).
      if (cols > 1) {
        for (let rowIdx = 0; rowIdx < t.rows.length; rowIdx += 1) {
          if (rowIdx === r) continue;
          const row = t.rows[rowIdx];
          const starts = startsOf(row);
          const j = starts.findIndex((s0, idx) => s0 === c0 && !row[idx].vContinue);
          if (j >= 0) row[j] = { ...row[j], span: (row[j].span || 1) + (cols - 1) };
        }
      }
      if (rows > 1) {
        for (let j = 0; j < t.rows[r].length; j += 1) {
          if (j >= k && j < k + cols) continue; // the new cells themselves -- not a neighbor
          t.rows[r][j] = { ...t.rows[r][j], rowspan: rows };
        }
        for (let i = 1; i < rows; i += 1) {
          const newRow = t.rows[r].map((neighbor, j) => (j >= k && j < k + cols
            ? blankLike(src)
            : { span: neighbor.span || 1, vContinue: true, paras: [{ text: '', props: { ...PLAIN } }] }));
          t.rows.splice(r + i, 0, newRow);
        }
      }
      // W05-4 (#ref, a maintainer's named weakness): the NEW-split L3 check compares "outside" owners
      // by a TEXT-ONLY, ORDER-INDEPENDENT multiset (table-merge-split.cjs verifyPreservation) --
      // which cannot see TWO outside cells trading their text with each other: the SET of texts
      // is unchanged, only which cell holds which. Modelled here so the weakness is MEASURED
      // (the canary offline test asserts this fault currently PASSES L3), not just asserted.
      if (f === 'outside-swap') {
        const row = t.rows[r];
        const leftIdx = 0;
        const rightIdx = row.length - 1;
        if (leftIdx < k && rightIdx >= k + cols && row[leftIdx] && row[rightIdx] && row[leftIdx] !== row[rightIdx]) {
          const a = row[leftIdx].paras[0].text;
          const b = row[rightIdx].paras[0].text;
          row[leftIdx] = { ...row[leftIdx], paras: [{ ...row[leftIdx].paras[0], text: b }] };
          row[rightIdx] = { ...row[rightIdx], paras: [{ ...row[rightIdx].paras[0], text: a }] };
        }
      }
      dirty = true;
    }
    // any other shape mismatch -- refused before dispatch; not modelled (guard miss -> unchanged)
  }

  /**
   * Wave 5: a header/footer write (meta mirrors the callCommand, which re-checks the plan's guards;
   * header/footer probes 1-2): the whole content becomes ONE plain paragraph (default run props,
   * left alignment); a created part exists only on that section. Faults model the negative controls.
   */
  function applyHeaderFooter({ part, variant, sectionIndex, sectionCount, create, text, activate = false, mode, paragraphIndex, paragraphCount, oldText, inheritedBy = [], pageNumber = false }) {
    // H-07a: the sections inheriting the part (the later ones up to the first with its own) -- exactly the named ones
    const inheritsOk = () => { const aff = []; for (let j = sectionIndex + 1; j < liveDoc.sections.length && !liveDoc.sections[j][part][variant]; j += 1) aff.push(j); return JSON.stringify(aff) === JSON.stringify(inheritedBy); };
    const secs = liveDoc.sections;
    if (secs.length !== sectionCount) return;
    const sec = secs[sectionIndex];
    // GH #289 paragraph mode: the named paragraph's one text run gets the text; everything else stays
    if (mode === 'paragraph') {
      if (variant === 'title' && !(sec && sec.titlePage === true)) return;
      if (variant === 'even' && liveDoc.evenAndOdd !== true) return;
      const own = sec && sec[part][variant];
      if (!own || own.nonPlain) return;
      if (!inheritsOk()) return;
      if (own.paras.length !== paragraphCount) return;
      const x = own.paras[paragraphIndex];
      const textOf = y => (Array.isArray(y.runs) ? y.runs.map(r => r.text).join('') : y.text);
      if (String(textOf(x)).replace(/\s+/g, ' ').trim() !== oldText) return;
      const f = faultActive() ? script.fault : null;
      if (f === 'ignore') return;
      const target = f === 'hf-para-other-paragraph' && own.paras.length > 1 ? own.paras[paragraphIndex === 0 ? 1 : 0] : x;
      if (Array.isArray(target.runs)) {
        const nonEmpty = target.runs.filter(r => r.text !== '');
        if (nonEmpty.length !== 1) return;
        nonEmpty[0].text = text;
      } else target.text = text;
      if (f === 'hf-para-drop-format') { target.props = { ...PLAIN }; target.alignment = 'left'; delete target.style; }
      if (f === 'hf-para-drop-inline') target.inline = [];
      if (f === 'hf-para-touch-table' && Array.isArray(own.tables) && own.tables.length) own.tables[0].rows[0][0] = `${own.tables[0].rows[0][0]} (edited)`;
      dirty = true;
      return;
    }
    // M5: the command re-reads the variant's flag (internal, read only) and refuses a changed state
    if (variant === 'title' && activate !== (sec && sec.titlePage !== true)) return;
    if (variant === 'even' && activate !== (liveDoc.evenAndOdd !== true)) return;
    if (!sec || !!sec[part][variant] === create) return;
    if (secs.length > 1) {
      if (!inheritsOk()) return;
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
    else sec[part][variant] = pageNumber && f !== 'hf-page-number-missing' ? { nonPlain: false, paras: [{ text: '', runs: [{ text }, { text: '' }], inline: ['pgNum'], props: { ...PLAIN }, alignment: 'left' }] } : fresh(text);
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
   * GH #268c: an image write (meta mirrors the callCommand -- image.cjs's own commandBody re-checks
   * the resolved index and, when the selector named one, the expected parent paragraph; for a wrap,
   * that the SOURCE is still inline, the one LIVE-measured transition (R4). Faults model the
   * negative controls: another image changing (preservation), the guard being bypassed ('ignore'
   * models the guard already catching a drifted index/paragraph upstream -- the fake's own guard
   * above does that for real, this flag is for a test that wants NO write at all for another reason).
   */
  function applyImageSize({ imageIndex, expectedParagraph, widthEmu, heightEmu }) {
    const im = liveDoc.images[imageIndex];
    if (!im) return;
    if (expectedParagraph !== null && im.paragraph !== expectedParagraph) return;
    const f = faultActive() ? script.fault : null;
    if (f === 'ignore') return;
    if (f === 'image-other-changed') { const other = liveDoc.images.find((x, i) => i !== imageIndex); if (other) other.cx = (other.cx || 0) + 1; }
    if (f === 'image-body-change' && live.length) live[0].text = `${live[0].text} (edited)`;
    im.cx = widthEmu;
    im.cy = heightEmu;
    dirty = true;
  }

  function applyImageWrap({ imageIndex, expectedParagraph, wrap }) {
    const im = liveDoc.images[imageIndex];
    if (!im) return;
    if (expectedParagraph !== null && im.paragraph !== expectedParagraph) return;
    if (im.drawingType !== 'inline') return;
    const f = faultActive() ? script.fault : null;
    if (f === 'ignore') return;
    if (f === 'image-other-changed') { const other = liveDoc.images.find((x, i) => i !== imageIndex); if (other) { other.drawingType = 'anchor'; other.wrapType = 'square'; } }
    if (f === 'image-body-change' && live.length) live[0].text = `${live[0].text} (edited)`;
    im.drawingType = 'anchor';
    im.wrapType = wrap;
    dirty = true;
  }

  /**
   * GH #269d: a reply / a comment on a range (meta mirrors the callCommand: target re-found by text + author + quote with EXACTLY one match; the range count re-checked against
   * the planned total). The author is the SESSION identity -- the command passes no author. Faults model the negative controls: nothing landing ('ignore'), the author EMPTY
   * (a forged NAME is not detectable offline: the product never sets one, pinned statically; LIVE compares the saved author with the session identity), the write applied TWICE, a sibling comment changing.
   */
  function applyReplyComment({ comment, text }) {
    const str = v => (v === null || v === undefined ? '' : String(v));
    const eol = v => str(v).replace(/[\r\n]+$/, '');
    const hits = liveDoc.comments.filter(c => str(c.text) === comment.text && str(c.author) === comment.author && eol(c.quote) === eol(comment.quote));
    if (hits.length !== 1) return;
    const f = faultActive() ? script.fault : null;
    if (f === 'ignore') return;
    const who = { ...SESSION, ...(script.session || {}) };
    const reply = { text, author: f === 'cw-empty-author' ? '' : who.currentName, userId: who.documentUserId };
    hits[0].replies.push(reply);
    if (f === 'cw-duplicate') hits[0].replies.push({ ...reply });
    if (f === 'cw-other-comment') { const o = liveDoc.comments.find(c => c !== hits[0]); if (o) o.solved = !o.solved; }
    dirty = true;
  }

  function applyAddComment({ matchText, occurrence, total, text }) {
    const where = [];
    live.forEach((p, i) => { for (let at = String(p.text).indexOf(matchText); at >= 0; at = String(p.text).indexOf(matchText, at + matchText.length)) where.push(i); });
    if (where.length !== total || where[occurrence - 1] === undefined) return;
    const f = faultActive() ? script.fault : null;
    if (f === 'ignore') return;
    const who = { ...SESSION, ...(script.session || {}) };
    const nextId = String(liveDoc.comments.reduce((m, c) => Math.max(m, Number(c.id) || 0), 0) + 1);
    const row = { id: nextId, text, author: f === 'cw-empty-author' ? '' : who.currentName, userId: who.documentUserId, timeUtc: 1760000009000, quote: matchText, solved: false, replies: [] };
    liveDoc.comments.push(row);
    if (f === 'cw-duplicate') liveDoc.comments.push({ ...row, id: String(Number(nextId) + 1) });
    if (f === 'cw-other-comment') { const o = liveDoc.comments.find(c => c !== row); if (o) o.solved = !o.solved; }
    dirty = true;
  }

  /**
   * GH #269c: a comment SetSolved (meta mirrors the callCommand: re-finds the comment by text + author +
   * quote, requires EXACTLY one match, calls SetSolved only). Faults model the negative controls: the
   * write silently not landing ('ignore'), a SIBLING comment's solved state flipping, the target's
   * author being forged (the SetAuthorName surface the product must never touch), the wrong direction.
   */
  function applyResolveComment({ comment, resolved }) {
    const str = v => (v === null || v === undefined ? '' : String(v));
    const eol = v => str(v).replace(/[\r\n]+$/, '');
    const hits = liveDoc.comments.filter(c => str(c.text) === comment.text && str(c.author) === comment.author && eol(c.quote) === eol(comment.quote));
    if (hits.length !== 1) return;
    const f = faultActive() ? script.fault : null;
    if (f === 'ignore') return;
    hits[0].solved = f === 'rc-wrong-direction' ? !resolved : resolved;
    if (f === 'rc-other-comment') { const o = liveDoc.comments.find(c => c !== hits[0]); if (o) o.solved = !o.solved; }
    if (f === 'rc-author-forged') hits[0].author = 'forged';
    dirty = true;
  }

  /**
   * W11: a page-setup write (meta mirrors the callCommand, which re-checks the section count/index
   * -- commandBody() in families/page-setup.cjs). Only the requested fields change; the ACT order
   * matches the real call order (size, margins, columns, startPageNumber). Faults model negative
   * controls: a sibling field on the SAME section changing, another section changing, the margins
   * landing in the WRONG positions (the exact bug the real order-discrimination probe caught,
   * #ref -- this fault proves the offline test would catch a regression of that same mistake).
   */
  function applyPageSetup({ sectionIndex, sectionCount, fields }) {
    const secs = liveDoc.sections;
    if (secs.length !== sectionCount) return;
    const sec = secs[sectionIndex];
    if (!sec) return;
    const f = faultActive() ? script.fault : null;
    if (f === 'ignore') return;
    if (fields.size) { sec.pageWidth = fields.size.width; sec.pageHeight = fields.size.height; }
    if (fields.margins) {
      // the CORRECT order is (left, top, right, bottom) at the real call sites; the fault
      // below writes the task's field NAMES into the WRONG getter slots, modelling exactly the bug
      // that measuring with four IDENTICAL values could not have caught.
      if (f === 'ps-margins-swapped') sec.margins = { top: fields.margins.left, right: fields.margins.top, bottom: fields.margins.right, left: fields.margins.bottom };
      else sec.margins = { ...fields.margins };
    }
    if (fields.columns) {
      sec.columnsCount = fields.columns.count;
      // W11-4: space IS independently re-readable (GetColumnsSpaces) -- the fault below
      // corrupts ONLY the space array, modelling the exact false-SATISFIED bug a same-count/
      // different-space retry must NOT fall for.
      const wrongSpace = f === 'ps-columns-space-wrong' ? fields.columns.space + 1 : fields.columns.space;
      sec.columnsSpaces = Array.from({ length: Math.max(0, fields.columns.count - 1) }, () => wrongSpace);
    }
    if (fields.startPageNumber !== undefined) sec.startPageNumber = fields.startPageNumber;
    if (f === 'ps-other-field-changed') sec.startPageNumber = (sec.startPageNumber || 1) + 1000; // a sibling field on the SAME section, not requested
    if (f === 'ps-other-section' && secs.length > 1) { const j = sectionIndex === 0 ? 1 : 0; secs[j].startPageNumber = (secs[j].startPageNumber || 1) + 1000; }
    if (f === 'ps-body-change' && live.length) live[0].text = `${live[0].text} (edited)`;
    dirty = true;
  }

  /** W-10 (style DEFINITIONS): `mode` is 'define' or 'modify' -- mirrors style-definition.cjs's
   * commandBody() defensive no-op shape (define never overwrites, modify never creates).
   *
   * UNIT BOUNDARY (after the LIVE size-unit mismatch): internally this fake
   * stores `text.size` in HALF-POINTS and `paragraph.indentLeft`/`paragraph.spacingAfter` in
   * TWIPS -- the SAME raw units the real driver's saved file / ToJSON carries -- NOT the
   * task-facing PT the family's `fields` use. This mirrors `commandBody`'s own write-side
   * conversion (`size * 2`, `indentLeft * 20`, `spacingAfter * 20`) and the driver's read-side
   * conversion (`/ 2`, `/ 20`) -- WITHOUT this boundary, the fake's round-trip is trivially
   * consistent (same value in, same value out) and CANNOT exercise the real unit conversion,
   * which is exactly why the real size-unit bug shipped invisibly through 14 offline tests. */
  function applyStyleDefinition({ mode, name, fields }) {
    const f = faultActive() ? script.fault : null;
    if (f === 'ignore') return;
    const exists = Object.prototype.hasOwnProperty.call(liveDoc.styles, name);
    if (mode === 'define') {
      if (exists) return;
      // GH #292 S3: a missing base writes nothing, the commandBody's own pre-create guard
      // GH #292 S5: and a base of another type writes nothing (SetBasedOn returns false; the commandBody checks GetType first)
      const type = fields.type === 'character' ? 'character' : 'paragraph';
      if (fields.basedOn !== undefined && (!Object.prototype.hasOwnProperty.call(liveDoc.styles, fields.basedOn) || (liveDoc.styles[fields.basedOn].type || 'paragraph') !== type)) return;
      liveDoc.styles[name] = { text: {}, paragraph: {}, ...(type === 'character' ? { type } : {}), ...(fields.basedOn !== undefined && f !== 'sd-basedon-dropped' ? { basedOn: fields.basedOn } : {}) };
      //: doc.GetAllStyles() is MEASURED to never see a style created this way --
      // this fake marks it 'custom' so readLive's style-builder can model the SAME blindness
      // (unless the caller's `request.styleNames` targets it, mirroring buildTargetedStyleQuery).
      liveDoc.customStyleNames.add(name);
    } else if (!exists) {
      return;
    }
    const st = liveDoc.styles[name];
    if (fields.text) {
      const t = { ...fields.text };
      if ('size' in t) t.size = Math.round(t.size * 2); // pt -> half-points (SetFontSize's own unit)
      Object.assign(st.text, t);
    }
    if (fields.paragraph) {
      const p = { ...fields.paragraph };
      if ('indentLeft' in p) p.indentLeft = Math.round(p.indentLeft * 20); // pt -> twips (SetIndLeft's own unit)
      if ('spacingAfter' in p) p.spacingAfter = Math.round(p.spacingAfter * 20); // pt -> twips (SetSpacingAfter's own unit)
      for (const k of ['spacingBefore', 'indentRight', 'indentFirstLine']) if (k in p) p[k] = Math.round(p[k] * 20); // GH #292: pt -> twips
      Object.assign(st.paragraph, p);
    }
    // GH #292 S4: the word /4 Bridge WRITE -- next on the style; a link on BOTH sides (faults: the setter does not take / only one side)
    if (fields.next !== undefined && f !== 'sd-nextlink-dropped') st.next = fields.next;
    if (fields.link !== undefined && f !== 'sd-nextlink-dropped') {
      st.link = fields.link;
      if (f !== 'sd-link-one-side' && liveDoc.styles[fields.link]) liveDoc.styles[fields.link].link = name;
    }
    // SD-other-style-changed: a sibling style changes too (the L3 preservation fault this family's
    // offline acceptance must catch, same class as ps-other-field-changed above).
    if (f === 'sd-other-style-changed') {
      const otherName = Object.keys(liveDoc.styles).find(n => n !== name) || '__ZZZ_SD_OTHER__';
      if (!liveDoc.styles[otherName]) liveDoc.styles[otherName] = { text: {}, paragraph: {} };
      liveDoc.styles[otherName].text.bold = !liveDoc.styles[otherName].text.bold;
    }
    if (f === 'sd-body-change' && live.length) live[0].text = `${live[0].text} (edited)`;
    dirty = true;
  }

  /**
   * W04-4 (table-format, #ref MEASURED shapes): a table-level format write (meta mirrors
   * commandBody() in families/table-format.cjs). `fields` arrives already in call-ready shape
   * (colors pre-split into r/g/b at lower() time, same discipline as every other raw-number
   * write this codebase made since W04-1c -- never a color object). Border edges are stored
   * keyed by their TASK-FACING name (top/bottom/left/right/insideH/insideV); the ToJSON/XML
   * start-end rename is modelled only where it is OBSERVED (the real driver's own
   * ToJSON extraction, and the LIVE acceptance test's raw-XML cross-check) -- not here, since the
   * fake driver's job is to model WRITE PERSISTENCE, not the editor's internal key-naming choice.
   */
  /** GH #291 faults: every vertical merge of `t` undone -- the owner loses its rowspan, its continuations become plain empty cells. */
  function breakVMerge(t) {
    for (const row of (t && t.rows) || []) for (const c of row) { if (c.rowspan) delete c.rowspan; if (c.vContinue) c.vContinue = false; }
    dirty = true;
  }

  function applyTableFormat({ tableAnchor, anchorText, fields }) {
    const n = x => String(x).replace(/\s+/g, ' ').trim();
    const cellText = c => n(c.paras.map(x => x.text).join('\r\n'));
    let tables = liveDoc.tables;
    if (tableAnchor !== null && tableAnchor !== undefined) {
      tables = tables.filter(t => t.rows.flat().some(c => cellText(c) === tableAnchor));
      if (tables.length !== 1) return;
    }
    const hits = tables.filter(t => t.rows.flat().some(c => cellText(c) === anchorText));
    if (hits.length !== 1) return;
    const t = hits[0];
    t.format = t.format || {};
    const f = faultActive() ? script.fault : null;
    if (f === 'ignore') return;
    if (f === 'tf-break-vmerge') breakVMerge(t);
    const BORDER_EDGES = ['top', 'bottom', 'left', 'right', 'insideH', 'insideV'];
    const EDGE_SWAP = { top: 'bottom', bottom: 'top', left: 'right', right: 'left', insideH: 'insideV', insideV: 'insideH' };
    if (fields.borders) {
      t.format.borders = t.format.borders || {};
      for (const edge of BORDER_EDGES) {
        const req = fields.borders[edge];
        if (!req) continue;
        const value = req.style === 'none'
          ? { value: 'none', sz: 0, space: 0, color: { auto: true, r: 0, g: 0, b: 0 } }
          : { value: req.style, sz: req.size, space: 0, color: { auto: false, r: req.r, g: req.g, b: req.b } };
        // the SAME regression class the real start/end rename could have hidden: a border write
        // landing on the WRONG edge is modelled explicitly here, not inferred from the rename.
        const slot = f === 'tf-border-edge-swapped' ? (EDGE_SWAP[edge] || edge) : edge;
        t.format.borders[slot] = value;
      }
    }
    if (fields.shading) t.format.shading = { val: 'clear', fill: { r: fields.shading.r, g: fields.shading.g, b: fields.shading.b }, color: { auto: false, r: fields.shading.r, g: fields.shading.g, b: fields.shading.b } };
    if (fields.cellMargins) {
      t.format.cellMargins = t.format.cellMargins || {};
      for (const edge of ['top', 'right', 'bottom', 'left']) if (fields.cellMargins[edge] !== undefined) t.format.cellMargins[edge] = { type: 'dxa', w: fields.cellMargins[edge] };
    }
    if (fields.indent !== undefined) t.format.indent = fields.indent;
    if (fields.width !== undefined) t.format.width = { type: 'dxa', w: fields.width };
    if (fields.style) {
      const match = (liveDoc.tableStyles || []).find(s => s.name === fields.style.name);
      if (match) t.format.style = String(match.styleId);
    }
    if (f === 'tf-other-field-changed') t.format.indent = (t.format.indent || 0) + 999; // a sibling field on the SAME table, not requested
    if (f === 'tf-other-table') { const o = liveDoc.tables.find(x => x !== t); if (o) { o.format = o.format || {}; o.format.indent = (o.format.indent || 0) + 999; } }
    if (f === 'tf-body-change' && live.length) live[0].text = `${live[0].text} (edited)`;
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
        // GH #291 H4 (LIVE): a boundary at or inside a short row's TRAILING gap puts the new cell before the whole gap
        let g0 = row.length;
        while (g0 > 0 && row[g0 - 1].gap) g0 -= 1;
        if (p > g0) p = g0;
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
        // GH #291 H4 (LIVE): deleting a column of a wider grid gap shrinks the gap by one
        if (row[k].gap && span > 1) { row[k].span = span - 1; return; }
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
  function applyBookmark({ targetText, name, matchText }) {
    if (inNote) return;
    const hits = exact(targetText);
    if (hits.length !== 1 || liveDoc.bookmarks.some(b => b.name === name)) return;
    // GH #266: matchText = a sub-range of the ONE parent paragraph, found exactly once inside it
    if (matchText !== undefined && occurrences(hits[0].text, matchText) !== 1) return;
    const f = faultActive() ? script.fault : null;
    if (f === 'ignore') return;
    const other = live.find(x => x !== hits[0]);
    const covered = matchText !== undefined ? matchText : targetText;
    // a sub-range bookmark remembers its parent (`parent`, stripped from every read) so the marker count stays on the right paragraph
    const entry = t => ({ name, text: t, ...(matchText !== undefined ? { parent: targetText } : {}) });
    liveDoc.bookmarks.push(entry(f === 'bookmark-wrong-target' && other ? other.text : (f === 'bookmark-whole-parent' ? targetText : covered)));
    if (f === 'bookmark-twice') liveDoc.bookmarks.push(entry(covered));
    dirty = true;
  }

  /** Wave 3: a hyperlink over the target paragraph, its text properties restored (source parity). */
  function applyLink({ targetText, url, tooltip, matchText }) {
    if (inNote) return;
    const hits = exact(targetText);
    if (hits.length !== 1) return;
    if (matchText !== undefined) {
      // GH #266: a link over a sub-range of the parent; the parent keeps its other links
      if (occurrences(hits[0].text, matchText) !== 1 || !tooltip) return;
      const g = faultActive() ? script.fault : null;
      if (g === 'ignore') return;
      const q = hits[0];
      q.links = [...(q.links || []), { url: g === 'link-wrong-url' ? `${url}/wrong` : url, tip: tooltip, text: g === 'link-whole-parent' ? q.text : matchText }];
      if (g === 'link-no-restore') q.props = { ...q.props, underline: true };
      if (g === 'link-drop-other-link' && q.links.length > 1) q.links.shift();
      dirty = true;
      return;
    }
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
    /**
     * GH #220 1b: the GH #100 U1/U3 warm-session surface, mirroring the REAL driver's own semantics
     * (docx-driver.cjs's release/resetForNextCall/health/documentKey) on this fake's `dirty`/`docxConfirmedAt`/
     * `live` state -- NOT a thin always-true mock: `release()`'s confirmation follows the SAME
     * `script.saveObserver`/`docxConfirmedAt` rule `close()` already uses above, `health().canSave` is the
     * SAME `dirty` flag `canSave()` reports, and `documentKey()` is a test seam (`script.documentKey`, a
     * value or a function) so a test can simulate an external rotation mid-run.
     */
    async release() {
      calls.release += 1;
      if (script.releaseFails) return { ok: false, detail: 'fake docx release refused' };
      const confirmation = !script.saveObserver || docxConfirmedAt === null ? null
        : (calls.dispatch > docxConfirmedAt ? { state: 'unconfirmed', reason: 'superseded-by-a-newer-save' } : { state: 'confirmed', reason: 'forcesave-confirmed' });
      return confirmation ? { ok: true, detail: 'session-kept', confirmation } : { ok: true, detail: 'session-kept' };
    },
    resetForNextCall() { calls.reset += 1; return { ok: !!live }; },
    async health() {
      calls.health += 1;
      if (script.healthThrows) throw new Error('fake docx health threw');
      return { alive: !!live, canSave: live ? dirty : null, ds: { open: true, sockets: 1, busy: false }, anomalies: 0 };
    },
    documentKey() {
      calls.documentKey += 1;
      if (typeof script.documentKey === 'function') return script.documentKey();
      return script.documentKey !== undefined ? script.documentKey : 'fake-docx-key-1';
    },
    async readLive(_handle, request = {}) {
      calls.readLive += 1;
      const requestedStyleNames = Array.isArray(request.styleNames) ? request.styleNames : [];
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
      // W-07: the per-level fine-grained format, shared per numId+level, mirroring levelFormats
      // above. W07-7/W07-8 (MEASURED on the real editor): an absent `start`
      // key represents 1 -- mirrored here (ONLY for `start`) so offline tests exercise the SAME
      // classify()/verifyPrimitive() behaviour the real driver now produces. `script.
      // numberingFormatFullyUnreadable` (a SEPARATE, explicit flag from the pre-existing
      // `numberingFormatUnreadable`) opts a test OUT of this default, for the family's own
      // defensive "genuinely unobservable" path -- never silently lost.
      const formatOfNumbering = p => {
        if (script.numberingFormatUnreadable) return undefined;
        const arr = liveDoc.numberingFormats[p.list.numId];
        const slot = arr ? arr[p.list.level] : undefined;
        if (script.numberingFormatFullyUnreadable) return slot || undefined;
        const base = slot || {};
        return Object.prototype.hasOwnProperty.call(base, 'start') ? base : { ...base, start: 1 };
      };
      // GH #293 N5: the definition identity (the abstractNum key) beside the format; each numId has its own unless
      // `liveDoc.abstractOf` maps several numIds to one (Word's shared definition); script.abstractUnreadable drops it
      const abstractOf = p => (script.abstractUnreadable ? null : ((liveDoc.abstractOf && liveDoc.abstractOf[p.list.numId]) || `abs-${p.list.numId}`));
      const numberingOf = p => (script.numberingUnreadable ? {} : { numbering: p.list ? { present: true, level: p.list.level, numId: p.list.numId, levelFormats: levelFormatsOf(p), ...(request.numberingFormat === true ? { format: formatOfNumbering(p), abstractId: abstractOf(p) } : {}) } : { present: false } });
      // M5 gate 8: the internal calculated grid (default: equal widths over the logical width)
      const gridOfTable = t => (script.gridUnreadable ? undefined : Array.isArray(t.gridCalc) ? [...t.gridCalc] : Array((t.rows[0] || []).reduce((a, c) => a + (c.span || 1), 0)).fill(20));
      // W04-4: the table-format read. `script.tableFormatUnreadable` -> the WHOLE field
      // UNOBSERVABLE (models ToJSON itself missing/throwing); `script.tableBorderEdgeUnreadable`
      // -> isolates ONE edge (models a sub-key extraction failure without faking the whole table).
      const formatOfTable = t => {
        if (script.tableFormatUnreadable) return 'UNOBSERVABLE:no-such-getter';
        const f = JSON.parse(JSON.stringify(t.format || {}));
        if (script.tableBorderEdgeUnreadable) { f.borders = f.borders || {}; f.borders[script.tableBorderEdgeUnreadable] = 'UNOBSERVABLE:no-such-getter'; }
        return f;
      };
      // Wave 3: a link and each bookmark marker are non-run children (live: 'hyperlink' / 'unsupported')
      const markers = p => (p.links ? p.links.length : 0) + 2 * liveDoc.bookmarks.filter(b => (b.parent !== undefined ? b.parent : b.text) === p.text).length;
      const richOf = p => (script.richUnreadable ? {} : { hyperlinks: (p.links || []).map(l => ({ url: l.url, tip: l.tip, text: l.text !== undefined ? l.text : p.text, style: 'Hyperlink' })), noteRefs: { footnote: 0, endnote: 0, ...(p.refs || {}) } });
      // GH #227: the B21 opt-in/runtime-gate contract the real driver's sectionVariantQuery/
      // __eoB21Variant produce -- three NAMED, mutually exclusive reasons a flag can be unread,
      // mirroring the real driver's own `UNAVAILABLE:<reason>` strings exactly (docx-driver.cjs
      // sectionVariantQuery / __eoB21Unavailable), never a generic fallback:
      //   request.includeSectionVariantState===false   -> 'UNAVAILABLE:not-requested' (opt-out)
      //   script.sectionVariantRuntimeNotVerified       -> 'UNAVAILABLE:runtime-not-verified' (gate closed)
      //   script.flagsUnreadable (pre-existing control)  -> 'UNAVAILABLE:bridge-unavailable' (gate open, bridge call failed)
      const variantFlagsOf = sec => {
        const unavailable = reason => ({ titlePage: `UNAVAILABLE:${reason}`, evenAndOdd: `UNAVAILABLE:${reason}` });
        if (request.includeSectionVariantState === false) return unavailable('not-requested');
        if (script.sectionVariantRuntimeNotVerified) return unavailable('runtime-not-verified');
        if (script.flagsUnreadable) return unavailable('bridge-unavailable');
        return { titlePage: sec.titlePage === true, evenAndOdd: liveDoc.evenAndOdd === true };
      };
      const paragraphs = live.map(p => {
        const rp = script.propsUnreadable ? {} : runPayload(p);
        if (!script.propsUnreadable && markers(p)) { rp.nonRunChildren += markers(p); }
        return { text: p.text, props: script.propsUnreadable ? null : p.props, format: script.propsUnreadable ? undefined : formatOf(p), ...rp, ...numberingOf(p), ...richOf(p) };
      });
      const document = script.docUnreadable ? null : { ...(Array.isArray(request.numberingMembers) && request.numberingMembers.length ? { numberingMembers: Object.fromEntries(request.numberingMembers.map(t => [t, membersOf(t)])) } : {}),
        bookmarks: liveDoc.bookmarks.map(b => ({ name: b.name, text: b.text })), footnotes: [...liveDoc.footnotes], endnotes: [...liveDoc.endnotes], // GH #269b/#269c: the opt-in comment read (request.includeComments); absent = unread, never an empty list.
        ...(request.includeComments === true && !script.commentsUnreadable ? { comments: liveDoc.comments.map(c => ({ ...c, replies: c.replies.map(r => ({ ...r })) })) } : {}),
        footnoteChains: liveDoc.footnotes.map(t => ({ paragraphs: [t], complete: true, stop: 'null' })), endnoteChains: liveDoc.endnotes.map(t => ({ paragraphs: [t], complete: true, stop: 'null' })), inMainBody: !inNote,
        ...(script.tablesUnreadable ? {} : { tables: liveDoc.tables.map(t => ({ rows: t.rows.map(r => r.filter(c => !c.gap).map(cellRead)), gridCalc: gridOfTable(t), format: formatOfTable(t) })).filter((_, i) => !(script.tableReadMissesOne && i === 0)) }),
        // W04-4: the style NAME -> id list `style` resolves against before dispatch.
        tableStyles: script.tableStylesUnreadable ? 'UNOBSERVABLE:no-such-getter' : (liveDoc.tableStyles || DEFAULT_TABLE_STYLES).map(s => ({ ...s })),
        // GH #268b: the images as the driver's page-side projection returns them (a scripted failure = the named string, never an empty list)
        images: script.imagesUnreadable ? 'UNOBSERVABLE:no-GetAllImages' : (liveDoc.images || []).map(i => ({ ...i })),
        // M5 gate 6: the document default text properties (the replacement's run properties = PLAIN)
        defaultText: script.defaultTextUnreadable ? null : { fontFamily: 'Arial', fontSize: 22 },
        ...(script.sectionsUnreadable ? {} : { sections: liveDoc.sections.map(sec => ({ header: Object.fromEntries(HF_VARIANTS.map(v => [v, partRead(sec.header[v])])), footer: Object.fromEntries(HF_VARIANTS.map(v => [v, partRead(sec.footer[v])])),
          ...variantFlagsOf(sec),
          // W11: page setup read, per-field unobservable (never a default number) when scripted.
          ...(script.pageSetupUnreadable ? { pageWidth: 'UNOBSERVABLE:no-such-getter', pageHeight: 'UNOBSERVABLE:no-such-getter', margins: { top: 'UNOBSERVABLE:no-such-getter', right: 'UNOBSERVABLE:no-such-getter', bottom: 'UNOBSERVABLE:no-such-getter', left: 'UNOBSERVABLE:no-such-getter' }, columnsCount: 'UNOBSERVABLE:no-such-getter', columnsSpaces: 'UNOBSERVABLE:no-such-getter', startPageNumber: 'UNOBSERVABLE:no-such-getter' }
            // W11-4: columnsSpaces -- a script.columnsSpacesUnreadable flag isolates JUST
            // this one getter as unobservable, so an offline test can exercise the narrower case
            // (count readable, space not) without faking the whole page-setup block unreadable.
            : { pageWidth: sec.pageWidth, pageHeight: sec.pageHeight, margins: { ...sec.margins }, columnsCount: sec.columnsCount, columnsSpaces: script.columnsSpacesUnreadable ? 'UNOBSERVABLE:no-such-getter' : [...sec.columnsSpaces], startPageNumber: sec.startPageNumber }),
        })) }),
        // W-10: style definitions, keyed by name. script.styleUnobservableFields: ['group.key', ...]
        // isolates ONE getter as unobservable (the real driver's per-field discipline), without
        // faking the whole styles map unreadable -- same narrowing pattern as columnsSpacesUnreadable.
        //
        //: GetAllStyles()-blindness for session-created ('custom') styles is modeled
        // here -- a custom name is enumerated ONLY if it is ALSO in `requestedStyleNames` (the
        // SAME targeted-read mechanism buildTargetedStyleQuery provides on the real driver).
        // Without this, `define_style`'s own offline acceptance tests would never have caught
        // the real bug either -- exactly the gap zero asked be closed.
        ...(script.stylesUnreadable ? {} : { styles: {
          ...Object.fromEntries(Object.entries(liveDoc.styles).filter(([name]) => !liveDoc.customStyleNames.has(name) || requestedStyleNames.includes(name)).map(([name, st]) => {
            const unobs = (script.styleUnobservableFields || []).includes.bind(script.styleUnobservableFields || []);
            const field = (group, key) => (unobs(`${group}.${key}`) ? 'UNOBSERVABLE:no-such-getter' : st[group][key]);
            //: the SAME read-side conversion the real driver applies (half-points/
            // twips -> pt) -- 'size' in group 'text' is stored in half-points, 'indentLeft'/
            // 'spacingAfter' in group 'paragraph' in twips; UNOBSERVABLE/'inherited' pass through.
            const convert = (v, divisor) => (typeof v === 'number' ? v / divisor : v);
            // GH #292: the further keys read 'inherited' when unset -- the real driver's value for an absent ToJSON key
            const newField = (group, key) => { const v = field(group, key); return v === undefined ? 'inherited' : v; };
            return [name, {
              // GH #292 S3: the real ToJSON carries the base's styleId; the fake's styleId is 'id:<name>', the driver-resolved basedOnName the name
              name, identity: { name, basedOn: st.basedOn ? `id:${st.basedOn}` : null, next: st.next ? `id:${st.next}` : null, link: st.link ? `id:${st.link}` : null, styleType: st.type === 'character' ? 'characterStyle' : 'paragraphStyle', basedOnName: st.basedOn || null, nextName: st.next || null, linkName: st.link || null }, // GH #292 S5: the measured ToJSON styleType
              text: { bold: field('text', 'bold'), size: convert(field('text', 'size'), 2), color: field('text', 'color'),
                italic: newField('text', 'italic'), underline: newField('text', 'underline'), strike: newField('text', 'strike'), fontFamily: newField('text', 'fontFamily') },
              paragraph: { align: field('paragraph', 'align'), indentLeft: convert(field('paragraph', 'indentLeft'), 20), spacingAfter: convert(field('paragraph', 'spacingAfter'), 20),
                spacingBefore: convert(newField('paragraph', 'spacingBefore'), 20), indentRight: convert(newField('paragraph', 'indentRight'), 20), indentFirstLine: convert(newField('paragraph', 'indentFirstLine'), 20),
                keepNext: newField('paragraph', 'keepNext'), keepLines: newField('paragraph', 'keepLines'), pageBreakBefore: newField('paragraph', 'pageBreakBefore') },
            }];
          })),
          // (W-10-11): a REQUESTED name the targeted read could not find becomes a
          // null-valued KEY, mirroring the real driver's buildTargetedStyleQuery (`out[name] =
          // null` for an unmatched name) merged by readLive's own union -- the exact shape that
          // made styleExists()'s key-presence-only check (W10-9) wrongly match a MISSING name.
          ...Object.fromEntries(requestedStyleNames.filter(n => !(n in liveDoc.styles)).map(n => [n, null])),
        } }) };
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
      else if (a === 'setNumbering' && command.meta.bridgeListDetach === true) { const r = applyListDetach(command.meta); if (!r.ok) return { ok: false, dispatched: true, outcome: 'b18-bridge-write-failed', evidence: r }; }
      else if (a === 'setNumbering') applySetNumbering(command.meta);
      else if (a === 'bookmark') applyBookmark(command.meta);
      else if (a === 'link') applyLink(command.meta);
      else if (a === 'note') applyNote(command.meta);
      else if (a === 'cell') applyCell(command.meta);
      else if (a === 'row') applyRow(command.meta);
      else if (a === 'column') applyColumn(command.meta);
      else if (a === 'merge') applyMerge(command.meta);
      else if (a === 'split') applySplit(command.meta);
      else if (a === 'headerFooter') applyHeaderFooter(command.meta);
      else if (a === 'imageSetSize') applyImageSize(command.meta);
      else if (a === 'imageSetWrap') applyImageWrap(command.meta);
      else if (a === 'pageSetup') applyPageSetup(command.meta);
      else if (a === 'resolveComment') applyResolveComment(command.meta);
      else if (a === 'replyComment') applyReplyComment(command.meta);
      else if (a === 'addComment') applyAddComment(command.meta);
      else if (a === 'tableFormat') applyTableFormat(command.meta);
      else if (a === 'styleDefinition') applyStyleDefinition(command.meta);
      // M5 constraints: a scripted SIDE EFFECT of the write on the live state (a change the command
      // should not have made -- another paragraph, a table, formatting), for the post-write L3 controls
      if (typeof script.sideEffect === 'function') { script.sideEffect({ live, liveDoc, dispatch: calls.dispatch }); dirty = true; }
      return { ok: true, dispatched: true };
    },
  };
}

module.exports = { createStore, createFakeDocxDriver, htmlOf, formatOf, runPayload };
