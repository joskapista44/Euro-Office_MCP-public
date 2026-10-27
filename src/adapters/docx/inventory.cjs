'use strict';

const crypto = require('node:crypto');

/**
 * DOCX live inventory normalization (M2 `normalizeInventory`) and stable semantic locators.
 *
 * EXTRACT_SEMANTICS from the source (f8c6d6cf, read-only):
 *   coedit-docx-inventory.cjs:inventoryFromHtml   block model: heading / paragraph / list / table
 *   docx-targeting.cjs:decorateStableLocators:33  fingerprint, headingPath, context {before, after}
 *
 * INPUT is the driver's LIVE payload: `{ html, paragraphs: [{ text, props, format, runs }] }`. `html` is the
 * editor's own projection (pluginMethod_GetFileHTML); `paragraphs` are run properties read by the
 * driver's READ-ONLY property query. No package input: the LIVE projection is the primary and only
 * targeting authority here (M4 brief point 8, "no package-first targeting").
 *
 * What is deliberately NOT carried over:
 *   - package heading promotion (source promotePackageHeadingSemantics): a `kind` that flips with
 *     whether the package was read would make the rewrite retry identity read-dependent. `kind`
 *     comes from the LIVE HTML tag only. Package heading semantics, if a family needs them, arrive
 *     through the M2 enrichment hook as additive, non-authoritative evidence.
 *   - any physical index as identity: `index` is kept as a HINT for evidence, never for targeting.
 */

function norm(value) {
  return String(value ?? '').replace(/\s+/g, ' ').trim();
}

function decodeHtml(s) {
  return String(s || '')
    .replace(/&nbsp;/gi, ' ').replace(/&amp;/gi, '&').replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>').replace(/&quot;/gi, '"').replace(/&#39;/gi, "'");
}

function stripTags(s) {
  return decodeHtml(String(s || '').replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]*>/g, ''));
}

function attr(tag, name) {
  const m = String(tag).match(new RegExp(`\\b${name}\\s*=\\s*["']([^"']*)["']`, 'i'));
  return m ? decodeHtml(m[1]) : null;
}

/**
 * Wave 3: the editor projection renders a FOOTNOTE as a reference anchor inside its paragraph
 * (`<a style="mso-footnote-id:ftnN" href="#_ftnN" ...>[N]</a>`) and its body in a trailing
 * `<div style="mso-element:footnote-list">` whose `<p class="MsoFootnoteText">` would otherwise
 * parse as extra paragraphs (measured live 2026-09-24, rich probe 3). Both are representation, not
 * document text: the anchor is removed from the paragraph and the list from the block stream. The
 * notes themselves are read through the driver's document read. Endnotes are not projected at all
 * (same probe: no marker, no list), so there is nothing to strip for them.
 */
function stripFootnoteProjection(html) {
  let s = String(html || '').replace(/<a\b[^>]*mso-footnote-id[^>]*>[\s\S]*?<\/a>/gi, '');
  for (;;) {
    const start = s.search(/<div\b[^>]*mso-element:\s*footnote-list[^>]*>/i);
    if (start < 0) return s;
    const re = /<\/?div\b[^>]*>/gi;
    re.lastIndex = start;
    let depth = 0;
    let end = -1;
    let m;
    while ((m = re.exec(s))) {
      depth += m[0][1] === '/' ? -1 : 1;
      if (depth === 0) { end = re.lastIndex; break; }
    }
    if (end < 0) return s.slice(0, start); // an unbalanced list runs to the end: nothing after it is body text
    s = s.slice(0, start) + s.slice(end);
  }
}

// The index just past the </table> that closes the <table> opened at `start` -- depth-aware, so a NESTED table's own </table> does not end the
// outer one (gnl3a LIVE: the lazy match ended the holding table there, and its remaining rows read as paragraphs OUTSIDE the tables). -1 if unbalanced.
function tableEnd(html, start) {
  const re = /<(\/?)table\b[^>]*>/gi;
  re.lastIndex = start;
  let depth = 0;
  let m;
  while ((m = re.exec(html))) { depth += m[1] ? -1 : 1; if (depth === 0) return re.lastIndex; }
  return -1;
}

// `inner` with every nested <table>...</table> replaced by a placeholder, so the row/cell matching sees only THIS table's own rows; `restore`
// puts the nested HTML back (a holding cell's text keeps the nested table's text, as the driver's cell text does).
function maskNestedTables(inner) {
  const saved = [];
  let out = '';
  let from = 0;
  const re = /<table\b[^>]*>/gi;
  let m;
  while ((m = re.exec(inner))) {
    const end = tableEnd(inner, m.index);
    if (end < 0) break;
    out += `${inner.slice(from, m.index)}\u0000${saved.length}\u0000`;
    saved.push(inner.slice(m.index, end));
    from = end;
    re.lastIndex = end;
  }
  return { masked: out + inner.slice(from), restore: x => x.replace(/\u0000(\d+)\u0000/g, (_, k) => saved[Number(k)]) };
}

function parseBlocks(html) {
  const blocks = [];
  const source = String(html || '');
  const blockRe = /<(h[1-6]|p|ul|ol|table)\b([^>]*)>([\s\S]*?)<\/\1>/gi;
  let m;
  while ((m = blockRe.exec(source))) {
    const tag = m[1].toLowerCase();
    let inner = m[3] || '';
    let restore = x => x;
    if (tag === 'table') {
      const end = tableEnd(source, m.index);
      if (end > 0) {
        const whole = source.slice(m.index, end);
        ({ masked: inner, restore } = maskNestedTables(whole.slice(whole.indexOf('>') + 1, whole.lastIndexOf('</'))));
        blockRe.lastIndex = end;
      }
    }
    if (/^h[1-6]$/.test(tag)) {
      blocks.push({ kind: 'heading', level: Number(tag.slice(1)), text: stripTags(inner) });
    } else if (tag === 'p') {
      blocks.push({ kind: 'paragraph', text: stripTags(inner), style: attr(m[2], 'style') });
    } else if (tag === 'ul' || tag === 'ol') {
      const items = [];
      const liRe = /<li\b[^>]*>([\s\S]*?)<\/li>/gi;
      let li;
      while ((li = liRe.exec(inner))) items.push({ text: stripTags(li[1]) });
      blocks.push({ kind: 'list', ordered: tag === 'ol', items });
    } else {
      const rows = [];
      const rowRe = /<tr\b[^>]*>([\s\S]*?)<\/tr>/gi;
      let r;
      while ((r = rowRe.exec(inner))) {
        const cells = [];
        const cellRe = /<(td|th)\b[^>]*>([\s\S]*?)<\/\1>/gi;
        let c;
        // Wave 4a: the cell's grid spans as projected (a merged cell renders colspan, table probe 1; rowspan read the same way)
        // M5 (row 14): a row's grid gap (w:gridAfter) is projected as an extra, EMPTY cell marked
        // `mso-cell-special:placeholder` that the driver read does not have (gridAfter probe, run
        // ...fcd2df); flagged here, reconciled only by the column family -- absent on every other cell
        while ((c = cellRe.exec(r[1]))) {
          const cell = { text: stripTags(restore(c[2])), span: Number(attr(c[0], 'colspan') || 1), rowspan: Number(attr(c[0], 'rowspan') || 1) };
          if (/mso-cell-special\s*:\s*placeholder/i.test(attr(c[0], 'style') || '')) cell.placeholder = true;
          cells.push(cell);
        }
        rows.push(cells);
      }
      blocks.push({ kind: 'table', rows });
    }
  }
  return blocks;
}

function blockText(block) {
  if (!block) return '';
  if (block.kind === 'list') return norm((block.items || []).map(x => x.text).join(' '));
  if (block.kind === 'table') return norm((block.rows || []).flat().map(x => x.text).join(' '));
  return norm(block.text);
}

function snippet(value, max = 80) {
  const s = norm(value);
  return s.length <= max ? s : s.slice(0, max);
}

function shortHash(value) {
  return crypto.createHash('sha256').update(String(value)).digest('hex').slice(0, 16);
}

/** Stable semantic locator per block (source docx-targeting.cjs:33-59, same field meaning). */
function decorate(blocks) {
  const stack = [];
  for (let i = 0; i < blocks.length; i += 1) {
    const block = blocks[i];
    if (block.kind === 'heading') {
      while (stack.length && stack[stack.length - 1].level >= block.level) stack.pop();
      stack.push({ level: block.level, text: norm(block.text) });
    }
    const headingPath = stack.map(x => ({ ...x }));
    const key = headingPath.map(h => `${h.level}:${h.text}`).join(' > ');
    block.index = i; // a HINT for evidence only -- never used to target
    block.locator = Object.freeze({
      version: 1,
      kind: block.kind,
      fingerprint: shortHash(`${block.kind}|${key}|${blockText(block)}`),
      headingPath,
      text: block.kind === 'table' ? null : norm(block.text),
      context: Object.freeze({
        before: i > 0 ? snippet(blockText(blocks[i - 1])) : '',
        after: i + 1 < blocks.length ? snippet(blockText(blocks[i + 1])) : '',
      }),
    });
  }
  return blocks;
}

/**
 * Run properties are attached to a paragraph/heading block only when the driver's property read
 * named EXACTLY one live paragraph with that text; otherwise `props` is null (unknown), never guessed.
 */
function attachProperties(blocks, paragraphs) {
  const list = Array.isArray(paragraphs) ? paragraphs : [];
  // Wave 2: list items carry the numbering read (level, identity) of the ONE live paragraph with
  // their text; otherwise `numbering` is null (unknown), never guessed.
  const numberingOf = text => {
    const hits = list.filter(p => norm(p.text) === norm(text));
    if (hits.length !== 1 || !hits[0].numbering) return null;
    const n = hits[0].numbering;
    // M5 gate 7: the per-level format read (EO_INTERNAL_DEPENDENCY READ); anything but 9 integers-or-null -> null (unread)
    const lf = Array.isArray(n.levelFormats) && n.levelFormats.length === 9 && n.levelFormats.every(v => v === null || Number.isInteger(v)) ? Object.freeze(n.levelFormats.slice()) : null;
    // W-07: the PUBLIC ToJSON()-derived level format -- each known key re-validated here too (the
    // same discipline as levelFormats above), never trusting the driver's shape unchecked; an
    // unexpected type on any key drops just that key, not the whole object. `start` is an
    // INTEGER-OR-ABSENT key (never coerced to 0/1 when missing -- the driver's own comment on why).
    const fmt = n.format && typeof n.format === 'object' ? Object.freeze({
      ...(typeof n.format.numFmt === 'string' ? { numFmt: n.format.numFmt } : {}),
      ...(typeof n.format.lvlText === 'string' ? { lvlText: n.format.lvlText } : {}),
      ...(typeof n.format.suff === 'string' ? { suff: n.format.suff } : {}),
      ...(typeof n.format.lvlJc === 'string' ? { lvlJc: n.format.lvlJc } : {}),
      ...(Number.isInteger(n.format.start) ? { start: n.format.start } : {}),
    }) : null;
    return Object.freeze({ ...n, levelFormats: lf, format: fmt });
  };
  for (const block of blocks) {
    if (block.kind === 'list') for (const item of block.items || []) item.numbering = numberingOf(item.text);
    if (block.kind === 'paragraph' || block.kind === 'heading') block.numbering = numberingOf(block.text);
  }
  // The same run/format read for a paragraph block and (W2.1) for a list item: a paragraph that a
  // later list step turned into a list item keeps its formatting evidence.
  const attachTo = node => {
    const hits = list.filter(p => norm(p.text) === norm(node.text));
    node.props = hits.length === 1 && hits[0].props ? Object.freeze({ ...hits[0].props }) : null;
    // The whole-paragraph formatting state (formatting family), read the same read-only way.
    node.format = hits.length === 1 && hits[0].format ? Object.freeze({ ...hits[0].format }) : null;
    // W1.1: the run list is formatting EVIDENCE for ranges inside this paragraph -- never a target
    // authority. It is attached only when it provably covers the live paragraph: exactly one live
    // paragraph with this text, no non-run child, and the runs' text is exactly its text. Otherwise
    // `runs` is null and every range claim on this node is unobservable (fail-closed).
    const h = hits.length === 1 ? hits[0] : null;
    const covered = !!h && Array.isArray(h.runs) && h.nonRunChildren === 0 && h.runsCoverText === true
      && h.runs.map(r => String(r.text)).join('') === String(h.text);
    // Wave 3: the paragraph's hyperlinks and note references (unknown = null, never an empty list).
    node.hyperlinks = h && Array.isArray(h.hyperlinks) ? Object.freeze(h.hyperlinks.map(x => Object.freeze({ ...x }))) : (h && typeof h.hyperlinks === 'string' && h.hyperlinks.startsWith('UNOBSERVABLE:') ? h.hyperlinks : null);   // a NAMED unreadable field stays named; absent stays null
    node.noteRefs = h && h.noteRefs && typeof h.noteRefs === 'object' ? Object.freeze({ ...h.noteRefs }) : (h && typeof h.noteRefs === 'string' && h.noteRefs.startsWith('UNOBSERVABLE:') ? h.noteRefs : null);   // a NAMED unreadable count stays named; absent stays null
    node.liveText = covered ? String(h.text) : null;
    node.runs = covered
      ? Object.freeze(h.runs.map(r => Object.freeze({ text: String(r.text), start: r.start, end: r.end, format: Object.freeze({ ...(r.format || {}) }) })))
      : null;
  };
  for (const block of blocks) {
    if (block.kind === 'paragraph' || block.kind === 'heading') attachTo(block);
    if (block.kind === 'list') for (const item of block.items || []) attachTo(item);
  }
  return blocks;
}

/** Images in the live projection (source docx-round-trip-safety.cjs: src / alt / title). */
function parseImages(html) {
  const images = [];
  const re = /<img\b[^>]*>/gi;
  let m;
  while ((m = re.exec(String(html || '')))) images.push({ src: attr(m[0], 'src'), alt: attr(m[0], 'alt'), title: attr(m[0], 'title') });
  return images;
}

/**
 * M8 build gate (owner decision A): per internal dependency, whether THIS session's build is verified for it (the
 * driver's runtime report). Absent or malformed -> null, and every consumer treats null as NOT verified (fail-closed).
 * NARROWED by the W-B1920-C1 bridge cutover: listLevelFormat/tableGrid no longer exist here -- those
 * two reads come from the Word bridge now, fail-closed at the SOURCE (null numbering / null gridCalc on any bridge
 * failure), not through this build-gate mechanism. M11.5d.4: the numbering removal (B18) left the internal path -- the gate now
 * reports whether THIS runtime is the one the B18 Bridge WRITE is proven on (numberingRemoveBridge).
 */
function internalGateOf(runtime) {
  const g = runtime && runtime.internalGate;
  if (!g || typeof g !== 'object') return null;
  return Object.freeze({ runtime: typeof g.runtime === 'string' ? g.runtime : null, numberingRemoveBridge: g.numberingRemoveBridge === true,
    styleNextLinkBridge: g.styleNextLinkBridge === true }); // GH #292 S4: the next / link Bridge WRITE is proven on this runtime (word /4)
}

/** B21 the title-page / even-odd flags: true | false | 'UNAVAILABLE:<why>' (unread, not requested, runtime not verified,
 * reader missing/threw) -- three states, never a null a caller could take for a real one. */
const isUnavailable = v => typeof v === 'string' && v.startsWith('UNAVAILABLE:');
const variantFlag = v => (typeof v === 'boolean' || isUnavailable(v) ? v : 'UNAVAILABLE:unread');
// W04-4 (MEASURED): the table-level tblPr fields the driver already extracted (per field,
// never a default) -- `null` = successfully read as absent (e.g. no border set on that edge),
// the driver's own 'UNOBSERVABLE:<reason>' string = the read itself failed, passed through AS-IS
// so a FAIL_CLOSED reason downstream can name it (same discipline as page-setup's `measured`).
const BORDER_EDGES = Object.freeze(['top', 'bottom', 'left', 'right', 'insideH', 'insideV']);
const MARGIN_EDGES = Object.freeze(['top', 'right', 'bottom', 'left']);
const unobservableOr = (v, fallback) => (typeof v === 'string' ? v : fallback);
function borderVal(v) {
  if (v && typeof v === 'object') return Object.freeze({ value: v.value, sz: v.sz, space: v.space, color: v.color && typeof v.color === 'object' ? Object.freeze({ auto: v.color.auto, r: v.color.r, g: v.color.g, b: v.color.b }) : null });
  return unobservableOr(v, null);
}
function marginVal(v) {
  if (v && typeof v === 'object') return Object.freeze({ type: v.type, w: v.w });
  return unobservableOr(v, null);
}
function widthVal(v) {
  if (v && typeof v === 'object') return Object.freeze({ type: v.type, w: v.w });
  return unobservableOr(v, null);
}
function shadingVal(v) {
  if (v && typeof v === 'object') return Object.freeze({ val: v.val, fill: v.fill, color: v.color });
  return unobservableOr(v, null);
}
/** GH #268b: one driver image row -> the inventory row: placement inline | floating (the ToJSON drawingType inline | anchor), the size in EMU, absent properties null, a failed ToJSON read named on every JSON-derived field. */
function imageRowOf(r) {
  const j = typeof r.json === 'string' ? r.json : null;
  const placement = j || (r.drawingType === 'inline' ? 'inline' : (r.drawingType === 'anchor' ? 'floating' : (r.drawingType === null || r.drawingType === undefined ? 'UNOBSERVABLE:drawing-type-absent' : `UNKNOWN:${r.drawingType}`)));
  return Object.freeze({
    index: r.index, alt: j || (r.descr === undefined ? null : r.descr), title: j || (r.title === undefined ? null : r.title),
    placement, wrap: j || (r.wrapType === undefined ? null : r.wrapType), widthEmu: j || (r.cx === undefined ? null : r.cx), heightEmu: j || (r.cy === undefined ? null : r.cy),
    paragraph: typeof r.paragraph === 'string' ? r.paragraph : null,
  });
}

function tableFormatOf(f) {
  if (typeof f === 'string') return Object.freeze({ borders: f, shading: f, cellMargins: f, indent: f, width: f, style: f });
  const src = f && typeof f === 'object' ? f : {};
  return Object.freeze({
    borders: Object.freeze(Object.fromEntries(BORDER_EDGES.map(e => [e, borderVal(src.borders && src.borders[e])]))),
    shading: shadingVal(src.shading),
    cellMargins: Object.freeze(Object.fromEntries(MARGIN_EDGES.map(e => [e, marginVal(src.cellMargins && src.cellMargins[e])]))),
    indent: typeof src.indent === 'number' && Number.isFinite(src.indent) ? src.indent : unobservableOr(src.indent, null),
    width: widthVal(src.width),
    style: src.style === undefined ? null : src.style,
  });
}

/** GH #226 3: the driver's per-note chain records, normalized: { paragraphs: string[], complete: boolean, stop: string }; anything malformed = that note NAMED incomplete (never dropped); a non-array = null (unread). */
function noteChainsOf(list) {
  if (!Array.isArray(list)) return null;
  return Object.freeze(list.map(c => (c && Array.isArray(c.paragraphs)
    ? Object.freeze({ paragraphs: Object.freeze(c.paragraphs.map(String)), complete: c.complete === true, stop: typeof c.stop === 'string' ? c.stop : 'unknown' })
    : Object.freeze({ paragraphs: Object.freeze([]), complete: false, stop: 'malformed' }))));
}

/** GH #269b: the driver's comment rows, normalized: a measured value or a NAMED 'UNOBSERVABLE:<reason>' string per field (kept as-is), a missing field = null; a non-array = null (unread, NOT an empty list); a malformed row = a row whose every field is NAMED malformed (never dropped). */
const COMMENT_MALFORMED = 'UNOBSERVABLE:malformed-row';
const unobs = v => typeof v === 'string' && v.indexOf('UNOBSERVABLE:') === 0;
function commentsOf(list) {
  if (!Array.isArray(list)) return null;
  const str = v => (typeof v === 'string' ? v : (v === undefined || v === null ? null : String(v)));
  const reply = r => (r && typeof r === 'object'
    ? Object.freeze({ text: str(r.text), author: str(r.author), userId: str(r.userId) })
    : Object.freeze({ text: COMMENT_MALFORMED, author: COMMENT_MALFORMED, userId: COMMENT_MALFORMED }));
  return Object.freeze(list.map(c => (c && typeof c === 'object'
    ? Object.freeze({
      id: str(c.id), text: str(c.text), author: str(c.author), userId: str(c.userId),
      timeUtc: unobs(c.timeUtc) ? c.timeUtc : (Number.isFinite(c.timeUtc) ? c.timeUtc : null),
      quote: str(c.quote), solved: unobs(c.solved) ? c.solved : (typeof c.solved === 'boolean' ? c.solved : null),
      replies: Array.isArray(c.replies) ? Object.freeze(c.replies.map(reply)) : (unobs(c.replies) ? c.replies : null),
    })
    : Object.freeze({ id: COMMENT_MALFORMED, text: COMMENT_MALFORMED, author: COMMENT_MALFORMED, userId: COMMENT_MALFORMED, timeUtc: COMMENT_MALFORMED, quote: COMMENT_MALFORMED, solved: COMMENT_MALFORMED, replies: COMMENT_MALFORMED }))));
}

// One driver table cell as the canonical inventory carries it. `nestedTables` (the nested tables bound to the HOLDING cell, read_tables returns them) is present only when the
// driver read one: an array of { rows } (recursively the same cell shape) or the driver's named 'UNOBSERVABLE:<reason>' string -- never defaulted, never dropped.
function cellOfDriver(c) {
  const nestedOf = n => (Array.isArray(n) ? Object.freeze(n.map(nt => Object.freeze({ rows: Object.freeze(((nt && nt.rows) || []).map(row => Object.freeze(row.map(cellOfDriver)))) }))) : (typeof n === 'string' ? n : undefined));
  const nested = nestedOf(c.nestedTables);
  return Object.freeze({
    text: norm(c.text), paragraphs: c.paragraphs, simple: c.simple === true, format: Object.freeze({ ...(c.format || {}) }),
    // Wave 4d: per paragraph (text + format); null = unread
    paraFormats: Array.isArray(c.paraFormats) ? Object.freeze(c.paraFormats.map(p => Object.freeze({ text: norm(p.text), format: Object.freeze({ ...(p.format || {}) }) }))) : null,
    // GH #289: the non-plain content types of the cell's paragraphs (a field, an image ...); null = unread
    inline: Array.isArray(c.inline) ? Object.freeze(c.inline.map(String)) : null,
    ...(nested !== undefined ? { nestedTables: nested } : {}),
  });
}

/** a plain JSON value, frozen all the way down */
function deepFreeze(v) {
  if (v && typeof v === 'object') { for (const k of Object.keys(v)) deepFreeze(v[k]); Object.freeze(v); }
  return v;
}

function normalizeInventory(payload) {
  if (!payload || typeof payload.html !== 'string' || payload.html.trim() === '') return null;
  const blocks = attachProperties(decorate(parseBlocks(stripFootnoteProjection(payload.html))), payload.paragraphs);
  // Wave 3: the document-level read (bookmarks, notes, cursor context); null = unread.
  const d = payload.document && typeof payload.document === 'object' ? payload.document : null;
  const document = d ? Object.freeze({
    bookmarks: Array.isArray(d.bookmarks) ? Object.freeze(d.bookmarks.map(b => Object.freeze({ name: String(b.name), text: typeof b.text === 'string' ? b.text : null }))) : null,
    footnotes: Array.isArray(d.footnotes) ? Object.freeze(d.footnotes.map(String)) : null,
    endnotes: Array.isArray(d.endnotes) ? Object.freeze(d.endnotes.map(String)) : null,
    // GH #226 3 (additive, the first-paragraph arrays above are unchanged): per note the FULL paragraph list through the driver's GetNext() chain + whether the chain is trusted (complete) + how it stopped; null = unread.
    footnoteChains: noteChainsOf(d.footnoteChains),
    endnoteChains: noteChainsOf(d.endnoteChains),
    // GH #269b (additive, opt-in on the driver side): the main document's comments; null = unread (NOT an empty list).
    comments: commentsOf(d.comments),
    inMainBody: typeof d.inMainBody === 'boolean' ? d.inMainBody : null,
    // Wave 4a: the driver's public table read (the cell-identity authority); null = unread.
    // the driver's NAMED reason when the table element walk failed (tables is then null); absent otherwise
    ...(typeof d.tablesUnobservable === 'string' ? { tablesUnobservable: d.tablesUnobservable } : {}),
    tables: Array.isArray(d.tables) ? Object.freeze(d.tables.map(t => Object.freeze({
      rows: Object.freeze((t.rows || []).map(row => Object.freeze(row.map(cellOfDriver)))),
      // M5 gate 8: the calculated grid column widths (EO_INTERNAL_DEPENDENCY READ); null = unread
      gridCalc: Array.isArray(t.gridCalc) && t.gridCalc.length > 0 && t.gridCalc.every(v => typeof v === 'number' && Number.isFinite(v)) ? Object.freeze(t.gridCalc.slice()) : null,
      // W04-4: the table-level tblPr fields (table-format family); 'UNOBSERVABLE:...'
      // string per sub-field if the driver's own read failed, never a default.
      format: tableFormatOf(t.format),
    }))) : null,
    // M5 gate 6: the document's default text properties ({ fontFamily, fontSize }); null = unread
    defaultText: d.defaultText && typeof d.defaultText.fontFamily === 'string' && Number.isFinite(d.defaultText.fontSize)
      ? Object.freeze({ fontFamily: d.defaultText.fontFamily, fontSize: d.defaultText.fontSize }) : null,
    // W04-4: the table-style NAME-to-id list (table-format family's `style` field, read
    // BEFORE any dispatch) -- every GetType()==='table' style, its name and its own ToJSON().
    // styleId (a bare number, measured). 'UNOBSERVABLE:unread' when the driver's own list build
    // failed entirely (never a guessed/empty list standing in for a real failure).
    tableStyles: Array.isArray(d.tableStyles) ? Object.freeze(d.tableStyles.map(s => Object.freeze({ name: String(s.name), styleId: s.styleId }))) : (typeof d.tableStyles === 'string' ? d.tableStyles : 'UNOBSERVABLE:unread'),
    // GH #268b: the images the driver's page-side projection returned (docx-images-read.cjs): per image a row of primitives; 'UNOBSERVABLE:<reason>' per field / for the whole list when a read failed.
    // NOT the preservation multiset `images` above (the HTML projection's src / alt / title): that one stays as it is.
    images: Array.isArray(d.images) ? Object.freeze(d.images.map(imageRowOf)) : (typeof d.images === 'string' ? d.images : 'UNOBSERVABLE:unread'),
    // Wave 5: per section, per kind (header/footer), per variant: null = no part of its own;
    // the part's paragraphs (normalized text + format). sections null = unread.
    sections: Array.isArray(d.sections) ? Object.freeze(d.sections.map((s, index) => {
      const part = p => (p && Array.isArray(p.paras) ? Object.freeze({ paragraphs: p.paragraphs, simple: p.simple === true,
        paras: Object.freeze(p.paras.map(x => Object.freeze({ text: norm(x.text), format: Object.freeze({ ...(x.format || {}) }),
          // M5 gate 6: the effective run properties; null = unread
          runs: Array.isArray(x.runs) ? Object.freeze(x.runs.map(r => Object.freeze({ ...r, text: String(r.text || '') }))) : null,
          // GH #289: the non-plain content types of the paragraph's public ToJSON (a field, an image ...); null = unread
          inline: Array.isArray(x.inline) ? Object.freeze(x.inline.map(String)) : null }))),
        // GH #289 H-05: the part's tables (element index + cell texts) and whether it is paragraphs of runs + tables only
        tables: Object.freeze((Array.isArray(p.tables) ? p.tables : []).map(t => Object.freeze({ at: t.at, rows: Object.freeze((t.rows || []).map(r => Object.freeze(r.map(x => norm(x))))) }))),
        withTables: p.withTables === true }) : null);
      const kind = k => Object.freeze(Object.fromEntries(['default', 'even', 'title'].map(v => [v, part(s && s[k] ? s[k][v] : null)])));
      // W11: page-setup fields. A number is passed through verbatim; the driver's own
      // 'UNOBSERVABLE:<reason>' string (a missing/throwing public getter, per-field, never a
      // default) is passed through AS-IS too, so a FAIL_CLOSED reason downstream can name it --
      // never silently coerced to null the way titlePage/evenAndOdd collapse an unread flag.
      const measured = v => (typeof v === 'number' && Number.isFinite(v) ? v : (typeof v === 'string' ? v : 'UNOBSERVABLE:unread'));
      // W11-4: GetColumnsSpaces() is an ARRAY read -- pass the array through verbatim when
      // it is one (frozen, defensive copy), else the driver's own 'UNOBSERVABLE:<reason>' string.
      const measuredArr = v => (Array.isArray(v) ? Object.freeze(v.slice()) : (typeof v === 'string' ? v : 'UNOBSERVABLE:unread'));
      const pageSetup = s && typeof s === 'object' ? {
        pageWidth: measured(s.pageWidth), pageHeight: measured(s.pageHeight),
        margins: Object.freeze({ top: measured(s.margins && s.margins.top), right: measured(s.margins && s.margins.right), bottom: measured(s.margins && s.margins.bottom), left: measured(s.margins && s.margins.left) }),
        columnsCount: measured(s.columnsCount), columnsSpaces: measuredArr(s.columnsSpaces), startPageNumber: measured(s.startPageNumber),
      } : { pageWidth: 'UNOBSERVABLE:unread', pageHeight: 'UNOBSERVABLE:unread', margins: Object.freeze({ top: 'UNOBSERVABLE:unread', right: 'UNOBSERVABLE:unread', bottom: 'UNOBSERVABLE:unread', left: 'UNOBSERVABLE:unread' }), columnsCount: 'UNOBSERVABLE:unread', columnsSpaces: 'UNOBSERVABLE:unread', startPageNumber: 'UNOBSERVABLE:unread' };
      return Object.freeze({ index, header: kind('header'), footer: kind('footer'), titlePage: variantFlag(s && s.titlePage), evenAndOdd: variantFlag(s && s.evenAndOdd), ...pageSetup });
    })) : null,
    // W-10 (style DEFINITIONS): keyed by NAME (never styleId -- the driver's own
    // comment explains why: ToJSON carries no styleId). A missing/throwing getter arrives as the
    // driver's 'UNOBSERVABLE:<reason>' string and is passed through AS-IS (the pageSetup `measured`
    // pattern above, same rule: never silently coerced to a guessed value).
    // GH #293 L7: per requested list-item text, the word /5 Bridge's whole affected set (docx.numbering.listMembers.read/v1), passed
    // through as read -- { unread: <reason> } off /5 or on a failed read; absent = not requested.
    ...(d.numberingMembers && typeof d.numberingMembers === 'object' ? { numberingMembers: deepFreeze(JSON.parse(JSON.stringify(d.numberingMembers))) } : {}),
    styles: d.styles && typeof d.styles === 'object' ? Object.freeze(Object.fromEntries(Object.entries(d.styles).map(([name, st]) => [name, st && typeof st === 'object' ? Object.freeze({
      name: st.name, identity: st.identity ? Object.freeze({ ...st.identity }) : null,
      text: Object.freeze({ ...(st.text || {}) }), paragraph: Object.freeze({ ...(st.paragraph || {}) }),
    }) : null]))) : null,
  }) : null;
  return Object.freeze({ source: 'LIVE', blocks, images: parseImages(payload.html), propertyRead: Array.isArray(payload.paragraphs), document, ...(typeof payload.richUnobservable === 'string' ? { richUnobservable: payload.richUnobservable } : {}), internalGate: internalGateOf(payload.runtime) });
}

/**
 * W2.1: the document's semantic TEXT UNITS in order -- paragraphs, headings and list items -- so a
 * paragraph that a later list step turned into a list item is still the same unit. `kind` is the
 * current representation; `prev` / `next` are the neighbouring units' texts (representation-free
 * context); `block` is the containing block (its locator carries the heading path).
 */
/**
 * GH #271: the tables as a body comparison sees them -- the stored content (rows, format), NOT the calculated grid.
 * `gridCalc` is a layout read through the bridge; on a LARGE document it was measured null for 30/30 tables in one
 * read and set for 30/30 in the next, with no write between them. null stays null (unread).
 */
function tablesContent(tables) {
  return Array.isArray(tables) ? tables.map(t => ({ rows: t.rows, format: t.format })) : null;
}

function textUnits(primary) {
  const units = [];
  for (const b of (primary && primary.blocks) || []) {
    if (b.kind === 'paragraph' || b.kind === 'heading') units.push({ kind: b.kind, text: b.text, node: b, block: b });
    else if (b.kind === 'list') for (const item of b.items || []) units.push({ kind: 'list_item', text: item.text, node: item, block: b });
  }
  return units.map((u, i) => ({ ...u, prev: i > 0 ? norm(units[i - 1].text) : '', next: i + 1 < units.length ? norm(units[i + 1].text) : '' }));
}

/** The ONE text unit with this text; absent or more than one -> not ok (never the first of many). */
function resolveTextUnit(primary, text) {
  const hits = textUnits(primary).filter(u => norm(u.text) === norm(text));
  return hits.length === 1 ? { ok: true, unit: hits[0] } : { ok: false, count: hits.length, reason: hits.length ? 'text-unit-ambiguous' : 'text-unit-not-found' };
}

function countExact(primary, text) {
  const wanted = norm(text);
  return ((primary && primary.blocks) || []).filter(b => blockText(b) === wanted).length;
}

module.exports = { normalizeInventory, stripFootnoteProjection, parseBlocks, parseImages, norm, blockText, countExact, textUnits, tablesContent, resolveTextUnit };
