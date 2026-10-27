'use strict';

const { norm, textUnits } = require('../inventory.cjs');
const { UNRESOLVED } = require('./text.cjs');
const { tableReadWithVerticalMerges, cellSelector, scopeOf } = require('./table.cjs');

/**
 * DOCX `table-format` family (post-M7, W04 -- a NEW capability, not a source-migrated intent;
 * registered via src/adapters/docx/capabilities.cjs's NEW_INTENTS, disjoint from the 28/28
 * migration inventory, same mechanism as `set_page_setup` (W11)). See
 * docs/W04-TABLE-FORMAT-DESIGN-20261007.md (v3) for the full design -- every field shape, key
 * mapping and comparison strategy below is cited there, with the LIVE evidence it came from.
 *
 * Measured live: W04-1a reflection, W04-1b/1c write + ToJSON key shapes for the `top`
 * border edge/margin, W04-2 ToJSON readback for shading/cellMargins.top/
 * indent/width/style, W04-4 for the remaining 5 border edges + `'none'` style,
 * the remaining 3 cellMargins edges, and the style NAME-to-id resolution. All six fields are
 * MEASURED end to end; nothing here is guessed.
 *
 * Shape follows the existing table families (table.cjs, table-row.cjs, ...): the TABLE is located
 * by one of its cells' text, via `cellSelector`/`scopeOf`, quoted verbatim, not invented. Unlike
 * those families, the WRITE targets the TABLE itself (its `tblPr`), not a cell.
 */

const BORDER_EDGES = Object.freeze(['top', 'bottom', 'left', 'right', 'insideH', 'insideV']);
const MARGIN_EDGES = Object.freeze(['top', 'right', 'bottom', 'left']);
const FIELDS = Object.freeze(['borders', 'shading', 'cellMargins', 'indent', 'width', 'style']);
// format_table is a NEW_INTENTS capability, not one of the 28 migrated intents -- it is
// deliberately NOT added to field-contract.cjs's STEP_KEYS, for the SAME reason page-setup.cjs
// states (that table's own acceptance test asserts Object.keys(STEP_KEYS) equals the 28-intent
// registry 1:1; confirmed the hard way during W11-3's own history). The fail-closed
// unknown-field check is re-implemented HERE, locally, same pattern as page-setup.cjs.
const STEP_KEYS = Object.freeze(['intent', 'id', 'dependsOn', 'table', ...FIELDS]);
const COLOR_RE = /^#[0-9A-Fa-f]{6}$/;

const isPosInt = v => Number.isInteger(v) && v > 0;
const isNonNegInt = v => Number.isInteger(v) && v >= 0;
const isUnobservable = v => typeof v === 'string' && v.indexOf('UNOBSERVABLE:') === 0;

/** #RRGGBB -> raw { r, g, b } NUMBERS -- never an Api.CreateRGBColor(...) object (v2 fix: the object shape measured the WRONG result, 000000, for SetShd; borders were only ever measured with raw numbers). */
function parseColor(hex) {
  const n = parseInt(hex.slice(1), 16);
  return { r: (n >> 16) & 255, g: (n >> 8) & 255, b: n & 255 };
}

function validateBorderEdge(edge, obj, index) {
  if (!obj || typeof obj !== 'object') return { ok: false, detail: `step ${index}: format_table borders.${edge} must be an object` };
  const keys = Object.keys(obj).sort().join(',');
  // 'none' is MEASURED but forbids size/color -- the measured write ignores them
  // (reads back sz:0, color.auto:true regardless of what was requested).
  if (obj.style === 'none') {
    if (keys !== 'style') return { ok: false, detail: `step ${index}: format_table borders.${edge} with style 'none' must not also set size or color` };
    return { ok: true, value: { style: 'none' } };
  }
  if (keys !== 'color,size,style') return { ok: false, detail: `step ${index}: format_table borders.${edge} requires all of style, size, color` };
  if (obj.style !== 'single') return { ok: false, detail: `step ${index}: format_table borders.${edge}.style must be 'single' or 'none' (the only measured values)` };
  // size is in EIGHTHS OF A POINT (measured: sz 4 = 0.5pt) -- the raw eighths value, not points.
  if (!isPosInt(obj.size)) return { ok: false, detail: `step ${index}: format_table borders.${edge}.size must be a positive integer (eighths of a point)` };
  if (typeof obj.color !== 'string' || !COLOR_RE.test(obj.color)) return { ok: false, detail: `step ${index}: format_table borders.${edge}.color must be #RRGGBB` };
  const { r, g, b } = parseColor(obj.color);
  return { ok: true, value: { style: obj.style, size: obj.size, r, g, b } };
}

function validateBorders(borders, index) {
  if (!borders || typeof borders !== 'object') return { ok: false, detail: `step ${index}: format_table borders must be an object` };
  const edges = Object.keys(borders);
  if (!edges.length) return { ok: false, detail: `step ${index}: format_table borders must name at least one edge` };
  const out = {};
  for (const e of edges) {
    if (!BORDER_EDGES.includes(e)) return { ok: false, detail: `field-unsupported: step ${index} (format_table).borders.${e}` };
    const v = validateBorderEdge(e, borders[e], index);
    if (!v.ok) return v;
    out[e] = v.value;
  }
  return { ok: true, value: out };
}

function validateShading(color, index) {
  if (typeof color !== 'string' || !COLOR_RE.test(color)) return { ok: false, detail: `step ${index}: format_table shading must be #RRGGBB` };
  return { ok: true, value: parseColor(color) };
}

function validateCellMargins(cm, index) {
  if (!cm || typeof cm !== 'object' || Object.keys(cm).sort().join(',') !== MARGIN_EDGES.slice().sort().join(',')) {
    return { ok: false, detail: `step ${index}: format_table cellMargins requires all four of ${MARGIN_EDGES.join(', ')}` };
  }
  const out = {};
  for (const e of MARGIN_EDGES) {
    if (!isNonNegInt(cm[e])) return { ok: false, detail: `step ${index}: format_table cellMargins.${e} must be a non-negative integer (twips)` };
    out[e] = cm[e];
  }
  return { ok: true, value: out };
}

function validateIndent(v, index) {
  return isNonNegInt(v) ? { ok: true, value: v } : { ok: false, detail: `step ${index}: format_table indent must be a non-negative integer (twips)` };
}
function validateWidth(v, index) {
  // the unit argument ('twips') is a FIXED internal constant (v2) -- only the numeric value is exposed.
  return isPosInt(v) ? { ok: true, value: v } : { ok: false, detail: `step ${index}: format_table width must be a positive integer (twips)` };
}
function validateStyle(v, index) {
  return (typeof v === 'string' && norm(v)) ? { ok: true, value: v } : { ok: false, detail: `step ${index}: format_table style must be a non-empty style name` };
}

function plan(step, index) {
  if (step.intent !== 'format_table') return { ok: false, detail: `step ${index}: ${step.intent} is not a format-table intent` };
  for (const key of Object.keys(step)) {
    if (!STEP_KEYS.includes(key)) return { ok: false, detail: `field-unsupported: step ${index} (format_table).${key}` };
  }
  // table.cjs's own cellSelector, quoted verbatim (design doc point 1) -- `table` names ONE of
  // the target table's cells; format_table never targets that cell, only the table it is in.
  const sel = cellSelector(step.table);
  if (!sel || !norm(sel.text)) return { ok: false, detail: `step ${index}: format_table requires table: { text } or { tableCell: {...} }` };
  const present = FIELDS.filter(f => step[f] !== undefined);
  if (present.length === 0) return { ok: false, detail: `step ${index}: format_table requires at least one of ${FIELDS.join(', ')}` };
  const fields = {};
  if (step.borders !== undefined) { const v = validateBorders(step.borders, index); if (!v.ok) return v; fields.borders = v.value; }
  if (step.shading !== undefined) { const v = validateShading(step.shading, index); if (!v.ok) return v; fields.shading = v.value; }
  if (step.cellMargins !== undefined) { const v = validateCellMargins(step.cellMargins, index); if (!v.ok) return v; fields.cellMargins = v.value; }
  if (step.indent !== undefined) { const v = validateIndent(step.indent, index); if (!v.ok) return v; fields.indent = v.value; }
  if (step.width !== undefined) { const v = validateWidth(step.width, index); if (!v.ok) return v; fields.width = v.value; }
  if (step.style !== undefined) { const v = validateStyle(step.style, index); if (!v.ok) return v; fields.style = v.value; }
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`, intent: step.intent, selector: Object.freeze({ table: step.table }),
      anchorText: norm(sel.text), tableAnchor: sel.tableAnchor === null ? null : norm(sel.tableAnchor),
      fields: Object.freeze(fields), preservationMode: 'PRIMITIVE_INTEGRAL',
      ...(Array.isArray(step.dependsOn) ? { dependsOn: step.dependsOn } : {}),
    },
  };
}

// ---- the observation ----------------------------------------------------------------------------

/** Which style NAME resolves to which id -- the ONLY way classify()/resolve() can compare an opaque tblStyle id against a requested NAME, without dispatching. */
function resolveStyleId(list, name) {
  if (!Array.isArray(list)) return { ok: false, reason: 'table-style-list-unobservable' };
  const hits = list.filter(s => s.name === name);
  if (hits.length !== 1) return { ok: false, reason: hits.length ? 'table-style-ambiguous' : 'table-style-not-found' };
  return { ok: true, styleId: hits[0].styleId };
}

function borderMatches(format, want) {
  const b = format.borders;
  if (isUnobservable(b)) return false;
  for (const edge of Object.keys(want)) {
    const req = want[edge];
    const obs = b && b[edge];
    if (isUnobservable(obs)) return false;
    if (req.style === 'none') {
      if (!obs || obs.value !== 'none') return false;
    } else if (!obs || obs.value !== req.style || obs.sz !== req.size || !obs.color || obs.color.auto !== false || obs.color.r !== req.r || obs.color.g !== req.g || obs.color.b !== req.b) return false;
  }
  return true;
}
function shadingMatches(format, want) {
  const s = format.shading;
  if (isUnobservable(s)) return false;
  // val == 'clear' is the FIXED sType constant (design doc point 1/3); 'color' (distinct from
  // 'fill') is read but its exact relationship to the request is NEM-MERT -- not compared.
  return !!s && s.val === 'clear' && !!s.fill && s.fill.r === want.r && s.fill.g === want.g && s.fill.b === want.b;
}
function cellMarginsMatches(format, want) {
  const m = format.cellMargins;
  if (isUnobservable(m)) return false;
  for (const edge of Object.keys(want)) {
    const obs = m && m[edge];
    if (isUnobservable(obs)) return false;
    if (!obs || obs.type !== 'dxa' || obs.w !== want[edge]) return false;
  }
  return true;
}
/** String() on both sides (design doc point 3): the style OBJECT's own styleId is measured as a bare number; the TABLE's own in-session tblPr.tblStyle representation was not captured un-truncated -- this comparison is written so that ambiguity cannot produce a false mismatch either way. */
function fieldMatches(format, field, want, expectedStyleId) {
  if (field === 'borders') return borderMatches(format, want);
  if (field === 'shading') return shadingMatches(format, want);
  if (field === 'cellMargins') return cellMarginsMatches(format, want);
  if (field === 'indent') return !isUnobservable(format.indent) && format.indent === want;
  if (field === 'width') return !isUnobservable(format.width) && !!format.width && format.width.type === 'dxa' && format.width.w === want;
  if (field === 'style') return !isUnobservable(format.style) && String(format.style) === String(expectedStyleId);
  return false;
}

/** Which of the REQUESTED fields currently read as UNOBSERVABLE (a missing/throwing read, never a guessed value) -- checked only on the sub-keys actually requested, never the whole axis. */
function unobservableFields(format, fields) {
  const out = [];
  if (fields.borders) {
    const b = format.borders;
    if (isUnobservable(b) || Object.keys(fields.borders).some(e => isUnobservable(b && b[e]))) out.push('borders');
  }
  if (fields.shading !== undefined && isUnobservable(format.shading)) out.push('shading');
  if (fields.cellMargins) {
    const m = format.cellMargins;
    if (isUnobservable(m) || Object.keys(fields.cellMargins).some(e => isUnobservable(m && m[e]))) out.push('cellMargins');
  }
  if (fields.indent !== undefined && isUnobservable(format.indent)) out.push('indent');
  if (fields.width !== undefined && isUnobservable(format.width)) out.push('width');
  if (fields.style !== undefined && isUnobservable(format.style)) out.push('style');
  return out;
}

/**
 * Resolves WHICH table `op` names (via table.cjs's own tableRead/scopeOf, reused verbatim --
 * the SAME cell-identity mechanism every table family uses), then reads that table's own
 * `format`. GH #270 (g270v1 G4 / G8): the vertical-merge refusal is PER TABLE, as it has been for replace_table_cell and the column family since W06-3 -- a merge in ANOTHER table of the document no
 * longer stops this one (tableReadWithVerticalMerges reconciles every table's two reads separately); only a merge in the TARGET table fails closed here, with the same reason as before (a merged grid
 * with a table-level format is not measured, so it is not guessed).
 */
function tableFormatState(primary, op) {
  const read = tableReadWithVerticalMerges(primary);
  if (!read.ok) return read;
  const scope = scopeOf(read, op.tableAnchor);
  if (!scope.ok) return scope;
  const hits = scope.cells.filter(c => c.cell.text === op.anchorText);
  if (hits.length !== 1) return { ok: false, reason: hits.length ? 'table-anchor-ambiguous' : 'table-anchor-not-found', tables: hits.length };
  const tableIndex = parseInt(hits[0].at.split(':')[0], 10);
  // GH #291 T-01: a vertical merge in the target table no longer refuses -- LIVE-measured (docs/evidence/gh291-vmerge-table-probe-20261021.json): borders incl.
  // insideH/insideV, shading, cellMargins, indent, width and a style written on such a table; the saved vMerge/gridSpan and every cell unchanged. L3 below now also compares the
  // target table's own grid (cell texts, paragraph counts, HTML spans, vertical merges), so a write that breaks the merge is never proven.
  const grid = { cells: read.cells.filter(c => c.at.split(':')[0] === String(tableIndex)).map(c => [c.at, c.cell.text, c.cell.paragraphs]), spans: read.html[tableIndex].rows.map(r => r.map(c => [c.span, c.rowspan])), vspans: (read.vspans && read.vspans[tableIndex]) || [] };
  const d = primary && primary.document;
  const tables = Array.isArray(d && d.tables) ? d.tables : null;
  if (!tables || !tables[tableIndex]) return { ok: false, reason: 'table-unreadable' };
  const format = tables[tableIndex].format;
  // Defensive only: inventory.cjs's own tableFormatOf() normalizes a top-level UNOBSERVABLE
  // string into a PER-FIELD breakdown (every sub-field set to that same string) before this
  // function ever sees it -- a bare string here would mean some OTHER pipeline stage skipped
  // that normalization, not the ordinary unobservable-read path (which unobservableFields()
  // below catches, per requested field, same as page-setup's own discipline).
  if (typeof format === 'string') return { ok: false, reason: `FAIL_CLOSED_UNSUPPORTED_SHAPE:table-format-unreadable:${format}` };
  if (!format || typeof format !== 'object') return { ok: false, reason: 'table-format-unreadable' };
  let expectedStyleId = null;
  if (op.fields.style !== undefined) {
    const s = resolveStyleId(d && d.tableStyles, op.fields.style);
    if (!s.ok) return { ok: false, reason: s.reason };
    expectedStyleId = s.styleId;
  }
  const fieldNames = Object.keys(op.fields);
  const matches = fieldNames.map(f => fieldMatches(format, f, op.fields[f], expectedStyleId));
  const done = matches.length > 0 && matches.every(Boolean);
  const unobservable = unobservableFields(format, op.fields);
  return { ok: true, tableIndex, format, done, fields: fieldNames, unobservable, expectedStyleId, grid };
}

// ---- classify / resolve -------------------------------------------------------------------------

function classify(primary, op) {
  const s = tableFormatState(primary, op);
  if (!s.ok) return { state: 'FAIL_CLOSED', evidence: { reason: s.reason } };
  if (s.done) {
    if (op.recovery === UNRESOLVED) return { state: 'FAIL_CLOSED', evidence: { reason: 'preservation-baseline-unavailable' } };
    return { state: 'SATISFIED', evidence: { table: s.tableIndex, fields: op.fields, proof: 'fresh-live-table-matches-every-requested-field' } };
  }
  if (s.unobservable.length) return { state: 'FAIL_CLOSED', evidence: { reason: `FAIL_CLOSED_UNSUPPORTED_SHAPE:table-format-field-unobservable:${s.unobservable.join(',')}` } };
  return { state: 'APPLY', evidence: { table: s.tableIndex, fields: op.fields } };
}

function resolve(primary, selector, ctx = {}) {
  const op = ctx.operation;
  if (!op || op.intent !== 'format_table') return { ok: false, target: null, selector, ambiguity: 'INVALID_SELECTOR' };
  const s = tableFormatState(primary, op);
  const up = r => r.toUpperCase().replace(/-/g, '_');
  if (!s.ok) return { ok: false, target: null, selector, ambiguity: up(s.reason) };
  if (!s.done && s.unobservable.length) return { ok: false, target: null, selector, ambiguity: `TABLE_FORMAT_FIELD_UNOBSERVABLE:${s.unobservable.join(',').toUpperCase()}` };
  return { ok: true, target: { kind: 'table_format', tableIndex: s.tableIndex, anchorText: op.anchorText, tableAnchor: op.tableAnchor, fields: op.fields, expectedStyleId: s.expectedStyleId }, selector, ambiguity: null };
}

function preDispatchCheck(op, target) {
  const same = target && target.kind === 'table_format' && target.anchorText === op.anchorText && target.tableAnchor === op.tableAnchor;
  return same ? { ok: true } : { ok: false, dispatched: false, detail: 'table-format-target-unresolved' };
}

// ---- lower --------------------------------------------------------------------------------------

/** Call-ready field shapes: colors already split into raw r/g/b numbers (v2 fix, never an Api.CreateRGBColor object); 'none' edges carry only {style:'none'} (size/color forbidden at plan time). */
function loweredFields(fields, expectedStyleId) {
  const out = {};
  if (fields.borders) {
    out.borders = {};
    for (const [edge, v] of Object.entries(fields.borders)) out.borders[edge] = v.style === 'none' ? { style: 'none' } : { style: v.style, size: v.size, r: v.r, g: v.g, b: v.b };
  }
  if (fields.shading) out.shading = { r: fields.shading.r, g: fields.shading.g, b: fields.shading.b };
  if (fields.cellMargins) out.cellMargins = { ...fields.cellMargins };
  if (fields.indent !== undefined) out.indent = fields.indent;
  if (fields.width !== undefined) out.width = fields.width;
  if (fields.style !== undefined) out.style = { name: fields.style, expectedStyleId };
  return out;
}

function lower(op, target) {
  const meta = Object.freeze({ family: 'tableFormat', action: 'tableFormat', tableAnchor: target.tableAnchor, anchorText: target.anchorText, fields: loweredFields(target.fields, target.expectedStyleId) });
  return { kind: 'callCommand', body: commandBody(meta), meta };
}

/**
 * One callCommand, public API only, re-resolving the table by cell text from scratch (table.cjs's
 * own re-check shape) -- never trusting a stale target. The style axis re-resolves BY NAME at
 * write time too (same re-check philosophy as every other family's commandBody).
 *
 * W04-6 (the reviewer's review): EVERY setter/resolution the requested fields need is checked
 * to exist BEFORE any write happens. A guard miss used to `continue`/skip just that ONE field
 * while still writing the others -- a PARTIAL WRITE: the caller's step asked for N fields, and
 * fewer than N were actually applied, silently. Any single missing setter or unresolved style
 * now aborts the WHOLE command, writing NOTHING -- the same atomic-or-nothing guarantee every
 * other family's commandBody already gives (table.cjs's own "any guard miss returns unchanged").
 */
function commandBody(meta) {
  return `
try {
  const m = ${JSON.stringify(meta)};
  const textOpts = { Numbering: false };
  const norm = s => String(s || '').replace(/\\s+/g, ' ').trim();
  const doc = Api.GetDocument();
  const cellsOf = t => { const out = []; for (let r = 0; r < t.GetRowsCount(); r += 1) { const row = t.GetRow(r); for (let c = 0; c < row.GetCellsCount(); c += 1) out.push(row.GetCell(c)); } return out; };
  const textOf = cell => norm(cell.GetContent().GetText(textOpts));
  let tables = doc.GetAllTables() || [];
  if (m.tableAnchor !== null) {
    tables = tables.filter(t => cellsOf(t).some(cell => textOf(cell) === m.tableAnchor));
    if (tables.length !== 1) return;
  }
  const hits = tables.filter(t => cellsOf(t).some(cell => textOf(cell) === m.anchorText));
  if (hits.length !== 1) return;
  const table = hits[0];
  const BORDER_SETTER = { top: 'SetTableBorderTop', bottom: 'SetTableBorderBottom', left: 'SetTableBorderLeft', right: 'SetTableBorderRight', insideH: 'SetTableBorderInsideH', insideV: 'SetTableBorderInsideV' };
  const MARGIN_SETTER = { top: 'SetTableCellMarginTop', right: 'SetTableCellMarginRight', bottom: 'SetTableCellMarginBottom', left: 'SetTableCellMarginLeft' };
  // ---- validate EVERY needed setter/resolution first -- nothing is written until ALL pass ----
  if (m.fields.borders) {
    for (const edge of Object.keys(m.fields.borders)) if (typeof table[BORDER_SETTER[edge]] !== 'function') return;
  }
  if (m.fields.shading && typeof table.SetShd !== 'function') return;
  if (m.fields.cellMargins) {
    for (const edge of Object.keys(m.fields.cellMargins)) if (typeof table[MARGIN_SETTER[edge]] !== 'function') return;
  }
  if (m.fields.indent !== undefined && typeof table.SetTableInd !== 'function') return;
  if (m.fields.width !== undefined && typeof table.SetWidth !== 'function') return;
  let matchedStyle = null;
  if (m.fields.style) {
    if (typeof doc.GetAllStyles !== 'function' || typeof table.SetStyle !== 'function') return;
    const styles = doc.GetAllStyles() || [];
    matchedStyle = styles.find(function (s) {
      try { return typeof s.GetType === 'function' && s.GetType() === 'table' && typeof s.GetName === 'function' && String(s.GetName()) === m.fields.style.name; } catch (e) { return false; }
    });
    if (!matchedStyle) return;
  }
  // ---- every precondition held -- write all requested fields ----
  if (m.fields.borders) {
    for (const edge of Object.keys(m.fields.borders)) {
      const req = m.fields.borders[edge];
      const setter = BORDER_SETTER[edge];
      if (req.style === 'none') table[setter]('none', 0, 0, 0, 0, 0);
      else table[setter]('single', req.size, 0, req.r, req.g, req.b);
    }
  }
  if (m.fields.shading) table.SetShd('clear', m.fields.shading.r, m.fields.shading.g, m.fields.shading.b);
  if (m.fields.cellMargins) {
    for (const edge of Object.keys(m.fields.cellMargins)) table[MARGIN_SETTER[edge]](m.fields.cellMargins[edge]);
  }
  if (m.fields.indent !== undefined) table.SetTableInd(m.fields.indent);
  if (m.fields.width !== undefined) table.SetWidth('twips', m.fields.width);
  if (m.fields.style) table.SetStyle(matchedStyle);
} catch (_) {}
`;
}

// ---- verify -------------------------------------------------------------------------------------

/** The main body as the reads see it, EXCLUDING tables (this family writes a table by design -- the per-table comparison below covers that). Same shape as page-setup.cjs's bodyKey, minus `tables`. */
function bodyKey(primary) {
  const d = (primary && primary.document) || {};
  return JSON.stringify({ units: textUnits(primary).map(u => [u.kind, norm(u.text)]), footnotes: d.footnotes || null, endnotes: d.endnotes || null, bookmarks: d.bookmarks || null, images: ((primary && primary.images) || []).length, inMainBody: d.inMainBody === undefined ? null : d.inMainBody });
}

function otherTablesKey(tables, exceptIndex) {
  return JSON.stringify(tables.map((t, i) => (i === exceptIndex ? null : { rows: t.rows, format: t.format })));
}

/** L1: every requested field now matches, via the SAME extraction as classify; the resolved table unchanged. */
function verifyPrimitive(op, before, after) {
  const b = tableFormatState(before && before.primary, op);
  if (!b.ok) return { ok: false, evidence: { reason: `before-${b.reason}` } };
  const a = tableFormatState(after && after.primary, op);
  if (!a.ok) return { ok: false, evidence: { reason: `after-${a.reason}` } };
  const ok = !b.done && a.done;
  return { ok, evidence: { table: a.tableIndex, fields: op.fields, formatReadAfter: a.format } };
}

/**
 * L3, PRIMITIVE_INTEGRAL (same owner placement as every other table family -- table-row.cjs:
 * "every other table unchanged; the text units outside the tables unchanged in order"). Preserved:
 * the body (bodyKey, tables excluded); every OTHER table's rows+format; every field of the SAME
 * table not requested by this step (same cross-axis-preservation requirement as page-setup.cjs's
 * verifyPreservation, reused by analogy -- a shading-only step must leave borders/cellMargins/
 * indent/width/style exactly as read before the step).
 */
function verifyPreservation(op, before, after) {
  const b = tableFormatState(before && before.primary, op);
  const a = tableFormatState(after && after.primary, op);
  if (!b.ok || !a.ok) return { ok: false, evidence: { reason: !b.ok ? `before-${b.reason}` : `after-${a.reason}` } };
  const requested = new Set(Object.keys(op.fields));
  const axisKeys = { borders: ['borders'], shading: ['shading'], cellMargins: ['cellMargins'], indent: ['indent'], width: ['width'], style: ['style'] };
  const changedOwnFields = Object.keys(axisKeys).filter(f => !requested.has(f)).filter(f => {
    const keyOf = fmt => JSON.stringify(axisKeys[f].map(k => fmt && fmt[k]));
    return keyOf(b.format) !== keyOf(a.format);
  });
  const bd = before.primary && before.primary.document;
  const ad = after.primary && after.primary.document;
  const bTables = Array.isArray(bd && bd.tables) ? bd.tables : [];
  const aTables = Array.isArray(ad && ad.tables) ? ad.tables : [];
  const sameOtherTables = bTables.length === aTables.length && otherTablesKey(bTables, b.tableIndex) === otherTablesKey(aTables, a.tableIndex);
  const sameBody = bodyKey(before.primary) === bodyKey(after.primary);
  const sameGrid = JSON.stringify(b.grid) === JSON.stringify(a.grid);
  const ok = changedOwnFields.length === 0 && sameOtherTables && sameBody && sameGrid;
  return { ok, evidence: { changedOwnFields, otherTablesUnchanged: sameOtherTables, bodyUnchanged: sameBody, targetGridUnchanged: sameGrid, tableCount: aTables.length } };
}

/** L2: the final table holds exactly the requested fields. */
function wholeTaskChecks(op, finalObservation) {
  const s = tableFormatState(finalObservation.primary, op);
  return [{ name: `${op.id}: table format matches ${Object.keys(op.fields).join('+')}`, ok: s.ok && s.done }];
}

// refs: none -- a table-format value is not a body text unit a later step can name by reference.
function producedText() {
  return null;
}

module.exports = {
  FIELDS, BORDER_EDGES, MARGIN_EDGES,
  plan, classify, resolve, preDispatchCheck, lower, commandBody,
  verifyPrimitive, verifyPreservation, wholeTaskChecks, producedText,
  tableFormatState, fieldMatches, resolveStyleId, bodyKey, otherTablesKey,
};
