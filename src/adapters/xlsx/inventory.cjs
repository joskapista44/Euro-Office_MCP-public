'use strict';

/**
 * XLSX live payload -> the adapter's primary observation. Pure.
 *
 * The payload is what the XLSX driver read through PUBLIC getters in one live read:
 *   sheets     [{ name, index, used }]            (used = the used range, A1, or null)
 *   cells      [{ sheet, a, v, f, count, merge, nf?, fill? }] every cell of every used range, plus the requested
 *              windows; v = GetValue, f = GetFormula ('' when none), count = WorksheetFunction.COUNT of
 *              the single-cell range (1 number / 0 not a number / null unread), merge = the cell's
 *              MergeArea address; nf / fill = GetNumberFormat / GetFillColor, only on window cells
 *              read with `format: true`
 *   truncated  the read stopped at the driver's cell ceiling
 *   unsupported [{ field, reason }] getters that were unavailable or threw
 *
 * `complete` is false when anything was truncated or unreadable: a verifier that needs the whole
 * workbook (L3) must then fail closed rather than compare a partial picture.
 */

const { analyze } = require('./dependencies.cjs');
const { withProvenPivots } = require('./pivot-source.cjs');
const { pivotRefsOf } = require('./pivot-refs.cjs');

const key = (sheet, a) => `${sheet}\u0000${a}`;

const deepFreeze = o => { if (o && typeof o === 'object' && !Object.isFrozen(o)) { Object.values(o).forEach(deepFreeze); Object.freeze(o); } return o; };

/** M7 tier H P9: the read's objects, every pivot with its public or content-proven source (pivot-source.cjs). */
const objectsOf = (payload, sheetNames) => withProvenPivots(payload.objects, payload.runtime && payload.runtime.identity, sheetNames);

function deps(sheets, cells, payload, formulas) {
  if (payload.dependents === undefined) return { dependencies: null, rules: null, definedNames: null };
  const d = payload.dependents;
  const internal = typeof d === 'string' ? { names: d, validations: d } : (d || {});
  // M7 S32 slice 2: a tiered read's cells are only the windows -- the analysis also sees every formula of the T3 index
  const all = formulas ? { ...Object.fromEntries(Object.entries(formulas).map(([k, f]) => [k, { f }])), ...cells } : cells;
  const r = analyze(sheets.map(s => s.name), all, { names: internal.names, validations: internal.validations, conditional: payload.conditional, objects: objectsOf(payload, sheets.map(s => s.name)) });
  return deepFreeze({ dependencies: r.dependencies, rules: r.rules, definedNames: r.names });
}

/** A cell-level bold/italic read (EO_INTERNAL_DEPENDENCY, build-gated in the driver): boolean passes
 * through; an already-tagged 'UNOBSERVABLE:<why>' string passes through; anything else is an
 * unexpected shape and is tagged generically -- never defaulted to a boolean. */
const internalFlag = x => (typeof x === 'boolean' ? x : (typeof x === 'string' && x.startsWith('UNOBSERVABLE') ? x : 'UNOBSERVABLE:not-boolean'));

function normalizeCell(c) {
  const f = typeof c.f === 'string' && c.f.startsWith('=') ? c.f : '';
  const count = c.count === 0 || c.count === 1 ? c.count : null;
  const v = c.v === undefined ? null : c.v;
  const out = { v, f, count, merge: typeof c.merge === 'string' ? c.merge.replace(/\$/g, '').toUpperCase() : null };
  // number format / fill: present only on window cells whose read asked for them (a clear's proof)
  if ('nf' in c) out.nf = c.nf === undefined ? null : c.nf;
  if ('fill' in c) out.fill = c.fill === undefined ? null : c.fill;
  // cell-level bold / italic (EO_INTERNAL_DEPENDENCY READ, build-gated in the driver): true | false | 'UNOBSERVABLE:<why>'
  if ('bold' in c) out.bold = internalFlag(c.bold);
  if ('italic' in c) out.italic = internalFlag(c.italic);
  // style (wave 4, format_range): wrap boolean | font / align / border object -- each 'UNOBSERVABLE:<why>' when its
  // reader could not answer (empty cell, mixed runs, unverified build ...); never defaulted
  const unob = x => (typeof x === 'string' && x.startsWith('UNOBSERVABLE') ? x : 'UNOBSERVABLE:unexpected-shape');
  if ('wrap' in c) out.wrap = typeof c.wrap === 'boolean' ? c.wrap : unob(c.wrap);
  // text rotation (X-F07a): PUBLIC GetOrientation(), a number | 'UNOBSERVABLE:<why>' -- never defaulted.
  if ('rot' in c) out.rot = typeof c.rot === 'number' ? c.rot : unob(c.rot);
  // X-F07b/c (owner GO 2026-10-05): readingOrder/indent (numbers) / shrink (boolean) -- the
  // internal Range-model align axes, gated separately from align.h/v/border; never defaulted.
  if ('readingOrder' in c) out.readingOrder = typeof c.readingOrder === 'number' ? c.readingOrder : unob(c.readingOrder);
  if ('indent' in c) out.indent = typeof c.indent === 'number' ? c.indent : unob(c.indent);
  if ('shrink' in c) out.shrink = typeof c.shrink === 'boolean' ? c.shrink : unob(c.shrink);
  for (const k of ['font', 'align', 'border']) if (k in c) out[k] = c[k] && typeof c[k] === 'object' ? deepFreeze(JSON.parse(JSON.stringify(c[k]))) : unob(c[k]);
  return Object.freeze(out);
}

function normalizeInventory(payload) {
  if (!payload || !Array.isArray(payload.sheets)) return null;
  const sheets = payload.sheets.map(s => Object.freeze({ name: String(s.name), index: Number(s.index), used: s.used ? String(s.used).replace(/\$/g, '').toUpperCase() : null }));
  const cells = {};
  for (const c of Array.isArray(payload.cells) ? payload.cells : []) {
    if (!c || typeof c.sheet !== 'string' || typeof c.a !== 'string') continue;
    cells[key(c.sheet, c.a.replace(/\$/g, '').toUpperCase())] = normalizeCell(c);
  }
  const unsupported = Array.isArray(payload.unsupported) ? payload.unsupported.slice() : [];
  const tier = tierFields(sheets, payload, unsupported);
  const formulas = formulaIndexOf(payload, unsupported);
  return Object.freeze({
    format: 'xlsx',
    sheets: Object.freeze(sheets),
    cells: Object.freeze(cells),
    complete: payload.truncated !== true && unsupported.length === 0,
    truncated: payload.truncated === true,
    unsupported: Object.freeze(unsupported),
    // per-sheet objects of the sheets the read named (charts / pivots / sheet-level names); an axis whose
    // reader failed is a reason string, never an empty list
    objects: Object.freeze(Object.fromEntries(Object.entries(objectsOf(payload, sheets.map(s => s.name))).map(([n, o]) => [n, Object.freeze({ ...o })]))),
    // M7 P3: the raw pivot identities of the same sheets (pivot-refs.cjs), for refresh_pivot only
    pivotRefs: deepFreeze(pivotRefsOf(payload.objects)),
    // column / row layout (wave 4), read only when a layout step asked: { sheet: { cols: { B: { width, hidden } },
    // rows: { '2': { height, hidden } } } | 'UNOBSERVABLE:<why>' }; null when not read
    layout: payload.layout === undefined || payload.layout === null ? null : deepFreeze(JSON.parse(JSON.stringify(payload.layout))),
    // freeze panes (wave 4), read only when a freeze step asked: { sheet: { state: 'frozen', row0, col0 } | { state:
    // 'none' } | 'UNOBSERVABLE:<why>' }, and the ACTIVE sheet's name; both null when not read
    panes: payload.panes === undefined || payload.panes === null ? null : deepFreeze(JSON.parse(JSON.stringify(payload.panes))),
    active: typeof payload.active === 'string' ? payload.active : null,
    // the public range-level validation reads (wave 6), keyed '<sheet>!<range>'; null when none was asked
    cfReads: payload.cfReads && typeof payload.cfReads === 'object' ? deepFreeze(JSON.parse(JSON.stringify(payload.cfReads))) : null,
    validationReads: payload.validationReads && typeof payload.validationReads === 'object' ? deepFreeze(JSON.parse(JSON.stringify(payload.validationReads))) : null,
    // the print-setting reads (wave 8, the model), keyed by sheet name; null when none was asked
    printReads: payload.printReads && typeof payload.printReads === 'object' ? deepFreeze(JSON.parse(JSON.stringify(payload.printReads))) : null,
    // the page-layout reads (wave 8), keyed by sheet name; null when none was asked
    pageReads: payload.pageReads && typeof payload.pageReads === 'object' ? deepFreeze(JSON.parse(JSON.stringify(payload.pageReads))) : null,
    // the chart reads (wave 7: charts + pivot names), keyed by sheet name; null when none was asked
    chartReads: payload.chartReads && typeof payload.chartReads === 'object' ? deepFreeze(JSON.parse(JSON.stringify(payload.chartReads))) : null,
    // drawings (M7 S2b): every sheet's image / shape / chart through the public getters, or null when not asked
    drawingReads: payload.drawingReads && typeof payload.drawingReads === 'object' ? deepFreeze(JSON.parse(JSON.stringify(payload.drawingReads))) : null,
    // the AutoFilter reads (wave 6), keyed by sheet name; null when none was asked
    filterReads: payload.filterReads && typeof payload.filterReads === 'object' ? deepFreeze(JSON.parse(JSON.stringify(payload.filterReads))) : null,
    // build = the editor build the driver saw; fileId = the document the session opened (a move's carried expectation is
    // bound to it, owner decision 2026-09-30); null when the driver did not name it
    // M7 S32 slice 1: the observation tier. 'tiered': the cells above are ONLY the task's windows, and `digests` (per
    // sheet: used range, excluded windows, entry count, SHA-256 of the non-blank content outside them) cover the rest
    ...tier,
    // M7 S32 slice 2: the T3 formula index of a tiered read -- '<sheet>\0<A1>' -> formula text, every formula cell
    // OUTSIDE the windows (the windows' own formulas are in `cells`); null when not read. Not cells: no value is known.
    formulas,
    runtime: Object.freeze({ identity: payload.runtime && typeof payload.runtime.identity === 'string' ? payload.runtime.identity : null, cacheKey: payload.runtime && typeof payload.runtime.cacheKey === 'string' ? payload.runtime.cacheKey : null, fileId: payload.runtime && typeof payload.runtime.fileId === 'string' && payload.runtime.fileId !== '' ? payload.runtime.fileId : null }),
    // the dependency readers (names / validations INTERNAL build-gated, conditional formats and objects
    // PUBLIC): read only when the task asked for them (payload.dependents present). Any reader gap or
    // unresolvable reference -> complete false with the reason; never an empty set (owner decision
    // 2026-09-28). delete refuses unless dependencies.complete; copy unless rules.complete.
    ...deps(sheets, cells, payload, formulas),
  });
}

/** The T3 index as a frozen map, or null; a malformed index is a gap (the observation is incomplete). */
function formulaIndexOf(payload, unsupported) {
  if (payload.tier !== 'tiered' || payload.formulaIndex === undefined) return null;
  const rows = payload.formulaIndex;
  const ok = Array.isArray(rows) && rows.every(r => Array.isArray(r) && typeof r[0] === 'string' && typeof r[1] === 'string' && typeof r[2] === 'string' && r[2].startsWith('='));
  if (!ok) { unsupported.push({ field: 'formula-index', reason: 'malformed' }); return null; }
  return Object.freeze(Object.fromEntries(rows.map(([sheet, a, f]) => [key(sheet, a.replace(/\$/g, '').toUpperCase()), f])));
}

/** Every formula the observation knows: the read cells' own, plus the T3 index ('<sheet>\0<A1>' -> text). */
function formulaEntries(primary) {
  const out = Object.entries(primary.cells || {}).filter(([, c]) => c.f).map(([k, c]) => [k, c.f]);
  for (const [k, f] of Object.entries(primary.formulas || {})) if (!(primary.cells && primary.cells[k])) out.push([k, f]);
  return out;
}

/** The tier fields of the inventory; a tiered payload's digest gap is pushed to `unsupported` (incomplete). */
function tierFields(sheets, payload, unsupported) {
  if (payload.tier !== 'tiered') return { tier: 'full', digests: null };
  const gap = digestGap(sheets, payload.digests);
  if (gap) unsupported.push({ field: 'digest', reason: gap });
  return { tier: 'tiered', digests: Array.isArray(payload.digests) ? deepFreeze(JSON.parse(JSON.stringify(payload.digests))) : null };
}

/** A tiered payload must carry exactly one well-formed digest per sheet; otherwise the reason (incomplete). */
function digestGap(sheets, digests) {
  if (!Array.isArray(digests)) return 'missing';
  for (const s of sheets) {
    const rows = digests.filter(d => d && d.sheet === s.name);
    if (rows.length !== 1) return `sheet ${s.name}: ${rows.length} digests`;
    if (typeof rows[0].hash !== 'string' || !/^[0-9a-f]{64}$/.test(rows[0].hash) || !Number.isInteger(rows[0].count) || !Array.isArray(rows[0].excluded)) return `sheet ${s.name}: malformed digest`;
  }
  return digests.length === sheets.length ? null : 'digest for an unknown sheet';
}

const cellAt = (primary, sheet, a) => (primary && primary.cells ? primary.cells[key(sheet, a)] || null : null);

/** A cell that has never been read (outside every used range and window) is not "blank": it is unknown. */
const isBlank = cell => !!cell && cell.f === '' && (cell.v === null || cell.v === '');

module.exports = { normalizeInventory, cellAt, isBlank, key, formulaEntries };
