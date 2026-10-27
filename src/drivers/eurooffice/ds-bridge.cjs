'use strict';

/**
 * Euro-Office DS Bridge v1 -- the MCP-side client (READ pilot; owner GO the owner, owner channel, 2026-10-06, internal tracker;
 * contract docs/DS-BRIDGE-V1-CONTRACT-20261006.md).
 *
 * The bridge lives INSIDE the patched Document Server (`Api.GetEuroOfficeBridge()`), so the client runs in the editor,
 * as part of a callCommand body. `BRIDGE_CLIENT_SRC` is that client: fail-closed gates, in order --
 * BRIDGE_UNAVAILABLE | MANIFEST_MALFORMED | BRIDGE_VERSION_MISMATCH | RUNTIME_VERSION_MISMATCH | PATCH_VERSION_MISMATCH |
 * CAPABILITY_MISSING | CAPABILITY_VERSION_MISMATCH | UNEXPECTED_SHAPE -- plus the bridge's own `{ ok: false, reason }`.
 * It never falls back to a direct internal call, never names an internal method, and only ever asks for a capability
 * ID the manifest advertises.
 *
 *   __eoBridgeRead(capabilityId, capabilityVersion, expectedRuntime, args, valueType)
 *     -> { ok: true, value } | { ok: false, reason }
 * expectedRuntime is runtime-identity.cjs RUNTIMES.<format> ({ documentServerVersion, patchVersion }) -- the served
 * manifest's stable identity (owner decision (B), internal tracker). The manifest's `documentServerBuild` (the DS
 * cache key at patch-build time) and the editor URL's cache key are NOT compared: a DS restart regenerates the key on
 * unchanged content.
 */
const { BRIDGE_MAJOR } = require('../../core/runtime-identity.cjs');

const BRIDGE_CAPABILITIES = Object.freeze({
  BOLD_READ: Object.freeze({ id: 'textStyle.bold.read/v1', version: 1, valueType: 'boolean' }),
  ITALIC_READ: Object.freeze({ id: 'textStyle.italic.read/v1', version: 1, valueType: 'boolean' }),
  APPEARANCE_READ: Object.freeze({ id: 'range.appearance.read/v1', version: 1, valueType: 'object' }),
  FILTER_FIELDS_READ: Object.freeze({ id: 'sheet.autoFilter.fieldIdentity.read/v1', version: 1, valueType: 'object' }),
  COL_BREAKS_READ: Object.freeze({ id: 'sheet.columnBreaks.read/v1', version: 1, valueType: 'object' }),
  FREEZE_READ: Object.freeze({ id: 'sheet.freezeState.read/v1', version: 1, valueType: 'object' }),
  DRAWING_ANCHOR_READ: Object.freeze({ id: 'drawing.anchor.read/v1', version: 1, valueType: 'object' }),
  DRAWING_TEXT_READ: Object.freeze({ id: 'drawing.text.read/v1', version: 1, valueType: 'string' }),
  COL_WIDTH_CANON_READ: Object.freeze({ id: 'sheet.columnWidth.canonicalize/v1', version: 1, valueType: 'object' }),
  CHART_LABELS_EFFECTIVE_READ: Object.freeze({ id: 'chart.dataLabels.effective.read/v1', version: 1, valueType: 'object' }),
  // Word bridge (W-B1920-C1 cutover, card an internal branch; contract
  // documentserver-patches/9.3.4.60/euro-office-bridge-word-v1/bridge-block.js, SZO SZERINT atvett ID/kontraktus):
  DOCX_LIST_LEVEL_FORMAT_READ: Object.freeze({ id: 'docx.numbering.levelFormat.read/v1', version: 1, valueType: 'object' }),
  DOCX_TABLE_GRID_READ: Object.freeze({ id: 'docx.table.effectiveGrid.read/v1', version: 1, valueType: 'object' }),
  // M11.5e.4 (M11.5 continuous mandate, internal tracker): B21 (section variant state read, word /3 only).
  DOCX_SECTION_VARIANT_STATE_READ: Object.freeze({ id: 'docx.section.variantState.read/v1', version: 1, valueType: 'object' }),
  // GH #293 L7 (word /5 only): one list item -> the WHOLE affected set of its numbering document-wide (table cells included), the
  // other lists on the same definition with their overrides and order; the set_numbering / set_list_type co-member guards read it.
  DOCX_LIST_MEMBERS_READ: Object.freeze({ id: 'docx.numbering.listMembers.read/v1', version: 1, valueType: 'object' }),
  // B11 /3 COMBINED (internal tracker): both READ capabilities of the /2->/3 operation; the xlsx driver's
  // dependency inventory (dependentsBridgeBody) reads through them after the B11 cutover.
  DEFINED_NAMES_READ: Object.freeze({ id: 'workbook.definedNames.read/v1', version: 1, valueType: 'object' }),
  VALIDATIONS_READ: Object.freeze({ id: 'range.validations.read/v1', version: 1, valueType: 'object' }),
  // Wave B READ /4 (owner order owner channel 2026-10-08 point 9; design docs/DS-BRIDGE-WAVE-B-READ-DESIGN-20261008.md):
  // the B05 READ, B09 and B13 cutovers read through these; a /3 runtime refuses them (PATCH_VERSION_MISMATCH).
  ALIGN_EXT_READ: Object.freeze({ id: 'range.alignmentExt.read/v1', version: 1, valueType: 'object' }),
  PRINT_SETTINGS_READ: Object.freeze({ id: 'sheet.printSettings.read/v1', version: 1, valueType: 'object' }),
  PIVOT_STATE_READ: Object.freeze({ id: 'pivot.state.read/v1', version: 1, valueType: 'object' }),
  // X-O03/6b (internal tracker; X-O03/4 LIVE r4 FOUND_READ_PATH, X-O03/5 design doc, X-O03/6a LIVE
  // r6): a scatter chart series' Y-value array -- the PUBLIC chartApi's GetValues() is a formula
  // string or THROWS on scatter (X-O03/1, internal tracker), this reads the internal model
  // instead, through the SAME entry point as chart.dataLabels.effective.read/v1 (X-O03/6a
  // confirmed `===`). Served only on patchVersion /9.
  SCATTER_Y_READ: Object.freeze({ id: 'chart.series.scatterY.read/v1', version: 1, valueType: 'object' }),
  // X-O04/10 (internal tracker; design docs/X-O04-9-CHART-STYLE-READ-BRIDGE-DESIGN-20261014.md,
  // reviewed and accepted): both chains are X-O04/4..X-O04/8 LIVE VERIFIED. { fill, outline } per series -- each a 6-hex-digit
  // RGB string or null (a PROVEN valid model state, not a guess). Served only on patchVersion /11.
  MARKER_STYLE_READ: Object.freeze({ id: 'chart.series.markerStyle.read/v1', version: 1, valueType: 'object' }),
  // { verticalGridlines, horizontalGridlines } per chart -- each a 6-hex-digit RGB string or null.
  // The cross-binding X-O04/8 found (SetMajorVerticalGridlines's colour lands on
  // getHorizontalAxis(), and vice versa) is hidden entirely inside the DS-side handler; this client
  // only ever sees the PUBLIC setter's own semantics. Served only on patchVersion /11.
  GRIDLINES_COLOR_READ: Object.freeze({ id: 'chart.axis.majorGridlinesColor.read/v1', version: 1, valueType: 'object' }),
  // X-O05/2 (internal tracker; X-O05/1 LIVE WORKS_PERSISTED_READABLE, reviewed and accepted): a
  // chart series' own c:order index -- the PUBLIC chartApi has no order getter; the internal
  // model's OWN getOrder() method does (series.private_GetSeries() on the public wrapper; the
  // Bridge's internal chart-type resolution already lands one level further in, so the handler
  // calls it directly on the resolved internal series object, same convention as
  // MARKER_STYLE_READ's `s.marker` one line above -- no `.private_GetSeries()` needed there
  // either). { order } -- a non-negative integer, never guessed; getOrder() itself, not a data
  // field, so no risk of reading an unsynchronized cache. Served only on patchVersion /12.
  SERIES_ORDER_READ: Object.freeze({ id: 'chart.series.order.read/v1', version: 1, valueType: 'object' }),
  // X-O05/2 (the OTHER X-O05/1 axis; LIVE re-measured and reviewed and accepted): a BAR chart's
  // own gapWidth/overlap -- there is no PUBLIC getter and no internal METHOD either, unlike every
  // other migrated axis in this file; the LIVE re-measurement read the internal DATA field
  // DIRECTLY and confirmed it survives both the same session (right after the setter) and a fresh
  // reopen with no further write -- not an unsynchronized write-time-only mirror. { gapWidth,
  // overlap }, both numbers, read ONLY on a bar-type chart (the LIVE measurement was bar-only;
  // every other chart type is unmeasured for this field and refuses by name, never a guessed
  // number or null). Served only on patchVersion /12.
  GAP_OVERLAP_READ: Object.freeze({ id: 'chart.bar.gapOverlap.read/v1', version: 1, valueType: 'object' }),
  // X-O08 (zero LIVE, 2026-10-15): a shape's fill/line colour -- the PUBLIC GetFill()/GetLine().GetFill()
  // wrappers only ever expose the TYPE ("solid"/"nofill", X-O08-7 WRITE_ONLY finding, confirmed
  // again by two independent LIVE rounds), never the RGB value. The handler reads the TYPE through
  // the public getter first (never guessed) and, ONLY when it is "solid", the actual byte value
  // through the internal model (the public wrapper's own `.Drawing` field -- the SAME access
  // `drawing.anchor.read/v1`/`drawing.text.read/v1` already use -- then `.spPr[.ln].Fill.fill.
  // color.color.RGBA`, X-O08 LIVE VERIFIED session+reopen, both the fill chain and the line chain,
  // on three shapes with three distinct colours + a no-stroke edge + a noFill edge). { fill, line}
  // -- each either a 6-hex-digit RGB string, the named reason `noFill`/`noLine`, a named
  // `notMeasured:<type>` for any non-solid fill type (gradient/pattern/picture -- never guessed,
  // X-O08 never measured those), or an UNOBSERVABLE chain-break reason if the type says solid but
  // the internal chain still breaks (a genuine anomaly, not swallowed).
  // Served on patchVersion /14 (the real rev14-candidate/bridge-block.js built on top of
  // /13's actual bytes, XLSX_V14 added to runtime-identity.cjs ACCEPTED_RUNTIMES.xlsx and to the
  // B03/B05 Bridge-write runtime lists -- the B03/B05 handlers are untouched, /14 adds only this
  // ONE READ capability). On any runtime older than /14 this client-side wiring degrades
  // gracefully to UNOBSERVABLE, by design (same shared `bridgeAxis` convention every other migrated
  // axis in this file uses).
  FILL_LINE_COLOR_READ: Object.freeze({ id: 'drawing.fillLineColor.read/v1', version: 1, valueType: 'object' }),
  // GH #139 B /16 (design docs/GH139-B16-CHART-STYLE-READ-DESIGN.md;
  // LIVE MEASURED: write+session+reopen round trip, S1==S2 on every
  // chart): the measured READ counterpart of the WRITE-ONLY `markerSymbol`/`lineStyle` style
  // targets (/15, MARKER_SYMBOL_WRITE_ID/LINE_STYLE_WRITE_ID, charts.cjs) -- the SAME
  // `series.marker` object `MARKER_STYLE_READ` already reads fill/outline off (symbol/size are
  // different fields on it), and the series' own `spPr.ln.prstDash` / `smooth` fields. { symbol,
  // size } / { dash, smooth }, each either the model's own measured value, null (never written --
  // the model's own unset state, never a guessed default), or a NAMED UNOBSERVABLE string for an
  // unmapped token/code or a field the series' own type never carries (e.g. `smooth` on a bar
  // series). Served only on patchVersion /16.
  MARKER_SYMBOL_READ: Object.freeze({ id: 'chart.series.markerSymbol.read/v1', version: 1, valueType: 'object' }),
  LINE_STYLE_READ: Object.freeze({ id: 'chart.series.lineStyle.read/v1', version: 1, valueType: 'object' }),
  // GH #326 / #327 (owner decision, path A): a sheet's protection state, the
  // 15 named Excel flags, and (with a range) the non-empty unlocked / hidden-formula cells -- the engine's own
  // sheetProtection object carries credential fields (algorithmName/hashValue/saltValue/spinCount/password/
  // temporaryPassword), NEVER read or enumerated by the handler; only the 15 named booleans + `protected` pass
  // through (docs/DS-9.3.5-BRIDGE17-PROTECTION-READ.md). { protected, forbids, cells? } -- `cells` only when a
  // range is asked, out of this unit's scope (client-wiring only; the #326 guard / #327 redaction units consume
  // it later). Served only on patchVersion /17; a /16 runtime correctly refuses with CAPABILITY_MISSING (same
  // shared `bridgeAxis` convention every other migrated axis in this file uses) -- compatibility-first, the
  // ACCEPTED list carries both /16 and /17 so a DS /16 -> /17 apply causes no product outage.
  PROTECTION_READ: Object.freeze({ id: 'sheet.protection.read/v1', version: 1, valueType: 'object' }),
});

/**
 * B15 /5 (owner GO the owner, owner channel 2026-10-09 msg 33257, internal tracker): NOT a READ. A RUNTIME_HYDRATION capability
 * makes a non-active sheet's ephemeral editor view available and returns no document data; it is served ONLY through the
 * bridge's own hydrate() entry (read() refuses it), so the READ contract above is untouched. Allowed: an ephemeral,
 * session-only view object. Forbidden: document content/structure, history/undo, persistent metadata, dirty/save state.
 */
const BRIDGE_HYDRATIONS = Object.freeze({
  SHEET_MODEL_READY: Object.freeze({ id: 'sheet.model.ensureReady/v1', version: 1 }),
});

const BRIDGE_CLIENT_SRC = `
  const __eoBridgeRead = (capabilityId, capabilityVersion, expectedRuntime, args, valueType) => {
    let bridge;
    try { bridge = typeof Api.GetEuroOfficeBridge === 'function' ? Api.GetEuroOfficeBridge() : null; } catch (err) { bridge = null; }
    if (!bridge || typeof bridge.manifest !== 'function' || typeof bridge.read !== 'function') return { ok: false, reason: 'BRIDGE_UNAVAILABLE' };
    let m;
    try { m = bridge.manifest(); } catch (err) { return { ok: false, reason: 'BRIDGE_UNAVAILABLE' }; }
    if (!m || typeof m !== 'object' || typeof m.bridgeVersion !== 'string' || typeof m.documentServerVersion !== 'string' || typeof m.patchVersion !== 'string' || !Array.isArray(m.capabilities)) return { ok: false, reason: 'MANIFEST_MALFORMED' };
    if (Number(m.bridgeVersion.split('.')[0]) !== ${BRIDGE_MAJOR}) return { ok: false, reason: 'BRIDGE_VERSION_MISMATCH' };
    const accepted = (Array.isArray(expectedRuntime) ? expectedRuntime : [expectedRuntime]).filter(e => e && typeof e === 'object');
    if (!accepted.some(e => m.documentServerVersion === e.documentServerVersion)) return { ok: false, reason: 'RUNTIME_VERSION_MISMATCH' };
    if (!accepted.some(e => m.documentServerVersion === e.documentServerVersion && m.patchVersion === e.patchVersion)) return { ok: false, reason: 'PATCH_VERSION_MISMATCH' };
    const named = m.capabilities.filter(c => c && c.id === capabilityId);
    if (named.length === 0) return { ok: false, reason: 'CAPABILITY_MISSING' };
    const versioned = named.filter(c => c.version === capabilityVersion);
    if (versioned.length === 0) return { ok: false, reason: 'CAPABILITY_VERSION_MISMATCH' };
    const cap = versioned.filter(c => c.kind === 'READ' && c.mutation === 'NO');
    if (cap.length !== 1) return { ok: false, reason: 'CAPABILITY_MISSING' };
    let r;
    try { r = bridge.read(capabilityId, args); } catch (err) { return { ok: false, reason: 'BRIDGE_THREW' }; }
    if (!r || typeof r !== 'object' || typeof r.ok !== 'boolean') return { ok: false, reason: 'UNEXPECTED_SHAPE' };
    const keys = Object.keys(r).sort().join(',');
    if (r.ok === true) return keys === 'ok,value' && typeof r.value === valueType ? { ok: true, value: r.value } : { ok: false, reason: 'UNEXPECTED_SHAPE' };
    return keys === 'ok,reason' && typeof r.reason === 'string' && r.reason ? { ok: false, reason: 'BRIDGE:' + r.reason } : { ok: false, reason: 'UNEXPECTED_SHAPE' };
  };
`;

/**
 * The hydration client, included ONLY in a body that hydrates (never in the READ client above, so a read-only query
 * stays byte-identical). Same fail-closed gates in the same order as __eoBridgeRead, then: the advertised entry must be
 * kind RUNTIME_HYDRATION + mutation EPHEMERAL_VIEW_ONLY, the bridge must expose hydrate(), and success is exactly
 * { ok: true, value: { modelReady: true } }. No fallback to a direct internal call exists here or in the driver.
 *   __eoBridgeHydrate(capabilityId, capabilityVersion, expectedRuntime, args) -> { ok: true } | { ok: false, reason }
 */
const BRIDGE_HYDRATE_CLIENT_SRC = `
  const __eoBridgeHydrate = (capabilityId, capabilityVersion, expectedRuntime, args) => {
    let bridge;
    try { bridge = typeof Api.GetEuroOfficeBridge === 'function' ? Api.GetEuroOfficeBridge() : null; } catch (err) { bridge = null; }
    if (!bridge || typeof bridge.manifest !== 'function') return { ok: false, reason: 'BRIDGE_UNAVAILABLE' };
    let m;
    try { m = bridge.manifest(); } catch (err) { return { ok: false, reason: 'BRIDGE_UNAVAILABLE' }; }
    if (!m || typeof m !== 'object' || typeof m.bridgeVersion !== 'string' || typeof m.documentServerVersion !== 'string' || typeof m.patchVersion !== 'string' || !Array.isArray(m.capabilities)) return { ok: false, reason: 'MANIFEST_MALFORMED' };
    if (Number(m.bridgeVersion.split('.')[0]) !== ${BRIDGE_MAJOR}) return { ok: false, reason: 'BRIDGE_VERSION_MISMATCH' };
    const accepted = (Array.isArray(expectedRuntime) ? expectedRuntime : [expectedRuntime]).filter(e => e && typeof e === 'object');
    if (!accepted.some(e => m.documentServerVersion === e.documentServerVersion)) return { ok: false, reason: 'RUNTIME_VERSION_MISMATCH' };
    if (!accepted.some(e => m.documentServerVersion === e.documentServerVersion && m.patchVersion === e.patchVersion)) return { ok: false, reason: 'PATCH_VERSION_MISMATCH' };
    const named = m.capabilities.filter(c => c && c.id === capabilityId);
    if (named.length === 0) return { ok: false, reason: 'CAPABILITY_MISSING' };
    const versioned = named.filter(c => c.version === capabilityVersion);
    if (versioned.length === 0) return { ok: false, reason: 'CAPABILITY_VERSION_MISMATCH' };
    const cap = versioned.filter(c => c.kind === 'RUNTIME_HYDRATION' && c.mutation === 'EPHEMERAL_VIEW_ONLY');
    if (cap.length !== 1 || typeof bridge.hydrate !== 'function') return { ok: false, reason: 'CAPABILITY_MISSING' };
    let r;
    try { r = bridge.hydrate(capabilityId, args); } catch (err) { return { ok: false, reason: 'BRIDGE_THREW' }; }
    if (!r || typeof r !== 'object' || typeof r.ok !== 'boolean') return { ok: false, reason: 'UNEXPECTED_SHAPE' };
    const keys = Object.keys(r).sort().join(',');
    if (r.ok === true) return keys === 'ok,value' && r.value && typeof r.value === 'object' && Object.keys(r.value).join(',') === 'modelReady' && r.value.modelReady === true ? { ok: true } : { ok: false, reason: 'UNEXPECTED_SHAPE' };
    return keys === 'ok,reason' && typeof r.reason === 'string' && r.reason ? { ok: false, reason: 'BRIDGE:' + r.reason } : { ok: false, reason: 'UNEXPECTED_SHAPE' };
  };
`;

/**
 * M11.5a.4 (owner GO the owner, owner channel, internal tracker): the FIRST Bridge WRITE the product calls -- B05-WRITE
 * (align.indent / align.shrink). Served only on /6 through the bridge's own write() entry; read() / hydrate() refuse it.
 */
const BRIDGE_WRITES = Object.freeze({
  ALIGN_EXT_WRITE: Object.freeze({ id: 'range.alignmentExt.write/v1', version: 1, kind: 'WRITE', mutation: 'TARGETED' }),
  // M11.5b.4: B10-WRITE (the eight print/page settings, one per call), served only on /7
  PRINT_SETTINGS_WRITE: Object.freeze({ id: 'sheet.printSettings.write/v1', version: 1, kind: 'WRITE', mutation: 'TARGETED' }),
  // M11.5c.4: B03-WRITE (one worksheet copy, appended last; its own receipt shape, parsed by the driver's B03 client), only /8
  WORKSHEET_COPY_WRITE: Object.freeze({ id: 'worksheet.copy/v1', version: 1, kind: 'WRITE', mutation: 'TARGETED' }),
  // M11.5d.4: B18-WRITE (one paragraph's numbering removed; its own receipt shape, parsed by the B18 client), only word /2
  PARAGRAPH_NUMBERING_REMOVE_WRITE: Object.freeze({ id: 'docx.paragraph.numbering.remove/v1', version: 1, kind: 'WRITE', mutation: 'TARGETED' }),
  // GH #292 S4: one style's next / linked style (both sides of a link), by style NAME; its own receipt shape, parsed by the S4 client; only word /4
  STYLE_NEXT_LINK_WRITE: Object.freeze({ id: 'docx.style.nextLink.write/v1', version: 1, kind: 'WRITE', mutation: 'TARGETED' }),
  // GH #293 L7: one list moved to its own copy of a shared definition (every member, label-preserving isolation only); its own receipt
  // shape, parsed by the L7 client; only word /5
  LIST_DETACH_WRITE: Object.freeze({ id: 'docx.numbering.listDetach.write/v1', version: 1, kind: 'WRITE', mutation: 'TARGETED' }),
  // X-O05/3 (internal tracker; reviewed and accepted): a BAR chart's own gapWidth/overlap, written
  // through the internal setGapWidth()/setOverlap() METHODS -- a direct assignment to the data
  // field chart.bar.gapOverlap.read/v1 reads APPEARS to apply in-session but does NOT survive a
  // reopen (PRE-DEPLOY LIVE, internal tracker); the setter is the SAME path X-O05/1's own
  // measurement found persisting. No public setter exists. Served only on /12.
  GAP_OVERLAP_WRITE: Object.freeze({ id: 'chart.bar.gapOverlap.write/v1', version: 1, kind: 'WRITE', mutation: 'TARGETED' }),
  // X-O05/3: a chart series' own c:order index, written through the SAME internal series object
  // chart.series.order.read/v1 reads (its own getOrder()) -- the internal setOrder() method,
  // never a data field (unlike gapOverlap above: an internal METHOD exists for this one, X-O05/1
  // LIVE VERIFIED). Served only on /12.
  SERIES_ORDER_WRITE: Object.freeze({ id: 'chart.series.order.write/v1', version: 1, kind: 'WRITE', mutation: 'TARGETED' }),
  // X-O07/2 (internal tracker; X-O07/1 LIVE MEASURED and reviewed): deletes the WHOLE chart
  // drawing through the internal deleteDrawingBase() METHOD (api.Chart, no-arg) -- the public
  // chart API has no delete method of its own (the X-O07 reflection probe found only
  // series-level RemoveSeria). The handler's own name-lookup is STRICTER than every other
  // chart.* handler: zero matches refuses chart-not-found, more than one refuses
  // chart-name-ambiguous (never guesses on an irreversible operation). Served only on /13.
  CHART_DELETE_WRITE: Object.freeze({ id: 'chart.drawing.delete.write/v1', version: 1, kind: 'WRITE', mutation: 'TARGETED' }),
  // X-O05/5-C (internal tracker; X-O05/5-B round 7b+8d LIVE MEASURED and reviewed): a chart
  // series' own marker symbol/size, written through the internal `marker.setSymbol(token)` /
  // `marker.setNoneMarker()` / `marker.setSize(n)` methods -- the SAME `series.marker` sub-object
  // MARKER_STYLE_READ above already reads fill/outline off; symbol/size are DIFFERENT fields on
  // that same object, never readable back through MARKER_STYLE_READ (no getter for either exists,
  // X-O04/10's own exhaustive enumeration). Receipt: { applied: [...] }. Served only on /15.
  MARKER_WRITE: Object.freeze({ id: 'chart.series.marker.write/v1', version: 1, kind: 'WRITE', mutation: 'TARGETED' }),
  // X-O05/5-C (internal tracker; X-O05/5-B round 8k/8m/8o/8q LIVE MEASURED and reviewed): a chart
  // series' own line dash/smooth. smooth: series.setSmooth(bool) unconditionally. dash: a
  // BRAND-NEW, standalone AscFormat.CLn(), seeded from the EXISTING line's own w/Fill fields when
  // present (never guessed), attached via the series' EXISTING spPr's own setLn when one exists
  // (R8o fix: never replaces an existing spPr, preserving the series' own fill/effects), or a
  // brand-new CSpPr when none exists yet. Receipt on a successful dash write: { applied: ['dash'],
  // widthSource, fillSource } -- the LAST two ('preserved'|'fallback') are an R8q addition
  // (internal tracker, the round 8p LIVE probe's own review) naming WHICH of the two seeding
  // paths fired; neither is ever readable back (no getter for the line's own width/fill exists
  // on the public or `markerStyle`-read internal surface). Served only on /15.
  LINE_STYLE_WRITE: Object.freeze({ id: 'chart.series.lineStyle.write/v1', version: 1, kind: 'WRITE', mutation: 'TARGETED' }),
});

/**
 * The write client, included ONLY in a body that writes B05 (a read-only query stays byte-identical). Two steps:
 *   __eoBridgeWriteGate(capabilityId, capabilityVersion, expectedRuntime) -> { ok: true, bridge } | { ok: false, reason }
 *     the same fail-closed gates in the same order as __eoBridgeRead, then the advertised entry must be kind WRITE +
 *     mutation TARGETED and the bridge must expose write(). Called ONCE, before any mutation of the command.
 *   __eoBridgeWrite(bridge, capabilityId, args, requested) -> { ok: true } | { ok: false, reason, uncertain?, partial?, applied? }
 *     ONE write() call. Success is exactly { ok: true, value: { applied } } with `applied` equal to `requested` (the
 *     requested keys in bridge order); a refusal is { ok: false, reason } or { ok: false, reason, partial, applied }.
 *     Anything else -> UNEXPECTED_SHAPE (uncertain). The receipt never verifies anything: the caller re-reads.
 * No fallback to a direct internal call exists here or in the driver.
 */
const BRIDGE_WRITE_CLIENT_SRC = `
  const __eoBridgeWriteGate = (capabilityId, capabilityVersion, expectedRuntime) => {
    let bridge;
    try { bridge = typeof Api.GetEuroOfficeBridge === 'function' ? Api.GetEuroOfficeBridge() : null; } catch (err) { bridge = null; }
    if (!bridge || typeof bridge.manifest !== 'function') return { ok: false, reason: 'BRIDGE_UNAVAILABLE' };
    let m;
    try { m = bridge.manifest(); } catch (err) { return { ok: false, reason: 'BRIDGE_UNAVAILABLE' }; }
    if (!m || typeof m !== 'object' || typeof m.bridgeVersion !== 'string' || typeof m.documentServerVersion !== 'string' || typeof m.patchVersion !== 'string' || !Array.isArray(m.capabilities)) return { ok: false, reason: 'MANIFEST_MALFORMED' };
    if (Number(m.bridgeVersion.split('.')[0]) !== ${BRIDGE_MAJOR}) return { ok: false, reason: 'BRIDGE_VERSION_MISMATCH' };
    const accepted = (Array.isArray(expectedRuntime) ? expectedRuntime : [expectedRuntime]).filter(e => e && typeof e === 'object');
    if (!accepted.some(e => m.documentServerVersion === e.documentServerVersion)) return { ok: false, reason: 'RUNTIME_VERSION_MISMATCH' };
    if (!accepted.some(e => m.documentServerVersion === e.documentServerVersion && m.patchVersion === e.patchVersion)) return { ok: false, reason: 'PATCH_VERSION_MISMATCH' };
    const named = m.capabilities.filter(c => c && c.id === capabilityId);
    if (named.length === 0) return { ok: false, reason: 'CAPABILITY_MISSING' };
    const versioned = named.filter(c => c.version === capabilityVersion);
    if (versioned.length === 0) return { ok: false, reason: 'CAPABILITY_VERSION_MISMATCH' };
    const cap = versioned.filter(c => c.kind === 'WRITE' && c.mutation === 'TARGETED');
    if (cap.length !== 1) return { ok: false, reason: 'CAPABILITY_KIND_MISMATCH' };
    if (typeof bridge.write !== 'function') return { ok: false, reason: 'CAPABILITY_MISSING' };
    return { ok: true, bridge };
  };
  const __eoBridgeWrite = (bridge, capabilityId, args, requested) => {
    let r;
    try { r = bridge.write(capabilityId, args); } catch (err) { return { ok: false, reason: 'BRIDGE_THREW', uncertain: true }; }
    if (!r || typeof r !== 'object' || typeof r.ok !== 'boolean') return { ok: false, reason: 'UNEXPECTED_SHAPE', uncertain: true };
    const keys = Object.keys(r).sort().join(',');
    if (r.ok === true) {
      const v = r.value;
      const shaped = keys === 'ok,value' && v && typeof v === 'object' && Object.keys(v).join(',') === 'applied' && Array.isArray(v.applied) && v.applied.join(',') === requested.join(',');
      return shaped ? { ok: true } : { ok: false, reason: 'UNEXPECTED_SHAPE', uncertain: true };
    }
    if (typeof r.reason !== 'string' || !r.reason) return { ok: false, reason: 'UNEXPECTED_SHAPE', uncertain: true };
    if (keys === 'ok,reason') return { ok: false, reason: 'BRIDGE:' + r.reason };
    if (keys === 'applied,ok,partial,reason' && typeof r.partial === 'boolean' && Array.isArray(r.applied)) return { ok: false, reason: 'BRIDGE:' + r.reason, partial: r.partial, applied: r.applied.slice() };
    return { ok: false, reason: 'UNEXPECTED_SHAPE', uncertain: true };
  };
`;

/**
 * X-O05/5-C (internal tracker; X-O05/5-B round 8q LIVE MEASURED and reviewed): the
 * marker.write/v1 / lineStyle.write/v1 write client, included ONLY in a set_chart_style body
 * that writes `markerSymbol`/`lineStyle` (after BRIDGE_WRITE_CLIENT_SRC, whose gate it uses).
 *   __eoBridgeSeriesStyleWrite(bridge, capabilityId, args, requested) -> { ok: true, widthSource?, fillSource? }
 *     | { ok: false, reason, uncertain?, partial?, applied? }
 *     ONE write() call. A DELIBERATE widening of __eoBridgeWrite's own strict
 *     `Object.keys(v).join(',') === 'applied'` shape check above: lineStyle.write/v1's own dash
 *     receipt carries TWO additional, OPTIONAL string fields on success (`widthSource`,
 *     `fillSource`, R8q -- naming which of the two width/fill seeding paths fired, the round 8p
 *     LIVE probe's own review) that marker.write/v1's plain `{applied}` receipt never has; both
 *     are forwarded on success (undefined when absent), never themselves verified --
 *     the caller re-reads, same as every other Bridge write here. __eoBridgeWrite itself is
 *     UNCHANGED (every other capability through it keeps the exact-`applied`-only check); this is
 *     the ONE place that tolerates the two named extra keys, and only when present AND a string.
 */
const SERIES_STYLE_WRITE_RECEIPT_EXTRA_KEYS = Object.freeze(['widthSource', 'fillSource']);
const BRIDGE_SERIES_STYLE_WRITE_CLIENT_SRC = `
  const __eoBridgeSeriesStyleWrite = (bridge, capabilityId, args, requested) => {
    let r;
    try { r = bridge.write(capabilityId, args); } catch (err) { return { ok: false, reason: 'BRIDGE_THREW', uncertain: true }; }
    if (!r || typeof r !== 'object' || typeof r.ok !== 'boolean') return { ok: false, reason: 'UNEXPECTED_SHAPE', uncertain: true };
    const keys = Object.keys(r).sort().join(',');
    if (r.ok === true) {
      const v = r.value;
      const extra = ${JSON.stringify(SERIES_STYLE_WRITE_RECEIPT_EXTRA_KEYS)};
      const vkeys = v && typeof v === 'object' ? Object.keys(v) : [];
      const onlyKnown = vkeys.every(k => k === 'applied' || extra.indexOf(k) !== -1);
      const extraTypeOk = extra.every(k => !Object.prototype.hasOwnProperty.call(v || {}, k) || typeof v[k] === 'string');
      const shaped = keys === 'ok,value' && v && typeof v === 'object' && Array.isArray(v.applied) && v.applied.join(',') === requested.join(',') && onlyKnown && extraTypeOk;
      if (!shaped) return { ok: false, reason: 'UNEXPECTED_SHAPE', uncertain: true };
      const out = { ok: true };
      extra.forEach(k => { if (typeof v[k] === 'string') out[k] = v[k]; });
      return out;
    }
    if (typeof r.reason !== 'string' || !r.reason) return { ok: false, reason: 'UNEXPECTED_SHAPE', uncertain: true };
    if (keys === 'ok,reason') return { ok: false, reason: 'BRIDGE:' + r.reason };
    if (keys === 'applied,ok,partial,reason' && typeof r.partial === 'boolean' && Array.isArray(r.applied)) return { ok: false, reason: 'BRIDGE:' + r.reason, partial: r.partial, applied: r.applied.slice() };
    return { ok: false, reason: 'UNEXPECTED_SHAPE', uncertain: true };
  };
`;

/**
 * M11.5c.4: the B03 copy write client, included ONLY in a copy_sheet body (after BRIDGE_WRITE_CLIENT_SRC, whose gate it uses).
 *   __eoBridgeCopyWrite(bridge, capabilityId, args) -> { ok: true, index } | { ok: false, reason, uncertain, mutation, before?, after?, errors? }
 *     ONE write() call. worksheet.copy/v1 has its own receipt: success is exactly { ok: true, value: { applied: ['copy'],
 *     name, index } } with name === args.name. A refusal the capability makes BEFORE its copy call (an explicit allowlist
 *     of the /8 handler's own pre-call reasons) -> mutation NONE; ANY other reason (copy-threw / copy-error /
 *     copy-not-observed / copy-unexpected, or an unknown one) -> uncertain, mutation UNKNOWN: the copy may exist (maintainer
 *     #ref, #ref: fail-closed). A malformed receipt -> UNEXPECTED_SHAPE (uncertain). Never retried; the caller re-reads.
 */
const BRIDGE_COPY_WRITE_CLIENT_SRC = `
  const __eoBridgeCopyWrite = (bridge, capabilityId, args) => {
    const odd = { ok: false, reason: 'UNEXPECTED_SHAPE', uncertain: true, mutation: 'UNKNOWN' };
    let r;
    try { r = bridge.write(capabilityId, args); } catch (err) { return { ok: false, reason: 'BRIDGE_THREW', uncertain: true, mutation: 'UNKNOWN' }; }
    if (!r || typeof r !== 'object' || typeof r.ok !== 'boolean') return odd;
    if (r.ok === true) {
      const v = r.value;
      const shaped = Object.keys(r).sort().join(',') === 'ok,value' && v && typeof v === 'object' && Object.keys(v).sort().join(',') === 'applied,index,name'
        && Array.isArray(v.applied) && v.applied.join(',') === 'copy' && v.name === args.name && typeof v.index === 'number';
      return shaped ? { ok: true, index: v.index } : odd;
    }
    if (typeof r.reason !== 'string' || !r.reason) return odd;
    // ALLOWLIST of the /8 handler's own refusals made BEFORE its copy call (: fail-closed); EVERY other reason
    // -- the known after-call ones and any a later revision adds -- is uncertain, mutation UNKNOWN
    const PRE = ['bad-args', 'unsupported-property', 'invalid-name', 'editor-missing', 'edit-not-allowed', 'workbook-locked', 'workbook-protected', 'name-taken', 'source-not-found', 'source-ambiguous'];
    const pre = PRE.indexOf(r.reason) >= 0 || /^(asc_copyWorksheet|asc_getWorksheetsCount|asc_getWorksheetName|asc_isProtectedWorkbook|asc_registerCallback|asc_unregisterCallback|canEdit)-missing$/.test(r.reason);
    return { ok: false, reason: 'BRIDGE:' + r.reason, uncertain: !pre, mutation: pre ? 'NONE' : 'UNKNOWN', before: r.before || null, after: r.after || null, errors: r.errors || null };
  };
`;

/**
 * X-O07/3: the chart-delete write client, included ONLY in a delete_chart body (after
 * BRIDGE_WRITE_CLIENT_SRC, whose gate it uses).
 *   __eoBridgeChartDeleteWrite(bridge, capabilityId, args) -> { ok: true } | { ok: false, reason, uncertain, mutation }
 *     ONE write() call. chart.drawing.delete.write/v1 has its own receipt (X-O07/2 LIVE
 *     MEASURED): success is exactly { ok: true, value: { deleted: true } } -- NOT the generic
 *     { applied: [requested keys] } shape __eoBridgeWrite expects (there is no "requested
 *     property" for a whole-chart delete), hence this dedicated client, the SAME reason
 *     __eoBridgeCopyWrite above exists for worksheet.copy/v1's own receipt shape. A refusal the
 *     handler makes BEFORE its deleteDrawingBase() call (an explicit allowlist of the handler's
 *     own pre-call reasons: bad-args, unsupported-property, sheet-not-found, charts-threw,
 *     chart-not-found, chart-name-ambiguous, deletedrawingbase-missing) -> mutation NONE; the one
 *     reason the call ITSELF can throw (deletedrawingbase-threw) -> uncertain, mutation UNKNOWN
 *     (fail-closed: the deletion may have partially happened). A malformed receipt ->
 *     UNEXPECTED_SHAPE (uncertain). Never retried; the caller re-reads (classify()'s own fresh
 *     chartReads loop, the same path rename_chart/move_chart already rely on).
 */
const BRIDGE_CHART_DELETE_WRITE_CLIENT_SRC = `
  const __eoBridgeChartDeleteWrite = (bridge, capabilityId, args) => {
    const odd = { ok: false, reason: 'UNEXPECTED_SHAPE', uncertain: true, mutation: 'UNKNOWN' };
    let r;
    try { r = bridge.write(capabilityId, args); } catch (err) { return { ok: false, reason: 'BRIDGE_THREW', uncertain: true, mutation: 'UNKNOWN' }; }
    if (!r || typeof r !== 'object' || typeof r.ok !== 'boolean') return odd;
    if (r.ok === true) {
      const v = r.value;
      const shaped = Object.keys(r).sort().join(',') === 'ok,value' && v && typeof v === 'object' && Object.keys(v).sort().join(',') === 'deleted' && v.deleted === true;
      return shaped ? { ok: true } : odd;
    }
    if (typeof r.reason !== 'string' || !r.reason) return odd;
    // ALLOWLIST of the /13 handler's own refusals made BEFORE its deleteDrawingBase() call
    // (X-O07/2's own reasons list, fail-closed); the one AFTER-call reason
    // (deletedrawingbase-threw) and any a later revision adds are uncertain, mutation UNKNOWN.
    const PRE = ['bad-args', 'unsupported-property', 'sheet-not-found', 'charts-threw', 'chart-not-found', 'chart-name-ambiguous', 'deletedrawingbase-missing'];
    const pre = PRE.indexOf(r.reason) >= 0;
    return { ok: false, reason: 'BRIDGE:' + r.reason, uncertain: !pre, mutation: pre ? 'NONE' : 'UNKNOWN' };
  };
`;

/**
 * M11.5d.4: the B18 numbering-removal write client, included ONLY in a list body that removes numbering (after
 * BRIDGE_WRITE_CLIENT_SRC, whose gate it uses).
 *   __eoBridgeNumberingRemove(bridge, capabilityId, args) -> { ok: true } | { ok: false, reason, uncertain, mutation }
 *     ONE write() call. Success is exactly { ok: true, value: { applied: ['numbering-removed'], before: { level, numId } } }
 *     with before.level === args.level. A refusal the word /2 handler makes BEFORE its RemoveNumPr call (an explicit
 *     allowlist) -> mutation NONE; ANY other reason (remove-threw / remove-not-observed, or an unknown one) -> uncertain,
 *     mutation UNKNOWN. A malformed receipt -> UNEXPECTED_SHAPE (uncertain). Never retried; the caller re-reads (a
 *     removed numbering on a retry is SATISFIED from the fresh read, never a second write).
 */
const BRIDGE_NUMBERING_REMOVE_WRITE_CLIENT_SRC = `
  const __eoBridgeNumberingRemove = (bridge, capabilityId, args) => {
    const odd = { ok: false, reason: 'UNEXPECTED_SHAPE', uncertain: true, mutation: 'UNKNOWN' };
    let r;
    try { r = bridge.write(capabilityId, args); } catch (err) { return { ok: false, reason: 'BRIDGE_THREW', uncertain: true, mutation: 'UNKNOWN' }; }
    if (!r || typeof r !== 'object' || typeof r.ok !== 'boolean') return odd;
    if (r.ok === true) {
      const v = r.value;
      const b = v && v.before;
      const shaped = Object.keys(r).sort().join(',') === 'ok,value' && v && typeof v === 'object' && Object.keys(v).sort().join(',') === 'applied,before'
        && Array.isArray(v.applied) && v.applied.join(',') === 'numbering-removed' && b && typeof b === 'object' && b.level === args.level;
      return shaped ? { ok: true } : odd;
    }
    if (typeof r.reason !== 'string' || !r.reason) return odd;
    // the handler's own refusals BEFORE its removal call; '<routine>-missing' is its routine-presence check (also pre-call)
    const PRE = ['bad-args', 'unsupported-property', 'invalid-level', 'document-unavailable', 'search-threw', 'target-not-found', 'target-ambiguous', 'paragraph-unreadable', 'no-GetNumbering', 'not-numbered', 'level-mismatch'];
    const pre = PRE.indexOf(r.reason) >= 0 || /^[A-Za-z]+-missing$/.test(r.reason);
    return { ok: false, reason: 'BRIDGE:' + r.reason, uncertain: !pre, mutation: pre ? 'NONE' : 'UNKNOWN' };
  };
`;

/**
 * GH #292 S4: the next / linked style write client, included ONLY in a modify_style body that sets one (after BRIDGE_WRITE_CLIENT_SRC,
 * whose gate it uses).
 *   __eoBridgeStyleNextLink(bridge, capabilityId, args) -> { ok: true } | { ok: false, reason, uncertain, mutation }
 *     ONE write() call. Success is exactly { ok: true, value: { applied: ['next-set' | 'link-set'], before: { value } } } with the
 *     applied entry matching args.field. A refusal the word /4 handler makes BEFORE its setter call (an explicit allowlist, incl.
 *     no-change) -> mutation NONE; set-threw / write-not-observed / an unknown reason -> uncertain, mutation UNKNOWN. A malformed
 *     receipt -> UNEXPECTED_SHAPE (uncertain). Never retried; the caller re-reads.
 */
const BRIDGE_STYLE_NEXT_LINK_WRITE_CLIENT_SRC = `
  const __eoBridgeStyleNextLink = (bridge, capabilityId, args) => {
    const odd = { ok: false, reason: 'UNEXPECTED_SHAPE', uncertain: true, mutation: 'UNKNOWN' };
    let r;
    try { r = bridge.write(capabilityId, args); } catch (err) { return { ok: false, reason: 'BRIDGE_THREW', uncertain: true, mutation: 'UNKNOWN' }; }
    if (!r || typeof r !== 'object' || typeof r.ok !== 'boolean') return odd;
    if (r.ok === true) {
      const v = r.value;
      const shaped = Object.keys(r).sort().join(',') === 'ok,value' && v && typeof v === 'object' && Object.keys(v).sort().join(',') === 'applied,before'
        && Array.isArray(v.applied) && v.applied.join(',') === (args.field === 'next' ? 'next-set' : 'link-set') && v.before && typeof v.before === 'object';
      return shaped ? { ok: true } : odd;
    }
    if (typeof r.reason !== 'string' || !r.reason) return odd;
    const PRE = ['bad-args', 'unsupported-property', 'unsupported-field', 'document-unavailable', 'style-not-found', 'target-not-found', 'type-threw', 'type-mismatch', 'internal-missing', 'getter-threw', 'no-change', 'already-linked-elsewhere'];
    const pre = PRE.indexOf(r.reason) >= 0;
    return { ok: false, reason: 'BRIDGE:' + r.reason, uncertain: !pre, mutation: pre ? 'NONE' : 'UNKNOWN' };
  };
`;

/**
 * GH #293 L7: the list detach write client, included ONLY in a set_numbering body with scope 'list' on a shared definition (after
 * BRIDGE_WRITE_CLIENT_SRC, whose gate it uses).
 *   __eoBridgeListDetach(bridge, capabilityId, args) -> { ok: true, moved } | { ok: false, reason, uncertain, mutation }
 *     ONE write() call. Success is exactly { ok: true, value: { applied: ['list-detached'], before: { numId }, moved } } with moved equal to
 *     the number of members the caller sent. A refusal the word /5 handler makes BEFORE its CNum.Copy() call (an explicit allowlist, incl.
 *     nothing-to-detach and not-isolatable) -> mutation NONE; copy-threw / set-threw / write-not-observed / an unknown reason -> uncertain,
 *     mutation UNKNOWN. A malformed receipt -> UNEXPECTED_SHAPE (uncertain). Never retried; the caller re-reads.
 */
const BRIDGE_LIST_DETACH_WRITE_CLIENT_SRC = `
  const __eoBridgeListDetach = (bridge, capabilityId, args) => {
    const odd = { ok: false, reason: 'UNEXPECTED_SHAPE', uncertain: true, mutation: 'UNKNOWN' };
    let r;
    try { r = bridge.write(capabilityId, args); } catch (err) { return { ok: false, reason: 'BRIDGE_THREW', uncertain: true, mutation: 'UNKNOWN' }; }
    if (!r || typeof r !== 'object' || typeof r.ok !== 'boolean') return odd;
    if (r.ok === true) {
      const v = r.value;
      const shaped = Object.keys(r).sort().join(',') === 'ok,value' && v && typeof v === 'object' && Object.keys(v).sort().join(',') === 'applied,before,moved'
        && Array.isArray(v.applied) && v.applied.join(',') === 'list-detached' && v.before && typeof v.before === 'object' && v.moved === args.members.length;
      return shaped ? { ok: true, moved: v.moved } : odd;
    }
    if (typeof r.reason !== 'string' || !r.reason) return odd;
    const PRE = ['bad-args', 'unsupported-property', 'invalid-level', 'document-unavailable', 'search-threw', 'target-not-found', 'target-ambiguous',
      'paragraph-unreadable', 'getNumbering-threw', 'level-mismatch', 'internal-missing', 'numpr-unreadable', 'abstract-unreadable', 'nums-threw',
      'nums-unreadable', 'paragraphs-unreadable', 'override-unreadable', 'member-unreadable', 'own-num-not-listed', 'not-numbered', 'members-mismatch',
      'others-mismatch', 'nothing-to-detach', 'order-unreadable', 'not-isolatable'];
    const pre = PRE.indexOf(r.reason) >= 0;
    return { ok: false, reason: 'BRIDGE:' + r.reason, uncertain: !pre, mutation: pre ? 'NONE' : 'UNKNOWN' };
  };
`;

module.exports = { BRIDGE_MAJOR, BRIDGE_CAPABILITIES, BRIDGE_CLIENT_SRC, BRIDGE_HYDRATIONS, BRIDGE_HYDRATE_CLIENT_SRC, BRIDGE_WRITES, BRIDGE_WRITE_CLIENT_SRC, BRIDGE_SERIES_STYLE_WRITE_CLIENT_SRC, BRIDGE_COPY_WRITE_CLIENT_SRC, BRIDGE_CHART_DELETE_WRITE_CLIENT_SRC, BRIDGE_NUMBERING_REMOVE_WRITE_CLIENT_SRC, BRIDGE_STYLE_NEXT_LINK_WRITE_CLIENT_SRC, BRIDGE_LIST_DETACH_WRITE_CLIENT_SRC };
