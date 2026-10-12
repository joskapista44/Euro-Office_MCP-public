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
  // B11 /3 COMBINED (internal tracker): both READ capabilities of the /2->/3 operation; the xlsx driver's
  // dependency inventory (dependentsBridgeBody) reads through them after the B11 cutover.
  DEFINED_NAMES_READ: Object.freeze({ id: 'workbook.definedNames.read/v1', version: 1, valueType: 'object' }),
  VALIDATIONS_READ: Object.freeze({ id: 'range.validations.read/v1', version: 1, valueType: 'object' }),
  // Wave B READ /4 (owner order owner channel 2026-10-08 point 9; design docs/DS-BRIDGE-WAVE-B-READ-DESIGN-20261008.md):
  // the B05 READ, B09 and B13 cutovers read through these; a /3 runtime refuses them (PATCH_VERSION_MISMATCH).
  ALIGN_EXT_READ: Object.freeze({ id: 'range.alignmentExt.read/v1', version: 1, valueType: 'object' }),
  PRINT_SETTINGS_READ: Object.freeze({ id: 'sheet.printSettings.read/v1', version: 1, valueType: 'object' }),
  PIVOT_STATE_READ: Object.freeze({ id: 'pivot.state.read/v1', version: 1, valueType: 'object' }),
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

module.exports = { BRIDGE_MAJOR, BRIDGE_CAPABILITIES, BRIDGE_CLIENT_SRC, BRIDGE_HYDRATIONS, BRIDGE_HYDRATE_CLIENT_SRC, BRIDGE_WRITES, BRIDGE_WRITE_CLIENT_SRC, BRIDGE_COPY_WRITE_CLIENT_SRC, BRIDGE_NUMBERING_REMOVE_WRITE_CLIENT_SRC };
