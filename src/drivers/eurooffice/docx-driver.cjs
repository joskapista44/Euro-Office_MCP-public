'use strict';

const { startConfigHost } = require('./config-host.cjs');
const { observePage, settleSaveAck, dsIdentity, confirmForceSave, closeResult } = require('./save-ack.cjs');
const { defaultWait } = require('../../core/session/open.cjs');
const { ACCEPTED_RUNTIMES, DOCX_RUNTIME_IDS, DOCX_B18_BRIDGE_WRITE_RUNTIME_ENTRIES, DOCX_B18_BRIDGE_WRITE_RUNTIME_IDS,
  DOCX_B21_BRIDGE_READ_RUNTIME_ENTRIES, DOCX_B21_BRIDGE_READ_RUNTIME_IDS, cacheKeyOf, HANDSHAKE_BODY, classifyHandshake } = require('../../core/runtime-identity.cjs');
const { BRIDGE_CAPABILITIES, BRIDGE_CLIENT_SRC, BRIDGE_WRITES, BRIDGE_WRITE_CLIENT_SRC, BRIDGE_NUMBERING_REMOVE_WRITE_CLIENT_SRC } = require('./ds-bridge.cjs');
const { resolveConfigSource } = require('./config-source.cjs');

/**
 * Euro-Office DOCX editor transport — the M1 driver contract on a real editor (M4).
 *
 *   open / close / forceSave / canSave / readLive / dispatch / readPackage
 *
 * This is TRANSPORT, not Word semantics and not lifecycle: M1 decides when each is called, and
 * the DOCX adapter decides what a command means. The transport behaviour mirrors the source
 * session module (joskapista44/euro-mcp f8c6d6cf docx-minimal-editor-session.cjs), which is the
 * measured reference for the editor runtime; it is re-expressed here, not copied.
 *
 * READS ARE READ-ONLY BY CONSTRUCTION. readLive returns the editor's HTML projection plus run
 * properties from PROPERTY_QUERY, a constant body that only calls Get* methods (static gate:
 * test/static/driver-read-query.test.cjs). No adapter code is ever executed as a read.
 *
 * Playwright is injected (`loadPlaywright`), so this module has no dependency the core suite needs.
 *
 * SAVE CONFIRMATION (M7, variant A -- owner GO 2026-10-03, owner channel): the XLSX driver's save sequence on the SAME
 * observer module (save-ack.cjs, reused as is; the DS word and cell SDKs speak the same save protocol, and the LIVE
 * probe measured it on this editor -- docs/evidence/M7-DOCX-SAVE-ACK-PROBE.md):
 *   open      the editor <-> DS socket of THIS session observed from its first frame (DS host + document key)
 *   dispatch  any write attempt marks the session mutated
 *   forceSave the last save acked FIRST, then ONE editor forceSave, then its OWN confirmation; a write-free session
 *             starts no forceSave
 *   close     the ack awaited BEFORE requestClose; the FINAL save-ack and confirmation evaluated AFTER requestClose and
 *             the teardown; a late send, a late anomaly or a failed close leaves no confirmation standing; a teardown
 *             failure is reported, never swallowed
 * A confirmation is the server's report of the save, not a check of the stored content.
 */

/**
 * M8 DOCX internal build gate (owner decision A, owner channel, 2026-10-05) -- NARROWED to the ONE
 * remaining internal WRITE dependency after the W-B1920-C1 bridge cutover (card
 * an internal branch; the owner, owner channel 9-10. pont: "Ha
 * B19/B20 PASS: a DOCX driver atallhat Bridge-re es a kozvetlen dependency eltavolithato. Ne legyen
 * production fallback BRIDGE -> DIRECT_INTERNAL ut."). The two internal READS this gate used to cover
 * (CNum.GetLvl(k).GetFormat, M5 gate 7; CTable.TableGridCalc, M5 gate 8) are GONE from this driver --
 * replaced by docx.numbering.levelFormat.read/v1 and docx.table.effectiveGrid.read/v1 through
 * propertyQuery()'s own BRIDGE_CLIENT_SRC call, with NO internal fallback: a bridge failure (any of
 * the five fail-closed gates, or the capability's own refusal) leaves the field unread (null), the
 * same "null = unread, never a default" contract this driver already honoured for them.
 *   CParagraph.RemoveNumPr()     WRITE -- CUT OVER (M11.5d.4) to the word /2 Bridge WRITE docx.paragraph.numbering.remove/v1
 *   CNum.GetLvl(k).GetFormat()   READ  -- CUT OVER to the Word bridge (BRIDGE_CAPABILITIES.DOCX_LIST_LEVEL_FORMAT_READ)
 *   CTable.TableGridCalc         READ  -- CUT OVER to the Word bridge (BRIDGE_CAPABILITIES.DOCX_TABLE_GRID_READ)
 */
/**
 * M11.5d.4 (owner: M11.5 continuous mandate, internal tracker): B18 runs ONLY through the word /2 Bridge WRITE
 * capability, so the numbering-removal command is narrowed to /2 -- the driver's second line behind the list family's own
 * refusal. No direct internal path and no fallback to one exists.
 */
const B18_BRIDGE_WRITE_RUNTIMES = DOCX_B18_BRIDGE_WRITE_RUNTIME_IDS;
/**
 * M11.5e.4 (owner: M11.5 continuous mandate, internal tracker): B21 runs ONLY through the word /3 Bridge READ
 * capability docx.section.variantState.read/v1, so the section-variant-state read is narrowed to /3 -- the driver's
 * second line behind the header/footer family's own refusal. No direct internal path and no fallback to one exists.
 */
const B21_BRIDGE_READ_RUNTIMES = DOCX_B21_BRIDGE_READ_RUNTIME_IDS;
/** The dependency name a command's meta.internal carries -> its build list. Empty since M11.5e.4 (B21 left it); kept for
 * any future named internal dependency and exercised by its own generic test. */
const DOCX_INTERNAL_WRITE_RUNTIMES = Object.freeze({});

/** The runtime gate of THIS session: its DS runtime identity, whether the B18 Bridge WRITE is proven on it, and whether the B21 Bridge READ is proven on it. */
function internalGateOf(runtime) {
  return Object.freeze({ runtime: runtime || null, numberingRemoveBridge: !!runtime && B18_BRIDGE_WRITE_RUNTIMES.includes(runtime),
    sectionVariant: !!runtime && B21_BRIDGE_READ_RUNTIMES.includes(runtime) });
}

/**
 * M11.5d.4: the in-editor B18 client, prepended ONLY to a numbering-removal command (every other command stays
 * byte-identical):
 *   __eoB18Remove(text, level)   the Bridge WRITE gate on the word /2 entry, then ONE docx.paragraph.numbering.remove/v1
 *                                write() through __eoBridgeNumberingRemove (ds-bridge.cjs: the receipt rule; a refusal
 *                                after the removal call is uncertain, mutation UNKNOWN). Never retried; the core re-reads.
 */
function b18BridgeWritePrelude() {
  const W = BRIDGE_WRITES.PARAGRAPH_NUMBERING_REMOVE_WRITE;
  return `${BRIDGE_CLIENT_SRC}${BRIDGE_WRITE_CLIENT_SRC}${BRIDGE_NUMBERING_REMOVE_WRITE_CLIENT_SRC}
  const __eoB18Remove = (text, level) => {
    const g = __eoBridgeWriteGate(${JSON.stringify(W.id)}, ${W.version}, ${JSON.stringify(DOCX_B18_BRIDGE_WRITE_RUNTIME_ENTRIES)});
    if (!g.ok) return { ok: false, stage: 'bridge-write-gate', reason: g.reason, mutation: 'NONE' };
    const w = __eoBridgeNumberingRemove(g.bridge, ${JSON.stringify(W.id)}, { text, level });
    return w.ok ? { ok: true } : Object.assign({ stage: 'bridge-write' }, w);
  };
`;
}
/** A B18 command is ok ONLY when its body returned { ok: true } (the bridge receipt); anything else is a failed write the core re-reads. */
function b18DispatchResult(r) {
  const res = r && r.result;
  if (r && r.ok === true && res && res.ok === true) return { ok: true, dispatched: true, outcome: r.outcome, error: null };
  return { ok: false, dispatched: true, outcome: r && r.ok === true ? 'b18-bridge-write-failed' : (r && r.outcome) || 'b18-command-failed', evidence: res || null, error: (r && r.error) || null };
}

/**
 * M11.5e.4: the in-editor B21 client, prepended ONLY to a title/even header/footer command (every other command stays
 * byte-identical):
 *   __eoB21Variant(section)   ONE docx.section.variantState.read/v1 read() through __eoBridgeRead, gated on the word /3
 *                             entry; any gate/capability/shape failure -> null, and the command leaves the state
 *                             unmutated (the same contract an unreadable/unverified reader had before the cutover).
 */
function b21BridgeReadPrelude() {
  const CAP = BRIDGE_CAPABILITIES.DOCX_SECTION_VARIANT_STATE_READ;
  return `${BRIDGE_CLIENT_SRC}
  const __eoB21Variant = (section) => {
    const r = __eoBridgeRead(${JSON.stringify(CAP.id)}, ${CAP.version}, ${JSON.stringify(DOCX_B21_BRIDGE_READ_RUNTIME_ENTRIES)}, { section }, ${JSON.stringify(CAP.valueType)});
    if (!r.ok) return null;
    const v = r.value;
    return v && typeof v === 'object' && Object.keys(v).sort().join(',') === 'evenAndOdd,titlePage' && typeof v.titlePage === 'boolean' && typeof v.evenAndOdd === 'boolean' ? v : null;
  };
`;
}

/** B21 runtime gate for the property query: an opt-out (`includeSectionVariantState: false`) or an unverified runtime
 * never runs the internal getters; the reason is reported in the flags. */
function sectionVariantQuery(request, runtimeId) {
  const wanted = !(request && request.includeSectionVariantState === false);
  const verified = internalGateOf(runtimeId).sectionVariant;
  return { sectionVariantState: wanted && verified, sectionVariantUnavailable: wanted ? 'runtime-not-verified' : 'not-requested' };
}

/** The dispatch-side refusal (second line): a B18 numbering-removal command off the word /2+/3 list, a B21 variant-read
 * command off word /3, or the internal dependencies a command names that its runtime is not verified for. */
function internalWriteRefusal(command, runtime) {
  if (command && command.meta && command.meta.bridgeNumberingRemove === true && !B18_BRIDGE_WRITE_RUNTIMES.includes(runtime)) {
    return { ok: false, dispatched: false, outcome: 'b18-bridge-write-runtime-not-proven', runtime, proven: B18_BRIDGE_WRITE_RUNTIMES };
  }
  if (command && command.meta && command.meta.bridgeSectionVariantRead === true && !B21_BRIDGE_READ_RUNTIMES.includes(runtime)) {
    return { ok: false, dispatched: false, outcome: 'b21-bridge-read-runtime-not-proven', runtime, proven: B21_BRIDGE_READ_RUNTIMES };
  }
  const internal = command && command.meta && Array.isArray(command.meta.internal) ? command.meta.internal : [];
  const unverified = internal.filter(d => !(DOCX_INTERNAL_WRITE_RUNTIMES[d] || []).includes(runtime));
  return unverified.length ? { ok: false, dispatched: false, outcome: 'internal-docx-runtime-not-verified', runtime, dependencies: unverified } : null;
}

/** Closes the browser; a browser.close() failure is REPORTED (local to this driver; the XLSX driver has its own). */
async function closeBrowser(browser) {
  if (!browser) return { ok: true, detail: 'no-browser' };
  return browser.close().then(() => ({ ok: true, detail: 'torn-down' }),
    err => ({ ok: false, detail: `browser-close-failed: ${String(err && err.message || err).slice(0, 160)}` }));
}

/**
 * W-B1920-C1 (word-bridge cutover): the numbering/table reads go through __eoBridgeRead, gated on the
 * served Word bundle's stable manifest identity (core/runtime-identity.cjs RUNTIMES.docx; owner decision (B), card
 * 9580767b) -- not on the editor URL's cache key. `PROPERTY_QUERY` stays exported (the same string).
 */
/**
 * `sectionVariantState` (default true): read the B21 section variant state (title-page / even-odd flags) through the
 * word /3 Bridge READ capability. When false the generated body does not contain the capability id at all and every
 * section reports `UNAVAILABLE:<sectionVariantUnavailable>` for both flags.
 */
function propertyQuery({ sectionVariantState = true, sectionVariantUnavailable = 'not-requested' } = {}) {
  const unavailable = JSON.stringify(`UNAVAILABLE:${sectionVariantUnavailable}`);
  const variantFields = sectionVariantState ? '...variant(sec, index)' : `titlePage: ${unavailable}, evenAndOdd: ${unavailable}`;
  return `
try {
  const textOpts = { Numbering: false };
  const doc = Api.GetDocument();
  // The DS Bridge client (fail-closed, no fallback): the two Word read capabilities this driver
  // now depends on for list-level format and table grid -- no internal fallback on any failure.
${BRIDGE_CLIENT_SRC}  const rd = (cap, args) => { const r = __eoBridgeRead(cap.id, cap.version, ${JSON.stringify(ACCEPTED_RUNTIMES.docx)}, args, cap.valueType); return r.ok ? r.value : null; };
  // The generic bridge client only checks typeof value === 'object' -- not the PER-CAPABILITY field
  // shape. A malformed/unexpected result (ok:true with the wrong shape) must still fail closed here.
  const validNumbering = v => {
    if (!v || typeof v !== 'object') return false;
    if (v.present === false) return true;
    if (v.present !== true) return false;
    const lf = v.levelFormats;
    return (v.level === null || Number.isInteger(v.level)) && (v.numId === null || typeof v.numId === 'string')
      && (lf === null || (Array.isArray(lf) && lf.length === 9 && lf.every(x => x === null || Number.isInteger(x))));
  };
  // W-08: colour's return shape, MEASURED (docx-w08-getcolor-shape-probe), not
  // guessed -- the first hypothesis's speculative branches (GetRGB as number/string, GetR/G/B,
  // bare r/g/b properties) are gone. GetColor() on a run with no explicit colour answers null,
  // a DEFINED, observable state ('auto'), not an unread one; after a write it answers an ApiColor
  // object whose IsAutoColor() / GetHex() are the measured reads. undefined (the getter itself
  // absent) stays unread -- the key is never set, never guessed.
  const colorHex = col => {
    if (col === undefined) return undefined;
    if (col === null) return 'auto';
    if (typeof col.IsAutoColor === 'function' && col.IsAutoColor()) return 'auto';
    if (typeof col.GetHex === 'function') {
      const h = String(col.GetHex() || '').replace('#', '');
      if (/^[0-9a-fA-F]{6}$/.test(h)) return '#' + h.toUpperCase();
    }
    return undefined;
  };
  const out = [];
  const n = doc.GetElementsCount();
  for (let i = 0; i < n; i += 1) {
    const el = doc.GetElement(i);
    if (!el || typeof el.GetClassType !== 'function' || el.GetClassType() !== 'paragraph') continue;
    const text = String(el.GetText(textOpts) || '').replace(/[\\r\\n]+$/, '');
    if (!text.trim()) continue;
    // W1.1: the paragraph's runs, needle-free -- text, offset and the EFFECTIVE text properties of
    // each run's own range (a run is uniform by definition; measured live 2026-09-24, probe 2:
    // run.GetTextPr() is DIRECT only, run.GetRange().GetTextPr() resolves inheritance). Any non-run
    // child is counted: then the run list does not cover the paragraph and no range claim is made.
    const runs = [];
    let nonRunChildren = 0;
    let offset = 0;
    const k = typeof el.GetElementsCount === 'function' ? el.GetElementsCount() : -1;
    if (k < 0) nonRunChildren = -1;
    for (let j = 0; j < k; j += 1) {
      const child = el.GetElement(j);
      if (!child || typeof child.GetClassType !== 'function' || child.GetClassType() !== 'run' || typeof child.GetRange !== 'function') { nonRunChildren += 1; continue; }
      const t = String(child.GetText(textOpts) || '');
      const rr = child.GetRange();
      const rp = rr && typeof rr.GetTextPr === 'function' ? rr.GetTextPr() : null;
      const rget = m => (rp && typeof rp[m] === 'function' ? rp[m]() : undefined);
      const rf = {};
      for (const [key, m] of [['bold', 'GetBold'], ['italic', 'GetItalic'], ['underline', 'GetUnderline'], ['strike', 'GetStrikeout']]) { const v = rget(m); if (typeof v === 'boolean') rf[key] = v; }
      const ff = rget('GetFontFamily'); if (ff) rf.fontFamily = String(ff);
      const fs = rget('GetFontSize'); if (fs !== undefined && fs !== null && Number.isFinite(Number(fs))) { const v = Number(fs); rf.fontSize = v > 16 ? v / 2 : v; }
      const lg = rget('GetLanguage'); if (lg) rf.language = String(lg);
      // W-08 (internal tracker): highlight / vertAlign read back as plain strings; colour via the
      // shared colorHex normalizer (never a raw object crosses the frame boundary).
      const col = colorHex(rget('GetColor')); if (col) rf.color = col;
      const hl = rget('GetHighlight'); if (hl) rf.highlight = String(hl);
      const va = rget('GetVertAlign'); if (va) rf.vertAlign = String(va);
      runs.push({ text: t, start: offset, end: offset + t.length, format: rf });
      offset += t.length;
    }
    // Wave 2 (list family) -> W-B1920-C1 cutover: the paragraph's numbering through
    // docx.numbering.levelFormat.read/v1 ONLY. The bridge's own implementation already re-derives
    // the public level (ApiNumberingLevel.GetLevelIndex) and the internal NumPr.NumId it used to
    // read here directly -- this driver no longer touches the internal per-level format reader at
    // all, and a bridge failure (any of its five fail-closed gates, or the capability's own refusal) leaves
    // numbering unread (null), never a guessed default and never a direct-internal fallback.
    const numReq = rd(${JSON.stringify(BRIDGE_CAPABILITIES.DOCX_LIST_LEVEL_FORMAT_READ)}, { paragraph: i });
    const numbering = validNumbering(numReq) ? numReq : null;
    // Wave 3 (rich-inline + notes), needle-free, measured live 2026-09-24 (rich probes 1-3):
    // a hyperlink is a child of ClassType 'hyperlink' (url, screen tip, displayed text, effective
    // format of its range); a note reference is an empty run whose DIRECT style is 'footnote
    // reference' / 'endnote reference'. The GetText above excludes both markers.
    const hyperlinks = [];
    const noteRefs = { footnote: 0, endnote: 0 };
    for (let j = 0; j < k; j += 1) {
      const child = el.GetElement(j);
      const cls = child && typeof child.GetClassType === 'function' ? child.GetClassType() : null;
      if (cls === 'hyperlink') {
        const hr = typeof child.GetRange === 'function' ? child.GetRange() : null;
        const hp = hr && typeof hr.GetTextPr === 'function' ? hr.GetTextPr() : null;
        const style = hp && typeof hp.GetStyle === 'function' ? hp.GetStyle() : null;
        hyperlinks.push({ url: String(child.GetLinkedText() || ''), tip: String(child.GetScreenTipText() || ''), text: String(child.GetDisplayedText() || ''),
          style: style && typeof style.GetName === 'function' ? String(style.GetName()) : null });
      }
      if (cls === 'run' && typeof child.GetTextPr === 'function') {
        const dp = child.GetTextPr();
        const st = dp && typeof dp.GetStyle === 'function' ? dp.GetStyle() : null;
        const sn = st && typeof st.GetName === 'function' ? String(st.GetName()) : '';
        if (sn === 'footnote reference') noteRefs.footnote += 1;
        if (sn === 'endnote reference') noteRefs.endnote += 1;
      }
    }
    const runInfo = { runs, nonRunChildren, runsCoverText: runs.map(r => r.text).join('') === text, hyperlinks, noteRefs, ...(numbering ? { numbering } : {}) };
    const hits = (doc.Search(text, { MatchCase: true }) || []).filter(r => String(r.GetText(textOpts) || '') === text);
    if (hits.length !== 1) { out.push({ text, props: null, ...runInfo }); continue; }
    const p = hits[0].GetTextPr();
    if (!p) { out.push({ text, props: null, ...runInfo }); continue; }
    const get = m => (typeof p[m] === 'function' ? p[m]() : undefined);
    const props = { bold: get('GetBold'), italic: get('GetItalic'), underline: get('GetUnderline'),
      strike: get('GetStrikeout'), fontFamily: get('GetFontFamily'), fontSize: get('GetFontSize'), language: get('GetLanguage'),
      color: colorHex(get('GetColor')), highlight: get('GetHighlight'), vertAlign: get('GetVertAlign') };
    // Whole-paragraph formatting, normalized as the source's direct read does
    // (docx-session-formatting.cjs readSemanticFormatting): booleans, font size in points,
    // alignment name, spacing in points, line spacing, style name, language.
    const format = {};
    for (const k of ['bold', 'italic', 'underline', 'strike']) if (props[k] !== undefined) format[k] = !!props[k];
    if (props.fontFamily) format.fontFamily = String(props.fontFamily);
    if (Number.isFinite(Number(props.fontSize))) { const n = Number(props.fontSize); format.fontSize = n > 16 ? n / 2 : n; }
    if (props.language) format.language = String(props.language);
    if (props.color) format.color = props.color;
    if (props.highlight) format.highlight = String(props.highlight);
    if (props.vertAlign) format.vertAlign = String(props.vertAlign);
    if (typeof el.GetStyle === 'function') { const st = el.GetStyle(); if (st && typeof st.GetName === 'function' && st.GetName()) format.style = String(st.GetName()); }
    const pp = typeof el.GetParaPr === 'function' ? el.GetParaPr() : null;
    if (pp) {
      if (typeof pp.GetJc === 'function') {
        const v = pp.GetJc();
        if (typeof v === 'string' && ['left', 'right', 'center', 'justify'].includes(v.toLowerCase())) format.alignment = v.toLowerCase();
        else { const jc = { 0: 'right', 1: 'left', 2: 'center', 3: 'justify' }; if (Object.prototype.hasOwnProperty.call(jc, v)) format.alignment = jc[v]; }
      }
      if (typeof pp.GetSpacingBefore === 'function') { const v = Number(pp.GetSpacingBefore()); if (Number.isFinite(v)) format.spaceBefore = v / 20; }
      if (typeof pp.GetSpacingAfter === 'function') { const v = Number(pp.GetSpacingAfter()); if (Number.isFinite(v)) format.spaceAfter = v / 20; }
      if (typeof pp.GetSpacingLineValue === 'function') { const n = Number(pp.GetSpacingLineValue()); if (Number.isFinite(n)) format.lineSpacing = n > 10 ? (n >= 200 ? n / 240 : (n + 1) / 100) : n; }
      // W-09 (indentation + shading): same points-externally/twips-internally convention as
      // spaceBefore/spaceAfter above (measured: SetIndLeft(720) round-trips as 720 in the saved
      // <w:ind>, i.e. the Builder's own unit is already twips). GetIndFirstLine() reports the
      // signed value as set, even when the file stores a negative first-line as <w:hanging> --
      // no sign handling needed here.
      if (typeof pp.GetIndLeft === 'function') { const v = Number(pp.GetIndLeft()); if (Number.isFinite(v)) format.indentLeft = v / 20; }
      if (typeof pp.GetIndRight === 'function') { const v = Number(pp.GetIndRight()); if (Number.isFinite(v)) format.indentRight = v / 20; }
      if (typeof pp.GetIndFirstLine === 'function') { const v = Number(pp.GetIndFirstLine()); if (Number.isFinite(v)) format.indentFirstLine = v / 20; }
      // shading as a 6-digit hex string, 'auto' when none is set -- an ABSENT key would be
      // unobservable (compare() treats undefined like null) and fail-close every first request.
      // a maintainer review: MEASURED shapes -- an untouched paragraph's GetShd() is null (W-09 matrix,
      // control), a set one an ApiColor whose GetRGB() is an {r,g,b} OBJECT and GetHex() '#RRGGBB'
      // (W-08 shape probe). null / IsAutoColor() -> 'auto'; else GetHex(); anything else stays unread.
      if (typeof pp.GetShd === 'function') {
        const shd = pp.GetShd();
        if (shd === null) format.shading = 'auto';
        else if (shd && typeof shd.IsAutoColor === 'function' && shd.IsAutoColor()) format.shading = 'auto';
        else if (shd && typeof shd.GetHex === 'function') {
          const h = String(shd.GetHex() || '').replace('#', '');
          if (/^[0-9a-fA-F]{6}$/.test(h)) format.shading = h.toUpperCase();
        }
      }
      // W-09 tabs / per-side border / pagination flags (W09-1a
      // through 1e): NO Get-counterpart exists for ANY of these (MEASURED, not merely unread) --
      // ParaPr.ToJSON() is the ONLY reader. An ABSENT key is a DEFINITE state here (OOXML's own
      // presence-only semantics: no w:tabs/w:pBdr/w:keepNext element means none/off), never
      // UNOBSERVABLE -- UNLIKE style-definition.cjs's 'inherited' convention, which exists
      // because a STYLE inherits from a base style; a plain document PARAGRAPH does not.
      // Units: tabs pos and border size round-trip UNCHANGED through the raw ToJSON value
      // (W09-1b/1c: SetTabs pos 720 -> saved w:pos=720 -> ToJSON pos 720; Set*Border sz 8 ->
      // saved w:sz=8 -> ToJSON sz 8), confirming TWIPS for pos and EIGHTHS-OF-A-POINT for sz --
      // divided back to points here (/20, /8), the SAME externally-points convention every other
      // paragraph measurement in this file already uses.
      if (typeof pp.ToJSON === 'function') {
        let ptj = null;
        try {
          const rawPp = pp.ToJSON();
          ptj = typeof rawPp === 'string' ? JSON.parse(rawPp) : rawPp;
        } catch (_) { ptj = null; }
        if (ptj && typeof ptj === 'object') {
          format.tabs = Array.isArray(ptj.tabs)
            ? ptj.tabs.map(t => ({ pos: Number.isFinite(Number(t && t.pos)) ? Number(t.pos) / 20 : null, align: t && typeof t.val === 'string' ? t.val : null }))
            : [];
          const pBdr = ptj.pBdr && typeof ptj.pBdr === 'object' ? ptj.pBdr : {};
          const hx2 = n => Number(n).toString(16).padStart(2, '0').toUpperCase();
          const borderSideFormat = side => {
            if (!side || typeof side !== 'object') return null;
            const c = side.color && typeof side.color === 'object' ? side.color : null;
            const color = c && Number.isInteger(c.r) && Number.isInteger(c.g) && Number.isInteger(c.b) ? hx2(c.r) + hx2(c.g) + hx2(c.b) : (c && c.auto === true ? 'auto' : null);
            return { style: typeof side.value === 'string' ? side.value : null, size: Number.isFinite(Number(side.sz)) ? Number(side.sz) / 8 : null, color };
          };
          format.borderTop = borderSideFormat(pBdr.top);
          format.borderBottom = borderSideFormat(pBdr.bottom);
          format.borderLeft = borderSideFormat(pBdr.left);
          format.borderRight = borderSideFormat(pBdr.right);
          format.keepNext = ptj.keepNext === true;
          format.keepLines = ptj.keepLines === true;
          format.pageBreakBefore = ptj.pageBreakBefore === true;
          format.widowControl = ptj.widowControl === true;
        }
      }
    }
    out.push({ text, props, format, ...runInfo });
  }
  // Wave 3, document level: bookmarks (name + the text of its range), the first paragraph of every
  // footnote/endnote, and whether the editor's cursor is in the MAIN body (GetCurrentFootEndnote()
  // null) -- after a note write it is not, until the note command restores it (rich probes 2-3).
  // Each field is read on its own; an unreadable one is absent (unknown), never an empty list.
  const docInfo = {};
  const noteText = p => String(p.GetText(textOpts) || '').replace(/^\\s+|[\\r\\n]+$/g, '');
  try { docInfo.bookmarks = (doc.GetAllBookmarksNames() || []).map(name => { const r = doc.GetBookmarkRange(name); return { name: String(name), text: r ? String(r.GetText(textOpts) || '') : null }; }); } catch (_) { /* unread */ }
  try { docInfo.footnotes = (doc.GetFootnotesFirstParagraphs() || []).map(noteText); } catch (_) { /* unread */ }
  try { docInfo.endnotes = (doc.GetEndNotesFirstParagraphs() || []).map(noteText); } catch (_) { /* unread */ }
  try { docInfo.inMainBody = doc.GetCurrentFootEndnote() === null; } catch (_) { /* unread */ }
  // Wave 4a (table family), needle-free, measured live 2026-09-24 (table probes 1-2): every table
  // through the public GetAllTables / GetRow / GetCell; per cell its content text (the editor ends
  // it with '\\t' or '\\r\\n', a merged cell joins its paragraphs with '\\r\\n'), its element count,
  // whether it holds ONLY paragraphs made ONLY of runs (a hyperlink, a note reference or a bookmark
  // marker is a non-run child), and the effective text/paragraph properties of its FIRST paragraph.
  //
  // Wave 4c (column probe 1, 2026-09-24): the font family is read per RUN, not from the paragraph
  // range. A cell paragraph made by the editor starts with an EMPTY run in the default font, so the
  // paragraph range's GetFontFamily() is undefined even on a uniform Courier New cell; the public
  // run.GetRange().GetTextPr().GetFontFamily() of the NON-EMPTY runs is exact. Runs that disagree
  // -> 'MIXED' (never one invented family); no non-empty run -> the paragraph range's value.
  function cellFirstParagraphFont(p) {
    const fams = new Set();
    for (let j = 0; j < p.GetElementsCount(); j += 1) {
      const ch = p.GetElement(j);
      if (!ch || ch.GetClassType() !== 'run' || String(ch.GetText(textOpts) || '') === '') continue;
      const rp = ch.GetRange().GetTextPr();
      const f = rp && typeof rp.GetFontFamily === 'function' ? rp.GetFontFamily() : undefined;
      fams.add(f ? String(f) : '');
    }
    if (fams.size > 1) return 'MIXED';
    if (fams.size === 1) return [...fams][0] || undefined;
    const pr = p.GetRange().GetTextPr();
    const f = pr && typeof pr.GetFontFamily === 'function' ? pr.GetFontFamily() : undefined;
    return f ? String(f) : undefined;
  }
  function paragraphFormat(p) {
    const format = {};
    const tp = p.GetRange().GetTextPr();
    const tget = k => (tp && typeof tp[k] === 'function' ? tp[k]() : undefined);
    for (const [key, g] of [['bold', 'GetBold'], ['italic', 'GetItalic'], ['underline', 'GetUnderline'], ['strike', 'GetStrikeout']]) { const v = tget(g); if (typeof v === 'boolean') format[key] = v; }
    const ff = cellFirstParagraphFont(p); if (ff) format.fontFamily = ff;
    const fs = tget('GetFontSize'); if (fs !== undefined && fs !== null && Number.isFinite(Number(fs))) { const v = Number(fs); format.fontSize = v > 16 ? v / 2 : v; }
    const st = p.GetStyle(); if (st && st.GetName()) format.style = String(st.GetName());
    const pp = p.GetParaPr(); const jc = pp ? pp.GetJc() : undefined; if (jc !== undefined && jc !== null) format.alignment = String(jc);
    return format;
  }
  try {
    docInfo.tables = (doc.GetAllTables() || []).map((t, ti) => {
      const rows = [];
      for (let r = 0; r < t.GetRowsCount(); r += 1) {
        const row = t.GetRow(r);
        const cells = [];
        for (let c = 0; c < row.GetCellsCount(); c += 1) {
          const content = row.GetCell(c).GetContent();
          const m = content.GetElementsCount();
          let simple = m > 0;
          for (let i = 0; i < m; i += 1) {
            const e = content.GetElement(i);
            if (!e || e.GetClassType() !== 'paragraph') { simple = false; continue; }
            for (let j = 0; j < e.GetElementsCount(); j += 1) { const ch = e.GetElement(j); if (!ch || ch.GetClassType() !== 'run') simple = false; }
          }
          const first = m > 0 ? content.GetElement(0) : null;
          const format = first && first.GetClassType() === 'paragraph' ? paragraphFormat(first) : {};
          // Wave 4d (merge/split probe 1): a merged cell keeps EACH source paragraph with its own runs
          // and alignment -- the per-paragraph text and format, read the same way as the first's
          const paraFormats = [];
          for (let i = 0; i < m; i += 1) { const e = content.GetElement(i); if (e && e.GetClassType() === 'paragraph') paraFormats.push({ text: String(e.GetText(textOpts) || ''), format: paragraphFormat(e) }); }
          cells.push({ text: String(content.GetText(textOpts) || ''), paragraphs: m, simple, format, paraFormats });
        }
        rows.push(cells);
      }
      // M5 gate 8 -> W-B1920-C1 cutover: the CALCULATED grid column widths through
      // docx.table.effectiveGrid.read/v1 ONLY -- no internal calculated-grid fallback. null =
      // unread (bridge failure or an invalid shape); the split family then refuses a column split
      // before dispatch, exactly as it did for an unverified-build internal read.
      const gridRead = rd(${JSON.stringify(BRIDGE_CAPABILITIES.DOCX_TABLE_GRID_READ)}, { table: ti });
      const gridCalc = gridRead && Array.isArray(gridRead.gridCalc) ? gridRead.gridCalc : null;
      return { rows, gridCalc };
    });
  } catch (_) { /* unread */ }
  // Wave 5 (header/footer probes 1-2, 2026-09-25): the HTML projection shows NO header/footer text,
  // so every section is read through the public GetSections / GetHeader(v, false) /
  // GetFooter(v, false) -- false: a read never creates a part. Per section, per kind, per variant:
  // null = the section has no part of its own; else the part's paragraphs (text + format) and
  // whether it holds ONLY paragraphs made ONLY of runs. The section object is a new wrapper per
  // call (GetFinalSection() !== GetSections()[i]), so the identity is the index.
  try {
    const partOf = hf => {
      if (!hf) return null;
      const m = hf.GetElementsCount();
      let simple = m > 0;
      const paras = [];
      for (let i = 0; i < m; i += 1) {
        const e = hf.GetElement(i);
        if (!e || e.GetClassType() !== 'paragraph') { simple = false; continue; }
        for (let j = 0; j < e.GetElementsCount(); j += 1) { const ch = e.GetElement(j); if (!ch || ch.GetClassType() !== 'run') simple = false; }
        // M5 gate 6 (owner decision 1 after 28947): the EFFECTIVE properties of every run (public
        // run.GetRange().GetTextPr(), the read header/footer probe 2 used) -- the replacement guard
        // compares them with what the full replacement yields. A throwing read -> runs null (unread).
        let runs = [];
        try {
          for (let j = 0; j < e.GetElementsCount(); j += 1) {
            const ch = e.GetElement(j);
            if (!ch || ch.GetClassType() !== 'run') continue;
            const tp = ch.GetRange().GetTextPr();
            const g = k => (tp && typeof tp[k] === 'function' ? tp[k]() : undefined);
            runs.push({ text: String(ch.GetText(textOpts) || ''), bold: g('GetBold'), italic: g('GetItalic'), underline: g('GetUnderline'), strike: g('GetStrikeout'), fontFamily: g('GetFontFamily'), fontSize: g('GetFontSize') });
          }
        } catch (_) { runs = null; }
        paras.push({ text: String(e.GetText(textOpts) || ''), format: paragraphFormat(e), runs });
      }
      return { paragraphs: m, simple, paras };
    };
    ${sectionVariantState ? `// B21 (M11.5e.4, owner decision after owner channel): the title-page and even/odd flags have no public getter;
    // they are read through the word /3 Bridge READ capability docx.section.variantState.read/v1 -- no direct internal
    // fallback. Any gate/capability/shape failure gives 'UNAVAILABLE:bridge-unavailable' for BOTH flags, never a guess;
    // the family then refuses a variant write before dispatch.
    const __eoB21Unavailable = { titlePage: 'UNAVAILABLE:bridge-unavailable', evenAndOdd: 'UNAVAILABLE:bridge-unavailable' };
    const variant = (sec, index) => {
      const v = rd(${JSON.stringify(BRIDGE_CAPABILITIES.DOCX_SECTION_VARIANT_STATE_READ)}, { section: index });
      return v && typeof v === 'object' && Object.keys(v).sort().join(',') === 'evenAndOdd,titlePage' && typeof v.titlePage === 'boolean' && typeof v.evenAndOdd === 'boolean' ? v : __eoB21Unavailable;
    };` : ''}
    // M5 gate 6: the document's DEFAULT text properties (public ApiDocument.GetDefaultTextPr) -- the run
    // properties a full replacement falls back to. Absent / throwing / non-scalar -> null (unread); the
    // header/footer family then refuses every EXISTING-part replacement before dispatch.
    try {
      const dt = typeof doc.GetDefaultTextPr === 'function' ? doc.GetDefaultTextPr() : null;
      const ff = dt && typeof dt.GetFontFamily === 'function' ? dt.GetFontFamily() : undefined;
      const fz = dt && typeof dt.GetFontSize === 'function' ? dt.GetFontSize() : undefined;
      docInfo.defaultText = typeof ff === 'string' && typeof fz === 'number' ? { fontFamily: ff, fontSize: fz } : null;
    } catch (_) { docInfo.defaultText = null; }
    docInfo.sections = (doc.GetSections() || []).map((sec, index) => {
      const one = { index, header: {}, footer: {}, ${variantFields} };
      for (const v of ['default', 'even', 'title']) { one.header[v] = partOf(sec.GetHeader(v, false)); one.footer[v] = partOf(sec.GetFooter(v, false)); }
      return one;
    });
  } catch (_) { /* unread */ }
  return { paragraphs: out, document: docInfo };
} catch (err) { return { error: String(err) }; }
`;
}
const PROPERTY_QUERY = propertyQuery();

/**
 * The M7 save bounds of the DOCX driver -- fixed, never extended at run time; every wait is condition-based.
 * forceSaveAckTimeoutMs: the wait for the editor's LAST saveChanges to be acked before the forceSave request. 120 s
 * (owner GO 2026-10-03, after the 7a1d5ea review): a 67-part / ~111 MB DOCX save was acked ~84 s after the change
 * (docs/evidence/M7-DOCX-WIRING-REVIEW.md 3.2). The confirmation and the close keep the XLSX driver's 30 s: their own
 * measured waits were ~1 s (the forcesave answer) and immediate (the ack is already in at close).
 */
const DOCX_SAVE_BOUNDS = Object.freeze({ forceSaveAckTimeoutMs: 120000, forceSaveConfirmTimeoutMs: 30000, closeAckTimeoutMs: 30000 });

function createEuroOfficeDocxDriver({ url, user, pass, loadPlaywright, timeoutMs = 60000, pollMs = 100, commandTimeoutMs = 15000,
  closeAckTimeoutMs = DOCX_SAVE_BOUNDS.closeAckTimeoutMs, forceSaveAckTimeoutMs = DOCX_SAVE_BOUNDS.forceSaveAckTimeoutMs,
  forceSaveConfirmTimeoutMs = DOCX_SAVE_BOUNDS.forceSaveConfirmTimeoutMs,
  // M11.2: the editor-config source (config-source.cjs). The PRODUCT passes the broker source and NO user/pass;
  // absent = the legacy Basic-auth source built from url/user/pass (test/harness only)
  configSource = null }) {
  const source = resolveConfigSource({ configSource, url, user, pass });
  const state = { browser: null, page: null, frame: null, host: null, config: null, runtimeId: null, cacheKey: null, runtimeRefusal: null, saveAck: null, mutated: false, forceSaveOp: null };

  // M11.2: the editor config comes from the config source seam (config-source.cjs) -- the driver never holds a
  // credential of its own; without a configSource the legacy Basic-auth source (test/harness only) is used
  async function fetchConfig(fileId, operation) {
    return source.getEditorConfig({ fileId: String(fileId), operation: operation || null });
  }

  // M11.3b.1: the session's loopback config host is token-protected (config-host.cjs)
  const startHost = startConfigHost;

  const inEditor = fn => state.frame.evaluate(fn);

  async function runInEditor(body) {
    return state.frame.evaluate(({ body: src, limit }) => new Promise(resolve => {
      const e = window.editor || (window.Asc || {}).editor;
      if (!e || typeof e.callCommand !== 'function') { resolve({ ok: false, outcome: 'callcommand-unavailable' }); return; }
      let done = false;
      const finish = v => { if (!done) { done = true; clearTimeout(timer); resolve(v); } };
      const timer = setTimeout(() => finish({ ok: false, outcome: 'command-callback-timeout' }), limit);
      try { e.callCommand(new Function(src), true, r => finish({ ok: true, outcome: 'command-completed', result: r === undefined ? null : r })); }
      catch (err) { finish({ ok: false, outcome: 'call-threw', error: String(err && (err.message || err)).slice(0, 500) }); }
    }), { body, limit: commandTimeoutMs });
  }

  /** The session's DS runtime identity -- one canonical bridge handshake (core/runtime-identity.cjs); the URL cache key is diagnostic only. */
  async function establishRuntime() {
    state.cacheKey = cacheKeyOf(state.frame.url());
    const r = await runInEditor(HANDSHAKE_BODY).catch(err => ({ ok: false, outcome: `threw:${String(err && err.message || err).slice(0, 80)}` }));
    const c = r && r.ok ? classifyHandshake(r.result, ACCEPTED_RUNTIMES.docx) : { ok: false, reason: `HANDSHAKE_FAILED:${(r && r.outcome) || 'unknown'}` };
    state.runtimeId = c.ok ? c.runtimeId : null;
    state.runtimeRefusal = c.ok ? null : c.reason;
  }

  async function teardown() {
    const torn = await closeBrowser(state.browser);
    if (state.host) await state.host.close();
    state.browser = state.page = state.frame = state.host = state.saveAck = state.forceSaveOp = state.runtimeId = state.cacheKey = state.runtimeRefusal = null;
    return torn;
  }

  async function canSaveNow() {
    return inEditor(() => {
      const e = window.editor || (window.Asc || {}).editor;
      try {
        const v = typeof e.asc_isDocumentCanSave === 'function' ? e.asc_isDocumentCanSave() : e.canSave;
        return typeof v === 'boolean' ? { ok: true, canSave: v } : { ok: false, canSave: null, detail: 'canSave not boolean' };
      } catch (err) { return { ok: false, canSave: null, detail: String(err) }; }
    });
  }

  const readCanSave = () => canSaveNow().then(r => (r && r.ok ? r.canSave : null), () => null);

  return {
    async open({ fileId, operation = null }) {
      const loaded = loadPlaywright();
      if (!loaded.ok) return { ok: false, detail: loaded.detail || 'playwright unavailable' };
      try {
        state.config = await fetchConfig(fileId, operation);
        state.host = await startHost(state.config);
        state.browser = await loaded.pw.chromium.launch();
        state.page = await state.browser.newPage({ viewport: { width: 1400, height: 900 } });
        // the editor <-> DS socket of THIS session: the config's DS host and document key (the socket's shardkey)
        state.mutated = false;
        state.forceSaveOp = null;
        state.saveAck = observePage(state.page, dsIdentity(state.config));
        await state.page.goto(state.host.url, { waitUntil: 'domcontentloaded', timeout: timeoutMs });
        await state.page.waitForFunction(() => window.__ready === true || window.__error, null, { timeout: timeoutMs, polling: pollMs });
        // Condition-based: the editor is open once its live model answers a non-empty projection.
        const deadline = Date.now() + timeoutMs;
        while (Date.now() < deadline) {
          state.frame = state.page.frames().find(f => /documenteditor/.test(f.url())) || null;
          if (state.frame) {
            const live = await inEditor(() => { const e = window.editor || (window.Asc || {}).editor; return !!(e && e.pluginMethod_GetFileHTML && String(e.pluginMethod_GetFileHTML() || '').trim()); }).catch(() => false);
            if (live) { await establishRuntime(); return { ok: true, id: `eo-${fileId}-${Date.now()}` }; }
          }
          await state.page.waitForTimeout(pollMs);
        }
        await teardown();
        return { ok: false, detail: 'live editor model not ready' };
      } catch (err) {
        // M11.3b.1: a navigation error carries the host URL -- the session token is redacted before anything leaves
        const redact = state.host ? state.host.redact : (x => x);
        await teardown();
        // a controlled config-source refusal (M11.2) keeps its product error class; the editor was never started
        if (err && err.name === 'ConfigSourceError') return { ok: false, code: err.code, detail: err.detail || err.code };
        return { ok: false, detail: redact(String(err && err.message || err)) };
      }
    },

    async close() {
      // the observer and the forceSave operation window are kept until the END of the close
      const obs = state.page ? state.saveAck : null;
      const op = state.forceSaveOp;
      const pre = await settleSaveAck(obs, { timeoutMs: closeAckTimeoutMs, pollMs, wait: defaultWait, mutated: state.mutated, readCanSave });
      const mark = obs ? obs.mark() : null;
      let lifecycle = { ok: false, detail: 'no editor' };
      if (state.page) {
        lifecycle = await state.page.evaluate(async () => {
          const d = window.__docEditor;
          if (!d || typeof d.requestClose !== 'function') return { ok: false, detail: 'requestClose-unavailable' };
          try { const r = d.requestClose(); if (r && typeof r.then === 'function') await r; return { ok: true, detail: 'requestClose-called' }; }
          catch (err) { return { ok: false, detail: String(err) }; }
        }).catch(err => ({ ok: false, detail: String(err) }));
      }
      const torn = await teardown();
      await defaultWait(0); // deliver the websocket events queued during the teardown before the final verdict
      return closeResult(obs, { pre, mark, op, lifecycle, torn });
    },

    async forceSave() {
      const r = await confirmForceSave(state.page && state.saveAck, { readCanSave, mutated: state.mutated, pollMs, wait: defaultWait,
        ackTimeoutMs: forceSaveAckTimeoutMs, confirmTimeoutMs: forceSaveConfirmTimeoutMs,
        requestForceSave: () => inEditor(() => {
          const e = window.editor || (window.Asc || {}).editor;
          if (!e || typeof e.forceSave !== 'function') return { ok: false, detail: 'forceSave-unavailable' };
          return { ok: true, sent: e.forceSave() === true };
        }) });
      if (r.forceSave && r.forceSave.op) state.forceSaveOp = r.forceSave.op; // re-evaluated at close
      return r;
    },

    async canSave() {
      return canSaveNow();
    },

    /** request.includeSectionVariantState (default true): false = the B21 internal getters are not called at all. */
    async readLive(_handle, request = {}) {
      const html = await inEditor(() => { const e = window.editor || (window.Asc || {}).editor; return String(e.pluginMethod_GetFileHTML() || ''); });
      if (!html.trim()) return { ok: false, detail: 'empty live projection' };
      const props = await runInEditor(propertyQuery(sectionVariantQuery(request, state.runtimeId)));
      const r = props.ok ? props.result : null;
      const paragraphs = r && Array.isArray(r.paragraphs) ? r.paragraphs : null;
      const document = r && r.document && typeof r.document === 'object' ? r.document : null;
      const gate = internalGateOf(state.runtimeId);
      return { ok: true, authority: 'LIVE_READ', payload: { html, paragraphs, document, runtime: { internalGate: gate, identity: state.runtimeId, cacheKey: state.cacheKey } } };
    },

    async dispatch(_handle, command) {
      // M8 runtime gate, second line: a command that names an internal dependency is refused on a runtime not verified
      // for it, before anything is sent (the adapter already refused it before the write)
      const refused = internalWriteRefusal(command, state.runtimeId);
      if (refused) return refused;
      state.mutated = true; // any write attempt: from here a close needs an observed, acked save
      if (!command || command.kind !== 'callCommand' || typeof command.body !== 'string') {
        return { ok: false, dispatched: false, detail: 'unsupported command kind' };
      }
      if (command.meta && command.meta.bridgeNumberingRemove === true) return b18DispatchResult(await runInEditor(b18BridgeWritePrelude() + command.body));
      const prelude = command.meta && command.meta.bridgeSectionVariantRead === true ? b21BridgeReadPrelude() : '';
      const r = await runInEditor(prelude + command.body);
      return { ok: r.ok === true, dispatched: true, outcome: r.outcome, error: r.error || null };
    },

    async readPackage(_handle) {
      const docUrl = state.config && state.config.document && state.config.document.url;
      if (!docUrl) return { ok: false, detail: 'no document url in the editor config' };
      const r = await fetch(docUrl);
      if (!r.ok) return { ok: false, detail: `package HTTP ${r.status}` };
      return { ok: true, bytes: Buffer.from(await r.arrayBuffer()) };
    },
  };
}

module.exports = { createEuroOfficeDocxDriver, PROPERTY_QUERY, propertyQuery, DOCX_SAVE_BOUNDS, internalGateOf, internalWriteRefusal,
  B18_BRIDGE_WRITE_RUNTIMES, B21_BRIDGE_READ_RUNTIMES, b18BridgeWritePrelude, b18DispatchResult, b21BridgeReadPrelude };
