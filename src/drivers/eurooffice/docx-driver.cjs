'use strict';

const { startConfigHost, readDsOpenState, openWaitFailureDetail } = require('./config-host.cjs');
const { observePage, settleSaveAck, dsIdentity, confirmForceSave, closeResult } = require('./save-ack.cjs');
const { defaultWait } = require('../../core/session/open.cjs');
const { ACCEPTED_RUNTIMES, DOCX_B18_BRIDGE_WRITE_RUNTIME_ENTRIES, DOCX_B18_BRIDGE_WRITE_RUNTIME_IDS,
  DOCX_B21_BRIDGE_READ_RUNTIME_ENTRIES, DOCX_B21_BRIDGE_READ_RUNTIME_IDS, DOCX_GH292S4_BRIDGE_WRITE_RUNTIME_ENTRIES, DOCX_GH292S4_BRIDGE_WRITE_RUNTIME_IDS,
  DOCX_GH293L7_BRIDGE_RUNTIME_ENTRIES, DOCX_GH293L7_BRIDGE_RUNTIME_IDS, cacheKeyOf, HANDSHAKE_BODY, classifyHandshake } = require('../../core/runtime-identity.cjs');
const { BRIDGE_CAPABILITIES, BRIDGE_CLIENT_SRC, BRIDGE_WRITES, BRIDGE_WRITE_CLIENT_SRC, BRIDGE_NUMBERING_REMOVE_WRITE_CLIENT_SRC, BRIDGE_STYLE_NEXT_LINK_WRITE_CLIENT_SRC, BRIDGE_LIST_DETACH_WRITE_CLIENT_SRC } = require('./ds-bridge.cjs');
const { resolveConfigSource } = require('./config-source.cjs');
const { IMAGES_READ_CALL } = require('./docx-images-read.cjs');

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
/** GH #292 S4: the next / linked style WRITE runs ONLY through the word /4 Bridge WRITE docx.style.nextLink.write/v1 -- the driver's
 * second line behind the style-definition family's own refusal. No direct internal path and no fallback to one exists. */
const S4_BRIDGE_WRITE_RUNTIMES = DOCX_GH292S4_BRIDGE_WRITE_RUNTIME_IDS;
/** GH #293 L7: the list members READ and the list detach WRITE run ONLY through the word /5 Bridge (docx.numbering.listMembers.read/v1,
 * docx.numbering.listDetach.write/v1) -- the driver's second line behind the numbering family's own refusal. No direct internal path. */
const L7_BRIDGE_RUNTIMES = DOCX_GH293L7_BRIDGE_RUNTIME_IDS;
/** The dependency name a command's meta.internal carries -> its build list. Empty since M11.5e.4 (B21 left it); kept for
 * any future named internal dependency and exercised by its own generic test. */
const DOCX_INTERNAL_WRITE_RUNTIMES = Object.freeze({});

/** The runtime gate of THIS session: its DS runtime identity, whether the B18 Bridge WRITE is proven on it, and whether the B21 Bridge READ is proven on it. */
function internalGateOf(runtime) {
  return Object.freeze({ runtime: runtime || null, numberingRemoveBridge: !!runtime && B18_BRIDGE_WRITE_RUNTIMES.includes(runtime),
    sectionVariant: !!runtime && B21_BRIDGE_READ_RUNTIMES.includes(runtime), styleNextLinkBridge: !!runtime && S4_BRIDGE_WRITE_RUNTIMES.includes(runtime),
    listScopeBridge: !!runtime && L7_BRIDGE_RUNTIMES.includes(runtime) });
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
/**
 * GH #292 S4: the in-editor next / link client, prepended ONLY to a modify_style command that sets one:
 *   __eoS4Gate()               the Bridge WRITE gate on the word /4 entry (the body calls it BEFORE any public setter)
 *   __eoS4Write(bridge, args)  ONE docx.style.nextLink.write/v1 write() through __eoBridgeStyleNextLink (ds-bridge.cjs: the receipt
 *                              rule). Never retried; the core re-reads.
 */
function s4BridgeWritePrelude() {
  const W = BRIDGE_WRITES.STYLE_NEXT_LINK_WRITE;
  return `${BRIDGE_CLIENT_SRC}${BRIDGE_WRITE_CLIENT_SRC}${BRIDGE_STYLE_NEXT_LINK_WRITE_CLIENT_SRC}
  const __eoS4Gate = () => __eoBridgeWriteGate(${JSON.stringify(W.id)}, ${W.version}, ${JSON.stringify(DOCX_GH292S4_BRIDGE_WRITE_RUNTIME_ENTRIES)});
  const __eoS4Write = (bridge, args) => __eoBridgeStyleNextLink(bridge, ${JSON.stringify(W.id)}, args);
`;
}
/**
 * GH #293 L7: the in-editor list detach client, prepended ONLY to a set_numbering scope 'list' command whose list shares its definition
 * (every other command stays byte-identical):
 *   __eoL7Gate()              the Bridge WRITE gate on the word /5 entry (a gate refusal writes nothing)
 *   __eoL7Detach(bridge, args) ONE docx.numbering.listDetach.write/v1 write() through __eoBridgeListDetach (ds-bridge.cjs: the receipt rule)
 */
function l7BridgeWritePrelude() {
  const W = BRIDGE_WRITES.LIST_DETACH_WRITE;
  return `${BRIDGE_CLIENT_SRC}${BRIDGE_WRITE_CLIENT_SRC}${BRIDGE_LIST_DETACH_WRITE_CLIENT_SRC}
  const __eoL7Gate = () => __eoBridgeWriteGate(${JSON.stringify(W.id)}, ${W.version}, ${JSON.stringify(DOCX_GH293L7_BRIDGE_RUNTIME_ENTRIES)});
  const __eoL7Detach = (bridge, args) => __eoBridgeListDetach(bridge, ${JSON.stringify(W.id)}, args);
`;
}

/**
 * GH #293 L7: the name-targeted list members read -- per requested item text, its level (public GetNumbering), then ONE
 * docx.numbering.listMembers.read/v1 read() gated on the word /5 entry. Per text: the capability's value, { present: false } for an item
 * that is not numbered, or { unread: <reason> } (a gate / capability / search failure -- never a guessed set). Off /5 every text is
 * { unread: 'PATCH_VERSION_MISMATCH' }: the numbering family then fails closed.
 */
function buildListMembersQuery(texts) {
  const CAP = BRIDGE_CAPABILITIES.DOCX_LIST_MEMBERS_READ;
  return `
try {
${BRIDGE_CLIENT_SRC}  const doc = Api.GetDocument();
  const out = {};
  for (const t of ${JSON.stringify(texts)}) {
    let hits = [];
    try { hits = (doc.Search(t, { MatchCase: true }) || []).filter(r => String(r.GetText({ Numbering: false }) || '') === t); } catch (e) { out[t] = { unread: 'search-threw' }; continue; }
    if (hits.length !== 1) { out[t] = { unread: hits.length ? 'target-ambiguous' : 'target-not-found' }; continue; }
    const p = hits[0].GetParagraph(0);
    const lv = p && typeof p.GetNumbering === 'function' ? p.GetNumbering() : null;
    if (!lv) { out[t] = { present: false }; continue; }
    const level = typeof lv.GetLevelIndex === 'function' ? lv.GetLevelIndex() : null;
    if (!Number.isInteger(level)) { out[t] = { unread: 'level-unreadable' }; continue; }
    const r = __eoBridgeRead(${JSON.stringify(CAP.id)}, ${CAP.version}, ${JSON.stringify(DOCX_GH293L7_BRIDGE_RUNTIME_ENTRIES)}, { text: t, level }, ${JSON.stringify(CAP.valueType)});
    out[t] = r.ok ? r.value : { unread: r.reason };
  }
  return out;
} catch (err) { return { error: String(err) }; }
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
  if (command && command.meta && command.meta.bridgeStyleNextLink === true && !S4_BRIDGE_WRITE_RUNTIMES.includes(runtime)) {
    return { ok: false, dispatched: false, outcome: 'gh292s4-bridge-write-runtime-not-proven', runtime, proven: S4_BRIDGE_WRITE_RUNTIMES };
  }
  if (command && command.meta && command.meta.bridgeListDetach === true && !L7_BRIDGE_RUNTIMES.includes(runtime)) {
    return { ok: false, dispatched: false, outcome: 'gh293l7-bridge-write-runtime-not-proven', runtime, proven: L7_BRIDGE_RUNTIMES };
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
 * GH #271 (LIVE, g271c + probes 2026-10-20, a 424-block / 30-table document on DS 9.3.5): ONE propertyQuery callCommand
 * held the editor's main thread for 41 s before a write and 175 s after it (the same body repeated in one session:
 * 43 -> 121 -> 152 -> 222 -> 282 s). No websocket frame went out meanwhile, the Document Server dropped the socket, and
 * the write's save never started (forceSave precondition: socket-closed-while-saving) -- the edit was lost, fail-closed.
 * The hot parts are the paragraph loop and the table walk; the per-call growth is the editor's own (none of Search,
 * GetRange/GetTextPr, ToJSON or the bridge reads grows on its own -- measured), so the fix is not a cheaper call but a
 * yielding one: a LARGE document is read in TIME-BUDGETED slices (paragraphs from a cursor, then the document part with
 * the tables from a cursor), each its own callCommand, so the socket's heartbeat runs between them. Every slice reports
 * the document's fingerprint (element count, text length, FNV-1a of the text); slices that disagree make the whole read
 * UNOBSERVABLE -- never a stitched read of two document states. A small document keeps the single, unchanged body.
 */
const PROPERTY_SLICE_BUDGET_MS = 4000;
const PROPERTY_SLICE_MIN_PARAGRAPH_TAGS = 300;
function sliceOf(slice) {
  if (!slice) return null;
  if ((slice.part !== 'paragraphs' && slice.part !== 'document') || !Number.isInteger(slice.from) || slice.from < 0 || !Number.isInteger(slice.budgetMs) || slice.budgetMs < 1) throw new Error('propertyQuery: invalid slice');
  return { part: slice.part, from: slice.from, budgetMs: slice.budgetMs };
}
const SLICE_PRELUDE = sl => `  const __slice = ${JSON.stringify(sl)};
  const __t0 = Date.now();
  let __next = null;
  const __fp = () => { const s = String(doc.GetText(textOpts) || ''); let h = 2166136261; for (let x = 0; x < s.length; x += 1) { h ^= s.charCodeAt(x); h = Math.imul(h, 16777619); } return doc.GetElementsCount() + ':' + s.length + ':' + (h >>> 0).toString(16); };
  const __sliceTables = (list, f) => { const acc = []; for (let k = __slice.from; k < list.length; k += 1) { if (k > __slice.from && Date.now() - __t0 > __slice.budgetMs) { __next = k; break; } acc.push(f(list[k])); } return acc; };
`;
/** The HTML projection's paragraph tags (cell paragraphs included) -- the size that decides a sliced read. */
const paragraphTagsOf = html => (String(html).match(/<p[\s>]/gi) || []).length;

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
function propertyQuery({ sectionVariantState = true, sectionVariantUnavailable = 'not-requested', numberingFormat = false, includeComments = false, slice = null } = {}) {
  const sl = sliceOf(slice);
  const unavailable = JSON.stringify(`UNAVAILABLE:${sectionVariantUnavailable}`);
  const variantFields = sectionVariantState ? '...variant(sec, index)' : `titlePage: ${unavailable}, evenAndOdd: ${unavailable}`;
  return `
try {
  const textOpts = { Numbering: false };
  const doc = Api.GetDocument();
${sl ? SLICE_PRELUDE(sl) : ''}  // A call that may throw, NAMED: the field becomes 'UNOBSERVABLE:<call>: <message>' instead of the throw aborting the WHOLE read (the outer catch returns { error } and every paragraph's rich entry, the
  // bookmark list included, is lost -- g266v1 ROUND4B). null still means 'absent', [] 'empty'; an unreadable field is this named string.
  const named = (what, f) => { try { return f(); } catch (e) { return 'UNOBSERVABLE:' + what + ': ' + String((e && e.message) || e).slice(0, 80); } };
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
  const searchFailures = [];
  const n = doc.GetElementsCount();
${sl ? `  for (let i = ${sl.part === 'paragraphs' ? sl.from : 'n'}; i < n; i += 1) {\n    if (i > __slice.from && Date.now() - __t0 > __slice.budgetMs) { __next = i; break; }\n` : '  for (let i = 0; i < n; i += 1) {\n'}    const el = doc.GetElement(i);
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
      let child;
      try { child = el.GetElement(j); } catch (_) { nonRunChildren += 1; continue; }   // a child that cannot be fetched is a non-run child: no range claim, the rest of the read goes on
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
${numberingFormat ? `    // W-07 (M7 backlog, numbering pattern/start): the OWN level's fine-grained format, through
    // the PUBLIC ApiNumbering.ToJSON() -- no internal dependency (measured live, W07-1d/W07-3
    // probes: ToJSON is present on every CreateNumbering() instance). The level index comes from
    // the bridge read above. Shape per the LIVE evidence:
    // abstractNum.<id>.lvl[ownLevel] = { numFmt: { val }, lvlText, suff, lvlJc, start, ... }
    // -- every field read on its own, absent/wrong-shape = unread, never guessed. The start
    // key is PRESENT only when explicitly set via SetStart(n); it is ABSENT on a
    // default-level numbering even though the SAVED FILE implicitly writes w:start=1 there --
    // absence was first read as UNREAD; the W07-7 measurement below settles it as 1 (start only).
    let numFormat = null;
    let numAbstract = null;
    if (numbering && numbering.present === true) {
      try {
        const lv = typeof el.GetNumbering === 'function' ? el.GetNumbering() : null;
        const num = lv && typeof lv.GetNumbering === 'function' ? lv.GetNumbering() : null;
        // LIVE run1 (a REAL bug the vm-test did not catch): ApiNumbering.ToJSON() returns a JSON
        // STRING, not an object -- the W-04 driver's tblFormatOf parses it the same way. Treating
        // the raw return as an object left the format field permanently null.
        const raw = num && typeof num.ToJSON === 'function' ? num.ToJSON() : null;
        const tj = typeof raw === 'string' ? JSON.parse(raw) : raw;
        const abs = tj && tj.abstractNum && typeof tj.abstractNum === 'object' ? Object.values(tj.abstractNum)[0] : null;
        // GH #293 N5 (LIVE, docs/evidence/gh293-n5-word-shared-abstractnum-public-recognition-20261022.json): the abstractNum key IS the
        // definition's identity -- two lists with their own numId on ONE abstractNum (Word's shape) answer the same key, an
        // independent list another one. A missing / several keys = unread (null).
        const absKeys = tj && tj.abstractNum && typeof tj.abstractNum === 'object' ? Object.keys(tj.abstractNum) : [];
        numAbstract = absKeys.length === 1 ? String(absKeys[0]) : null;
        const ownLvl = Number.isInteger(numbering.level) ? numbering.level : 0;
        const lvlObj = abs && Array.isArray(abs.lvl) ? abs.lvl[ownLvl] : null;
        if (lvlObj && typeof lvlObj === 'object') {
          const f = {};
          if (lvlObj.numFmt && typeof lvlObj.numFmt === 'object' && lvlObj.numFmt.val !== undefined) f.numFmt = String(lvlObj.numFmt.val);
          if (typeof lvlObj.lvlText === 'string') f.lvlText = lvlObj.lvlText;
          if (typeof lvlObj.suff === 'string') f.suff = lvlObj.suff;
          if (typeof lvlObj.lvlJc === 'string') f.lvlJc = lvlObj.lvlJc;
          // W07-7 probe (evidence docx-w07-7-start1-20261007.json): ToJSON represents start=1 by
          // OMITTING the key -- measured in BOTH directions (a level explicitly moved 5 -> 1 loses
          // the key; a level NEVER set and a level explicitly set to 1 are IDENTICAL, key-absent,
          // and the saved file agrees: w:start=1 in every one of those cases). Absence is therefore
          // read as 1 -- but ONLY for this one key; a present, non-integer start stays unread, and
          // every OTHER missing field (numFmt/lvlText/suff/lvlJc) still means "unread".
          if (Number.isInteger(lvlObj.start)) f.start = lvlObj.start;
          else if (!('start' in lvlObj)) f.start = 1;
          numFormat = f;
        }
      } catch (_) { numFormat = null; }
    }
` : ''}    // Wave 3 (rich-inline + notes), needle-free, measured live 2026-09-24 (rich probes 1-3):
    // a hyperlink is a child of ClassType 'hyperlink' (url, screen tip, displayed text, effective
    // format of its range); a note reference is an empty run whose DIRECT style is 'footnote
    // reference' / 'endnote reference'. The GetText above excludes both markers.
    const hyperlinks = [];
    const noteRefs = { footnote: 0, endnote: 0 };
    let childFailure = null;   // ONE child that cannot be read names itself for THIS paragraph; it never aborts the other paragraphs
    for (let j = 0; j < k; j += 1) {
      try {
        const child = el.GetElement(j);
        const cls = child && typeof child.GetClassType === 'function' ? child.GetClassType() : null;
        if (cls === 'hyperlink') {
          hyperlinks.push({ url: named('GetLinkedText', () => String(child.GetLinkedText() || '')), tip: named('GetScreenTipText', () => String(child.GetScreenTipText() || '')), text: named('GetDisplayedText', () => String(child.GetDisplayedText() || '')),
            style: named('GetRange/GetTextPr/GetStyle', () => { const hr = typeof child.GetRange === 'function' ? child.GetRange() : null; const hp = hr && typeof hr.GetTextPr === 'function' ? hr.GetTextPr() : null; const style = hp && typeof hp.GetStyle === 'function' ? hp.GetStyle() : null; return style && typeof style.GetName === 'function' ? String(style.GetName()) : null; }) });
        }
        if (cls === 'run' && typeof child.GetTextPr === 'function') {
          const dp = child.GetTextPr();
          const st = dp && typeof dp.GetStyle === 'function' ? dp.GetStyle() : null;
          const sn = st && typeof st.GetName === 'function' ? String(st.GetName()) : '';
          if (sn === 'footnote reference') noteRefs.footnote += 1;
          if (sn === 'endnote reference') noteRefs.endnote += 1;
        }
      } catch (e) { if (childFailure === null) childFailure = 'UNOBSERVABLE:child ' + j + ': ' + String((e && e.message) || e).slice(0, 80); }
    }
    const runInfo = { runs, nonRunChildren, runsCoverText: runs.map(r => r.text).join('') === text, hyperlinks: childFailure !== null ? childFailure : hyperlinks, noteRefs: childFailure !== null ? childFailure : noteRefs, ...(numbering ? { numbering${numberingFormat ? ': numbering.present === true ? { ...numbering, format: numFormat, abstractId: numAbstract } : numbering' : ''} } : {}) };
    // The lookup is its own failure domain (gbolv2/gbolv3 LIVE): doc.Search THROWS "this.Content[nPos].ConvertParaContentPosToRangePos is not a
    // function" while converting a hit INSIDE a paragraph whose hyperlink holds a bookmark (S3: bookmark inside the link text, S5: exactly on the
    // link range -- the stored file carries the link and the bookmark in both). It is a throw, not a non-unique hit. ONLY that paragraph is
    // affected; every other paragraph and the bookmark list stay readable. For that paragraph the SAME common text properties are read from the
    // paragraph's own range (p.GetRange().GetTextPr(), the call the table-cell paragraph reads already use) -- NO Search; if that fails too, the
    // paragraph names the Search throw as before. A value read this way carries propsVia: 'paragraph-range'.
    let hits;
    let searchError = null;
    let p = null;
    let propsVia = null;
    try { hits = (doc.Search(text, { MatchCase: true }) || []).filter(r => String(r.GetText(textOpts) || '') === text); } catch (e) { searchError = String((e && e.message) || e).slice(0, 120); }
    if (searchError !== null) {
      try { const pr0 = typeof el.GetRange === 'function' ? el.GetRange() : null; p = pr0 && typeof pr0.GetTextPr === 'function' ? pr0.GetTextPr() : null; } catch (_) { p = null; }
      if (!p) { searchFailures.push({ index: i, error: searchError }); out.push({ text, props: null, searchUnobservable: 'UNOBSERVABLE:search-throws: ' + searchError, ...runInfo }); continue; }
      propsVia = 'paragraph-range';
    } else {
      if (hits.length !== 1) { out.push({ text, props: null, ...runInfo }); continue; }
      p = hits[0].GetTextPr();
      if (!p) { out.push({ text, props: null, ...runInfo }); continue; }
    }
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
      // GH #231b (docs/GH231-FORMAT-TEXT-PARAGRAPH-PROPS-DECISION.md section 8): the ORIGINAL formula here
      // (ported offline from the source oracle, 5d9caa1d, never independently LIVE-verified by this team)
      // was n>10 ? (n>=200 ? n/240 : (n+1)/100) : n -- NOT a clean inverse of this family's own write,
      // SetSpacingLine(wanted.lineSpacing * 240, 'auto') (formatting.cjs:543/623): for any written multiplier
      // below 200/240≈0.833 the old formula read back a WRONG value (write 0.5 -> old formula gives 1.21).
      // Every write in THIS codebase always uses the 'auto' rule, whose OOXML unit is UNCONDITIONALLY
      // 240ths-of-a-line (ECMA-376 w:spacing/@w:line under lineRule='auto') -- so the single, uniform
      // inverse n/240 is the correct conversion for every value this family can ever write. The
      // n>=200/(n+1)/100 branch is PENDING LIVE CONFIRMATION to have been a genuine bug (not a real,
      // alternate OOXML convention this family also produces) -- test/live/gh231b-linespacing-read-roundtrip-probe.live.cjs
      // (not run by this agent) measures the RAW GetSpacingLineValue()/GetSpacingLineRule() returns across
      // a value grid to settle this before the fix lands LIVE.
      if (typeof pp.GetSpacingLineValue === 'function') { const n = Number(pp.GetSpacingLineValue()); if (Number.isFinite(n)) format.lineSpacing = n / 240; }
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
          // GH #290 (docs/evidence/gh290-paragraph-format-public-api-20261021.json): the boolean keys above are the DIRECT
          // value with an absent flag read as false -- but a paragraph DOES inherit these from its style (Heading 1 keepNext /
          // keepLines, measured). paginationDirect keeps the distinction: true | false (explicit w:val) | null (absent).
          const dflag = v => (typeof v === 'boolean' ? v : null);
          format.paginationDirect = { keepNext: dflag(ptj.keepNext), keepLines: dflag(ptj.keepLines), pageBreakBefore: dflag(ptj.pageBreakBefore), widowControl: dflag(ptj.widowControl) };
        }
      }
    }
    out.push({ text, props, format, ...(propsVia ? { propsVia } : {}), ...runInfo });
  }
${sl && sl.part === 'paragraphs' ? '  return { paragraphs: out, searchFailures, next: __next, fp: __fp() };\n' : ''}  // Wave 3, document level: bookmarks (name + the text of its range), the first paragraph of every
  // footnote/endnote, and whether the editor's cursor is in the MAIN body (GetCurrentFootEndnote()
  // null) -- after a note write it is not, until the note command restores it (rich probes 2-3).
  // Each field is read on its own; an unreadable one is absent (unknown), never an empty list.
  const docInfo = {};
  const noteText = p => String(p.GetText(textOpts) || '').replace(/^\\s+|[\\r\\n]+$/g, '');
  // per bookmark: ONE bookmark whose range cannot be read names ITS OWN call in ITS text, the other bookmarks stay readable (the whole list is absent only if the NAMES cannot be listed)
  try { docInfo.bookmarks = (doc.GetAllBookmarksNames() || []).map(name => { const r = named('GetBookmarkRange', () => doc.GetBookmarkRange(name)); return { name: String(name), text: typeof r === 'string' ? r : (r ? named('GetText', () => String(r.GetText(textOpts) || '')) : null) }; }); } catch (_) { /* unread */ }
  try { docInfo.footnotes = (doc.GetFootnotesFirstParagraphs() || []).map(noteText); } catch (_) { /* unread */ }
  try { docInfo.endnotes = (doc.GetEndNotesFirstParagraphs() || []).map(noteText); } catch (_) { /* unread */ }
  // GH #226 3 (additive; the first-paragraph fields above are UNCHANGED for every old consumer): the FULL paragraph list of each note through the
  // PUBLIC GetNext() chain of its first paragraph (GH #226 2 LIVE: FULL on DS 9.3.5, read-only, no selection change). The chain is TRUSTED only when it
  // ends by GetNext() === null within NOTE_CHAIN_MAX steps and never reaches the first paragraph of ANOTHER note (a text match used ONLY as an overrun
  // DETECTOR that marks the note incomplete -- never as a boundary to cut at); anything else is NAMED: { paragraphs: the ones read, complete: false,
  // stop: 'no-getnext' | 'error' | 'overrun-suspected' | 'bound' }, never a guessed full text. A chain that is unreadable as a whole is absent (unknown).
  const NOTE_CHAIN_MAX = 40;
  const noteChains = firsts => {
    const firstTexts = firsts.map(noteText);
    return firsts.map((first, i) => {
      const paragraphs = [firstTexts[i]];
      if (!first || typeof first.GetNext !== 'function') return { paragraphs, complete: false, stop: 'no-getnext' };
      let p = first;
      for (let n = 0; n < NOTE_CHAIN_MAX; n += 1) {
        let next = null;
        try { next = p.GetNext(); } catch (_) { return { paragraphs, complete: false, stop: 'error' }; }
        if (!next) return { paragraphs, complete: true, stop: 'null' };
        const t = noteText(next);
        if (t !== '' && firstTexts.some((ft, j) => j !== i && ft === t)) return { paragraphs, complete: false, stop: 'overrun-suspected' };
        paragraphs.push(t);
        p = next;
      }
      return { paragraphs, complete: false, stop: 'bound' };
    });
  };
  try { docInfo.footnoteChains = noteChains(doc.GetFootnotesFirstParagraphs() || []); } catch (_) { /* unread */ }
  try { docInfo.endnoteChains = noteChains(doc.GetEndNotesFirstParagraphs() || []); } catch (_) { /* unread */ }
${includeComments ? `  // GH #269b (opt-in: only a read_comments request pays for it; READ-only -- no Set*/Add*/Delete/Push, no selection change): every top-level comment of the MAIN
  // document through the PUBLIC doc.GetAllComments(), per field in its own guard (a failed field is NAMED 'UNOBSERVABLE:<reason>', a missing getter likewise,
  // never a guessed value); the replies through GetRepliesCount()/GetReply(i). The range anchor has NO public getter (docs/GH269-DOCX-COMMENTS-READ-WRITE-RESEARCH.md
  // 2.) -- the quote is the text at CREATION time, not a live anchor. A header/footer/note comment is not enumerated (the main document's CommentsManager only).
  try {
    const unob = e => 'UNOBSERVABLE:' + String((e && e.message) || e).slice(0, 80);
    const pick = (o, k, f) => { try { o[k] = f(); } catch (e) { o[k] = unob(e); } };
    docInfo.comments = (doc.GetAllComments() || []).map(c => {
      const row = {};
      pick(row, 'id', () => c.GetId()); pick(row, 'text', () => c.GetText()); pick(row, 'author', () => c.GetAuthorName()); pick(row, 'userId', () => c.GetUserId());
      pick(row, 'timeUtc', () => c.GetTimeUTC()); pick(row, 'quote', () => c.GetQuoteText()); pick(row, 'solved', () => c.IsSolved());
      row.replies = [];
      try {
        const n = c.GetRepliesCount();
        for (let k = 0; k < n; k += 1) { const r = c.GetReply(k); const rr = {}; pick(rr, 'text', () => r.GetText()); pick(rr, 'author', () => r.GetAuthorName()); pick(rr, 'userId', () => r.GetUserId()); row.replies.push(rr); }
      } catch (e) { row.replies = unob(e); }
      return row;
    });
  } catch (_) { /* unread: GetAllComments is absent or threw -- the field stays absent (unknown), not an empty list */ }
` : ''}  try { docInfo.inMainBody = doc.GetCurrentFootEndnote() === null; } catch (_) { /* unread */ }
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
  // GH #289 (LIVE probes 2-3, docs/evidence/gh289-*-hidden-content-20261021.json): a PAGE field reads as a run with EMPTY
  // text and an image's run carries no text either, so the run classes alone call a paragraph "plain" -- a full
  // replacement (header/footer, table cell) dropped the field / the image silently. The paragraph's public ToJSON()
  // names them (pgNum, paraDrawing, image ...): every content type outside the plain set, sorted; a throwing /
  // unparsable read -> ['UNREAD'] (the caller refuses).
  const PLAIN_PARAGRAPH_TYPES = ['paragraph', 'paraPr', 'endRun', 'run', 'textPr'];
  function inlineTypesOf(e) {
    try {
      const raw = e.ToJSON(false, false);
      const o = typeof raw === 'string' ? JSON.parse(raw) : raw;
      if (!o || typeof o !== 'object') return ['UNREAD'];
      const seen = new Set();
      const walk = (v, depth) => {
        if (!v || typeof v !== 'object') return;
        if (depth > 12) { seen.add('TOO_DEEP'); return; }
        if (typeof v.type === 'string' && !PLAIN_PARAGRAPH_TYPES.includes(v.type)) seen.add(v.type);
        for (const k of Object.keys(v)) walk(v[k], depth + 1);
      };
      walk(o, 0);
      return [...seen].sort();
    } catch (_) { return ['UNREAD']; }
  }
  // W04-6 (the LIVE catch): these three consts + tblFormatOf() MUST be declared BEFORE
  // the docInfo.tables = (...).map(...) call below -- that map runs SYNCHRONOUSLY and calls
  // tblFormatOf(t) immediately, so if this block is declared AFTER (as W04-5 had it), the consts
  // are still in the TEMPORAL DEAD ZONE at call time: a ReferenceError, caught by the outer
  // try/catch, silently discards docInfo.tables ENTIRELY -- breaking every table family's read
  // (table, table-row, table-column, merge-split, format_table alike), not just table-format's own
  // fields. LIVE run1 caught this as 'tables-unreadable' on every round; fail-closed, no write, no
  // corruption -- but a landed version of this bug would have broken every DOCX table operation.
  //
  // W04-4 (MEASURED): per table, the table-level tblPr fields through the public
  // ToJSON() -- a JSON STRING, parsed once (never relied on as an object directly; the W04-4
  // probe's own style-object read confirmed ToJSON() returns a string there too). Each of the
  // 6 fields fails UNOBSERVABLE independently (never one failure dragging another field down,
  // same discipline as read1 above): a missing ToJSON, an unparsable string, or a missing tblPr
  // leaves ALL 6 unobservable (there is nothing to extract from); a missing SUB-key (e.g. no
  // border set on an edge, no cellMargin on an edge, no tblStyle) is a successfully-read ABSENT
  // value ('null'), not UNOBSERVABLE -- the read worked, it just found nothing there.
  //
  // Border key-mapping (MEASURED): ToJSON's tblBorders renames ONLY left/right to
  // start/end; top/bottom/insideH/insideV keep their name. The SAVED XML, independently, keeps
  // w:left/w:right -- the rename exists ONLY in this ToJSON() read, nowhere else. cellMargins
  // has NO such rename (left/right keep their name in ToJSON too, confirmed).
  const BORDER_EDGES = ['top', 'bottom', 'left', 'right', 'insideH', 'insideV'];
  const BORDER_JSON_KEY = { left: 'start', right: 'end' };
  const MARGIN_EDGES = ['top', 'right', 'bottom', 'left'];
  function tblFormatOf(t) {
    let tblPr;
    try {
      if (typeof t.ToJSON !== 'function') return { borders: 'UNOBSERVABLE:no-such-getter', shading: 'UNOBSERVABLE:no-such-getter', cellMargins: 'UNOBSERVABLE:no-such-getter', indent: 'UNOBSERVABLE:no-such-getter', width: 'UNOBSERVABLE:no-such-getter', style: 'UNOBSERVABLE:no-such-getter' };
      const raw = t.ToJSON();
      const obj = typeof raw === 'string' ? JSON.parse(raw) : raw;
      tblPr = obj && obj.tblPr;
      if (!tblPr || typeof tblPr !== 'object') return { borders: 'UNOBSERVABLE:no-tblPr', shading: 'UNOBSERVABLE:no-tblPr', cellMargins: 'UNOBSERVABLE:no-tblPr', indent: 'UNOBSERVABLE:no-tblPr', width: 'UNOBSERVABLE:no-tblPr', style: 'UNOBSERVABLE:no-tblPr' };
    } catch (e) {
      const reason = 'UNOBSERVABLE:' + String(e && e.message || e).slice(0, 80);
      return { borders: reason, shading: reason, cellMargins: reason, indent: reason, width: reason, style: reason };
    }
    const borders = {};
    for (const edge of BORDER_EDGES) {
      try {
        const key = BORDER_JSON_KEY[edge] || edge;
        const b = tblPr.tblBorders && tblPr.tblBorders[key];
        borders[edge] = b && typeof b === 'object' ? { value: b.value, sz: b.sz, space: b.space, color: b.color && typeof b.color === 'object' ? { auto: b.color.auto, r: b.color.r, g: b.color.g, b: b.color.b } : null } : null;
      } catch (e) { borders[edge] = 'UNOBSERVABLE:' + String(e && e.message || e).slice(0, 80); }
    }
    let shading;
    try { const s = tblPr.shd; shading = s && typeof s === 'object' ? { val: s.val, fill: s.fill, color: s.color } : null; } catch (e) { shading = 'UNOBSERVABLE:' + String(e && e.message || e).slice(0, 80); }
    const cellMargins = {};
    for (const edge of MARGIN_EDGES) {
      try { const m = tblPr.tblCellMar && tblPr.tblCellMar[edge]; cellMargins[edge] = m && typeof m === 'object' ? { type: m.type, w: m.w } : null; } catch (e) { cellMargins[edge] = 'UNOBSERVABLE:' + String(e && e.message || e).slice(0, 80); }
    }
    let indent;
    try { indent = typeof tblPr.tblInd === 'number' && Number.isFinite(tblPr.tblInd) ? tblPr.tblInd : (tblPr.tblInd === undefined ? null : 'UNOBSERVABLE:non-numeric-read'); } catch (e) { indent = 'UNOBSERVABLE:' + String(e && e.message || e).slice(0, 80); }
    let width;
    try { const w = tblPr.tblW; width = w && typeof w === 'object' ? { type: w.type, w: w.w } : null; } catch (e) { width = 'UNOBSERVABLE:' + String(e && e.message || e).slice(0, 80); }
    let style;
    try { style = tblPr.tblStyle === undefined ? null : tblPr.tblStyle; } catch (e) { style = 'UNOBSERVABLE:' + String(e && e.message || e).slice(0, 80); }
    return { borders, shading, cellMargins, indent, width, style };
  }
  // The tables list is the document TOP-LEVEL tables in document order, found by the element walk (gntv1 LIVE: the walk finds them and the cell holding a nested table is
  // still there). GetAllTables() also lists the nested tables, so its length differed from the HTML projection (top-level only) and every table family refused with
  // table-read-shape-mismatch. all is a table GetAllTables() position: the bridge grid read keeps that index space (top-level index + the nested tables seen so far).
  // A throwing element is NAMED (tablesUnobservable) and the list stays absent -- never a silently narrowed list.
  // One table cell as the product reads it; depth 0 = a top-level table, a nested table is read one level deeper (at most NESTED_READ_DEPTH levels).
  const NESTED_READ_DEPTH = 3;
  function readCell(cellObj, depth) {
      const content = cellObj.GetContent();
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
      // GH #289: the non-plain content types of the cell's paragraphs (a field, an image ...), the union, sorted
      const inlineSet = new Set();
      for (let i = 0; i < m; i += 1) { const e = content.GetElement(i); if (e && e.GetClassType() === 'paragraph') for (const t of inlineTypesOf(e)) inlineSet.add(t); }
      const info = { text: String(content.GetText(textOpts) || ''), paragraphs: m, simple, format, paraFormats, inline: [...inlineSet].sort() };
      // nested tables of this cell (read_tables returns them bound to the holding cell); absent when the cell holds none. A throw here names itself on THIS cell only.
      try {
        const nested = [];
        for (let i = 0; i < m; i += 1) { const e = content.GetElement(i); if (e && e.GetClassType() === 'table') nested.push(e); }
        if (nested.length > 0) {
          const read = depth >= NESTED_READ_DEPTH ? 'UNOBSERVABLE:nested-depth-limit' : nested.map(nt => readTableRows(nt, depth + 1));
          // gnpq2 LIVE (DS 9.3.5): a top-level table carrying its nested tables as OBJECTS does not come back through the callCommand callback (undefined, the whole
          // query lost); the same tables without them and the nested part alone both do. So the holding cell of a TOP-LEVEL table hands its nested tables over as a JSON
          // string (nestedTablesJson); readLive (richOutcome -> inflateNestedTables) turns it back into cell.nestedTables. Deeper levels sit inside that string.
          if (depth === 0 && Array.isArray(read)) info.nestedTablesJson = JSON.stringify(read); else info.nestedTables = read;
        }
      } catch (e) { info.nestedTables = 'UNOBSERVABLE:nested-table: ' + String((e && e.message) || e).slice(0, 80); }
      return info;

  }
  function readTableRows(tb, depth) {
    const nrows = [];
    for (let r = 0; r < tb.GetRowsCount(); r += 1) {
      const row = tb.GetRow(r);
      const ncells = [];
      for (let c = 0; c < row.GetCellsCount(); c += 1) ncells.push(readCell(row.GetCell(c), depth));
      nrows.push(ncells);
    }
    return { rows: nrows };
  }
  const walkedTables = [];
  let tablesWalkFailure = null;
  const countNestedIn = tb => {
    let k = 0;
    for (let r = 0; r < tb.GetRowsCount(); r += 1) {
      const row = tb.GetRow(r);
      for (let c = 0; c < row.GetCellsCount(); c += 1) {
        const content = row.GetCell(c).GetContent();
        for (let e = 0; e < content.GetElementsCount(); e += 1) { const x = content.GetElement(e); if (x && x.GetClassType() === 'table') k += 1 + countNestedIn(x); }
      }
    }
    return k;
  };
  try {
    let nestedSeen = 0;
    const bodyCount = doc.GetElementsCount();
    for (let i = 0; i < bodyCount; i += 1) {
      let el;
      try {
        el = doc.GetElement(i);
        if (el && el.GetClassType() === 'table') { walkedTables.push({ t: el, all: walkedTables.length + nestedSeen }); nestedSeen += countNestedIn(el); }
      } catch (e) { tablesWalkFailure = 'UNOBSERVABLE:table-walk: element ' + i + ': ' + String((e && e.message) || e).slice(0, 80); break; }
    }
  } catch (e) { tablesWalkFailure = 'UNOBSERVABLE:table-walk: ' + String((e && e.message) || e).slice(0, 80); }
  if (tablesWalkFailure !== null) docInfo.tablesUnobservable = tablesWalkFailure;
  try {
    if (tablesWalkFailure === null) docInfo.tables = ${sl ? '__sliceTables(walkedTables, w => {' : 'walkedTables.map(w => {'}
      const t = w.t;
      const rows = [];
      for (let r = 0; r < t.GetRowsCount(); r += 1) {
        const row = t.GetRow(r);
        const cells = [];
        for (let c = 0; c < row.GetCellsCount(); c += 1) cells.push(readCell(row.GetCell(c), 0));
        rows.push(cells);
      }
      // M5 gate 8 -> W-B1920-C1 cutover: the CALCULATED grid column widths through
      // docx.table.effectiveGrid.read/v1 ONLY -- no internal calculated-grid fallback. null =
      // unread (bridge failure or an invalid shape); the split family then refuses a column split
      // before dispatch, exactly as it did for an unverified-build internal read.
      const gridRead = rd(${JSON.stringify(BRIDGE_CAPABILITIES.DOCX_TABLE_GRID_READ)}, { table: w.all });
      const gridCalc = gridRead && Array.isArray(gridRead.gridCalc) ? gridRead.gridCalc : null;
      return { rows, gridCalc, format: tblFormatOf(t) };
    });
  } catch (_) { /* unread */ }
${sl ? '  if (__next !== null) return { paragraphs: out, document: { tables: docInfo.tables }, searchFailures, next: __next, fp: __fp() };\n' : ''}  // W04-4 (MEASURED): the table-style NAME-to-id mapping a 'style' request needs BEFORE
  // dispatch (classify()/resolve() see only this passively-read list, never calling GetAllStyles()
  // themselves). Selection discipline matches the probe exactly: GetType()==='table' only, and the
  // id is the style object's OWN ToJSON().styleId (a bare number, measured) -- never re-derived or
  // guessed. A style missing a name or whose ToJSON does not parse is skipped, not defaulted.
  try {
    docInfo.tableStyles = (doc.GetAllStyles() || []).filter(s => { try { return typeof s.GetType === 'function' && s.GetType() === 'table'; } catch (_) { return false; } }).map(s => {
      let name = null; let styleId = null;
      try { name = typeof s.GetName === 'function' ? String(s.GetName()) : null; } catch (_) { name = null; }
      try { const raw = typeof s.ToJSON === 'function' ? s.ToJSON() : null; const obj = typeof raw === 'string' ? JSON.parse(raw) : raw; styleId = obj && obj.styleId !== undefined ? obj.styleId : null; } catch (_) { styleId = null; }
      return { name, styleId };
    }).filter(s => s.name !== null && s.styleId !== null);
  } catch (_) { docInfo.tableStyles = null; }
  // GH #268b: the images -- doc.GetAllImages() + the page-side projection of ApiDrawing.ToJSON() (docx-images-read.cjs); the image BYTES ToJSON carries never leave the editor.
  ${IMAGES_READ_CALL}
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
      // GH #289 H-05: a part of paragraphs of runs AND tables -- each table's cell texts, in element order (the paragraph
      // mode leaves them alone and proves it); anything else makes withTables false
      let runsOnly = true;
      let other = false;
      const tables = [];
      const paras = [];
      for (let i = 0; i < m; i += 1) {
        const e = hf.GetElement(i);
        if (e && e.GetClassType() === 'table') {
          simple = false;
          try {
            const rows = [];
            for (let r = 0; r < e.GetRowsCount(); r += 1) { const row = e.GetRow(r); const cs = []; for (let c = 0; c < row.GetCellsCount(); c += 1) cs.push(String(row.GetCell(c).GetContent().GetText(textOpts) || '')); rows.push(cs); }
            tables.push({ at: i, rows });
          } catch (_) { other = true; }
          continue;
        }
        if (!e || e.GetClassType() !== 'paragraph') { simple = false; other = true; continue; }
        for (let j = 0; j < e.GetElementsCount(); j += 1) { const ch = e.GetElement(j); if (!ch || ch.GetClassType() !== 'run') { simple = false; runsOnly = false; } }
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
        paras.push({ text: String(e.GetText(textOpts) || ''), format: paragraphFormat(e), runs, inline: inlineTypesOf(e) });
      }
      return { paragraphs: m, simple, paras, tables, withTables: tables.length > 0 && runsOnly && !other };
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
    // W11: page-setup, PUBLIC getters only, confirmed live (docx-w11-1a-section-
    // reflection-probe.live.cjs / docx-w11-1b-section-write-probe.live.cjs). EACH getter
    // in its own try, same pattern as flag() above: a missing or throwing getter sets ONLY that
    // field to 'UNOBSERVABLE:<reason>' -- never a default number -- and never drags any other
    // field down with it.
    const read1 = (obj, name) => {
      try {
        if (!obj || typeof obj[name] !== 'function') return 'UNOBSERVABLE:no-such-getter';
        const v = obj[name]();
        return typeof v === 'number' && Number.isFinite(v) ? v : 'UNOBSERVABLE:non-numeric-read';
      } catch (e) { return 'UNOBSERVABLE:' + String(e && e.message || e).slice(0, 80); }
    };
    // W11-4: GetColumnsSpaces() IS independently re-readable (confirmed live, session +
    // two reopens, docs/evidence/docx-w11-1b-section-write-20261007.json R2/R3) -- the earlier
    // belief that only GetColumnsWidths (a DERIVED value) existed was wrong; GetColumnsCount's
    // sibling array getter was simply never checked. Same per-field discipline as read1: a
    // non-array, or any non-numeric element, is UNOBSERVABLE, never a guessed array.
    const readArr = (obj, name) => {
      try {
        if (!obj || typeof obj[name] !== 'function') return 'UNOBSERVABLE:no-such-getter';
        const v = obj[name]();
        if (!Array.isArray(v) || !v.every(x => typeof x === 'number' && Number.isFinite(x))) return 'UNOBSERVABLE:non-numeric-array-read';
        return v.slice();
      } catch (e) { return 'UNOBSERVABLE:' + String(e && e.message || e).slice(0, 80); }
    };
    const pageSetupOf = sec => ({
      pageWidth: read1(sec, 'GetPageWidth'),
      pageHeight: read1(sec, 'GetPageHeight'),
      margins: {
        top: read1(sec, 'GetPageMarginTop'), right: read1(sec, 'GetPageMarginRight'),
        bottom: read1(sec, 'GetPageMarginBottom'), left: read1(sec, 'GetPageMarginLeft'),
      },
      columnsCount: read1(sec, 'GetColumnsCount'),
      columnsSpaces: readArr(sec, 'GetColumnsSpaces'),
      startPageNumber: read1(sec, 'GetStartPageNumber'),
    });
    docInfo.sections = (doc.GetSections() || []).map((sec, index) => {
      const one = { index, header: {}, footer: {}, ${variantFields}, ...pageSetupOf(sec) };
      for (const v of ['default', 'even', 'title']) { one.header[v] = partOf(sec.GetHeader(v, false)); one.footer[v] = partOf(sec.GetFooter(v, false)); }
      return one;
    });
  } catch (_) { /* unread */ }
  // W-10 (style DEFINITIONS, CORRECTED after LIVE run1 showed every round
  // verify-uncertain): styles are identified by NAME, never styleId -- ToJSON() does NOT carry
  // a styleId field (measured live, W10-1c: "a ToJSON-ban styleId NINCS, a fajlban van");
  // doc.GetStyle(name) is the session-side identity. ToJSON itself can be a STRING (the W07-6
  // lesson) -- typeof-guarded before JSON.parse, never assumed to be an object.
  //
  // The FIRST version of this block read the style's own GetTextPr()/GetParaPr() GETTERS
  // directly -- that surface was NEVER independently measured (only the matching SETTERS were,
  // W10-1b/1c/W10-2, write-only) and LIVE run1 confirmed it is NOT readable: every
  // round came back verify-uncertain, because every getter degraded to UNOBSERVABLE. The review named
  // the ALREADY-MEASURED fix: the style's ToJSON() itself carries readable rPr/pPr sub-fields,
  // confirmed live AFTER A REOPEN (W10-1c evidence R2, docs/evidence/
  // docx-w10-1c-style-write-run2-20261007.json, + the W10-2 SetColor/SetSpacingAfter evidence):
  //   rPr.b = true | rPr.sz = <half-points, the SAME unit as SetFontSize's own argument>
  //   rPr.color = { auto, r, g, b }  (the Builder's RGB(255,0,0) call -> r:255 g:0 b:0)
  //   pPr.jc = 'center' | pPr.ind.left = <twips> | pPr.spacing.after = <twips>
  //     (pPr.spacing.afterAutoSpacing appears ONLY when SetSpacingAfter's 2nd arg was true --
  //     v1 always passes false, so this key is not read here)
  // CORRECTED (LIVE run2): rPr.sz is in the SAME unit as SetFontSize's own
  // WRITE argument (half-points) -- but the family's task-facing size field (inside 'text')
  // is PT (the SAME convention indentLeft/spacingAfter already externalize at, and the one
  // lower() converts FROM when writing, size times 2). The read side must convert BACK
  // (divide by 2), the same way indentLeft/spacingAfter already divide by 20 (twips -> pt)
  // below -- the earlier comment here ('SAME unit, no conversion') was WRONG: that is true
  // for the WRITE argument, not for round-tripping the read back into the task-facing pt
  // field. LIVE run2 measured the exact mismatch: a requested size of 26pt wrote rPr.sz=52
  // (half-pt, correct), but the UN-halved read compared 52 against the requested 26 and
  // never matched -- every define_style/modify_style round with a size field was uncertain
  // for this reason alone.
  // A key ABSENT from rPr/pPr is the style's OWN value NOT explicitly set (inherited from its
  // base style) -- named 'inherited', never guessed as a default or treated as UNOBSERVABLE
  // (UNOBSERVABLE stays reserved for a genuine read failure: no ToJSON method, a throw, or a
  // parse failure -- cases where inheritance cannot even be distinguished from a read error).
  try {
    const styleIdOf = {};
    const hex2 = n => Number(n).toString(16).padStart(2, '0').toUpperCase();
    const colorOf = (c) => {
      if (!c || typeof c !== 'object') return 'UNOBSERVABLE:non-object-color';
      if (c.auto === true) return 'auto';
      if (Number.isInteger(c.r) && Number.isInteger(c.g) && Number.isInteger(c.b)) return hex2(c.r) + hex2(c.g) + hex2(c.b);
      return 'UNOBSERVABLE:unrecognized-color-shape';
    };
    const styleOf = (name) => {
      // W-10 (LIVE run4): doc.GetStyle(<unknown name>) was MEASURED to NOT return
      // null -- it falls back to the 'Normal' style object. A match is only valid when the
      // returned style's own GetName() equals the REQUESTED name exactly; otherwise treat it as
      // not-found. The SAME guard applies in buildTargetedStyleQuery's verbatim copy below.
      const raw = typeof doc.GetStyle === 'function' ? doc.GetStyle(name) : null;
      const st = raw && typeof raw.GetName === 'function' && raw.GetName() === name ? raw : null;
      if (!st) return null;
      let tj = null;
      let tjReadFailed = false;
      try {
        const rawTj = typeof st.ToJSON === 'function' ? st.ToJSON() : null;
        tj = typeof rawTj === 'string' ? JSON.parse(rawTj) : rawTj;
        if (!tj || typeof tj !== 'object') tjReadFailed = true;
      } catch (_) { tj = null; tjReadFailed = true; }
      if (tjReadFailed) {
        const u = 'UNOBSERVABLE:no-readable-ToJSON';
        return { name, identity: null, text: { bold: u, size: u, color: u, italic: u, underline: u, strike: u, fontFamily: u }, paragraph: { align: u, indentLeft: u, spacingAfter: u, spacingBefore: u, indentRight: u, indentFirstLine: u, keepNext: u, keepLines: u, pageBreakBefore: u } };
      }
      const identity = { name: typeof tj.name === 'string' ? tj.name : null,
        basedOn: tj.basedOn === undefined ? null : tj.basedOn,
        // GH #292 S4 (LIVE, docs/evidence/gh292s4-next-link-internal-probe-20261022.json): next / link are styleIds in the same ToJSON, resolved to names below
        next: tj.next === undefined ? null : tj.next, link: tj.link === undefined ? null : tj.link,
        styleType: typeof tj.styleType === 'string' ? tj.styleType : null };
      // GH #292 S3 (LIVE, docs/evidence/gh292-style-definition-probe-20261021.json): basedOn is the base style's ToJSON styleId, never its name --
      // the name is resolved after the enumeration below (basedOnName), from the styleIds this read itself saw
      if (tj.styleId !== undefined && tj.styleId !== null) styleIdOf[name] = String(tj.styleId);
      const rPr = tj.rPr && typeof tj.rPr === 'object' ? tj.rPr : null;
      const pPr = tj.pPr && typeof tj.pPr === 'object' ? tj.pPr : null;
      const bold = !rPr ? 'inherited' : (!('b' in rPr) ? 'inherited' : !!rPr.b);
      const size = !rPr ? 'inherited' : (!('sz' in rPr) ? 'inherited' : (typeof rPr.sz === 'number' && Number.isFinite(rPr.sz) ? rPr.sz / 2 : 'UNOBSERVABLE:non-numeric-sz'));
      const color = !rPr ? 'inherited' : (!('color' in rPr) ? 'inherited' : colorOf(rPr.color));
      const ind = pPr && pPr.ind && typeof pPr.ind === 'object' ? pPr.ind : null;
      const spacing = pPr && pPr.spacing && typeof pPr.spacing === 'object' ? pPr.spacing : null;
      // GH #292 S6 (LIVE): a style's ToJSON reads left / right / justify as 'start' / 'end' / 'both' -- normalized to the task-facing names
      const align = !pPr ? 'inherited' : (!('jc' in pPr) ? 'inherited' : (typeof pPr.jc === 'string' ? ({ start: 'left', end: 'right', both: 'justify' }[pPr.jc] || pPr.jc) : 'UNOBSERVABLE:non-string-jc'));
      const indentLeft = !ind ? 'inherited' : (!('left' in ind) ? 'inherited' : (typeof ind.left === 'number' && Number.isFinite(ind.left) ? ind.left / 20 : 'UNOBSERVABLE:non-numeric-ind-left'));
      const spacingAfter = !spacing ? 'inherited' : (!('after' in spacing) ? 'inherited' : (typeof spacing.after === 'number' && Number.isFinite(spacing.after) ? spacing.after / 20 : 'UNOBSERVABLE:non-numeric-spacing-after'));
      // GH #292 S1 / S2 (LIVE, docs/evidence/gh292-style-definition-probe-20261021.json): the further keys a style's ToJSON carries --
      // rPr i / u / strike (booleans), rFonts.ascii; pPr spacing.before, ind.right, ind.firstLine (twips; negative = a hanging indent),
      // keepNext / keepLines / pageBreakBefore (booleans). An absent key reads 'inherited', as above.
      const flag = (o, k) => (!o ? 'inherited' : (!(k in o) ? 'inherited' : !!o[k]));
      const twipsPt = (o, k, why) => (!o ? 'inherited' : (!(k in o) ? 'inherited' : (typeof o[k] === 'number' && Number.isFinite(o[k]) ? o[k] / 20 : 'UNOBSERVABLE:' + why)));
      const fonts = rPr && rPr.rFonts && typeof rPr.rFonts === 'object' ? rPr.rFonts : null;
      const fontFamily = !rPr ? 'inherited' : (!fonts || !('ascii' in fonts) ? 'inherited' : (typeof fonts.ascii === 'string' ? fonts.ascii : 'UNOBSERVABLE:non-string-font'));
      const more = { text: { italic: flag(rPr, 'i'), underline: flag(rPr, 'u'), strike: flag(rPr, 'strike'), fontFamily },
        paragraph: { spacingBefore: twipsPt(spacing, 'before', 'non-numeric-spacing-before'), indentRight: twipsPt(ind, 'right', 'non-numeric-ind-right'),
          indentFirstLine: twipsPt(ind, 'firstLine', 'non-numeric-ind-first-line'), keepNext: flag(pPr, 'keepNext'), keepLines: flag(pPr, 'keepLines'), pageBreakBefore: flag(pPr, 'pageBreakBefore') } };
      return {
        name, identity,
        text: { bold, size, color, ...more.text },
        paragraph: { align, indentLeft, spacingAfter, ...more.paragraph },
      };
    };
    const all = typeof doc.GetAllStyles === 'function' ? (doc.GetAllStyles() || []) : [];
    docInfo.styles = {};
    for (const st of all) {
      try {
        const nm = typeof st.GetName === 'function' ? String(st.GetName()) : null;
        if (nm) docInfo.styles[nm] = styleOf(nm);
      } catch (_) { /* skip this one style */ }
    }
    // GH #292 S3: basedOn styleId -> the base style's NAME, among the styles enumerated here; a base this enumeration did not see is named, never guessed
    const nameOfId = {}; for (const [nm, sid] of Object.entries(styleIdOf)) nameOfId[sid] = nm;
    for (const v of Object.values(docInfo.styles)) if (v && v.identity) {
      const nameOf = (id, why) => (id === null ? null : (Object.prototype.hasOwnProperty.call(nameOfId, String(id)) ? nameOfId[String(id)] : 'UNOBSERVABLE:' + why));
      v.identity.basedOnName = nameOf(v.identity.basedOn, 'base-style-not-enumerated');
      v.identity.nextName = nameOf(v.identity.next, 'next-style-not-enumerated'); v.identity.linkName = nameOf(v.identity.link, 'linked-style-not-enumerated');
    }
  } catch (_) { /* unread */ }
  return { paragraphs: out, document: docInfo, searchFailures${sl ? ', next: __next, fp: __fp()' : ''} };
} catch (err) { return { error: String(err) }; }
`;
}
const PROPERTY_QUERY = propertyQuery();

/**
 * What readLive hands the product from the page-side property query. The query's outer catch returns { error }, and a failed call returns { ok: false, outcome, error }: either way
 * `paragraphs` and `document` used to become null with NOTHING saying why (g266v1 ROUND4B / gbolv1: every rich field and the bookmark list null, in a document whose pieces all read fine).
 * `richUnobservable` is now the NAMED reason ('UNOBSERVABLE:<where>: <message>'), absent when the query read both parts; the payload stays ok:true so the HTML-only reads keep working.
 *
 * gntw3 (nested-table document, 2026-10-19): the query's own script never legitimately returns null -- it
 * always ends `return { paragraphs, document, searchFailures }` or `return { error }` (offline-VM confirmed
 * on a 3-top-level + 1-nested-table document: no throw, no leaked API object anywhere in the tree, fully
 * JSON.stringify-able). So a `props.ok === true` call whose own `result` is `null` means the callCommand
 * callback fired with NO value at all -- it escaped the native bridge before runInEditor's `r === undefined
 * ? null : r` line turned the missing value into null -- not "the script computed and returned an empty or
 * malformed read". Named on its own, ahead of the generic branches, so the symptom reads "the result never
 * arrived" instead of the misleading "no paragraphs array" (which reads as a shape problem on a result that
 * was actually never received).
 */
/**
 * inflateNestedTables(document) -> the same document with every top-level holding cell's `nestedTablesJson` string turned back into `cell.nestedTables` (the page hands it
 * over as a string, see readCell); an unparseable string is NAMED ('UNOBSERVABLE:nested-table-json: <message>'), never dropped. Cells without it are untouched.
 */
function inflateNestedTables(document) {
  if (!document || !Array.isArray(document.tables)) return document;
  for (const t of document.tables) {
    for (const row of (t && Array.isArray(t.rows) ? t.rows : [])) {
      for (const c of (Array.isArray(row) ? row : [])) {
        if (!c || typeof c.nestedTablesJson !== 'string') continue;
        try { c.nestedTables = JSON.parse(c.nestedTablesJson); } catch (e) { c.nestedTables = `UNOBSERVABLE:nested-table-json: ${String((e && e.message) || e).slice(0, 80)}`; }
        delete c.nestedTablesJson;
      }
    }
  }
  return document;
}

function richOutcome(props) {
  const r = props && props.ok ? props.result : null;
  const paragraphs = r && Array.isArray(r.paragraphs) ? r.paragraphs : null;
  const document = r && r.document && typeof r.document === 'object' ? inflateNestedTables(r.document) : null;
  const clip = x => String(x).slice(0, 160);
  let richUnobservable;
  if (!props || !props.ok) richUnobservable = `UNOBSERVABLE:property-query-call: ${clip((props && (props.error || props.outcome)) || 'no result')}`;
  else if (r === null) richUnobservable = 'UNOBSERVABLE:property-query: command-completed with no result value (escaped before delivery, not computed as empty)';
  else if (typeof r.error === 'string') richUnobservable = `UNOBSERVABLE:property-query: ${clip(r.error)}`;
  else if (!paragraphs) richUnobservable = 'UNOBSERVABLE:property-query: no paragraphs array';
  else if (!document) richUnobservable = 'UNOBSERVABLE:property-query: no document object';
  else if (Array.isArray(r.searchFailures) && r.searchFailures.length > 0) richUnobservable = `UNOBSERVABLE:search-throws: paragraph ${r.searchFailures.map(x => x.index).join(', ')}: ${clip(r.searchFailures[0].error)}`;
  return { paragraphs, document, richUnobservable };
}

/**
 * slicedPropertyRead(runInEditor, opts) -> the same { ok, outcome, result: { paragraphs, document, searchFailures } } one
 * propertyQuery call gives, read in time-budgeted slices (see PROPERTY_SLICE_BUDGET_MS). A failed slice is returned as
 * is (richOutcome names it); a slice that does not advance, or slices reporting different document fingerprints, end
 * the read with a NAMED error. Tables are merged in walk order; a slice whose table list is unread leaves it unread.
 */
async function slicedPropertyRead(runInEditor, opts, budgetMs = PROPERTY_SLICE_BUDGET_MS) {
  const fail = error => ({ ok: true, outcome: 'command-completed', result: { error } });
  const fps = new Set();
  const paragraphs = [];
  const searchFailures = [];
  let slices = 0;
  const step = async (part, from) => {
    slices += 1;
    const r = await runInEditor(propertyQuery({ ...opts, slice: { part, from, budgetMs } }));
    const v = r && r.ok ? r.result : null;
    if (!v || typeof v.error === 'string') return { stop: r };
    if (typeof v.fp !== 'string') return { stop: fail(`sliced read: ${part} slice ${from} reported no fingerprint`) };
    fps.add(v.fp);
    if (v.next !== null && !(Number.isInteger(v.next) && v.next > from)) return { stop: fail(`sliced read: ${part} slice ${from} did not advance`) };
    return { v };
  };
  for (let from = 0; ;) {
    const { v, stop } = await step('paragraphs', from);
    if (stop) return stop;
    if (!Array.isArray(v.paragraphs)) return fail(`sliced read: paragraph slice ${from} has no paragraphs array`);
    paragraphs.push(...v.paragraphs);
    if (Array.isArray(v.searchFailures)) searchFailures.push(...v.searchFailures);
    if (v.next === null) break;
    from = v.next;
  }
  let tables = [];
  let document = null;
  for (let from = 0; ;) {
    const { v, stop } = await step('document', from);
    if (stop) return stop;
    const d = v.document && typeof v.document === 'object' ? v.document : null;
    if (!d) return fail(`sliced read: document slice ${from} has no document object`);
    tables = Array.isArray(tables) && Array.isArray(d.tables) ? tables.concat(d.tables) : undefined;
    if (v.next === null) { document = d; break; }
    from = v.next;
  }
  if (fps.size !== 1) return fail(`sliced read: the document changed between slices (${fps.size} fingerprints in ${slices} slices)`);
  const rest = { ...document };
  delete rest.tables;
  return { ok: true, outcome: 'command-completed', result: { paragraphs, document: tables === undefined ? rest : { ...rest, tables }, searchFailures }, slices };
}

/**
 * W-10 (style DEFINITIONS, LIVE run3): `doc.GetAllStyles()` was MEASURED to return
 * ONLY the built-in/default styles -- a style just created via `doc.CreateStyle(name,
 * 'paragraph')` is ABSENT from it, in the SAME session, after a save, and after a reopen,
 * EVEN THOUGH `doc.GetStyle(name)` finds it every time. The `document.styles` map built by
 * PROPERTY_QUERY above is therefore STRUCTURALLY BLIND to any newly created, never-yet-
 * enumerated style -- this is what made every `define_style` round verify-uncertain.
 *
 * DESIGN DECISION (the review asked for one, stated here): a TARGETED, name-by-name read,
 * NOT a parameterized PROPERTY_QUERY. `PROPERTY_QUERY` stays the EXACT SAME frozen string
 * constant it always was (four other files already import it as a plain string --
 * test/static/table-font-read.test.cjs, test/static/m4-closure-gates.test.cjs,
 * test/live/m9-2-read-anatomy.live.cjs, test/live/docx-range-probe.live.cjs -- turning it into
 * a function would break all four for a change that only this ONE family needs). Instead,
 * `readLive(handle, request)` -- a SECOND parameter that was ALREADY plumbed end-to-end
 * (src/core/session/handle.cjs's `readLive` -> `driverOf(handle).readLive(handle, request)`,
 * `request` sourced from `options.observation.live` via `createExecutor(task, { observation })`)
 * but UNUSED until now -- runs this SEPARATE, SMALL, name-targeted query ADDITIONALLY, and
 * merges its results into `document.styles` (by name; a targeted hit WINS over a same-named
 * GetAllStyles() entry, since GetStyle(name) is the measured-reliable identity).
 *
 * `families/style-definition.cjs` exports `styleNamesOf(task)` so ANY caller (this driver's own
 * LIVE harness, and -- once wired -- a real MCP entry point outside this source snapshot) can
 * derive the right `styleNames` from a task's `define_style`/`modify_style` steps without
 * knowing this plumbing's internals.
 *
 * NAMED LIMITATION, NOT SOLVED HERE: `verifyPreservation`'s "every OTHER style unchanged" claim
 * remains bounded to GetAllStyles()'s own enumeration UNION the names THIS task's own
 * operations request -- a THIRD custom style, neither built-in nor named by this task, stays
 * invisible to it. This is the SAME pre-existing enumeration gap, not a new one, and is not
 * closed by this change.
 *
 * NAMED DUPLICATION, NOT REFACTORED HERE: the bold/size/color/align/indentLeft/spacingAfter
 * extraction logic below is a VERBATIM COPY of PROPERTY_QUERY's own `styleOf`/`colorOf`/`hex2`
 * (not factored into a shared fragment, to avoid a nested-backtick restructuring under this
 * unit's time budget -- the SAME backtick-inside-a-template-literal class of bug this session
 * already hit once today). If one copy's unit/shape logic changes, the other needs the SAME
 * change -- flagged, not automated.
 *
 * W-10-7 (LIVE run4): `doc.GetStyle(<name>)` was MEASURED to NOT return null for an
 * unknown name -- it falls back to the `Normal` style object (own diagnostic: a NEG_CONTROL
 * query for a never-created name came back `{found:true, identity:{name:'Normal'}}`). Both
 * `styleOf` copies (PROPERTY_QUERY's own, and this one) now require `GetName() === name` before
 * accepting a `GetStyle` result; `families/style-definition.cjs`'s `commandBody` carries the SAME
 * guard for its `define`/`modify` branches (it independently called `doc.GetStyle(m.name)` and
 * tested bare truthiness, which this fallback made ALWAYS true -- `define` never created a style
 * in LIVE, explaining run1-3's `style-not-found-after-write`).
 */
function buildTargetedStyleQuery(names) {
  return `
try {
  const doc = Api.GetDocument();
  const styleIdOf = {};
  const hex2 = n => Number(n).toString(16).padStart(2, '0').toUpperCase();
  const colorOf = (c) => {
    if (!c || typeof c !== 'object') return 'UNOBSERVABLE:non-object-color';
    if (c.auto === true) return 'auto';
    if (Number.isInteger(c.r) && Number.isInteger(c.g) && Number.isInteger(c.b)) return hex2(c.r) + hex2(c.g) + hex2(c.b);
    return 'UNOBSERVABLE:unrecognized-color-shape';
  };
  const styleOf = (name) => {
    const rawSt = typeof doc.GetStyle === 'function' ? doc.GetStyle(name) : null;
    const st = rawSt && typeof rawSt.GetName === 'function' && rawSt.GetName() === name ? rawSt : null;
    if (!st) return null;
    let tj = null;
    let tjReadFailed = false;
    try {
      const raw = typeof st.ToJSON === 'function' ? st.ToJSON() : null;
      tj = typeof raw === 'string' ? JSON.parse(raw) : raw;
      if (!tj || typeof tj !== 'object') tjReadFailed = true;
    } catch (_) { tj = null; tjReadFailed = true; }
    if (tjReadFailed) {
      const u = 'UNOBSERVABLE:no-readable-ToJSON';
      return { name, identity: null, text: { bold: u, size: u, color: u, italic: u, underline: u, strike: u, fontFamily: u }, paragraph: { align: u, indentLeft: u, spacingAfter: u, spacingBefore: u, indentRight: u, indentFirstLine: u, keepNext: u, keepLines: u, pageBreakBefore: u } };
    }
    const identity = { name: typeof tj.name === 'string' ? tj.name : null,
      basedOn: tj.basedOn === undefined ? null : tj.basedOn,
      // GH #292 S4 (LIVE, docs/evidence/gh292s4-next-link-internal-probe-20261022.json): next / link are styleIds in the same ToJSON, resolved to names below
      next: tj.next === undefined ? null : tj.next, link: tj.link === undefined ? null : tj.link,
      styleType: typeof tj.styleType === 'string' ? tj.styleType : null };
    if (tj.styleId !== undefined && tj.styleId !== null) styleIdOf[name] = String(tj.styleId); // GH #292 S3: see PROPERTY_QUERY
    const rPr = tj.rPr && typeof tj.rPr === 'object' ? tj.rPr : null;
    const pPr = tj.pPr && typeof tj.pPr === 'object' ? tj.pPr : null;
    const bold = !rPr ? 'inherited' : (!('b' in rPr) ? 'inherited' : !!rPr.b);
    const size = !rPr ? 'inherited' : (!('sz' in rPr) ? 'inherited' : (typeof rPr.sz === 'number' && Number.isFinite(rPr.sz) ? rPr.sz / 2 : 'UNOBSERVABLE:non-numeric-sz'));
    const color = !rPr ? 'inherited' : (!('color' in rPr) ? 'inherited' : colorOf(rPr.color));
    const ind = pPr && pPr.ind && typeof pPr.ind === 'object' ? pPr.ind : null;
    const spacing = pPr && pPr.spacing && typeof pPr.spacing === 'object' ? pPr.spacing : null;
    // GH #292 S6: the same jc normalization as PROPERTY_QUERY's styleOf
    const align = !pPr ? 'inherited' : (!('jc' in pPr) ? 'inherited' : (typeof pPr.jc === 'string' ? ({ start: 'left', end: 'right', both: 'justify' }[pPr.jc] || pPr.jc) : 'UNOBSERVABLE:non-string-jc'));
    const indentLeft = !ind ? 'inherited' : (!('left' in ind) ? 'inherited' : (typeof ind.left === 'number' && Number.isFinite(ind.left) ? ind.left / 20 : 'UNOBSERVABLE:non-numeric-ind-left'));
    const spacingAfter = !spacing ? 'inherited' : (!('after' in spacing) ? 'inherited' : (typeof spacing.after === 'number' && Number.isFinite(spacing.after) ? spacing.after / 20 : 'UNOBSERVABLE:non-numeric-spacing-after'));
    // GH #292 S1 / S2 (LIVE, docs/evidence/gh292-style-definition-probe-20261021.json): the further keys a style's ToJSON carries --
    // rPr i / u / strike (booleans), rFonts.ascii; pPr spacing.before, ind.right, ind.firstLine (twips; negative = a hanging indent),
    // keepNext / keepLines / pageBreakBefore (booleans). An absent key reads 'inherited', as above.
    const flag = (o, k) => (!o ? 'inherited' : (!(k in o) ? 'inherited' : !!o[k]));
    const twipsPt = (o, k, why) => (!o ? 'inherited' : (!(k in o) ? 'inherited' : (typeof o[k] === 'number' && Number.isFinite(o[k]) ? o[k] / 20 : 'UNOBSERVABLE:' + why)));
    const fonts = rPr && rPr.rFonts && typeof rPr.rFonts === 'object' ? rPr.rFonts : null;
    const fontFamily = !rPr ? 'inherited' : (!fonts || !('ascii' in fonts) ? 'inherited' : (typeof fonts.ascii === 'string' ? fonts.ascii : 'UNOBSERVABLE:non-string-font'));
    const more = { text: { italic: flag(rPr, 'i'), underline: flag(rPr, 'u'), strike: flag(rPr, 'strike'), fontFamily },
      paragraph: { spacingBefore: twipsPt(spacing, 'before', 'non-numeric-spacing-before'), indentRight: twipsPt(ind, 'right', 'non-numeric-ind-right'),
        indentFirstLine: twipsPt(ind, 'firstLine', 'non-numeric-ind-first-line'), keepNext: flag(pPr, 'keepNext'), keepLines: flag(pPr, 'keepLines'), pageBreakBefore: flag(pPr, 'pageBreakBefore') } };
    return { name, identity, text: { bold, size, color, ...more.text }, paragraph: { align, indentLeft, spacingAfter, ...more.paragraph } };
  };
  const names = ${JSON.stringify(names)};
  const out = {};
  for (const name of names) { try { out[name] = styleOf(name); } catch (_) { out[name] = null; } }
  // GH #292 S3: basedOn styleId -> name -- first among the targeted styles (a session-created base GetAllStyles is blind to), then, only when
  // still unresolved, the GetAllStyles enumeration (the built-in styles); a base found in neither is named, never guessed
  const nameOfId = {}; for (const [nm, sid] of Object.entries(styleIdOf)) nameOfId[sid] = nm;
  const pending = Object.values(out).filter(v => v && v.identity && [v.identity.basedOn, v.identity.next, v.identity.link].some(id => id !== null && !Object.prototype.hasOwnProperty.call(nameOfId, String(id))));
  if (pending.length && typeof doc.GetAllStyles === 'function') {
    for (const st of (doc.GetAllStyles() || [])) {
      try { const raw = st.ToJSON(); const j = typeof raw === 'string' ? JSON.parse(raw) : raw; if (j && j.styleId !== undefined && j.styleId !== null && typeof st.GetName === 'function') nameOfId[String(j.styleId)] = String(st.GetName()); } catch { /* skip this one style */ }
    }
  }
  const nameOf = (id, why) => (id === null ? null : (Object.prototype.hasOwnProperty.call(nameOfId, String(id)) ? nameOfId[String(id)] : 'UNOBSERVABLE:' + why));
  for (const v of Object.values(out)) if (v && v.identity) {
    v.identity.basedOnName = nameOf(v.identity.basedOn, 'base-style-not-enumerated');
    v.identity.nextName = nameOf(v.identity.next, 'next-style-not-enumerated'); v.identity.linkName = nameOf(v.identity.link, 'linked-style-not-enumerated');
  }
  return out;
} catch (err) { return { error: String(err) }; }
`;
}

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

  // GH #220 1b: the GH #100 U1/U3 warm-session surface -- the SAME helper xlsx-driver.cjs's own aliveNow uses
  // (a cheap inEditor evaluate, false on any failure, never throws); docx-driver.cjs had no such helper yet.
  const aliveNow = () => (state.page && state.frame ? inEditor(() => true).then(() => true, () => false) : Promise.resolve(false));

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
        try {
          await state.page.waitForFunction(() => window.__ready === true || window.__error, null, { timeout: timeoutMs, polling: pollMs });
        } catch {
          // GitHub #77: this is reachable only when NEITHER __ready NOR __error ever became truthy in the
          // whole window (onError sets __error, so it resolves the wait above, not this catch) -- read back
          // onWarning/onInfo, or confirm literally nothing arrived
          const e = new Error(openWaitFailureDetail(await readDsOpenState(state.page)));
          e.name = 'EditorWaitTimeout';
          throw e;
        }
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
        // GitHub #77: the ready/error wait above resolves as soon as DS reports EITHER __ready or __error --
        // an onError does NOT fail-fast (whether onDocumentReady still arrives afterward is NEM-MERT, so a
        // successful open must not be cut short), which means a real onError lands HERE, not in the catch
        // below. Read back whatever DS reported before the generic message swallows it.
        // the DS description is page-reported text -- redacted like every other detail that leaves open()
        const dsDetail = (state.host ? state.host.redact : (x => x))(openWaitFailureDetail(await readDsOpenState(state.page)));
        await teardown();
        return { ok: false, detail: `live editor model not ready (${dsDetail})` };
      } catch (err) {
        // M11.3b.1: a navigation error carries the host URL -- the session token is redacted before anything leaves
        const redact = state.host ? state.host.redact : (x => x);
        await teardown();
        // a controlled config-source refusal (M11.2) keeps its product error class; the editor was never started
        if (err && err.name === 'ConfigSourceError') return { ok: false, code: err.code, detail: err.detail || err.code };
        // GitHub #77: the ready/error wait timed out -- err.message already names the DS's own code/description,
        // what onWarning/onInfo recorded, or 'no-ds-event' (config-host.cjs openWaitFailureDetail)
        if (err && err.name === 'EditorWaitTimeout') return { ok: false, detail: redact(err.message) };
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

    /**
     * release -- GH #100 U1/U3 (docx U1b, the #220-named blocking gap, mirroring xlsx-driver.cjs:1480-1495
     * method-for-method -- docx's own state shape (`state.saveAck`/`state.mutated`/`state.forceSaveOp`) is
     * ALREADY byte-identical to xlsx's for these three fields, so no adaptation of the save-ack pairing
     * itself was needed, only this wrapper): the CALL END of a session that STAYS open. close() minus its
     * two destructive steps -- the same bounded wait for the DS acknowledgement of the editor's last
     * saveChanges and the same re-evaluation of the forceSave confirmation, but NO requestClose and NO
     * teardown. `ok` means "the call ended with every change sent and acked"; a release that is not ok is a
     * failed call end, and the caller must close() the session instead of reusing it. close() itself is
     * untouched by this addition.
     */
    async release() {
      const obs = state.page ? state.saveAck : null;
      const op = state.forceSaveOp;
      const pre = await settleSaveAck(obs, { timeoutMs: closeAckTimeoutMs, pollMs, wait: defaultWait, mutated: state.mutated, readCanSave });
      const mark = obs ? obs.mark() : null;
      const live = await aliveNow();
      return closeResult(obs, { pre, mark, op, lifecycle: { ok: live, detail: live ? 'session-kept' : 'no live editor' }, torn: { ok: true, detail: 'kept' } });
    },

    /**
     * resetForNextCall -- GH #100 U1 (docx U1b): the per-call state a kept session must not carry into the
     * next call. `mutated` is "a write was attempted in this call" (a close/release needs an acked save from
     * there) and `forceSaveOp` is the last forceSave operation window; both belong to ONE call. The save
     * observer is NOT reset: its pairing is per DS socket and per send (the same reason xlsx-driver.cjs's
     * own resetForNextCall leaves it alone).
     */
    resetForNextCall() {
      state.mutated = false;
      state.forceSaveOp = null;
      return { ok: !!(state.page && state.frame) };
    },

    /**
     * health -- GH #100 U3 (docx U1b): what the call-start check of a KEPT session needs, in one read-only
     * call that never throws. `canSave` is null when it could not be read (never read as "nothing unsaved");
     * `ds.open` is an authenticated editor <-> DS socket that is still open; `anomalies` counts the
     * observer's anomalies since open -- the SAME `state.saveAck.evidence()` shape xlsx-driver.cjs's own
     * health() reads, since both drivers share save-ack.cjs unchanged.
     */
    async health() {
      const live = await aliveNow();
      const canSave = live ? await readCanSave() : null;
      const ev = state.saveAck ? state.saveAck.evidence() : null;
      const sockets = ev ? ev.dsSockets : [];
      return { alive: live, canSave, ds: { open: sockets.some(x => x.open && x.authed), sockets: sockets.length, busy: sockets.some(x => x.busy) }, anomalies: ev ? ev.anomalies.length : null };
    },

    /** documentKey -- GH #100 U3 (docx U1b): the DS document key of the config this session was OPENED
     * with (null when none). Compared in memory, never logged -- same discipline as xlsx-driver.cjs's own. */
    documentKey() {
      const k = state.config && state.config.document ? state.config.document.key : null;
      return typeof k === 'string' && k !== '' ? k : null;
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
    /** request.numberingFormat (default false): true = also read each listed paragraph's own-level format through the
     * PUBLIC ApiNumbering.ToJSON() (W-07 set_numbering); off by default, so every other read stays Bridge-only. */
    /** request.includeComments (default false, GH #269b): true = also read the main document's comments through the PUBLIC doc.GetAllComments() (read_comments only). */
    /** request.styleNames (W-10): names for the additional, name-targeted style read (buildTargetedStyleQuery). */
    async readLive(_handle, request = {}) {
      const html = await inEditor(() => { const e = window.editor || (window.Asc || {}).editor; return String(e.pluginMethod_GetFileHTML() || ''); });
      if (!html.trim()) return { ok: false, detail: 'empty live projection' };
      const queryOpts = { ...sectionVariantQuery(request, state.runtimeId), numberingFormat: request.numberingFormat === true, includeComments: request.includeComments === true };
      const props = paragraphTagsOf(html) >= PROPERTY_SLICE_MIN_PARAGRAPH_TAGS ? await slicedPropertyRead(runInEditor, queryOpts) : await runInEditor(propertyQuery(queryOpts));
      const { paragraphs, document, richUnobservable } = richOutcome(props);
      // W-10: doc.GetAllStyles() cannot see a freshly created style -- a SEPARATE,
      // name-targeted read (buildTargetedStyleQuery) covers exactly the names THIS call's
      // caller asked about, merged in by name (a targeted hit wins -- doc.GetStyle(name) is the
      // measured-reliable identity, GetAllStyles() is not). No-op if no names were requested.
      const styleNames = Array.isArray(request.styleNames) ? request.styleNames.filter(n => typeof n === 'string' && n) : [];
      if (styleNames.length && document) {
        const targeted = await runInEditor(buildTargetedStyleQuery(styleNames));
        const byName = targeted.ok && targeted.result && typeof targeted.result === 'object' && !targeted.result.error ? targeted.result : {};
        document.styles = { ...(document.styles || {}), ...byName };
      }
      // GH #293 L7: the whole affected set of each requested list item, through the word /5 Bridge (unread off /5)
      const memberTexts = Array.isArray(request.numberingMembers) ? [...new Set(request.numberingMembers.filter(t => typeof t === 'string' && t))] : [];
      if (memberTexts.length && document) {
        const m = await runInEditor(buildListMembersQuery(memberTexts));
        const got = m.ok && m.result && typeof m.result === 'object' && !m.result.error ? m.result : null;
        document.numberingMembers = Object.fromEntries(memberTexts.map(t => [t, got && got[t] && typeof got[t] === 'object' ? got[t] : { unread: got ? 'not-returned' : `query-failed:${(m && m.outcome) || 'unknown'}` }]));
      }
      const gate = internalGateOf(state.runtimeId);
      return { ok: true, authority: 'LIVE_READ', payload: { html, paragraphs, document, ...(richUnobservable ? { richUnobservable } : {}), runtime: { internalGate: gate, identity: state.runtimeId, cacheKey: state.cacheKey } } };
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
      // GH #292 S4: the same receipt rule as B18 -- ok ONLY when the body returned { ok: true }
      if (command.meta && command.meta.bridgeStyleNextLink === true) return b18DispatchResult(await runInEditor(s4BridgeWritePrelude() + command.body));
      // GH #293 L7: the same receipt rule -- ok ONLY when the body returned { ok: true } (the detach receipt AND the level write)
      if (command.meta && command.meta.bridgeListDetach === true) return b18DispatchResult(await runInEditor(l7BridgeWritePrelude() + command.body));
      const prelude = command.meta && command.meta.bridgeSectionVariantRead === true ? b21BridgeReadPrelude() : '';
      const r = await runInEditor(prelude + command.body);
      return { ok: r.ok === true, dispatched: true, outcome: r.outcome, error: r.error || null };
    },

    async readPackage() {
      const docUrl = state.config && state.config.document && state.config.document.url;
      if (!docUrl) return { ok: false, detail: 'no document url in the editor config' };
      const r = await fetch(docUrl);
      if (!r.ok) return { ok: false, detail: `package HTTP ${r.status}` };
      return { ok: true, bytes: Buffer.from(await r.arrayBuffer()) };
    },
  };
}

module.exports = { createEuroOfficeDocxDriver, PROPERTY_QUERY, propertyQuery, slicedPropertyRead, paragraphTagsOf, PROPERTY_SLICE_BUDGET_MS, PROPERTY_SLICE_MIN_PARAGRAPH_TAGS, richOutcome, inflateNestedTables, buildTargetedStyleQuery, DOCX_SAVE_BOUNDS, internalGateOf, internalWriteRefusal,
  B18_BRIDGE_WRITE_RUNTIMES, B21_BRIDGE_READ_RUNTIMES, S4_BRIDGE_WRITE_RUNTIMES, L7_BRIDGE_RUNTIMES, b18BridgeWritePrelude, b18DispatchResult, b21BridgeReadPrelude, s4BridgeWritePrelude,
  l7BridgeWritePrelude, buildListMembersQuery };
