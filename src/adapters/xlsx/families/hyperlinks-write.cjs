'use strict';

const { parseRange, rangeText } = require('../a1.cjs');
const { cellAt, isBlank } = require('../inventory.cjs');
const cells = require('./cells.cjs');
const { APOSTROPHE_REASON } = require('../sheet-name.cjs');

/**
 * XLSX `hyperlinks` family: set_hyperlink / remove_hyperlink (GH #300 L2, owner decision 2026-10-27,
 * "2: A ... 3: A" -- full READ/WRITE symmetry on top of L1). CANDIDATE, OFFLINE-TESTED ONLY --
 * NOT in XLSX_WRITE_ALLOWLIST (write-tools.cjs), pending its own LIVE PASS (the same WITHHELD convention
 * docx's set_list_type used). The read mirror and the verifier is read_hyperlinks (L1, families/
 * hyperlinks-read.cjs) -- its own observation (`hyperlinkSheets`, the DS Bridge /18) is reused here
 * unchanged for the gate/verify.
 *
 * ENGINE PATH (decided from EXISTING evidence, no new live run -- see
 * docs/DS-9.3.5-BRIDGE18-L2-HYPERLINK-WRITE-PLAN.md "1. Engine path"): the PUBLIC `ApiWorksheet.SetHyperlink` /
 * `ApiRange.ClearHyperlinks`, through callCommand -- the SAME pattern every other WRITE family in this
 * adapter uses (comments.cjs's AddComment is the direct precedent). The custom DS Bridge (/17, /18) is
 * reserved for READS the public getters do not expose; the public surface probe
 * (docs/evidence/m10-xlsx-public-surface-probe-20261009.json, LIVE, 2026-10-09) already lists
 * `ApiWorksheet.SetHyperlink` and `ApiRange.ClearHyperlinks` -- no Bridge WRITE handler (a /19) is needed
 * or built here.
 *
 * SIGNATURE (LIVE-measured already, as the L1 read candidate's own seed -- NOT a new run; test/live/
 * xlsx-gh300-gh298-engine-access-probe.live.cjs, docs/evidence/gh300-gh298-engine-access-probe-20261027.json):
 *   ws.SetHyperlink(sRange, sAddress, sSubAddress, sScreenTipText, sTextToDisplay)
 *   ws.GetRange(sRange).ClearHyperlinks()
 * Neither threw on any of the three measured calls; an empty-string arg is stored as `null`
 * (Location/Tooltip, measured on the reopened engine); getHyperlinkType() was 1 ("external") on every
 * measured link -- an address-only / location-only ("internal") link was NEVER exercised live, on
 * EITHER side (L1's own doc already names this the same way: 'internal' "NOT_MEASURED live, named by
 * convention"). v1 here is therefore EXTERNAL-ONLY: `address` is required; an address-empty /
 * location-only write is refused, named, not guessed into an untested shape.
 *
 * sTextToDisplay (LIVE-MEASURED 2026-10-28, test/live/xlsx-gh300-l2-hyperlink-empty-cell-probe.live.cjs on
 * the production Document Server, cell Bridge /18): SetHyperlink(.., '') REWRITES the cell's VALUE to the address, on every cell type measured
 * (empty, text AND number -- a number cell became the address string). SetHyperlink(.., <the cell's own,
 * already-read text>) leaves the value exactly as it was (a number cell stays a number). v1.1 therefore
 * NEVER passes '': `lower()` takes the resolved `target.text` (the cell's own GetValue, read at
 * classify/resolve time, the SAME fresh observation every other family's gate already reads) and passes
 * that as sTextToDisplay. The target cell's own font (color, underline) changes on every measured write
 * and ClearHyperlinks/retry does not restore it -- that is the engine's own measured cost of a hyperlink,
 * not this family's decision, and the L3 check below (verifyPreservation, reusing cells.cjs unchanged)
 * does not look at font/style at all, so it neither asserts nor blocks that change; it blocks only a
 * VALUE change on the target cell -- which the always-own-text write now avoids.
 *
 * A BLANK cell (no value) or a FORMULA cell at the target is refused before dispatch, named, not guessed
 * into an untested shape: a blank cell has no "own text" to preserve (passing '' would be the exact
 * measured corruption above), and a formula cell's own-text write was never measured and could turn the
 * formula into a literal. See `textGate()` below (xlsx-hyperlink-cell-empty /
 * xlsx-hyperlink-cell-formula-unmeasured).
 *
 * MULTI-CELL RANGE (GH #353): refused before dispatch, named `xlsx-hyperlink-range-unsupported` (see `rangeProblem()` below)
 * -- v1 final behaviour for both intents: one A1 cell only.
 *
 * NOT_MEASURED, and NOT run here (the engine-path choice above needed none of these; a live probe of
 * these specific points is this candidate's own prerequisite before any install/live step -- named in
 * full in the plan doc):
 *   - SetHyperlink's behaviour on a cell that already holds a (different) hyperlink -- comments.cjs's
 *     AddComment is measured as a silent no-op on an occupied cell; SetHyperlink was only ever called on
 *     a cell with NO prior link. v1 therefore REFUSES before the write when the cell already holds a
 *     DIFFERENT hyperlink (xlsx-hyperlink-cell-occupied), the SAME conservative posture as comments.cjs,
 *     not an assumed overwrite
 *   - SetHyperlink's own return value (THREW: is absent for all three live calls, so no throw -- but the
 *     value itself never survived the probe's own serialisation boundary; same for ClearHyperlinks and
 *     FormatAsTable, consistently, across six independent calls -- consistent with a void/undefined
 *     return, not re-asserted as fact)
 */

/*
 * GH #352 (H01) -- the INTERNAL link (LIVE-measured 2026-10-30, test/live/xlsx-gh352-internal-hyperlink-probe.live.cjs, four
 * rounds on the production Document Server, docs/evidence/gh352-internal-hyperlink-probe-live-20261030.json). The engine's own
 * ApiWorksheet.SetHyperlink source (read in round 3): sAddress is REQUIRED; a non-empty sAddress that is not an allowed-protocol
 * URL takes the internal branch, which needs sSubAddress "Sheet!Ref" (a bare "Ref" means this sheet); the sheet must exist by
 * its PLAIN name (a quoted 'Name' is refused, "Invalid subAddress"), and the reference must pass the engine's range check.
 * Round 4: SetHyperlink(cell, "<Sheet>!<Ref>", "<Sheet>!<Ref>", tip, <own text>) writes it -- same-sheet, cross-sheet with a
 * space in the name and a range target alike; stored <hyperlink location="'<Sheet>'!<Ref>"> (always the QUOTED form, no r:id);
 * a fresh product read_hyperlinks reports kind "internal" and that location; the cell value is kept (a number stays a number).
 * So `location` takes "Ref" / "Sheet!Ref" / "'Sheet'!Ref", is canonicalised to the stored quoted form, and the target sheet is
 * checked against the fresh sheet list before the write (xlsx-hyperlink-target-sheet-not-found).
 * address AND location together: the engine takes the external branch and DROPS the location (round 1: stored location null)
 * -- a write that would land and then fail its own verify; refused before dispatch instead.
 * Font: as with the external link, the engine applies its hyperlink style (blue, single underline) and the cell's earlier bold /
 * colour is lost -- measured; the same named cost the header above already states for the external write.
 */
const INTENTS = Object.freeze(['set_hyperlink', 'remove_hyperlink']);
const SET_FIELDS = Object.freeze(['intent', 'id', 'sheet', 'cell', 'address', 'location', 'tooltip']);
const REMOVE_FIELDS = Object.freeze(['intent', 'id', 'sheet', 'cell']);
// our implementation ceiling, not an engine limit (the same shape as comments.cjs's MAX_TEXT)
const MAX_ADDRESS = 2000;
const MAX_LOCATION = 2000;
const MAX_TOOLTIP = 2000;
const CONTROL = /[\u0000-\u0009\u000B-\u001F\u007F]/;

/**
 * GH #353 (H02): a multi-cell range is refused with a NAMED, pre-dispatch reason. The engine fact is LIVE-measured
 * (docs/DS-9.3.5-BRIDGE18-L2-HYPERLINK-WRITE-PLAN.md section 4): ws.SetHyperlink(<range>, ...) is a SILENT NO-OP -- it does
 * not throw and no link is created (read_hyperlinks and the stored file stay empty), so a write that lands in nothing would
 * only be caught by its own verify. remove_hyperlink on a range (ApiRange.ClearHyperlinks) was never measured; v1 refuses it
 * the same way rather than guess. Per-cell expansion would be new behaviour and is a separate decision (not made here).
 * Any ':' form is refused, A1:A1 included: v1 takes the bare A1 cell only.
 */
const RANGE_UNSUPPORTED = 'xlsx-hyperlink-range-unsupported';
function rangeProblem(at, cell) {
  if (typeof cell !== 'string' || !cell.includes(':')) return null;
  return { ok: false, detail: `${at} ${RANGE_UNSUPPORTED}: "${cell}" is a cell range; the engine's SetHyperlink on a range is a silent no-op (no error, no link -- LIVE-measured), and ClearHyperlinks on a range is unmeasured; give one cell (A1)` };
}

function textProblem(at, name, v, max, required) {
  if (v === undefined || v === null) return required ? { ok: false, detail: `${at} requires a non-empty ${name}` } : null;
  if (typeof v !== 'string') return { ok: false, detail: `${at} ${name} must be a string` };
  // empty-string is never meaningful here (omit the field instead) -- required or not
  if (v.trim() === '') return { ok: false, detail: `${at} requires a non-empty ${name} when given (omit it instead of '')` };
  if (v.length > max) return { ok: false, detail: `${at} ${name} has ${v.length} characters, over the ${max} ceiling` };
  if (CONTROL.test(v)) return { ok: false, detail: `${at} ${name} holds a control character (only printable text)` };
  return null;
}

/** GH #352: "Ref" | "Sheet!Ref" | "'Sheet'!Ref" -> { sheet, ref, stored, engine } or { problem } */
function parseLocation(v, ownSheet) {
  const bang = v.lastIndexOf('!');
  let sheet = ownSheet;
  let ref = v;
  if (bang !== -1) {
    const raw = v.slice(0, bang);
    ref = v.slice(bang + 1);
    const quoted = /^'(.*)'$/.exec(raw);
    sheet = quoted ? quoted[1] : raw;
    if (sheet === '') return { problem: 'names an empty sheet' };
  }
  // the target sheet's EXISTENCE is checked against the fresh sheet list at classify; here only the shapes the engine cannot
  // carry: an apostrophe (the sheet-name rule's measured refusal), a '!' (the engine splits the sub-address on it), the
  // characters / length no sheet name can hold
  if (sheet.includes("'")) return { problem: `target sheet "${sheet}": ${APOSTROPHE_REASON}` };
  if (sheet.includes('!') || sheet.length > 31 || /[\u0000-\u001f\[\]:*?/\\]/.test(sheet)) return { problem: `target sheet "${sheet}" is not a sheet name` };
  if (ref.includes('$')) return { problem: 'holds a $ (absolute) marker -- unmeasured, give the plain reference' };
  const g = parseRange(ref);
  if (!g) return { problem: `"${ref}" is not an A1 cell or range reference` };
  const r = rangeText(g);
  return { sheet, ref: r, stored: `'${sheet}'!${r}`, engine: `${sheet}!${r}` };
}

/** a STORED / read-back location -> "sheet\u0000REF", or null. The engine quotes the sheet name only when it needs to
 * (LIVE: "'H352'!D9" -- cell-reference shaped -- and "'GH352_x B'!A1" quoted, "GH352_x!C3" not), so a written link is
 * compared by its resolved (sheet, ref) pair, never by the string. */
function locationKey(v) {
  if (typeof v !== 'string' || v === '') return null;
  const bang = v.lastIndexOf('!');
  if (bang === -1) return null;
  const raw = v.slice(0, bang);
  const quoted = /^'(.*)'$/.exec(raw);
  const g = parseRange(v.slice(bang + 1));
  return g ? `${quoted ? quoted[1].replace(/''/g, "'") : raw}\u0000${rangeText(g)}` : null;
}

function planSet(step, index) {
  const at = `step ${index}: set_hyperlink`;
  const unknown = Object.keys(step).filter(k => !SET_FIELDS.includes(k));
  if (unknown.length) return { ok: false, detail: `${at} does not support field(s): ${unknown.join(', ')}` };
  if (typeof step.sheet !== 'string' || step.sheet === '') return { ok: false, detail: `${at} requires a sheet name` };
  const rangeP = rangeProblem(at, step.cell);
  if (rangeP) return rangeP;
  const g = parseRange(step.cell);
  if (!g || g.count !== 1) return { ok: false, detail: `${at} requires exactly one cell (A1)` };
  const hasAddress = step.address !== undefined && step.address !== null;
  const hasLocation = step.location !== undefined && step.location !== null;
  if (hasAddress && hasLocation) return { ok: false, detail: `${at} takes an address (external) OR a location (internal), not both -- the engine keeps only the address and drops the location (GH #352 LIVE probe)` };
  if (!hasAddress && !hasLocation) return { ok: false, detail: `${at} requires an address (external link) or a location (internal link)` };
  const addrP = textProblem(at, 'address', step.address, MAX_ADDRESS, false);
  if (addrP) return addrP;
  const locP = textProblem(at, 'location', step.location, MAX_LOCATION, false);
  if (locP) return locP;
  const tipP = textProblem(at, 'tooltip', step.tooltip, MAX_TOOLTIP, false);
  if (tipP) return tipP;
  let loc = null;
  if (hasLocation) {
    loc = parseLocation(step.location, step.sheet);
    if (loc.problem) return { ok: false, detail: `${at} location ${loc.problem}` };
  }
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`,
      intent: 'set_hyperlink',
      family: 'hyperlinks',
      selector: Object.freeze({ sheet: step.sheet, cell: rangeText(g) }),
      kind: loc ? 'internal' : 'external',
      address: loc ? null : step.address,
      // internal: the STORED (quoted) form read_hyperlinks reports; external: always null (no location is written)
      location: loc ? loc.stored : null,
      targetSheet: loc ? loc.sheet : null,
      engineLocation: loc ? loc.engine : null,
      tooltip: step.tooltip === undefined ? null : step.tooltip,
      // a hyperlink owns no cell content: the cell's own value is preserved, not claimed
      cells: Object.freeze([]),
      preservationMode: 'PRIMITIVE_INTEGRAL',
    },
  };
}

function planRemove(step, index) {
  const at = `step ${index}: remove_hyperlink`;
  const unknown = Object.keys(step).filter(k => !REMOVE_FIELDS.includes(k));
  if (unknown.length) return { ok: false, detail: `${at} does not support field(s): ${unknown.join(', ')}` };
  if (typeof step.sheet !== 'string' || step.sheet === '') return { ok: false, detail: `${at} requires a sheet name` };
  const rangeP = rangeProblem(at, step.cell);
  if (rangeP) return rangeP;
  const g = parseRange(step.cell);
  if (!g || g.count !== 1) return { ok: false, detail: `${at} requires exactly one cell (A1)` };
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`,
      intent: 'remove_hyperlink',
      family: 'hyperlinks',
      selector: Object.freeze({ sheet: step.sheet, cell: rangeText(g) }),
      cells: Object.freeze([]),
      preservationMode: 'PRIMITIVE_INTEGRAL',
    },
  };
}

function plan(step, index) {
  if (step.intent === 'set_hyperlink') return planSet(step, index);
  return planRemove(step, index);
}

/** The sheet's hyperlink list (read_hyperlinks' own `hyperlinkReads`, DS Bridge /18), or the named reason it cannot be told. */
function gate(p, op) {
  if (!p || p.format !== 'xlsx') return { ok: false, reason: 'xlsx-observation-missing' };
  if (!p.complete) return { ok: false, reason: p.truncated ? 'xlsx-observation-truncated' : 'xlsx-observation-unreadable-getters', unsupported: p.unsupported };
  const s = p.sheets.filter(x => x.name === op.selector.sheet);
  if (s.length !== 1) return { ok: false, reason: s.length > 1 ? 'xlsx-sheet-ambiguous' : 'xlsx-sheet-not-found', available: p.sheets.map(x => x.name) };
  const read = p.hyperlinkReads ? p.hyperlinkReads[op.selector.sheet] : undefined;
  if (read === undefined) return { ok: false, reason: 'xlsx-hyperlinks-unobservable', why: 'not-read' };
  if (typeof read === 'string') return { ok: false, reason: 'xlsx-hyperlinks-unobservable', why: read };
  if (!read || !Array.isArray(read.links)) return { ok: false, reason: 'xlsx-hyperlinks-unobservable', why: 'malformed-read' };
  return { ok: true, links: read.links, at: read.links.find(l => l.ref === op.selector.cell) || null };
}

const norm = v => (v === undefined || v === '' ? null : v);
/** exactly the written hyperlink: address, and location/tooltip, '' on write normalised the same way the engine measured (-> null) */
const locationMatches = (l, op) => (op.kind === 'internal' ? locationKey(l.location) === locationKey(op.location) : norm(l.location) === null);
const isWritten = (l, op) => !!l && norm(l.address) === norm(op.address) && locationMatches(l, op) && norm(l.tooltip) === norm(op.tooltip);

/**
 * set_hyperlink's own prerequisite (family decision from the 2026-10-28 LIVE probe above): the
 * target cell's own, already-read text, for sTextToDisplay -- never ''. A cell outside every read
 * window/used-range is unread (xlsx-target-cell-unread, defensive -- `windowsOf` below always asks
 * for this exact cell, so this should not occur live); a cell with no value is refused named
 * (xlsx-hyperlink-cell-empty, 0 writes -- an empty sTextToDisplay would silently write the address
 * as the cell's value, the measured corruption this family now avoids); a formula cell is refused
 * named too (xlsx-hyperlink-cell-formula-unmeasured -- an own-text write over a formula was never
 * measured and could turn the formula into a literal).
 */
function textGate(p, op) {
  const cell = cellAt(p, op.selector.sheet, op.selector.cell);
  if (!cell) return { ok: false, reason: 'xlsx-target-cell-unread' };
  if (isBlank(cell)) return { ok: false, reason: 'xlsx-hyperlink-cell-empty' };
  if (cell.f !== '') return { ok: false, reason: 'xlsx-hyperlink-cell-formula-unmeasured' };
  return { ok: true, text: String(cell.v) };
}

function classifySet(p, op) {
  const g = gate(p, op);
  if (!g.ok) return { state: 'FAIL_CLOSED', evidence: g };
  const where = { sheet: op.selector.sheet, cell: op.selector.cell };
  if (isWritten(g.at, op)) return { state: 'SATISFIED', evidence: { ...where, proof: 'fresh-live-the-cell-holds-this-hyperlink' } };
  if (g.at) return { state: 'FAIL_CLOSED', evidence: { ...where, reason: 'xlsx-hyperlink-cell-occupied' } };
  if (op.kind === 'internal' && !p.sheets.some(x => x.name === op.targetSheet)) return { state: 'FAIL_CLOSED', evidence: { ...where, reason: 'xlsx-hyperlink-target-sheet-not-found', targetSheet: op.targetSheet } };
  const t = textGate(p, op);
  if (!t.ok) return { state: 'FAIL_CLOSED', evidence: { ...where, reason: t.reason } };
  return { state: 'APPLY', evidence: where };
}

function classifyRemove(p, op) {
  const g = gate(p, op);
  if (!g.ok) return { state: 'FAIL_CLOSED', evidence: g };
  const where = { sheet: op.selector.sheet, cell: op.selector.cell };
  if (!g.at) return { state: 'SATISFIED', evidence: { ...where, proof: 'fresh-live-no-hyperlink-at-the-cell' } };
  return { state: 'APPLY', evidence: where };
}

function classify(p, op) {
  return op.intent === 'set_hyperlink' ? classifySet(p, op) : classifyRemove(p, op);
}

function resolve(p, selector, ctx = {}) {
  const op = ctx.operation;
  const g = gate(p, op);
  if (!g.ok) return { ok: false, target: null, selector, ambiguity: g.reason, detail: g };
  if (op && op.intent === 'set_hyperlink') {
    const t = textGate(p, op);
    if (!t.ok) return { ok: false, target: null, selector, ambiguity: t.reason, detail: t };
    return { ok: true, target: Object.freeze({ sheet: selector.sheet, cell: selector.cell, text: t.text }) };
  }
  return { ok: true, target: Object.freeze({ sheet: selector.sheet, cell: selector.cell }) };
}

/** One callCommand each: exact-name sheet check, then the public ApiWorksheet/ApiRange writer.
 * set_hyperlink needs `target.text` (the cell's own read text, resolve()'s own textGate) -- resolve()
 * always runs before this in the real pipeline (adapter.cjs's executePrimitive); a direct call without
 * a resolved target is a caller error, not a silent '' fallback (that is the corruption this family now
 * refuses to reproduce). */
function lower(op, target) {
  const { sheet, cell } = op.selector;
  if (op.intent === 'set_hyperlink') {
    if (!target || typeof target.text !== 'string') throw new Error('hyperlinks-write.lower: set_hyperlink requires a resolved target with text -- call resolve() first');
    // internal (GH #352): sAddress AND sSubAddress are both "<Sheet>!<Ref>" (plain name) -- the measured round-4 shape
    const address = op.kind === 'internal' ? op.engineLocation : op.address;
    const location = op.kind === 'internal' ? op.engineLocation : '';
    const tooltip = op.tooltip === null ? '' : op.tooltip;
    return {
      kind: 'callCommand',
      body: `
try {
  const name = ${JSON.stringify(sheet)};
  const ws = Api.GetSheet(name);
  if (!ws || ws.GetName() !== name) return { ok: false, stage: 'sheet-identity' };
  ws.SetHyperlink(${JSON.stringify(cell)}, ${JSON.stringify(address)}, ${JSON.stringify(location)}, ${JSON.stringify(tooltip)}, ${JSON.stringify(target.text)});
  return { ok: true };
} catch (err) { return { ok: false, stage: 'threw', error: String(err) }; }
`,
      meta: Object.freeze({ family: 'hyperlinks', action: 'set', sheet, cell, kind: op.kind, address: op.address, location: op.location, tooltip: op.tooltip, text: target.text }),
    };
  }
  return {
    kind: 'callCommand',
    body: `
try {
  const name = ${JSON.stringify(sheet)};
  const ws = Api.GetSheet(name);
  if (!ws || ws.GetName() !== name) return { ok: false, stage: 'sheet-identity' };
  ws.GetRange(${JSON.stringify(cell)}).ClearHyperlinks();
  return { ok: true };
} catch (err) { return { ok: false, stage: 'threw', error: String(err) }; }
`,
    meta: Object.freeze({ family: 'hyperlinks', action: 'remove', sheet, cell }),
  };
}

function verifyPrimitive(op, before, after) {
  const g = gate(after && after.primary, op);
  if (!g.ok) return { ok: false, evidence: { reason: g.reason, why: g.why || null } };
  if (op.intent === 'set_hyperlink') {
    const l = g.at;
    return { ok: isWritten(l, op), evidence: { sheet: op.selector.sheet, cell: op.selector.cell, found: !!l,
      addressMatches: !!l && norm(l.address) === norm(op.address), locationMatches: !!l && locationMatches(l, op),
      tooltipMatches: !!l && norm(l.tooltip) === norm(op.tooltip), proof: 'fresh-live-read_hyperlinks' } };
  }
  return { ok: !g.at, evidence: { sheet: op.selector.sheet, cell: op.selector.cell, stillPresent: !!g.at, proof: 'fresh-live-read_hyperlinks' } };
}

const sameLink = (x, y) => JSON.stringify(x) === JSON.stringify(y);

/** every OTHER hyperlink of the sheet unchanged, and every cell (a hyperlink owns none) */
function verifyPreservation(op, before, after) {
  const b = gate(before && before.primary, op);
  const a = gate(after && after.primary, op);
  if (!b.ok || !a.ok) return { ok: false, evidence: { reason: 'xlsx-hyperlink-preservation-unobservable', before: b.ok ? null : b.reason, after: a.ok ? null : a.reason } };
  const others = list => list.filter(l => l.ref !== op.selector.cell);
  const was = others(b.links);
  const now = others(a.links);
  const changed = was.filter(l => !now.some(x => sameLink(x, l))).map(l => l.ref);
  const added = now.filter(l => !was.some(x => x.ref === l.ref)).map(l => l.ref);
  const rest = cells.verifyPreservation(op, before, after);
  return { ok: changed.length === 0 && added.length === 0 && now.length === was.length && rest.ok, evidence: { ...rest.evidence, linksCompared: was.length, linksChanged: changed.slice(0, 20), linksAdded: added.slice(0, 20) } };
}

function wholeTaskChecks(op, finalObservation) {
  const g = gate(finalObservation && finalObservation.primary, op);
  const ok = op.intent === 'set_hyperlink' ? (g.ok && isWritten(g.at, op)) : (g.ok && !g.at);
  const name = op.intent === 'set_hyperlink' ? `${op.selector.sheet}!${op.selector.cell} hyperlink matches` : `${op.selector.sheet}!${op.selector.cell} has no hyperlink`;
  return [{ name, ok, detail: ok ? null : { reason: g.reason || 'not-the-expected-state' } }];
}

const windowsOf = op => [{ sheet: op.selector.sheet, range: op.selector.cell }];
const liveExtras = op => ({ hyperlinkSheets: [op.selector.sheet] });

module.exports = { RANGE_UNSUPPORTED, INTENTS, MAX_ADDRESS, MAX_LOCATION, MAX_TOOLTIP, plan, classify, resolve, lower, verifyPrimitive, verifyPreservation, wholeTaskChecks, windowsOf, liveExtras };
