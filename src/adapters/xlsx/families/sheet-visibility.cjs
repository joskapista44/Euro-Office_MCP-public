'use strict';

const cells = require('./cells.cjs');
const { sheetNameProblem } = require('../sheet-name.cjs');

/**
 * XLSX `sheet-visibility` family: set_sheet_visibility (GH #303 R1) -- one sheet 'visible' | 'hidden' through the public
 * ApiWorksheet.SetVisible(bool). The READ mirror and the verifier is read_sheet_state's `visibility` (GetVisible).
 * The R1 LIVE design probe (docs/evidence/gh303-r1-sheet-visibility-probe-20261026.json):
 *   - hide / show is effective in session, after a fresh reopen, in the saved <sheet state="hidden"> and in the installed
 *     product's read_sheet_state; SetVisible returns nothing (its effect is the only answer)
 *   - hiding the ACTIVE sheet moves the active sheet to the next visible one (the active sheet is not part of this
 *     intent's L3: a product write stores the last-changed sheet as active anyway -- #303 R2, closed as expected)
 *   - *** the engine HIDES THE LAST VISIBLE SHEET TOO (0 visible sheets left) *** -- so that refusal is OURS, made
 *     before the write from the read AND re-checked in the dispatch, never left to the engine
 * 'hidden' is the plain hidden state (the public setter's only one); a very-hidden sheet reads 'hidden' like any other.
 * Refused before the write (named): no such sheet, a visibility not read, hiding the last visible sheet.
 * SATISFIED: the sheet already reads this visibility. L1: its read visibility. L3: every cell, and every OTHER sheet's
 * visibility and the sheet list unchanged.
 */

const INTENTS = Object.freeze(['set_sheet_visibility']);
const FIELDS = Object.freeze(['intent', 'id', 'sheet', 'visibility']);
const STATES = Object.freeze({ visible: true, hidden: false });

function plan(step, index) {
  const at = `step ${index}: set_sheet_visibility`;
  const unknown = Object.keys(step).filter(k => !FIELDS.includes(k));
  if (unknown.length) return { ok: false, detail: `${at} does not support field(s): ${unknown.join(', ')}` };
  const why = sheetNameProblem(step.sheet);
  if (why) return { ok: false, detail: `${at}: ${why}` };
  if (!Object.hasOwn(STATES, step.visibility)) return { ok: false, detail: `${at}: visibility must be 'visible' or 'hidden' (read_sheet_state's own values)` };
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`,
      intent: 'set_sheet_visibility',
      family: 'sheet-visibility',
      selector: Object.freeze({ sheet: step.sheet, cellFree: true, claims: Object.freeze([`sheet-visibility:${step.sheet}`]) }),
      visibility: step.visibility,
      cells: Object.freeze([]),
      preservationMode: 'PRIMITIVE_INTEGRAL',
    },
  };
}

/** the sheet's read visibility and how many sheets read visible, or the named reason */
function gate(p, op) {
  if (!p || p.format !== 'xlsx') return { ok: false, reason: 'xlsx-observation-missing' };
  if (!p.complete) return { ok: false, reason: p.truncated ? 'xlsx-observation-truncated' : 'xlsx-observation-unreadable-getters', unsupported: p.unsupported };
  if (p.sheets.filter(x => x.name === op.selector.sheet).length !== 1) return { ok: false, reason: 'xlsx-sheet-not-found', available: p.sheets.map(x => x.name) };
  const vis = p.visibility;
  const bad = !vis ? ['not-read'] : p.sheets.map(x => x.name).filter(n => typeof vis[n] !== 'boolean');
  if (bad.length) return { ok: false, reason: 'xlsx-sheet-visibility-unobservable', sheets: bad.slice(0, 20) };
  return { ok: true, now: vis[op.selector.sheet], visibleCount: p.sheets.filter(x => vis[x.name] === true).length };
}

function classify(p, op) {
  const g = gate(p, op);
  if (!g.ok) return { state: 'FAIL_CLOSED', evidence: g };
  const want = STATES[op.visibility];
  const where = { sheet: op.selector.sheet, visibility: op.visibility };
  if (g.now === want) return { state: 'SATISFIED', evidence: { ...where, proof: 'fresh-live-GetVisible-reads-this-state' } };
  if (!want && g.visibleCount <= 1) return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-sheet-visibility-last-visible-sheet', ...where, visibleCount: g.visibleCount } };
  return { state: 'APPLY', evidence: where };
}

function resolve(p, selector, ctx = {}) {
  const g = gate(p, ctx.operation);
  if (!g.ok) return { ok: false, target: null, selector, ambiguity: g.reason, detail: g };
  return { ok: true, target: Object.freeze({ sheet: selector.sheet }) };
}

/** One callCommand: exact-name sheet; a hide re-checks that another sheet stays visible (the engine would not). */
function lower(op) {
  const J = JSON.stringify;
  const hide = !STATES[op.visibility];
  return {
    kind: 'callCommand',
    body: `
try {
  const ws = Api.GetSheet(${J(op.selector.sheet)});
  if (!ws || ws.GetName() !== ${J(op.selector.sheet)}) return { ok: false, stage: 'sheet-identity' };${hide ? `
  const others = (Api.GetSheets() || []).filter(s => s.GetName() !== ${J(op.selector.sheet)} && s.GetVisible() === true);
  if (others.length < 1) return { ok: false, stage: 'last-visible-sheet' };` : ''}
  ws.SetVisible(${!hide});
  return { ok: true };
} catch (err) { return { ok: false, stage: 'threw', error: String(err) }; }
`,
    meta: Object.freeze({ family: 'sheet-visibility', sheet: op.selector.sheet, visibility: op.visibility }),
  };
}

function verifyPrimitive(op, before, after) {
  const g = gate(after && after.primary, op);
  if (!g.ok) return { ok: false, evidence: { reason: g.reason } };
  return { ok: g.now === STATES[op.visibility], evidence: { sheet: op.selector.sheet, now: g.now ? 'visible' : 'hidden', want: op.visibility, proof: 'fresh-live-GetVisible' } };
}

/** L3: every cell; the sheet list and every OTHER sheet's visibility unchanged */
function verifyPreservation(op, before, after) {
  const rest = cells.verifyPreservation(op, before, after);
  if (!rest.ok && rest.evidence && rest.evidence.reason) return rest;
  const b = before.primary;
  const a = after.primary;
  const others = p => (p && p.visibility ? p.sheets.filter(x => x.name !== op.selector.sheet).map(x => [x.name, x.index, p.visibility[x.name]]) : null);
  const kept = !!others(b) && JSON.stringify(others(b)) === JSON.stringify(others(a));
  return { ok: rest.ok && kept, evidence: { ...rest.evidence, otherSheetsKept: kept } };
}

function wholeTaskChecks(op, finalObservation) {
  const v = verifyPrimitive(op, null, finalObservation);
  return [{ name: `${op.selector.sheet}: ${op.visibility}`, ok: v.ok, detail: v.ok ? null : v.evidence }];
}

const windowsOf = () => [];
const liveExtras = () => ({ sheetVisibility: true });

module.exports = { INTENTS, plan, classify, resolve, lower, verifyPrimitive, verifyPreservation, wholeTaskChecks, windowsOf, liveExtras };
