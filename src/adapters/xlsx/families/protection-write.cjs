'use strict';

const { parseRange, rangeText, cellsOf } = require('../a1.cjs');
const { XLSX_B27_BRIDGE_WRITE_RUNTIME_IDS } = require('../runtime.cjs');
const cells = require('./cells.cjs');

/**
 * XLSX `protection` write family (GH #338 P01a protect, #339 P01b unprotect, #340 P02a cell locked, #341 P02b formula hidden,
 * #342 P03 allowed sheet operations): protect_sheet {sheet, forbids?} | unprotect_sheet {sheet} | set_cell_protection {sheet, range,
 * locked?, hidden?}. Each is the ONLY step of its task. PASSWORDLESS only -- the product never sets, reads or removes a password.
 *
 * ENGINE PATH: ONE Bridge WRITE (DS Bridge /22 ONLY; design docs/DS-9.3.5-BRIDGE22-SHEET-PROTECTION-WRITE-DESIGN.md) --
 * sheet.protection.write/v1 (the engine model's setProtectedSheet, one history point) or sheet.cellProtection.write/v1 (Range.setLocked /
 * setHiddenFormulas). The public API has no protection setter (#299 PR-B01); the editor's asc_setProtectedSheet blocks the session
 * (#299 run 2). On /16../21 the capability does not exist: the plan is FAIL_CLOSED (named, no dispatch).
 *
 * READ <-> WRITE mapping: `forbids` is EXACTLY sheet.protection.read/v1's `forbids` (the same 15 keys, the same meaning: true = the
 * protection forbids the operation); an unnamed flag takes the Excel default. The verifier is that read (`protectionReads` /
 * `protectionRangeReads`, the #326 guard's own axes, fetched for every op by protection-gate.targetsOf).
 * Cell protection: the read lists the NON-EMPTY cells only (an empty cell is the Excel default: locked, not hidden), so the target is
 * judged per cell of the range against `cells.unlocked` / `cells.hidden`.
 *
 * Refused BEFORE the write, named: set_cell_protection on a PROTECTED sheet (the #326 guard, xlsx-protected-sheet-forbids:cellProtection
 * -- Excel refuses it too, and it could turn #327's redaction off); a credential-protected sheet is the HANDLER's refusal
 * (BRIDGE:credential-protected, mutation NONE) -- the read never exposes a credential, so the product cannot tell it.
 * RETRY: the fresh read already holding the target -> SATISFIED (0 write), 'protection-already-at-the-target'.
 *
 * GH #366 2nd unit (docs/GH366-LARGE-BOOK-FAMILY-PARITY-SOURCE-CONTRACT-20261105.md): `gate()` used to read
 * ONLY `sheets`/`protectionReads`/`protectionRangeReads` and never looked at the observation's own
 * `complete`/`truncated` flag -- unlike `delete-cells.cjs`'s `gate()`, which refuses named
 * (`xlsx-observation-truncated`) above the read bound BEFORE dispatch. A LIVE measurement showed that this
 * family, above the same bound, DISPATCHES anyway and leaves `DISPATCHED_OUTCOME_UNKNOWN` /
 * `INTEGRAL_PRESERVATION_NOT_PROVEN` -- the hypothesis that it "already worked safely on large books" was
 * REFUTED. `gate()` now refuses PRE-dispatch above the bound, the same named reason, the same mechanism.
 */

const INTENTS = Object.freeze(['protect_sheet', 'unprotect_sheet', 'set_cell_protection']);
const FLAGS = Object.freeze(['objects', 'scenarios', 'formatCells', 'formatColumns', 'formatRows', 'insertColumns', 'insertRows', 'insertHyperlinks',
  'deleteColumns', 'deleteRows', 'selectLockedCells', 'sort', 'autoFilter', 'pivotTables', 'selectUnlockedCells']);
/** the Excel default of each flag (true = forbidden), the engine's setDefaultInterface() */
const DEFAULT_FORBIDS = Object.freeze(Object.fromEntries(FLAGS.map(f => [f, !(f === 'selectLockedCells' || f === 'selectUnlockedCells')])));
const FIELDS = Object.freeze({ protect_sheet: ['forbids'], unprotect_sheet: [], set_cell_protection: ['range', 'locked', 'hidden'] });
const CAPABILITY = Object.freeze({ protect_sheet: 'sheet.protection.write/v1', unprotect_sheet: 'sheet.protection.write/v1', set_cell_protection: 'sheet.cellProtection.write/v1' });
const B27_BRIDGE_WRITE_RUNTIMES = XLSX_B27_BRIDGE_WRITE_RUNTIME_IDS;
// the same ceiling as write_range (every cell of the range is judged); the handler's own cap is 10000
const MAX_CELLS = 1000;
const NAMED = 'protection-already-at-the-target';

/** protect_sheet: the canonical forbids (every flag, the read order, the Excel default for the unnamed), or the named problem */
function forbidsOf(step, at) {
  const fb = step.forbids === undefined ? {} : step.forbids;
  if (!fb || typeof fb !== 'object' || Array.isArray(fb)) return { detail: `${at} forbids must be an object of flag: boolean` };
  const bad = Object.keys(fb).filter(k => !FLAGS.includes(k));
  if (bad.length) return { detail: `${at} forbids has unknown flag(s): ${bad.join(', ')} (one of ${FLAGS.join(', ')})` };
  const nb = Object.keys(fb).find(k => typeof fb[k] !== 'boolean');
  if (nb) return { detail: `${at} forbids.${nb} must be a boolean` };
  const has = f => Object.prototype.hasOwnProperty.call(fb, f);
  return { forbids: Object.freeze(Object.fromEntries(FLAGS.map(f => [f, has(f) ? fb[f] : DEFAULT_FORBIDS[f]]))) };
}

/** set_cell_protection: the normalised range and the locked / hidden target, or the named problem */
function cellTargetOf(step, at) {
  const g = parseRange(step.range);
  if (!g) return { detail: `${at} requires a plain A1 range` };
  if (g.count > MAX_CELLS) return { detail: `${at} range has ${g.count} cells, over the ${MAX_CELLS} ceiling` };
  const keys = ['locked', 'hidden'].filter(k => step[k] !== undefined);
  if (!keys.length) return { detail: `${at} requires locked and / or hidden` };
  const nb = keys.find(k => typeof step[k] !== 'boolean');
  if (nb) return { detail: `${at} ${nb} must be a boolean` };
  return { range: rangeText(g), want: Object.fromEntries(keys.map(k => [k, step[k]])) };
}

function plan(step, index) {
  const intent = step.intent;
  const at = `step ${index}: ${intent}`;
  if (!FIELDS[intent]) return { ok: false, detail: `${at} is not a protection intent` };
  const fields = ['intent', 'id', 'sheet', ...FIELDS[intent]];
  const unknown = Object.keys(step).filter(k => !fields.includes(k));
  if (unknown.length) return { ok: false, detail: `${at} does not support field(s): ${unknown.join(', ')}` };
  if (typeof step.sheet !== 'string' || step.sheet === '') return { ok: false, detail: `${at} requires a sheet name` };
  let want = {};
  let range = null;
  if (intent === 'protect_sheet') {
    const f = forbidsOf(step, at);
    if (f.detail) return { ok: false, detail: f.detail };
    want = { forbids: f.forbids };
  } else if (intent === 'set_cell_protection') {
    const c = cellTargetOf(step, at);
    if (c.detail) return { ok: false, detail: c.detail };
    ({ range, want } = c);
  }
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`,
      intent,
      family: 'protection',
      selector: Object.freeze(range ? { sheet: step.sheet, range } : { sheet: step.sheet }),
      want: Object.freeze(want),
      // protection owns no cell content: every value and formula is preserved, not claimed
      cells: Object.freeze([]),
      preservationMode: 'PRIMITIVE_INTEGRAL',
    },
  };
}

/** the sheet's protection read (sheet-level, or the range read for set_cell_protection), or the named reason it cannot be told */
function gate(p, op) {
  if (!p || p.format !== 'xlsx') return { ok: false, reason: 'xlsx-observation-missing' };
  // GH #366 2nd unit: this gate never looked at primary.complete/truncated at all (unlike delete-cells.cjs:127) --
  // LIVE measured (gh366 v3 r2, docs/GH366-LARGE-BOOK-FAMILY-PARITY-SOURCE-CONTRACT-20261105.md) that above
  // the read bound the family DISPATCHES anyway and leaves DISPATCHED_OUTCOME_UNKNOWN / INTEGRAL_PRESERVATION_NOT_PROVEN
  // instead of a named refusal -- the same safety limit every other family already enforces PRE-dispatch.
  if (!p.complete) return { ok: false, reason: p.truncated ? 'xlsx-observation-truncated' : 'xlsx-observation-unreadable-getters', unsupported: p.unsupported };
  const s = (p.sheets || []).filter(x => x.name === op.selector.sheet);
  if (s.length !== 1) return { ok: false, reason: s.length > 1 ? 'xlsx-sheet-ambiguous' : 'xlsx-sheet-not-found', available: (p.sheets || []).map(x => x.name) };
  const ranged = op.intent === 'set_cell_protection';
  const read = ranged ? (p.protectionRangeReads || {})[`${op.selector.sheet}!${op.selector.range}`] : (p.protectionReads || {})[op.selector.sheet];
  if (read === undefined) return { ok: false, reason: 'xlsx-protection-unobservable', why: 'not-read' };
  if (typeof read === 'string') return { ok: false, reason: 'xlsx-protection-unobservable', why: read };
  if (!read || typeof read.protected !== 'boolean') return { ok: false, reason: 'xlsx-protection-unobservable', why: 'malformed-read' };
  if (read.protected && (!read.forbids || FLAGS.some(f => typeof read.forbids[f] !== 'boolean'))) return { ok: false, reason: 'xlsx-protection-unobservable', why: 'malformed-forbids' };
  if (ranged && (!read.cells || !Array.isArray(read.cells.unlocked) || !Array.isArray(read.cells.hidden))) return { ok: false, reason: 'xlsx-protection-unobservable', why: 'cells-not-read' };
  return { ok: true, read };
}

/** the cells of the range NOT holding the target (an unlisted cell = the Excel default: locked, not hidden) */
function cellsOff(read, op) {
  const unlocked = new Set(read.cells.unlocked);
  const hidden = new Set(read.cells.hidden);
  const off = [];
  for (const a of cellsOf(parseRange(op.selector.range))) {
    if (op.want.locked !== undefined && (!unlocked.has(a)) !== op.want.locked) off.push(a);
    else if (op.want.hidden !== undefined && hidden.has(a) !== op.want.hidden) off.push(a);
  }
  return off;
}

/** does the read hold the op's target? */
function holds(read, op) {
  if (op.intent === 'unprotect_sheet') return read.protected === false;
  if (op.intent === 'protect_sheet') return read.protected === true && FLAGS.every(f => read.forbids[f] === op.want.forbids[f]);
  return cellsOff(read, op).length === 0;
}

function classify(p, op) {
  const runtime = p && p.runtime ? p.runtime.identity : null;
  const where = { ...op.selector };
  if (!B27_BRIDGE_WRITE_RUNTIMES.includes(runtime)) {
    return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-protection-bridge-write-runtime-not-proven', capability: 'CAPABILITY_MISSING', ...where, runtime, proven: B27_BRIDGE_WRITE_RUNTIMES } };
  }
  const g = gate(p, op);
  if (!g.ok) return { state: 'FAIL_CLOSED', evidence: { ...where, ...g } };
  if (holds(g.read, op)) return { state: 'SATISFIED', evidence: { ...where, observation: NAMED, proof: 'fresh-live-sheet.protection.read' } };
  if (op.intent === 'set_cell_protection') {
    if (g.read.protected) return { state: 'FAIL_CLOSED', evidence: { ...where, reason: 'xlsx-protected-sheet-forbids:cellProtection' } };
    return { state: 'APPLY', evidence: { ...where, want: op.want, cells: cellsOff(g.read, op).slice(0, 20) } };
  }
  return { state: 'APPLY', evidence: { ...where, from: { protected: g.read.protected, forbids: g.read.forbids }, to: op.intent === 'protect_sheet' ? { protected: true, forbids: op.want.forbids } : { protected: false } } };
}

function resolve(p, selector, ctx = {}) {
  const g = gate(p, ctx.operation);
  if (!g.ok) return { ok: false, target: null, selector, ambiguity: g.reason, detail: g };
  return { ok: true, target: Object.freeze({ ...selector }) };
}

/** One callCommand: the exact-name sheet check, then ONE Bridge write (the driver prelude supplies the gate / write / fail helpers). */
function lower(op) {
  const sheet = op.selector.sheet;
  let args;
  let requested;
  if (op.intent === 'protect_sheet') { args = { sheet, protect: true, forbids: op.want.forbids }; requested = ['protect']; } else if (op.intent === 'unprotect_sheet') { args = { sheet, protect: false }; requested = ['unprotect']; } else {
    args = { sheet, range: op.selector.range };
    requested = [];
    if (op.want.locked !== undefined) { args.locked = op.want.locked; requested.push('locked'); }
    if (op.want.hidden !== undefined) { args.hidden = op.want.hidden; requested.push('hidden'); }
  }
  return {
    kind: 'callCommand',
    body: `
try {
  const name = ${JSON.stringify(sheet)};
  const ws = Api.GetSheet(name);
  if (!ws || ws.GetName() !== name) return { ok: false, stage: 'sheet-identity' };
  const g = __eoProtGate();
  if (!g.ok) return g;
  const w = __eoProtWrite(${JSON.stringify(args)}, ${JSON.stringify(requested)});
  if (!w.ok) return __eoProtFail(w);
  return { ok: true, satisfied: w.satisfied === true };
} catch (err) { return { ok: false, stage: 'threw', error: String(err) }; }
`,
    meta: Object.freeze({ family: 'protection', intent: op.intent, sheet, ...(op.selector.range ? { range: op.selector.range } : {}), want: op.want, bridgeProtectionWrite: CAPABILITY[op.intent] }),
  };
}

function verifyPrimitive(op, before, after) {
  const g = gate(after && after.primary, op);
  if (!g.ok) return { ok: false, evidence: { reason: g.reason, why: g.why || null } };
  const ok = holds(g.read, op);
  const observed = op.intent === 'set_cell_protection' ? { protected: g.read.protected, unlocked: g.read.cells.unlocked.slice(0, 50), hidden: g.read.cells.hidden.slice(0, 50) } : { protected: g.read.protected, forbids: g.read.forbids };
  return { ok, evidence: { ...op.selector, expected: op.want, observed, ...(ok || op.intent !== 'set_cell_protection' ? {} : { cellsOff: cellsOff(g.read, op).slice(0, 20) }), proof: 'fresh-live-sheet.protection.read' } };
}

/** every cell's value and formula unchanged (protection owns none) */
function verifyPreservation(op, before, after) {
  return cells.verifyPreservation(op, before, after);
}

function wholeTaskChecks(op, finalObservation) {
  const g = gate(finalObservation && finalObservation.primary, op);
  const ok = g.ok && holds(g.read, op);
  const label = op.intent === 'protect_sheet' ? 'is protected with the requested forbids' : op.intent === 'unprotect_sheet' ? 'is not protected' : `${op.selector.range} holds ${JSON.stringify(op.want)}`;
  return [{ name: `${op.selector.sheet} ${label}`, ok, detail: ok ? null : { reason: g.ok ? 'differs' : g.reason } }];
}

const windowsOf = op => (op.selector.range ? [{ sheet: op.selector.sheet, range: op.selector.range }] : []);
const liveExtras = () => ({});

module.exports = { INTENTS, FLAGS, DEFAULT_FORBIDS, CAPABILITY, B27_BRIDGE_WRITE_RUNTIMES, MAX_CELLS, plan, classify, resolve, lower, verifyPrimitive, verifyPreservation, wholeTaskChecks, windowsOf, liveExtras, holds };
