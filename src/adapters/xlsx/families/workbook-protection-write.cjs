'use strict';

const { XLSX_B28_BRIDGE_WRITE_RUNTIME_IDS } = require('../runtime.cjs');
const cells = require('./cells.cjs');

/**
 * XLSX `workbook-protection` write family (GH #338-#342 remainder, W3, docs/DS-9.3.5-BRIDGE24-WORKBOOK-PROTECTION-DESIGN.md):
 * protect_workbook {} | unprotect_workbook {}. Each is the ONLY step of its task (planOperations(), the SAME mechanism the
 * sheet-level 'protection' family already uses) -- the design's own "Sole-step rule" closes the review-blocking W2/W3
 * interaction this way: a task mixing either intent with ANYTHING else (a W2-gated structural intent, once that later unit
 * lands, or any other op) is refused whole, at planning, before a single step of it dispatches.
 *
 * ENGINE PATH: ONE Bridge WRITE (DS Bridge /24 ONLY) -- workbook.protection.write/v1 (the engine model's
 * setProtectedWorkbook, one history point, `asc_setLockStructure(protect)` ONLY -- no password member touched, the
 * SAME passwordless-only shape as /22's sheet.protection.write/v1). On /16../23 the capability does not exist: the
 * plan is FAIL_CLOSED (named, no dispatch).
 *
 * READ <-> WRITE mapping: the workbook's `lockStructure`, read through workbook.protection.read/v1 (W1,
 * protection-read.cjs's additive `workbook` field). RETRY: the fresh read already holding the target -> SATISFIED
 * (0 write), 'workbook-protection-already-at-the-target'.
 *
 * Review-blocking security defect, fixed in the Bridge candidate (62133fca), not here: `unprotect_workbook` on a
 * CURRENTLY LOCKED workbook is refused `credential-unobservable` by the handler itself, fail-closed, BEFORE any
 * history point or setter call -- the fix does not need a workbook-level credential detector to exist (engine
 * finding 4 left it unsourced) because it fires on "is currently locked", not "carries a credential". This family
 * adds the SAME refusal pre-dispatch, named here too, the moment the plan-time read already shows locked -- a
 * redundant, cheap short-circuit in front of the Bridge's own (authoritative) check, not a replacement for it.
 */

const INTENTS = Object.freeze(['protect_workbook', 'unprotect_workbook']);
const CAPABILITY = Object.freeze({ protect_workbook: 'workbook.protection.write/v1', unprotect_workbook: 'workbook.protection.write/v1' });
const B28_BRIDGE_WRITE_RUNTIMES = XLSX_B28_BRIDGE_WRITE_RUNTIME_IDS;
const NAMED = 'workbook-protection-already-at-the-target';
const SOLE_STEP_REASON = 'xlsx-workbook-protection-change-must-be-sole-step';

function plan(step, index) {
  const intent = step.intent;
  const at = `step ${index}: ${intent}`;
  if (intent !== 'protect_workbook' && intent !== 'unprotect_workbook') return { ok: false, detail: `${at} is not a workbook-protection intent` };
  const unknown = Object.keys(step).filter(k => k !== 'intent' && k !== 'id');
  if (unknown.length) return { ok: false, detail: `${at} does not support field(s): ${unknown.join(', ')}` };
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`,
      intent,
      family: 'workbook-protection',
      // workbook-scoped, no sheet and no cell footprint of its own -- the same `cellFree` convention
      // page.cjs / chart-point-labels.cjs use for a sheet-level, cell-free claim, one level up
      selector: Object.freeze({ sheet: '\u0000workbook', range: 'A1', cellFree: true, claims: Object.freeze(['workbook-protection']) }),
      want: Object.freeze({ protect: intent === 'protect_workbook' }),
      // owns no cell content: every value and formula is preserved, not claimed
      cells: Object.freeze([]),
      preservationMode: 'PRIMITIVE_INTEGRAL',
    },
  };
}

/** the workbook's current lockStructure read, or the named reason it cannot be told */
function gate(p) {
  if (!p || p.format !== 'xlsx') return { ok: false, reason: 'xlsx-observation-missing' };
  // GH #366 4th unit: this gate never looked at primary.complete/truncated at all (the same gap
  // protect_sheet had, protection-write.cjs, 2nd unit) -- LIVE measured (gh366 (A) group r1) that
  // above the read bound the family DISPATCHES anyway and leaves DISPATCHED_OUTCOME_UNKNOWN /
  // INTEGRAL_PRESERVATION_NOT_PROVEN instead of a named refusal -- the same safety limit every
  // other family already enforces PRE-dispatch.
  if (!p.complete) return { ok: false, reason: p.truncated ? 'xlsx-observation-truncated' : 'xlsx-observation-unreadable-getters', unsupported: p.unsupported };
  const read = p.workbookProtectionRead;
  if (read === undefined || read === null) return { ok: false, reason: 'xlsx-workbook-protection-unobservable', why: 'not-read' };
  if (typeof read === 'string') return { ok: false, reason: 'xlsx-workbook-protection-unobservable', why: read };
  if (!read || typeof read.lockStructure !== 'boolean') return { ok: false, reason: 'xlsx-workbook-protection-unobservable', why: 'malformed-read' };
  return { ok: true, read };
}

function classify(p, op) {
  const runtime = p && p.runtime ? p.runtime.identity : null;
  if (!B28_BRIDGE_WRITE_RUNTIMES.includes(runtime)) {
    return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-workbook-protection-bridge-write-runtime-not-proven', capability: 'CAPABILITY_MISSING', runtime, proven: B28_BRIDGE_WRITE_RUNTIMES } };
  }
  const g = gate(p);
  if (!g.ok) return { state: 'FAIL_CLOSED', evidence: g };
  if (g.read.lockStructure === op.want.protect) return { state: 'SATISFIED', evidence: { observation: NAMED, proof: 'fresh-live-workbook.protection.read' } };
  if (op.want.protect === false && g.read.lockStructure === true) {
    return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-workbook-unprotect-credential-unobservable' } };
  }
  return { state: 'APPLY', evidence: { from: { lockStructure: g.read.lockStructure }, to: { lockStructure: op.want.protect } } };
}

function resolve(p, selector) {
  const g = gate(p);
  if (!g.ok) return { ok: false, target: null, selector, ambiguity: g.reason, detail: g };
  return { ok: true, target: Object.freeze({ ...selector }) };
}

/** One callCommand: the ONE Bridge write (the driver prelude supplies the gate / write / fail helpers). */
function lower(op) {
  const args = { protect: op.want.protect };
  const requested = ['lockStructure'];
  return {
    kind: 'callCommand',
    body: `
try {
  const g = __eoWbProtGate();
  if (!g.ok) return g;
  const w = __eoWbProtWrite(${JSON.stringify(args)}, ${JSON.stringify(requested)});
  if (!w.ok) return __eoWbProtFail(w);
  return { ok: true, satisfied: w.satisfied === true };
} catch (err) { return { ok: false, stage: 'threw', error: String(err) }; }
`,
    meta: Object.freeze({ family: 'workbook-protection', intent: op.intent, want: op.want, bridgeWorkbookProtectionWrite: true }),
  };
}

function verifyPrimitive(op, before, after) {
  const g = gate(after && after.primary);
  if (!g.ok) return { ok: false, evidence: { reason: g.reason, why: g.why || null } };
  const ok = g.read.lockStructure === op.want.protect;
  return { ok, evidence: { expected: op.want, observed: { lockStructure: g.read.lockStructure }, proof: 'fresh-live-workbook.protection.read' } };
}

/** every cell's value and formula unchanged (workbook-protection owns none) */
function verifyPreservation(op, before, after) {
  return cells.verifyPreservation(op, before, after);
}

function wholeTaskChecks(op, finalObservation) {
  const g = gate(finalObservation && finalObservation.primary);
  const ok = g.ok && g.read.lockStructure === op.want.protect;
  const label = op.intent === 'protect_workbook' ? 'is protected (lockStructure)' : 'is not protected (lockStructure)';
  return [{ name: `workbook ${label}`, ok, detail: ok ? null : { reason: g.ok ? 'differs' : g.reason } }];
}

const windowsOf = () => [];
const liveExtras = () => ({ workbookProtection: true });

module.exports = { INTENTS, CAPABILITY, B28_BRIDGE_WRITE_RUNTIMES, SOLE_STEP_REASON, plan, classify, resolve, lower, verifyPrimitive, verifyPreservation, wholeTaskChecks, windowsOf, liveExtras };
