'use strict';

const { parseRange, rangeText } = require('../a1.cjs');
const cells = require('./cells.cjs');

/**
 * XLSX `validation` family: set_validation / clear_validation (M6 wave 6).
 *
 * Two readers, both already in the driver: the PUBLIC range-level GetValidation() getters (the source's measured
 * reader, B-D16) for the rule's parameters, and the W3-proven, build-gated INTERNAL dataValidations reader for the
 * rule's exact RANGES (the public reader answers for a range, not per cell -- a partial rule is invisible to it).
 * Audit: B-D16 / D25 MIGRATE (whole number, between, stop); B-D17..D22 NOT_MEASURED; B-D23 an existing DIFFERENT
 * rule was silently overwritten through Modify.
 *
 * Contract (the W6 probe measured the public getters; nothing wider is supported): set supports the measured shape only -- type 'wholeNumber',
 * operator 'between', alertStyle 'stop', integer bounds. Refused before the write: a validation that overlaps the
 * target PARTIALLY or covers it with another range, and an existing rule on exactly the target with other
 * parameters (no silent Modify); an unobservable reader. clear: SATISFIED when no rule touches the target; refused
 * when a rule reaches beyond it (Delete would clear cells outside the request -- unmeasured).
 * L1 the public read + exactly one internal rule on exactly the range (set) / none touching it (clear);
 * L3 every other rule of every sheet and every cell. Retry SATISFIED 0/0.
 *
 * M7 R17 (owner M7 GO 2026-10-03; tier D probe docs/evidence/M7-D-PROBE.md, every getter read back exactly in
 * session and after a reopen): type wholeNumber | decimal | textLength with operator between | notBetween (min, max)
 * or greater | lessEqual (value); type list with `items` (inline, no comma / quote in an item) or `listRange` (an
 * absolute range on the same sheet); alertStyle stop | warning | information. Still refused (measured or not
 * measured): date (formula1 reads back as locale text), custom (the formula is re-anchored to A1 -- its meaning
 * changes), time, the other operators, input / error messages and the ignore-blank / dropdown flags.
 */

const INTENTS = Object.freeze(['set_validation', 'clear_validation']);
const MAX_CELLS = 1000;

function plan(step, index) {
  const at = `step ${index}: ${step.intent}`;
  const set = step.intent === 'set_validation';
  const fields = set ? ['intent', 'id', 'sheet', 'range', 'type', 'operator', 'alertStyle', 'min', 'max', 'value', 'items', 'listRange'] : ['intent', 'id', 'sheet', 'range'];
  const unknown = Object.keys(step).filter(k => !fields.includes(k));
  if (unknown.length) return { ok: false, detail: `${at} does not support field(s): ${unknown.join(', ')}` };
  if (typeof step.sheet !== 'string' || step.sheet === '') return { ok: false, detail: `${at} requires a sheet name` };
  const g = parseRange(step.range);
  if (!g) return { ok: false, detail: `${at} requires a plain A1 range` };
  if (g.count > MAX_CELLS) return { ok: false, detail: `${at} range has ${g.count} cells, over the ${MAX_CELLS} ceiling` };
  let want = null;
  if (set) {
    want = ruleOf(step);
    if (want.problem) return { ok: false, detail: `${at}: ${want.problem}` };
  }
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`,
      intent: step.intent,
      family: 'validation',
      selector: Object.freeze({ sheet: step.sheet, range: rangeText(g) }),
      want: want ? Object.freeze(want) : null,
      cells: Object.freeze([]),
      preservationMode: 'PRIMITIVE_INTEGRAL',
    },
  };
}

const TYPES = Object.freeze({ wholeNumber: 'xlValidateWholeNumber', decimal: 'xlValidateDecimal', textLength: 'xlValidateTextLength', list: 'xlValidateList' });
const ALERTS = Object.freeze({ stop: 'xlValidAlertStop', warning: 'xlValidAlertWarning', information: 'xlValidAlertInformation' });
const TWO = Object.freeze({ between: 'xlBetween', notBetween: 'xlNotBetween' });
const ONE = Object.freeze({ greater: 'xlGreater', lessEqual: 'xlLessEqual' });
const plainNumber = n => typeof n === 'number' && Number.isFinite(n) && !/e/i.test(String(n));

/** The rule a set_validation step asks for, in the editor's terms -- or { problem }. */
function ruleOf(step) {
  const type = TYPES[step.type];
  const alertStyle = ALERTS[step.alertStyle];
  if (!type) return { problem: `type must be one of ${Object.keys(TYPES).join(' | ')} (date / time / custom are not supported: measured or not measured)` };
  if (!alertStyle) return { problem: `alertStyle must be one of ${Object.keys(ALERTS).join(' | ')}` };
  if (step.type === 'list') return listRule(step, type, alertStyle);
  for (const k of ['items', 'listRange']) if (k in step) return { problem: `${k} belongs to type list` };
  return boundsRule(step, type, alertStyle);
}

/** wholeNumber / decimal / textLength: two bounds (between, notBetween) or one value (greater, lessEqual). */
function boundsRule(step, type, alertStyle) {
  const isNum = n => (step.type === 'decimal' ? plainNumber(n) : Number.isInteger(n)) && (step.type !== 'textLength' || n >= 0);
  const what = step.type === 'decimal' ? 'plain finite numbers' : (step.type === 'textLength' ? 'non-negative integers' : 'integers');
  if (TWO[step.operator]) {
    if ('value' in step) return { problem: `operator ${step.operator} takes min and max, not value` };
    if (!isNum(step.min) || !isNum(step.max) || step.min > step.max) return { problem: `min and max must be ${what}, min <= max` };
    return { type, alertStyle, operator: TWO[step.operator], formula1: String(step.min), formula2: String(step.max) };
  }
  if (ONE[step.operator]) {
    if ('min' in step || 'max' in step) return { problem: `operator ${step.operator} takes value, not min / max` };
    if (!isNum(step.value)) return { problem: `value must be one of the ${what}` };
    return { type, alertStyle, operator: ONE[step.operator], formula1: String(step.value), formula2: '' };
  }
  return { problem: `operator must be one of ${[...Object.keys(TWO), ...Object.keys(ONE)].join(' | ')} (the others are not measured)` };
}

function listRule(step, type, alertStyle) {
  for (const k of ['operator', 'min', 'max', 'value']) if (k in step) return { problem: `type list does not take ${k}` };
  if (('items' in step) === ('listRange' in step)) return { problem: 'type list takes exactly one of items | listRange' };
  if ('items' in step) {
    const ok = Array.isArray(step.items) && step.items.length > 0 && step.items.length <= 64 && step.items.every(x => typeof x === 'string' && x !== '' && x === x.trim() && !/[,"]/.test(x));
    if (!ok) return { problem: 'items must be 1..64 non-empty strings without a comma, a double quote or outer spaces' };
    return { type, alertStyle, operator: 'xlBetween', formula1: `"${step.items.join(',')}"`, formula2: '' };
  }
  const g = parseRange(step.listRange);
  if (!g) return { problem: 'listRange must be a plain A1 range on the same sheet' };
  const abs = `$${rangeText(g).replace(/^([A-Z]+)(\d+)/, '$1$$$2').replace(/:([A-Z]+)(\d+)$/, ':$$$1$$$2')}`;
  return { type, alertStyle, operator: 'xlBetween', formula1: `=${abs}`, formula2: '' };
}

const keyOf = op => `${op.selector.sheet}!${op.selector.range}`;
const overlaps = (a, b) => a.c1 <= b.c2 && b.c1 <= a.c2 && a.r1 <= b.r2 && b.r1 <= a.r2;

/** The internal validation rules of the target sheet that touch the target: exact / other. */
function touching(p, op) {
  const t = parseRange(op.selector.range);
  const rules = (p.rules.bySheet[op.selector.sheet] || []).filter(r => r.kind === 'validation');
  const exact = [];
  const other = [];
  for (const r of rules) {
    const parts = String(r.range).split(',').map(x => parseRange(x.trim()));
    if (!parts.some(x => x && overlaps(x, t))) continue;
    if (parts.length === 1 && parts[0] && rangeText(parts[0]) === op.selector.range) exact.push(r); else other.push(r);
  }
  return { exact, other };
}

const bare = f => String(f === undefined || f === null ? '' : f).replace(/^=/, '');
const sameRule = (read, want) => !!read && !read.absent && ['type', 'alertStyle', 'operator'].every(k => read[k] === want[k]) && bare(read.formula1) === bare(want.formula1) && bare(read.formula2) === bare(want.formula2);

function gate(p, op) {
  if (!p || p.format !== 'xlsx') return { ok: false, reason: 'xlsx-observation-missing' };
  if (!p.complete) return { ok: false, reason: p.truncated ? 'xlsx-observation-truncated' : 'xlsx-observation-unreadable-getters', unsupported: p.unsupported };
  const s = p.sheets.filter(x => x.name === op.selector.sheet);
  if (s.length !== 1) return { ok: false, reason: s.length > 1 ? 'xlsx-sheet-ambiguous' : 'xlsx-sheet-not-found', available: p.sheets.map(x => x.name) };
  if (!p.rules || !p.rules.complete || !p.rules.bySheet) return { ok: false, reason: 'xlsx-validation-rules-unobservable', unobservable: p.rules ? p.rules.unobservable : 'not-read' };
  const read = p.validationReads ? p.validationReads[keyOf(op)] : undefined;
  if (read === undefined || typeof read === 'string') return { ok: false, reason: 'xlsx-validation-read-unobservable', why: read === undefined ? 'not-read' : read };
  return { ok: true, read, ...touching(p, op) };
}

function classify(p, op) {
  const g = gate(p, op);
  if (!g.ok) return { state: 'FAIL_CLOSED', evidence: g };
  const where = { sheet: op.selector.sheet, range: op.selector.range };
  if (g.other.length) return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-validation-overlaps-another-range', rules: g.other.map(r => r.range).slice(0, 20) } };
  if (op.intent === 'clear_validation') {
    if (!g.exact.length && g.read.absent) return { state: 'SATISFIED', evidence: { ...where, proof: 'fresh-live-no-validation' } };
    return { state: 'APPLY', evidence: where };
  }
  if (g.exact.length > 1) return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-validation-ambiguous', count: g.exact.length } };
  if (g.exact.length === 1) {
    return sameRule(g.read, op.want) ? { state: 'SATISFIED', evidence: { ...where, proof: 'fresh-live-the-rule-on-exactly-the-range' } }
      : { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-validation-exists-different', now: g.read } };
  }
  return { state: 'APPLY', evidence: where };
}

function resolve(p, selector, ctx = {}) {
  const g = gate(p, ctx.operation);
  if (!g.ok) return { ok: false, target: null, selector, ambiguity: g.reason, detail: g };
  return { ok: true, target: Object.freeze({ sheet: selector.sheet, range: selector.range }) };
}

function lower(op) {
  const { sheet, range } = op.selector;
  const w = op.want;
  // a single-formula rule is added with four arguments (the measured call); a two-formula rule with five (M6 body)
  const args = w ? [w.type, w.alertStyle, w.operator, w.formula1, ...(w.formula2 === '' ? [] : [w.formula2])].map(x => JSON.stringify(x)).join(', ') : '';
  const call = op.intent === 'set_validation' ? `r.GetValidation().Add(${args});` : 'r.GetValidation().Delete();';
  return {
    kind: 'callCommand',
    body: `
try {
  const name = ${JSON.stringify(sheet)};
  const ws = Api.GetSheet(name);
  if (!ws || ws.GetName() !== name) return { ok: false, stage: 'sheet-identity' };
  const r = ws.GetRange(${JSON.stringify(range)});
  ${call}
  return { ok: true };
} catch (err) { return { ok: false, stage: 'threw', error: String(err) }; }
`,
    meta: Object.freeze({ family: 'validation', action: op.intent, sheet, range, want: w }),
  };
}

function verifyPrimitive(op, _before, after) {
  const g = gate(after && after.primary, op);
  if (!g.ok) return { ok: false, evidence: g };
  const ok = op.intent === 'set_validation' ? g.exact.length === 1 && g.other.length === 0 && sameRule(g.read, op.want) : g.exact.length === 0 && g.other.length === 0 && g.read.absent === true;
  return { ok, evidence: { sheet: op.selector.sheet, range: op.selector.range, read: g.read, exact: g.exact.length, other: g.other.length, proof: 'fresh-live-public-read-and-internal-ranges' } };
}

/** L3: every cell, and every rule other than the target's validation unchanged (all sheets). */
function verifyPreservation(op, before, after) {
  const rest = cells.verifyPreservation(op, before, after);
  if (!rest.ok && rest.evidence && rest.evidence.reason) return rest;
  const strip = p => Object.fromEntries(Object.entries(p.rules.bySheet || {}).map(([n, rs]) => [n, rs.filter(r => !(n === op.selector.sheet && r.kind === 'validation' && r.range === op.selector.range))]));
  const b = before.primary;
  const a = after.primary;
  const rulesKept = !!(b.rules && a.rules && b.rules.bySheet && a.rules.bySheet) && JSON.stringify(strip(b)) === JSON.stringify(strip(a));
  return { ok: rest.ok && rulesKept, evidence: { ...rest.evidence, otherRulesKept: rulesKept } };
}

function wholeTaskChecks(op, finalObservation) {
  const v = verifyPrimitive(op, finalObservation, finalObservation);
  const what = op.intent !== 'set_validation' ? 'has no validation' : (op.want.type === 'xlValidateWholeNumber' && op.want.operator === 'xlBetween' ? 'validates a whole number between the bounds' : `validates ${op.want.type} ${op.want.operator}`);
  return [{ name: `${op.selector.sheet}!${op.selector.range} ${what}`, ok: v.ok, detail: v.ok ? null : v.evidence }];
}

const windowsOf = () => [];
const liveExtras = op => ({ dependents: true, validationRanges: [{ sheet: op.selector.sheet, range: op.selector.range }] });

module.exports = { INTENTS, plan, classify, resolve, lower, verifyPrimitive, verifyPreservation, wholeTaskChecks, windowsOf, liveExtras };
