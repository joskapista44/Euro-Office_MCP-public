'use strict';

const formula = require('../formula.cjs');
const { parseRange, colName, colNum, MAX_COL, overlaps } = require('../a1.cjs');
const cells = require('./cells.cjs');
const { nameUses } = require('../dependencies.cjs');
const { formulaEntries } = require('../inventory.cjs');

/**
 * XLSX `names` family: set_defined_name / delete_defined_name, WORKBOOK scope (M6 wave 6).
 *
 * The reader is the W3-proven, build-gated EO_INTERNAL_DEPENDENCY READ of every defined name (asc_getDefinedNames:
 * { name, ref, scopeSheet }, `ref` without its '='). The writes are public: Api.AddDefName / ApiDefName.Delete.
 * Audit: B-D26 / D30 MIGRATE; B-D28 (re-point) had NO preservation guard -- an unrelated existing name was silently
 * re-pointed; B-D30 had no dependent guard; B-D31 constants NOT_MEASURED.
 *
 * Contract (the W6 probe: SetRefersTo does NOT survive reopen, a constant throws -- both stay refused):
 *   set:    refersTo an ABSOLUTE reference ('=Sheet!$A$1', '=Sheet!$A$1:$B$2', quoted names) -- a relative one or a
 *           constant is refused. SATISFIED: exactly one workbook-level name of that name (case-insensitive) whose
 *           ref equals it (canonical). An EXISTING name with ANOTHER ref -> refused (no silent re-point).
 *   delete: SATISFIED when absent; refused while any formula, rule or other name USES it (it would break).
 * L1 the name's state; L3 every other name, every cell, every rule unchanged. Retry SATISFIED 0/0.
 *
 * M7 R16 (tier D probe, m7-d-probe1-20261003.json NM; W3 measured ws.AddDefName in the bare form and its scope through
 * the same internal reader): an optional `scope: '<sheet>'` makes the name SHEET-level -- written with
 * ws.AddDefName / ws.GetDefName(n).Delete on that sheet, read through the reader's scopeSheet. The probe: the same name
 * at workbook AND sheet level breaks a user formula (`#NAME?`, session and reopen) -- so a set or a delete is refused
 * while the name exists in ANY other scope. The in-use guard counts every formula that uses the name, whatever its
 * scope (conservative).
 */

const INTENTS = Object.freeze(['set_defined_name', 'delete_defined_name', 'set_print_area', 'clear_print_area', 'set_print_titles', 'clear_print_titles']);
const PRINT = Object.freeze({ set_print_area: 'Print_Area', clear_print_area: 'Print_Area', set_print_titles: 'Print_Titles', clear_print_titles: 'Print_Titles' });
const writes = op => op.intent === 'set_defined_name' || op.intent === 'set_print_area' || op.intent === 'set_print_titles';
const NAME = /^[A-Za-z_\\][A-Za-z0-9_.]{0,254}$/;
const RESERVED = /^(print_area|print_titles|_xlnm\..*)$/i;
const ABS_REF = /^=(?:'(?:[^']|'')+'|[A-Za-z_À-￿][\w.À-￿]*)!\$[A-Z]{1,3}\$\d+(?::\$[A-Z]{1,3}\$\d+)?$/;

/** The scope (M7 R16) and the reference of a step. */
function targetProblem(step) {
  if ('scope' in step && (typeof step.scope !== 'string' || step.scope === '')) return 'scope must be a sheet name (omit it for a workbook-level name)';
  if (step.intent === 'set_defined_name' && (typeof step.refersTo !== 'string' || !ABS_REF.test(step.refersTo))) return "refersTo must be an ABSOLUTE reference such as '=Sheet!$A$1' (a relative reference or a constant is not measured yet)";
  return null;
}

/**
 * X-P02 (smell-regression fix): validates range|areas and builds the comma-joined
 * refersTo for a SET print-area step. Split out of planPrint() so the branching stays in ONE
 * place (a maintainer's complexity-budget review, base = main 764f279) -- the behavior and the 4 X-P02
 * tests are unchanged.
 */
function printAreaRefersTo(step, at, q) {
  if ('range' in step && 'areas' in step) return { ok: false, detail: `${at}: exactly one of range or areas, not both` };
  if (!('range' in step) && !('areas' in step)) return { ok: false, detail: `${at} requires range (one A1 rectangle) or areas (an array of 2+)` };
  const raw = 'areas' in step ? step.areas : [step.range];
  if ('areas' in step && (!Array.isArray(raw) || raw.length < 2)) return { ok: false, detail: `${at}: areas must be an array of at least 2 ranges (use range for a single area)` };
  const parsed = [];
  for (const one of raw) {
    const g = parseRange(one);
    if (!g) return { ok: false, detail: `${at}: ${'areas' in step ? 'each area' : 'range'} must be one plain A1 range` };
    parsed.push(g);
  }
  for (let i = 0; i < parsed.length; i += 1) {
    for (let j = i + 1; j < parsed.length; j += 1) {
      if (overlaps(parsed[i], parsed[j])) return { ok: false, detail: `${at}: areas must not overlap` };
    }
  }
  return { refersTo: `=${parsed.map(g => `${q}!$${colName(g.c1)}$${g.r1}:$${colName(g.c2)}$${g.r2}`).join(',')}` };
}

/**
 * M7 F Q1 / Q2 / Q4 / Q5 (tier F probe m7-f-probe1-20261003.json, session = reopen): the print area and the print titles
 * are the SHEET-level names Print_Area / Print_Titles written through the PUBLIC ws.AddDefName with a bare, unquoted-
 * when-possible reference -- exactly what the editor path (the W8b oracle) stores -- and cleared through
 * ws.GetDefName(n).Delete(). Measured shapes only: one plain absolute rectangle for the area.
 *
 * M7 X-P03 (the tier F probe's own successor: docs/evidence/m7-p57-print-titles-columns-probe-20261005.json,
 * the asc_changePrintTitles oracle): titles take `rows: 'n:m'`, `columns: 'A:B'`, or BOTH -- at least one. The measured
 * stored shape with BOTH is COLUMNS FIRST, comma-joined, the sheet name repeated in EACH member
 * (`<Sheet>!$A:$B,<Sheet>!$1:$2`), never the other order and never a shared sheet prefix. Everything else follows the
 * sheet-scope name path (R16): L1 the name on its sheet, L3 every other name, cell and rule unchanged.
 *
 * M7 X-P02 (this session's own oracle: docs/evidence/m7-p61-print-area-multi-probe-20261005.json): a SET
 * area takes either `range` (the single rectangle, unchanged -- X-P01 stays exactly as it was) or `areas` (an array of
 * 2+ plain A1 rectangles, all on this ONE `sheet`, no per-area sheet -- the measured write, `AddDefName('Print_Area',
 * '<Sheet>!$A$1:$B$2,<Sheet>!$D$1:$E$2')`, is BAJTRA stable across 3 reads and the engine does NOT reorder on save
 * (measured both ways: forward and reversed), so the stored refersTo preserves the given array order verbatim. Two
 * areas overlapping within the SAME step is refused at plan() -- no "which one wins" is measured or wanted.
 */
function planPrint(step, index) {
  const at = `step ${index}: ${step.intent}`;
  const set = step.intent.startsWith('set_');
  const area = PRINT[step.intent] === 'Print_Area';
  const fields = ['intent', 'id', 'sheet', ...(set ? (area ? ['range', 'areas'] : ['rows', 'columns']) : [])];
  const unknown = Object.keys(step).filter(k => !fields.includes(k));
  if (unknown.length) return { ok: false, detail: `${at} does not support field(s): ${unknown.join(', ')}` };
  if (typeof step.sheet !== 'string' || step.sheet === '') return { ok: false, detail: `${at} requires a sheet name` };
  const q = /^[A-Za-z_][\w.]*$/.test(step.sheet) ? step.sheet : `'${step.sheet.replace(/'/g, "''")}'`;
  let refersTo = null;
  if (set && area) {
    const r = printAreaRefersTo(step, at, q);
    if (r.ok === false) return r;
    refersTo = r.refersTo;
  } else if (set) {
    if (!('rows' in step) && !('columns' in step)) return { ok: false, detail: `${at} requires at least one of rows ('n:m') or columns ('A:B')` };
    const parts = [];
    if ('columns' in step) {
      const m = /^([A-Z]{1,3}):([A-Z]{1,3})$/.exec(String(step.columns || ''));
      const c1 = m && colNum(m[1]);
      const c2 = m && colNum(m[2]);
      if (!m || !(c1 <= c2 && c2 <= MAX_COL)) return { ok: false, detail: `${at}: columns must be 'A:B' (plain, uppercase column letters, first at or before second)` };
      parts.push(`${q}!$${colName(c1)}:$${colName(c2)}`);
    }
    if ('rows' in step) {
      const m = /^(\d+):(\d+)$/.exec(String(step.rows || ''));
      if (!m || Number(m[1]) < 1 || Number(m[2]) < Number(m[1])) return { ok: false, detail: `${at}: rows must be 'n:m'` };
      parts.push(`${q}!$${m[1]}:$${m[2]}`);
    }
    // MEASURED order (m7-p57-print-titles-columns-probe-20261005.json): columns FIRST when both are given.
    refersTo = `=${parts.join(',')}`;
  }
  const name = PRINT[step.intent];
  return { ok: true, operation: { id: step.id || `s${index}`, intent: step.intent, family: 'names',
    selector: Object.freeze({ sheet: `\u0000name:${step.sheet.toLowerCase()}!${name.toLowerCase()}`, range: 'A1', name, claims: Object.freeze([`name:${step.sheet.toLowerCase()}!${name.toLowerCase()}`]) }),
    name, refersTo, scope: step.sheet, cells: Object.freeze([]), preservationMode: 'PRIMITIVE_INTEGRAL' } };
}

function planName(step, index) {
  const at = `step ${index}: ${step.intent}`;
  const fields = step.intent === 'set_defined_name' ? ['intent', 'id', 'name', 'refersTo', 'scope'] : ['intent', 'id', 'name', 'scope'];
  const unknown = Object.keys(step).filter(k => !fields.includes(k));
  if (unknown.length) return { ok: false, detail: `${at} does not support field(s): ${unknown.join(', ')}` };
  if (typeof step.name !== 'string' || !NAME.test(step.name) || /^[A-Za-z]{1,3}\d+$/.test(step.name) || /^(TRUE|FALSE)$/i.test(step.name)) return { ok: false, detail: `${at}: name must be a valid defined-name identifier (not cell-reference shaped)` };
  // M7 S1 (owner GO 2026-09-30): the reserved print names are not ordinary names -- what EO does with a workbook-level
  // one is not measured (M7 audit Q17), and print area / titles are not a supported write
  if (RESERVED.test(step.name)) return { ok: false, detail: `${at}: '${step.name}' is a reserved name (Print_Area, Print_Titles, _xlnm.*) and is refused` };
  const problem = targetProblem(step);
  if (problem) return { ok: false, detail: `${at}: ${problem}` };
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`,
      intent: step.intent,
      family: 'names',
      // names live outside any sheet: a per-name pseudo sheet (the planner's overlap check: two steps on ONE name clash)
      // a delete OWNS the name for the task: a formula step that uses it clashes (selector.uses, W9 composition)
      selector: Object.freeze({ sheet: `\u0000name:${step.name.toLowerCase()}`, range: 'A1', name: step.name, ...(step.intent === 'delete_defined_name' ? { claims: Object.freeze([`name:${step.name.toLowerCase()}`]) } : {}) }),
      name: step.name,
      refersTo: step.refersTo || null,
      scope: 'scope' in step ? step.scope : null,
      cells: Object.freeze([]),
      preservationMode: 'PRIMITIVE_INTEGRAL',
    },
  };
}

const plan = (step, index) => (PRINT[step.intent] ? planPrint : planName)(step, index);

const canon = ref => formula.canonical(ref.startsWith('=') ? ref : `=${ref}`);
const sameName = (a, b) => a.toLowerCase() === b.toLowerCase();
const scopeOf = op => op.scope || null;
const scopedHits = (p, op) => (Array.isArray(p.definedNames) ? p.definedNames.filter(d => d.scopeSheet === scopeOf(op) && sameName(d.name, op.name)) : null);
// a print name lives once per sheet (the tier F probe read two sheets' Print_Area side by side): only a WORKBOOK-level one
// is another scope for it (the measured collision, R16); for any other name every other scope is (conservative)
const otherScopes = (p, op) => p.definedNames.filter(d => d.scopeSheet !== scopeOf(op) && sameName(d.name, op.name) && (!PRINT[op.intent] || d.scopeSheet === null));

/** The scope sheet must exist; the name must not exist in any OTHER scope (the probe: both scopes break a formula). */
function scopeProblem(p, op) {
  if (op.scope && !(p.sheets || []).some(s => s.name === op.scope)) return { reason: 'xlsx-sheet-not-found', scope: op.scope };
  const elsewhere = otherScopes(p, op);
  return elsewhere.length ? { reason: 'xlsx-name-exists-in-another-scope', name: op.name, scopes: elsewhere.map(d => d.scopeSheet) } : null;
}

function gate(p, op) {
  if (!p || p.format !== 'xlsx') return { ok: false, reason: 'xlsx-observation-missing' };
  if (!p.complete) return { ok: false, reason: p.truncated ? 'xlsx-observation-truncated' : 'xlsx-observation-unreadable-getters', unsupported: p.unsupported };
  if (!Array.isArray(p.definedNames)) return { ok: false, reason: 'xlsx-names-unobservable' };
  if (!p.dependencies || !p.dependencies.complete) return { ok: false, reason: 'xlsx-names-dependency-observation-incomplete', unobservable: p.dependencies ? p.dependencies.unobservable.slice(0, 20) : 'not-read' };
  const problem = scopeProblem(p, op);
  return problem ? { ok: false, ...problem } : { ok: true, hits: scopedHits(p, op) };
}

/** Every formula text in the workbook that uses the name: cells, rules, other names. */
function usersOf(p, n) {
  const out = [];
  const uses = text => { const u = nameUses(text || ''); return !u.problem && u.names.some(x => sameName(x, n)); };
  // the read cells' formulas, plus the T3 formula index of a tiered read (M7 S32 slice 2)
  for (const [k, f] of formulaEntries(p)) if (uses(f)) out.push(k.replace('\u0000', '!'));
  if (p.rules && p.rules.bySheet) for (const [s, rs] of Object.entries(p.rules.bySheet)) rs.forEach((r, i) => { if (uses(r.formula) || uses(r.formula2)) out.push(`${s}:${r.kind}#${i}`); });
  for (const d of p.definedNames || []) if (!sameName(d.name, n) && uses(`=${d.ref}`)) out.push(`name:${d.name}`);
  return out;
}

function classify(p, op) {
  const g = gate(p, op);
  if (!g.ok) return { state: 'FAIL_CLOSED', evidence: g };
  if (writes(op)) {
    if (g.hits.length > 1) return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-name-ambiguous', name: op.name } };
    if (g.hits.length === 1) {
      return canon(g.hits[0].ref) === canon(op.refersTo)
        ? { state: 'SATISFIED', evidence: { name: op.name, refersTo: op.refersTo, proof: 'fresh-live-name-with-the-reference' } }
        : { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-name-exists-with-another-reference', name: op.name, now: g.hits[0].ref } };
    }
    return { state: 'APPLY', evidence: { name: op.name, refersTo: op.refersTo } };
  }
  if (g.hits.length === 0) return { state: 'SATISFIED', evidence: { name: op.name, proof: 'fresh-live-name-absent' } };
  const users = usersOf(p, op.name);
  if (users.length) return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-name-in-use', name: op.name, users: users.slice(0, 20), count: users.length } };
  return { state: 'APPLY', evidence: { name: op.name } };
}

function resolve(p, selector, ctx = {}) {
  const g = gate(p, ctx.operation);
  if (!g.ok) return { ok: false, target: null, selector, ambiguity: g.reason, detail: g };
  return { ok: true, target: Object.freeze({ name: ctx.operation.name }) };
}

function lower(op) {
  const n = JSON.stringify(op.name);
  const owner = op.scope ? 'ws' : 'Api';
  // WITHOUT the leading '=': AddDefName stores its text verbatim, and a name stored with '=' never resolves (W3e; W5-W9
  // LIVE run 1, 2026-09-29: the '=' form was read back as UNOBSERVABLE:name-text-starts-with-equals). W3g v2 proved the
  // bare form resolves.
  const stored = op.refersTo ? op.refersTo.replace(/^=/, '') : null;
  const call = writes(op) ? `${owner}.AddDefName(${n}, ${JSON.stringify(stored)});` : `const d = ${owner}.GetDefName(${n}); if (!d) return { ok: false, stage: 'name-missing' }; d.Delete();`;
  const sheet = op.scope ? `const ws = Api.GetSheet(${JSON.stringify(op.scope)});
  if (!ws || ws.GetName() !== ${JSON.stringify(op.scope)}) return { ok: false, stage: 'sheet-identity' };
  ` : '';
  return {
    kind: 'callCommand',
    body: `
try {
  ${sheet}${call}
  return { ok: true };
} catch (err) { return { ok: false, stage: 'threw', error: String(err) }; }
`,
    meta: Object.freeze({ family: 'names', action: op.intent, name: op.name, refersTo: stored, scope: op.scope }),
  };
}

function verifyPrimitive(op, _before, after) {
  const a = after && after.primary;
  const hits = a ? scopedHits(a, op) : null;
  const ok = Array.isArray(hits) && (writes(op) ? hits.length === 1 && canon(hits[0].ref) === canon(op.refersTo) : hits.length === 0);
  return { ok, evidence: { name: op.name, refersTo: op.refersTo, now: hits, proof: 'fresh-live-defined-names' } };
}

/** L3: every cell (cells.verifyPreservation), every OTHER name, every rule unchanged. */
function verifyPreservation(op, before, after) {
  const rest = cells.verifyPreservation(op, before, after);
  if (!rest.ok && rest.evidence && rest.evidence.reason) return rest;
  const b = before.primary;
  const a = after.primary;
  const key = d => `${d.scopeSheet || ''}\u0000${d.name.toLowerCase()}`;
  const others = x => (x.definedNames || []).filter(d => !(d.scopeSheet === scopeOf(op) && sameName(d.name, op.name))).map(d => `${key(d)}=${canon(d.ref)}`).sort();
  const namesKept = Array.isArray(b.definedNames) && Array.isArray(a.definedNames) && JSON.stringify(others(b)) === JSON.stringify(others(a));
  const rulesKept = !!b.rules && !!a.rules && JSON.stringify(b.rules.bySheet) === JSON.stringify(a.rules.bySheet);
  return { ok: rest.ok && namesKept && rulesKept, evidence: { ...rest.evidence, otherNamesKept: namesKept, rulesKept } };
}

function wholeTaskChecks(op, finalObservation) {
  const v = verifyPrimitive(op, finalObservation, finalObservation);
  return [{ name: `${op.scope ? `${op.scope}!` : ''}${op.name} ${writes(op) ? `refers to ${op.refersTo}` : 'is absent'}`, ok: v.ok, detail: v.ok ? null : v.evidence }];
}

const windowsOf = () => [];
const liveExtras = () => ({ dependents: true });

module.exports = { INTENTS, plan, classify, resolve, lower, verifyPrimitive, verifyPreservation, wholeTaskChecks, windowsOf, liveExtras };
