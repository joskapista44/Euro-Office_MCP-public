'use strict';

const { addr } = require('../a1.cjs');
const cells = require('./cells.cjs');

/**
 * XLSX `freeze` family: freeze_panes / unfreeze_panes (M6 wave 4).
 *
 * Measured (W4b2 LIVE): the public GetFreezePanes() acts on the ACTIVE sheet, write AND read, whatever sheet
 * object it is called on; with the target active, FreezeRows(n) / FreezeColumns(m) / FreezeAt(cell) freeze as
 * asked, FreezeAt(C3) freezing $A$1:$C$3 INCLUSIVE -- so n rows AND m columns = FreezeAt(the cell at column m,
 * row n).
 * Owner decisions (owner channel, #ref / #ref): the operation activates the target, freezes, and RESTORES the
 * originally active sheet in the same command (the restore also on a failure); the proof and the reopened retry
 * read the target's OWN pane WITHOUT activating it -- through the DS Bridge sheet.freezeState.read/v1 (B06 cutover;
 * a missing bridge or another build refuses every freeze step before the write). An unobservable pane is never
 * "not frozen".
 * The expected pane: frozen at the first UNfrozen cell { row0: rows, col0: columns } (0-based); unfreeze: none.
 * L1: the target's pane as expected AND the active sheet the same as before. L3: every other sheet's pane and
 * every cell unchanged. One freeze per task; in a mixed task its LAST step (M7 S5 S23 -- it moves the active sheet
 * during the command, so every other step is done and verified before it).
 * Retry = the terminal state on a fresh read -> SATISFIED, 0 write / 0 command.
 */

const INTENTS = Object.freeze(['freeze_panes', 'unfreeze_panes']);
// X-L09 (owner decision "A", internal tracker): raised from 1000 rows / 100 columns to the MEASURED range. LIVE
// (test/live/xlsx-xl09-freeze-above-cap-probe.live.cjs, LIVE probe, session + reopen): a freeze at 1001 / 2000 / 100000 rows, 101 / 200 / 1000
// columns and 1001 x 101 together lands in the file at exactly that pane -- the old 1000 / 100 was OUR implementation ceiling, not an engine limit.
// ABOVE these the plan still refuses, named (rows / columns must be an integer in [1, N]); nothing above was measured.
const MAX_ROWS = 100000;
const MAX_COLS = 1000;

function plan(step, index) {
  const at = `step ${index}: ${step.intent}`;
  const fields = step.intent === 'freeze_panes' ? ['intent', 'id', 'sheet', 'rows', 'columns'] : ['intent', 'id', 'sheet'];
  const unknown = Object.keys(step).filter(k => !fields.includes(k));
  if (unknown.length) return { ok: false, detail: `${at} does not support field(s): ${unknown.join(', ')}` };
  if (typeof step.sheet !== 'string' || step.sheet === '') return { ok: false, detail: `${at} requires a sheet name` };
  let expect = Object.freeze({ state: 'none' });
  if (step.intent === 'freeze_panes') {
    if (!('rows' in step) && !('columns' in step)) return { ok: false, detail: `${at} requires rows and / or columns` };
    if ('rows' in step && !(Number.isInteger(step.rows) && step.rows >= 1 && step.rows <= MAX_ROWS)) return { ok: false, detail: `${at}: rows must be an integer in [1, ${MAX_ROWS}]` };
    if ('columns' in step && !(Number.isInteger(step.columns) && step.columns >= 1 && step.columns <= MAX_COLS)) return { ok: false, detail: `${at}: columns must be an integer in [1, ${MAX_COLS}]` };
    expect = Object.freeze({ state: 'frozen', row0: step.rows || 0, col0: step.columns || 0 });
  }
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`,
      intent: step.intent,
      family: 'freeze',
      selector: Object.freeze({ sheet: step.sheet, range: 'A1' }),
      rows: step.rows || 0,
      columns: step.columns || 0,
      expect,
      cells: Object.freeze([]),
      preservationMode: 'PRIMITIVE_INTEGRAL',
    },
  };
}

const unobservable = x => typeof x === 'string';
const same = (p, e) => !!p && typeof p === 'object' && p.state === e.state && (e.state === 'none' || (p.row0 === e.row0 && p.col0 === e.col0));

function gate(primary, op) {
  if (!primary || primary.format !== 'xlsx') return { ok: false, reason: 'xlsx-observation-missing' };
  if (!primary.complete) return { ok: false, reason: primary.truncated ? 'xlsx-observation-truncated' : 'xlsx-observation-unreadable-getters', unsupported: primary.unsupported };
  const s = primary.sheets.filter(x => x.name === op.selector.sheet);
  if (s.length !== 1) return { ok: false, reason: s.length > 1 ? 'xlsx-sheet-ambiguous' : 'xlsx-sheet-not-found', available: primary.sheets.map(x => x.name) };
  if (typeof primary.active !== 'string' || !primary.sheets.some(x => x.name === primary.active)) return { ok: false, reason: 'xlsx-active-sheet-unread', active: primary.active };
  if (!primary.panes) return { ok: false, reason: 'xlsx-freeze-pane-unobservable', sheet: op.selector.sheet, why: 'not-read' };
  // X-L10 (LIVE-measured, test/live/xlsx-freeze-active-sheet-persistence.live.cjs): the SAVED file's activeTab follows
  // the freeze TARGET -- the in-command restore of the originally active sheet holds in the session (L1) but is not what
  // the editor saves -- so a HIDDEN target would open as the active (shown) sheet. A hidden target is refused before
  // the write; an unread visibility is never "visible".
  const vis = primary.visibility ? primary.visibility[op.selector.sheet] : undefined;
  if (vis === undefined || typeof vis !== 'boolean') return { ok: false, reason: 'xlsx-freeze-sheet-visibility-unobservable', sheet: op.selector.sheet, why: vis === undefined ? 'not-read' : vis };
  if (vis === false) return { ok: false, reason: 'xlsx-freeze-hidden-sheet-refused', sheet: op.selector.sheet };
  // every sheet's pane: the target's is the proof, the others' the L3 baseline
  for (const x of primary.sheets) {
    const p = primary.panes[x.name];
    if (p === undefined || unobservable(p)) return { ok: false, reason: 'xlsx-freeze-pane-unobservable', sheet: x.name, why: p === undefined ? 'not-read' : p };
  }
  return { ok: true, pane: primary.panes[op.selector.sheet], active: primary.active };
}

function classify(primary, op) {
  const g = gate(primary, op);
  if (!g.ok) return { state: 'FAIL_CLOSED', evidence: g };
  const where = { sheet: op.selector.sheet, expect: op.expect };
  if (same(g.pane, op.expect)) return { state: 'SATISFIED', evidence: { ...where, proof: 'fresh-live-own-pane-without-activation' } };
  return { state: 'APPLY', evidence: { ...where, now: g.pane, active: g.active } };
}

function resolve(primary, selector, ctx = {}) {
  const g = gate(primary, ctx.operation);
  if (!g.ok) return { ok: false, target: null, selector, ambiguity: g.reason, detail: g };
  return { ok: true, target: Object.freeze({ sheet: selector.sheet, pane: g.pane, active: g.active }) };
}

/**
 * One callCommand: exact-name sheet check; record the active sheet; activate the target (checked); the public
 * freeze call; restore the original active sheet -- ALSO after a failure -- and report whether it was restored.
 */
function lower(op) {
  const { sheet } = op.selector;
  const call = op.intent === 'unfreeze_panes' ? 'fp.Unfreeze();'
    : (op.rows && op.columns ? `fp.FreezeAt(ws.GetRange(${JSON.stringify(addr(op.columns, op.rows))}));`
      : (op.rows ? `fp.FreezeRows(${op.rows});` : `fp.FreezeColumns(${op.columns});`));
  return {
    kind: 'callCommand',
    body: `
const name = ${JSON.stringify(sheet)};
let origin = null;
let back = null;
let result;
try {
  const ws = Api.GetSheet(name);
  if (!ws || ws.GetName() !== name) return { ok: false, stage: 'sheet-identity' };
  origin = Api.GetActiveSheet().GetName();
  back = Api.GetSheet(origin);
  ws.SetActive();
  if (Api.GetActiveSheet().GetName() !== name) result = { ok: false, stage: 'activate' };
  else { const fp = ws.GetFreezePanes(); ${call} result = { ok: true }; }
} catch (err) { result = { ok: false, stage: 'threw', error: String(err) }; }
if (back === null) return { ...result, ok: false, restored: null };
try { back.SetActive(); } catch (err) { return { ...result, ok: false, restored: false, restoreError: String(err) }; }
const restored = Api.GetActiveSheet().GetName() === origin;
return { ...result, ok: result.ok && restored, origin, restored };
`,
    meta: Object.freeze({ family: 'freeze', action: op.intent, sheet, rows: op.rows, columns: op.columns }),
  };
}

/** L1: the target's own pane (no activation) as expected, and the active sheet the one before the command. */
function verifyPrimitive(op, before, after) {
  const b = before && before.primary;
  const a = after && after.primary;
  const pane = a && a.panes ? a.panes[op.selector.sheet] : undefined;
  const paneOk = same(pane, op.expect);
  const activeKept = !!b && !!a && typeof a.active === 'string' && a.active === b.active;
  return { ok: paneOk && activeKept, evidence: { sheet: op.selector.sheet, expect: op.expect, pane: pane === undefined ? 'not-read' : pane, activeBefore: b ? b.active : null, activeAfter: a ? a.active : null, activeKept, proof: 'fresh-live-own-pane-without-activation' } };
}

/** L3: every cell's content, and every OTHER sheet's pane unchanged. */
function verifyPreservation(op, before, after) {
  const rest = cells.verifyPreservation(op, before, after);
  if (!rest.ok && rest.evidence && rest.evidence.reason) return rest;
  const b = before.primary.panes || {};
  const a = after.primary.panes || {};
  const changed = Object.keys(b).filter(n => n !== op.selector.sheet && JSON.stringify(a[n]) !== JSON.stringify(b[n])).map(n => ({ sheet: n, before: b[n], after: a[n] === undefined ? 'not-read' : a[n] }));
  // X-L13/6: once tiered, report preservationScope FULL -- `panes`/`sheetVisibility` are read per-SHEET
  // (xlsx-driver.cjs: `for (const s of sheets) panes[s.row.name] = ...`), unconditional on used-cell count,
  // so every other sheet's pane stays fully checked above. Below 45000 unchanged.
  const tiered = before.primary.tier === 'tiered';
  return { ok: rest.ok && changed.length === 0, evidence: { ...rest.evidence, paneChanged: changed.slice(0, 20), paneChangedCount: changed.length, ...(tiered ? { preservationScope: 'FULL' } : {}) } };
}

function wholeTaskChecks(op, finalObservation) {
  const p = finalObservation && finalObservation.primary;
  const pane = p && p.panes ? p.panes[op.selector.sheet] : undefined;
  return [{ name: `${op.selector.sheet} ${op.intent === 'freeze_panes' ? `frozen at row0 ${op.rows}, col0 ${op.columns}` : 'not frozen'}`, ok: same(pane, op.expect), detail: same(pane, op.expect) ? null : { pane: pane === undefined ? 'not-read' : pane } }];
}

const windowsOf = () => [];
const liveExtras = () => ({ panes: true, sheetVisibility: true });

module.exports = { INTENTS, plan, classify, resolve, lower, verifyPrimitive, verifyPreservation, wholeTaskChecks, windowsOf, liveExtras };
