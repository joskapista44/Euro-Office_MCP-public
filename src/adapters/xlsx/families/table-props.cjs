'use strict';

const { parseRange, parseCell, rangeText, MAX_ROW } = require('../a1.cjs');
const { XLSX_B25_BRIDGE_WRITE_RUNTIME_IDS, XLSX_RUNTIME_24_ID, XLSX_RUNTIME_25_ID, XLSX_RUNTIME_26_ID, XLSX_RUNTIME_27_ID, XLSX_RUNTIME_28_ID, XLSX_RUNTIME_29_ID } = require('../runtime.cjs');
const { referencesBlock } = require('../shift.cjs');
const { formulaEntries } = require('../inventory.cjs');
const { tableRenameRewrite, sheetRefs, tokenize } = require('../formula.cjs');
const ts = require('./table-structure.cjs');

/**
 * XLSX `table-props` family (GH #332 T02 rename, #333 T03 resize, #334 T04a style + stripe / banding flags, #335 T04b totals row):
 * set_table_style {sheet, table, style} | set_table_flag {sheet, table, flag, value} | set_table_totals_row {sheet, table, value: true} |
 * rename_table {sheet, table, name} | resize_table {sheet, table, ref}. Each is the ONLY step of its task.
 *
 * ENGINE PATH: ONE sheet.table.write/v1 Bridge WRITE (DS Bridge /19 ONLY) -- the editor's own table actions; the public API has no
 * table setter (PUBLIC-FIRST probe, docs/evidence/gh332-gh335-public-probe-20261030.json: 23 candidate calls, all unchanged after a
 * fresh reopen). On /16../18 the capability does not exist: the plan is FAIL_CLOSED (named, no dispatch). The handler itself verifies
 * (read-after-write through sheet.tables.read/v1) and answers with named refusals; THIS family adds the product-side refusals the
 * measurements demand and the L1 / L3 verification from a fresh read_tables + cells.
 *
 * MEASURED hazards (docs/DS-9.3.5-BRIDGE19-TABLE-WRITE-DESIGN.md, engine-access probe) -> named refusals BEFORE the write:
 *   rename / resize while ANY structured reference to the table exists     xlsx-table-structured-reference
 *       (rename -> #NAME?; resize -> the STORED value of the reference stays stale)
 *   a table without a header row (totals) / with a totals row (resize)     xlsx-table-no-header-row / xlsx-table-totals-row-unmeasured
 *   totals row: the sheet's last row / content already under the table    xlsx-table-totals-row-no-room / -below-occupied
 *   resize: another origin / a width change / a height under 2 / an overlap with another table / data, a merge or a pivot in the
 *       extension                                                         xlsx-table-resize-*
 *   a validation / rule / chart / defined name reaching the sheet           xlsx-table-props-dependent-not-measured (conservative)
 * NOT offered in v1: header row OFF (it deletes the header cell contents), per-column totals functions (GH #335 remainder,
 *   LIVE: docs/evidence/gh335-totals-header-live-1-20261102.json -- neither a dedicated engine setter nor write_range reaches
 *   a per-column totals function; a totals cell already holding a formula refuses write_range as xlsx-write-over-formula-refused).
 * totals OFF (GH #335 remainder): the EDITOR path (asc_changeFormatTableInfo) is measured clean (LIVE, a two-session
 *   ON->OFF round-trip with an in-session mid-read, docs/evidence/gh335-totals-header-live-1-20261102.json). The
 *   Bridge's OWN sheet.table.write/v1 handler refuses any value !== true on /19, /21, /22, /23 (`totals-off-not-
 *   measured`, verified against every shipped revision's own candidate block) -- a value: false step is still a
 *   NAMED, plan-time-proven-runtime refusal (xlsx-table-totals-off-bridge-runtime-not-proven) on those. /24 (GH
 *   #338-#342 remainder, docs/DS-9.3.5-BRIDGE24-WORKBOOK-PROTECTION-DESIGN.md) is the ONE revision that DOES offer
 *   it (TOTALS_OFF_BRIDGE_RUNTIMES, below), now proven by the same byte-identity warrant every other capability-
 *   specific runtime list carries (test/static/ds935-bridge24-write-handler-identity.test.cjs). /25 (the W2
 *   regression fix, workbook.protection.read/v1 only -- sheet.table.write/v1 untouched, byte-identical to /24's
 *   handler, test/static/ds935-bridge25-write-handler-identity.test.cjs) carries the SAME totalsRow value:false
 *   acceptance for that reason. The schema and classify()/lower() plumbing for value: false already mirrored
 *   value: true before /24 existed, so the activation was a one-line runtime-list change, not a redesign.
 * RETRY: the fresh read already holding the target value -> SATISFIED (0 write) with the named observation
 * 'table-property-already-at-the-target-value'. rename: the old name absent and exactly one table of the new name -> SATISFIED
 * ('table-already-renamed').
 * The workbook-wide name collision is the HANDLER's refusal (it reads every sheet); this family only sees the named sheet.
 */

const INTENTS = Object.freeze(['set_table_style', 'set_table_flag', 'set_table_totals_row', 'rename_table', 'resize_table']);
const PROPERTY = Object.freeze({ set_table_style: 'style', set_table_flag: 'flag', set_table_totals_row: 'totalsRow', rename_table: 'name', resize_table: 'ref' });
const FLAGS = Object.freeze(['rowStripes', 'columnStripes', 'firstColumn', 'lastColumn']);
const FIELDS = Object.freeze({
  set_table_style: ['style'], set_table_flag: ['flag', 'value'], set_table_totals_row: ['value'], rename_table: ['name'], resize_table: ['ref'],
});
const B25_BRIDGE_WRITE_RUNTIMES = XLSX_B25_BRIDGE_WRITE_RUNTIME_IDS;
// GH #335 remainder (internal tracker): the Bridge's OWN sheet.table.write/v1 handler refuses value !== true on
// /19, /21, /22, /23 (grep 'totals-off-not-measured' in each rev{19,20,21,22,23}-candidate bridge-block.js -> 1/1/1/1/1)
// -- NOT a family-wide capability (so it does not live in runtime.cjs next to B25/B26/B27's own already-landed
// lists), but a SEPARATE, per-revision list of its own. /24 is the one revision proven to accept value:false
// (test/static/ds935-bridge24-write-handler-identity.test.cjs's byte-identity warrant names sheet.table.write/v1
// as the ONE exception on /24, the totalsRow branch); /25 (the W2 regression fix) carries the SAME handler
// byte-identical (test/static/ds935-bridge25-write-handler-identity.test.cjs names workbook.protection.read/v1 as
// /25's ONE exception, not this one), so it accepts value:false too. /26 (review fix, f0451f47: the two handler
// guards switch to the workbook.protection.read/v1 getter) touches sheet.table.write/v1's PROTECTION GUARD only --
// the totalsRow value:false branch it checks here is outside that guard and stays byte-identical
// (test/acceptance/ds935-bridge26-worksheet-copy-table-write-protection.test.cjs), so it accepts value:false too.
// /27 (GH #346: chart.series.style.read/v1 changed, chart.series.dataPoint.style.write/v1 added) carries sheet.table.write/v1
// byte-identical (test/static/ds935-bridge27-write-handler-identity.test.cjs), so it accepts value:false too.
// /28 (GH #361/#363 pre-install readiness gate): carries sheet.table.write/v1 byte-identical too
// (test/static/ds935-bridge28-table-autofilter-criteria-write-handler-identity.test.cjs's WARRANT -- sheet.table.
// write/v1 is one of the /27 handlers it proves unchanged), so it accepts value:false too.
// Frozen, pinned by the SAME static-identity test every other capability-specific runtime list (B25/B26/B27/B28) is pinned by.
const TOTALS_OFF_BRIDGE_RUNTIMES = Object.freeze([XLSX_RUNTIME_24_ID, XLSX_RUNTIME_25_ID, XLSX_RUNTIME_26_ID, XLSX_RUNTIME_27_ID, XLSX_RUNTIME_28_ID, XLSX_RUNTIME_29_ID]);
const NAMED = 'table-property-already-at-the-target-value';

/** the engine's predefined table styles: Light 1..21, Medium 1..28, Dark 1..11 (the handler gate is the same pattern) */
function styleOk(name) {
  const m = /^TableStyle(Light|Medium|Dark)([1-9][0-9]?)$/.exec(String(name));
  return !!m && Number(m[2]) <= { Light: 21, Medium: 28, Dark: 11 }[m[1]];
}
const nameOk = n => typeof n === 'string' && n.length >= 1 && n.length <= 255 && /^[A-Za-z_][A-Za-z0-9_.]*$/.test(n) && !/^[A-Za-z]{1,3}[0-9]+$/.test(n) && !/^[RrCc][0-9]*$/.test(n);
const normRef = r => String(r).replace(/\$/g, '').toUpperCase();

function plan(step, index) {
  const intent = step.intent;
  const at = `step ${index}: ${intent}`;
  if (!FIELDS[intent]) return { ok: false, detail: `${at} is not a table-props intent` };
  const fields = ['intent', 'id', 'sheet', 'table', ...FIELDS[intent]];
  const unknown = Object.keys(step).filter(k => !fields.includes(k));
  if (unknown.length) return { ok: false, detail: `${at} does not support field(s): ${unknown.join(', ')}` };
  if (typeof step.sheet !== 'string' || step.sheet === '') return { ok: false, detail: `${at} requires a sheet name` };
  if (typeof step.table !== 'string' || step.table.trim() === '') return { ok: false, detail: `${at} requires the table name` };
  const want = {};
  if (intent === 'set_table_style') {
    if (!styleOk(step.style)) return { ok: false, detail: `${at} requires style: one of TableStyleLight1..21, TableStyleMedium1..28, TableStyleDark1..11` };
    want.style = step.style;
  } else if (intent === 'set_table_flag') {
    if (!FLAGS.includes(step.flag)) return { ok: false, detail: `${at} requires flag: one of ${FLAGS.join(', ')}` };
    if (typeof step.value !== 'boolean') return { ok: false, detail: `${at} requires value: a boolean` };
    Object.assign(want, { flag: step.flag, value: step.value });
  } else if (intent === 'set_table_totals_row') {
    if (typeof step.value !== 'boolean') return { ok: false, detail: `${at} requires value: a boolean` };
    want.value = step.value;
  } else if (intent === 'rename_table') {
    if (!nameOk(step.name)) return { ok: false, detail: `${at} requires name: 1..255 characters, a letter or underscore first, then letters, digits, underscore or dot; not a cell reference` };
    want.name = step.name;
  } else {
    if (typeof step.ref !== 'string' || !parseRange(normRef(step.ref)) || !/^[A-Z]{1,3}[1-9][0-9]{0,6}:[A-Z]{1,3}[1-9][0-9]{0,6}$/.test(normRef(step.ref))) return { ok: false, detail: `${at} requires ref: an A1 range such as A1:C7 (the new table range)` };
    want.ref = normRef(step.ref);
  }
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`,
      intent,
      family: 'table-props',
      selector: Object.freeze({ sheet: step.sheet, table: step.table }),
      want: Object.freeze(want),
      // the family owns its own cell L1 / L3; the generic cell claim is empty
      cells: Object.freeze([]),
      preservationMode: 'PRIMITIVE_INTEGRAL',
    },
  };
}

/** the sheet's tables and the named one (rename may name the NEW name when the old one is gone), or the named reason they cannot be told */
function gate(p, op, { fullRead = false } = {}) {
  if (!p || p.format !== 'xlsx') return { ok: false, reason: 'xlsx-observation-missing' };
  if (fullRead) {
    if (!p.complete) return { ok: false, reason: p.truncated ? 'xlsx-observation-truncated' : 'xlsx-observation-unreadable-getters', unsupported: p.unsupported };
    if (p.tier && p.tier !== 'full') return { ok: false, reason: 'xlsx-table-props-needs-a-full-read', tier: p.tier };
  }
  const s = p.sheets.filter(x => x.name === op.selector.sheet);
  if (s.length !== 1) return { ok: false, reason: s.length > 1 ? 'xlsx-sheet-ambiguous' : 'xlsx-sheet-not-found', available: p.sheets.map(x => x.name) };
  const read = p.tableReads ? p.tableReads[op.selector.sheet] : undefined;
  if (read === undefined) return { ok: false, reason: 'xlsx-tables-unobservable', why: 'not-read' };
  if (typeof read === 'string') return { ok: false, reason: 'xlsx-tables-unobservable', why: read };
  if (!read || !Array.isArray(read.tables)) return { ok: false, reason: 'xlsx-tables-unobservable', why: 'malformed-read' };
  const find = n => read.tables.filter(t => String(t.name).toLowerCase() === String(n).toLowerCase());
  const named = find(op.selector.table);
  if (named.length > 1) return { ok: false, reason: 'xlsx-table-name-ambiguous', available: read.tables.map(t => t.name) };
  if (named.length === 0) return { ok: false, reason: 'xlsx-table-not-found', available: read.tables.map(t => t.name), tables: read.tables, find };
  const g = parseRange(String(named[0].ref));
  if (!g) return { ok: false, reason: 'xlsx-table-ref-unparsable', table: named[0].name, ref: named[0].ref };
  return { ok: true, tables: read.tables, table: named[0], g, find };
}

const blank = c => !c || (c.f === '' && (c.v === null || c.v === ''));
const within = (b, r, c) => r >= b.r1 && r <= b.r2 && c >= b.c1 && c <= b.c2;

/** cells holding content inside a rectangle (needs the full read) */
function occupied(p, sheet, rect) {
  const hit = [];
  for (const [a, c] of ts.cellsOfSheet(p, sheet)) {
    const q = parseRange(a);
    if (q && within(rect, q.r1, q.c1) && !blank(c)) hit.push(a);
  }
  return hit;
}

function classify(p, op) {
  const runtime = p && p.runtime ? p.runtime.identity : null;
  const where = { sheet: op.selector.sheet, table: op.selector.table };
  if (!B25_BRIDGE_WRITE_RUNTIMES.includes(runtime)) {
    return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-table-props-bridge-write-runtime-not-proven', capability: 'CAPABILITY_MISSING', ...where, runtime, proven: B25_BRIDGE_WRITE_RUNTIMES } };
  }
  // GH #335 remainder: the Bridge itself refuses value !== true on every revision shipped today -- a SEPARATE,
  // narrower runtime gate than the family-wide one above (which only proves sheet.table.write/v1 EXISTS, not that
  // it accepts totals OFF). Named so it reads as "not yet", not as the old, pre-measurement flat refusal.
  if (op.intent === 'set_table_totals_row' && op.want.value === false && !TOTALS_OFF_BRIDGE_RUNTIMES.includes(runtime)) {
    return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-table-totals-off-bridge-runtime-not-proven', capability: 'CAPABILITY_MISSING', ...where, runtime, proven: TOTALS_OFF_BRIDGE_RUNTIMES } };
  }
  const needsFull = op.intent === 'rename_table' || op.intent === 'resize_table' || op.intent === 'set_table_totals_row';
  const t = gate(p, op, { fullRead: needsFull });
  if (!t.ok) {
    // a rename retry: the old name is gone and exactly one table holds the new name
    if (op.intent === 'rename_table' && t.reason === 'xlsx-table-not-found') {
      const now = t.find(op.want.name);
      if (now.length === 1) return { state: 'SATISFIED', evidence: { ...where, newName: now[0].name, observation: 'table-already-renamed', proof: 'fresh-live-read_tables' } };
    }
    return { state: 'FAIL_CLOSED', evidence: t };
  }
  const tb = t.table;
  const w = op.want;
  const named = { ...where, table: tb.name };
  const sat = extra => ({ state: 'SATISFIED', evidence: { ...named, observation: NAMED, ...extra, proof: 'fresh-live-read_tables' } });
  const dep = ts.dependencyConflict(p, op.selector.sheet, tb.name);

  if (op.intent === 'set_table_style') {
    if (tb.style && tb.style.name === w.style) return sat({ style: w.style });
    return { state: 'APPLY', evidence: { ...named, from: tb.style ? tb.style.name : null, to: w.style } };
  }
  if (op.intent === 'set_table_flag') {
    if (!tb.style) return { state: 'FAIL_CLOSED', evidence: { ...named, reason: 'xlsx-table-style-info-missing' } };
    if (tb.style[w.flag] === w.value) return sat({ flag: w.flag, value: w.value });
    return { state: 'APPLY', evidence: { ...named, flag: w.flag, from: tb.style[w.flag], to: w.value } };
  }
  if (op.intent === 'set_table_totals_row') {
    if (w.value === true) {
      if (tb.totalsRow === true) return sat({ totalsRow: true });
      if (tb.headerRow !== true) return { state: 'FAIL_CLOSED', evidence: { ...named, reason: 'xlsx-table-no-header-row' } };
      if (t.g.r2 + 1 >= MAX_ROW) return { state: 'FAIL_CLOSED', evidence: { ...named, reason: 'xlsx-table-totals-row-no-room', note: 'the table or the row below it is the last row of the sheet' } };
      const below = { r1: t.g.r2 + 1, r2: t.g.r2 + 1, c1: t.g.c1, c2: t.g.c2 };
      const occ = occupied(p, op.selector.sheet, below);
      if (occ.length) return { state: 'FAIL_CLOSED', evidence: { ...named, reason: 'xlsx-table-totals-row-below-occupied', cells: occ.slice(0, 20) } };
      const b = ts.bandConflict(p, op, t, below);
      if (b) return { state: 'FAIL_CLOSED', evidence: { ...named, ...b } };
      if (dep) return { state: 'FAIL_CLOSED', evidence: { ...named, ...retag(dep) } };
      return { state: 'APPLY', evidence: { ...named, ref: tb.ref, row: t.g.r2 + 1 } };
    }
    // GH #335 remainder: value: false -- removing an existing totals row (measured clean on LIVE, two separate
    // editor sessions with an in-session read of the mid-state between the ON call and the OFF call). NOT measured
    // LIVE: a formula that reaches the totals row itself (the probe's structured-reference cells did not point at
    // it) -- conservative, same rule family as rename/resize: a structured reference anywhere, or a plain A1
    // reference into the totals row, refuses before the write (GH #335 remainder, internal tracker).
    //
    // EXCLUDED from both checks below: the totals row's OWN cells. The engine fills them with its own formula
    // when the totals row is switched ON (e.g. a SUBTOTAL over a structured column reference) -- that formula is
    // itself a structured/A1 reference to the table, but it is removed WITH the row, so it never goes stale
    // (measured LIVE on a disposable test server: the own-cell SUBTOTAL formula was tripping the structured-
    // reference check on every totals-row table, making totals OFF unusable in practice).
    if (tb.totalsRow !== true) return sat({ totalsRow: false });
    const totalsBand = { sheet: op.selector.sheet, r1: t.g.r2, r2: t.g.r2, c1: t.g.c1, c2: t.g.c2 };
    const ownTotalsCell = key => {
      const sep = key.indexOf('!');
      if (sep < 0 || key.slice(0, sep) !== totalsBand.sheet) return false;
      const cell = parseCell(key.slice(sep + 1));
      return !!cell && cell.r === totalsBand.r2 && cell.c >= totalsBand.c1 && cell.c <= totalsBand.c2;
    };
    const totalsRowRefs = ts.structuredRefs(p, tb.name).filter(k => !ownTotalsCell(k));
    if (totalsRowRefs.length) return { state: 'FAIL_CLOSED', evidence: { ...named, reason: 'xlsx-table-structured-reference', formulas: totalsRowRefs.slice(0, 20) } };
    const touching = formulaEntries(p).filter(([k, f]) => referencesBlock(f, k.split('\u0000')[0], totalsBand)).map(([k]) => k.replace('\u0000', '!')).filter(k => !ownTotalsCell(k));
    if (touching.length) return { state: 'FAIL_CLOSED', evidence: { ...named, reason: 'xlsx-table-totals-row-referenced', formulas: touching.slice(0, 20) } };
    // the same own-row exclusion, one layer down: the generic structural-dependency reader (dependencies.cjs)
    // marks the totals row's own structured-reference formula UNOBSERVABLE too, which on its own made the whole
    // workbook's dependency set 'incomplete' -- masked by the two checks above until they were scoped to exclude
    // the own row. Narrowed the same way via dependencyConflict's ignoreUnobservable predicate -- the REACH scan
    // (a validation / conditional-format / chart / defined-name that DOES depend on the table) always still runs,
    // so a totals-row-only gap is forgiven but a real dependent elsewhere is not.
    const unobsOwnRowOnly = u => {
      if (u.host !== totalsBand.sheet) return false;
      const parts = String(u.axis || '').split(':');
      const cell = parseCell(parts[parts.length - 1]);
      return !!cell && cell.r === totalsBand.r2 && cell.c >= totalsBand.c1 && cell.c <= totalsBand.c2;
    };
    const depOff = ts.dependencyConflict(p, op.selector.sheet, tb.name, unobsOwnRowOnly);
    if (depOff) return { state: 'FAIL_CLOSED', evidence: { ...named, ...retag(depOff) } };
    return { state: 'APPLY', evidence: { ...named, ref: tb.ref, row: t.g.r2 } };
  }
  if (op.intent === 'rename_table') {
    if (tb.name === w.name) return sat({ name: w.name });
    const clash = t.find(w.name).filter(x => x !== tb);
    if (clash.length) return { state: 'FAIL_CLOSED', evidence: { ...named, reason: 'xlsx-table-name-taken', other: clash[0].name } };
    // GH #332/#333 (owner decision A): rename no longer refuses a structured reference to THIS table --
    // MEASURED LIVE (the GH #332/#333 research probe, prod): renaming alone leaves the old name in the dependent formula (#NAME?
    // on reopen). The family now rewrites each one, in the SAME callCommand, to the new name (lower()
    // below); a reference the rewrite cannot safely reach (the name appears inside a string literal --
    // e.g. INDIRECT("Table1[Qty]")) stays a NAMED refusal, unchanged.
    const refs = ts.structuredRefs(p, tb.name);
    let rewrites = [];
    if (refs.length) {
      rewrites = buildRenameRewrites(p, refs, tb.name, w.name);
      if (!rewrites) return { state: 'FAIL_CLOSED', evidence: { ...named, reason: 'xlsx-table-structured-reference', formulas: refs.slice(0, 20) } };
    }
    const renameIgnore = combineIgnore(refAxisIgnore(refs), otherTableAxisIgnore(p, [tb.name, w.name]));
    const renameDep = ts.dependencyConflict(p, op.selector.sheet, tb.name, renameIgnore);
    if (renameDep) return { state: 'FAIL_CLOSED', evidence: { ...named, ...retag(renameDep) } };
    op._renameFormulaRewrites = Object.freeze(rewrites.map(r => Object.freeze(r)));
    return { state: 'APPLY', evidence: { ...named, from: tb.name, to: w.name, ...(rewrites.length ? { formulasRewritten: rewrites.map(r => r.address) } : {}) } };
  }
  // resize_table
  const want = parseRange(w.ref);
  if (normRef(tb.ref) === w.ref) return sat({ ref: w.ref });
  if (tb.totalsRow !== false) return { state: 'FAIL_CLOSED', evidence: { ...named, reason: 'xlsx-table-totals-row-unmeasured' } };
  if (tb.headerRow !== true) return { state: 'FAIL_CLOSED', evidence: { ...named, reason: 'xlsx-table-no-header-row' } };
  if (want.r1 !== t.g.r1 || want.c1 !== t.g.c1) return { state: 'FAIL_CLOSED', evidence: { ...named, reason: 'xlsx-table-resize-moves-origin', ref: tb.ref, want: w.ref } };
  if (want.c2 !== t.g.c2) return { state: 'FAIL_CLOSED', evidence: { ...named, reason: 'xlsx-table-resize-width-change-unmeasured', ref: tb.ref, want: w.ref } };
  if (want.r2 < want.r1 + 1) return { state: 'FAIL_CLOSED', evidence: { ...named, reason: 'xlsx-table-resize-too-small', want: w.ref } };
  if (want.r2 > MAX_ROW) return { state: 'FAIL_CLOSED', evidence: { ...named, reason: 'xlsx-table-resize-out-of-sheet', want: w.ref } };
  for (const o of t.tables) {
    if (o === tb) continue;
    const og = parseRange(String(o.ref));
    if (!og) return { state: 'FAIL_CLOSED', evidence: { ...named, reason: 'xlsx-table-ref-unparsable', table: o.name, ref: o.ref } };
    if (og.r2 >= want.r1 && og.r1 <= want.r2 && og.c2 >= want.c1 && og.c1 <= want.c2) return { state: 'FAIL_CLOSED', evidence: { ...named, reason: 'xlsx-table-resize-overlaps-table', other: o.name, ref: o.ref } };
  }
  // GH #332/#333 (owner decision A): a structured reference to THIS table no
  // longer refuses resize before dispatch -- MEASURED LIVE (the GH #332/#333 research probe, prod DS /26): resize does NOT
  // leave the dependent formula stale. The stored cache reads correctly after a save+reopen; the ONLY
  // artefact is an in-session read of '' until the engine recalculates, which verifyPreservation below
  // excludes for these specific cells (op._resizeStructuredRefs). `refs` still feeds the dependency-gate's
  // ignore predicate, so an UNRELATED structured reference (another table) is unaffected.
  const refs = ts.structuredRefs(p, tb.name);
  const ext = want.r2 > t.g.r2 ? { r1: t.g.r2 + 1, r2: want.r2, c1: want.c1, c2: want.c2 } : null;
  if (ext) {
    const occ = occupied(p, op.selector.sheet, ext);
    if (occ.length) return { state: 'FAIL_CLOSED', evidence: { ...named, reason: 'xlsx-table-resize-extension-has-data', cells: occ.slice(0, 20) } };
    const b = ts.bandConflict(p, op, t, ext);
    if (b) return { state: 'FAIL_CLOSED', evidence: { ...named, ...b } };
  }
  const resizeIgnore = combineIgnore(refAxisIgnore(refs), otherTableAxisIgnore(p, [tb.name]));
  const resizeDep = ts.dependencyConflict(p, op.selector.sheet, tb.name, resizeIgnore);
  if (resizeDep) return { state: 'FAIL_CLOSED', evidence: { ...named, ...retag(resizeDep) } };
  if (refs.length) op._resizeStructuredRefs = Object.freeze(refs);
  return { state: 'APPLY', evidence: { ...named, ref: tb.ref, to: w.ref, ...(refs.length ? { structuredReferencesPresent: refs.slice(0, 20) } : {}) } };
}

/** the structure family's dependency reason, named for this family */
const retag = d => ({ ...d, reason: d.reason === 'xlsx-table-structure-dependent-not-measured' ? 'xlsx-table-props-dependent-not-measured' : d.reason });

/**
 * GH #332/#333: a dependencyConflict ignore-predicate that forgives ONLY the unobservable-formula
 * entries `refs` already names (a structured reference to THIS table, already measured/handled by the
 * narrower check above) -- any OTHER unobservable entry (a reference to a different table, an external
 * workbook, INDIRECT) stays un-ignored, same convention as totals-OFF's `unobsOwnRowOnly`.
 */
function refAxisIgnore(refs) {
  const axes = new Set(refs.map(r => { const i = r.indexOf('!'); return `formula:${r.slice(0, i)}:${r.slice(i + 1)}`; }));
  return u => axes.has(u.axis);
}

/** Either predicate forgiving the entry. */
const combineIgnore = (a, b) => u => a(u) || b(u);

/** A fresh `Name[` matcher (word-boundary on the left, same class `structuredRefs()` uses) -- 'g' so matchAll() works. */
const nameBracketRe = () => /(^|[^A-Za-z0-9_.À-￿])([A-Za-z_À-￿][\w.À-￿]*)\s*\[/g;

/**
 * GH #332/#333 (a LIVE review of the forgiveness below): every `Name[` identifier found in `formula`'s 'other'
 * tokens -- OUTSIDE string literals and quoted sheet names, via formula.cjs's own `tokenize()` -- plus
 * whether a literal/quoted-name token ALSO holds one (e.g. INDIRECT("Table1[Qty]")): that shape is the
 * same 'cannot safely reach it' case tableRenameRewrite's sawInLiteral names, so it is reported and the
 * caller never forgives on it, regardless of which name it names. `null` for an unparsable formula
 * (unterminated literal) or one with no leading '=' -- the caller already filtered those via sheetRefs.
 */
function namedBracketRefsOutsideLiterals(formula) {
  const toks = tokenize(String(formula).replace(/^=/, ''));
  if (!toks) return null;
  const names = [];
  for (const t of toks) {
    if (t.kind !== 'other') {
      if (nameBracketRe().test(t.text)) return { names: null, sawInLiteral: true };
      continue;
    }
    for (const m of t.text.matchAll(nameBracketRe())) names.push(m[2]);
  }
  return { names, sawInLiteral: false };
}

/**
 * GH #332/#333: a dependencyConflict ignore-predicate that ALSO forgives a structured reference to a
 * DIFFERENT, NAMED table elsewhere in the workbook. MEASURED (the GH #332/#333 structured-ref-rewrite
 * research probe, prod -- a `[current]`-rename and a `[t2]`-resize seeded side by side): the dependency
 * reader (dependencies.cjs) marks EVERY structured-reference formula in the WHOLE workbook unobservable,
 * not just the ones naming THIS table -- so an unrelated table's own dependent formula (one that cannot
 * possibly be affected by renaming/resizing THIS table) blocked rename/resize too, on both operations.
 * Forgiven ONLY when the formula text:
 *   - is NOT an external-workbook reference (formula.sheetRefs' `external` flag, the same bracket-before-
 *     '!' shape that already tells it apart from a structured reference) -- an external link stays
 *     un-ignored, unchanged;
 *   - holds NO `Name[` inside a string literal / quoted sheet name (a dynamic reference text the
 *     structural read cannot see at all, e.g. INDIRECT("Table1[Qty]") -- MEASURED: an earlier version that
 *     scanned the whole formula text, literals included, wrongly forgave this shape when it also named a
 *     known other table elsewhere);
 *   - names, before every '[' found OUTSIDE a literal, EITHER one of the OTHER tables the current read
 *     found (on ANY sheet -- a table name is workbook-unique, so cross-sheet is fine, matching
 *     `structuredRefs`' own detection) OR `ownNames` -- a formula naming even ONE unresolvable identifier
 *     (e.g. `=SUM(t2[Qty])+Ghost[Qty]`) stays un-ignored, same as today's single-reference
 *     case. An ambiguous bracket that names no known OTHER table (e.g. a bare `[@Qty]`/`[#This Row]`
 *     relative reference, whose own table cannot be told from the text alone) is NOT a match, so a
 *     formula with ZERO resolved names also stays un-ignored: fail-closed, the same convention as
 *     totals-OFF's `unobsOwnRowOnly`.
 */
function otherTableAxisIgnore(p, ownNames) {
  const own = new Set(ownNames.map(n => String(n).toLowerCase()));
  const others = new Set(
    Object.values(p.tableReads || {})
      .filter(r => r && Array.isArray(r.tables))
      .flatMap(r => r.tables.map(t => t.name))
      .map(n => String(n).toLowerCase())
      .filter(n => !own.has(n))
  );
  if (!others.size) return () => false;
  const known = new Set([...own, ...others]);
  const byAxis = new Map(formulaEntries(p).map(([k, f]) => [`formula:${k.replace('\u0000', ':')}`, f]));
  return u => {
    const f = byAxis.get(u.axis);
    if (typeof f !== 'string' || !f.startsWith('=')) return false;
    const sr = sheetRefs(f);
    if (!sr.ok || sr.external) return false;
    const scan = namedBracketRefsOutsideLiterals(f);
    if (!scan || scan.sawInLiteral || !scan.names.length) return false;
    return scan.names.every(n => known.has(String(n).toLowerCase()));
  };
}

/**
 * GH #332: for each cell `refs` named, a rewritten formula text with the table's old structured
 * reference renamed to the new name (formula.cjs's tableRenameRewrite -- string literals and quoted
 * sheet names kept verbatim, so an INDIRECT("Old[Qty]") text is never touched). `null` -> at least one
 * formula names the OLD name only inside a literal (the rename cannot safely reach it); the caller
 * refuses the whole rename, unchanged. A ref with NO match at all (e.g. an in-table `[@Qty]` relative
 * reference, which never embeds the table name) is simply skipped -- nothing to rewrite there.
 */
function buildRenameRewrites(p, refs, fromName, toName) {
  const want = new Set(refs);
  const out = [];
  for (const [k, f] of formulaEntries(p)) {
    const addr = k.replace('\u0000', '!');
    if (!want.has(addr)) continue;
    const r = tableRenameRewrite(f, fromName, toName);
    if (!r || r.sawInLiteral) return null;
    if (r.changed) out.push({ address: addr, from: f, to: r.formula });
  }
  return out;
}

function resolve(p, selector, ctx = {}) {
  const t = gate(p, ctx.operation);
  if (!t.ok) {
    const op = ctx.operation;
    if (op && op.intent === 'rename_table' && t.reason === 'xlsx-table-not-found' && t.find(op.want.name).length === 1) {
      const n = t.find(op.want.name)[0];
      return { ok: true, target: Object.freeze({ sheet: selector.sheet, table: n.name, ref: n.ref }) };
    }
    return { ok: false, target: null, selector, ambiguity: t.reason, detail: t };
  }
  return { ok: true, target: Object.freeze({ sheet: selector.sheet, table: t.table.name, ref: t.table.ref }) };
}

/** One callCommand: the exact-name sheet check, then ONE sheet.table.write/v1 Bridge write (the driver prelude supplies the gate / write / fail helpers). */
function lower(op, target) {
  const w = op.want;
  const base = { sheet: op.selector.sheet, table: op.selector.table };
  const args = {
    set_table_style: { ...base, property: 'style', name: w.style },
    set_table_flag: { ...base, property: 'flag', flag: w.flag, value: w.value },
    set_table_totals_row: { ...base, property: 'totalsRow', value: w.value },
    rename_table: { ...base, property: 'name', name: w.name },
    resize_table: { ...base, property: 'ref', ref: w.ref },
  }[op.intent];
  // GH #332 (owner decision A): rename_table's dependent-formula rewrite rides in the SAME callCommand,
  // right after the table rename itself succeeds -- a separate task/session would leave a window where
  // the table is renamed but the dependent formulas still name the old one. `op._renameFormulaRewrites`
  // is built in classify() (the SAME op-mutation convention adapter.cjs's own `_protectionRaceCheck`
  // already uses); each entry's `to` is written per-cell via SetValue (the measured formula-write API --
  // cells.cjs, families/range.cjs:409 -- NOT a `SetFormula` method, which this engine does not expose).
  const rewrites = op.intent === 'rename_table' ? (op._renameFormulaRewrites || []) : [];
  const rewriteStep = rewrites.length ? `
  for (const rw of ${JSON.stringify(rewrites)}) {
    const sep = rw.address.indexOf('!');
    const rsheet = Api.GetSheet(rw.address.slice(0, sep));
    if (!rsheet) return { ok: false, stage: 'rename-rewrite-sheet-missing', address: rw.address };
    try { rsheet.GetRange(rw.address.slice(sep + 1)).SetValue(rw.to); }
    catch (rerr) { return { ok: false, stage: 'rename-rewrite-threw', address: rw.address, error: String(rerr) }; }
  }` : '';
  return {
    kind: 'callCommand',
    body: `
try {
  const name = ${JSON.stringify(op.selector.sheet)};
  const ws = Api.GetSheet(name);
  if (!ws || ws.GetName() !== name) return { ok: false, stage: 'sheet-identity' };
  const g = __eoTableGate();
  if (!g.ok) return g;
  const w = __eoTableWrite(${JSON.stringify(args)});
  if (!w.ok) return __eoTableFail(w);${rewriteStep}
  return { ok: true };
} catch (err) { return { ok: false, stage: 'threw', error: String(err) }; }
`,
    meta: Object.freeze({
      family: 'table-props', intent: op.intent, sheet: op.selector.sheet, table: op.selector.table, property: PROPERTY[op.intent], want: w, ref: String(target && target.ref), bridgeTableWrite: true,
      ...(rewrites.length ? { renameFormulaRewrites: rewrites } : {}),
    }),
  };
}

// GH #361 2nd unit (bridge /28): restOf()/gate() work on the RAW
// `inv.tableReads[sheet].tables` entries (xlsx-driver.cjs's own read, same shape inventory.cjs
// forwards) -- NOT the read_tables projection. On that raw shape a table carrying its own
// AutoFilter gets `autoFilterDetail` (the candidate sheet.table.autoFilter.criteria.read/v1
// result, `{ present, range, fields }` when readable, a named 'UNOBSERVABLE:...' string or
// absent otherwise -- see tables-read.cjs's own autoFilterCriteriaOf(), which only exists to turn
// THIS field into the `autoFilterCriteria` the read_tables projection exposes; the raw object
// gate() sees never carries `autoFilterCriteria` at all). `range` moves WITH the table's own
// `ref` (same cell range, same text); `fields` (the filter conditions) do not change on a
// ref-only move. A ref-changing intent must therefore exclude `autoFilterDetail.range` from
// restOf()'s comparison the SAME way it excludes `ref` itself, while KEEPING `.fields` (and any
// other detail field) compared (a resize/totals-row toggle must not silently swallow an
// unrelated filter-condition change). String-shaped or absent autoFilterDetail (every runtime
// before /28, or /28 itself when the candidate capability did not produce a clean read) is
// untouched -- the `typeof` guard means behaviour there stays byte-identical to before this fix.
function normAutoFilterRange(c) {
  if (c.autoFilterDetail && typeof c.autoFilterDetail === 'object') c.autoFilterDetail.range = null;
}

/** the table without the fields the property is ALLOWED to change (a stable string) */
function restOf(t, op) {
  const c = JSON.parse(JSON.stringify(t));
  const p = op.intent;
  if (p === 'set_table_style') { if (c.style) c.style.name = null; } else if (p === 'set_table_flag') { if (c.style) c.style[op.want.flag] = null; } else if (p === 'set_table_totals_row') { c.totalsRow = null; c.ref = null; c.autoFilter = null; normAutoFilterRange(c); } else if (p === 'rename_table') { c.name = null; } else { c.ref = null; normAutoFilterRange(c); }
  return JSON.stringify(c);
}

function targetOf(p, op) {
  const t = gate(p, op);
  if (t.ok) return t;
  if (op.intent === 'rename_table' && t.reason === 'xlsx-table-not-found') {
    const now = t.find(op.want.name);
    if (now.length === 1) return { ok: true, tables: t.tables, table: now[0], g: parseRange(String(now[0].ref)), find: t.find };
  }
  return t;
}

function verifyPrimitive(op, before, after) {
  const b = gate(before && before.primary, op);
  const a = targetOf(after && after.primary, op);
  if (!b.ok || !a.ok) return { ok: false, evidence: { reason: 'xlsx-table-props-unobservable', before: b.ok ? null : b.reason, after: a.ok ? null : a.reason } };
  const w = op.want;
  const t = a.table;
  let exp = true;
  let observed;
  if (op.intent === 'set_table_style') { observed = t.style ? t.style.name : null; exp = observed === w.style; } else if (op.intent === 'set_table_flag') { observed = t.style ? t.style[w.flag] : null; exp = observed === w.value; } else if (op.intent === 'set_table_totals_row') {
    observed = { totalsRow: t.totalsRow, ref: t.ref };
    exp = w.value === true
      ? t.totalsRow === true && normRef(t.ref) === rangeText({ r1: b.g.r1, c1: b.g.c1, r2: b.g.r2 + 1, c2: b.g.c2 })
      : t.totalsRow !== true && normRef(t.ref) === rangeText({ r1: b.g.r1, c1: b.g.c1, r2: b.g.r2 - 1, c2: b.g.c2 });
  } else if (op.intent === 'rename_table') { observed = t.name; exp = t.name === w.name; } else { observed = t.ref; exp = normRef(t.ref) === w.ref; }
  const restKept = restOf(t, op) === restOf(b.table, op);
  const othersKept = JSON.stringify(a.tables.filter(x => x !== t)) === JSON.stringify(b.tables.filter(x => x !== b.table));
  // GH #332 (owner decision A): the receipt names which cells the rename's rewrite touched -- this is
  // the ONLY layer whose evidence survives into appliedOperations[0].evidence for an APPLIED step
  // (classify()'s own evidence, where op._renameFormulaRewrites was computed, is a planning-time
  // artefact and is not otherwise re-surfaced).
  const formulasRewritten = op.intent === 'rename_table' && Array.isArray(op._renameFormulaRewrites) && op._renameFormulaRewrites.length
    ? op._renameFormulaRewrites.map(r => r.address) : undefined;
  return { ok: exp && restKept && othersKept, evidence: { sheet: op.selector.sheet, table: t.name, property: PROPERTY[op.intent], expected: op.want, observed, propertyOk: exp, restOfTableKept: restKept, otherTablesKept: othersKept, proof: 'fresh-live-read_tables', ...(formulasRewritten ? { formulasRewritten } : {}) } };
}

/** every cell outside the allowed region unchanged, other tables kept, no new error value */
function verifyPreservation(op, before, after) {
  const b = gate(before && before.primary, op);
  const a = targetOf(after && after.primary, op);
  if (!b.ok || !a.ok) return { ok: false, evidence: { reason: 'xlsx-table-props-preservation-unobservable', before: b.ok ? null : b.reason, after: a.ok ? null : a.reason } };
  const sheet = op.selector.sheet;
  let region = null;
  if (op.intent === 'set_table_totals_row') region = op.want.value === true ? { r1: b.g.r2 + 1, r2: b.g.r2 + 1, c1: b.g.c1, c2: b.g.c2 } : { r1: b.g.r2, r2: b.g.r2, c1: b.g.c1, c2: b.g.c2 };
  if (op.intent === 'resize_table') { const want = parseRange(op.want.ref); region = { r1: Math.min(b.g.r2, want.r2) + 1, r2: Math.max(b.g.r2, want.r2), c1: b.g.c1, c2: b.g.c2 }; }
  const inRegion = (s, g) => region && s === sheet && g && within(region, g.r1, g.c1);
  // GH #332/#333 (owner decision A): a structured reference to this table that resize left in place, or
  // that rename rewrote, can read '' in-session right after the Bridge write -- MEASURED LIVE (the GH
  // #332/#333 research probe, prod): the engine does not recalculate the dependent cell's VALUE until
  // the file is saved and reopened, while the STORED bytes are already correct by then (whether the value is actually
  // right is the stored-content gate's job -- the LIVE acceptance test, GH #141 model -- not this
  // in-session check). The FORMULA TEXT is not delayed the same way: SetValue writes it synchronously,
  // so it is still checked here -- resize expects it UNCHANGED, rename expects the REWRITTEN text (the
  // check a skipped/broken rewrite still fails on).
  const expectedFormulaAt = new Map(); // raw key -> the <f> text this address must hold (its value is excused)
  if (op.intent === 'resize_table' && Array.isArray(op._resizeStructuredRefs)) {
    for (const addr of op._resizeStructuredRefs) {
      const i = addr.indexOf('!'); const k = `${addr.slice(0, i)}\u0000${addr.slice(i + 1)}`;
      const was = (before.primary.cells || {})[k];
      if (was) expectedFormulaAt.set(k, was.f);
    }
  }
  if (op.intent === 'rename_table' && Array.isArray(op._renameFormulaRewrites)) {
    for (const rw of op._renameFormulaRewrites) {
      const i = rw.address.indexOf('!');
      expectedFormulaAt.set(`${rw.address.slice(0, i)}\u0000${rw.address.slice(i + 1)}`, rw.to);
    }
  }
  const pb = before.primary; const pa = after.primary;
  const changed = [];
  for (const [k, c] of Object.entries(pb.cells || {})) {
    const [s, a1] = k.split('\u0000');
    const g = parseRange(a1);
    if (expectedFormulaAt.has(k)) {
      const now = pa.cells ? pa.cells[k] : undefined;
      if (!now || now.f !== expectedFormulaAt.get(k)) changed.push(`${s}!${a1}:formula`);
      continue;
    }
    if (inRegion(s, g)) continue;
    const now = pa.cells ? pa.cells[k] : undefined;
    if (!now) { if (!blank(c)) changed.push(`${s}!${a1}:gone`); continue; }
    if (now.f !== c.f || String(now.v) !== String(c.v)) changed.push(`${s}!${a1}`);
  }
  for (const [k, c] of Object.entries(pa.cells || {})) {
    if (pb.cells && pb.cells[k]) continue;
    if (blank(c)) continue;
    const [s, a1] = k.split('\u0000');
    if (!inRegion(s, parseRange(a1)) && !expectedFormulaAt.has(k)) changed.push(`${s}!${a1}:appeared`);
  }
  const errs = pp => Object.entries(pp.cells || {}).filter(([k, c]) => !expectedFormulaAt.has(k) && (/^#(NAME\?|REF!)/i.test(String(c.v)) || /#REF!/i.test(String(c.f || '')))).map(([k]) => k.replace('\u0000', '!'));
  const was = new Set(errs(pb));
  const newErrors = errs(pa).filter(x => !was.has(x));
  return { ok: changed.length === 0 && newErrors.length === 0, evidence: { cellsChangedOutsideTheRegion: changed.slice(0, 20), newErrorValues: newErrors.slice(0, 20), region: region ? rangeText(region) : null, staleValueExcludedFromRegion: expectedFormulaAt.size ? [...expectedFormulaAt.keys()].map(k => k.replace('\u0000', '!')) : undefined } };
}

function wholeTaskChecks(op, finalObservation) {
  const t = targetOf(finalObservation && finalObservation.primary, op);
  const w = op.want;
  let ok = false;
  let label;
  if (t.ok) {
    if (op.intent === 'set_table_style') { ok = !!t.table.style && t.table.style.name === w.style; label = `style is ${w.style}`; } else if (op.intent === 'set_table_flag') { ok = !!t.table.style && t.table.style[w.flag] === w.value; label = `${w.flag} is ${w.value}`; } else if (op.intent === 'set_table_totals_row') { ok = t.table.totalsRow === w.value; label = w.value ? 'has a totals row' : 'has no totals row'; } else if (op.intent === 'rename_table') { ok = t.table.name === w.name; label = `is named ${w.name}`; } else { ok = normRef(t.table.ref) === w.ref; label = `has the range ${w.ref}`; }
  } else label = 'is readable';
  return [{ name: `${op.selector.sheet} table ${op.selector.table} ${label}`, ok, detail: ok ? null : { reason: t.reason || 'differs', current: t.ok ? t.table : null } }];
}

const windowsOf = () => [];
// GH #332/#333 (MEASURED LIVE): tableSheets named only the operation's own sheet, so otherTableAxisIgnore's
// `others` set (built from p.tableReads) was always empty for a dependent formula naming a table on a
// DIFFERENT sheet -- '*' matches the same driver convention objectSheets already uses (every sheet).
const liveExtras = op => ({ tableSheets: ['*'], objectSheets: [op.selector.sheet], dependents: true });

module.exports = { INTENTS, PROPERTY, FLAGS, B25_BRIDGE_WRITE_RUNTIMES, TOTALS_OFF_BRIDGE_RUNTIMES, plan, classify, resolve, lower, verifyPrimitive, verifyPreservation, wholeTaskChecks, windowsOf, liveExtras, styleOk, nameOk };
