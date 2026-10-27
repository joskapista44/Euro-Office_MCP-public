'use strict';

const { XLSX_B03_BRIDGE_WRITE_RUNTIME_IDS } = require('../runtime.cjs');
const { isBlank, key } = require('../inventory.cjs');
const { refOnTarget } = require('../dependencies.cjs');
const formula = require('../formula.cjs');
const { sheetNameProblem } = require('../sheet-name.cjs');
const { copyMatches, objectsComplete, brokenNames, refKept, objectDiffs } = require('../sheet-copy.cjs');
const drawings = require('../drawings.cjs');

/**
 * XLSX `sheets` family: create_sheet, rename_sheet, delete_sheet, move_sheet (M6 wave 3).
 *
 * The primitives are MEASURED (wave-3 LIVE probe, run m6-w3p-1790584404615-f856e6 at 4ed3489,
 * docs/evidence/M6-W3-SHEET-PROBE.md): PUBLIC Api.AddSheet / ApiWorksheet.SetName / Delete / Move;
 * there is no public Copy (the source used the internal asc_copyWorksheet) -- copy_sheet is refused.
 *   - AddSheet: visible in the same command; a duplicate exact name THROWS; but an invalid name
 *     ('bad/x') is accepted AS IS and a 32-character name silently becomes 'Sheet2' -- both ok, both
 *     result null. So the name is validated BEFORE dispatch and the created name is proven exactly.
 *   - SetName: a formula referencing the sheet follows the rename (value unchanged).
 *   - Delete: a referencing formula becomes '= #REF!*2' / '#REF!' and stays so after a reopen.
 *   - Move(before, undefined) / Move(undefined, after): both directions work.
 *
 * Contract, where it differs from the source (audit rows B-S05..B-S09, B-S26, B-S27):
 *   - sheet names are limited to the class the formula identity was MEASURED on (L-21/2, sheet-name.cjs: letters incl.
 *     accents, digits, underscore, words joined by single spaces, not cell- or R1C1-shaped, at most 31 characters; an
 *     apostrophe is refused by name -- the engine mishandles it); the source left name validity to the runtime
 *     (NOT_MEASURED there, silently wrong here);
 *   - uniqueness is case-insensitive (a formula naming a sheet with a case-only difference is silently
 *     dropped by the editor -- W2 probe set b);
 *   - create is SATISFIED only by an EMPTY sheet of that exact name (the source accepted any sheet);
 *   - delete refuses a sheet that a formula on another sheet references (the source had no guard;
 *     the probe measured the silent #REF!), AND refuses before any write unless the observation carries
 *     a COMPLETE, fresh inventory of the other dependents naming none of them on the sheet (owner decision
 *     2026-09-28: an incomplete dependency check is not a limit): defined names (workbook and sheet
 *     level, and a formula THROUGH a name), validations, conditional formats, chart series, pivot
 *     sources (dependencies.cjs); a reader gap or an unresolvable / broken / ambiguous reference refuses;
 *   - a sheet operation is the only step of its task (no rename graphs, no delete-after-write; the
 *     source's batch combinations belong to composition, W9);
 *   - L3: every cell of every other sheet keeps its formula (a rename rewrites the references, and
 *     only those) and a non-formula cell keeps its value; the sheet list changes exactly as asked.
 * Retry = the terminal state on a fresh read: rename with the source gone and the target present, and
 * delete with the sheet absent, are SATISFIED -- a fresh read cannot tell them from "never existed"
 * (the same limit as the source, stated).
 */

const INTENTS = Object.freeze(['create_sheet', 'rename_sheet', 'delete_sheet', 'move_sheet', 'copy_sheet']);

/**
 * M11.5c.4 (owner: M11.5 continuous mandate, internal tracker): copy_sheet writes ONLY through the DS Bridge WRITE
 * worksheet.copy/v1, whose runtime list is /8 .. /12 (L-05; the driver keeps the same list as its second line; a static test compares
 * them). The former EO_INTERNAL_COPY_RUNTIMES direct-internal path is gone; any other runtime refuses before the write.
 */
const B03_BRIDGE_WRITE_RUNTIMES = XLSX_B03_BRIDGE_WRITE_RUNTIME_IDS;
const FIELDS = Object.freeze({
  create_sheet: ['intent', 'id', 'name'],
  rename_sheet: ['intent', 'id', 'sheet', 'name'],
  delete_sheet: ['intent', 'id', 'sheet'],
  move_sheet: ['intent', 'id', 'sheet', 'before', 'after'],
  copy_sheet: ['intent', 'id', 'sheet', 'name'],
});

/** null when the name is inside the released (measured) class, otherwise why not -- the ONE rule, adapters/xlsx/sheet-name.cjs (L-21/2). */
const nameProblem = sheetNameProblem;

/** Per intent: the selector from a step, or why it is refused (names checked by nameProblem). */
const SELECTORS = {
  create_sheet: st => ({ name: ['name', st.name] }),
  rename_sheet: st => (typeof st.name === 'string' && typeof st.sheet === 'string' && st.name.toLowerCase() === st.sheet.toLowerCase()
    ? 'the new name equals the old one, case-insensitively'
    : { sheet: ['sheet', st.sheet], name: ['name', st.name] }),
  delete_sheet: st => ({ sheet: ['sheet', st.sheet] }),
  copy_sheet: st => (typeof st.name === 'string' && typeof st.sheet === 'string' && st.name.toLowerCase() === st.sheet.toLowerCase()
    ? 'the copy name equals the source name, case-insensitively'
    : { sheet: ['sheet', st.sheet], name: ['name', st.name] }),
  move_sheet: st => {
    if ((st.before !== undefined) === (st.after !== undefined)) return 'exactly one of before / after must name the reference sheet';
    const where = st.before !== undefined ? 'before' : 'after';
    const ref = st[where];
    if (typeof ref === 'string' && typeof st.sheet === 'string' && ref.toLowerCase() === st.sheet.toLowerCase()) return 'a sheet cannot be moved relative to itself';
    return { sheet: ['sheet', st.sheet], reference: [where, ref], where: [null, where] };
  },
};

function plan(step, index) {
  const intent = step.intent;
  const unknown = Object.keys(step).filter(k => !FIELDS[intent].includes(k));
  const refuse = why => ({ ok: false, detail: `step ${index}: ${intent} refused before dispatch: ${why}` });
  if (unknown.length) return refuse(`unsupported field(s): ${unknown.join(', ')}`);
  const spec = SELECTORS[intent](step);
  if (typeof spec === 'string') return refuse(spec);
  const selector = {};
  for (const [k, [field, value]] of Object.entries(spec)) {
    const why = field === null ? null : nameProblem(value);
    if (why) return refuse(`${field} ${JSON.stringify(value)}: ${why}`);
    selector[k] = value;
  }
  return {
    ok: true,
    operation: { id: step.id || `s${index}`, intent, family: 'sheets', selector: Object.freeze(selector), preservationMode: 'PRIMITIVE_INTEGRAL' },
  };
}

const namesOf = primary => primary.sheets.map(s => s.name);
const exact = (primary, n) => primary.sheets.filter(s => s.name === n);
const folded = (primary, n) => primary.sheets.filter(s => s.name.toLowerCase() === String(n).toLowerCase());
const sheetEmpty = (primary, n) => Object.entries(primary.cells).every(([k, c]) => !k.startsWith(`${n}\u0000`) || isBlank(c));

/** Formula cells on OTHER sheets that reference `n` (case-insensitive: the editor folds case). */
function referencesTo(primary, n) {
  const out = [];
  for (const [k, c] of Object.entries(primary.cells)) {
    const [sheet, a] = k.split('\u0000');
    if (sheet === n || c.f === '') continue;
    if (formula.sheetRefs(c.f).names.some(x => x.toLowerCase() === n.toLowerCase())) out.push(`${sheet}!${a}`);
  }
  return out;
}

const fail = (reason, extra = {}) => ({ state: 'FAIL_CLOSED', evidence: { reason, ...extra } });

const CLASSIFY = {
  copy_sheet(primary, s, available) {
    // M7 R26c: only create / rename (cell formulas) / move / delete are measured for a name that needs quotes
    if (!formula.bareSheetName(s.sheet) || !formula.bareSheetName(s.name)) return fail('xlsx-copy-quoted-name-not-measured', { sheet: s.sheet, name: s.name });
    const blocked = copyBlocked(primary, s, available);
    if (blocked) return blocked;
    const tgt = exact(primary, s.name).length;
    if (tgt > 1 || folded(primary, s.name).some(x => x.name !== s.name)) return fail('xlsx-sheet-name-case-collision', { name: s.name, available });
    if (tgt === 1) {
      const m = copyMatches(primary, s.sheet, s.name);
      // the desired end state: a sheet of that name, LAST, equal to the source (cells, objects, rules).
      // Met -> SATISFIED, but NOT this run's copy: provenance is not proven (no retry token by design).
      if (m.ok && available[available.length - 1] === s.name) return { state: 'SATISFIED', evidence: { sheet: s.name, proof: 'fresh-live-end-state-equal-to-the-desired-copy', provenance: 'NOT_PROVEN' } };
      return fail('xlsx-copy-target-exists-and-differs', { name: s.name, last: available[available.length - 1] === s.name, diffs: m.diffs, count: m.count });
    }
    return { state: 'APPLY', evidence: { copy: s.sheet, to: s.name } };
  },
  create_sheet(primary, s, available) {
    const e = exact(primary, s.name);
    if (e.length > 1) return fail('xlsx-sheet-ambiguous', { available });
    if (e.length === 1) {
      return sheetEmpty(primary, s.name)
        ? { state: 'SATISFIED', evidence: { sheet: s.name, proof: 'fresh-live-empty-sheet-of-that-name' } }
        : fail('xlsx-create-name-taken-by-a-non-empty-sheet', { sheet: s.name });
    }
    if (folded(primary, s.name).length) return fail('xlsx-sheet-name-case-collision', { name: s.name, available });
    return { state: 'APPLY', evidence: { create: s.name } };
  },
  rename_sheet(primary, s, available) {
    const src = exact(primary, s.sheet).length;
    const tgt = exact(primary, s.name).length;
    if (tgt > 1 || folded(primary, s.name).some(x => x.name !== s.name)) return fail('xlsx-sheet-name-case-collision', { name: s.name, available });
    if (src === 1 && tgt === 1) return fail('xlsx-rename-source-and-target-both-present', { available });
    if (src === 0 && tgt === 1) return { state: 'SATISFIED', evidence: { sheet: s.name, proof: 'fresh-live-source-absent-target-present', limit: 'cannot tell a done rename from a pre-existing sheet of that name' } };
    if (src === 0) return missing(primary, s.sheet, available);
    return renameBlocked(primary, s) || { state: 'APPLY', evidence: { rename: s.sheet, to: s.name } };
  },
  delete_sheet(primary, s, available) {
    if (exact(primary, s.sheet).length === 0) {
      return folded(primary, s.sheet).length
        ? fail('xlsx-sheet-case-mismatch', { available })
        : { state: 'SATISFIED', evidence: { sheet: s.sheet, proof: 'fresh-live-sheet-absent', limit: 'cannot tell a done delete from a sheet that never existed' } };
    }
    if (available.length < 2) return fail('xlsx-delete-last-sheet-refused');
    return deleteBlocked(primary, s.sheet) || { state: 'APPLY', evidence: { delete: s.sheet } };
  },
  move_sheet(primary, s, available) {
    if (exact(primary, s.sheet).length !== 1) return missing(primary, s.sheet, available);
    if (exact(primary, s.reference).length !== 1) return fail('xlsx-move-reference-not-found', { reference: s.reference, available });
    const i = available.indexOf(s.sheet);
    const j = available.indexOf(s.reference);
    if (i === j + (s.where === 'before' ? -1 : 1)) return { state: 'SATISFIED', evidence: { order: available, proof: 'fresh-live-already-adjacent' } };
    return { state: 'APPLY', evidence: { from: i, where: s.where, reference: s.reference } };
  },
};

/**
 * delete_sheet's refusals before any write: a formula on another sheet names it; any INDIRECT (a text
 * reference no parser resolves); the dependency observation incomplete (a reader gap or an
 * unresolvable / broken / ambiguous reference -- owner decision 2026-09-28); any dependent naming it.
 */
function deleteBlocked(primary, n) {
  const refs = referencesTo(primary, n);
  if (refs.length) return fail('xlsx-delete-sheet-referenced-by-formulas', { sheet: n, references: refs.slice(0, 20), count: refs.length });
  const indirect = Object.entries(primary.cells).filter(([, c]) => /\bINDIRECT\s*\(/i.test(c.f)).map(([k]) => k.replace('\u0000', '!'));
  if (indirect.length) return fail('xlsx-delete-indirect-reference-unresolvable', { sheet: n, cells: indirect.slice(0, 20) });
  const d = primary.dependencies;
  if (!d || !d.complete) return fail('xlsx-delete-dependency-observation-incomplete', { sheet: n, unobservable: d ? d.unobservable.slice(0, 20) : ['dependencies not read'], count: d ? d.unobservable.length : null });
  const ref = refOnTarget(d, [n]);
  if (ref) return fail(ref.reason, { sheet: n, ...ref });
  // a dependent hosted ON the deleted sheet goes with it -- except a sheet-scoped name (unmeasured: refused)
  const same = x => x && x.toLowerCase() === n.toLowerCase();
  const deps = d.refs.filter(r => !(r.kind !== 'name' && same(r.host)) && r.sheets.some(same));
  return deps.length ? fail('xlsx-delete-sheet-referenced-by-dependents', { sheet: n, dependents: deps.slice(0, 20) }) : null;
}

/**
 * rename_sheet's refusals before any write (owner, owner channel 2026-09-28): measured (W3f run 2,
 * docs/evidence/M6-W3F-COPY-RULES-PROBE.md) a conditional format referencing the renamed sheet is NOT
 * rewritten -- not even the sheet's own -- so the rename would leave it naming a sheet that no longer
 * exists. Every conditional format of the workbook is read fresh (a name resolved to its sheets); an
 * incomplete read or an unresolvable dependent refuses too.
 */
function renameBlocked(primary, s) {
  const d = primary.dependencies;
  if (!d || !d.complete) return fail('xlsx-rename-dependency-observation-incomplete', { sheet: s.sheet, unobservable: d ? d.unobservable.slice(0, 20) : ['dependencies not read'], count: d ? d.unobservable.length : null });
  const ref = refOnTarget(d, [s.sheet]);
  if (ref) return fail(ref.reason, { sheet: s.sheet, ...ref });
  const same = x => x && x.toLowerCase() === s.sheet.toLowerCase();
  const cf = d.refs.filter(r => r.kind === 'conditional' && r.sheets.some(same));
  if (cf.length) return fail('xlsx-rename-conditional-format-reference-not-followed', { sheet: s.sheet, rules: cf.slice(0, 20), count: cf.length });
  // measured (W3g): a pivot's source follows in the session, but after save + reopen GetSource throws --
  // for a non-renamed control pivot too: an OBSERVATION LIMIT (no engine defect proven). The renamed
  // state cannot be read back, so the rename is refused (a chart series follows: allowed)
  const pv = d.refs.filter(r => r.kind === 'pivot' && r.sheets.some(same));
  if (pv.length) return fail('xlsx-rename-pivot-source-unobservable-after-reopen', { sheet: s.sheet, pivots: pv.slice(0, 20), count: pv.length });
  // M7 R26c: for a name that needs quotes only a CELL formula's following is measured -- a validation, a defined name
  // or a chart series referencing such a sheet is not
  if (formula.bareSheetName(s.sheet) && formula.bareSheetName(s.name)) return null;
  const quoted = d.refs.filter(r => r.kind !== 'formula' && r.sheets.some(same));
  return quoted.length ? fail('xlsx-rename-quoted-name-dependent-not-measured', { sheet: s.sheet, refs: quoted.slice(0, 20), count: quoted.length }) : null;
}

/** copy_sheet's refusals before any write (runtime identity, source, every unobservable axis), or null. */
function copyBlocked(primary, s, available) {
  const runtime = primary.runtime ? primary.runtime.identity : null;
  if (!B03_BRIDGE_WRITE_RUNTIMES.includes(runtime)) return fail('xlsx-copy-bridge-write-runtime-not-proven', { runtime, proven: B03_BRIDGE_WRITE_RUNTIMES });
  if (exact(primary, s.sheet).length !== 1) return missing(primary, s.sheet, available);
  // an unobservable axis refuses BEFORE the write (owner rule): the copy could not be proven complete
  if (!primary.rules || !primary.rules.complete) return fail('xlsx-copy-rule-observation-incomplete', { unobservable: primary.rules ? primary.rules.unobservable.slice(0, 20) : ['rules not read'] });
  // measured (W3f run 1, docs/evidence/M6-W3F-COPY-RULES-PROBE.md): asc_copyWorksheet leaves a validation /
  // conditional-format reference to the source sheet ITSELF pointing at the source -- also after a reopen --
  // while it maps the cells' and names' self references. Such a copy cannot be correct: refused before the write.
  const selfRules = selfReferencingRules(primary, s.sheet);
  if (selfRules.length) return fail('xlsx-copy-rule-self-reference-not-mapped', { sheet: s.sheet, rules: selfRules.slice(0, 10), count: selfRules.length });
  // M7 S1 made an xlCellValue rule OBSERVABLE (the read), not copyable: what asc_copyWorksheet does with one hosted on
  // the copied sheet is not measured -> still refused before the write (only xlExpression was measured in W3f)
  const unmeasuredCopy = (primary.rules.bySheet[s.sheet] || []).filter(r => r.kind === 'conditional' && r.type !== 'xlExpression');
  if (unmeasuredCopy.length) return fail('xlsx-copy-conditional-type-not-measured-for-copy', { sheet: s.sheet, rules: unmeasuredCopy.slice(0, 10), count: unmeasuredCopy.length });
  if (!objectsComplete(primary.objects[s.sheet])) return fail('xlsx-copy-source-objects-unobservable', { objects: primary.objects[s.sheet] || null });
  // the defined names must be observable, and nothing on the source may be unresolvable
  const d = primary.dependencies;
  // ...and every axis the L3 compares on the OTHER sheets must be observable: a reader gap there would be
  // compared as "unchanged" on both sides (a pivot on another sheet: unreadable after a reopen, W3g)
  const l3Axis = a => a === 'defined-names' || a === 'data-validation' || /^(objects|charts|pivots|names|conditional-formatting):/.test(a);
  const gap = !d ? ['dependencies not read'] : d.unobservable.filter(u => l3Axis(u.axis) || (u.host && u.host.toLowerCase() === s.sheet.toLowerCase()));
  if (gap.length) return fail('xlsx-copy-source-dependency-unobservable', { unobservable: gap.slice(0, 20) });
  // the copied sheet carries the source's texts: a #REF! there is copied in an unmeasured way
  const ref = d ? refOnTarget(d, [s.sheet]) : null;
  if (ref) return fail(ref.reason, { sheet: s.sheet, ...ref });
  const broken = brokenNames(primary, s.sheet);
  return broken.length ? fail('xlsx-copy-source-carries-broken-name', { names: broken }) : null;
}

/** The source's validation / conditional-format rules whose formula (1 or 2) names the source sheet itself. */
function selfReferencingRules(primary, n) {
  const names = f => (f === null || f === undefined || f === '' ? [] : formula.sheetRefs(`=${String(f).replace(/^=/, '')}`).names);
  const own = f => names(f).some(x => x.toLowerCase() === n.toLowerCase());
  return (primary.rules.bySheet[n] || []).filter(r => own(r.formula) || own(r.formula2)).map(r => ({ kind: r.kind, range: r.range, formula: r.formula, formula2: r.formula2 }));
}

const missing = (primary, n, available) => fail(folded(primary, n).length ? 'xlsx-sheet-case-mismatch' : 'xlsx-sheet-not-found', { sheet: n, available });

function classify(primary, op) {
  if (!primary || primary.format !== 'xlsx') return fail('xlsx-observation-missing');
  if (!primary.complete) return fail(primary.truncated ? 'xlsx-observation-truncated' : 'xlsx-observation-unreadable-getters', { unsupported: primary.unsupported });
  const s = op.selector;
  // an ambiguous source (the same exact name twice) is never a target
  if (s.sheet !== undefined && exact(primary, s.sheet).length > 1) return fail('xlsx-sheet-ambiguous', { available: namesOf(primary) });
  const c = CLASSIFY[op.intent](primary, s, namesOf(primary));
  return c.state === 'APPLY' ? (drawingBlocked(primary, op) || c) : c;
}

/**
 * M7 S2b (owner GO 2026-10-01, owner channel): the L3 claims every sheet's drawings, so an unreadable one refuses the
 * write; copy_sheet of a source hosting an IMAGE or a SHAPE refuses (the image content and a shape's geometry on a
 * fresh sheet are not observable -- docs/evidence/M7-S2B-DRAWING-PROBE.md). A chart-only source keeps the proven path.
 */
function drawingBlocked(primary, op) {
  const o = drawings.observe(primary.drawingReads, namesOf(primary));
  if (!o.complete) return fail('xlsx-sheets-drawings-unobservable', { unobservable: o.gaps.slice(0, 20), count: o.gaps.length });
  if (op.intent !== 'copy_sheet') return null;
  const hosted = (o.bySheet[op.selector.sheet] || []).filter(d => d.kind === 'image' || d.kind === 'shape');
  return hosted.length ? fail('xlsx-copy-source-hosts-images-or-shapes', { sheet: op.selector.sheet, drawings: hosted.slice(0, 20).map(d => ({ kind: d.kind, name: d.name })), count: hosted.length }) : null;
}

function resolve(primary, selector, ctx = {}) {
  const c = classify(primary, ctx.operation);
  if (c.state === 'FAIL_CLOSED') return { ok: false, target: null, selector, ambiguity: c.evidence.reason, detail: c.evidence };
  return { ok: true, target: Object.freeze({ ...selector, sheetsBefore: Object.freeze(namesOf(primary)) }) };
}

/** One callCommand, public API only; identity checks inside the command so a wrong target writes nothing. */
function lower(op) {
  const s = op.selector;
  const J = JSON.stringify;
  if (op.intent === 'copy_sheet') {
    // M11.5c.4: ONE worksheet.copy/v1 Bridge WRITE through the driver's injected B03 client (gate + write + receipt parse)
    return {
      kind: 'callCommand',
      body: `\ntry {\n  return __eoB03Copy(${J(s.sheet)}, ${J(s.name)});\n} catch (err) { return { ok: false, stage: 'threw', error: String(err) }; }\n`,
      meta: Object.freeze({ family: 'sheets', action: op.intent, ...s, bridgeCopyWrite: true }),
    };
  }
  const bodies = {
    create_sheet: `
  const name = ${J(s.name)};
  if (Api.GetSheets().some(w => w.GetName().toLowerCase() === name.toLowerCase())) return { ok: false, stage: 'name-taken' };
  Api.AddSheet(name);
  const hit = Api.GetSheets().filter(w => w.GetName() === name).length;
  return hit === 1 ? { ok: true } : { ok: false, stage: 'created-name-mismatch', hit };`,
    rename_sheet: `
  const ws = Api.GetSheet(${J(s.sheet)});
  if (!ws || ws.GetName() !== ${J(s.sheet)}) return { ok: false, stage: 'sheet-identity' };
  ws.SetName(${J(s.name)});
  return { ok: true };`,
    delete_sheet: `
  const ws = Api.GetSheet(${J(s.sheet)});
  if (!ws || ws.GetName() !== ${J(s.sheet)}) return { ok: false, stage: 'sheet-identity' };
  ws.Delete();
  return { ok: true };`,
    move_sheet: `
  const ws = Api.GetSheet(${J(s.sheet)});
  const ref = Api.GetSheet(${J(s.reference)});
  if (!ws || ws.GetName() !== ${J(s.sheet)} || !ref || ref.GetName() !== ${J(s.reference)}) return { ok: false, stage: 'sheet-identity' };
  ${s.where === 'before' ? 'ws.Move(ref, undefined);' : 'ws.Move(undefined, ref);'}
  return { ok: true };`,
  };
  return {
    kind: 'callCommand',
    body: `\ntry {${bodies[op.intent]}\n} catch (err) { return { ok: false, stage: 'threw', error: String(err) }; }\n`,
    meta: Object.freeze({ family: 'sheets', action: op.intent, ...s }),
  };
}

/** The sheet list the operation must produce from `before` (names, in order). */
function expectedSheets(op, before) {
  const s = op.selector;
  if (op.intent === 'create_sheet') return null; // position is not part of the contract: checked by removal
  if (op.intent === 'rename_sheet') return before.map(n => (n === s.sheet ? s.name : n));
  if (op.intent === 'delete_sheet') return before.filter(n => n !== s.sheet);
  if (op.intent === 'copy_sheet') return [...before, s.name];
  const rest = before.filter(n => n !== s.sheet);
  const j = rest.indexOf(s.reference);
  rest.splice(s.where === 'before' ? j : j + 1, 0, s.sheet);
  return rest;
}

function verifyPrimitive(op, before, after) {
  const b = before && before.primary;
  const a = after && after.primary;
  if (!b || !a) return { ok: false, evidence: { reason: 'xlsx-observation-missing' } };
  const s = op.selector;
  const nb = namesOf(b);
  const na = namesOf(a);
  let ok;
  if (op.intent === 'create_sheet') ok = exact(a, s.name).length === 1 && sheetEmpty(a, s.name) && na.length === nb.length + 1;
  else ok = JSON.stringify(na) === JSON.stringify(expectedSheets(op, nb));
  const copy = op.intent === 'copy_sheet' && ok ? copyMatches(a, s.sheet, s.name) : null;
  if (copy) ok = copy.ok;
  return { ok, evidence: { intent: op.intent, before: nb, after: na, proof: 'fresh-live-sheet-list', ...(copy ? { copy } : {}) } };
}

/**
 * L3: the sheet list changed exactly as asked, and every cell of every surviving sheet keeps its
 * formula and (non-formula) value. A rename allows exactly one formula change: a reference to the old
 * name rewritten to the new one; a formula that referenced the old name and did NOT follow is a change.
 */
/** The sheet list after the operation is exactly what it asked for (create: only the new name added). */
function listAsExpected(op, nb, na) {
  if (op.intent === 'create_sheet') return JSON.stringify(na.filter(n => n !== op.selector.name)) === JSON.stringify(nb);
  return JSON.stringify(na) === JSON.stringify(expectedSheets(op, nb));
}

/**
 * Is formula f1 the kept form of f0? Byte-equal -- except after a rename, where a formula that
 * referenced the old name must now carry the new one (and nothing else changed).
 */
function formulaKept(op, f0, f1) {
  const s = op.selector;
  const refersToOld = op.intent === 'rename_sheet' && f0 !== '' && formula.sheetRefs(f0).names.some(x => x.toLowerCase() === s.sheet.toLowerCase());
  return refersToOld ? formula.canonicalRenamed(f0, s.sheet, s.name) === formula.canonical(f1) : f0 === f1;
}

/** The change, if any, from a cell read before to its counterpart read after. */
function cellChange(op, cell, now) {
  if (!now) return 'unreadable-after';
  if (!formulaKept(op, cell.f, now.f)) return 'formula';
  return cell.f === '' && JSON.stringify(now.v) !== JSON.stringify(cell.v) ? 'value' : null;
}

function verifyPreservation(op, before, after) {
  const b = before && before.primary;
  const a = after && after.primary;
  if (!b || !a || !b.complete || !a.complete) return { ok: false, evidence: { reason: 'xlsx-preservation-observation-incomplete' } };
  const s = op.selector;
  const nb = namesOf(b);
  const na = namesOf(a);
  if (!listAsExpected(op, nb, na)) return { ok: false, evidence: { reason: 'xlsx-sheet-list-changed-beyond-the-operation', before: nb, after: na } };
  const mapSheet = n => (op.intent === 'rename_sheet' && n === s.sheet ? s.name : n);
  const changed = [];
  const seen = new Set();
  for (const [k, cell] of Object.entries(b.cells)) {
    const [sheet, addr] = k.split('\u0000');
    if (op.intent === 'delete_sheet' && sheet === s.sheet) continue;
    const k2 = key(mapSheet(sheet), addr);
    seen.add(k2);
    const change = cellChange(op, cell, a.cells[k2]);
    if (change) changed.push({ cell: `${mapSheet(sheet)}!${addr}`, change, ...(change === 'formula' ? { before: cell.f, after: a.cells[k2].f } : {}) });
  }
  for (const [k, cell] of Object.entries(a.cells)) {
    if (op.intent === 'copy_sheet' && k.startsWith(`${s.name}\u0000`)) continue; // the copy itself: proven by L1
    if (!seen.has(k) && !isBlank(cell)) changed.push({ cell: k.replace('\u0000', '!'), change: 'appeared' });
  }
  const side = workbookSide(op, b, a);
  // the drawings of every sheet: the requested change only -- a rename moves the parent name on purpose, a delete takes
  // the deleted sheet's drawings on purpose, a copy's own drawings are L1's (copyMatches); everything else equal
  const lc = n => String(n).toLowerCase();
  const draw = drawings.drawingDiffs(drawings.observe(b.drawingReads, nb), drawings.observe(a.drawingReads, na), {
    mapSheet: mapSheet,
    dropSheet: n => op.intent === 'delete_sheet' && lc(n) === lc(s.sheet),
    skipAfter: n => op.intent === 'copy_sheet' && n === s.name,
  });
  const ok = changed.length === 0 && side.ok && draw.length === 0;
  return { ok, evidence: { sheets: na, compared: seen.size, changed: changed.slice(0, 20), changedCount: changed.length, ...(side.checked ? { workbook: side } : {}), drawingsKept: draw.length === 0, drawingDiffs: draw } };
}

/**
 * delete / copy read the dependents: then every OTHER sheet's rules and objects, and the defined names
 * not scoped to the deleted / copied sheet, must be unchanged (the copy's own are L1's job). An
 * after-read that is incomplete where the before-read was complete is a failure, not a skip.
 */
function workbookSide(op, b, a) {
  if (!b.dependencies) return { ok: true, checked: false };
  if (!a.dependencies || !a.rules) return { ok: false, checked: true, reason: 'xlsx-dependents-not-read-after' };
  const s = op.selector;
  if (op.intent === 'rename_sheet') return renameSide(b, a, s.sheet, s.name);
  const target = (op.intent === 'copy_sheet' ? s.name : s.sheet).toLowerCase();
  const own = n => n.toLowerCase() === target;
  const diffs = [...ruleSideDiffs(b, a, own), ...objectSideDiffs(b, a, own), ...nameSideDiffs(b, a, own)];
  return { ok: diffs.length === 0, checked: true, diffs: diffs.slice(0, 20) };
}

/**
 * rename reads the dependents (its guard): then EVERY sheet's rules and objects and every defined name
 * must be the before-state with each reference to the old name rewritten to the new one -- nothing else
 * changed, and a reference that did NOT follow is a change (as formulaKept for the cells).
 */
function renameSide(b, a, from, to) {
  const map = n => (n && n.toLowerCase() === from.toLowerCase() ? to : n);
  const diffs = [];
  if (b.rules.complete && !a.rules.complete) diffs.push({ axis: 'rules', change: 'unobservable-after' });
  if (b.rules.complete && a.rules.complete) {
    for (const [n, list] of Object.entries(b.rules.bySheet)) {
      const after = a.rules.bySheet[map(n)] || [];
      const same = list.length === after.length && list.every((r, i) => r.kind === after[i].kind && r.type === after[i].type && r.range === after[i].range
        && refKept(r.formula, after[i].formula, from, to) && refKept(r.formula2, after[i].formula2, from, to));
      if (!same) diffs.push({ axis: 'rules', sheet: map(n) });
    }
  }
  for (const [n, o] of Object.entries(b.objects)) {
    const t = a.objects[map(n)];
    const d = objectsComplete(o) && objectsComplete(t) ? objectDiffs(o, t, from, to) : (JSON.stringify(o) === JSON.stringify(t) ? [] : [{ change: 'objects-unobservable' }]);
    if (d.length) diffs.push({ axis: 'objects', sheet: map(n), diffs: d.slice(0, 5) });
  }
  if (Array.isArray(b.definedNames)) {
    const an = Array.isArray(a.definedNames) ? a.definedNames : [];
    const kept = b.definedNames.length === an.length && b.definedNames.every((x, i) => x.name === an[i].name && map(x.scopeSheet) === an[i].scopeSheet && refKept(x.ref, an[i].ref, from, to));
    if (!kept) diffs.push({ axis: 'defined-names', before: b.definedNames.slice(0, 10), after: an.slice(0, 10) });
  }
  return { ok: diffs.length === 0, checked: true, diffs: diffs.slice(0, 20) };
}

function ruleSideDiffs(b, a, own) {
  if (!b.rules.complete) return [];
  if (!a.rules.complete) return [{ axis: 'rules', change: 'unobservable-after' }];
  return Object.entries(b.rules.bySheet).filter(([n, list]) => !own(n) && JSON.stringify(list) !== JSON.stringify(a.rules.bySheet[n])).map(([n]) => ({ axis: 'rules', sheet: n }));
}

const objectSideDiffs = (b, a, own) => Object.entries(b.objects).filter(([n, o]) => !own(n) && JSON.stringify(o) !== JSON.stringify(a.objects[n])).map(([n]) => ({ axis: 'objects', sheet: n }));

function nameSideDiffs(b, a, own) {
  const names = p => (Array.isArray(p.definedNames) ? JSON.stringify(p.definedNames.filter(x => !(x.scopeSheet && own(x.scopeSheet)))) : null);
  return names(b) !== null && names(a) !== names(b) ? [{ axis: 'defined-names', before: names(b), after: names(a) }] : [];
}

function wholeTaskChecks(op, finalObservation) {
  const p = finalObservation && finalObservation.primary;
  const s = op.selector;
  if (!p) return [{ name: 'final observation present', ok: false, detail: null }];
  const na = namesOf(p);
  if (op.intent === 'create_sheet') return [{ name: `exactly one sheet named ${s.name}`, ok: exact(p, s.name).length === 1, detail: { sheets: na } }];
  if (op.intent === 'rename_sheet') return [{ name: `${s.sheet} renamed to ${s.name}`, ok: exact(p, s.sheet).length === 0 && exact(p, s.name).length === 1, detail: { sheets: na } }];
  if (op.intent === 'copy_sheet') { const m = copyMatches(p, s.sheet, s.name); return [{ name: `exactly one sheet ${s.name}, equal to ${s.sheet}`, ok: exact(p, s.name).length === 1 && m.ok, detail: { sheets: na, diffs: m.diffs } }]; }
  if (op.intent === 'delete_sheet') return [{ name: `no sheet named ${s.sheet}`, ok: folded(p, s.sheet).length === 0, detail: { sheets: na } }];
  const i = na.indexOf(s.sheet);
  const j = na.indexOf(s.reference);
  const ok = i >= 0 && j >= 0 && (s.where === 'before' ? i === j - 1 : i === j + 1);
  return [{ name: `${s.sheet} directly ${s.where} ${s.reference}`, ok, detail: { sheets: na } }];
}

/** No cell window: the used ranges the driver always reads are the whole proof. */
const windowsOf = () => [];
/** copy_sheet reads the source and the copy with format and objects. */
const liveExtras = op => {
  if (op.intent === 'copy_sheet') return { formatSheets: [op.selector.sheet, op.selector.name], objectSheets: [op.selector.sheet, op.selector.name], dependents: true, drawingSheets: ['*'] };
  if (op.intent === 'delete_sheet' || op.intent === 'rename_sheet') return { dependents: true, drawingSheets: ['*'] };
  return { drawingSheets: ['*'] };
};

module.exports = { INTENTS, B03_BRIDGE_WRITE_RUNTIMES, liveExtras, plan, classify, resolve, lower, verifyPrimitive, verifyPreservation, wholeTaskChecks, windowsOf, nameProblem, referencesTo };
