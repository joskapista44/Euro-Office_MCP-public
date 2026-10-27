'use strict';

/**
 * GH #76 MODEL 1 -- the pure, shared comparison the stored-content verification uses (owner decision
 * 2026-10-18, "one implementation, not two"): both the optional product-side `verifyStoredContent`
 * (write-tools.cjs) and the release gate (test/live/stored-content-gate.cjs) diff an EXPECTED
 * post-state against one or more READ authorities through these two functions. Neither caller keeps
 * its own copy any more. PURE: no I/O, never throws.
 *
 * Lives under adapters/xlsx, not core: the comparison is over cells/formulas/numFmt -- XLSX concepts
 * the core may never name (test/static/no-format-knowledge.test.cjs).
 */

const normFormula = f => String(f || '').replace(/^=/, '').replace(/\s+/g, '').toUpperCase();
const hex = v => (typeof v === 'number' ? v.toString(16).toUpperCase().padStart(6, '0') : String(v || '').toUpperCase());

/**
 * cellDiffs(ref, expected, actual, authority) -> [named difference, ...]
 *   expected  { value?, formula?, numFmt?, bold?, fill?, text? } -- only the fields present are checked
 *             { blank: true } is exclusive with every other key: a cleared cell has no stored
 *             <c> element at all (missing from the parsed bytes is the EXPECTED stored shape, not
 *             a defect) -- without this, a `clear_range`-style expectation would read as a false
 *             MISMATCH (`:stored:missing`) for exactly the operation it is meant to confirm
 *             (GH #141 final product-level READ<->WRITE acceptance, 2026-10-18).
 *   actual    one authority's view of the cell:
 *     'product' (a read_range result cell): { v, f, nf, bold, fill, type }
 *     'stored'  (the independently parsed file bytes): { t, value, formula, numFmt, bold, fill }
 *   authority is a LABEL in the returned reason string, not a dispatch key -- callers with only ONE
 *   authority (e.g. the product-side verify, which has no independent byte parser) pass 'product'.
 */
function cellDiffs(ref, expected, actual, authority) {
  if (expected.blank === true) {
    // 'stored': an omitted <c> element IS the blank shape -- `!actual` passes, never `:missing`.
    // 'product': the read is still expected to NAME every requested address, blank or not -- an
    // absent entry there is a real read gap, not evidence of blankness.
    const empty = x => x === null || x === '' || x === undefined;
    // a blank cell holds neither a value nor a formula (a formula with an empty result is NOT blank)
    if (authority === 'stored') return (!actual || (empty(actual.value) && empty(actual.formula))) ? [] : [`${ref}:${authority}:not-blank`];
    if (!actual) return [`${ref}:${authority}:missing`];
    return (empty(actual.v) && empty(actual.f)) ? [] : [`${ref}:${authority}:not-blank`];
  }
  if (!actual) return [`${ref}:${authority}:missing`];
  const out = [];
  const v = authority === 'stored' ? actual.value : actual.v;
  const f = authority === 'stored' ? actual.formula : actual.f;
  const numFmt = authority === 'stored' ? actual.numFmt : actual.nf;
  if ('value' in expected && String(v) !== String(expected.value)) out.push(`${ref}:${authority}:value`);
  if ('formula' in expected && normFormula(f) !== normFormula(expected.formula)) out.push(`${ref}:${authority}:formula`);
  if (!('formula' in expected) && normFormula(f) !== '') out.push(`${ref}:${authority}:unexpected-formula`);
  if ('bold' in expected && (actual.bold === true) !== expected.bold) out.push(`${ref}:${authority}:bold`);
  if ('numFmt' in expected && numFmt !== expected.numFmt) out.push(`${ref}:${authority}:numFmt`);
  if ('fill' in expected && !hex(actual.fill).endsWith(String(expected.fill).replace('#', '').toUpperCase())) out.push(`${ref}:${authority}:fill`);
  if (expected.text === true && (authority === 'stored' ? actual.t === null : actual.type !== 'text')) out.push(`${ref}:${authority}:not-text`);
  return out;
}

/** The product read results (an array of read_range-shaped { ok, result: { cells } } entries) -> { address: cell }. */
function productCells(results) {
  const out = {};
  for (const r of results || []) for (const c of ((r && r.result && r.result.cells) || [])) out[c.address] = c;
  return out;
}

/**
 * GH #141 final acceptance (2026-10-18, follow-on from the review above) -- three MORE non-cell
 * comparators, for the operations where BOTH authorities give a reviewed, code-grounded shape:
 *   defined names (set_defined_name)   read_defined_names (defined-names-read.cjs) vs
 *                                       file-verifier's readDefinedNames()
 *   print area (set_print_area)        read_print_area (print-area-read.cjs) vs the SAME stored
 *                                       defined-names list, filtered to the 'Print_Area' name
 *   print fit (set_print_fit)          read_page_setup's `.print` half (page-read.cjs) vs
 *                                       file-verifier's readPageSetup() (fitToWidth/fitToHeight/
 *                                       fitToPage only -- see the NOT-covered note below)
 *
 * NOT covered here, named rather than guessed (each is a SEPARATE, real blocker, not a bulk
 * estimate): `set_page_layout`'s own field (`orientation`) has NO stored-bytes reader in
 * `test/live/file-verifier.cjs` at all (its `readPageSetup()` reads `paperSize`/`scale`/
 * `firstPageNumber`/`fitToPage`/`fitToWidth`/`fitToHeight` only) -- stays `comparable: false` in
 * the harness. `freeze_panes` gained a stored-bytes reader in the GH #141 comparator-gaps batch 1
 * patch (2026-10-18) -- see `freezePanesDiffs` below. `set_validation`/`add_conditional_format`
 * gained a NARROW rule-content comparator (not the full general one -- see `validationRuleDiffs`/
 * `conditionalFormatRuleDiffs` below), restricted to the ONE enum pair each has real evidence for.
 */

/**
 * definedNamesByName(list, authority) -> { name: { refersTo, scope } }, normalizing the TWO
 * different scope shapes into one comparable form:
 *   'product' (defined-names-read.cjs's projectDefinedNamesRead): [{ name, refersTo, scope? }] --
 *     scope is already a plain sheet-name string, or absent for workbook scope.
 *   'stored' (file-verifier's readDefinedNames): [{ name, refersTo, scope: { kind, sheet?, index? } }] --
 *     scope.kind 'sheet' -> scope.sheet; 'workbook'/'sheet-unresolved' -> undefined (never guessed).
 */
function definedNamesByName(list, authority) {
  const out = {};
  for (const e of list || []) {
    const scope = authority === 'stored' ? (e.scope && e.scope.kind === 'sheet' ? e.scope.sheet : undefined) : e.scope;
    out[e.name] = { refersTo: e.refersTo, scope };
  }
  return out;
}

/** definedNameDiffs(name, expected, actual, authority) -> [named difference, ...]. `expected`:
 * { refersTo?, scope? } -- `refersTo` compared via `normFormula` (a stored '=' is never required;
 * a product-side one always is -- both sides are stripped before comparing, so neither convention
 * has to match the other's). */
function definedNameDiffs(name, expected, actual, authority) {
  if (!actual) return [`${name}:${authority}:missing`];
  const out = [];
  if ('refersTo' in expected && normFormula(actual.refersTo) !== normFormula(expected.refersTo)) out.push(`${name}:${authority}:refersTo`);
  if ('scope' in expected && actual.scope !== expected.scope) out.push(`${name}:${authority}:scope`);
  return out;
}

const PRINT_AREA_PART = /^('(?:[^']|'')+'|[^'!]+)!\$([A-Z]{1,3})\$(\d+)(?::\$([A-Z]{1,3})\$(\d+))?$/;

/**
 * storedPrintAreaOf(storedDefinedNames, sheet) -> the SAME shape `read_print_area` returns
 * ({ set: false } | { set: true, range } | { set: true, areas } | 'UNOBSERVABLE:<why>'), built
 * from the stored `Print_Area` defined name on `sheet` (mirrors print-area-read.cjs's own
 * parsing, using the shared a1.cjs range primitives -- not that file's private helper, to keep
 * this test-verification module decoupled from a product WRITE-side file's internals).
 */
function storedPrintAreaOf(storedDefinedNames, sheet) {
  // the engine stores the BUILT-IN name `_xlnm.Print_Area` (LIVE finalacc-d1 file), not bare `Print_Area`
  const hits = (storedDefinedNames || []).filter(d => d.name && String(d.name).toLowerCase().replace(/^_xlnm\./, '') === 'print_area' && d.scope && d.scope.kind === 'sheet' && d.scope.sheet === sheet);
  if (hits.length === 0) return { set: false };
  if (hits.length > 1) return 'UNOBSERVABLE:more-than-one-print-area';
  const ref = String(hits[0].refersTo || '').replace(/^=/, '');
  const parts = [];
  for (const part of ref.split(',')) {
    const m = PRINT_AREA_PART.exec(part.trim());
    if (!m) return 'UNOBSERVABLE:not-a-plain-rectangle';
    const owner = m[1].startsWith("'") ? m[1].slice(1, -1).replace(/''/g, "'") : m[1];
    if (owner !== sheet) return 'UNOBSERVABLE:part-on-another-sheet';
    parts.push(m[4] ? `${m[2]}${m[3]}:${m[4]}${m[5]}` : `${m[2]}${m[3]}`);
  }
  return parts.length === 1 ? { set: true, range: parts[0] } : { set: true, areas: parts };
}

/** printAreaDiffs(expected, actual, authority) -> [named difference, ...]. `expected`/`actual`
 * share the `read_print_area` result shape (section header above); a STRING `actual` is always an
 * `UNOBSERVABLE:*`/error reason from that authority, never compared as a value. */
function printAreaDiffs(expected, actual, authority) {
  if (typeof actual === 'string') return [`printArea:${authority}:${actual}`];
  if (!actual || typeof actual.set !== 'boolean') return [`printArea:${authority}:missing`];
  if (actual.set !== expected.set) return [`printArea:${authority}:set`];
  if (expected.set) {
    if ('range' in expected && actual.range !== expected.range) return [`printArea:${authority}:range`];
    if ('areas' in expected && JSON.stringify(actual.areas || null) !== JSON.stringify(expected.areas)) return [`printArea:${authority}:areas`];
  }
  return [];
}

/** printFitDiffs(expected, actual, authority) -> [named difference, ...]. `expected`:
 * { fitToWidth?, fitToHeight?, fitToPage? } (the ONLY three fields verified -- see the
 * NOT-covered note above for why `scale`/`paper`/the rest of `read_page_setup` are excluded).
 * `actual` is the product read's `.print` half, or the stored `readPageSetup()` value -- both
 * carry these three field names directly, no translation needed. */
function printFitDiffs(expected, actual, authority) {
  if (!actual || typeof actual !== 'object') return [`printFit:${authority}:missing`];
  const out = [];
  for (const k of ['fitToWidth', 'fitToHeight', 'fitToPage']) {
    if (k in expected && actual[k] !== expected[k]) out.push(`printFit:${authority}:${k}`);
  }
  return out;
}

/**
 * GH #141 minimal-tasks batch 1 (2026-10-18) -- three MORE non-cell comparators, for
 * `unmerge_range` and the four sheet-structural intents (`rename_sheet`/`delete_sheet`/
 * `move_sheet`/`copy_sheet`), plus a narrow ABSENCE-only comparator for `clear_validation`
 * (never the general validation-content comparator, which stays blocked -- see the batch's own
 * writeup for why: an OOXML<->engine enum translation this repo does not have).
 */

/** mergeDiffs(expected, actual, authority) -> [named difference, ...].
 *   expected  { ref: '<cell>', range: '<A1 merge range>', merged: boolean } -- `ref` is the
 *             product-side anchor cell to inspect; `range` is the FULL merge area to look for in
 *             the stored `<mergeCells>` list (both always given, regardless of `merged`'s value,
 *             so BOTH authorities can be checked for either outcome without a second shape).
 *   actual    'product': the read_range result cell at `ref` (or undefined if the read never
 *             named it) -- `.merge` is a range-text string when merged, `null` otherwise
 *             (src/drivers/eurooffice/xlsx-driver.cjs: `out.merge = m ? addrOf(m) : null`).
 *             An UNMERGED cell reports its OWN address (e.g. 'A1'), which counts as not merged.
 *             'stored': the array of merge-range strings from `storedMergeRangesOf()` below
 *             (NEVER the raw `readMergeCells()` result -- that one reports "mergeCells-not-present"
 *             as `{ok:false}`, which this function's caller has already normalized to `[]`, the
 *             correct reading of "no merges on this sheet", not a read failure).
 */
function mergeDiffs(expected, actual, authority) {
  if (authority === 'stored') {
    const ranges = Array.isArray(actual) ? actual : [];
    const present = ranges.includes(expected.range);
    if (expected.merged) return present ? [] : ['merge:stored:not-merged'];
    return present ? ['merge:stored:still-merged'] : [];
  }
  if (!actual) return ['merge:product:missing'];
  // an UNMERGED cell's own MergeArea is its own address, never null (merge.cjs: "an unmerged cell's
  // own reported area IS its own address") -- LIVE finalacc-p3 read 'still-merged' on a real unmerge
  const mergedRange = typeof actual.merge === 'string' ? actual.merge.replace(/\$/g, '') : null;
  const ownCell = mergedRange === null || mergedRange === expected.ref;
  if (expected.merged) return !ownCell && mergedRange === expected.range ? [] : [`merge:product:${ownCell ? 'not-merged' : 'wrong-range'}`];
  return ownCell ? [] : ['merge:product:still-merged'];
}

/** storedMergeRangesOf(readMergeCellsResult) -> string[] -- normalizes file-verifier's
 * `readMergeCells()` three outcomes (present / absent / genuinely unreadable) into a plain
 * range-string list for `mergeDiffs`'s 'stored' authority: "no `<mergeCells>` element at all" is
 * the SAME as "an empty one" (both mean "nothing merged here"), never a read failure. Any OTHER
 * failure reason (sheet not found, ambiguous element, malformed ref) stays a failure, named. */
function storedMergeRangesOf(readMergeCellsResult) {
  if (readMergeCellsResult && readMergeCellsResult.ok === false && readMergeCellsResult.reason === 'mergeCells-not-present') return { ok: true, ranges: [] };
  if (readMergeCellsResult && readMergeCellsResult.ok === true) return { ok: true, ranges: readMergeCellsResult.value.ranges };
  return { ok: false, reason: (readMergeCellsResult && readMergeCellsResult.reason) || 'unreadable' };
}

/** sheetListDiffs(expected, names, authority) -> [named difference, ...]. `names` is the ordered
 * sheet-name list (product: `workbook_overview`'s `sheets[].name`, already in document order;
 * stored: the raw `xl/workbook.xml` `<sheets>` name list, same order). `expected`:
 *   contains:   string[] -- every name here MUST appear (existence, not full-list equality --
 *               a shared workbook runs MANY rounds at once; comparing the WHOLE list would make
 *               one round's assertion depend on every OTHER round's concurrent sheet mutations)
 *   excludes:   string[] -- none of these may appear
 *   adjacentTo: { name, ref, side: 'before'|'after' } -- `name` must sit immediately before/after
 *               `ref` in the list (a RELATIVE, not absolute, position check -- robust to unrelated
 *               sheets existing anywhere else in the list)
 */
function sheetListDiffs(expected, names, authority) {
  if (!Array.isArray(names)) return [`sheetList:${authority}:unreadable`];
  const out = [];
  for (const n of expected.contains || []) if (!names.includes(n)) out.push(`sheetList:${authority}:missing:${n}`);
  for (const n of expected.excludes || []) if (names.includes(n)) out.push(`sheetList:${authority}:present:${n}`);
  if (expected.adjacentTo) {
    const { name, ref, side } = expected.adjacentTo;
    const i = names.indexOf(name);
    const j = names.indexOf(ref);
    const ok = i >= 0 && j >= 0 && (side === 'before' ? i === j - 1 : i === j + 1);
    if (!ok) out.push(`sheetList:${authority}:adjacency:${name}`);
  }
  return out;
}

/** validationAbsentDiffs(expected, actual, authority) -> [named difference, ...]. PROVES ABSENCE
 * ONLY -- deliberately narrower than `validationRuleDiffs` below (which covers exactly one
 * type+operator pair; this one needs none of that vocabulary, because "absent" has only one
 * shape on either side). `expected`: { cell: '<A1 ref>' } (the `merged`/`range` style flag
 * doesn't apply here -- this comparator has exactly one outcome, "no validation rule covers this
 * cell any more").
 *   actual  'product': the read_validations result's `.validation` object -- `{ absent: true }`
 *           is the driver's OWN "nothing here" marker (xlsx-driver.cjs: `if (!v) return
 *           {absent:true};`), never decoded further.
 *           'stored': `{ legacy: [...], x14: [...] }` from `readWorksheetDataValidations()` --
 *           each entry's OWN `sqref` (OOXML: SPACE-separated range tokens, never comma) is
 *           checked for an EXACT token match against `expected.cell` only -- a single-cell/
 *           single-range string equality, not general range-containment (this round's own `cell`
 *           is always written as the SAME exact token the rule was created with, so this is
 *           sufficient for what this comparator proves; it is NOT a general sqref-membership test).
 */
function validationAbsentDiffs(expected, actual, authority) {
  if (authority === 'product') return (actual && actual.absent === true) ? [] : ['validation:product:not-absent'];
  if (!actual || !Array.isArray(actual.legacy) || !Array.isArray(actual.x14)) return ['validation:stored:unreadable'];
  const covers = rule => typeof rule.sqref === 'string' && rule.sqref.split(/\s+/).includes(expected.cell);
  return (actual.legacy.some(covers) || actual.x14.some(covers)) ? ['validation:stored:still-present'] : [];
}

/**
 * validationRuleDiffs(expected, actual, authority) -> [named difference, ...]. The GENERAL
 * rule-content comparator narrowed to the ONE type+operator pair this unit found real evidence
 * for, both halves independently -- GH #141 krit. 6: `wholeNumber`/`greater`.
 *   ENGINE side (the 'product' vocabulary -- `read_validations`'s own output, no translation:
 *     xlsx-driver.cjs's validation reader is `{type: v.GetType(), operator: v.GetOperator(), ...}`
 *     verbatim): `wholeNumber` -> `xlValidateWholeNumber`, `greater` -> `xlGreater` -- this is
 *     validation.cjs's OWN write-side table (`TYPES.wholeNumber`, `ONE.greater`), and a direct
 *     LIVE read-back of exactly this rule already exists (`docs/evidence/m7-final-m7-d-20261004.json`
 *     and two sibling M7 evidence files: `read: {type:"xlValidateWholeNumber",
 *     alertStyle:"xlValidAlertStop", operator:"xlGreater", formula1:"4", formula2:""}`).
 *   OOXML side (the 'stored' vocabulary, ECMA-376 Part 1 §18.3.1.23/§18.18.21): `wholeNumber` ->
 *     `'whole'` and `greater` -> `'greaterThan'` are each independently LIVE-evidenced in
 *     `docs/evidence/b11-v3-validation-live-acceptance-20261007.json`'s own FILE/x14 reading --
 *     `type:'whole'` for a different rule's `operator:'between'` (same file, x14[2]), and
 *     `operator:'greaterThan'` for a different rule's `type:'decimal'` (same file, x14[3]). No
 *     sample pairs the two in the SAME rule -- that combination is not itself a LIVE sample, but
 *     it is not a guess either: OOXML's `type`/`operator` attributes are orthogonal per the spec
 *     (an operator's spelling does not depend on which validation type it qualifies), so each
 *     half's own independent proof composes. Named here, not hidden: if this orthogonality
 *     assumption is ever found wrong, this is the one comparator it would break.
 *   expected  { range, value } (only the round's own field -- `type`/`operator` are fixed by
 *             this narrow comparator, not general parameters).
 *   actual  'product': `read_validations`'s `.validation` object, engine vocabulary directly.
 *           'stored': `{ legacy: [...], x14: [...] }` from `readWorksheetDataValidations()`,
 *           matched by an EXACT `sqref === expected.range` token (same convention as
 *           `validationAbsentDiffs`'s `cell` match, one level up to a general range string).
 */
function validationRuleDiffs(expected, actual, authority) {
  if (authority === 'product') {
    if (!actual || typeof actual !== 'object' || typeof actual.type !== 'string') return ['validationRule:product:unreadable'];
    if (actual.type !== 'xlValidateWholeNumber') return ['validationRule:product:wrong-type'];
    if (actual.operator !== 'xlGreater') return ['validationRule:product:wrong-operator'];
    return String(actual.formula1) === String(expected.value) ? [] : ['validationRule:product:wrong-value'];
  }
  if (!actual || !Array.isArray(actual.legacy) || !Array.isArray(actual.x14)) return ['validationRule:stored:unreadable'];
  const matches = r => r.type === 'whole' && r.operator === 'greaterThan' && typeof r.sqref === 'string' && r.sqref.split(/\s+/).includes(expected.range) && String(r.formula1) === String(expected.value);
  return (actual.legacy.some(matches) || actual.x14.some(matches)) ? [] : ['validationRule:stored:not-present'];
}

/**
 * GH #141 minimal-tasks batch 2 (2026-10-18) -- ten MORE pure comparators, for
 * `delete_defined_name` (an ABSENCE-only companion to
 * `definedNameDiffs`, same shape as batch 1's `validationAbsentDiffs`), the two print-titles
 * intents (`set_print_titles`/`clear_print_titles`, names.cjs's SAME `Print_Titles`/`Print_Area`
 * mechanism -- `storedPrintTitlesOf`/`productPrintTitlesOf` mirror `storedPrintAreaOf` exactly,
 * one axis wider: `{columns?, rows?}` instead of a rectangle), and six print.cjs settings
 * (`set_print_scale`/`set_paper_size`/`set_first_page_number`/`set_print_centering`/
 * `set_header_footer`/`set_header_footer_flag`/`set_page_break` -- seven names, one shared
 * `read_page_setup`/`readPageSetup`/`readPrintOptions`/`readRowBreaks`/`readColBreaks`/
 * `readHeaderFooter` source family page-read.cjs/file-verifier.cjs already read for `printFit`).
 */

/** definedNameAbsentDiffs(name, list, authority) -> [named difference, ...]. The companion to
 * `definedNameDiffs` for `delete_defined_name`: proves ABSENCE, not content -- `list` is the SAME
 * raw entry list both `definedNamesByName()` and the `definedName` NON_CELL_KINDS storedOf/product
 * unwrap already produce (`{name, ...}[]`), never the by-name map (a map built by
 * `definedNamesByName` would already have discarded WHICH entries exist, and this only needs a
 * presence check). A non-array list (the authority's read failed or was never attempted) is
 * always named `:unreadable`, never vacuously "absent". */
function definedNameAbsentDiffs(name, list, authority) {
  if (!Array.isArray(list)) return [`${name}:${authority}:unreadable`];
  const present = list.some(e => e && typeof e.name === 'string' && e.name.toLowerCase() === String(name).toLowerCase());
  return present ? [`${name}:${authority}:still-present`] : [];
}

const TITLES_COLS = /^('(?:[^']|'')+'|[^'!,]+)!\$([A-Z]{1,3}):\$([A-Z]{1,3})$/;
const TITLES_ROWS = /^('(?:[^']|'')+'|[^'!,]+)!\$(\d+):\$(\d+)$/;
const unquoteSheet = s => (s.startsWith("'") ? s.slice(1, -1).replace(/''/g, "'") : s);

/**
 * storedPrintTitlesOf(storedDefinedNames, sheet) -> the SAME shape `read_defined_names`'s own
 * `titles` projection uses (`{ set: false }` | `{ set: true, columns?, rows? }` |
 * `'UNOBSERVABLE:<why>'`), built from the stored `Print_Titles` defined name on `sheet` -- mirrors
 * `storedPrintAreaOf` exactly (same absence/ambiguity/sheet-ownership rules), one axis wider: a
 * part is either a column range (`Sheet!$A:$B`) or a row range (`Sheet!$1:$2`), never a rectangle.
 * Decoupled from `defined-names-read.cjs`'s own `titlesOf` on purpose (same reasoning as
 * `storedPrintAreaOf`'s own doc: this module stays independent of that product-side file's
 * internals), using local regexes rather than `a1.cjs`'s `colNum` (this file has no requires).
 */
function storedPrintTitlesOf(storedDefinedNames, sheet) {
  const hits = (storedDefinedNames || []).filter(d => d.name && String(d.name).toLowerCase().replace(/^_xlnm\./, '') === 'print_titles' && d.scope && d.scope.kind === 'sheet' && d.scope.sheet === sheet);
  if (hits.length === 0) return { set: false };
  if (hits.length > 1) return 'UNOBSERVABLE:more-than-one-print-titles';
  const ref = String(hits[0].refersTo || '').replace(/^=/, '');
  const out = {};
  for (const part of ref.split(',')) {
    const trimmed = part.trim();
    const c = TITLES_COLS.exec(trimmed);
    const m = c || TITLES_ROWS.exec(trimmed);
    if (!m) return 'UNOBSERVABLE:not-a-titles-shape';
    if (unquoteSheet(m[1]) !== sheet) return 'UNOBSERVABLE:part-on-another-sheet';
    const key = c ? 'columns' : 'rows';
    if (out[key]) return 'UNOBSERVABLE:not-a-titles-shape';
    out[key] = `${m[2]}:${m[3]}`;
  }
  return { set: true, ...(out.columns ? { columns: out.columns } : {}), ...(out.rows ? { rows: out.rows } : {}) };
}

/** productPrintTitlesOf(names, sheet) -> the same `{set:false}` / `{set:true, columns?, rows?}` /
 * `'UNOBSERVABLE:<why>'` shape, from `read_defined_names`'s own entry list (`defined-names-read.cjs`'s
 * `builtIn: 'Print_Titles'` + `titles` marker on the sheet-scoped entry -- never guessed from a
 * bare name match, since an UNSCOPED `Print_Titles` is a different, reserved-but-not-this entry). */
function productPrintTitlesOf(names, sheet) {
  const hit = (names || []).find(e => e && e.builtIn === 'Print_Titles' && e.scope === sheet);
  if (!hit) return { set: false };
  if (typeof hit.titles === 'string') return hit.titles;
  return { set: true, ...hit.titles };
}

/** printTitlesDiffs(expected, actual, authority) -> [named difference, ...]. `expected`/`actual`
 * share the above result shape; `expected.sheet` (present for lookup only, in the caller's
 * `expect` object) is never itself compared. A STRING `actual` is always an `UNOBSERVABLE:*`
 * reason from that authority, never compared as a value (mirrors `printAreaDiffs`). */
function printTitlesDiffs(expected, actual, authority) {
  if (typeof actual === 'string') return [`printTitles:${authority}:${actual}`];
  if (!actual || typeof actual.set !== 'boolean') return [`printTitles:${authority}:missing`];
  if (actual.set !== expected.set) return [`printTitles:${authority}:set`];
  if (expected.set) {
    if ('columns' in expected && actual.columns !== expected.columns) return [`printTitles:${authority}:columns`];
    if ('rows' in expected && actual.rows !== expected.rows) return [`printTitles:${authority}:rows`];
  }
  return [];
}

/** printScaleDiffs(expected, actual, authority) -> [named difference, ...]. `expected`:
 * { scale, fitToPage: false } -- MEASURED together (print.cjs `PAGE.set_print_scale.holds`: a
 * scale write always turns `fitToPage` off; `test/live/xlsx-l10b-preseeded-fixture.cjs`'s `PR_SCALE`
 * case, an already-passing acceptance fixture, confirms the read side: `{scale:80, fitToPage:false,
 * paper:'A4', ...}`). `actual` is the product read's `.print` half or the stored `readPageSetup()`
 * value -- both carry `scale`/`fitToPage` directly, no translation (same as `printFitDiffs`). */
function printScaleDiffs(expected, actual, authority) {
  if (!actual || typeof actual !== 'object') return [`printScale:${authority}:missing`];
  const out = [];
  if ('scale' in expected && actual.scale !== expected.scale) out.push(`printScale:${authority}:scale`);
  if ('fitToPage' in expected && actual.fitToPage !== expected.fitToPage) out.push(`printScale:${authority}:fitToPage`);
  return out;
}

/** The OOXML `pageSetup/@paperSize` (ST_PaperSize) id for a paper NAME print.cjs's own `PAPER`
 * table can write -- MEASURED, not the general spec table: `test/live/xlsx-l10b-preseeded-fixture.cjs`'s
 * `PAPER_ID` (an already-passing acceptance fixture) confirms exactly these three: A3=8, A4=9,
 * 'US Letter'=1. Every OTHER name in print.cjs's `PAPER` table has no reviewed numeric id anywhere
 * in this repo and is deliberately NOT guessed here (`paperSizeDiffs` names the gap instead). */
const PAPER_SIZE_ID = Object.freeze({ A3: 8, A4: 9, 'US Letter': 1 });

/** paperSizeDiffs(expected, actual, authority) -> [named difference, ...]. `expected`: { paper }
 * (one of print.cjs's `PAPER` names). `actual` is the product read's `.print` half (`.paper`, the
 * SAME name, via `page-read.cjs`'s own `paperOf()` width/height match) or the stored
 * `readPageSetup()` value (`.paperSize`, the OOXML numeric id -- translated through
 * `PAPER_SIZE_ID` above, never compared as a raw name). */
function paperSizeDiffs(expected, actual, authority) {
  if (!actual || typeof actual !== 'object') return [`paperSize:${authority}:missing`];
  if (authority === 'stored') {
    const want = PAPER_SIZE_ID[expected.paper];
    if (want === undefined) return [`paperSize:stored:no-measured-id-for:${expected.paper}`];
    return actual.paperSize === want ? [] : ['paperSize:stored:paperSize'];
  }
  return actual.paper === expected.paper ? [] : ['paperSize:product:paper'];
}

/** firstPageNumberDiffs(expected, actual, authority) -> [named difference, ...]. `expected`:
 * { value } -- a positive integer, or null (clear). `actual.firstPageNumber` on EITHER authority
 * already folds the `useFirstPage`/`useFirstPageNumber` flag in (page-read.cjs's `printHalf`:
 * `r.useFirstPage===true && Number.isInteger(r.firstPage) ? r.firstPage : null`; `file-verifier`'s
 * `readPageSetup` keeps the flag as its OWN field) -- SET compares the folded value on both sides;
 * CLEAR (print.cjs F2: "the number goes, the useFirstPage flag stays as it was") checks only that
 * `firstPageNumber` reads null, never the flag itself (which this comparator does not claim to
 * predict after a clear). */
function firstPageNumberDiffs(expected, actual, authority) {
  if (!actual || typeof actual !== 'object') return [`firstPageNumber:${authority}:missing`];
  if (expected.value === null) return actual.firstPageNumber === null ? [] : [`firstPageNumber:${authority}:firstPageNumber`];
  if (authority === 'product') return actual.firstPageNumber === expected.value ? [] : ['firstPageNumber:product:firstPageNumber'];
  return (actual.useFirstPageNumber === true && actual.firstPageNumber === expected.value) ? [] : ['firstPageNumber:stored:firstPageNumber'];
}

/** printCenteringDiffs(expected, actual, authority) -> [named difference, ...]. `expected`:
 * { horizontal? } and/or { vertical? } (print.cjs: one axis per step, but both MAY be asserted
 * independently here). `actual` 'product': the read's `.print.horizontal`/`.vertical` (booleans,
 * page-read.cjs's own `hCenter`/`vCenter` rename). 'stored': `readPrintOptions()`'s
 * `horizontalCentered`/`verticalCentered` -- a DIFFERENT field name on this authority only, so
 * this comparator (unlike `printFitDiffs`/`printScaleDiffs`) translates per authority. */
function printCenteringDiffs(expected, actual, authority) {
  if (!actual || typeof actual !== 'object') return [`printCentering:${authority}:missing`];
  const out = [];
  if ('horizontal' in expected) {
    const v = authority === 'stored' ? actual.horizontalCentered : actual.horizontal;
    if (v !== expected.horizontal) out.push(`printCentering:${authority}:horizontal`);
  }
  if ('vertical' in expected) {
    const v = authority === 'stored' ? actual.verticalCentered : actual.vertical;
    if (v !== expected.vertical) out.push(`printCentering:${authority}:vertical`);
  }
  return out;
}

/** headerFooterDiffs(expected, actual, authority) -> [named difference, ...]. `expected`:
 * { slot, value } -- `slot` one of the six HF_SLOTS names, `value` a code string or null (clear).
 * `actual` is the product read's `.print` half or the stored `readHeaderFooter()` value -- both
 * key the six slots by the SAME field names (page-read.cjs's `HF_SLOTS` / file-verifier's
 * `HF_SLOTS`, same set), no translation. */
function headerFooterDiffs(expected, actual, authority) {
  if (!actual || typeof actual !== 'object') return [`headerFooter:${authority}:missing`];
  return actual[expected.slot] === expected.value ? [] : [`headerFooter:${authority}:${expected.slot}`];
}

/** headerFooterFlagDiffs(expected, actual, authority) -> [named difference, ...]. `expected`:
 * { flag, value } -- `flag` 'differentFirst' | 'differentOddEven'. Same field names on both
 * authorities (page-read.cjs's `differentFirst`/`differentOddEven` rename of `diffFirst`/
 * `diffOddEven`; file-verifier's `readHeaderFooter()` already names them `differentFirst`/
 * `differentOddEven` directly, ECMA-376's own attribute names). */
function headerFooterFlagDiffs(expected, actual, authority) {
  if (!actual || typeof actual !== 'object') return [`headerFooterFlag:${authority}:missing`];
  return actual[expected.flag] === expected.value ? [] : [`headerFooterFlag:${authority}:${expected.flag}`];
}

const colNumLocal = s => [...String(s).toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0);

/** pageBreakDiffs(expected, actual, authority) -> [named difference, ...]. `expected`:
 * { axis: 'row'|'column', at, remove? } -- `at` a row number (1-based) or column letter, the SAME
 * `set_page_break` step shape (print.cjs: a manual break is stored at id `at-1`). `actual`
 * 'product': the read's `.print.rowBreaks`/`.columnBreaks` (ALREADY the `at`-shaped list,
 * page-read.cjs's own `manualAts()` conversion -- `at` compared directly, no further translation).
 * 'stored': `{ breaks: [{id,man}] }`, the caller's own per-axis pick of `readRowBreaks`/
 * `readColBreaks` (passed pre-selected, since this comparator has no sheet/file access of its
 * own) -- `id` translated from `at` locally (`colNumLocal`, this file has no requires so `a1.cjs`'s
 * `colNum` is not reachable), and only a MANUAL break (`man===true`) counts, never an automatic
 * one at the same id (mirrors `stepGap`'s own refusal in print.cjs for a removal). Absence of the
 * `<rowBreaks>`/`<colBreaks>` element at all is passed in as an EMPTY array by the caller (the
 * same "absence means nothing here" rule `storedMergeRangesOf` already established for merges),
 * never surfaced as its own failure here. */
function pageBreakDiffs(expected, actual, authority) {
  const field = expected.axis === 'column' ? 'columnBreaks' : 'rowBreaks';
  if (authority === 'product') {
    if (!actual || !Array.isArray(actual[field])) return ['pageBreak:product:missing'];
    const present = actual[field].includes(expected.at);
    return present === !expected.remove ? [] : [`pageBreak:product:${expected.remove ? 'still-present' : 'not-present'}`];
  }
  if (!actual || !Array.isArray(actual.breaks)) return ['pageBreak:stored:missing'];
  const id = (expected.axis === 'column' ? colNumLocal(expected.at) : Number(expected.at)) - 1;
  const present = actual.breaks.some(b => b.id === id && b.man === true);
  return present === !expected.remove ? [] : [`pageBreak:stored:${expected.remove ? 'still-present' : 'not-present'}`];
}

/**
 * GH #141 minimal-tasks batch 3 (2026-10-18) -- the final 12 intents: the chart family's
 * title/axis-title/legend (comparable, grounded in file-verifier's existing chart readers plus
 * one new one, `readChartTitleText`), rename/delete_chart (comparable, reusing `sheetListDiffs`
 * unchanged -- a chart-name list is not a sheet-name list, but the existence/exclusion check is
 * identical, and `sheetListDiffs` never assumed its `names` argument was sheets specifically),
 * `insert_image` (comparable, a drawing-presence-by-name check, mirroring the chart-presence
 * pattern one layer up: a sheet's DRAWING part instead of a chart's CONTENT part), and
 * `delete_conditional_format` (comparable, an ABSENCE-only companion to the family's own
 * `add_conditional_format`, same shape as `validationAbsentDiffs`/`definedNameAbsentDiffs`).
 * `set_chart_style`/`move_chart`/`copy_chart`/`unfreeze_panes`/`clear_filter` stay
 * `comparable: false`, named -- see the harness's own `readOf` cases for why each one specifically
 * (style is mostly WRITE-ONLY by the family's own documentation; move/copy need chart POSITION,
 * which no file-verifier reader carries at all; freeze/filter have no stored-bytes reader in this
 * repo at all, same boundary `freeze_panes`/`filter_range` already stayed behind).
 */

const trimChartTitle = x => String(x === null || x === undefined ? '' : x).replace(/[\r\n]+$/, '');

/** chartTitleDiffs(expected, actual, authority) -> [named difference, ...]. `expected`:
 * { title } -- `actual` is the RAW title value for that authority (product: the chart object's
 * own `.title` field, a string MEASURED to carry a trailing CRLF the write side strips before
 * comparing (charts.cjs's own `trimTitle`) -- `undefined` if the authority never found the chart
 * at all, which this comparator keeps distinct from `null`/`''` (a chart WITH no title, a real,
 * comparable state, never conflated with "the chart itself is missing"). */
function chartTitleDiffs(expected, actual, authority) {
  if (actual === undefined) return [`chartTitle:${authority}:missing`];
  const text = actual === null ? '' : actual;
  if (typeof text !== 'string') return [`chartTitle:${authority}:missing`];
  return trimChartTitle(text) === expected.title ? [] : [`chartTitle:${authority}:title`];
}

/** chartAxisTitleDiffs(expected, actual, authority) -> [named difference, ...]. `expected`:
 * { axis: 'horizontal'|'vertical', title } -- `actual` is the RAW axis-title value the CALLER has
 * already selected for `expected.axis` (product: `chart.horAxisTitle`/`.verAxisTitle`; stored:
 * `readChartAxisTitles()`'s own `.horizontal`/`.vertical`) -- this comparator never picks the axis
 * itself, only compares the one value it is handed. `null` (no `c:title` element / no axis title
 * set) is a real, comparable "no title" state, kept distinct from `undefined` (chart/axis not
 * found at all). Named per axis, same convention as `printCenteringDiffs`'s per-axis naming.
 *
 * LIVE finalacc batch3 p9 (2026-10-18, a maintainer's own round-trip on the kept copy): the
 * PRODUCT side's `horAxisTitle` carries a trailing CRLF (`"GH141 Axis Title\r\n"`) -- the SAME
 * trailing-CRLF shape `chartTitleDiffs`'s own doc already measured for the MAIN title, here on
 * the AXIS title instead. Reuses `trimChartTitle` unchanged, never a second normalizer. */
function chartAxisTitleDiffs(expected, actual, authority) {
  if (actual === undefined) return [`chartAxisTitle:${authority}:missing`];
  const text = actual === null ? '' : actual;
  if (typeof text !== 'string') return [`chartAxisTitle:${authority}:missing`];
  return trimChartTitle(text) === expected.title ? [] : [`chartAxisTitle:${authority}:${expected.axis}`];
}

/** The OOXML `c:legendPos/@val` for a `set_chart_legend` position NAME -- MEASURED: `charts.cjs`'s
 * own `LEGEND_POS` enum (`bottom`/`top`/`left`/`right`/`none`) is the product's write/read
 * vocabulary (`GetLegendPos()` returns the SAME string the setter took, charts.cjs's
 * `classifyLegend`: `hits[0].legend === op.legend`); the FILE stores the ECMA-376 abbreviation
 * instead (`b`/`t`/`l`/`r`; 'none' has no `c:legendPos` at all -- no element, see
 * `readChartLegend`'s own doc). Every OOXML `ST_LegendPos` value NOT in `LEGEND_POS` (e.g. `tr`)
 * is unreachable from this product's own write side, so it is not in this table at all --
 * `chartLegendDiffs` below never guesses one. */
const CHART_LEGEND_POS_OOXML = Object.freeze({ bottom: 'b', top: 't', left: 'l', right: 'r' });

/** chartLegendDiffs(expected, actual, authority) -> [named difference, ...]. `expected`:
 * { legend } (one of `LEGEND_POS`). `actual` 'product': the chart object's own `.legend` string,
 * compared directly (SAME vocabulary, no translation) -- `undefined` if the chart was not found.
 * 'stored': `readChartLegend()`'s own `{present, pos}` result (or `undefined`/non-object if the
 * chart part could not be resolved/read at all) -- translated through `CHART_LEGEND_POS_OOXML`
 * above; 'none' compares against `present === false` (no `c:legend` element), never a `pos`
 * value. */
function chartLegendDiffs(expected, actual, authority) {
  if (authority === 'product') {
    if (actual === undefined) return ['chartLegend:product:missing'];
    return actual === expected.legend ? [] : ['chartLegend:product:legend'];
  }
  if (!actual || typeof actual.present !== 'boolean') return ['chartLegend:stored:missing'];
  if (expected.legend === 'none') return actual.present === false ? [] : ['chartLegend:stored:legend'];
  const want = CHART_LEGEND_POS_OOXML[expected.legend];
  if (want === undefined) return [`chartLegend:stored:no-measured-pos-for:${expected.legend}`];
  return (actual.present === true && actual.pos === want) ? [] : ['chartLegend:stored:legend'];
}

/** drawingPresentDiffs(expected, actual, authority) -> [named difference, ...]. `expected`:
 * { name } -- the drawing's `cNvPr/@name` (`insert_image`'s own `name` field). `actual`
 * 'product': the sheet's `read_drawings` array (or `undefined`/non-array if the whole axis is
 * unreadable) -- presence is `.some(d => d.name === expected.name)`, kind-blind (this comparator
 * proves a drawing of that NAME exists, not specifically that it is an image -- see the harness's
 * own `drawingPresent` kind doc for why that narrower claim is not attempted). 'stored': the
 * `readDrawingAnchor()` result for that name (`{ok:true,...}` present, `{ok:false,
 * reason:'drawing-not-found'}` absent, any OTHER reason an unreadable gap, never conflated with
 * absence). */
function drawingPresentDiffs(expected, actual, authority) {
  if (authority === 'product') {
    if (!Array.isArray(actual)) return ['drawingPresent:product:unreadable'];
    return actual.some(d => d && d.name === expected.name) ? [] : ['drawingPresent:product:not-present'];
  }
  if (!actual || typeof actual.ok !== 'boolean') return ['drawingPresent:stored:unreadable'];
  if (actual.ok) return [];
  return actual.reason === 'drawing-not-found' ? ['drawingPresent:stored:not-present'] : [`drawingPresent:stored:unreadable:${actual.reason}`];
}

/** storedConditionalFormatRulesOf(readConditionalFormattingResult) -> {ok:true,rules:[...]} --
 * normalizes `readConditionalFormatting()`'s "no block at this exact sqref at all" outcome
 * (`{ok:false, reason:'conditionalFormatting-not-present'}`) into an EMPTY rule list (the SAME
 * "absence means nothing here" rule `storedMergeRangesOf` already established for merges), never
 * a read failure. Any OTHER failure reason (a malformed rule, an unreadable styles part, ...)
 * stays a named failure, untouched. */
function storedConditionalFormatRulesOf(readConditionalFormattingResult) {
  if (readConditionalFormattingResult && readConditionalFormattingResult.ok === false && readConditionalFormattingResult.reason === 'conditionalFormatting-not-present') return { ok: true, rules: [] };
  if (readConditionalFormattingResult && readConditionalFormattingResult.ok === true) return { ok: true, rules: readConditionalFormattingResult.value.rules };
  return { ok: false, reason: (readConditionalFormattingResult && readConditionalFormattingResult.reason) || 'unreadable' };
}

/** conditionalFormatAbsentDiffs(expected, actual, authority) -> [named difference, ...]. PROVES
 * ABSENCE ONLY, the `delete_conditional_format` companion to `add_conditional_format`'s own
 * `conditionalFormatRuleDiffs` below -- this one only ever asks "is the rule list at this
 * sqref/range now empty", unconditionally (no type/operator narrowing needed for absence).
 * `expected` ({sheet,range}) is carried for the caller's own lookup only, never compared here.
 * `actual` 'product': `read_conditional_formats`'s own `{rules}` projection (or `undefined`/
 * non-array `rules` if the read failed). 'stored': `storedConditionalFormatRulesOf()`'s own
 * `{ok,rules}` (or a non-object/`ok:false` reader failure). */
function conditionalFormatAbsentDiffs(expected, actual, authority) {
  if (!actual || !Array.isArray(actual.rules)) return [`conditionalFormat:${authority}:unreadable`];
  return actual.rules.length === 0 ? [] : [`conditionalFormat:${authority}:still-present`];
}

/**
 * conditionalFormatRuleDiffs(expected, actual, authority) -> [named difference, ...]. The
 * GENERAL rule-content comparator narrowed to the ONE CF type `xl13.ROUNDS`'s own
 * `add_conditional_format` round already writes: `aboveAverage` (no operator concept -- it is
 * never one of `CF_OPERATORS`). No vocabulary translation is needed on either side for this
 * type: `read_conditional_formats`'s product projection (cf-read.cjs's `aboveAverage()`,
 * `BY_TYPE.xlAboveAverageCondition`) already reports the write schema's own words
 * (`{type:'aboveAverage', above, numStdDev}`) -- the engine->write-schema translation cf-read.cjs
 * does for every CF type, not special-cased here. On the OOXML side, `aboveAverage` is one of
 * `test/live/file-verifier.cjs`'s own `CF_TYPES`/`X14_CF_TYPES` entries -- its `@type` attribute
 * is this exact literal string (ECMA-376 Part 1 §18.18.12 `ST_CfType`), read by
 * `readAboveAverageRule()` into `{type:'aboveAverage', above, stdDev}` -- same field names as the
 * product side, modulo `numStdDev`/`stdDev`'s own naming (not asserted here, see below).
 *   expected  { range, above } (`numStdDev`/`stdDev` is not asserted -- it defaults to 0 on both
 *             sides and the round's own task never sets it).
 *   actual  'product': `read_conditional_formats`'s own `{rules}` projection, write-schema
 *           vocabulary directly, no translation.
 *           'stored': `storedConditionalFormatRulesOf()`'s own `{ok,rules}`, OOXML-reader shape.
 */
function conditionalFormatRuleDiffs(expected, actual, authority) {
  if (!actual || !Array.isArray(actual.rules)) return [`conditionalFormatRule:${authority}:unreadable`];
  const found = actual.rules.some(r => r && r.type === 'aboveAverage' && r.above === expected.above);
  return found ? [] : [`conditionalFormatRule:${authority}:not-present`];
}

/**
 * GH #141 comparator-gaps batch 1 (2026-10-18) -- two MORE pure comparators, for `freeze_panes`/
 * `unfreeze_panes` and `filter_range`/`clear_filter`, the SAME `storedOf(sheet)=>buf=>storedResult`
 * / `diffsOf({expect,verifyResults,stored})=>[named difference,...]` NON_CELL_KINDS pattern every
 * kind above already uses.
 */

/** freezePanesDiffs(expected, actual, authority) -> [named difference, ...]. `expected`:
 * `{ state: 'none' }` | `{ state: 'frozen', row0, col0 }` -- the SAME shape freeze.cjs's own
 * `op.expect` already uses (that family's `same(p, e)` comparator, cited: `p.state === e.state &&
 * (e.state === 'none' || (p.row0 === e.row0 && p.col0 === e.col0))`).
 *   actual  'product': `read_freeze_panes`'s own projection (`panes-read.cjs`'s `projectPanesRead`)
 *           -- `{ sheet, frozen: false }` | `{ sheet, frozen: true, rows?, columns? }` | `{ sheet,
 *           frozen: null, reason }` (a gap, `rows`/`columns` ABSENT when 0, mirroring the write
 *           schema's own "present only when > 0").
 *           'stored': `readFreezePanes()`'s own `{ state, row0?, col0? }` value.
 */
function freezePanesDiffs(expected, actual, authority) {
  if (authority === 'product') {
    if (!actual || typeof actual !== 'object' || typeof actual.frozen !== 'boolean') return ['freezePanes:product:unreadable'];
    if (expected.state === 'none') return actual.frozen === false ? [] : ['freezePanes:product:still-frozen'];
    if (actual.frozen !== true) return ['freezePanes:product:not-frozen'];
    return ((actual.rows || 0) === expected.row0 && (actual.columns || 0) === expected.col0) ? [] : ['freezePanes:product:wrong-pane'];
  }
  if (!actual || typeof actual !== 'object' || typeof actual.state !== 'string') return ['freezePanes:stored:unreadable'];
  if (expected.state === 'none') return actual.state === 'frozen' ? ['freezePanes:stored:still-frozen'] : [];
  if (actual.state !== 'frozen') return ['freezePanes:stored:not-frozen'];
  return (actual.row0 === expected.row0 && actual.col0 === expected.col0) ? [] : ['freezePanes:stored:wrong-pane'];
}

/** autoFilterDiffs(expected, actual, authority) -> [named difference, ...]. PROVES STRUCTURE ONLY
 * (an AutoFilter is active on the exact range / no AutoFilter is active at all) -- NOT the
 * field/value criteria content (the same OOXML<->engine vocabulary boundary `set_validation`'s own
 * gap entry already refuses; `read_filter`'s own `filters[]` projection carries that content, but
 * no comparator here decodes it). `expected`: `{ active: true, range }` (filter_range) |
 * `{ active: false }` (clear_filter).
 *   actual  'product': `read_filter`'s own projection (`filter-read.cjs`'s `projectFilterRead`) --
 *           `{ sheet, range, filters, engineShapes }`; "active" = at least one `filters[]` entry
 *           whose `on` is not `false` (the SAME `active()` helper filter.cjs's own `classify` uses).
 *           'stored': `readAutoFilter()`'s own `{ present, ref, filterColumns }` value --
 *           filter.cjs's own header comment documents that `clear_filter`'s `ShowAllData()` "keeps
 *           the controls" (the AutoFilter range persists) while removing the active criteria, so
 *           "cleared" is decided by `filterColumns === 0`, never by the bare element's absence.
 */
function autoFilterDiffs(expected, actual, authority) {
  if (authority === 'product') {
    if (!actual || typeof actual !== 'object' || !Array.isArray(actual.filters)) return ['autoFilter:product:unreadable'];
    const activeCount = actual.filters.filter(f => f && f.on !== false).length;
    if (!expected.active) return activeCount === 0 ? [] : ['autoFilter:product:still-active'];
    if (actual.range === null) return ['autoFilter:product:no-range'];
    if (String(actual.range).replace(/\$/g, '') !== expected.range) return ['autoFilter:product:wrong-range'];
    return activeCount > 0 ? [] : ['autoFilter:product:not-active'];
  }
  if (!actual || typeof actual !== 'object' || typeof actual.present !== 'boolean' || typeof actual.filterColumns !== 'number') return ['autoFilter:stored:unreadable'];
  if (!expected.active) return actual.filterColumns === 0 ? [] : ['autoFilter:stored:still-active'];
  if (!actual.present) return ['autoFilter:stored:missing'];
  if (actual.ref !== expected.range) return ['autoFilter:stored:wrong-range'];
  return actual.filterColumns > 0 ? [] : ['autoFilter:stored:not-active'];
}

/**
 * GH #141 comparator-gaps batch 2 (2026-10-18) -- two more pure comparators, for
 * `set_page_layout` and `move_chart`. `set_chart`/`copy_chart`/`insert_shape`/`create_pivot` reuse
 * `chartList`/`drawingPresent`/`sheetListDiffs` (via a new `pivotList` kind) unchanged -- no new
 * comparator code for those four.
 */

/** pageOrientationDiffs(expected, actual, authority) -> [named difference, ...]. `expected`:
 * { orientation: 'portrait'|'landscape' } (the two writable values, page.cjs's own ORIENT map).
 * `actual` 'product': `read_page_setup`'s own `.print` half's `orientation` field (page-read.cjs's
 * `ORIENT_OF`-translated string, or `null`/`undefined` on a gap -- never guessed). 'stored':
 * `readPageSetup()`'s own `orientation` value (ECMA-376 `@orientation`, same string vocabulary,
 * no translation needed, the SAME convention `printFitDiffs`'s fitToWidth/fitToHeight already use). */
function pageOrientationDiffs(expected, actual, authority) {
  if (!actual || typeof actual !== 'object' || typeof actual.orientation !== 'string') return [`pageOrientation:${authority}:missing`];
  return actual.orientation === expected.orientation ? [] : [`pageOrientation:${authority}:orientation`];
}

/** chartPositionDiffs(expected, actual, authority) -> [named difference, ...]. `expected`:
 * { col, row } (0-based, the SAME `op.at.col`/`op.at.row` the `move_chart` write schema takes --
 * charts.cjs's own `SetPosition(col, 0, row, 0)`, cell-aligned, offset always 0).
 *   actual  'product': ONE raw `read_drawings` record (the CALLER has already found it by name --
 *           `drawings.cjs`'s own doc: `pos: { fromCol, colOffset, fromRow, rowOffset }`) --
 *           `undefined` if the record was not found at all (whole-sheet gap or no drawing of that
 *           name), compared via `pos.fromCol`/`pos.fromRow` directly (same field names, no
 *           translation -- `op.at.col`/`row` map onto them with zero offset, per the family's own
 *           measured call shape).
 *           'stored': `readDrawingAnchor()`'s full `{ok, value|reason}` result (the SAME shape
 *           `drawingPresentDiffs`'s stored branch already takes) -- its ADDITIVE `value.from`
 *           field (`{col,row}` or `null` on an anchor type with no position, e.g. absoluteAnchor).
 */
function chartPositionDiffs(expected, actual, authority) {
  if (authority === 'product') {
    if (!actual || typeof actual !== 'object' || !actual.pos || typeof actual.pos !== 'object') return ['chartPosition:product:missing'];
    const { fromCol, fromRow } = actual.pos;
    if (!Number.isInteger(fromCol) || !Number.isInteger(fromRow)) return ['chartPosition:product:missing'];
    return (fromCol === expected.col && fromRow === expected.row) ? [] : ['chartPosition:product:position'];
  }
  if (!actual || typeof actual.ok !== 'boolean') return ['chartPosition:stored:unreadable'];
  if (!actual.ok) return actual.reason === 'drawing-not-found' ? ['chartPosition:stored:not-present'] : [`chartPosition:stored:unreadable:${actual.reason}`];
  const from = actual.value.from;
  if (!from) return ['chartPosition:stored:no-position'];
  return (from.col === expected.col && from.row === expected.row) ? [] : ['chartPosition:stored:position'];
}

/** layoutWidthDiffs(expected, actual, authority) -> [named difference, ...]. `expected`:
 * { col, width } -- `width` in the WRITE schema's own units (layout.cjs: the engine's character-
 * width unit, (0, 255]), the SAME unit `read_layout`'s projection reports (layout-read.cjs's own
 * doc comment: "the MIRROR of the layout_range write schema ... the same units").
 *   actual  'product': `read_layout`'s `columns[col]` entry (layout-read.cjs `entryOf()`) --
 *           `{ width, hidden }` or a gap (`hidden: true` / `reason` set, no numeric `width`).
 *           Compared DIRECTLY against `expected.width`, no translation: both sides speak the same
 *           unit. A requested width of 20 round-trips to a canonical readback of 20 with ZERO
 *           pixel-quantisation drift (the engine's own column-width chain, MDW=7/PAD=5:
 *           `floor((20*7+5)/7*256)/256` quantises to the SAME mcw a plain 20 already is) -- this
 *           comparator does not attempt the general fractional-width case (M7-COLUMN-WIDTH-
 *           CONTRACT-20261005.md: that needs a tolerance, an open design question), only this
 *           round's own exact, drift-free value.
 *           'stored': `readColGeometry()`'s `.width` -- the OOXML `cols/col/@width`, NOT the
 *           character count: ECMA-376 18.3.1.13's own padded form, `floor((chars*MDW+5)/MDW*256)/256`
 *           with MDW (Maximum Digit Width) = 7 px, this workbook's default font (Calibri 11 @
 *           96dpi) -- LIVE MEASURED (`test/live/xlsx-xl13-8-bounded-live-acceptance.live.cjs`:
 *           a requested width of 20 persists as `20.7109375`, which this exact formula
 *           reproduces). `storedColWidthForChars()` below is that formula, not a fresh guess.
 */
const LAYOUT_WIDTH_MDW_PX = 7;
function storedColWidthForChars(chars) { return Math.floor(((chars * LAYOUT_WIDTH_MDW_PX + 5) / LAYOUT_WIDTH_MDW_PX) * 256) / 256; }
function layoutWidthDiffs(expected, actual, authority) {
  if (!actual || typeof actual !== 'object' || typeof actual.width !== 'number') return [`layoutWidth:${authority}:unreadable`];
  const want = authority === 'stored' ? storedColWidthForChars(expected.width) : expected.width;
  return actual.width === want ? [] : [`layoutWidth:${authority}:wrong-width`];
}

/** runBoldDiffs(expected, actual, authority) -> [named difference, ...]. `expected`:
 * { text, bold } -- the round's own task writes ONE run covering the WHOLE cell text with a
 * single font field (`bold`), so BOTH authorities must report exactly one run spanning that text;
 * more than one run means the format did not land uniformly (a genuine mismatch, never widened
 * away). Only `bold` is compared -- the task sets no other font field, and the two authorities
 * represent an UNSET field differently (product: the engine's own concrete default, e.g.
 * `italic: false`, a real font name/size; stored: `null`, ECMA-376's "rPr absent" -- GH136/GH137's
 * own boundary, not a new one), so this comparator does not claim anything about them.
 *   actual  'product': `read_runs`'s own `runMap` (`{ text, runs: [{ start, length, font }] }`,
 *           xlsx-driver.cjs's per-character merge) -- `runs[0].font.bold`.
 *           'stored': `readRichTextRuns()`'s own `{ runs: [{ text, bold, ... }] }` (sharedStrings
 *           `<si><r><rPr><b/></rPr></r></si>`, X-F04/GH137) -- `runs[0].bold`.
 */
function runBoldDiffs(expected, actual, authority) {
  if (authority === 'product') {
    if (!actual || typeof actual !== 'object' || !Array.isArray(actual.runs)) return ['runBold:product:unreadable'];
    if (actual.text !== expected.text) return ['runBold:product:wrong-text'];
    if (actual.runs.length !== 1) return ['runBold:product:run-count'];
    const font = actual.runs[0].font;
    return font && font.bold === expected.bold ? [] : ['runBold:product:wrong-bold'];
  }
  if (!actual || typeof actual !== 'object' || !Array.isArray(actual.runs)) return ['runBold:stored:unreadable'];
  if (actual.runs.length !== 1) return ['runBold:stored:run-count'];
  const r = actual.runs[0];
  if (r.text !== expected.text) return ['runBold:stored:wrong-text'];
  return r.bold === expected.bold ? [] : ['runBold:stored:wrong-bold'];
}

module.exports = {
  cellDiffs, productCells, normFormula, hex,
  definedNamesByName, definedNameDiffs, storedPrintAreaOf, printAreaDiffs, printFitDiffs,
  mergeDiffs, storedMergeRangesOf, sheetListDiffs, validationAbsentDiffs, validationRuleDiffs,
  definedNameAbsentDiffs, storedPrintTitlesOf, productPrintTitlesOf, printTitlesDiffs,
  printScaleDiffs, PAPER_SIZE_ID, paperSizeDiffs, firstPageNumberDiffs, printCenteringDiffs,
  headerFooterDiffs, headerFooterFlagDiffs, pageBreakDiffs,
  chartTitleDiffs, chartAxisTitleDiffs, CHART_LEGEND_POS_OOXML, chartLegendDiffs,
  drawingPresentDiffs, storedConditionalFormatRulesOf, conditionalFormatAbsentDiffs, conditionalFormatRuleDiffs,
  freezePanesDiffs, autoFilterDiffs,
  pageOrientationDiffs, chartPositionDiffs,
  storedColWidthForChars, layoutWidthDiffs, runBoldDiffs,
};
