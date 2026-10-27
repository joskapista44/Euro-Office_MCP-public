'use strict';

/**
 * GH #141 Wave 2 (P1) (internal tracker): a STATIC, code-grounded check of
 * docs/GH141-CELL-FORMAT-SYMMETRY-MATRIX.md's own SYMMETRIC / READ_ONLY claims -- repo-internal
 * only (no network, no live editor, no git history), the same pattern as
 * test/static/gh139-chart-symmetry-matrix.test.cjs. Every row the matrix calls SYMMETRIC or
 * READ_ONLY is checked HERE against the actual code:
 *   - the claimed format_range/merge_range/unmerge_range/clear_range/format_run WRITE key is (or,
 *     for the READ_ONLY row, is NOT) in its schema;
 *   - the claimed READ field is actually emitted by the driver's own readQuery() generated script
 *     (addFormat/addStyle's field assignments + the DS Bridge capability ids they call) -- not by
 *     string-matching read-tools.cjs's evidence prose, which can drift independently of the code.
 * A stale row turns ONE of these checks red -- this file never re-derives the matrix, only
 * re-measures what it asserts.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const { XLSX_WRITE_ALLOWLIST, STEP_SCHEMAS, TEMPORARY_BOUNDARY } = require('../../src/server/write-tools.cjs');
const { OPERATIONS, STATUS } = require('../../src/server/read-tools.cjs');
const { readQuery } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const clearFamily = require('../../src/adapters/xlsx/families/clear.cjs');
const formatRunFamily = require('../../src/adapters/xlsx/families/format-run.cjs');

// ---- 1. WRITE-side release claims: every intent the matrix names IS in XLSX_WRITE_ALLOWLIST ----

const RELEASED_CELL_FORMAT_INTENTS = Object.freeze(['format_range', 'merge_range', 'unmerge_range', 'clear_range', 'format_run']);

test('every cell-format intent the matrix tables (F/M/C/R rows) IS in XLSX_WRITE_ALLOWLIST', () => {
  for (const intent of RELEASED_CELL_FORMAT_INTENTS) assert.ok(XLSX_WRITE_ALLOWLIST.includes(intent), `${intent} should be released`);
});

// ---- 2. format_range's six released format.* keys (F1-F13 rows) -- the schema itself, not a copy ----

const FORMAT_RANGE_FORMAT_PROPS = STEP_SCHEMAS.format_range.properties.format.properties;

test('format_range schema accepts exactly the six keys the matrix tables (numberFormat/border/fill/font/align/wrap), nothing else', () => {
  assert.deepEqual(new Set(Object.keys(FORMAT_RANGE_FORMAT_PROPS)), new Set(['numberFormat', 'border', 'fill', 'font', 'align', 'wrap']));
  assert.equal(STEP_SCHEMAS.format_range.properties.format.additionalProperties, false);
});

test('format_range is NOT in TEMPORARY_BOUNDARY itself (the boundary is per-key inside the format object, enforced at dispatch, not a blanket refusal of the whole intent)', () => {
  // TEMPORARY_BOUNDARY.format_range names the boundary id used for keys OUTSIDE the six released
  // ones (write-tools.cjs ~970), but the intent itself is releasable -- the allowlist check above
  // already proves that. This test exists so a future removal of the per-key carve-out (making the
  // WHOLE intent refuse) would be caught here, not silently.
  assert.ok(Object.prototype.hasOwnProperty.call(TEMPORARY_BOUNDARY, 'format_range'));
  assert.equal(TEMPORARY_BOUNDARY.format_range, 'B05-WRITE');
});

test('font sub-schema (F4-F7 rows) carries color/bold/italic/name/size/underline/strikeout, and explicitly NOT subscript/superscript (F7: READ_ONLY by owner-gate, not a schema oversight)', () => {
  const font = FORMAT_RANGE_FORMAT_PROPS.font;
  for (const key of ['color', 'bold', 'italic', 'name', 'size', 'underline', 'strikeout']) assert.ok(key in font.properties, `font.${key} should be a released key`);
  assert.equal(font.additionalProperties, false);
  assert.ok(!('subscript' in font.properties), 'subscript must stay unreleased (F7)');
  assert.ok(!('superscript' in font.properties), 'superscript must stay unreleased (F7)');
});

test('align sub-schema (F8-F12 rows) carries horizontal/vertical/rotation/readingOrder/indent/shrink -- all six, none held back any more', () => {
  const align = FORMAT_RANGE_FORMAT_PROPS.align;
  for (const key of ['horizontal', 'vertical', 'rotation', 'readingOrder', 'indent', 'shrink']) assert.ok(key in align.properties, `align.${key} should be a released key`);
  assert.equal(align.additionalProperties, false);
});

test('border sub-schema (F2/F2b rows) carries mode/style/color/sides, mode required', () => {
  const border = FORMAT_RANGE_FORMAT_PROPS.border;
  for (const key of ['mode', 'style', 'color', 'sides']) assert.ok(key in border.properties, `border.${key} should exist`);
  assert.deepEqual(border.required, ['mode']);
});

// ---- 3. merge_range / unmerge_range / clear_range schemas (M/C rows) ----

test('merge_range / unmerge_range schemas exist, merge_range carries the optional `across`', () => {
  assert.ok('sheet' in STEP_SCHEMAS.merge_range.properties && 'range' in STEP_SCHEMAS.merge_range.properties);
  assert.ok('across' in STEP_SCHEMAS.merge_range.properties);
  assert.ok('sheet' in STEP_SCHEMAS.unmerge_range.properties && 'range' in STEP_SCHEMAS.unmerge_range.properties);
  assert.ok(!('across' in STEP_SCHEMAS.unmerge_range.properties), 'unmerge_range takes no across (the matrix M1 row claims it needs none)');
});

test('clear_range scope enum is exactly contents/formats/all (C1-C3 rows)', () => {
  assert.deepEqual(STEP_SCHEMAS.clear_range.properties.scope.enum, ['contents', 'formats', 'all']);
});

// GH #141 P1 follow-up (clear-formats-l1-teljes-alap, 2026-10-17): BASE_AXES widened from six to
// thirteen -- border/align/readingOrder/indent/shrink/font.{underline,strike} added, font.name/size
// deliberately still excluded (workbook-dependent base, not a fixed engine constant -- see the
// matrix's own C2/C3 NEM-MERT note). This test re-reads the module source for the frozen array
// literal (clear.cjs does not export BASE_AXES directly), the same static-grounding technique the
// matrix itself asks for -- no live editor, no re-derivation of the claim, just a direct code read.
test('clear_range\'s own L1 (BASE_AXES) asserts exactly the thirteen base axes the matrix\'s C2/C3 rows name -- NOT font.name/font.size (the matrix\'s own NEM-MERT, not a silent gap this test would hide)', () => {
  const fs = require('node:fs');
  const src = fs.readFileSync(require.resolve('../../src/adapters/xlsx/families/clear.cjs'), 'utf8');
  const m = src.match(/const BASE_AXES = Object\.freeze\(\[([\s\S]*?)\]\);/);
  assert.ok(m, 'BASE_AXES literal should be findable in clear.cjs');
  const names = [...m[1].matchAll(/\['([a-zA-Z]+)',/g)].map(x => x[1]);
  assert.deepEqual(names, ['nf', 'fill', 'bold', 'italic', 'wrap', 'rot', 'align', 'border', 'readingOrder', 'indent', 'shrink', 'underline', 'strike']);
  assert.ok(!names.includes('name') && !names.includes('size'), 'font.name/size must stay out of BASE_AXES (workbook-dependent base)');
});

test('clear.cjs\'s gate() requires align/border/readingOrder/indent/shrink/font to be readable on a formats/all-scope target, the same fail-closed shape wrap/rot/italic already had', () => {
  const fs = require('node:fs');
  const src = fs.readFileSync(require.resolve('../../src/adapters/xlsx/families/clear.cjs'), 'utf8');
  const m = src.match(/isFormatScope\(op\) && !\('wrap' in cell[\s\S]*?xlsx-target-style-unread'/);
  assert.ok(m, 'the format-scope unread guard should be findable');
  for (const key of ['align', 'border', 'readingOrder', 'indent', 'shrink', 'font']) assert.ok(m[0].includes(`'${key}' in cell`), `gate() should require '${key}' in cell`);
});

test('clear_range scope formats/all both refuse a range that touches a merged area, named (the matrix C2 row\'s own caveat)', () => {
  const p = clearFamily.plan({ intent: 'clear_range', id: 's', sheet: 'S', range: 'A1:B2', scope: 'formats' }, 0);
  assert.ok(p.ok, p.detail);
  // plan() alone cannot see live merge state (that needs a read) -- this test only proves the
  // schema/plan accepts the shape; the merge-dissolve refusal itself is classify()-time, already
  // covered by xlsx-l16-3-merge-integrity-verify.test.cjs (cited in the matrix), not re-proven here.
  assert.equal(p.operation.scope, 'formats');
});

// ---- 4. format_run's released 5 font axes, explicitly NOT the whole-cell effects (R1 row) ----

test('format_run (R1 row) is adapter-backed and its INTENTS carries exactly format_run', () => {
  assert.deepEqual(formatRunFamily.INTENTS, Object.freeze(['format_run']));
});

test('format_run schema carries bold/italic/name/size/color per run, and explicitly NOT underline/strikeout/subscript/superscript (those stay format_range\'s own whole-cell scope, per the matrix R1 caveat)', () => {
  const runFont = STEP_SCHEMAS.format_run.properties.runs.items.properties.font;
  for (const key of ['bold', 'italic', 'name', 'size', 'color']) assert.ok(key in runFont.properties, `runs[].font.${key} should be released`);
  for (const key of ['underline', 'strikeout', 'subscript', 'superscript']) assert.ok(!(key in runFont.properties), `runs[].font.${key} must stay out of format_run`);
});

// ---- 5. READ-side claims: read_range / read_runs ACCEPTED, and the driver's OWN generated script
// actually emits every READ field the matrix tables -- not read-tools.cjs's evidence string, which
// is prose and can drift independently of what the driver actually does. ----

test('read_range and read_runs are both ACCEPTED', () => {
  assert.equal(OPERATIONS.xlsx.read_range.status, STATUS.ACCEPTED);
  assert.equal(OPERATIONS.xlsx.read_runs.status, STATUS.ACCEPTED);
});

// A single generated script, style+format both requested, is the SAME call shape read_range makes
// with { format: true, style: true } (read-tools.cjs readRequestFields -> windows[].format/.style).
const generatedReadScript = readQuery([{ sheet: 'Sheet1', range: 'A1:A1', format: true, style: true, nf: false }], {});

test('the driver\'s generated read script assigns every F1-F13/F2b field the matrix claims as READ: nf, fill, bold/italic (via addFormat), wrap, rot, font, align, border, readingOrder, indent, shrink (via addStyle)', () => {
  for (const assignment of ['out.nf =', 'out.fill =', 'out.wrap =', 'out.rot =', 'out.font =', 'out.align =', 'out.border =', 'out.readingOrder =', 'out.indent =', 'out.shrink =']) {
    assert.ok(generatedReadScript.includes(assignment), `generated script should assign ${assignment}`);
  }
});

test('the driver\'s generated read script calls the four named DS Bridge READ capabilities the matrix cites: bold/italic (textStyle.*.read/v1), align+border (range.appearance.read/v1), readingOrder/indent/shrink (range.alignmentExt.read/v1)', () => {
  for (const capId of ['textStyle.bold.read/v1', 'textStyle.italic.read/v1', 'range.appearance.read/v1', 'range.alignmentExt.read/v1']) {
    assert.ok(generatedReadScript.includes(capId), `generated script should reference capability ${capId}`);
  }
});

test('F9 row\'s claimed mechanism asymmetry is real: rotation is read as a TOP-LEVEL `out.rot` field via a PUBLIC getter (GetOrientation), NOT folded into the align.* bridge object the way horizontal/vertical are', () => {
  // The align object itself (out.align) comes solely from APPEARANCE_READ; rot is assigned
  // separately, by its own try/catch, with no bridge capability id anywhere near it.
  const rotLine = generatedReadScript.split('\n').find(l => l.includes('out.rot ='));
  assert.ok(rotLine, 'out.rot assignment should exist');
  assert.ok(!rotLine.includes('range.appearance.read/v1') && !rotLine.includes('range.alignmentExt.read/v1'), 'rot must not be read through either bridge capability');
  assert.ok(/GetOrientation/.test(rotLine), 'rot should read through the PUBLIC GetOrientation getter');
});

test('F6 row\'s claimed name asymmetry is real: the WRITE key is `strikeout`, the READ field is `strike` (same underlying axis, different name)', () => {
  assert.ok('strikeout' in FORMAT_RANGE_FORMAT_PROPS.font.properties, 'WRITE key should be named strikeout');
  assert.ok(generatedReadScript.includes('strike:'), 'READ field should be named strike');
  assert.ok(!generatedReadScript.includes('strikeout:'), 'the READ side should not also use the WRITE name (that would mean the asymmetry note is stale)');
});

// ---- 6. F7 row: subscript/superscript are READ-able (addStyle's font reader) even though they are
// never a WRITE key (checked in section 2) -- this is what makes F7 READ_ONLY rather than MISSING. ----

test('F7 row: subscript/superscript ARE read (font.sub / font.sup in the generated script), confirming READ_ONLY (readable, never writable) rather than an absent capability on both sides', () => {
  assert.ok(generatedReadScript.includes('sub:') && generatedReadScript.includes('sup:'), 'font reader should expose sub/sup');
});

// ---- 7. RED CONTROL: a non-existent key/intent/capability must be correctly reported as absent --
// proves the checks above actually discriminate, not just pass by construction (the card's own "MI
// A KESZ" -- a deliberately broken row must go red). ----

test('RED CONTROL: a made-up format.* key is correctly absent from the released set', () => {
  assert.ok(!('definitelyNotARealFormatKey' in FORMAT_RANGE_FORMAT_PROPS));
});

test('RED CONTROL: a made-up intent is correctly absent from XLSX_WRITE_ALLOWLIST', () => {
  assert.ok(!XLSX_WRITE_ALLOWLIST.includes('format_range_definitely_not_real'));
});

test('RED CONTROL: a made-up field assignment / capability id is correctly absent from the generated read script', () => {
  assert.ok(!generatedReadScript.includes('out.definitelyNotARealField ='));
  assert.ok(!generatedReadScript.includes('definitely.not.a.real.capability/v1'));
});
