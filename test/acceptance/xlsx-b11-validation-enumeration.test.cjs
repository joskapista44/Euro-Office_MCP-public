'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

/**
 * B11-V1 -- range.validations.read/v1 Bridge capability, OFFLINE (card
 * an internal branch; dispatch, tracker 9580767b
 * #ref). DESIGN + OFFLINE IMPLEMENTATION + OFFLINE TEST + PATCH_BUILD only -- NOT a live apply,
 * NOT a driver cutover (see docs/B11-VALIDATION-ENUMERATION-BRIDGE-DESIGN-20261007.md).
 *
 * The GAP this covers (measured live, agents/an internal branch/handoff/maintainer-b11-bridge-20261007/
 * b11-public-first-run3.json): the public ApiRange.GetValidation() is sufficient for a range
 * covered by exactly one data validation (reads back correctly in-session, after reopen, and in
 * the saved file), but THROWS "Multiple validations exist" on a range whose area is covered by
 * more than one. Production's own per-sheet internal reader (xlsx-driver.cjs dependentsInFrame,
 * `wb.getWorksheet(i).dataValidations.elems`) already reads {type, ranges, formula1, formula2} for
 * every validation on a sheet; this capability exposes the SAME field set, scoped to one query
 * range, as a sanctioned bridge read instead of a direct internal dependency.
 *
 * ROUTING DECISION (the owner, owner channel): route EVERY case through the internal model, never
 * public-first-then-internal -- "determinisztikus szemantika... fontosabb, mint hogy minden
 * single case-t eroltetve publicon tartsunk." A single-validation query and a ten-validation query
 * differ only in the length of the returned list, not the code path or the error semantics.
 *
 * This test evaluates propertyQuery's sibling -- the ACTUAL bridge-block.js text -- via
 * `new Function('Api', source)`, the same sandboxed-Api pattern as
 * test/static/ds-bridge-client.test.cjs and the W-B1920 word-bridge-runtime test, against a fake
 * Api exposing GetSheet/GetRange (public, MEASURED: GetRow/GetCol/GetRowsCount/GetColumnsCount, b11
 * run3 axis3) and ws.worksheet.dataValidations.elems (the SAME internal shape dependentsInFrame
 * already reads).
 *
 * HYPOTHESIS, not LIVE-verified (named in the capability's own comment and the design doc): the
 * internal elems' r1/c1 are 0-based, matching GetRow()-1/GetCol()-1. If a LIVE measurement shows
 * otherwise, the capability's -1 offset must be corrected before this ships.
 */

const BLOCK_PATH = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.4.60', 'euro-office-bridge-v1', 'bridge-block.js');
const SOURCE = fs.readFileSync(BLOCK_PATH, 'utf8').replace(/,\s*$/, ''); // the file's own trailing comma belongs to its larger sdk-all.js context, not a standalone statement

function colOf(letters) { let n = 0; for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64); return n; }
function a1ToBounds(a1) {
  const m = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(a1);
  if (!m) throw new Error(`bad a1: ${a1}`);
  const c1 = colOf(m[1]); const r1 = Number(m[2]);
  const c2 = m[3] ? colOf(m[3]) : c1; const r2 = m[4] ? Number(m[4]) : r1;
  return { row: Math.min(r1, r2), col: Math.min(c1, c2), rows: Math.abs(r2 - r1) + 1, cols: Math.abs(c2 - c1) + 1 };
}

/** `elems` undefined means NO sheet (sheet-not-found path is tested separately); null means ws.worksheet.dataValidations itself is null (no validations, measured "[] if proven empty" contract). */
function fakeApi(elems, { sheetExists = true } = {}) {
  const ws = {
    GetRange(a1) {
      const b = a1ToBounds(a1);
      return { GetRow: () => b.row, GetCol: () => b.col, GetRowsCount: () => b.rows, GetColumnsCount: () => b.cols };
    },
    worksheet: { dataValidations: elems === null ? null : { elems } },
  };
  return { GetSheet: name => (sheetExists && name === 'S' ? ws : undefined) };
}

function bridgeFor(elems, opts) {
  return new Function('Api', `${SOURCE}; return Api.GetEuroOfficeBridge();`)(fakeApi(elems, opts));
}

const read = (elems, args, opts) => bridgeFor(elems, opts).read('range.validations.read/v1', args);

// Internal-model fixtures, 0-based r1/c1 (the stated HYPOTHESIS), matching dependentsInFrame's shape.
const V_E1 = { type: 6, ranges: [{ r1: 0, c1: 4, r2: 0, c2: 4 }], formula1: { text: '=D!$A$1' }, formula2: null };
const V_E5 = { type: 6, ranges: [{ r1: 4, c1: 4, r2: 4, c2: 4 }], formula1: { text: '=D!$A$1' }, formula2: null };
const V_WHOLE = { type: 1, ranges: [{ r1: 9, c1: 0, r2: 9, c2: 0 }], formula1: { text: '1' }, formula2: { text: '10' } };
const V_BLOCK = { type: 6, ranges: [{ r1: 0, c1: 0, r2: 9, c2: 9 }], formula1: { text: 'x' }, formula2: null };

test('0 validations: [] -- never a reason in place of an empty list', () => {
  assert.deepEqual(read([], { sheet: 'S', range: 'A1' }), { ok: true, value: [] });
});

test('0 validations via ws.worksheet.dataValidations === null -- also []', () => {
  assert.deepEqual(read(null, { sheet: 'S', range: 'A1' }), { ok: true, value: [] });
});

test('1 validation: a direct hit returns it; a miss outside its range returns []', () => {
  assert.deepEqual(read([V_E1], { sheet: 'S', range: 'E1' }), { ok: true, value: [{ type: 6, ranges: [{ r1: 0, c1: 4, r2: 0, c2: 4 }], formula1: '=D!$A$1', formula2: null }] });
  assert.deepEqual(read([V_E1], { sheet: 'S', range: 'F1' }), { ok: true, value: [] });
});

test('multiple IDENTICAL validations (same type/formula, different ranges) on a wide query -- both returned, deterministic order (public GetValidation() would throw here)', () => {
  const r = read([V_E1, V_E5], { sheet: 'S', range: 'A1:Z100' });
  assert.equal(r.ok, true);
  assert.equal(r.value.length, 2);
  assert.deepEqual(r.value.map(x => x.ranges[0]), [{ r1: 0, c1: 4, r2: 0, c2: 4 }, { r1: 4, c1: 4, r2: 4, c2: 4 }], 'sorted by r1,c1');
});

test('multiple DIFFERENT validations (different type/operator/formula) on a wide query -- both returned', () => {
  const r = read([V_E1, V_WHOLE], { sheet: 'S', range: 'A1:Z100' });
  assert.equal(r.ok, true);
  assert.deepEqual(r.value.map(x => x.type), [6, 1], 'sorted by r1,c1 -- V_WHOLE at r1=9 comes after V_E1 at r1=0');
});

test('PARTIALLY OVERLAPPING: a query range fully inside a single, larger validation range is still a hit', () => {
  const r = read([V_BLOCK], { sheet: 'S', range: 'B2:C3' });
  assert.deepEqual(r, { ok: true, value: [{ type: 6, ranges: [{ r1: 0, c1: 0, r2: 9, c2: 9 }], formula1: 'x', formula2: null }] });
});

test('PARTIALLY OVERLAPPING: a query range that only clips one edge of a validation range is still a hit', () => {
  const r = read([V_BLOCK], { sheet: 'S', range: 'I9:L12' }); // I9 (col 9, row 9) is inside r1..9,c1..9; L12 is outside
  assert.equal(r.ok, true);
  assert.equal(r.value.length, 1);
});

test('determinism: repeated calls on the SAME fixture return byte-identical JSON (order is not insertion-order-dependent by accident)', () => {
  const a = JSON.stringify(read([V_WHOLE, V_E1, V_E5], { sheet: 'S', range: 'A1:Z100' }));
  const b = JSON.stringify(read([V_E5, V_WHOLE, V_E1], { sheet: 'S', range: 'A1:Z100' })); // same set, different insertion order
  assert.equal(a, b, 'the sort key (first range r1,c1) is the elems ORDER-independent, not an accident of insertion order');
});

test('operator/alertStyle: read BEST-EFFORT when the model exposes them, OMITTED (not guessed) when it does not', () => {
  const withOperator = { ...V_E1, operator: 3, alertStyle: 1 };
  assert.deepEqual(read([withOperator], { sheet: 'S', range: 'E1' }).value[0].operator, 3);
  assert.deepEqual(read([withOperator], { sheet: 'S', range: 'E1' }).value[0].alertStyle, 1);
  const r = read([V_E1], { sheet: 'S', range: 'E1' });
  assert.ok(!('operator' in r.value[0]), 'operator key absent, not null/undefined -- never guessed');
  assert.ok(!('alertStyle' in r.value[0]));
});

// FAIL CLOSED branches
test('FAIL CLOSED: bad-args (missing range / missing sheet / wrong types)', () => {
  assert.deepEqual(read([], { sheet: 'S' }), { ok: false, reason: 'bad-args' });
  assert.deepEqual(read([], { range: 'A1' }), { ok: false, reason: 'bad-args' });
  assert.deepEqual(read([], { sheet: 1, range: 'A1' }), { ok: false, reason: 'bad-args' });
  assert.deepEqual(read([], null), { ok: false, reason: 'bad-args' });
});

test('FAIL CLOSED: sheet-not-found', () => {
  assert.deepEqual(read([], { sheet: 'NOPE', range: 'A1' }, { sheetExists: false }), { ok: false, reason: 'sheet-not-found' });
});

test('FAIL CLOSED: range-not-found (GetRange throws)', () => {
  const api = fakeApi([]);
  api.GetSheet('S').GetRange = () => { throw new Error('bad range'); };
  const bridge = new Function('Api', `${SOURCE}; return Api.GetEuroOfficeBridge();`)(api);
  assert.deepEqual(bridge.read('range.validations.read/v1', { sheet: 'S', range: 'A1' }), { ok: false, reason: 'range-not-found' });
});

test('FAIL CLOSED: dataValidations-missing (the model property itself absent, not null)', () => {
  const api = fakeApi([]);
  delete api.GetSheet('S').worksheet.dataValidations;
  const bridge = new Function('Api', `${SOURCE}; return Api.GetEuroOfficeBridge();`)(api);
  assert.deepEqual(bridge.read('range.validations.read/v1', { sheet: 'S', range: 'A1' }), { ok: false, reason: 'dataValidations-missing' });
});

test('FAIL CLOSED: dataValidations-elems-not-a-list (malformed model shape)', () => {
  const api = fakeApi([]);
  api.GetSheet('S').worksheet.dataValidations = { elems: 'not-an-array' };
  const bridge = new Function('Api', `${SOURCE}; return Api.GetEuroOfficeBridge();`)(api);
  assert.deepEqual(bridge.read('range.validations.read/v1', { sheet: 'S', range: 'A1' }), { ok: false, reason: 'dataValidations-elems-not-a-list' });
});

test('FAIL CLOSED: an individual validation entry with an unexpected shape refuses the whole read, not a partial list', () => {
  assert.deepEqual(read([{ type: 6, ranges: 'not-an-array', formula1: null, formula2: null }], { sheet: 'S', range: 'A1:Z100' }),
    { ok: false, reason: 'dataValidations-elems-not-a-list' });
  assert.deepEqual(read([{ type: 6, ranges: [{ r1: 0, c1: 0, r2: 'x', c2: 0 }], formula1: null, formula2: null }], { sheet: 'S', range: 'A1:Z100' }),
    { ok: false, reason: 'dataValidations-range-unexpected-shape' });
  assert.deepEqual(read([{ type: 6, ranges: [{ r1: 0, c1: 0, r2: 0, c2: 0 }], formula1: {}, formula2: null }], { sheet: 'S', range: 'A1' }),
    { ok: false, reason: 'validation-formula-unexpected-shape' });
});

test('FAIL CLOSED: range-bounds-unavailable when GetRow/GetCol/GetRowsCount/GetColumnsCount do not give integers', () => {
  const api = fakeApi([]);
  api.GetSheet('S').GetRange = () => ({ GetRow: () => undefined, GetCol: () => 1, GetRowsCount: () => 1, GetColumnsCount: () => 1 });
  const bridge = new Function('Api', `${SOURCE}; return Api.GetEuroOfficeBridge();`)(api);
  assert.deepEqual(bridge.read('range.validations.read/v1', { sheet: 'S', range: 'A1' }), { ok: false, reason: 'range-bounds-unavailable' });
});

test('capability is advertised in the manifest exactly once, at version 1, kind READ, mutation NO', () => {
  const bridge = bridgeFor([]);
  const m = bridge.manifest();
  assert.equal(m.patchVersion, 'euro-office-bridge-v1/9.3.4.60/5');
  const caps = m.capabilities.filter(c => c.id === 'range.validations.read/v1');
  assert.equal(caps.length, 1);
  assert.deepEqual([caps[0].version, caps[0].kind, caps[0].mutation, caps[0].format], [1, 'READ', 'NO', 'XLSX']);
});

test('NEGATIVE CONTROL: an unadvertised capability id is refused the same way as a real bridge would', () => {
  assert.deepEqual(bridgeFor([]).read('range.validations.doesNotExist/v1', { sheet: 'S', range: 'A1' }), { ok: false, reason: 'capability-not-advertised' });
});

test('the OTHER ten capabilities are unchanged by this patch (regression, mirrors the family\'s own getter-inventory invariant); the manifest now also carries workbook.definedNames.read/v1 (: ONE combined /2->/3 operation)', () => {
  const m = bridgeFor([]).manifest();
  const ids = m.capabilities.map(c => c.id);
  assert.equal(ids.length, 16, 'the /4 revision adds three Wave B READ capabilities (range.alignmentExt / sheet.printSettings / pivot.state); /5 adds the B15 RUNTIME_HYDRATION entry (sheet.model.ensureReady/v1)');
  for (const id of ['textStyle.bold.read/v1', 'textStyle.italic.read/v1', 'range.appearance.read/v1', 'sheet.autoFilter.fieldIdentity.read/v1',
    'sheet.columnBreaks.read/v1', 'chart.dataLabels.effective.read/v1', 'sheet.freezeState.read/v1', 'drawing.anchor.read/v1',
    'drawing.text.read/v1', 'sheet.columnWidth.canonicalize/v1', 'workbook.definedNames.read/v1']) {
    assert.ok(ids.includes(id), `missing capability: ${id}`);
  }
});
