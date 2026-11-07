'use strict';

/**
 * internal tracker, reopened: the first fix (ba4b0d7b) built its before/after
 * fixtures with `autoFilterCriteria` directly on the RAW table object -- a shape that never
 * occurs in production. `gate()`/`restOf()` in this file work on `inv.tableReads[sheet].tables`
 * (the driver's own raw read, same shape inventory.cjs forwards); on that shape a table's own
 * AutoFilter detail lives under `autoFilterDetail` (`{ present, range, fields }` when readable --
 * see xlsx-driver.cjs's bridgeAxis(TABLE_AUTOFILTER_CRITERIA_READ) attach and tables-read.cjs's
 * autoFilterCriteriaOf(), which exists ONLY to turn this raw field into the `autoFilterCriteria`
 * the read_tables projection exposes to callers -- the raw object gate() sees never carries
 * `autoFilterCriteria`). The old fixtures' field name meant normAutoFilterRange() never found
 * anything to null, so the LIVE /28 false-bukas (RESIZE_UP/RESIZE_DOWN/TOTALS all BLOCKED on an
 * otherwise-correct write) reproduced byte-identically after that "fix" landed (the production host, 2026-11-06
 * 21:34, same failed-check list as the original reg28/reg28b runs).
 *
 * This suite replaces the retired one with the REAL raw shape (`autoFilterDetail`), the only
 * shape the code under test ever actually receives.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const fam = require('../../src/adapters/xlsx/families/table-props.cjs');

const CRITERIA_NOT_READ = 'UNOBSERVABLE:table-autofilter-criteria-not-read';

function obs(tables) {
  return { primary: { format: 'xlsx', sheets: [{ name: 'A' }], tableReads: { A: { tables } } } };
}

const baseTable = extra => ({
  name: 'Table1', headerRow: true, totalsRow: false,
  columns: [{ name: 'Col1' }, { name: 'Col2' }, { name: 'Col3' }],
  style: { name: 'TableStyleLight1', rowStripes: true, columnStripes: false, firstColumn: false, lastColumn: false },
  autoFilter: true,
  ...extra,
});

const FIELDS_V1 = [{ colId: 0, values: ['x'] }];
const FIELDS_V2 = [{ colId: 0, values: ['x', 'y'] }]; // a genuinely different filter condition

test('resize_table, /28-shaped autoFilterDetail (the REAL raw field): range moves WITH ref, fields UNCHANGED -> ok (the LIVE-measured false-bukas is fixed)', () => {
  const op = { intent: 'resize_table', selector: { sheet: 'A', table: 'Table1' }, want: { ref: 'A1:C7' } };
  const before = obs([baseTable({ ref: 'A1:C5', autoFilterDetail: { present: true, range: 'A1:C5', fields: FIELDS_V1 } })]);
  const after = obs([baseTable({ ref: 'A1:C7', autoFilterDetail: { present: true, range: 'A1:C7', fields: FIELDS_V1 } })]);
  const r = fam.verifyPrimitive(op, before, after);
  assert.equal(r.ok, true, JSON.stringify(r.evidence));
  assert.equal(r.evidence.restOfTableKept, true);
});

test('NEGATIVE CONTROL: resize_table where the fields ALSO changed -> restOfTableKept must still catch it (fields are never excluded by the fix)', () => {
  const op = { intent: 'resize_table', selector: { sheet: 'A', table: 'Table1' }, want: { ref: 'A1:C7' } };
  const before = obs([baseTable({ ref: 'A1:C5', autoFilterDetail: { present: true, range: 'A1:C5', fields: FIELDS_V1 } })]);
  const after = obs([baseTable({ ref: 'A1:C7', autoFilterDetail: { present: true, range: 'A1:C7', fields: FIELDS_V2 } })]);
  const r = fam.verifyPrimitive(op, before, after);
  assert.equal(r.ok, false, JSON.stringify(r.evidence));
  assert.equal(r.evidence.restOfTableKept, false);
});

test('NEGATIVE CONTROL 2: a present:true detail whose OTHER field (not range/fields) changed is still caught -- the fix nulls range ONLY', () => {
  const op = { intent: 'resize_table', selector: { sheet: 'A', table: 'Table1' }, want: { ref: 'A1:C7' } };
  const before = obs([baseTable({ ref: 'A1:C5', autoFilterDetail: { present: true, range: 'A1:C5', fields: FIELDS_V1, reason: null } })]);
  const after = obs([baseTable({ ref: 'A1:C7', autoFilterDetail: { present: true, range: 'A1:C7', fields: FIELDS_V1, reason: 'something-changed' } })]);
  const r = fam.verifyPrimitive(op, before, after);
  assert.equal(r.ok, false, JSON.stringify(r.evidence));
});

test('set_table_totals_row, /28-shaped autoFilterDetail: range moves with the totals-row-extended ref, fields UNCHANGED -> ok', () => {
  const op = { intent: 'set_table_totals_row', selector: { sheet: 'A', table: 'Table1' }, want: { value: true } };
  const before = obs([baseTable({ ref: 'A1:C5', totalsRow: false, autoFilterDetail: { present: true, range: 'A1:C5', fields: FIELDS_V1 } })]);
  const after = obs([baseTable({ ref: 'A1:C6', totalsRow: true, autoFilterDetail: { present: true, range: 'A1:C6', fields: FIELDS_V1 } })]);
  const r = fam.verifyPrimitive(op, before, after);
  assert.equal(r.ok, true, JSON.stringify(r.evidence));
  assert.equal(r.evidence.restOfTableKept, true);
});

test('NEGATIVE CONTROL: set_table_totals_row where the fields ALSO changed -> restOfTableKept must still catch it', () => {
  const op = { intent: 'set_table_totals_row', selector: { sheet: 'A', table: 'Table1' }, want: { value: true } };
  const before = obs([baseTable({ ref: 'A1:C5', totalsRow: false, autoFilterDetail: { present: true, range: 'A1:C5', fields: FIELDS_V1 } })]);
  const after = obs([baseTable({ ref: 'A1:C6', totalsRow: true, autoFilterDetail: { present: true, range: 'A1:C6', fields: FIELDS_V2 } })]);
  const r = fam.verifyPrimitive(op, before, after);
  assert.equal(r.ok, false, JSON.stringify(r.evidence));
});

test('rename_table and set_table_style are UNTOUCHED by this fix: a table carrying /28-shaped autoFilterDetail whose range does NOT move still compares restOf correctly', () => {
  // these intents never touch `ref`, so autoFilterDetail.range is not expected to move at all --
  // restOf() does not call normAutoFilterRange() for them, and the comparison must still pass when
  // nothing (including autoFilterDetail) actually changed.
  const detail = { present: true, range: 'A1:C5', fields: FIELDS_V1 };
  for (const [op, mutate] of [
    [{ intent: 'rename_table', selector: { sheet: 'A', table: 'Table1' }, want: { name: 'Renamed' } }, t => ({ ...t, name: 'Renamed' })],
    [{ intent: 'set_table_style', selector: { sheet: 'A', table: 'Table1' }, want: { style: 'TableStyleMedium2' } }, t => ({ ...t, style: { ...t.style, name: 'TableStyleMedium2' } })],
  ]) {
    const before = obs([baseTable({ ref: 'A1:C5', autoFilterDetail: detail })]);
    const afterTable = mutate(baseTable({ ref: 'A1:C5', autoFilterDetail: detail }));
    const after = obs([afterTable]);
    const r = fam.verifyPrimitive(op, before, after);
    assert.equal(r.ok, true, `${op.intent}: ${JSON.stringify(r.evidence)}`);
  }
});

test('/27-shaped (string) autoFilterDetail: behaviour stays byte-identical to before this fix -- resize still ok when only ref moves', () => {
  const op = { intent: 'resize_table', selector: { sheet: 'A', table: 'Table1' }, want: { ref: 'A1:C7' } };
  const before = obs([baseTable({ ref: 'A1:C5', autoFilterDetail: CRITERIA_NOT_READ })]);
  const after = obs([baseTable({ ref: 'A1:C7', autoFilterDetail: CRITERIA_NOT_READ })]);
  const r = fam.verifyPrimitive(op, before, after);
  assert.equal(r.ok, true, JSON.stringify(r.evidence));
});

test('a table with autoFilter: false (no autoFilterDetail field at all) is unaffected -- normAutoFilterRange() is a no-op when the field is absent', () => {
  const op = { intent: 'resize_table', selector: { sheet: 'A', table: 'Table1' }, want: { ref: 'A1:C7' } };
  const before = obs([baseTable({ ref: 'A1:C5', autoFilter: false })]);
  const after = obs([baseTable({ ref: 'A1:C7', autoFilter: false })]);
  const r = fam.verifyPrimitive(op, before, after);
  assert.equal(r.ok, true, JSON.stringify(r.evidence));
});
