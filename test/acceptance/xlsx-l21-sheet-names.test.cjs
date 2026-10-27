'use strict';

/**
 * L-21/2 acceptance (offline): the released sheet-name class of the name-only / rename intents. ONE rule
 * (src/adapters/xlsx/sheet-name.cjs) now answers at all three layers that used to carry their own copy:
 * the write-tools schema + CONSTRAINTS, the sheets family (create / rename) and the charts family (non-creation intents).
 * Evidence: docs/L-21-SHEET-NAME-DESIGN-20261015.md + LIVE r2 (a maintainer, internal trackera965 #ref): space and accent names work
 * end to end; an apostrophe name does not (#NAME? after a reopen, rename not kept, chart series not followed).
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const W = require('../../src/server/write-tools.cjs');
const { xlsxAdapter } = require('../../src/adapters/xlsx/adapter.cjs');
const { sheetNameProblem, APOSTROPHE_REASON } = require('../../src/adapters/xlsx/sheet-name.cjs');
const { VALID, req } = require('../fakes/write-refusals.cjs');

/** every released intent whose `sheet` (or, for create_sheet / rename_sheet, `name`) is a sheet name, with the field(s) */
const SHEET_FIELDS = Object.freeze({
  insert_rows: ['sheet'],
  delete_rows: ['sheet'],
  rename_sheet: ['sheet', 'name'],
  freeze_panes: ['sheet'],
  unfreeze_panes: ['sheet'],
  set_chart_title: ['sheet'],
  rename_chart: ['sheet'],
  set_chart_axis_title: ['sheet'],
  set_chart_style: ['sheet'],
  create_sheet: ['name'],
});
const RELEASED = ['Munkalap 1', 'Kimutatás', 'Bevétel 2026', 'Q 1', 'Budget_2026', 'Ünnep', 'Sheet1', '_x'];
const APOSTROPHE = ["O'Brien", "D'Arcy", "'lead", "trail'", "it's here"];
/** not measured: punctuation, a leading digit, padded / doubled spaces -- refused with their own reason, not the apostrophe one */
const NOT_MEASURED = ['2026', 'Bevétel-2026', 'Összesítő (új)', 'a.b', ' lead', 'trail ', 'two  spaces'];

// rename_sheet: the new name must differ from the old one -> the field NOT under test is pinned to a third, released name
const step = (intent, field, value) => ({ ...VALID[intent], ...(intent === 'rename_sheet' ? { sheet: 'Alt_1', name: 'Alt_2' } : {}), [field]: value });
const validate = (intent, field, value) => W.validateWriteRequest(req([step(intent, field, value)]));

test('the shared rule: released names pass, the apostrophe has its OWN named reason, the rest are class-not-measured', () => {
  for (const n of RELEASED) assert.equal(sheetNameProblem(n), null, n);
  for (const n of APOSTROPHE) assert.equal(sheetNameProblem(n), APOSTROPHE_REASON, n);
  for (const n of NOT_MEASURED) assert.match(sheetNameProblem(n), /^sheet-name-class-not-measured/, JSON.stringify(n));
  for (const n of ['A1', 'XFD7', 'R1C1', 'rc', 'History', 'TRUE']) assert.notEqual(sheetNameProblem(n), null, n);
  assert.match(sheetNameProblem('L'.padEnd(32, 'x')), /longer than 31/);
  assert.equal(sheetNameProblem('L'.padEnd(31, 'x')), null);
  assert.notEqual(sheetNameProblem(''), null);
  assert.notEqual(sheetNameProblem(undefined), null);
});

test('PRODUCT SURFACE (write-tools): every released intent accepts the released names and refuses an apostrophe BY NAME with 0 further layers reached', () => {
  for (const [intent, fields] of Object.entries(SHEET_FIELDS)) {
    for (const field of fields) {
      for (const n of RELEASED.filter(x => !/^(A1|XFD7)$/.test(x))) {
        const v = validate(intent, field, n);
        assert.equal(v.ok, true, `${intent}.${field} ${JSON.stringify(n)}: ${v.reason}`);
      }
      for (const n of APOSTROPHE) {
        const v = validate(intent, field, n);
        assert.equal(v.ok, false, `${intent}.${field} ${JSON.stringify(n)}`);
        assert.equal(v.code, 'REFUSED_SHAPE', `${intent}.${field}`);
        assert.match(v.reason, /sheet-name-apostrophe-unsupported/, `${intent}.${field} ${JSON.stringify(n)}`);
      }
      for (const n of NOT_MEASURED) {
        const v = validate(intent, field, n);
        assert.equal(v.ok, false, `${intent}.${field} ${JSON.stringify(n)}`);
        assert.equal(v.code, 'REFUSED_SHAPE', `${intent}.${field}`);
      }
    }
  }
});

test('PRODUCT SURFACE: a released name reaches the executor task verbatim (no quoting, no rewriting at the product layer)', () => {
  const v = validate('freeze_panes', 'sheet', 'Munkalap 1');
  assert.equal(v.ok, true);
  assert.equal(v.task.steps[0].sheet, 'Munkalap 1');
  const r = validate('rename_sheet', 'name', 'Kimutatás');
  assert.equal(r.task.steps[0].name, 'Kimutatás');
});

test('FAMILY LAYER (planner alone, no product schema): the same rule at the sheets family and the charts family', () => {
  const plan = s => xlsxAdapter.planOperations({ id: 't', steps: [s] });
  for (const n of ['Munkalap 1', 'Kimutatás']) {
    assert.equal(plan({ intent: 'create_sheet', name: n }).ok, true, `create ${n}`);
    assert.equal(plan({ intent: 'rename_sheet', sheet: 'Old', name: n }).ok, true, `rename ${n}`);
    assert.equal(plan({ intent: 'set_chart_title', sheet: n, name: 'C1', title: 'T' }).ok, true, `chart title ${n}`);
    assert.equal(plan({ intent: 'set_chart_axis_title', sheet: n, name: 'C1', axis: 'horizontal', title: 'T' }).ok, true, `axis ${n}`);
    assert.equal(plan({ intent: 'rename_chart', sheet: n, name: 'C1', newName: 'C2' }).ok, true, `rename chart ${n}`);
  }
  for (const n of APOSTROPHE) {
    for (const s of [{ intent: 'create_sheet', name: n }, { intent: 'rename_sheet', sheet: 'Old', name: n }, { intent: 'rename_sheet', sheet: n, name: 'New' },
      { intent: 'set_chart_title', sheet: n, name: 'C1', title: 'T' }, { intent: 'rename_chart', sheet: n, name: 'C1', newName: 'C2' }]) {
      const p = plan(s);
      assert.equal(p.ok, false, `${s.intent} ${JSON.stringify(n)}`);
      assert.match(p.detail, /sheet-name-apostrophe-unsupported/, `${s.intent} ${JSON.stringify(n)}`);
    }
  }
});

test('set_chart (creation) is NOT widened: it builds the series reference text unquoted, so it keeps the bare class', () => {
  const add = sheet => xlsxAdapter.planOperations({ id: 't', steps: [{ intent: 'set_chart', sheet, name: 'C1', range: 'A1:C4', type: 'bar', title: 'T' }] });
  assert.equal(add('Budget').ok, true);
  for (const n of ['Munkalap 1', 'Kimutatás', "O'Brien"]) {
    const p = add(n);
    assert.equal(p.ok, false, n);
    assert.match(p.detail, /needs no quoting/, n);
  }
});

test('clear_print_area keeps its (SHEET_ADDRESS) class: it is not a name-only / rename intent of this slice', () => {
  assert.equal(validate('clear_print_area', 'sheet', 'Munkalap 1').ok, true);
  assert.equal(validate('clear_print_area', 'sheet', 'Kimutatás').ok, true);
  assert.equal(validate('clear_print_area', 'sheet', 'A1').ok, false);
});

test('REGRESSION: the old bare names are still released and cell-reference / reserved names are still refused on every intent', () => {
  for (const [intent, fields] of Object.entries(SHEET_FIELDS)) {
    for (const field of fields) {
      assert.equal(validate(intent, field, 'Budget_2026').ok, true, `${intent}.${field}`);
      for (const bad of ['A1', 'History']) assert.equal(validate(intent, field, bad).ok, false, `${intent}.${field} ${bad}`);
    }
  }
});
