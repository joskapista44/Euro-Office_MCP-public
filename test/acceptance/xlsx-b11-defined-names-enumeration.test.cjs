'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

/**
 * B11 /3 COMBINED -- workbook.definedNames.read/v1 Bridge capability, OFFLINE (card
 * an internal branch, tracker 9580767b). Added to
 * the SAME /2->/3 operation as range.validations.read/v1 (abandoning the validation-only /3,
 * 2719e511, per the owner's decision via zero).
 *
 * Reader (a maintainer's own static analysis of the served /2 bundle, s2.js -- MEASURED, not guessed):
 * `wbModel.dependencyFormulas.defNames` = `{ wb: {...}, sheet: { <sheetId>: {...} } }`, each value a
 * `DefName`-shaped object `{ name, ref, sheetId, hidden, type, isXLNM }`. The capability reads the
 * STORED `ref` field directly -- NEVER `getRef(true)` (production's own lazy-parse/cache path) and
 * NEVER `getAscCDefName()` (which mutates the internal `type` field as a side effect of building a
 * public wrapper, a maintainer's own finding) -- so this capability is structurally non-creating, proven
 * below with a spy on both methods.
 *
 * Args: {} (stateless full inventory, no active-sheet dependency). Value: every defined name,
 * workbook-scope first then sheet-scope (by sheet INDEX), name-ordered within each scope. The same
 * name in two scopes is two SEPARATE entries. Hidden/built-in/table/ref-less filtering is explicitly
 * a DRIVER decision, not this capability's -- it returns the full inventory, unfiltered.
 *
 * FIXED (review of the first delivery -- a negative-control read of the served /2
 * bundle's own structure):
 *   1. `hidden` accepts null AND undefined as false -- the bundle's own table-name writer
 *      (`addDefNameOpen(table.DisplayName, ref, null, null, c_oAscDefNameType.table)`) and some
 *      file-open paths store `hidden` as `null`, not a boolean. The OLD code only accepted
 *      `undefined`, so ANY workbook containing one table lost its ENTIRE names inventory to
 *      `name-unexpected-shape` -- a false NEGATIVE on the whole capability, not just that one entry.
 *   2. New `table` field: true iff the stored `type` equals the table-name enum value
 *      (`Asc.c_oAscDefNameType.table`). Table names are NOT present in workbook.xml's own
 *      `<definedNames>` (they live in the table definition instead), so a FILE-level parity check
 *      must be able to exclude them deterministically -- this flag is how. If the enum itself is
 *      unavailable, the WHOLE read fails closed (`defname-type-enum-missing`), never silently
 *      reporting `table: false` for an entry that might actually be one.
 *   3. The model resolution no longer tries `Api.GetDocument().wbModel` at all -- in the served
 *      bundle, `Api.GetDocument = new ApiDocument(private_GetLogicDocument())` is a wrapper
 *      construction on a path that never returns `wbModel` anyway, so the call was pure waste (and
 *      risk) on every read. Only `Asc.editor.wbModel` is used now -- the SAME path the public
 *      `Api.GetDefName` itself resolves through. Proven non-creating with a SECOND spy: a fake
 *      `Api.GetDocument` that THROWS if ever invoked (every test implicitly asserts this; one test
 *      asserts it explicitly with its own positive control).
 */
const BLOCK_PATH = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.4.60', 'euro-office-bridge-v1', 'bridge-block.js');
const SOURCE = fs.readFileSync(BLOCK_PATH, 'utf8').replace(/,\s*$/, '');
const CAP_ID = 'workbook.definedNames.read/v1';
const TABLE_TYPE = 1; // an arbitrary sentinel value for the fake enum -- its own identity is irrelevant, only equality matters

function fakeApi(spy = {}) {
  return { GetDocument: () => { spy.getDocumentCalls = (spy.getDocumentCalls || 0) + 1; throw new Error('GetDocument must never be called by workbook.definedNames.read/v1'); } };
}

function fakeAsc(defNames, sheets, { tableType = TABLE_TYPE, noEnum = false, noTableKey = false } = {}) {
  const wsById = {};
  sheets.forEach((nm, idx) => { wsById[idx] = { getName: () => nm, getIndex: () => idx }; });
  const wb = { dependencyFormulas: { defNames }, getWorksheetById: id => wsById[id] };
  const asc = { editor: { wbModel: wb } };
  if (!noEnum) asc.c_oAscDefNameType = noTableKey ? {} : { table: tableType };
  return asc;
}

function spyDefName(fields, spy) {
  return {
    ...fields,
    getRef: (...a) => { spy.getRef = (spy.getRef || 0) + 1; return fields.ref; },
    getAscCDefName: (...a) => { spy.getAscCDefName = (spy.getAscCDefName || 0) + 1; return {}; },
  };
}

function bridgeOf(apiSpy, ascOpts, defNames, sheets) {
  return new Function('Api', 'Asc', `${SOURCE}; return Api.GetEuroOfficeBridge();`)(fakeApi(apiSpy), fakeAsc(defNames, sheets, ascOpts));
}

function read(defNames, sheets, args, ascOpts) {
  const bridge = bridgeOf({}, ascOpts, defNames, sheets);
  return bridge.read(CAP_ID, args === undefined ? {} : args);
}

test('0 names: []', () => {
  assert.deepEqual(read({ wb: {}, sheet: {} }, []), { ok: true, value: [] });
});

test('workbook-scope entry', () => {
  const r = read({ wb: { a: { name: 'MyName', ref: 'Sheet1!$A$1', hidden: false } }, sheet: {} }, ['Sheet1']);
  assert.deepEqual(r, { ok: true, value: [{ name: 'MyName', scope: { kind: 'workbook' }, refersTo: 'Sheet1!$A$1', refError: false, hidden: false, builtIn: false, table: false }] });
});

test('sheet-scope entry, sheet identified by NAME via getWorksheetById', () => {
  const r = read({ wb: {}, sheet: { 7: { x: { name: 'Local', ref: 'A1' } } } }, [null, null, null, null, null, null, null, 'TheSheet']);
  assert.deepEqual(r, { ok: true, value: [{ name: 'Local', scope: { kind: 'sheet', sheet: 'TheSheet', index: 7 }, refersTo: 'A1', refError: false, hidden: false, builtIn: false, table: false }] });
});

test('multiple sheets: entries sorted by sheet INDEX, name within each sheet', () => {
  const r = read({ wb: {}, sheet: {
    1: { z: { name: 'Zeta', ref: 'B1' }, a: { name: 'Alpha', ref: 'B2' } },
    0: { x: { name: 'First', ref: 'A1' } },
  } }, ['Sheet0', 'Sheet1']);
  assert.equal(r.ok, true);
  assert.deepEqual(r.value.map(e => [e.scope.index, e.name]), [[0, 'First'], [1, 'Alpha'], [1, 'Zeta']]);
});

test('SAME name in workbook-scope AND sheet-scope: two SEPARATE entries, workbook first', () => {
  const r = read({ wb: { a: { name: 'Dup', ref: 'X1' } }, sheet: { 0: { b: { name: 'Dup', ref: 'Y2' } } } }, ['OnlySheet']);
  assert.equal(r.ok, true);
  assert.equal(r.value.length, 2);
  assert.deepEqual(r.value[0], { name: 'Dup', scope: { kind: 'workbook' }, refersTo: 'X1', refError: false, hidden: false, builtIn: false, table: false });
  assert.deepEqual(r.value[1], { name: 'Dup', scope: { kind: 'sheet', sheet: 'OnlySheet', index: 0 }, refersTo: 'Y2', refError: false, hidden: false, builtIn: false, table: false });
});

test('absolute range ref, read verbatim (no parse attempted)', () => {
  const r = read({ wb: { a: { name: 'AbsRange', ref: 'Sheet1!$A$1:$C$10' } }, sheet: {} }, ['Sheet1']);
  assert.equal(r.value[0].refersTo, 'Sheet1!$A$1:$C$10');
});

test('#REF! in the stored ref -> refError: true, refersTo kept verbatim (not dropped, not nulled)', () => {
  const r = read({ wb: { a: { name: 'Broken', ref: '#REF!' } }, sheet: {} }, []);
  assert.deepEqual(r.value[0], { name: 'Broken', scope: { kind: 'workbook' }, refersTo: '#REF!', refError: true, hidden: false, builtIn: false, table: false });
});

test('hidden + _xlnm built-in name: both flags carried, NOT filtered out (filtering is the driver\'s job)', () => {
  const r = read({ wb: { a: { name: '_xlnm.Print_Area', ref: 'Sheet1!$A$1:$B$2', hidden: true, isXLNM: true } }, sheet: {} }, ['Sheet1']);
  assert.deepEqual(r.value[0], { name: '_xlnm.Print_Area', scope: { kind: 'workbook' }, refersTo: 'Sheet1!$A$1:$B$2', refError: false, hidden: true, builtIn: true, table: false });
});

test('builtIn via isXLNM alone, or via the name containing "_xlnm" alone -- either sets the flag', () => {
  const byFlag = read({ wb: { a: { name: 'Ordinary', ref: 'A1', isXLNM: true } }, sheet: {} }, []);
  assert.equal(byFlag.value[0].builtIn, true);
  const byName = read({ wb: { a: { name: 'has_xlnm_in_it', ref: 'A1' } }, sheet: {} }, []);
  assert.equal(byName.value[0].builtIn, true);
  const neither = read({ wb: { a: { name: 'Plain', ref: 'A1' } }, sheet: {} }, []);
  assert.equal(neither.value[0].builtIn, false);
});

test('no ref at all: refersTo null, NEVER dropped from the inventory', () => {
  const r = read({ wb: { a: { name: 'NoRef' } }, sheet: {} }, []);
  assert.deepEqual(r, { ok: true, value: [{ name: 'NoRef', scope: { kind: 'workbook' }, refersTo: null, refError: false, hidden: false, builtIn: false, table: false }] });
});

// -- fix 1: hidden accepts null AND undefined as false; anything else non-boolean is bad --

test(': hidden: null -> hidden: false (the bundle\'s own table-name writer and some file-open paths store this, not undefined)', () => {
  const r = read({ wb: { a: { name: 'HiddenNull', ref: 'A1', hidden: null } }, sheet: {} }, []);
  assert.equal(r.ok, true);
  assert.equal(r.value[0].hidden, false);
});

test(': hidden: undefined (field entirely absent) -> hidden: false', () => {
  const r = read({ wb: { a: { name: 'HiddenAbsent', ref: 'A1' } }, sheet: {} }, []);
  assert.equal(r.ok, true);
  assert.equal(r.value[0].hidden, false);
});

test('a hidden value that is NEITHER null, undefined, NOR boolean still fails the whole read closed', () => {
  assert.deepEqual(read({ wb: { a: { name: 'X', ref: 'A1', hidden: 'yes' } }, sheet: {} }, []), { ok: false, reason: 'name-unexpected-shape' });
});

// -- fix 2: the table flag + its own fail-closed enum gate --

test(': table: true iff the stored type equals the table-name enum value', () => {
  const r = read({ wb: { a: { name: 'TableName', ref: null, type: TABLE_TYPE } }, sheet: {} }, []);
  assert.equal(r.ok, true);
  assert.equal(r.value[0].table, true);
});

test(': table: false for an ordinary name (type absent) and for a DIFFERENT, non-matching type', () => {
  const absent = read({ wb: { a: { name: 'Ordinary', ref: 'A1' } }, sheet: {} }, []);
  assert.equal(absent.value[0].table, false);
  const other = read({ wb: { a: { name: 'OtherType', ref: 'A1', type: 999 } }, sheet: {} }, []);
  assert.equal(other.value[0].table, false);
});

test(': FAIL CLOSED defname-type-enum-missing when Asc.c_oAscDefNameType itself is absent', () => {
  const r = read({ wb: { a: { name: 'N', ref: 'A1' } }, sheet: {} }, [], {}, { noEnum: true });
  assert.deepEqual(r, { ok: false, reason: 'defname-type-enum-missing' });
});

test(': FAIL CLOSED defname-type-enum-missing when Asc.c_oAscDefNameType.table itself is absent', () => {
  const r = read({ wb: { a: { name: 'N', ref: 'A1' } }, sheet: {} }, [], {}, { noTableKey: true });
  assert.deepEqual(r, { ok: false, reason: 'defname-type-enum-missing' });
});

test('FAIL CLOSED: sheet-scope-unresolved (a sheetId with no matching worksheet)', () => {
  assert.deepEqual(read({ wb: {}, sheet: { 99: { x: { name: 'Orphan', ref: 'A1' } } } }, []), { ok: false, reason: 'sheet-scope-unresolved' });
});

test('FAIL CLOSED: name-unexpected-shape (name not a string, ref not string/undefined/null)', () => {
  assert.deepEqual(read({ wb: { a: { name: 123, ref: 'A1' } }, sheet: {} }, []), { ok: false, reason: 'name-unexpected-shape' });
  assert.deepEqual(read({ wb: { a: { name: 'X', ref: 42 } }, sheet: {} }, []), { ok: false, reason: 'name-unexpected-shape' });
});

test(': FAIL CLOSED: model-missing (no Asc.editor.wbModel reachable; Api is NEVER consulted for the model anymore)', () => {
  const bridge = new Function('Api', 'Asc', `${SOURCE}; return Api.GetEuroOfficeBridge();`)({}, {});
  assert.deepEqual(bridge.read(CAP_ID, {}), { ok: false, reason: 'model-missing' });
});

test('FAIL CLOSED: defNames-missing (dependencyFormulas present but no defNames)', () => {
  const bridge = new Function('Api', 'Asc', `${SOURCE}; return Api.GetEuroOfficeBridge();`)({}, { editor: { wbModel: { dependencyFormulas: {}, getWorksheetById: () => null } }, c_oAscDefNameType: { table: TABLE_TYPE } });
  assert.deepEqual(bridge.read(CAP_ID, {}), { ok: false, reason: 'defNames-missing' });
});

test('FAIL CLOSED: bad-args (any non-empty args object, or a non-object)', () => {
  assert.deepEqual(read({ wb: {}, sheet: {} }, [], { extra: 1 }), { ok: false, reason: 'bad-args' });
  assert.deepEqual(read({ wb: {}, sheet: {} }, [], 'nope'), { ok: false, reason: 'bad-args' });
});

test('NON-CREATING, proven with TWO spies: getRef()/getAscCDefName() on a DefName are NEVER called, AND Api.GetDocument() is NEVER called -- only Asc.editor.wbModel + the plain .ref/.name/.hidden/.isXLNM/.type fields are read', () => {
  const apiSpy = {};
  const defNameSpy = {};
  const sheets = [];
  const wsById = {};
  const wb = { dependencyFormulas: { defNames: { wb: { a: spyDefName({ name: 'N1', ref: 'A1', hidden: false }, defNameSpy) }, sheet: {} } }, getWorksheetById: () => null };
  const Asc = { editor: { wbModel: wb }, c_oAscDefNameType: { table: TABLE_TYPE } };
  const bridge = new Function('Api', 'Asc', `${SOURCE}; return Api.GetEuroOfficeBridge();`)(fakeApi(apiSpy), Asc);
  const r = bridge.read(CAP_ID, {});
  assert.deepEqual(r, { ok: true, value: [{ name: 'N1', scope: { kind: 'workbook' }, refersTo: 'A1', refError: false, hidden: false, builtIn: false, table: false }] });
  assert.equal(defNameSpy.getRef, undefined, 'getRef() must never be called');
  assert.equal(defNameSpy.getAscCDefName, undefined, 'getAscCDefName() must never be called');
  assert.equal(apiSpy.getDocumentCalls, undefined, 'Api.GetDocument() must never be called');
});

test('POSITIVE CONTROL for the spies themselves: calling the spied methods directly DOES increment/throw (the spies are not vacuously silent)', () => {
  const defNameSpy = {};
  const dn = spyDefName({ name: 'N1', ref: 'A1' }, defNameSpy);
  dn.getRef(); dn.getAscCDefName();
  assert.deepEqual(defNameSpy, { getRef: 1, getAscCDefName: 1 });
  const apiSpy = {};
  assert.throws(() => fakeApi(apiSpy).GetDocument(), /GetDocument must never be called/);
  assert.equal(apiSpy.getDocumentCalls, 1);
});

test('determinism: result is independent of the object-key insertion order of wb/sheet maps', () => {
  const a = read({ wb: { x: { name: 'B', ref: '1' }, y: { name: 'A', ref: '2' } }, sheet: {} }, []);
  const b = read({ wb: { y: { name: 'A', ref: '2' }, x: { name: 'B', ref: '1' } }, sheet: {} }, []);
  assert.deepEqual(a, b);
});

test('capability advertised exactly once, version 1, kind READ, mutation NO, alongside range.validations.read/v1 (the combined /3)', () => {
  const bridge = bridgeOf({}, {}, { wb: {}, sheet: {} }, []);
  const m = bridge.manifest();
  assert.equal(m.patchVersion, 'euro-office-bridge-v1/9.3.4.60/5');
  const caps = m.capabilities.filter(c => c.id === CAP_ID);
  assert.equal(caps.length, 1);
  assert.deepEqual([caps[0].version, caps[0].kind, caps[0].mutation, caps[0].format], [1, 'READ', 'NO', 'XLSX']);
  assert.ok(m.capabilities.some(c => c.id === 'range.validations.read/v1'), 'both capabilities land in the SAME /3 manifest');
});
