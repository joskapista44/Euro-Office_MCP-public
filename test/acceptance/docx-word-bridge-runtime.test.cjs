'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { propertyQuery } = require('../../src/drivers/eurooffice/docx-driver.cjs');
const { BRIDGE_CAPABILITIES } = require('../../src/drivers/eurooffice/ds-bridge.cjs');

/**
 * W-B1920-C1 word-bridge cutover (+ the three mandatory conditions from #ref, the owner's
 * cutover GO 2026-10-07 14:4x): the "fake bridge-runtime" condition (2) asks for -- NOT
 * fake-docx-driver.cjs's higher-level transport fake (which hand-builds the `numbering`/`gridCalc`
 * shapes directly and never executes this driver's generated query text at all), but the SAME
 * sandboxed-Api pattern test/static/ds-bridge-client.test.cjs already uses for the generic client:
 * evaluate propertyQuery(build)'s actual GENERATED STRING with `new Function('Api', query)`, against
 * a scriptable fake `Api.GetEuroOfficeBridge()`, to prove the DRIVER'S OWN wiring (not just the
 * shared __eoBridgeRead gates) calls the right capability id/version/args for a paragraph and a
 * table, that FAIL-CLOSED holds for capability MISSING / WRONG VERSION / MALFORMED result on BOTH
 * capabilities, and -- condition (1), fallback forbidden even in hidden/debug form -- that NEITHER
 * ever reaches an internal GetNumbering()/TableGridCalc read EVEN WHEN the fake document exposes
 * them (a spy proves zero calls), with no mutation.
 *
 * The fake document surface is the minimum the query touches before it reaches a bridge call: every
 * document-level read past it (bookmarks, footnotes, sections, ...) is individually try/caught in
 * the driver and safely no-ops when the method is absent.
 */
const BUILD = '9.3.4-9475e9f12a1905e2ada1fbc4bb87d01b';
const LIST_CAP = BRIDGE_CAPABILITIES.DOCX_LIST_LEVEL_FORMAT_READ;
const GRID_CAP = BRIDGE_CAPABILITIES.DOCX_TABLE_GRID_READ;
// DS9346 (owner D5=B, VALTAS, docs/DS9346-RUNTIME-RETIREMENT.md): the default manifest now carries the
// ACCEPTED 9.3.5.0 identity -- the historical word /1 (9.3.4.60) baseline is retired; since the GH #292 S4 D-state the
// single accepted Word runtime is 9.3.5.0/5.
const manifest = (over = {}) => ({
  bridgeVersion: '1.0.0', documentServerVersion: '9.3.5', documentServerBuild: BUILD,
  patchVersion: 'euro-office-bridge-word-v1/9.3.5.0/5',
  capabilities: [
    { id: LIST_CAP.id, version: 1, kind: 'READ', format: 'DOCX', mutation: 'NO', semanticContract: 'x' },
    { id: GRID_CAP.id, version: 1, kind: 'READ', format: 'DOCX', mutation: 'NO', semanticContract: 'x' },
  ],
  ...over,
});

/**
 * One paragraph (docIndex 0), one table (index 0, zero rows -- the row loop is not what's under
 * test here). `internalSpy` exposes the old internal surface (GetNumbering / Table.TableGridCalc)
 * SO THAT a hidden fallback would have something to call -- and the spy counts whether it did.
 */
function fakeDoc(internalSpy) {
  const paragraph = {
    GetClassType: () => 'paragraph', GetText: () => 'ZQ bridge runtime probe', GetElementsCount: () => 0, GetElement: () => null,
    GetNumbering: () => { internalSpy.getNumbering += 1; return { present: true, level: 0 }; },
  };
  const table = { GetClassType: () => 'table', GetRowsCount: () => 0, get Table() { internalSpy.tableGridCalc += 1; return { TableGridCalc: [1, 2, 3] }; } };
  return { GetElementsCount: () => 2, GetElement: i => (i === 0 ? paragraph : (i === 1 ? table : null)), GetAllTables: () => [table], Search: () => [] };
}

function fakeApi(read, spy = { getNumbering: 0, tableGridCalc: 0 }) {
  return { GetDocument: () => fakeDoc(spy), GetEuroOfficeBridge: () => ({ manifest: () => manifest(), read }), __spy: spy };
}

function evaluate(Api) {
  return new Function('Api', propertyQuery())(Api);
}

/** falsy (null, omitted-key -> undefined) -- the same "unread" signal inventory.cjs's own `!x` check already treats identically. */
const unread = x => !x;

test('the generated query calls docx.numbering.levelFormat.read/v1 with the paragraph INDEX, and uses the bridge value as-is (no internal call)', () => {
  const seen = [];
  const Api = fakeApi((id, args) => { seen.push([id, args]); return { ok: true, value: { present: true, level: 2, numId: '7', levelFormats: [5, 5, 5, 5, 5, 5, 5, 5, 5] } }; });
  const r = evaluate(Api);
  assert.equal(r.error, undefined, JSON.stringify(r));
  assert.deepEqual(seen[0], [LIST_CAP.id, { paragraph: 0 }]);
  assert.deepEqual(r.paragraphs[0].numbering, { present: true, level: 2, numId: '7', levelFormats: [5, 5, 5, 5, 5, 5, 5, 5, 5] });
  assert.equal(Api.__spy.getNumbering, 0, 'the bridge value is used as-is -- GetNumbering() is never called');
});

test('the generated query calls docx.table.effectiveGrid.read/v1 with the table INDEX, and uses gridCalc as-is (no internal call)', () => {
  const seen = [];
  const Api = fakeApi((id, args) => { seen.push([id, args]); return { ok: true, value: { gridCalc: [14.586, 14.586, 18.9] } }; });
  const r = evaluate(Api);
  assert.equal(r.error, undefined, JSON.stringify(r));
  assert.deepEqual(seen[1], [GRID_CAP.id, { table: 0 }]);
  assert.deepEqual(r.document.tables[0].gridCalc, [14.586, 14.586, 18.9]);
  assert.equal(Api.__spy.tableGridCalc, 0, 'the bridge value is used as-is -- Table.TableGridCalc is never read');
});

/** Condition (2): FAIL-CLOSED for capability MISSING / WRONG VERSION / MALFORMED result, on BOTH capabilities. */
const FAIL_CASES = [
  ['capability MISSING (not advertised in the manifest)', m => manifest({ capabilities: m.capabilities.filter(c => c.id !== LIST_CAP.id) }), () => ({ ok: true, value: { present: false } }), 'list'],
  ['capability MISSING (not advertised in the manifest)', m => manifest({ capabilities: m.capabilities.filter(c => c.id !== GRID_CAP.id) }), () => ({ ok: true, value: { gridCalc: [1] } }), 'grid'],
  ['WRONG VERSION (bridge major version mismatch)', () => manifest({ bridgeVersion: '2.0.0' }), () => ({ ok: true, value: { present: false } }), 'list'],
  ['WRONG VERSION (bridge major version mismatch)', () => manifest({ bridgeVersion: '2.0.0' }), () => ({ ok: true, value: { gridCalc: [1] } }), 'grid'],
  ['WRONG VERSION (capability version mismatch)', m => manifest({ capabilities: m.capabilities.map(c => (c.id === LIST_CAP.id ? { ...c, version: 2 } : c)) }), () => ({ ok: true, value: { present: false } }), 'list'],
  ['WRONG VERSION (capability version mismatch)', m => manifest({ capabilities: m.capabilities.map(c => (c.id === GRID_CAP.id ? { ...c, version: 2 } : c)) }), () => ({ ok: true, value: { gridCalc: [1] } }), 'grid'],
  ['MALFORMED result (ok:true but the wrong valueType)', () => manifest(), () => ({ ok: true, value: 'not-an-object' }), 'list'],
  ['MALFORMED result (ok:true but the wrong valueType)', () => manifest(), () => ({ ok: true, value: 'not-an-object' }), 'grid'],
  ['MALFORMED result (an unexpected shape, neither ok:true+value nor ok:false+reason)', () => manifest(), () => ({ ok: true, oops: 1 }), 'list'],
  ['MALFORMED result (an unexpected shape, neither ok:true+value nor ok:false+reason)', () => manifest(), () => ({ ok: true, oops: 1 }), 'grid'],
];

for (const [label, buildManifest, read, axis] of FAIL_CASES) {
  test(`FAIL CLOSED (${axis}): ${label} -> unread, no internal call, no mutation`, () => {
    const spy = { getNumbering: 0, tableGridCalc: 0 };
    const Api = { GetDocument: () => fakeDoc(spy), GetEuroOfficeBridge: () => ({ manifest: () => buildManifest(manifest()), read }) };
    const r = evaluate(Api);
    assert.equal(r.error, undefined, JSON.stringify(r));
    assert.ok(unread(r.paragraphs[0].numbering), `numbering should be unread: ${JSON.stringify(r.paragraphs[0].numbering)}`);
    assert.ok(unread(r.document.tables[0].gridCalc), `gridCalc should be unread: ${JSON.stringify(r.document.tables[0].gridCalc)}`);
    assert.equal(spy.getNumbering, 0, 'condition (1): no hidden fallback to GetNumbering()');
    assert.equal(spy.tableGridCalc, 0, 'condition (1): no hidden fallback to Table.TableGridCalc');
  });
}

test('FAIL CLOSED: BRIDGE_UNAVAILABLE (no Api.GetEuroOfficeBridge at all) leaves numbering/gridCalc unread, no internal call', () => {
  const spy = { getNumbering: 0, tableGridCalc: 0 };
  const r = evaluate({ GetDocument: () => fakeDoc(spy) });
  assert.equal(r.error, undefined, JSON.stringify(r));
  assert.ok(unread(r.paragraphs[0].numbering));
  assert.ok(unread(r.document.tables[0].gridCalc));
  assert.equal(spy.getNumbering, 0);
  assert.equal(spy.tableGridCalc, 0);
});

for (const [label, over] of [['RUNTIME_VERSION_MISMATCH (another DS runtime version)', { documentServerVersion: '9.3.4' }],
  ['PATCH_VERSION_MISMATCH (another Word bridge revision)', { patchVersion: 'euro-office-bridge-word-v1/9.3.4.60/4' }],
  ['PATCH_VERSION_MISMATCH (the CELL bridge manifest)', { patchVersion: 'euro-office-bridge-v1/9.3.4.60/3' }]]) {
  test(`FAIL CLOSED: ${label} -> unread, not the bridge value, no internal call`, () => {
    const spy = { getNumbering: 0, tableGridCalc: 0 };
    const Api = { GetDocument: () => fakeDoc(spy), GetEuroOfficeBridge: () => ({ manifest: () => manifest(over), read: () => ({ ok: true, value: { present: false } }) }) };
    const r = evaluate(Api);
    assert.equal(r.error, undefined, JSON.stringify(r));
    assert.ok(unread(r.paragraphs[0].numbering));
    assert.equal(spy.getNumbering, 0);
  });
}

test('CACHE-KEY CHANGE: the manifest\'s documentServerBuild (a patch-time cache key) differs -- the Word reads still PASS (not a gate)', () => {
  const spy = { getNumbering: 0, tableGridCalc: 0 };
  const value = { present: true, level: 2, numId: '7', levelFormats: [5, 5, 5, 5, 5, 5, 5, 5, 5] };
  const Api = { GetDocument: () => fakeDoc(spy), GetEuroOfficeBridge: () => ({ manifest: () => manifest({ documentServerBuild: '9.3.4-94bc577ecccd9583be1b130bd5a6b407' }), read: id => (id === LIST_CAP.id ? { ok: true, value } : { ok: true, value: { gridCalc: [1] } }) }) };
  const r = evaluate(Api);
  assert.equal(r.error, undefined, JSON.stringify(r));
  assert.equal(unread(r.paragraphs[0].numbering), false, 'the reads are not refused by a cache-key difference');
  assert.equal(spy.getNumbering, 0);
});

test('POSITIVE CONTROL for the spy itself: a fake that is NOT exercised through the bridge path still lets the spy count a direct call (proves the spy is not vacuously zero)', () => {
  const spy = { getNumbering: 0, tableGridCalc: 0 };
  const d = fakeDoc(spy);
  d.GetElement(0).GetNumbering();
  d.GetAllTables()[0].Table;
  assert.deepEqual(spy, { getNumbering: 1, tableGridCalc: 1 });
});
