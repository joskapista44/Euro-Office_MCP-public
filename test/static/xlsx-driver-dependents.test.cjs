'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const drv = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { encodeWire } = require('../../src/drivers/eurooffice/wire-codec.cjs');
const { nameUses, resolveRef } = require('../../src/adapters/xlsx/dependencies.cjs');

const { installBridge, BRIDGE_RUNTIME_ID } = require('../fakes/ds-bridge-runtime.cjs');
const { HANDSHAKE_BODY } = require('../../src/core/runtime-identity.cjs');

/**
 * B11 (MIGRATED to the DS Bridge, owner GO owner channel): the dependency inventory is ONE callCommand through
 * `workbook.definedNames.read/v1` + `range.validations.read/v1`. Here the body runs against a fake `Api` carrying the
 * REAL bridge block (ds-bridge-runtime.cjs) and the model shapes the bridge reads; every gap is 'UNOBSERVABLE:<why>'
 * -- never an empty list -- and there is no direct-internal fallback.
 */
const TABLE = 3;
function model({ sheets: list, wbNames = {}, sheetNames = {}, dv = {} }) {
  const ids = list.map((n, i) => `id${i}`);
  const wsModel = list.map((n, i) => ({ getName: () => n, getIndex: () => i }));
  const asc = {
    editor: { wbModel: { getWorksheetById: id => wsModel[ids.indexOf(id)] || null, dependencyFormulas: { defNames: { wb: wbNames, sheet: Object.fromEntries(Object.entries(sheetNames).map(([n, v]) => [ids[list.indexOf(n)], v])) } } } },
    c_oAscDefNameType: { table: TABLE },
  };
  const api = {
    GetSheets: () => list.map(n => ({ GetName: () => n })),
    GetSheet: n => (list.includes(n) ? {
      GetRange: () => ({ GetRow: () => 1, GetCol: () => 1, GetRowsCount: () => 1048576, GetColumnsCount: () => 16384 }),
      worksheet: { dataValidations: dv[n] === undefined ? null : dv[n] },
    } : null),
  };
  installBridge(api);
  return { api, asc };
}
const runBody = ({ api, asc }) => {
  global.Asc = asc;
  // eslint-disable-next-line no-new-func
  try { return new Function('Api', drv.dependentsBridgeBody())(api); } finally { delete global.Asc; }
};
/** The served manifest with another patch revision -- an unaccepted runtime (owner decision (B): the identity is the manifest). */
const withPatch = (m, patchVersion) => { const real = m.api.GetEuroOfficeBridge; m.api.GetEuroOfficeBridge = () => { const b = real(); const man = b.manifest(); return { manifest: () => ({ ...man, patchVersion }), read: b.read }; }; return m; };
const def = (name, ref, extra = {}) => ({ name, ref, hidden: false, ...extra });

test('B11 bridge body: names in the production shape and filter (hidden/_xlnm-named/ref-less dropped; table and the prefix-stripped isXLNM Print_Titles KEPT; name order); validations per sheet, a sheet with none as []', () => {
  const r = runBody(model({
    sheets: ['A', 'H', 'D'],
    wbNames: { wb: def('Wb', 'D!$B$2'), hid: def('Hid', 'D!$A$1', { hidden: true }), pa: def('_xlnm.Print_Area', 'D!$A$1:$B$2', { isXLNM: true }), pt: def('Print_Titles', 'D!$1:$1', { isXLNM: true }), nr: def('NoRef', null), t: def('Table1', 'A!$F$2:$G$3', { hidden: null, type: TABLE }), dup: def('Dup', 'A!$D$4') },
    sheetNames: { H: { loc: def('Loc', 'D!$B$2'), dup: def('Dup', 'H!$E$5') } },
    dv: { H: { elems: [{ type: 4, ranges: [{ r1: 0, c1: 4, r2: 0, c2: 4 }], formula1: { text: 'D!$A$1:$A$2' }, formula2: null, operator: 0 }] } },
  }));
  assert.deepEqual(r.names, [
    { name: 'Dup', ref: 'A!$D$4', scopeSheet: null }, { name: 'Dup', ref: 'H!$E$5', scopeSheet: 'H' },
    { name: 'Loc', ref: 'D!$B$2', scopeSheet: 'H' }, { name: 'Print_Titles', ref: 'D!$1:$1', scopeSheet: null }, { name: 'Table1', ref: 'A!$F$2:$G$3', scopeSheet: null }, { name: 'Wb', ref: 'D!$B$2', scopeSheet: null },
  ]);
  assert.deepEqual(r.validations, { A: [], H: [{ type: 4, ranges: [{ r1: 0, c1: 4, r2: 0, c2: 4 }], formula1: 'D!$A$1:$A$2', formula2: null }], D: [] });
});

test('B11 bridge body: every gap is UNOBSERVABLE on its own axis, never an empty list, no internal fallback', () => {
  const base = { sheets: ['A'], wbNames: { x: def('X', 'A!$A$1') } };
  // an unaccepted runtime (another patch revision) -> the bridge client's own identity gate
  assert.deepEqual(runBody(withPatch(model(base), 'euro-office-bridge-v1/9.3.4.60/2')), { names: 'UNOBSERVABLE:PATCH_VERSION_MISMATCH', validations: 'UNOBSERVABLE:PATCH_VERSION_MISMATCH' });
  // DS9346 (owner D5=B, docs/DS9346-RUNTIME-RETIREMENT.md): the served block is now the 9.3.5.0/16 identity.
  assert.equal(BRIDGE_RUNTIME_ID, '9.3.5|euro-office-bridge-v1/9.3.5.0/16', 'POSITIVE CONTROL: the served block is the accepted runtime');
  // no bridge at all
  const nb = model(base); delete nb.api.GetEuroOfficeBridge;
  assert.deepEqual(runBody(nb), { names: 'UNOBSERVABLE:BRIDGE_UNAVAILABLE', validations: 'UNOBSERVABLE:BRIDGE_UNAVAILABLE' });
  // a capability missing from the manifest
  const nc = model(base); const realB = nc.api.GetEuroOfficeBridge;
  nc.api.GetEuroOfficeBridge = () => { const b = realB(); const m = b.manifest(); m.capabilities = m.capabilities.filter(c => c.id !== 'workbook.definedNames.read/v1'); return { manifest: () => m, read: b.read }; };
  const rnc = runBody(nc);
  assert.equal(rnc.names, 'UNOBSERVABLE:CAPABILITY_MISSING');
  assert.deepEqual(rnc.validations, { A: [] }, 'the OTHER axis is unaffected');
  // a malformed capability result
  const mf = model(base); const realM = mf.api.GetEuroOfficeBridge;
  mf.api.GetEuroOfficeBridge = () => { const b = realM(); return { manifest: b.manifest, read: () => ({ ok: true, value: 'x' }) }; };
  assert.deepEqual(runBody(mf), { names: 'UNOBSERVABLE:UNEXPECTED_SHAPE', validations: 'UNOBSERVABLE:UNEXPECTED_SHAPE' });
  // a capability's own refusal (a sheet scope that does not resolve) -> its reason, prefixed
  const bad = model({ sheets: ['A'], sheetNames: {} });
  bad.asc.editor.wbModel.dependencyFormulas.defNames.sheet = { ghost: { g: def('G', 'A!$A$1') } };
  assert.equal(runBody(bad).names, 'UNOBSERVABLE:BRIDGE:sheet-scope-unresolved');
  // POSITIVE CONTROL: the same base reads fine
  assert.deepEqual(runBody(model(base)), { names: [{ name: 'X', ref: 'A!$A$1', scopeSheet: null }], validations: { A: [] } });
});

test('B11: no direct-internal dependency read is left in the driver (DIRECT_INTERNAL_FALLBACK=0)', () => {
  const src = require('node:fs').readFileSync(require.resolve('../../src/drivers/eurooffice/xlsx-driver.cjs'), 'utf8');
  for (const needle of ['asc_getDefinedNames', 'dataValidations', 'dependentsInFrame', 'EO_INTERNAL_DEPENDENCY_BUILDS']) assert.equal(src.includes(needle), false, needle);
  assert.equal(src.includes('dependentsBridgeBody'), true, 'POSITIVE CONTROL: the scan reads the driver source');
  assert.equal(drv.dependentsInFrame, undefined);
});

test('readLive: the dependency read is one callCommand through the bridge body; a failed command is UNOBSERVABLE, never an empty inventory', async () => {
  for (const [reply, want, wantGate] of [
    [{ ok: true, outcome: 'command-completed', result: { names: [], validations: { S: [] } } }, { names: [], validations: { S: [] } }, true],
    [{ ok: true, outcome: 'command-completed', result: { names: 'UNOBSERVABLE:CAPABILITY_MISSING', validations: { S: [] } } }, { names: 'UNOBSERVABLE:CAPABILITY_MISSING', validations: { S: [] } }, false],
    [{ ok: true, outcome: 'command-completed', result: null }, 'UNOBSERVABLE:dependency-read-failed:command-completed', false],
    [{ ok: false, outcome: 'command-callback-timeout' }, 'UNOBSERVABLE:dependency-read-failed:command-callback-timeout', false],
  ]) {
    let bodies = 0;
    const build = '9.3.4-9475e9f12a1905e2ada1fbc4bb87d01b';
    const page = { evaluate: async () => ({ ok: true, detail: 'stub' }), goto: async () => {}, waitForFunction: async () => {}, waitForTimeout: async () => {}, frames: () => [frameObj] };
    const frameObj = {
      url: () => `https://ds/${build}/web-apps/apps/spreadsheeteditor/main/index.html`,
      evaluate: async (fn, arg) => {
        if (arg && typeof arg.body === 'string' && arg.body.includes('workbook.definedNames.read/v1') && !arg.body.includes('__eoWireEncode')) { bodies += 1; return reply; }
        if (arg && typeof arg.body === 'string' && /GetActiveSheet/.test(arg.body)) return { ok: true, result: 'S' };
        if (arg && arg.body === HANDSHAKE_BODY) return { ok: true, result: null }; // no manifest: the session identity stays null (not under test here)
        const canned = { ok: true, sheets: [{ name: 'S', index: 0, used: null }], cells: [], truncated: false, unsupported: [], objects: { S: { charts: [], pivots: [], names: [] } }, conditional: { S: [] } };
        // M9-2 O1: the read body crosses the boundary through the wire codec -- the stand-in editor answers in kind
        return { ok: true, result: arg && typeof arg.body === 'string' && arg.body.includes('__eoWireEncode') ? encodeWire(canned) : canned };
      },
    };
    const pw = { chromium: { launch: async () => ({ newPage: async () => page, close: async () => {} }) } };
    const realFetch = global.fetch;
    global.fetch = async () => ({ ok: true, json: async () => ({ ocs: { data: { documentServerUrl: 'https://ds/', document: { key: 'k' } } } }) });
    try {
      const d = drv.createEuroOfficeXlsxDriver({ url: 'https://nc', user: 'u', pass: 'p', loadPlaywright: () => ({ ok: true, pw }), timeoutMs: 1000 });
      assert.equal((await d.open({ fileId: '1' })).ok, true);
      const r = await d.readLive(null, { dependents: true });
      assert.equal(bodies, 1);
      assert.deepEqual(r.payload.dependents, want);
      assert.equal(r.payload.runtime.dependencyRuntime, wantGate);
      await d.close();
    } finally { global.fetch = realFetch; }
  }
});

test('reference resolution: name uses, a formula through a name, and each unresolvable shape', () => {
  assert.deepEqual(nameUses('= SUM(Sales_Data!C2:C10)+MyName*2+IF(TRUE,A:A,1E5)+$B$2+"Lit!A1"').names, ['MyName']);
  const ctx = { sheets: ['D', 'H'], names: [{ name: 'N', ref: 'D!$B$2', scopeSheet: null }, { name: 'L', ref: 'N+1', scopeSheet: 'H' }] };
  assert.deepEqual(resolveRef('=L*2', 'H', ctx), { sheets: ['D'] }, 'through a sheet-level name through a workbook name');
  assert.deepEqual(resolveRef('=L*2', 'D', ctx), { problem: 'unknown-name:L' }, 'a sheet-level name is not visible from another sheet');
  assert.deepEqual(resolveRef('=Gone!A1', 'H', ctx), { problem: 'missing-sheet:Gone' });
  // a #REF! points at nothing: no dependency, and the rest of the formula is still read (the live D!A1 is seen)
  assert.deepEqual(resolveRef('= #REF!*2', 'H', ctx), { sheets: [] });
  assert.deepEqual(resolveRef('=D!A1+#REF!', 'H', ctx), { sheets: ['D'] });
  assert.deepEqual(resolveRef('=INDIRECT("D!A1")', 'H', ctx), { problem: 'indirect' });
  const cyc = { sheets: ['D'], names: [{ name: 'A_', ref: 'B_', scopeSheet: null }, { name: 'B_', ref: 'A_', scopeSheet: null }] };
  assert.match(resolveRef('=A_', 'D', cyc).problem, /name-cycle/);
});

test('copy rule comparison: formula2 is compared too (self reference mapped), not only formula1', () => {
  const { copyMatches } = require('../../src/adapters/xlsx/sheet-copy.cjs');
  const rule = f2 => ({ kind: 'validation', type: 1, range: 'E1', formula: 'S!$A$1', formula2: f2 });
  const p = (f2copy) => ({ cells: {}, objects: { S: { charts: [], pivots: [], names: [] }, C: { charts: [], pivots: [], names: [] } },
    rules: { complete: true, bySheet: { S: [rule('S!$A$2')], C: [{ ...rule(f2copy), formula: 'C!$A$1' }] } } });
  assert.equal(copyMatches(p('C!$A$2'), 'S', 'C').ok, true, 'positive control: mapped formula2');
  assert.equal(copyMatches(p('S!$A$2'), 'S', 'C').ok, false, 'formula2 still pointing at the source');
  // the mismatch names both sides, so a LIVE failure says WHAT differed (run 5 said only "rules")
  const d = copyMatches(p('S!$A$2'), 'S', 'C').diffs.find(x => x.change === 'rules');
  assert.deepEqual([d.sourceCount, d.copyCount, d.source[0].formula2, d.copy[0].formula2], [1, 1, 'S!$A$2', 'S!$A$2']);
});
