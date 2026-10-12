'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

/**
 * M11.5c.1 (owner: M11.5 continuous mandate, owner channel, internal tracker) -- OFFLINE acceptance of the /8
 * CANDIDATE block (documentserver-patches/9.3.4.60/euro-office-bridge-v1/rev8-candidate/bridge-block.js): the B03-WRITE
 * Bridge capability worksheet.copy/v1. The ACTUAL candidate text runs against a fake editor (Asc.editor) that records
 * every call. PARITY: the bridge makes the product's B03 call -- asc_copyWorksheet(-1, [name], [index]) with the SAME
 * preconditions and success rule as copyInFrame (src/drivers/eurooffice/xlsx-driver.cjs); the M11.5c.0 probe measured
 * that this editor routine, called from the Bridge scope, gives a faithful, persisted copy appended last.
 */

const DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.4.60', 'euro-office-bridge-v1');
const read = p => fs.readFileSync(p, 'utf8');
const V7 = read(path.join(DIR, 'rev7-candidate', 'bridge-block.js'));
const V8 = read(path.join(DIR, 'rev8-candidate', 'bridge-block.js'));
const bridgeOf = (src, asc) => new Function('Api', 'AscCommon', 'Asc', `${src.replace(/,\s*$/, '')}; return Api.GetEuroOfficeBridge();`)({ GetSheet: () => null }, {}, asc);
const W = 'sheet.printSettings.write/v1';
const C = 'worksheet.copy/v1';
const DRIVER = read(path.join(__dirname, '..', '..', 'src', 'drivers', 'eurooffice', 'xlsx-driver.cjs'));
// the product's former B03 path, recorded from the last source that had it (9f9d3dd; M11.5c.4 removed it)
const PRE_CUTOVER = JSON.parse(read(path.join(__dirname, '..', 'fixtures', 'b03-direct-internal-copy-path-9f9d3dd.json'))).copyInFrame;

/** fake editor: sheets list + every API call recorded; `copy` decides what asc_copyWorksheet does to the list. */
function fakeEditor({ sheets = ['L', 'H'], missing = [], canEdit = true, lock = 0, prot = false, copy = 'append', error = false, throws = false } = {}) {
  const calls = []; const list = sheets.slice(); let onError = null;
  const e = {
    asc_getWorksheetsCount: () => list.length,
    asc_getWorksheetName: i => list[i],
    asc_isProtectedWorkbook: () => { calls.push(['asc_isProtectedWorkbook']); return prot; },
    asc_registerCallback: (n, f) => { calls.push(['asc_registerCallback', n]); onError = f; },
    asc_unregisterCallback: (n, f) => { calls.push(['asc_unregisterCallback', n, f === onError]); },
    canEdit: () => canEdit,
    collaborativeEditing: { getGlobalLock: () => lock },
    asc_copyWorksheet: (where, names, idx) => {
      calls.push(['asc_copyWorksheet', where, names, idx]);
      if (throws) throw new Error('boom');
      if (error && onError) onError(-1, 1);
      if (copy === 'append') list.push(names[0]); else if (copy === 'twice') list.push(names[0], names[0]); else if (copy === 'first') list.unshift(names[0]);
    },
  };
  for (const k of missing) delete e[k];
  return { asc: { editor: e }, calls, list, copies: () => calls.filter(c => c[0] === 'asc_copyWorksheet') };
}
const write = (env, args, id = C) => bridgeOf(V8, env.asc).write(id, args);

test('/8 manifest: patchVersion /8, /7 capabilities byte-identical, exactly one new WRITE entry', () => {
  const m7 = bridgeOf(V7, {}).manifest(); const m8 = bridgeOf(V8, {}).manifest();
  assert.equal(m7.patchVersion, 'euro-office-bridge-v1/9.3.4.60/7');
  assert.equal(m8.patchVersion, 'euro-office-bridge-v1/9.3.4.60/8');
  assert.equal(m8.capabilities.length, m7.capabilities.length + 1);
  assert.deepEqual(m8.capabilities.slice(0, m7.capabilities.length), m7.capabilities);
  const w = m8.capabilities.filter(c => c.id === C);
  assert.equal(w.length, 1);
  assert.deepEqual([w[0].version, w[0].kind, w[0].format, w[0].mutation], [1, 'WRITE', 'XLSX', 'TARGETED']);
  for (const k of ['bridgeVersion', 'documentServerVersion', 'documentServerBuild']) assert.equal(m8[k], m7[k]);
});

test('/8 = /7 + two inserted segments + the patchVersion bump (every /7 implementation byte-identical)', () => {
  let t = V8.replace('patchVersion: "euro-office-bridge-v1/9.3.4.60/8"', 'patchVersion: "euro-office-bridge-v1/9.3.4.60/7"');
  const cut = (s, startMark, endMark) => { const i = s.indexOf(startMark); const j = s.indexOf(endMark, i); assert.ok(i > 0 && j > i, startMark); return s.slice(0, i) + s.slice(j); };
  t = cut(t, '{ id: "worksheet.copy/v1"', '], };');
  t = cut(t, '"worksheet.copy/v1": function', '}; var EO_BRIDGE_HYDRATIONS = ');
  assert.equal(t, V7);
});

test('class separation: read()/hydrate() refuse the WRITE id; write() refuses other ids/versions; /7 has no copy', () => {
  const env = fakeEditor(); const b = bridgeOf(V8, env.asc); const args = { sheet: 'L', name: 'X' };
  assert.deepEqual(b.read(C, args), { ok: false, reason: 'capability-not-advertised' });
  assert.deepEqual(b.hydrate(C, args), { ok: false, reason: 'capability-not-advertised' });
  for (const id of ['worksheet.copy/v2', 'worksheet.copy.read/v1', 'nope']) assert.deepEqual(b.write(id, args), { ok: false, reason: 'capability-not-advertised' }, id);
  assert.deepEqual(env.copies(), []);
  assert.deepEqual(bridgeOf(V7, env.asc).write(C, args), { ok: false, reason: 'capability-not-advertised' }, '/7: B03 WRITE missing is detectable');
  assert.equal(typeof bridgeOf(V8, env.asc).write(W, { sheet: 'S', setting: 'scale', scale: 50 }).reason, 'string', '/8 write() still serves B10');
});

test('PARITY: the product B03 call (asc_copyWorksheet(-1, [name], [index])) with asc_onError registered and released; receipt only', () => {
  assert.match(PRE_CUTOVER, /e\.asc_copyWorksheet\(-1, \[dst\], \[index\]\)/, 'the former product routine this capability moved behind the bridge');
  assert.doesNotMatch(DRIVER, /\.asc_copyWorksheet\s*\(/, 'M11.5c.4: the product no longer calls it directly');
  const env = fakeEditor({ sheets: ['A', 'L', 'H'] });
  assert.deepEqual(write(env, { sheet: 'L', name: 'L_c' }), { ok: true, value: { applied: ['copy'], name: 'L_c', index: 3 } });
  assert.deepEqual(env.copies(), [['asc_copyWorksheet', -1, ['L_c'], [1]]]);
  assert.deepEqual(env.calls.map(c => c[0]), ['asc_isProtectedWorkbook', 'asc_registerCallback', 'asc_copyWorksheet', 'asc_unregisterCallback']);
  assert.equal(env.calls[3][2], true, 'the same handler unregistered');
  assert.deepEqual(env.list, ['A', 'L', 'H', 'L_c']);
});

test('refusals BEFORE the copy call (zero copies): arguments, name rules, collision (case-insensitive), source, permissions, missing API', () => {
  const cases = [
    [{}, null, 'bad-args'], [{ sheet: 'L' }, null, 'bad-args'], [{ sheet: '', name: 'X' }, null, 'bad-args'], [{ sheet: 'L', name: 1 }, null, 'bad-args'],
    [{ sheet: 'L', name: 'X', at: 1 }, null, 'unsupported-property'],
    [{ sheet: 'L', name: '' }, null, 'invalid-name'], [{ sheet: 'L', name: 'x'.repeat(32) }, null, 'invalid-name'],
    ...['a/b', 'a\\b', 'a[b', 'a]b', 'a:b', 'a*b', 'a?b', "'ab", "ab'"].map(n => [{ sheet: 'L', name: n }, null, 'invalid-name']),
    [{ sheet: 'L', name: 'h' }, null, 'name-taken'], [{ sheet: 'L', name: 'L' }, null, 'name-taken'],
    [{ sheet: 'Nope', name: 'X' }, null, 'source-not-found'],
    [{ sheet: 'L', name: 'X' }, { sheets: ['L', 'L'] }, 'source-ambiguous'],
    [{ sheet: 'L', name: 'X' }, { canEdit: false }, 'edit-not-allowed'],
    [{ sheet: 'L', name: 'X' }, { lock: 1 }, 'workbook-locked'],
    [{ sheet: 'L', name: 'X' }, { prot: true }, 'workbook-protected'],
    [{ sheet: 'L', name: 'X' }, { prot: null }, 'workbook-protected'],
    ...['asc_copyWorksheet', 'asc_getWorksheetsCount', 'canEdit', 'asc_registerCallback'].map(k => [{ sheet: 'L', name: 'X' }, { missing: [k] }, `${k}-missing`]),
  ];
  for (const [args, opts, reason] of cases) {
    const env = fakeEditor(opts || {});
    assert.equal(write(env, args).reason, reason, JSON.stringify([args, opts]));
    assert.deepEqual(env.copies(), [], JSON.stringify(args));
  }
  assert.deepEqual(bridgeOf(V8, {}).write(C, { sheet: 'L', name: 'X' }), { ok: false, reason: 'editor-missing' });
  assert.equal(write(fakeEditor(), { sheet: 'L', name: 'x'.repeat(31) }).ok, true, 'the 31-char boundary is a valid name');
});

test('after the call: error / throw / no effect / unexpected list -> ok:false with the observed lists, never retried inside', () => {
  const cases = [[{ error: true }, 'copy-error'], [{ throws: true }, 'copy-threw'], [{ copy: 'none' }, 'copy-not-observed'], [{ copy: 'twice' }, 'copy-unexpected'], [{ copy: 'first' }, 'copy-unexpected']];
  for (const [opts, reason] of cases) {
    const env = fakeEditor(opts);
    const r = write(env, { sheet: 'L', name: 'X' });
    assert.deepEqual([r.ok, r.reason], [false, reason], reason);
    assert.equal(env.copies().length, 1, `${reason}: exactly one call, no retry`);
    assert.ok(Array.isArray(r.before) && Array.isArray(r.after), `${reason}: the observed lists`);
    assert.equal(env.calls.filter(c => c[0] === 'asc_unregisterCallback').length, 1, `${reason}: the handler released`);
  }
});

test('the receipt carries no copied content (verification is the caller\'s fresh read)', () => {
  const r = write(fakeEditor(), { sheet: 'L', name: 'X' });
  assert.deepEqual(Object.keys(r.value).sort(), ['applied', 'index', 'name']);
});
