
'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { RUNTIMES, ACCEPTED_RUNTIMES } = require('../../src/core/runtime-identity.cjs');
const { BRIDGE_CLIENT_SRC, BRIDGE_HYDRATE_CLIENT_SRC, BRIDGE_CAPABILITIES, BRIDGE_HYDRATIONS } = require('../../src/drivers/eurooffice/ds-bridge.cjs');

/**
 * M11.5a.2 DUAL-RUNTIME COMPATIBILITY (owner GO: the owner, owner channel), historical.
 * DS9346 (owner D5=B, VALTAS, docs/DS9346-RUNTIME-RETIREMENT.md): the entire 9.3.4.60 chain this file compared
 * for parity (/4 through /17) is RETIRED -- ACCEPTED_RUNTIMES.xlsx is now the single 9.3.5.0/16 entry, and
 * EVERY historical identity now fails at the RUNTIME level (RUNTIME_VERSION_MISMATCH), not just the patch
 * level. The underlying mechanism this file tested -- that a served block's manifest identity is checked
 * against the accepted list, and that a wrong identity never silently serves a value -- is retained, now
 * exercised against the ONE accepted (9.3.5.0) block plus a representative sample of the now-fully-off-list
 * historical blocks, rather than the full /4..17 progression (that progression's own byte-identity across the
 * chain is proven elsewhere: test/static/b05-write-handler-identity.test.cjs and its siblings).
 */
const DIR935 = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.5.0', 'euro-office-bridge-v1');
const DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.4.60', 'euro-office-bridge-v1');
const V935 = fs.readFileSync(path.join(DIR935, 'bridge-block.js'), 'utf8').replace(/,\s*$/, '');
const V5 = fs.readFileSync(path.join(DIR, 'bridge-block.js'), 'utf8').replace(/,\s*$/, '');
const V6 = fs.readFileSync(path.join(DIR, 'rev6-candidate', 'bridge-block.js'), 'utf8').replace(/,\s*$/, '');
const V16 = fs.readFileSync(path.join(DIR, 'rev16-candidate', 'bridge-block.js'), 'utf8').replace(/,\s*$/, '');
const ALIGN = BRIDGE_CAPABILITIES.ALIGN_EXT_READ;

function api(src) {
  const calls = [];
  const inner = { getAlign: () => ({ getReadingOrder: () => 0, getIndent: () => 2, getShrinkToFit: () => true }), setIndent: v => calls.push(['setIndent', v]), setShrinkToFit: v => calls.push(['setShrinkToFit', v]) };
  const ws = { GetRange: () => ({ range: inner }), worksheet: { getIndex: () => 0, workbook: { oApi: { wb: { getWorksheet: () => ({ v: 1 }) } } } } };
  const Api = { GetSheet: n => (n === 'S' ? ws : null) };
  const bridge = new Function('Api', `${src}; return Api.GetEuroOfficeBridge();`)(Api);
  const realWrite = bridge.write; if (realWrite) bridge.write = (...a) => { calls.push(['bridge.write', a[0]]); return realWrite(...a); };
  Api.GetEuroOfficeBridge = () => bridge;
  return { Api, calls };
}
const read = (src, expected, cap = ALIGN, args = { sheet: 'S', cell: 'A1' }) => { const e = api(src); return { r: new Function('Api', `${BRIDGE_CLIENT_SRC}; return __eoBridgeRead(${JSON.stringify(cap.id)}, ${cap.version}, ${JSON.stringify(expected)}, ${JSON.stringify(args)}, ${JSON.stringify(cap.valueType || 'object')});`)(e.Api), calls: e.calls }; };
const hydrate = (src, expected) => { const e = api(src); const c = BRIDGE_HYDRATIONS.SHEET_MODEL_READY; return new Function('Api', `${BRIDGE_HYDRATE_CLIENT_SRC}; return __eoBridgeHydrate(${JSON.stringify(c.id)}, ${c.version}, ${JSON.stringify(expected)}, { sheet: 'S' });`)(e.Api); };

test('READ client with the accepted list: the 9.3.5.0 block serves the value; EVERY historical 9.3.4.60 block (/5, /6, /16) is refused -- RUNTIME_VERSION_MISMATCH, not just PATCH_VERSION_MISMATCH', () => {
  const v935 = read(V935, ACCEPTED_RUNTIMES.xlsx);
  assert.deepEqual(v935.r, { ok: true, value: { readingOrder: 0, indent: 2, shrink: true } });
  for (const src of [V5, V6, V16]) assert.deepEqual(read(src, ACCEPTED_RUNTIMES.xlsx).r, { ok: false, reason: 'RUNTIME_VERSION_MISMATCH' });
  assert.deepEqual(read(V6, RUNTIMES.xlsx).r, { ok: false, reason: 'RUNTIME_VERSION_MISMATCH' }, 'the single-9.3.5.0 form refuses the historical /6 too (positive control of the list)');
});

test('HYDRATION client with the accepted list: the 9.3.5.0 block PASSES; EVERY historical 9.3.4.60 block (/5, /6, /16) is refused', () => {
  assert.equal(hydrate(V935, ACCEPTED_RUNTIMES.xlsx).ok, true);
  for (const src of [V5, V6, V16]) assert.deepEqual(hydrate(src, ACCEPTED_RUNTIMES.xlsx), { ok: false, reason: 'RUNTIME_VERSION_MISMATCH' });
});

test('PRESENT_BUT_UNUSED: the READ client cannot be pointed at the WRITE entry, and no client path ever calls bridge.write', () => {
  const w = read(V935, ACCEPTED_RUNTIMES.xlsx, { id: 'range.alignmentExt.write/v1', version: 1, valueType: 'object' }, { sheet: 'S', cell: 'A1', indent: 3 });
  assert.deepEqual(w.r, { ok: false, reason: 'CAPABILITY_MISSING' }, 'kind WRITE is never a READ');
  assert.deepEqual(w.calls, [], 'BRIDGE_WRITE_CALLS=0, no setter touched');
  for (const cap of Object.values(BRIDGE_CAPABILITIES)) assert.notEqual(cap.id, 'range.alignmentExt.write/v1');
});

test('STATIC (M11.5a.4 supersedes the M11.5a.2 "0 references" pin): the write id and the bridge write() call live ONLY in ds-bridge.cjs', () => {
  const root = path.join(__dirname, '../../src'); const hits = [];
  const walk = d => { for (const e of fs.readdirSync(d, { withFileTypes: true })) { const p = path.join(d, e.name); if (e.isDirectory()) walk(p); else if (p.endsWith('.cjs')) { const s = fs.readFileSync(p, 'utf8'); if (/alignmentExt\.write|bridge\.write\s*\(|\.write\(\s*capabilityId/.test(s)) hits.push(path.relative(root, p)); } } };
  walk(root);
  assert.deepEqual(hits, ['drivers/eurooffice/ds-bridge.cjs']);
});
