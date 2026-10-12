
'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { RUNTIMES, ACCEPTED_RUNTIMES } = require('../../src/core/runtime-identity.cjs');
const { BRIDGE_CLIENT_SRC, BRIDGE_HYDRATE_CLIENT_SRC, BRIDGE_CAPABILITIES, BRIDGE_HYDRATIONS } = require('../../src/drivers/eurooffice/ds-bridge.cjs');

/**
 * M11.5a.2 DUAL-RUNTIME COMPATIBILITY (owner GO: the owner, owner channel). The product accepts the cell bridge /5 AND
 * /6 by an EXPLICIT list; the /6-only capability range.alignmentExt.write/v1 is PRESENT_BUT_UNUSED. Evaluated against the
 * ACTUAL bridge blocks (/5 = the served revision, /6 = the M11.5a.1 candidate) with a fake Api.
 */
const DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.4.60', 'euro-office-bridge-v1');
const V5 = fs.readFileSync(path.join(DIR, 'bridge-block.js'), 'utf8').replace(/,\s*$/, '');
const V6 = fs.readFileSync(path.join(DIR, 'rev6-candidate', 'bridge-block.js'), 'utf8').replace(/,\s*$/, '');
const V9 = V6.replace('patchVersion: "euro-office-bridge-v1/9.3.4.60/6"', 'patchVersion: "euro-office-bridge-v1/9.3.4.60/9"');
const V7 = fs.readFileSync(path.join(DIR, 'rev7-candidate', 'bridge-block.js'), 'utf8').replace(/,\s*$/, '');
const V8 = fs.readFileSync(path.join(DIR, 'rev8-candidate', 'bridge-block.js'), 'utf8').replace(/,\s*$/, '');
const V4 = V5.replace('patchVersion: "euro-office-bridge-v1/9.3.4.60/5"', 'patchVersion: "euro-office-bridge-v1/9.3.4.60/4"');
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

test('READ client with the accepted list: /5, the REAL /6, /7 and /8 blocks serve the same value; /4 and /9 -> PATCH_VERSION_MISMATCH', () => {
  const v5 = read(V5, ACCEPTED_RUNTIMES.xlsx), v6 = read(V6, ACCEPTED_RUNTIMES.xlsx);
  assert.deepEqual(v5.r, { ok: true, value: { readingOrder: 0, indent: 2, shrink: true } });
  assert.deepEqual(v6.r, v5.r, 'PARITY /5 vs /6');
  assert.deepEqual(read(V7, ACCEPTED_RUNTIMES.xlsx).r, v5.r, 'PARITY /5 vs /7');
  assert.deepEqual(read(V8, ACCEPTED_RUNTIMES.xlsx).r, v5.r, 'PARITY /5 vs /8');
  for (const src of [V4, V9]) assert.deepEqual(read(src, ACCEPTED_RUNTIMES.xlsx).r, { ok: false, reason: 'PATCH_VERSION_MISMATCH' });
  assert.deepEqual(read(V6, RUNTIMES.xlsx).r, { ok: false, reason: 'PATCH_VERSION_MISMATCH' }, 'the single-/5 form still refuses /6 (positive control of the list)');
});

test('HYDRATION client with the accepted list: /5, /6, /7 and /8 PASS (parity), /9 refused', () => {
  assert.deepEqual(hydrate(V6, ACCEPTED_RUNTIMES.xlsx), hydrate(V5, ACCEPTED_RUNTIMES.xlsx));
  assert.deepEqual(hydrate(V7, ACCEPTED_RUNTIMES.xlsx), hydrate(V5, ACCEPTED_RUNTIMES.xlsx));
  assert.deepEqual(hydrate(V8, ACCEPTED_RUNTIMES.xlsx), hydrate(V5, ACCEPTED_RUNTIMES.xlsx));
  assert.equal(hydrate(V5, ACCEPTED_RUNTIMES.xlsx).ok, true);
  assert.deepEqual(hydrate(V9, ACCEPTED_RUNTIMES.xlsx), { ok: false, reason: 'PATCH_VERSION_MISMATCH' });
});

test('PRESENT_BUT_UNUSED: the READ client cannot be pointed at the /6 WRITE entry, and no client path ever calls bridge.write', () => {
  const w = read(V6, ACCEPTED_RUNTIMES.xlsx, { id: 'range.alignmentExt.write/v1', version: 1, valueType: 'object' }, { sheet: 'S', cell: 'A1', indent: 3 });
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
