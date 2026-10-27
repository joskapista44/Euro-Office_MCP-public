'use strict';

// ds935-live-harness-hardcodes-20261018: two LIVE harnesses had DS-9.3.4-specific values baked in, so the U4 LIVE
// round on a 9.3.5 Document Server would have found-nothing-wrong-but-wrong (a hardcoded /5 expectation refusing a
// genuinely accepted 9.3.5 manifest; a hardcoded 9.3.4.60 block path reading the wrong candidate file).
//   1) test/live/ds-runtime-identity-restart-acceptance.live.cjs classified the handshake against the single
//      historical RUNTIMES.xlsx entry (/5) instead of the product's own ACCEPTED_RUNTIMES.xlsx list.
//   2) test/live/xlsx-m12-chart-read-ab-parity.live.cjs hardcoded the 9.3.4.60 chart-read block path.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { ACCEPTED_RUNTIMES, XLSX_RUNTIME_IDS, classifyHandshake } = require('../../src/core/runtime-identity.cjs');
const { needsPrivate } = require('../helpers/private-source.cjs');

const RESTART_FILE = path.join(__dirname, '..', 'live', 'ds-runtime-identity-restart-acceptance.live.cjs');
const CHART_PARITY_FILE = path.join(__dirname, '..', 'live', 'xlsx-m12-chart-read-ab-parity.live.cjs');

const manifest935 = () => ({
  bridgeVersion: '1.0.0', documentServerVersion: '9.3.5', documentServerBuild: '9.3.5-9475e9f12a1905e2ada1fbc4bb87d01b',
  patchVersion: 'euro-office-bridge-v1/9.3.5.0/16',
  capabilities: [],
});
const hsResult935 = { ok: true, manifest: manifest935() };

test('POSITIVE CONTROL: the two live-harness files exist and are non-trivially long (the probe itself can see content)', { skip: needsPrivate('test/live/ds-runtime-identity-restart-acceptance.live.cjs', 'test/live/xlsx-m12-chart-read-ab-parity.live.cjs') }, () => {
  assert.ok(fs.statSync(RESTART_FILE).size > 2000);
  assert.ok(fs.statSync(CHART_PARITY_FILE).size > 2000);
});

test('NEGATIVE CONTROL: a freshly invented string is not found in either file', { skip: needsPrivate('test/live/ds-runtime-identity-restart-acceptance.live.cjs', 'test/live/xlsx-m12-chart-read-ab-parity.live.cjs') }, () => {
  const a = fs.readFileSync(RESTART_FILE, 'utf8');
  const b = fs.readFileSync(CHART_PARITY_FILE, 'utf8');
  assert.equal(a.includes('ZSIRAF7788_NEVER_WRITTEN_HERE'), false);
  assert.equal(b.includes('ZSIRAF7788_NEVER_WRITTEN_HERE'), false);
});

test('ds-runtime-identity-restart-acceptance.live.cjs no longer classifies against the single historical RUNTIMES.xlsx entry', { skip: needsPrivate('test/live/ds-runtime-identity-restart-acceptance.live.cjs') }, () => {
  const src = fs.readFileSync(RESTART_FILE, 'utf8');
  assert.equal(/classifyHandshake\([^)]*,\s*RUNTIMES\.xlsx\)/.test(src), false, 'the old, /5-pinned classifyHandshake call must be gone');
  assert.equal(/__eoBridgeRead\([^;]*\$\{J\(RUNTIMES\.xlsx\)\}/.test(src), false, 'the old, /5-pinned BRIDGE_READ body must be gone');
  assert.match(src, /classifyHandshake\(hs && hs\.ok \? hs\.result : null, ACCEPTED_RUNTIMES\.xlsx\)/);
  assert.match(src, /\$\{J\(ACCEPTED_RUNTIMES\.xlsx\)\}/);
  assert.match(src, /XLSX_RUNTIME_IDS\.includes\(out\.HANDSHAKE\.runtimeId\)/);
});

test('RED (the old pattern, reconstructed) would have refused a real 9.3.5 handshake; GREEN (the fixed pattern) accepts it', () => {
  // DS9346 (owner D5=B, VALTAS, docs/DS9346-RUNTIME-RETIREMENT.md): RUNTIMES.xlsx / XLSX_RUNTIME_ID are now
  // THEMSELVES the 9.3.5.0/16 identity (the historical 9.3.4.60/5 entry this RED reconstruction depended on
  // is retired) -- so the RED demonstration below uses a FROZEN, explicit copy of the old single-entry value,
  // rather than the live (now-repointed) export, to keep testing what it originally proved.
  const OLD_RUNTIMES_XLSX = { documentServerVersion: '9.3.4', patchVersion: 'euro-office-bridge-v1/9.3.4.60/5' };
  const OLD_XLSX_RUNTIME_ID = '9.3.4|euro-office-bridge-v1/9.3.4.60/5';

  // RED: the exact call the harness used to make -- a single runtime, the historical /5 entry (documentServerVersion
  // '9.3.4'). A 9.3.5 manifest fails on the FIRST gate (documentServerVersion itself), before patchVersion is even
  // compared -- RUNTIME_VERSION_MISMATCH, not PATCH_VERSION_MISMATCH (that reason is reserved for a 9.3.5 manifest
  // against a list that DOES have a 9.3.5 entry but the wrong patchVersion -- see test/static/runtime-identity.test.cjs).
  const red = classifyHandshake(hsResult935, OLD_RUNTIMES_XLSX);
  assert.deepEqual(red, { ok: false, reason: 'RUNTIME_VERSION_MISMATCH' });

  // GREEN: the fixed call -- the full accepted list, same as the product's own establishRuntime/dsSection.
  const green = classifyHandshake(hsResult935, ACCEPTED_RUNTIMES.xlsx);
  assert.equal(green.ok, true);
  assert.equal(green.runtimeId, '9.3.5|euro-office-bridge-v1/9.3.5.0/16');

  // the BRIDGE_PASS check: XLSX_RUNTIME_IDS.includes(...) accepts the matched candidate id, the old
  // runtimeId === OLD_XLSX_RUNTIME_ID (the historical /5 id) would not have.
  assert.equal(XLSX_RUNTIME_IDS.includes(green.runtimeId), true);
  assert.equal(green.runtimeId === OLD_XLSX_RUNTIME_ID, false, 'the candidate id is NOT the historical /5 id -- the old equality check would still refuse it');
});

test('BRIDGE_READ (exported) embeds the full ACCEPTED_RUNTIMES.xlsx list, including the 9.3.5 candidate patchVersion', { skip: needsPrivate('test/live/ds-runtime-identity-restart-acceptance.live.cjs') }, () => {
  const { BRIDGE_READ } = require(RESTART_FILE);
  assert.equal(typeof BRIDGE_READ, 'string');
  for (const r of ACCEPTED_RUNTIMES.xlsx) assert.ok(BRIDGE_READ.includes(JSON.stringify(r.patchVersion)), `missing ${r.patchVersion}`);
  // NEGATIVE: a patchVersion from the DOCX family must not appear (confirms the embedded list is the xlsx one, not a wildcard match)
  assert.equal(BRIDGE_READ.includes(JSON.stringify('euro-office-bridge-word-v1/9.3.4.60/3')), false);
});

test('xlsx-m12-chart-read-ab-parity.live.cjs: the chart-read block path is env-overridable via EURO_M12_CHART_READ_BLOCK, default unchanged', { skip: needsPrivate('test/live/xlsx-m12-chart-read-ab-parity.live.cjs') }, () => {
  const src = fs.readFileSync(CHART_PARITY_FILE, 'utf8');
  assert.match(src, /process\.env\.EURO_M12_CHART_READ_BLOCK/);
  delete process.env.EURO_M12_CHART_READ_BLOCK;
  const m = require(CHART_PARITY_FILE);
  const defaultPath = m.resolveBlock();
  assert.equal(defaultPath, m.DEFAULT_BLOCK);
  assert.equal(defaultPath, path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.4.60', 'euro-office-chart-read-v1', 'chart-read-block.js'));

  const candidate = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.5.0', 'euro-office-chart-read-v1', 'chart-read-block.js');
  process.env.EURO_M12_CHART_READ_BLOCK = candidate;
  try {
    assert.equal(m.resolveBlock(), path.resolve(candidate));
    // the install source (between the markers) is byte-identical between the two patch versions -- only the
    // identity fields ABOVE the markers (runtime/version comment) differ, and those are not part of installSource().
    const defaultSource = (() => { delete process.env.EURO_M12_CHART_READ_BLOCK; return m.installSource(); })();
    process.env.EURO_M12_CHART_READ_BLOCK = candidate;
    const candidateSource = m.installSource();
    assert.equal(candidateSource, defaultSource);
    assert.ok(candidateSource.length > 500);
  } finally {
    delete process.env.EURO_M12_CHART_READ_BLOCK;
  }
});

test('NEGATIVE CONTROL: installSource() is NOT vacuously equal for every file -- a genuinely different block reads differently', { skip: needsPrivate('test/live/xlsx-m12-chart-read-ab-parity.live.cjs') }, () => {
  const m = require(CHART_PARITY_FILE);
  const os = require('node:os');
  const defaultSource = m.resolveBlock && (() => { delete process.env.EURO_M12_CHART_READ_BLOCK; return m.installSource(); })();
  const real = fs.readFileSync(m.DEFAULT_BLOCK, 'utf8');
  const mutated = real.replace('/*EO-CHART-READ-INSTALL-BEGIN*/', '/*EO-CHART-READ-INSTALL-BEGIN*/var __zsiraf_marker_7788 = true;\n');
  assert.notEqual(mutated, real, 'the mutation must actually change the file');
  const tmp = path.join(os.tmpdir(), `ds935-chart-read-mutated-${process.pid}.js`);
  fs.writeFileSync(tmp, mutated);
  try {
    process.env.EURO_M12_CHART_READ_BLOCK = tmp;
    const mutatedSource = m.installSource();
    assert.notEqual(mutatedSource, defaultSource);
    assert.ok(mutatedSource.includes('__zsiraf_marker_7788'));
  } finally {
    delete process.env.EURO_M12_CHART_READ_BLOCK;
    fs.rmSync(tmp, { force: true });
  }
});

test('the runbook (docs/DS-9.3.5-LIVE-RUNBOOK.md) section 2c names EURO_M12_CHART_READ_BLOCK, not the old manual-edit workaround', { skip: needsPrivate('docs/DS-9.3.5-LIVE-RUNBOOK.md') }, () => {
  const runbook = fs.readFileSync(path.join(__dirname, '..', '..', 'docs', 'DS-9.3.5-LIVE-RUNBOOK.md'), 'utf8');
  assert.match(runbook, /EURO_M12_CHART_READ_BLOCK/);
  assert.equal(runbook.includes('EGYSOROS kód-módosítás a harnessben a futtatás előtt'), false, 'the old manual-edit instruction must be removed');
});
