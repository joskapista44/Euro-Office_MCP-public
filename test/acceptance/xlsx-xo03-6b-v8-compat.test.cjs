'use strict';

const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');
const assert = require('node:assert/strict');
const { ACCEPTED_RUNTIMES } = require('../../src/core/runtime-identity.cjs');
const { BRIDGE_CLIENT_SRC, BRIDGE_CAPABILITIES } = require('../../src/drivers/eurooffice/ds-bridge.cjs');
const { installBridge } = require('../fakes/ds-bridge-runtime.cjs');

/**
 * X-O03/6b COMPATIBILITY (point 3). TWO checks, because the "live" patchVersion and
 * a maintainer's named "/8" turned out NOT to be the same thing here -- MEASURED, not assumed (see the
 * sanity assertion below, which deliberately does not hardcode a version number):
 *
 *   installBridge() (documentserver-patches/9.3.4.60/euro-office-bridge-v1/manifest.json) -- the
 *   SAME root-manifest fake every OTHER acceptance test in this suite installs -- reports
 *   *** /5 ***, not /8. This is the last SDK-backed, formally promoted root manifest; whether the
 *   REAL the production host host is running /8 today (a maintainer's claim) or still /5 is NOT something this repo can
 *   answer (that promotion needs the real served SDK file, gated to the designated operator role --
 *   see rev9-candidate/build-rev9.py's own docstring). *** NOT RESOLVED HERE, FLAGGED BACK. ***
 *
 * Both checks use the SAME `bridgeAxis` mechanism the driver actually calls: an existing,
 * already-served capability still reads correctly, and `chart.series.scatterY.read/v1` degrades to
 * a named reason -- never a runtime rejection, never a crash -- on EITHER version.
 */
const readAs = (Api, cap, args, expected) => new Function('Api', `${BRIDGE_CLIENT_SRC}; return __eoBridgeRead(${JSON.stringify(cap.id)}, ${cap.version}, ${JSON.stringify(expected)}, ${JSON.stringify(args)}, ${JSON.stringify(cap.valueType || 'object')});`)(Api);
const read = (Api, cap, args) => readAs(Api, cap, args, ACCEPTED_RUNTIMES.xlsx);
const alignExtArgs = { sheet: 'S', cell: 'A1' };
/** Builds the fake `Api` with GetSheet ALREADY on it, then installs the bridge onto the SAME
 * object (the dispatch closures capture this exact object by reference -- a later spread-copy
 * with GetSheet added would be invisible to them, since they still call the ORIGINAL Api.GetSheet). */
function apiWithAlign() {
  const inner = { getAlign: () => ({ getReadingOrder: () => 0, getIndent: () => 2, getShrinkToFit: () => true }) };
  const ws = { GetRange: () => ({ range: inner }), worksheet: { getIndex: () => 0, workbook: { oApi: { wb: { getWorksheet: () => ({ v: 1 }) } } } } };
  const Api = { GetSheet: n => (n === 'S' ? ws : null) };
  return Api;
}

test('installBridge() -- the SAME live-served fake every other acceptance test uses -- reports a patchVersion; an existing capability still PASSES on it, scatterY degrades to a named reason (never a crash, never a runtime rejection)', () => {
  const Api = apiWithAlign();
  const runtimeId = installBridge(Api);
  // DS9346 (owner D5=B, VALTAS, docs/DS9346-RUNTIME-RETIREMENT.md): installBridge() now loads the 9.3.5.0
  // block (the ACCEPTED identity), which carries the FULL capability set -- scatterY is present, so on this
  // fake (no chart support at all) it no longer degrades to CAPABILITY_MISSING; it fails INSIDE the bridge's
  // own handler instead (BRIDGE:charts-threw) -- still a named reason, never a crash, never a runtime rejection.
  assert.match(runtimeId, /^9\.3\.5\|euro-office-bridge-v1\/9\.3\.5\.0\/\d+$/, 'sanity: a real runtime id came back (not asserting WHICH /N -- that is the point of this test)');

  const existing = read(Api, BRIDGE_CAPABILITIES.ALIGN_EXT_READ, alignExtArgs);
  assert.deepEqual(existing, { ok: true, value: { readingOrder: 0, indent: 2, shrink: true } }, 'an existing capability keeps working against the live-served fake, whatever its /N');

  const scatterY = read(Api, BRIDGE_CAPABILITIES.SCATTER_Y_READ, { sheet: 'S', chart: 'C1', series: 0 });
  assert.deepEqual(scatterY, { ok: false, reason: 'BRIDGE:charts-threw' }, 'scatterY degrades to a named reason -- not a crash, not a runtime rejection');
});

test('specifically against rev8-candidate/bridge-block.js (zero named "/8" explicitly): same two properties hold', () => {
  const v8Src = fs.readFileSync(path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.4.60', 'euro-office-bridge-v1', 'rev8-candidate', 'bridge-block.js'), 'utf8').replace(/,\s*$/, '');
  const Api = apiWithAlign();
  const bridge = new Function('Api', `${v8Src}; return Api.GetEuroOfficeBridge();`)(Api);
  Api.GetEuroOfficeBridge = () => bridge;
  const runtimeId = `${bridge.manifest().documentServerVersion}|${bridge.manifest().patchVersion}`;
  assert.equal(runtimeId, '9.3.4|euro-office-bridge-v1/9.3.4.60/8');

  // DS9346 (owner D5=B, VALTAS, docs/DS9346-RUNTIME-RETIREMENT.md): rev8-candidate's identity (9.3.4.60/8) is
  // now off the ACCEPTED list entirely -- this test is about this SPECIFIC historical block's capability
  // mechanics (an existing capability vs. a missing one), not about current acceptance, so it checks the
  // capability-level behavior with an EXPLICIT expected-runtime matching this exact block's own identity.
  const EXPECTED_V8 = { documentServerVersion: '9.3.4', patchVersion: 'euro-office-bridge-v1/9.3.4.60/8' };
  const existing = readAs(Api, BRIDGE_CAPABILITIES.ALIGN_EXT_READ, alignExtArgs, EXPECTED_V8);
  assert.deepEqual(existing, { ok: true, value: { readingOrder: 0, indent: 2, shrink: true } }, 'an existing capability keeps working against /8');

  const scatterY = readAs(Api, BRIDGE_CAPABILITIES.SCATTER_Y_READ, { sheet: 'S', chart: 'C1', series: 0 }, EXPECTED_V8);
  assert.deepEqual(scatterY, { ok: false, reason: 'CAPABILITY_MISSING' }, 'scatterY on /8: a named reason, not a crash and not a runtime rejection');
});
