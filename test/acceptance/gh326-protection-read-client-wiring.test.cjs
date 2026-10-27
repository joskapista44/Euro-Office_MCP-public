'use strict';

/**
 * GH #326 / #327 (U4a, owner decision owner channel answer to 41211, path A): the CLIENT-side wiring of
 * sheet.protection.read/v1 into the product inventory (xlsx-driver.cjs's `protectionReads`, via the SAME
 * bridgeAxis convention X-O08's FILL_LINE_COLOR_READ already uses) -- NOT the guard (#326) or the redaction
 * (#327), both later units; this unit only proves the axis itself degrades/reads correctly.
 *
 * Offline, against the REAL bridge-block code (not an imitation): the shipped /16 block
 * (test/fakes/ds-bridge-runtime.cjs installBridge, the SAME fixture xlsx-m7-s2b.test.cjs uses) does not carry
 * the capability at all -- CAPABILITY_MISSING, never a guessed value. The /17 CANDIDATE block
 * (documentserver-patches/9.3.5.0/euro-office-bridge-v1/rev17-candidate/bridge-block.js, a maintainer's own commit
 * 361e9284) is loaded the SAME way test/acceptance/ds935-bridge17-protection-read.test.cjs already does, onto a
 * fake worksheet in the LIVE-measured shape (worksheet.sheetProtection).
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const { readQuery } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { installBridge } = require('../fakes/ds-bridge-runtime.cjs');

const DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.5.0', 'euro-office-bridge-v1');
const V17 = fs.readFileSync(path.join(DIR, 'rev17-candidate', 'bridge-block.js'), 'utf8');

/** The SAME fake worksheet shape a maintainer's own ds935-bridge17-protection-read.test.cjs uses, plus GetName (the
 * driver's own per-sheet existence check needs it -- the bridge handler itself only reads `.worksheet`). */
function sheet(name, sp = null) {
  return { GetName: () => name, GetIndex: () => 0, GetUsedRange: () => null, worksheet: { sheetProtection: sp }, GetRange: () => { throw new Error('range walk not exercised by this unit'); } };
}

const protectedSp = Object.freeze({
  sheet: true, algorithmName: 'SENTINEL-ALG', hashValue: 'SENTINEL-HASH', saltValue: 'SENTINEL-SALT', spinCount: 'SENTINEL-SPIN', password: 'SENTINEL-PW', temporaryPassword: 'SENTINEL-TMP',
  objects: false, scenarios: false, formatCells: false, formatColumns: false, formatRows: false, insertColumns: false, insertRows: true, insertHyperlinks: false,
  deleteColumns: false, deleteRows: false, selectLockedCells: false, sort: false, autoFilter: false, pivotTables: false, selectUnlockedCells: false,
});

const run = (api, protectionSheets) => new Function('Api', readQuery([], { protectionSheets }))(api).protectionReads;

test('a /16 runtime (the shipped block) does not advertise the capability at all -- UNOBSERVABLE:CAPABILITY_MISSING, never a guessed value', () => {
  const ws = sheet('Sheet1');
  const api = { GetSheets: () => [ws], GetSheet: n => (n === 'Sheet1' ? ws : null), WorksheetFunction: { COUNT: () => 0 } };
  const runtimeId = installBridge(api);
  assert.equal(runtimeId, '9.3.5|euro-office-bridge-v1/9.3.5.0/16', 'calibration: the shipped block really is /16');
  assert.deepEqual(run(api, ['Sheet1']), { Sheet1: 'UNOBSERVABLE:CAPABILITY_MISSING' });
});

test('a /17 runtime (the candidate block) reads protected + forbids through the SAME bridgeAxis convention', () => {
  const ws = sheet('P', protectedSp);
  const api = { GetSheets: () => [ws], GetSheet: n => (n === 'P' ? ws : null), WorksheetFunction: { COUNT: () => 0 } };
  new Function('Api', V17.replace(/\}\s*,\s*$/, '}'))(api);
  const m = api.GetEuroOfficeBridge().manifest();
  assert.equal(m.patchVersion, 'euro-office-bridge-v1/9.3.5.0/17', 'calibration: the candidate block really is /17');
  const out = run(api, ['P']);
  assert.equal(out.P.protected, true);
  assert.equal(out.P.forbids.insertRows, true);
  assert.equal(out.P.forbids.sort, false);
  assert.deepEqual(Object.keys(out.P), ['protected', 'forbids'], 'no `cells` field -- the range axis is out of this unit\'s scope');
  // the owner condition (never a credential in the output), re-checked at THIS client's own call site
  const serialized = JSON.stringify(out);
  for (const v of ['SENTINEL-ALG', 'SENTINEL-HASH', 'SENTINEL-SALT', 'SENTINEL-SPIN', 'SENTINEL-PW', 'SENTINEL-TMP']) assert.ok(!serialized.includes(v), `leaked ${v}`);
});

test('an unprotected sheet on /17 reads protected: false, forbids: null', () => {
  const ws = sheet('U', null);
  const api = { GetSheets: () => [ws], GetSheet: n => (n === 'U' ? ws : null), WorksheetFunction: { COUNT: () => 0 } };
  new Function('Api', V17.replace(/\}\s*,\s*$/, '}'))(api);
  assert.deepEqual(run(api, ['U']), { U: { protected: false, forbids: null } });
});

test('not asked -- protectionReads is generated empty, no bridge call at all (same convention as drawingReads/chartReads)', () => {
  assert.doesNotMatch(readQuery(), /protectionAsked/, 'the per-sheet loop (and its bridgeAxis call) is generated only inside the conditional block');
  const ws = sheet('P', protectedSp);
  const api = { GetSheets: () => [ws], GetSheet: n => (n === 'P' ? ws : null), WorksheetFunction: { COUNT: () => 0 } };
  installBridge(api);
  assert.deepEqual(new Function('Api', readQuery())(api).protectionReads, {});
});

test('"*" asks every sheet, the SAME all-sheets convention read_charts/read_drawings already use', () => {
  const a = sheet('A'); const b = sheet('B', protectedSp);
  const api = { GetSheets: () => [a, b], GetSheet: n => ({ A: a, B: b }[n] || null), WorksheetFunction: { COUNT: () => 0 } };
  new Function('Api', V17.replace(/\}\s*,\s*$/, '}'))(api);
  const out = run(api, ['*']);
  assert.deepEqual(Object.keys(out).sort(), ['A', 'B']);
  assert.equal(out.A.protected, false);
  assert.equal(out.B.protected, true);
});

test('a missing sheet names the gap, never a guessed empty object', () => {
  const api = { GetSheets: () => [], GetSheet: () => null, WorksheetFunction: { COUNT: () => 0 } };
  installBridge(api);
  assert.deepEqual(run(api, ['NoSuchSheet']), { NoSuchSheet: 'UNOBSERVABLE:sheet-missing' });
});
