'use strict';

/**
 * GH #346 unit 2 (a maintainer's review #ref, point 4): chart.series.dataPoint.style.write/v1's own runtime --
 * /27 ONLY (the capability does not exist on /16..26), the next free Bridge-write label after /24's shipped
 * XLSX_B28_BRIDGE_WRITE_RUNTIME_IDS, B25..B28's own pattern. Core and adapter wired symmetrically, pinned to
 * each other by a static test (the adapter imports nothing outside its own tree -- a static test is the
 * ONLY thing that keeps the two copies from drifting).
 *
 * Scope: ONLY the identity-list wiring. The capability is NOT yet wired into any adapter family -- the
 * MCP-level dataPoint target release is a separate unit (3), out of this card's scope (the card's own
 * "MI NEM TARTOZIK BELE").
 */

const test = require('node:test');
const assert = require('node:assert/strict');

const core = require('../../src/core/runtime-identity.cjs');
const adapter = require('../../src/adapters/xlsx/runtime.cjs');

const ID27 = '9.3.5|euro-office-bridge-v1/9.3.5.0/27';

test('XLSX_B29_BRIDGE_WRITE_RUNTIME_IDS is /27 ONLY, core and adapter alike, and core ENTRIES.map(runtimeIdOf) matches core IDS', () => {
  assert.deepEqual([...core.XLSX_B29_BRIDGE_WRITE_RUNTIME_IDS], [ID27], 'core');
  assert.deepEqual([...adapter.XLSX_B29_BRIDGE_WRITE_RUNTIME_IDS], [ID27], 'adapter');
  assert.deepEqual(core.XLSX_B29_BRIDGE_WRITE_RUNTIME_ENTRIES.map(core.runtimeIdOf), [...core.XLSX_B29_BRIDGE_WRITE_RUNTIME_IDS]);
});

test('MUTATION CONTROL: removing /27 from the entries list changes the pinned value -- the list is not vacuously equal to anything', () => {
  const withoutLast = core.XLSX_B29_BRIDGE_WRITE_RUNTIME_ENTRIES.slice(0, -1).map(core.runtimeIdOf);
  assert.notDeepEqual(withoutLast, [...core.XLSX_B29_BRIDGE_WRITE_RUNTIME_IDS], 'a /27-less list must differ from the real, pinned one');
  assert.deepEqual(withoutLast, [], 'the entries array holds exactly one runtime (/27) -- removing it empties the list');
});

test('RED CONTROL: the new B29 list does NOT widen backwards -- /26, /24 and the pre-/19 runtimes are all refused by it', () => {
  for (const n of [16, 17, 18, 19, 21, 22, 23, 24, 25, 26]) {
    const id = `9.3.5|euro-office-bridge-v1/9.3.5.0/${n}`;
    assert.equal([...adapter.XLSX_B29_BRIDGE_WRITE_RUNTIME_IDS].includes(id), false, `adapter must refuse /${n}`);
    assert.equal([...core.XLSX_B29_BRIDGE_WRITE_RUNTIME_IDS].includes(id), false, `core must refuse /${n}`);
  }
});

test('the module exports are present on both sides (no silent typo in the export list)', () => {
  assert.equal('XLSX_B29_BRIDGE_WRITE_RUNTIME_ENTRIES' in core, true);
  assert.equal('XLSX_B29_BRIDGE_WRITE_RUNTIME_IDS' in core, true);
  assert.equal('XLSX_B29_BRIDGE_WRITE_RUNTIME_IDS' in adapter, true);
});
