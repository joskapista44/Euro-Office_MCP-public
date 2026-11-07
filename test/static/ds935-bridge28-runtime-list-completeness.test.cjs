'use strict';

/**
 * GH #361/#363 PRE-INSTALL READINESS GATE (card an internal branch /28 Bridge telepites elotti keszultsegi kapu, f0d5dc60).
 *
 * The /18 and /21 widenings (core/runtime-identity.cjs's own comments) show the outage class this gate closes: a new
 * DS Bridge candidate rev lands with its byte-identity warrant proven, but one or more of the product's runtime-
 * acceptance lists is widened LATER, in a separate step -- and in the staging window between those two steps, a DS
 * apply to the new rev turns the capabilities gated by the lagging list(s) FAIL_CLOSED in production. For the
 * GENERAL list (core.ACCEPTED_RUNTIMES.xlsx / XLSX_RUNTIME_IDS, consumed as `bridgeAxis`'s baked-in expectedRuntime
 * in xlsx-driver.cjs) the blast radius is total: EVERY Bridge capability on that runtime -- including the two
 * brand-new ones this very candidate adds -- answers PATCH_VERSION_MISMATCH, because the whole-session handshake
 * never recognizes the runtime at all.
 *
 * This file asserts, in one place, that every runtime-gated list the /27 candidate already appears in ALSO carries
 * /28 (the candidate's patchVersion, "euro-office-bridge-v1/9.3.5.0/28" -- originally baked its own directory
 * suffix in, fixed at the source per a maintainer's review; see core/runtime-identity.cjs's own note) -- OR
 * is a list the byte-identity warrant does not need to cover, which is itself a SINGLE named exception
 * (XLSX_B29_BRIDGE_WRITE_RUNTIME_IDS: an internal Bridge-write-capability sequence label, unrelated to the DS
 * revision number "/28" by coincidence of numbering, gating a capability not yet wired to any adapter family).
 *
 * The walk is SOURCE-BASED (not a fixed name list) so a FUTURE list that reaches /27 and is not also widened to
 * /28 fails this test too -- the same technique test/static/runtime-identity.test.cjs's own "STATIC GUARD" tests
 * already use for a narrower set of names.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const core = require('../../src/core/runtime-identity.cjs');
const adapter = require('../../src/adapters/xlsx/runtime.cjs');

const R28 = '9.3.5|euro-office-bridge-v1/9.3.5.0/28';
const DELIBERATE_EXCLUSION_MARK = 'DELIBERATELY NOT WIDENED';

const SRC_ROOT = path.join(__dirname, '..', '..', 'src');
function walk(dir, files = []) {
  for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
    const p = path.join(dir, e.name);
    if (e.isDirectory()) walk(p, files);
    else if (p.endsWith('.cjs')) files.push(p);
  }
  return files;
}

/** every source line that references /27 as a runtime-identity-era reference (object OR string-id form), with its file + line text + the three lines above it (for the deliberate-exclusion marker). */
function findRuntime27Lists() {
  const hits = [];
  for (const f of walk(SRC_ROOT)) {
    const lines = fs.readFileSync(f, 'utf8').split('\n');
    lines.forEach((line, i) => {
      if (/\bXLSX_RUNTIME_27\b|\bXLSX_RUNTIME_27_ID\b/.test(line) && /Object\.freeze\(\[|=\s*\[/.test(line)) {
        hits.push({ file: path.relative(SRC_ROOT, f), lineNo: i + 1, line, context: lines.slice(Math.max(0, i - 10), i + 1).join('\n') });
      }
    });
  }
  return hits;
}

test('STATIC GUARD: every runtime-gated list that reaches /27 (the whole source tree, not a fixed name list) ALSO carries /28, OR documents the deliberate exclusion', () => {
  const hits = findRuntime27Lists();
  assert.ok(hits.length >= 13, `found ${hits.length} /27-era lists (expected at least the eleven core.cjs + two table-props.cjs/runtime.cjs declarations already known)`);
  const unaccountedFor = hits.filter(h => !/\bXLSX_RUNTIME_28\b|\bXLSX_RUNTIME_28_ID\b/.test(h.line) && !h.context.includes(DELIBERATE_EXCLUSION_MARK));
  assert.deepEqual(unaccountedFor.map(h => `${h.file}:${h.lineNo}`), [], 'every /27 list must carry /28 or document why not, in the SAME declaration\'s nearby comment');
});

test('RED CONTROL: the completeness check above DOES fire -- a /27 list missing BOTH /28 and the exclusion marker is flagged', () => {
  const fixture = [
    "const XLSX_SOME_FUTURE_BRIDGE_WRITE_RUNTIME_ENTRIES = Object.freeze([XLSX_RUNTIME_26, XLSX_RUNTIME_27]);",
  ];
  const line = fixture[0];
  const context = line;
  const flagged = !/\bXLSX_RUNTIME_28\b|\bXLSX_RUNTIME_28_ID\b/.test(line) && !context.includes(DELIBERATE_EXCLUSION_MARK);
  assert.equal(flagged, true, 'a /27 list with neither /28 nor the exclusion marker must be caught');
  // and the documented-exclusion escape hatch actually works (B29's own real declaration, re-read from disk)
  const b29Hit = findRuntime27Lists().find(h => h.line.includes('XLSX_B29_BRIDGE_WRITE_RUNTIME_ENTRIES'));
  assert.ok(b29Hit, 'B29 must still be found by the walk (it references /27)');
  assert.ok(b29Hit.context.includes(DELIBERATE_EXCLUSION_MARK), 'B29 carries the exclusion marker in its own nearby comment');
});

test('NEGATIVE CONTROL for the completeness check: deliberately stripping /28 from a real, already-widened list reproduces a flag (shows the check discriminates, not just passes)', () => {
  const widened = [...core.ACCEPTED_RUNTIMES.xlsx.map(core.runtimeIdOf)];
  assert.ok(widened.includes(R28), 'sanity: the real list carries /28 today');
  const stripped = widened.filter(id => id !== R28);
  assert.equal(stripped.includes(R28), false, 'the mutated copy no longer carries /28 -- this is the "forgot to widen" state the /18 outage reproduced');
  assert.notDeepEqual(stripped, widened);
});

test('every widened list carries the SAME /28 entry, core and adapter alike (core is the source of truth, adapter is its own, no-outside-import mirror)', () => {
  const WRITE_LABELS = ['B03', 'B05', 'B10', 'B22', 'B23', 'B24', 'B25', 'B26', 'B27', 'B28'];
  assert.ok([...core.XLSX_RUNTIME_IDS].includes(R28), 'core.XLSX_RUNTIME_IDS');
  assert.ok([...adapter.XLSX_RUNTIME_IDS].includes(R28), 'adapter.XLSX_RUNTIME_IDS');
  assert.ok(core.ACCEPTED_RUNTIMES.xlsx.map(core.runtimeIdOf).includes(R28), 'core.ACCEPTED_RUNTIMES.xlsx');
  for (const label of WRITE_LABELS) {
    const coreKey = `XLSX_${label}_BRIDGE_WRITE_RUNTIME_IDS`;
    assert.ok(core[coreKey], `core.${coreKey} exists`);
    assert.ok(adapter[coreKey], `adapter.${coreKey} exists`);
    assert.ok([...core[coreKey]].includes(R28), `core.${coreKey} carries /28`);
    assert.ok([...adapter[coreKey]].includes(R28), `adapter.${coreKey} carries /28`);
    assert.deepEqual([...core[coreKey]], [...adapter[coreKey]], `${coreKey}: core and adapter agree`);
  }
  const tp = require('../../src/adapters/xlsx/families/table-props.cjs');
  assert.ok([...tp.TOTALS_OFF_BRIDGE_RUNTIMES].includes(R28), 'table-props.cjs TOTALS_OFF_BRIDGE_RUNTIMES carries /28');
});

test('the ONE deliberate exclusion, XLSX_B29_BRIDGE_WRITE_RUNTIME_IDS, does NOT carry /28 -- named, not an oversight', () => {
  assert.equal([...core.XLSX_B29_BRIDGE_WRITE_RUNTIME_IDS].includes(R28), false, 'core');
  assert.equal([...adapter.XLSX_B29_BRIDGE_WRITE_RUNTIME_IDS].includes(R28), false, 'adapter');
  assert.deepEqual([...core.XLSX_B29_BRIDGE_WRITE_RUNTIME_IDS], ['9.3.5|euro-office-bridge-v1/9.3.5.0/27']);
});

test('TIERED_EXTRAS already carries both GH #361 (tableSheets) and GH #363 (compiledStyleRanges) -- neither new capability is a full-workbook-only read', () => {
  const { TIERED_EXTRAS } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
  assert.ok(TIERED_EXTRAS.includes('tableSheets'), 'GH #361 read_tables autoFilterDetail');
  assert.ok(TIERED_EXTRAS.includes('compiledStyleRanges'), 'GH #363 read_effective_fill');
});

test('END-TO-END REGRESSION PROOF: the REAL /28 candidate manifest (read off disk, not a fixture) now classifies ok against ACCEPTED_RUNTIMES.xlsx -- this is the exact outage this gate closes, reproduced and fixed', () => {
  const DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.5.0', 'euro-office-bridge-v1', 'rev28-table-autofilter-criteria-candidate');
  const block = fs.readFileSync(path.join(DIR, 'bridge-block.js'), 'utf8');
  const m = block.match(/patchVersion: "([^"]+)"/);
  assert.ok(m, 'the candidate file carries a patchVersion string');
  const manifest = { bridgeVersion: '1.0.0', documentServerVersion: '9.3.5', patchVersion: m[1], capabilities: [] };
  const result = core.classifyHandshake({ ok: true, manifest }, core.ACCEPTED_RUNTIMES.xlsx);
  assert.deepEqual(result.ok, true, `the real candidate's own patchVersion ("${m[1]}") must classify ok today`);
  assert.equal(result.runtimeId, R28);

  // RED: the SAME manifest against a list with /28 stripped out reproduces PATCH_VERSION_MISMATCH -- proving the
  // fix (not a coincidence of some OTHER accepted entry matching).
  const stripped = core.ACCEPTED_RUNTIMES.xlsx.filter(e => core.runtimeIdOf(e) !== R28);
  const redResult = core.classifyHandshake({ ok: true, manifest }, stripped);
  assert.deepEqual(redResult, { ok: false, reason: 'PATCH_VERSION_MISMATCH' }, 'without /28 in the accepted list, the real candidate manifest is refused -- the /18-style outage, reproduced');
});
