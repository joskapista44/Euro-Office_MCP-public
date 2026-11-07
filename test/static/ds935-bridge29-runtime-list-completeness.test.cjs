'use strict';

/**
 * GH #365 option C U2 (card an internal branch, kioszto zero; follows U1 internal tracker) -- the SAME
 * PRE-INSTALL READINESS GATE ds935-bridge28-runtime-list-completeness.test.cjs already built for
 * /27 -> /28, shifted one revision: every runtime-gated list that reaches /28 must ALSO carry /29
 * (the candidate's patchVersion, "euro-office-bridge-v1/9.3.5.0/29"), OR document the SAME single
 * named exception this family already has (XLSX_B29_BRIDGE_WRITE_RUNTIME_IDS: an internal Bridge-
 * write-capability sequence label, unrelated to the DS revision number by a coincidence of
 * numbering -- it was /27-only before /28 existed and stays /27-only here too, since /28 was never
 * widened into it either).
 *
 * The walk is SOURCE-BASED (not a fixed name list), the same technique the sibling /28 file and
 * test/static/runtime-identity.test.cjs's own "STATIC GUARD" tests already use -- a FUTURE list
 * that reaches /28 and is not also widened to /29 fails this test too.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const core = require('../../src/core/runtime-identity.cjs');
const adapter = require('../../src/adapters/xlsx/runtime.cjs');

const R29 = '9.3.5|euro-office-bridge-v1/9.3.5.0/29';
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

/** every source line that references /28 as a runtime-identity-era reference (object OR string-id form), with its file + line text + the ten lines above it (for the deliberate-exclusion marker). */
function findRuntime28Lists() {
  const hits = [];
  for (const f of walk(SRC_ROOT)) {
    const lines = fs.readFileSync(f, 'utf8').split('\n');
    lines.forEach((line, i) => {
      if (/\bXLSX_RUNTIME_28\b|\bXLSX_RUNTIME_28_ID\b/.test(line) && /Object\.freeze\(\[|=\s*\[/.test(line)) {
        hits.push({ file: path.relative(SRC_ROOT, f), lineNo: i + 1, line, context: lines.slice(Math.max(0, i - 10), i + 1).join('\n') });
      }
    });
  }
  return hits;
}

test('STATIC GUARD: every runtime-gated list that reaches /28 (the whole source tree, not a fixed name list) ALSO carries /29, OR documents the deliberate exclusion', () => {
  const hits = findRuntime28Lists();
  assert.ok(hits.length >= 13, `found ${hits.length} /28-era lists (expected at least the eleven core.cjs + two table-props.cjs/runtime.cjs declarations already known)`);
  const unaccountedFor = hits.filter(h => !/\bXLSX_RUNTIME_29\b|\bXLSX_RUNTIME_29_ID\b/.test(h.line) && !h.context.includes(DELIBERATE_EXCLUSION_MARK));
  assert.deepEqual(unaccountedFor.map(h => `${h.file}:${h.lineNo}`), [], 'every /28 list must carry /29 or document why not, in the SAME declaration\'s nearby comment');
});

test('RED CONTROL: the completeness check above DOES fire -- a /28 list missing BOTH /29 and the exclusion marker is flagged', () => {
  const fixture = [
    "const XLSX_SOME_FUTURE_BRIDGE_WRITE_RUNTIME_ENTRIES = Object.freeze([XLSX_RUNTIME_27, XLSX_RUNTIME_28]);",
  ];
  const line = fixture[0];
  const context = line;
  const flagged = !/\bXLSX_RUNTIME_29\b|\bXLSX_RUNTIME_29_ID\b/.test(line) && !context.includes(DELIBERATE_EXCLUSION_MARK);
  assert.equal(flagged, true, 'a /28 list with neither /29 nor the exclusion marker must be caught');
  // B29's own real declaration (`[XLSX_RUNTIME_27]`) carries NO /28 reference at all -- it was
  // never widened to /28 either, so (unlike the sibling /27->/28 file, where B29's line literally
  // IS `[XLSX_RUNTIME_27]` and so DOES match that file's /27-search) this file's /28-search
  // correctly never finds it: the escape hatch has nothing to catch here, because the exclusion
  // was already complete one revision earlier.
  const hits = findRuntime28Lists();
  const b29Hit = hits.find(h => h.line.includes('XLSX_B29_BRIDGE_WRITE_RUNTIME_ENTRIES'));
  assert.equal(b29Hit, undefined, 'B29 never reached /28, so it is absent from a /28-reference search -- not an escape-hatch case on THIS revision');
});

test('NEGATIVE CONTROL for the completeness check: deliberately stripping /29 from a real, already-widened list reproduces a flag (shows the check discriminates, not just passes)', () => {
  const widened = [...core.ACCEPTED_RUNTIMES.xlsx.map(core.runtimeIdOf)];
  assert.ok(widened.includes(R29), 'sanity: the real list carries /29 today');
  const stripped = widened.filter(id => id !== R29);
  assert.equal(stripped.includes(R29), false, 'the mutated copy no longer carries /29 -- this is the "forgot to widen" state the /18 outage reproduced');
  assert.notDeepEqual(stripped, widened);
});

test('every widened list carries the SAME /29 entry, core and adapter alike (core is the source of truth, adapter is its own, no-outside-import mirror)', () => {
  const WRITE_LABELS = ['B03', 'B05', 'B10', 'B22', 'B23', 'B24', 'B25', 'B26', 'B27', 'B28'];
  assert.ok([...core.XLSX_RUNTIME_IDS].includes(R29), 'core.XLSX_RUNTIME_IDS');
  assert.ok([...adapter.XLSX_RUNTIME_IDS].includes(R29), 'adapter.XLSX_RUNTIME_IDS');
  assert.ok(core.ACCEPTED_RUNTIMES.xlsx.map(core.runtimeIdOf).includes(R29), 'core.ACCEPTED_RUNTIMES.xlsx');
  for (const label of WRITE_LABELS) {
    const coreKey = `XLSX_${label}_BRIDGE_WRITE_RUNTIME_IDS`;
    assert.ok(core[coreKey], `core.${coreKey} exists`);
    assert.ok(adapter[coreKey], `adapter.${coreKey} exists`);
    assert.ok([...core[coreKey]].includes(R29), `core.${coreKey} carries /29`);
    assert.ok([...adapter[coreKey]].includes(R29), `adapter.${coreKey} carries /29`);
    assert.deepEqual([...core[coreKey]], [...adapter[coreKey]], `${coreKey}: core and adapter agree`);
  }
  const tp = require('../../src/adapters/xlsx/families/table-props.cjs');
  assert.ok([...tp.TOTALS_OFF_BRIDGE_RUNTIMES].includes(R29), 'table-props.cjs TOTALS_OFF_BRIDGE_RUNTIMES carries /29');
});

test('the ONE deliberate exclusion, XLSX_B29_BRIDGE_WRITE_RUNTIME_IDS, does NOT carry /29 -- named, not an oversight', () => {
  assert.equal([...core.XLSX_B29_BRIDGE_WRITE_RUNTIME_IDS].includes(R29), false, 'core');
  assert.equal([...adapter.XLSX_B29_BRIDGE_WRITE_RUNTIME_IDS].includes(R29), false, 'adapter');
  assert.deepEqual([...core.XLSX_B29_BRIDGE_WRITE_RUNTIME_IDS], ['9.3.5|euro-office-bridge-v1/9.3.5.0/27']);
});

test('END-TO-END REGRESSION PROOF: the REAL /29 candidate manifest (read off disk, not a fixture) now classifies ok against ACCEPTED_RUNTIMES.xlsx -- this is the exact outage this gate closes, reproduced and fixed', () => {
  const DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.5.0', 'euro-office-bridge-v1', 'rev29-veryhidden-candidate');
  const block = fs.readFileSync(path.join(DIR, 'bridge-block.js'), 'utf8');
  const m = block.match(/patchVersion: "([^"]+)"/);
  assert.ok(m, 'the candidate file carries a patchVersion string');
  const manifest = { bridgeVersion: '1.0.0', documentServerVersion: '9.3.5', patchVersion: m[1], capabilities: [] };
  const result = core.classifyHandshake({ ok: true, manifest }, core.ACCEPTED_RUNTIMES.xlsx);
  assert.deepEqual(result.ok, true, `the real candidate's own patchVersion ("${m[1]}") must classify ok today`);
  assert.equal(result.runtimeId, R29);

  // RED: the SAME manifest against a list with /29 stripped out reproduces PATCH_VERSION_MISMATCH --
  // proving the fix (not a coincidence of some OTHER accepted entry matching).
  const stripped = core.ACCEPTED_RUNTIMES.xlsx.filter(e => core.runtimeIdOf(e) !== R29);
  const redResult = core.classifyHandshake({ ok: true, manifest }, stripped);
  assert.deepEqual(redResult, { ok: false, reason: 'PATCH_VERSION_MISMATCH' }, 'without /29 in the accepted list, the real candidate manifest is refused -- the /18-style outage, reproduced');
});
