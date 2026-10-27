'use strict';

/**
 * v1.1.0 (B3): documentserver-patches/9.3.5.0/fresh-install-v1.1.0.json is the install plan of the release -- it must
 * install exactly the Bridge revisions the product accepts for writes (cell /17, word /5), pin the canonical block
 * bytes, and start from the same unpatched 9.3.5 bundles as the earlier /16 + /3 plan (fresh-install.json, kept for
 * the record). The output hashes were measured by applying the plan to the real unpatched bundles (a throwaway path)
 * and equal the bundles the production Document Server serves (cell 1a8c9ab1..., word dd9c4a9a..., 2026-10-27).
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const crypto = require('node:crypto');
const { ACCEPTED_RUNTIMES } = require('../../src/core/runtime-identity.cjs');

const DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.5.0');
const PLAN = JSON.parse(fs.readFileSync(path.join(DIR, 'fresh-install-v1.1.0.json'), 'utf8'));
const OLD = JSON.parse(fs.readFileSync(path.join(DIR, 'fresh-install.json'), 'utf8'));
const sha = file => crypto.createHash('sha256').update(fs.readFileSync(path.join(DIR, file))).digest('hex');
const HEX64 = /^[0-9a-f]{64}$/;

test('the plan installs cell /17 and word /5 -- the accepted runtimes, the last block carries the patchVersion', () => {
  assert.equal(PLAN.cell.patchVersion, 'euro-office-bridge-v1/9.3.5.0/17');
  assert.equal(PLAN.word.patchVersion, 'euro-office-bridge-word-v1/9.3.5.0/5');
  for (const [target, format] of [['cell', 'xlsx'], ['word', 'docx']]) {
    assert.ok(ACCEPTED_RUNTIMES[format].some(r => r.patchVersion === PLAN[target].patchVersion), `${target}: not an accepted runtime`);
    assert.equal(PLAN[target].blocks.at(-1).layer, PLAN[target].patchVersion);
  }
});

test('every pinned block hash equals the canonical block file', () => {
  for (const target of ['cell', 'word']) {
    for (const b of PLAN[target].blocks) assert.equal(sha(b.file), b.sha256, `${target}: ${b.file}`);
  }
});

test('the inputs are the same unpatched 9.3.5 bundles as the earlier plan; the outputs are pinned and differ from it', () => {
  for (const target of ['cell', 'word']) {
    assert.equal(PLAN[target].input_sha256, OLD[target].input_sha256);
    assert.equal(PLAN[target].file, OLD[target].file);
    assert.match(PLAN[target].output_sha256, HEX64);
    assert.notEqual(PLAN[target].output_sha256, OLD[target].output_sha256);
  }
});

/** The two checks above, as functions over any plan, so the RED control runs the same logic on a broken copy. */
const blockErrors = plan => ['cell', 'word'].flatMap(t => plan[t].blocks.filter(b => sha(b.file) !== b.sha256).map(b => `${t}:${b.file}`));
const runtimeErrors = plan => [['cell', 'xlsx'], ['word', 'docx']]
  .filter(([t, f]) => !ACCEPTED_RUNTIMES[f].some(r => r.patchVersion === plan[t].patchVersion)).map(([t]) => t);

test('RED CONTROL: the same checks fail on a copy with a wrong block hash, or with a patchVersion no runtime accepts', () => {
  assert.deepEqual(blockErrors(PLAN), []);
  assert.deepEqual(runtimeErrors(PLAN), []);
  const bad = JSON.parse(JSON.stringify(PLAN));
  bad.cell.blocks.at(-1).sha256 = OLD.cell.blocks.at(-1).sha256;
  bad.word.patchVersion = 'euro-office-bridge-word-v1/9.3.5.0/3';
  assert.deepEqual(blockErrors(bad), ['cell:euro-office-bridge-v1/rev17-candidate/bridge-block.js']);
  assert.deepEqual(runtimeErrors(bad), ['word']);
});
