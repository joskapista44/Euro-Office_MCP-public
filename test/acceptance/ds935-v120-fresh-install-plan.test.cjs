'use strict';

/**
 * v1.2.0 (B3'): documentserver-patches/9.3.5.0/fresh-install-v1.2.0.json installs cell /29 (chart-read /1 +
 * GH #362 Print_Titles save-fix + GH #365 veryHidden R1-R3 + the rev29-veryhidden bridge block) and word /5
 * (unchanged since v1.1.0). Modelled on ds935-v110-fresh-install-plan.test.cjs.
 *
 * cell.output_sha256 was MEASURED (2026-11-07) by running this repo's own
 * plan-rev29-veryhidden.py against the real unpatched 9.3.5.0 cell sdk-all.js taken from the Document Server 9.3.5.0 image;
 * it equals the cell bundle served by the reference production deployment.
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const crypto = require('node:crypto');
const { ACCEPTED_RUNTIMES } = require('../../src/core/runtime-identity.cjs');

const DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.5.0');
const PLAN = JSON.parse(fs.readFileSync(path.join(DIR, 'fresh-install-v1.2.0.json'), 'utf8'));
const V110 = JSON.parse(fs.readFileSync(path.join(DIR, 'fresh-install-v1.1.0.json'), 'utf8'));
const sha = file => crypto.createHash('sha256').update(fs.readFileSync(path.join(DIR, file))).digest('hex');
const HEX64 = /^[0-9a-f]{64}$/;

test('the plan installs cell /29 and word /5 -- both are accepted runtimes, the last insert block carries the patchVersion', () => {
  assert.equal(PLAN.cell.patchVersion, 'euro-office-bridge-v1/9.3.5.0/29');
  assert.equal(PLAN.word.patchVersion, 'euro-office-bridge-word-v1/9.3.5.0/5');
  for (const [target, format] of [['cell', 'xlsx'], ['word', 'docx']]) {
    assert.ok(ACCEPTED_RUNTIMES[format].some(r => r.patchVersion === PLAN[target].patchVersion), `${target}: not an accepted runtime`);
  }
  const inserts = PLAN.cell.blocks.filter(b => 'file' in b);
  assert.equal(inserts.at(-1).layer, PLAN.cell.patchVersion);
});

test('every pinned INSERT block hash equals its canonical file (REPLACE steps carry no file)', () => {
  for (const target of ['cell', 'word']) {
    for (const b of PLAN[target].blocks) {
      if ('file' in b) assert.equal(sha(b.file), b.sha256, `${target}: ${b.file}`);
      else assert.ok(typeof b.old === 'string' && typeof b.new === 'string' && b.old !== b.new, `${target}: malformed replace step ${b._step || ''}`);
    }
  }
});

test('cell carries exactly the 4 named GH #362/#365 replace steps, in order, before the insert blocks', () => {
  const steps = PLAN.cell.blocks.filter(b => 'old' in b).map(b => b._step);
  assert.deepEqual(steps, ['xlnm-print-titles (GH #362)', 'veryhidden-R1 (GH #365)', 'veryhidden-R2 (GH #365)', 'veryhidden-R3 (GH #365)']);
  const firstInsertIndex = PLAN.cell.blocks.findIndex(b => 'file' in b);
  assert.equal(firstInsertIndex, 4, 'replace steps must run before any insert block, per install_bridge.py plan order');
});

test('the inputs are the same unpatched 9.3.5 bundles as the earlier plans; word output is pinned and unchanged from v1.1.0', () => {
  for (const target of ['cell', 'word']) {
    assert.equal(PLAN[target].input_sha256, V110[target].input_sha256, `${target}: input bundle must not change between releases`);
    assert.equal(PLAN[target].file, V110[target].file);
  }
  assert.equal(PLAN.word.output_sha256, V110.word.output_sha256, 'word Bridge has not advanced since v1.1.0 -- same output');
  assert.match(PLAN.word.output_sha256, HEX64);
});

test('cell.output_sha256 is measured (not fabricated), differs from the v1.1.0 output, and equals the production-served cell', () => {
  assert.match(PLAN.cell.output_sha256, HEX64);
  assert.notEqual(PLAN.cell.output_sha256, V110.cell.output_sha256, 'the /29 output must differ from the /17 output');
  assert.equal(PLAN.cell.output_sha256, '75da09a457559333c75583a92c267d62b5e0b0c7c3eb7e26b18498411ed49f97', 'must equal the production-served cell');
  assert.ok(!('_pending' in PLAN.cell), 'the pending marker must be removed once the real hash lands');
});

/** The two checks above, as functions over any plan, so the RED control runs the same logic on a broken copy. */
const blockErrors = plan => ['cell', 'word'].flatMap(t => plan[t].blocks.filter(b => 'file' in b && sha(b.file) !== b.sha256).map(b => `${t}:${b.file}`));
const runtimeErrors = plan => [['cell', 'xlsx'], ['word', 'docx']]
  .filter(([t, f]) => !ACCEPTED_RUNTIMES[f].some(r => r.patchVersion === plan[t].patchVersion)).map(([t]) => t);

test('RED CONTROL: the same checks fail on a copy with a wrong block hash, or with a patchVersion no runtime accepts', () => {
  assert.deepEqual(blockErrors(PLAN), []);
  assert.deepEqual(runtimeErrors(PLAN), []);
  const bad = JSON.parse(JSON.stringify(PLAN));
  bad.cell.blocks.at(-1).sha256 = V110.cell.blocks.at(-1).sha256;
  bad.word.patchVersion = 'euro-office-bridge-word-v1/9.3.5.0/3';
  assert.deepEqual(blockErrors(bad), ['cell:euro-office-bridge-v1/rev29-veryhidden-candidate/bridge-block.js']);
  assert.deepEqual(runtimeErrors(bad), ['word']);
});
