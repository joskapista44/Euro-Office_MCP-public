'use strict';

/**
 * DS v9.3.5 cell Bridge fresh install (documentserver-patches/9.3.5.0/fresh-install.json), using the EXISTING,
 * UNMODIFIED documentserver-patches/9.3.4.60/install_bridge.py via --plan (the tool has no hardcoded version
 * string). Synthetic bundles here, same as test/acceptance/ds-fresh-install.test.cjs; the real-9.3.5-bundle run is
 * docs/evidence/ds935-u03-fresh-install-20261018.txt, docs/DS-9.3.5-CELL-BRIDGE-U0-U3.md.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const { spawnSync } = require('node:child_process');
const crypto = require('node:crypto');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

const TOOL = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.4.60', 'install_bridge.py');
const PLAN_DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.5.0');
const PLAN = path.join(PLAN_DIR, 'fresh-install.json');
const made = [];
const tmp = () => { const d = fs.mkdtempSync(path.join(os.tmpdir(), 'eo-dsfi935-test-')); made.push(d); return d; };
test.after(() => { for (const d of made) fs.rmSync(d, { recursive: true, force: true }); });
const sha = b => crypto.createHash('sha256').update(b).digest('hex');
const run = (plan, input, ...extra) => spawnSync('python3', [TOOL, '--target', 'cell', '--plan', plan, '--input', input, ...extra], { encoding: 'utf8' });

test('the shipped 9.3.5.0 cell plan names a real file with a matching hash, distinct from 9.3.4.60', () => {
  const plan = JSON.parse(fs.readFileSync(PLAN, 'utf8'));
  assert.ok(plan.cell, 'cell key present');
  // ds935-word-bridge-20261018 added a "word" key to this SAME plan file -- this test still pins ONLY the cell
  // entry (unchanged by that unit); the word entry has its own test, test/acceptance/ds935-word-bridge-fresh-install.test.cjs
  assert.ok(plan.word, 'word key present (added by ds935-word-bridge-20261018, a later unit sharing this plan file)');
  for (const b of plan.cell.blocks) {
    assert.equal(sha(fs.readFileSync(path.join(PLAN_DIR, b.file))), b.sha256, b.file);
  }
  for (const k of ['input_sha256', 'output_sha256']) assert.match(plan.cell[k], /^[0-9a-f]{64}$/);
  assert.match(plan.cell.anchor, /GetSheets/);
  // SUPERSEDED by ds935-chart-read-20261018: the chart read layer joined the plan, same family order as 9.3.4.60
  // (chart-read at its own anchor, Bridge at the target anchor) -- see ds935-chart-read-fresh-install.test.cjs.
  assert.deepEqual(plan.cell.blocks.map(b => b.file.split('/')[0]), ['euro-office-chart-read-v1', 'euro-office-bridge-v1']);
  assert.equal(plan.cell.blocks[0].anchor, 'AscBuilder.Cell.init=function(');
  assert.equal(plan.cell.blocks[1].anchor, undefined);

  const oldPlan = JSON.parse(fs.readFileSync(path.join(PLAN_DIR, '..', '9.3.4.60', 'fresh-install.json'), 'utf8'));
  assert.notEqual(plan.cell.input_sha256, oldPlan.cell.input_sha256, '9.3.5 input must differ from the 9.3.4.60 one');
  assert.notEqual(plan.cell.output_sha256, oldPlan.cell.output_sha256);

  // the block content is the LIVE accepted rev16, with EXACTLY its own EO_BRIDGE_MANIFEST identity
  // (documentServerVersion / documentServerBuild / patchVersion) rewritten for 9.3.5 -- every other
  // byte, including every capability entry, stays byte-identical (ds935-cell-bridge-identity-runtime-20261018).
  const rev16 = fs.readFileSync(
    path.join(PLAN_DIR, '..', '9.3.4.60', 'euro-office-bridge-v1', 'rev16-candidate', 'bridge-block.js'), 'utf8');
  const shipped = fs.readFileSync(path.join(PLAN_DIR, 'euro-office-bridge-v1', 'bridge-block.js'), 'utf8');
  assert.notEqual(shipped, rev16, 'the identity fields must actually differ');
  const old3 = ['documentServerVersion: "9.3.4"', 'documentServerBuild: "9.3.4-9475e9f12a1905e2ada1fbc4bb87d01b"', 'patchVersion: "euro-office-bridge-v1/9.3.4.60/16"'];
  const new3 = ['documentServerVersion: "9.3.5"', 'documentServerBuild: "9.3.5-201f1ecbbedf03b95470dec61d9f30bd"', 'patchVersion: "euro-office-bridge-v1/9.3.5.0/16"'];
  for (const s of old3) assert.equal(rev16.includes(s) && rev16.split(s).length - 1, 1, `rev16 carries ${s} exactly once`);
  for (const s of new3) assert.equal(shipped.includes(s) && shipped.split(s).length - 1, 1, `shipped carries ${s} exactly once`);
  // replacing the 3 old literals in rev16 with the 3 new ones reproduces the shipped block EXACTLY -- i.e. the only
  // difference between rev16 and the shipped 9.3.5.0 block is those 3 substitutions, nothing else moved
  let reconstructed = rev16;
  for (let i = 0; i < old3.length; i += 1) reconstructed = reconstructed.replace(old3[i], new3[i]);
  assert.equal(reconstructed, shipped, 'rev16 with ONLY the 3 identity fields substituted must equal the shipped 9.3.5.0 block');
});

test('synthetic PATCHABLE/ALREADY_PATCHED/UNKNOWN_INPUT/BLOCK_MISMATCH against the 9.3.5.0 plan mechanics', () => {
  const d = tmp();
  const bundle = 'var a=1;Api.GetSheets=function(){var b=2;';
  const block = 'BRIDGE_9_3_5();';
  fs.writeFileSync(path.join(d, 'block.js'), block);
  const out = bundle.replace('Api.GetSheets=function(){', block + 'Api.GetSheets=function(){');
  const plan = { cell: { input_sha256: sha(bundle), anchor: 'Api.GetSheets=function(){', blocks: [{ file: 'block.js', sha256: sha(block) }], output_sha256: sha(out), patchVersion: 'euro-office-bridge-v1/9.3.5.0/synthetic' } };
  fs.writeFileSync(path.join(d, 'plan.json'), JSON.stringify(plan));
  fs.writeFileSync(path.join(d, 'in.js'), bundle);

  const c = run(path.join(d, 'plan.json'), path.join(d, 'in.js'), '--check');
  assert.equal(c.status, 0, c.stderr); assert.match(c.stdout, /STATUS=PATCHABLE/);

  const a = run(path.join(d, 'plan.json'), path.join(d, 'in.js'), '--apply', '--output', path.join(d, 'o.js'));
  assert.equal(a.status, 0, a.stderr);
  assert.equal(fs.readFileSync(path.join(d, 'o.js'), 'utf8'), out);
  assert.deepEqual(fs.readdirSync(d).filter(f => f.startsWith('.install-bridge-')), []);

  const again = run(path.join(d, 'plan.json'), path.join(d, 'o.js'), '--apply', '--output', path.join(d, 'x.js'));
  assert.equal(again.status, 0); assert.match(again.stdout, /STATUS=ALREADY_PATCHED/);
  assert.ok(!fs.existsSync(path.join(d, 'x.js')));

  fs.writeFileSync(path.join(d, 'other.js'), 'not the real bundle');
  const un = run(path.join(d, 'plan.json'), path.join(d, 'other.js'), '--apply', '--output', path.join(d, 'y.js'));
  assert.equal(un.status, 2); assert.match(un.stdout, /STATUS=UNKNOWN_INPUT/); assert.ok(!fs.existsSync(path.join(d, 'y.js')));

  fs.writeFileSync(path.join(d, 'block.js'), 'TAMPERED();');
  const bm = run(path.join(d, 'plan.json'), path.join(d, 'in.js'), '--check');
  assert.equal(bm.status, 2); assert.match(bm.stdout, /STATUS=BLOCK_MISMATCH/);
});
