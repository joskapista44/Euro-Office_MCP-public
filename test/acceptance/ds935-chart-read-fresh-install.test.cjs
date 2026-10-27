'use strict';

/**
 * DS v9.3.5 chart-read layer joins the cell plan (documentserver-patches/9.3.5.0/fresh-install.json), at its own
 * anchor, same family order as 9.3.4.60 (chart-read before the Bridge's target anchor -- ds935-chart-read-20261018).
 * Uses the EXISTING, UNMODIFIED documentserver-patches/9.3.4.60/install_bridge.py via --plan.
 * Real-9.3.5-bundle run: docs/DS-9.3.5-CHART-READ-20261018.md, docs/evidence/ds935-chart-read-fresh-install-20261018.txt.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const { spawnSync, execFileSync } = require('node:child_process');
const crypto = require('node:crypto');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

const TOOL = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.4.60', 'install_bridge.py');
const PLAN_DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.5.0');
const PLAN = path.join(PLAN_DIR, 'fresh-install.json');
const made = [];
const tmp = () => { const d = fs.mkdtempSync(path.join(os.tmpdir(), 'eo-dscr935-test-')); made.push(d); return d; };
test.after(() => { for (const d of made) fs.rmSync(d, { recursive: true, force: true }); });
const sha = b => crypto.createHash('sha256').update(b).digest('hex');
const run = (plan, input, ...extra) => spawnSync('python3', [TOOL, '--target', 'cell', '--plan', plan, '--input', input, ...extra], { encoding: 'utf8' });

test('the shipped 9.3.5.0 cell plan carries chart-read at its own anchor, Bridge at the target anchor, same order as 9.3.4.60', () => {
  const plan = JSON.parse(fs.readFileSync(PLAN, 'utf8'));
  assert.deepEqual(plan.cell.blocks.map(b => b.file.split('/')[0]), ['euro-office-chart-read-v1', 'euro-office-bridge-v1']);
  assert.equal(plan.cell.blocks[0].anchor, 'AscBuilder.Cell.init=function(');
  assert.equal(plan.cell.blocks[1].anchor, undefined);
  for (const b of plan.cell.blocks) assert.equal(sha(fs.readFileSync(path.join(PLAN_DIR, b.file))), b.sha256, b.file);

  // the 9.3.4.60 plan uses the SAME order and the SAME anchor for the same two families
  const oldPlan = JSON.parse(fs.readFileSync(path.join(PLAN_DIR, '..', '9.3.4.60', 'fresh-install.json'), 'utf8'));
  assert.deepEqual(oldPlan.cell.blocks.map(b => b.file.split('/')[0]), ['euro-office-chart-read-v1', 'euro-office-bridge-v1']);
  assert.equal(oldPlan.cell.blocks[0].anchor, 'AscBuilder.Cell.init=function(');
});

test('the shipped chart-read block is the live 9.3.4.60 one with ONLY its 2 version literals rewritten for 9.3.5.0', () => {
  const live = fs.readFileSync(
    path.join(PLAN_DIR, '..', '9.3.4.60', 'euro-office-chart-read-v1', 'chart-read-block.js'), 'utf8');
  const shipped = fs.readFileSync(path.join(PLAN_DIR, 'euro-office-chart-read-v1', 'chart-read-block.js'), 'utf8');
  assert.notEqual(shipped, live, 'the identity fields must actually differ');
  const old2 = ['euro-office-chart-read-v1/9.3.4.60/1', 'runtime: "9.3.4.60"'];
  const new2 = ['euro-office-chart-read-v1/9.3.5.0/1', 'runtime: "9.3.5.0"'];
  for (const s of old2) assert.equal(live.includes(s) && live.split(s).length - 1, 1, `live carries ${s} exactly once`);
  for (const s of new2) assert.equal(shipped.includes(s) && shipped.split(s).length - 1, 1, `shipped carries ${s} exactly once`);
  let reconstructed = live;
  for (let i = 0; i < old2.length; i += 1) reconstructed = reconstructed.replace(old2[i], new2[i]);
  assert.equal(reconstructed, shipped, 'live with ONLY the 2 identity literals substituted must equal the shipped 9.3.5.0 block');
});

// The identity accessor is searched in src/ AND test/static/ (a static gate pinning it would call it); the
// family name 'chart-read-v1' only in src/ -- test/static/ legitimately names the block PATH (e.g. the
// ds935-live-harness-hardcodes test), which is not an identity check.
const gitGrep = (pattern, paths) => {
  try {
    return execFileSync('git', ['grep', '-l', pattern, '--', ...paths], { cwd: path.join(__dirname, '..', '..'), encoding: 'utf8' });
  } catch (e) {
    return e.stdout || '';
  }
};

test('the product has no chart-read-layer identity check to extend -- accessor in src/+test/static/, family name in src/, zero hits', () => {
  assert.equal(gitGrep('GetEuroOfficeChartReadLayer', ['src/', 'test/static/']).trim(), '', 'no src/ or test/static/ file calls the chart-read layer identity accessor');
  assert.equal(gitGrep('chart-read-v1', ['src/']).trim(), '', 'no src/ file references the chart-read family -- nothing to extend for 9.3.5');
});

test('RED CONTROL: the same grep finds a tracked file that does name the family (the search is not blind)', () => {
  assert.notEqual(gitGrep('chart-read-v1', ['test/static/']).trim(), '', 'test/static/ names the block path somewhere');
});

test('synthetic PATCHABLE/ALREADY_PATCHED/UNKNOWN_INPUT/BLOCK_MISMATCH with chart-read at its own anchor + Bridge at the target anchor', () => {
  const d = tmp();
  const bundle = 'var a=1;AscBuilder.Cell.init=function(){var m=2;Api.GetSheets=function(){var b=3;';
  const chartBlock = 'CHARTREAD_9_3_5();';
  const bridgeBlock = 'BRIDGE_9_3_5();';
  fs.writeFileSync(path.join(d, 'chart.js'), chartBlock);
  fs.writeFileSync(path.join(d, 'bridge.js'), bridgeBlock);
  const out = bundle
    .replace('AscBuilder.Cell.init=function(', chartBlock + 'AscBuilder.Cell.init=function(')
    .replace('Api.GetSheets=function(){', bridgeBlock + 'Api.GetSheets=function(){');
  const plan = {
    cell: {
      input_sha256: sha(bundle),
      anchor: 'Api.GetSheets=function(){',
      blocks: [
        { file: 'chart.js', sha256: sha(chartBlock), anchor: 'AscBuilder.Cell.init=function(' },
        { file: 'bridge.js', sha256: sha(bridgeBlock) },
      ],
      output_sha256: sha(out),
      patchVersion: 'euro-office-bridge-v1/9.3.5.0/synthetic',
    },
  };
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

  fs.writeFileSync(path.join(d, 'chart.js'), 'TAMPERED();');
  const bm = run(path.join(d, 'plan.json'), path.join(d, 'in.js'), '--check');
  assert.equal(bm.status, 2); assert.match(bm.stdout, /STATUS=BLOCK_MISMATCH/);
});
