'use strict';

/**
 * M12 DS Bridge fresh install (documentserver-patches/9.3.4.60/install_bridge.py): one block inserted before one
 * anchor, accepted only when the result hash is the plan's output hash. Synthetic bundles; the real-bundle run is
 * docs/evidence/m12-2/ds-fresh-install-20261012.txt.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const { spawnSync } = require('node:child_process');
const crypto = require('node:crypto');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

const TOOL = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.4.60', 'install_bridge.py');
const made = [];
const tmp = () => { const d = fs.mkdtempSync(path.join(os.tmpdir(), 'eo-dsfi-test-')); made.push(d); return d; };
test.after(() => { for (const d of made) fs.rmSync(d, { recursive: true, force: true }); });
const sha = b => crypto.createHash('sha256').update(b).digest('hex');

function setup({ bundle = 'var a=1;ANCHOR();var b=2;', block = 'BRIDGE();', anchor = 'ANCHOR();' } = {}) {
  const d = tmp();
  fs.writeFileSync(path.join(d, 'block.js'), block);
  const out = bundle.replace(anchor, block + anchor);
  const plan = { cell: { input_sha256: sha(bundle), anchor, blocks: [{ file: 'block.js', sha256: sha(block) }], output_sha256: sha(out), patchVersion: 'test/1' } };
  fs.writeFileSync(path.join(d, 'plan.json'), JSON.stringify(plan));
  fs.writeFileSync(path.join(d, 'in.js'), bundle);
  return { d, out };
}
const run = (d, input, ...extra) => spawnSync('python3', [TOOL, '--target', 'cell', '--plan', path.join(d, 'plan.json'), '--input', input, ...extra], { encoding: 'utf8' });

test('PATCHABLE: --check writes nothing, --apply writes exactly the expected bytes', () => {
  const { d, out } = setup();
  const c = run(d, path.join(d, 'in.js'), '--check');
  assert.equal(c.status, 0, c.stderr); assert.match(c.stdout, /STATUS=PATCHABLE/);
  assert.ok(!fs.existsSync(path.join(d, 'o.js')));
  const a = run(d, path.join(d, 'in.js'), '--apply', '--output', path.join(d, 'o.js'));
  assert.equal(a.status, 0, a.stderr);
  assert.equal(fs.readFileSync(path.join(d, 'o.js'), 'utf8'), out);
  assert.deepEqual(fs.readdirSync(d).filter(f => f.startsWith('.install-bridge-')), []);
});

test('ALREADY_PATCHED is a no-op; UNKNOWN_INPUT, ANCHOR_NOT_UNIQUE, BLOCK_MISMATCH refuse and write nothing', () => {
  const { d, out } = setup();
  fs.writeFileSync(path.join(d, 'done.js'), out);
  const ap = run(d, path.join(d, 'done.js'), '--apply', '--output', path.join(d, 'x.js'));
  assert.equal(ap.status, 0); assert.match(ap.stdout, /STATUS=ALREADY_PATCHED/); assert.ok(!fs.existsSync(path.join(d, 'x.js')));
  fs.writeFileSync(path.join(d, 'other.js'), 'var a=1;ANCHOR();var b=3;');
  const un = run(d, path.join(d, 'other.js'), '--apply', '--output', path.join(d, 'x.js'));
  assert.equal(un.status, 2); assert.match(un.stdout, /STATUS=UNKNOWN_INPUT/); assert.ok(!fs.existsSync(path.join(d, 'x.js')));
  const two = setup({ bundle: 'ANCHOR();x;ANCHOR();' });
  const t = run(two.d, path.join(two.d, 'in.js'), '--check');
  assert.equal(t.status, 2); assert.match(t.stdout, /STATUS=ANCHOR_NOT_UNIQUE/);
  const bad = setup();
  fs.writeFileSync(path.join(bad.d, 'block.js'), 'TAMPERED();');
  const b = run(bad.d, path.join(bad.d, 'in.js'), '--check');
  assert.equal(b.status, 2); assert.match(b.stdout, /STATUS=BLOCK_MISMATCH/);
});

test('the shipped plan names real files with matching hashes; cell = own chart read layer + Bridge, no historical chart patch', () => {
  const dir = path.dirname(TOOL);
  const plan = JSON.parse(fs.readFileSync(path.join(dir, 'fresh-install.json'), 'utf8'));
  for (const t of ['cell', 'word']) {
    for (const b of plan[t].blocks) assert.equal(sha(fs.readFileSync(path.join(dir, b.file))), b.sha256, `${t}: ${b.file}`);
    for (const k of ['input_sha256', 'output_sha256']) assert.match(plan[t][k], /^[0-9a-f]{64}$/);
    for (const a of [plan[t].anchor, ...plan[t].blocks.map(b => b.anchor).filter(Boolean)]) {
      assert.ok(a.length > 0 && a.length < 60, `${t}: a short anchor, not a copied upstream passage`);
    }
  }
  // the chart read layer must sit at the end of the cell builder scope (after ApiChart.prototype is replaced), the Bridge at its own anchor
  assert.equal(plan.cell.blocks[0].anchor, 'AscBuilder.Cell.init=function(');
  assert.equal(plan.cell.blocks[1].anchor, undefined);
  assert.deepEqual(plan.cell.blocks.map(b => b.file.split('/')[0]), ['euro-office-chart-read-v1', 'euro-office-bridge-v1']);
  assert.doesNotMatch(JSON.stringify(plan), /chart-public-getters/);
});

test('two blocks are inserted in plan order before the one anchor', () => {
  const d = tmp();
  const bundle = 'x;ANCHOR();y;';
  fs.writeFileSync(path.join(d, 'a.js'), 'A();'); fs.writeFileSync(path.join(d, 'b.js'), 'B();');
  const out = 'x;A();B();ANCHOR();y;';
  fs.writeFileSync(path.join(d, 'plan.json'), JSON.stringify({ cell: { input_sha256: sha(bundle), anchor: 'ANCHOR();', blocks: [{ file: 'a.js', sha256: sha('A();') }, { file: 'b.js', sha256: sha('B();') }], output_sha256: sha(out), patchVersion: 'test/2' } }));
  fs.writeFileSync(path.join(d, 'in.js'), bundle);
  const r = run(d, path.join(d, 'in.js'), '--apply', '--output', path.join(d, 'o.js'));
  assert.equal(r.status, 0, r.stdout + r.stderr);
  assert.equal(fs.readFileSync(path.join(d, 'o.js'), 'utf8'), out);
});

test('a block with its own anchor goes before that anchor; a missing or repeated own anchor refuses', () => {
  const d = tmp();
  const bundle = 'x;FIRST();y;LAST();z;';
  fs.writeFileSync(path.join(d, 'a.js'), 'A();'); fs.writeFileSync(path.join(d, 'b.js'), 'B();');
  const out = 'x;B();FIRST();y;A();LAST();z;';
  const plan = anchorA => ({ cell: { input_sha256: sha(bundle), anchor: 'FIRST();', blocks: [{ file: 'a.js', sha256: sha('A();'), anchor: anchorA }, { file: 'b.js', sha256: sha('B();') }], output_sha256: sha(out), patchVersion: 'test/3' } });
  fs.writeFileSync(path.join(d, 'plan.json'), JSON.stringify(plan('LAST();')));
  fs.writeFileSync(path.join(d, 'in.js'), bundle);
  const r = run(d, path.join(d, 'in.js'), '--apply', '--output', path.join(d, 'o.js'));
  assert.equal(r.status, 0, r.stdout + r.stderr);
  assert.equal(fs.readFileSync(path.join(d, 'o.js'), 'utf8'), out);
  for (const bad of ['NOWHERE();', ';']) {
    fs.writeFileSync(path.join(d, 'plan.json'), JSON.stringify(plan(bad)));
    const x = run(d, path.join(d, 'in.js'), '--check');
    assert.equal(x.status, 2, bad); assert.match(x.stdout, /STATUS=ANCHOR_NOT_UNIQUE/, bad);
  }
});
