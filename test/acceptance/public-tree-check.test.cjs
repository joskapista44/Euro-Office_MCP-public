'use strict';

/**
 * M12 public-repository CI gate (scripts/public-tree-check.cjs): the checkout must be byte-for-byte the derivation
 * recorded in PUBLIC-DERIVATION.json, with no symlink and no credential shape. Synthetic trees only.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const { spawnSync } = require('node:child_process');
const crypto = require('node:crypto');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

const TOOL = path.join(__dirname, '..', '..', 'scripts', 'public-tree-check.cjs');
const made = [];
test.after(() => { for (const d of made) fs.rmSync(d, { recursive: true, force: true }); });
const sha = b => crypto.createHash('sha256').update(b).digest('hex');

function tree() {
  const d = fs.mkdtempSync(path.join(os.tmpdir(), 'eo-ptc-test-'));
  made.push(d);
  const files = { 'README.md': ['# x\n', 0o644], 'bin/run.cjs': ['#!/usr/bin/env node\n', 0o755] };
  const ENTRIES = [];
  for (const [p, [text, mode]] of Object.entries(files)) {
    fs.mkdirSync(path.dirname(path.join(d, p)), { recursive: true });
    fs.writeFileSync(path.join(d, p), text); fs.chmodSync(path.join(d, p), mode);
    ENTRIES.push({ path: p, mode: mode.toString(8), sha256: sha(text) });
  }
  fs.writeFileSync(path.join(d, 'PUBLIC-DERIVATION.json'), JSON.stringify({ ENTRIES }));
  fs.mkdirSync(path.join(d, '.git')); fs.writeFileSync(path.join(d, '.git', 'HEAD'), 'ref\n');
  return d;
}
const run = d => spawnSync(process.execPath, [TOOL, '--root', d], { encoding: 'utf8' });

test('PASS: the checkout is exactly the recorded derivation (.git ignored)', () => {
  const r = run(tree());
  assert.equal(r.status, 0, r.stderr); assert.match(r.stdout, /PUBLIC_TREE_CHECK PASS/);
});

test('each deviation fails with its own named reason', () => {
  const cases = [
    ['EXTRA', d => fs.writeFileSync(path.join(d, 'new.txt'), 'x')],
    ['CHANGED', d => fs.appendFileSync(path.join(d, 'README.md'), 'edited')],
    ['MISSING', d => fs.rmSync(path.join(d, 'README.md'))],
    ['MODE', d => fs.chmodSync(path.join(d, 'bin/run.cjs'), 0o644)],
    ['SYMLINK', d => fs.symlinkSync('README.md', path.join(d, 'link.md'))],
    ['SECRET', d => fs.writeFileSync(path.join(d, 'leak.txt'), `${'gh'}p_${'A'.repeat(36)}`)],
    ['MANIFEST', d => fs.rmSync(path.join(d, 'PUBLIC-DERIVATION.json'))],
  ];
  for (const [reason, sabotage] of cases) {
    const d = tree(); sabotage(d);
    const r = run(d);
    assert.equal(r.status, 2, `${reason}: ${r.stdout}`);
    assert.match(r.stderr, new RegExp(`^${reason}: `, 'm'), reason);
  }
});
