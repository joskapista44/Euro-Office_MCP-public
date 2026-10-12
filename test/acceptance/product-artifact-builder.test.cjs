'use strict';

/**
 * Artifact builder (owner decision, owner channel 2026-10-08, point 15): RC_SHA -> clean git archive -> npm ci ->
 * artifact tree -> release.json -> tgz -> SHA256. Run on a throw-away repository, offline (no dependency to fetch).
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const { execFileSync, spawnSync } = require('node:child_process');
const crypto = require('node:crypto');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

const BUILDER = path.join(__dirname, '..', '..', 'scripts', 'build-artifact.cjs');
// every temp dir is removed when the file's tests end (the suite runs often; /tmp is shared by the fleet)
const made = [];
const tmp = () => { const d = fs.mkdtempSync(path.join(os.tmpdir(), 'eo-build-test-')); made.push(d); return d; };
test.after(() => { for (const d of made) fs.rmSync(d, { recursive: true, force: true }); });
const git = (cwd, ...a) => execFileSync('git', a, { cwd, encoding: 'utf8', env: { ...process.env, GIT_AUTHOR_DATE: '2026-10-08T10:00:00Z', GIT_COMMITTER_DATE: '2026-10-08T10:00:00Z' } });

function fixtureRepo({ lockfile = true } = {}) {
  const r = tmp();
  git(r, 'init', '-q');
  git(r, 'config', 'user.email', 't@test.invalid'); git(r, 'config', 'user.name', 't');
  const pkg = { name: 'euro-office-mcp', version: '0.1.0-test', private: true, engines: { node: '>=20' } };
  fs.writeFileSync(path.join(r, 'package.json'), JSON.stringify(pkg, null, 2));
  if (lockfile) fs.writeFileSync(path.join(r, 'package-lock.json'), JSON.stringify({ name: pkg.name, version: pkg.version, lockfileVersion: 3, requires: true, packages: { '': { name: pkg.name, version: pkg.version } } }, null, 2));
  for (const d of ['bin', 'src', 'test', 'docs']) fs.mkdirSync(path.join(r, d));
  fs.writeFileSync(path.join(r, 'bin', 'euro-office-mcp.cjs'), '#!/usr/bin/env node\n');
  fs.writeFileSync(path.join(r, 'src', 'x.cjs'), 'module.exports = 1;\n');
  fs.writeFileSync(path.join(r, 'test', 'x.test.cjs'), '');
  fs.writeFileSync(path.join(r, 'docs', 'big.md'), 'doc');
  git(r, 'add', '-A'); git(r, 'commit', '-q', '-m', 'rc');
  return { repo: r, sha: git(r, 'rev-parse', 'HEAD').trim() };
}

const build = (repo, rc, out) => spawnSync(process.execPath, [BUILDER, '--rc', rc, '--out', out, '--repo', repo], { encoding: 'utf8' });
const kv = s => Object.fromEntries(s.trim().split('\n').map(l => [l.slice(0, l.indexOf('=')), l.slice(l.indexOf('=') + 1)]));

test('builds from a committed RC: release.json, runtime files only, SHA256 recorded; reproducible', () => {
  const { repo, sha } = fixtureRepo();
  const out1 = tmp(); const out2 = tmp();
  const a = build(repo, sha, out1); const b = build(repo, sha, out2);
  assert.equal(a.status, 0, a.stderr); assert.equal(b.status, 0, b.stderr);
  const ka = kv(a.stdout);
  assert.equal(ka.SOURCE_SHA, sha);
  assert.equal(ka.ARTIFACT_VERSION, `0.1.0-test-${sha.slice(0, 12)}`);
  const tgz = path.join(out1, ka.ARTIFACT);
  assert.equal(crypto.createHash('sha256').update(fs.readFileSync(tgz)).digest('hex'), ka.ARTIFACT_SHA256);
  assert.equal(kv(b.stdout).ARTIFACT_SHA256, ka.ARTIFACT_SHA256, 'same RC -> byte-identical artifact');
  const manifest = JSON.parse(fs.readFileSync(path.join(out1, `euro-office-mcp-${ka.ARTIFACT_VERSION}.manifest.json`), 'utf8'));
  assert.equal(manifest.ARTIFACT_SHA256, ka.ARTIFACT_SHA256);
  const list = execFileSync('tar', ['-tzf', tgz], { encoding: 'utf8' }).split('\n').filter(Boolean);
  const top = `euro-office-mcp-${ka.ARTIFACT_VERSION}/`;
  assert.ok(list.includes(`${top}release.json`) && list.includes(`${top}bin/euro-office-mcp.cjs`) && list.includes(`${top}src/x.cjs`));
  assert.equal(list.some(f => f.startsWith(`${top}test/`) || f.startsWith(`${top}docs/`)), false, 'runtime files only');
  const rel = JSON.parse(execFileSync('tar', ['-xzOf', tgz, `${top}release.json`], { encoding: 'utf8' }));
  assert.deepEqual([rel.PRODUCT, rel.SOURCE_COMMIT, rel.ARTIFACT_VERSION], ['Euro-Office_MCP', sha, ka.ARTIFACT_VERSION]);
  assert.match(rel.LOCKFILE_SHA256, /^[0-9a-f]{64}$/);
  assert.equal('ARTIFACT_SHA256' in rel, false, 'a tarball cannot carry its own hash');
});

test('refuses a dirty tree, an unknown RC and a missing lockfile', () => {
  const { repo, sha } = fixtureRepo();
  fs.writeFileSync(path.join(repo, 'src', 'x.cjs'), 'module.exports = 2;\n');
  const dirty = build(repo, sha, tmp());
  assert.equal(dirty.status, 2); assert.match(dirty.stderr, /DIRTY_TREE/);
  const untracked = fixtureRepo();
  fs.writeFileSync(path.join(untracked.repo, 'stray.txt'), 'x');
  assert.match(build(untracked.repo, untracked.sha, tmp()).stderr, /DIRTY_TREE/);
  const clean = fixtureRepo();
  const unknown = build(clean.repo, 'deadbeef', tmp());
  assert.equal(unknown.status, 2); assert.match(unknown.stderr, /RC_NOT_FOUND/);
  const nolock = fixtureRepo({ lockfile: false });
  const nl = build(nolock.repo, nolock.sha, tmp());
  assert.equal(nl.status, 2); assert.match(nl.stderr, /NO_LOCKFILE/);
});

/** A fixture RC with ONE real dependency that ships a CLI (a local tarball, offline): the shape npm writes bin links for. */
function fixtureRepoWithBinDependency() {
  const { repo } = fixtureRepo();
  const dep = tmp();
  fs.writeFileSync(path.join(dep, 'package.json'), JSON.stringify({ name: 'dep-with-bin', version: '1.0.0', bin: { 'dep-with-bin': 'cli.js' } }));
  fs.writeFileSync(path.join(dep, 'cli.js'), '#!/usr/bin/env node\n');
  fs.writeFileSync(path.join(dep, 'index.js'), 'module.exports = 1;\n');
  const packed = execFileSync('npm', ['pack', '--silent', '--pack-destination', repo], { cwd: dep, encoding: 'utf8' }).trim().split('\n').pop();
  execFileSync('npm', ['install', '--package-lock-only', '--offline', '--ignore-scripts', '--no-audit', '--no-fund', `./${packed}`], { cwd: repo, stdio: 'ignore' });
  git(repo, 'add', '-A'); git(repo, 'commit', '-q', '-m', 'dependency with a bin');
  return { repo, sha: git(repo, 'rev-parse', 'HEAD').trim() };
}

test('a dependency with a CLI: two builds are byte-identical and no symlink in the artifact points outside it (M11.3b.2 defect)', () => {
  const { repo, sha } = fixtureRepoWithBinDependency();
  const out1 = tmp(); const out2 = tmp();
  const a = build(repo, sha, out1); const b = build(repo, sha, out2);
  assert.equal(a.status, 0, a.stderr); assert.equal(b.status, 0, b.stderr);
  assert.equal(kv(b.stdout).ARTIFACT_SHA256, kv(a.stdout).ARTIFACT_SHA256, 'BYTE_IDENTICAL with a real dependency');
  const tgz = path.join(out1, kv(a.stdout).ARTIFACT);
  const entries = execFileSync('tar', ['-tzvf', tgz], { encoding: 'utf8' }).split('\n').filter(Boolean);
  assert.ok(entries.some(l => l.includes('node_modules/dep-with-bin/index.js')), 'the dependency is in the artifact');
  assert.deepEqual(entries.filter(l => l.startsWith('l')), [], 'no symlink in the artifact (no bin links)');
});

test('SYMLINK_ESCAPES_TREE: a committed symlink pointing outside the artifact tree refuses the build (NEG control)', () => {
  const { repo } = fixtureRepo();
  fs.symlinkSync('/etc/hostname', path.join(repo, 'src', 'escape.cjs'));
  git(repo, 'add', '-A'); git(repo, 'commit', '-q', '-m', 'escaping link');
  const r = build(repo, git(repo, 'rev-parse', 'HEAD').trim(), tmp());
  assert.equal(r.status, 2);
  assert.match(r.stderr, /SYMLINK_ESCAPES_TREE: src\/escape\.cjs/);
});
