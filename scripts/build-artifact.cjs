#!/usr/bin/env node
'use strict';

/**
 * Euro-Office_MCP artifact builder (owner decision, owner channel 2026-10-08, point 15).
 *
 *   node scripts/build-artifact.cjs --rc <commit> --out <dir> [--repo <dir>]
 *
 * RC_SHA -> clean `git archive` -> `npm ci` (lockfile) -> artifact tree (runtime files only) -> release.json ->
 * deterministic tgz -> SHA256. Refuses a dirty working tree (an RC is built only from committed state) and a missing
 * lockfile. Output: <out>/euro-office-mcp-<ARTIFACT_VERSION>.tgz, .tgz.sha256, .manifest.json (the manifest carries
 * ARTIFACT_SHA256; the installer later places it next to the unpacked tree as artifact.json -- a tarball cannot hold
 * its own hash). Nothing here deploys, installs or switches a symlink.
 */

const { execFileSync } = require('node:child_process');
const crypto = require('node:crypto');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

const PRODUCT = 'Euro-Office_MCP';
const RUNTIME_PATHS = ['bin', 'src', 'config', 'package.json', 'package-lock.json'];

function fail(code, msg) { process.stderr.write(`build-artifact: ${msg}\n`); process.exit(code); }

/** Every symlink in the tree whose target is absolute or resolves outside the tree (non-reproducible, dangling once installed). */
function symlinksEscaping(root) {
  const out = [];
  const walk = dir => {
    for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
      const p = path.join(dir, e.name);
      if (e.isSymbolicLink()) {
        const target = fs.readlinkSync(p);
        const rel = path.relative(root, path.resolve(dir, target));
        if (path.isAbsolute(target) || rel.startsWith('..') || path.isAbsolute(rel)) out.push(path.relative(root, p));
      } else if (e.isDirectory()) walk(p);
    }
  };
  walk(root);
  return out;
}

function parseArgs(argv) {
  const a = { repo: process.cwd() };
  for (let i = 0; i < argv.length; i += 1) {
    if (argv[i] === '--rc') a.rc = argv[++i];
    else if (argv[i] === '--out') a.out = argv[++i];
    else if (argv[i] === '--repo') a.repo = argv[++i];
    else fail(64, `unknown argument: ${argv[i]}`);
  }
  if (!a.rc || !a.out) fail(64, 'usage: build-artifact.cjs --rc <commit> --out <dir> [--repo <dir>]');
  return a;
}

const git = (repo, args, opts = {}) => execFileSync('git', ['-C', repo, ...args], { encoding: opts.encoding === null ? null : 'utf8', maxBuffer: 512 * 1024 * 1024 });
const sha256 = buf => crypto.createHash('sha256').update(buf).digest('hex');

function main() {
  const a = parseArgs(process.argv.slice(2));
  const repo = path.resolve(a.repo);
  const dirty = git(repo, ['status', '--porcelain']).trim();
  if (dirty) fail(2, `DIRTY_TREE: refusing to build an RC artifact from uncommitted state:\n${dirty.split('\n').slice(0, 10).join('\n')}`);
  let sha;
  try { sha = git(repo, ['rev-parse', '--verify', `${a.rc}^{commit}`]).trim(); } catch { fail(2, `RC_NOT_FOUND: ${a.rc}`); }
  const commitEpoch = git(repo, ['show', '-s', '--format=%ct', sha]).trim();

  const work = fs.mkdtempSync(path.join(os.tmpdir(), 'eo-artifact-'));
  // a refusal inside the build exits the process -- remove the temp dir first (fail() never returns, so finally would not run)
  const failClean = (code, msg) => { fs.rmSync(work, { recursive: true, force: true }); fail(code, msg); };
  try {
    const src = path.join(work, 'src');
    fs.mkdirSync(src);
    fs.writeFileSync(path.join(work, 'source.tar'), git(repo, ['archive', '--format=tar', sha], { encoding: null }));
    execFileSync('tar', ['-xf', path.join(work, 'source.tar'), '-C', src]);
    if (!fs.existsSync(path.join(src, 'package-lock.json'))) failClean(2, 'NO_LOCKFILE: package-lock.json is required (npm ci)');
    const pkg = JSON.parse(fs.readFileSync(path.join(src, 'package.json'), 'utf8'));
    // --no-bin-links: the product never runs a dependency CLI, and a bin link is the one symlink npm writes into the tree
    execFileSync('npm', ['ci', '--omit=dev', '--ignore-scripts', '--no-audit', '--no-fund', '--no-bin-links'], { cwd: src, stdio: ['ignore', 'ignore', 'inherit'] });

    const version = `${pkg.version}-${sha.slice(0, 12)}`;
    const name = `euro-office-mcp-${version}`;
    const tree = path.join(work, 'stage', name);
    fs.mkdirSync(tree, { recursive: true });
    // verbatimSymlinks: cpSync would otherwise rewrite a relative link into an absolute one pointing into this temp dir
    for (const p of [...RUNTIME_PATHS, 'node_modules']) if (fs.existsSync(path.join(src, p))) fs.cpSync(path.join(src, p), path.join(tree, p), { recursive: true, verbatimSymlinks: true });
    const escaping = symlinksEscaping(tree);
    if (escaping.length) failClean(2, `SYMLINK_ESCAPES_TREE: ${escaping.slice(0, 5).join(', ')}`);
    const release = {
      PRODUCT, SOURCE_COMMIT: sha, ARTIFACT_VERSION: version, PACKAGE_VERSION: pkg.version,
      LOCKFILE_SHA256: sha256(fs.readFileSync(path.join(src, 'package-lock.json'))),
      NODE_ENGINES: (pkg.engines && pkg.engines.node) || null, SOURCE_COMMIT_EPOCH: Number(commitEpoch),
    };
    fs.writeFileSync(path.join(tree, 'release.json'), `${JSON.stringify(release, null, 2)}\n`);

    // deterministic archive: sorted entries, the commit time as mtime, numeric root ownership, gzip without a name/time
    fs.mkdirSync(path.resolve(a.out), { recursive: true });
    const tgz = path.join(path.resolve(a.out), `${name}.tgz`);
    const tar = execFileSync('tar', ['--sort=name', `--mtime=@${commitEpoch}`, '--owner=0', '--group=0', '--numeric-owner',
      '--format=gnu', '-cf', '-', '-C', path.join(work, 'stage'), name], { maxBuffer: 512 * 1024 * 1024 });
    fs.writeFileSync(tgz, execFileSync('gzip', ['-n', '-9'], { input: tar, maxBuffer: 512 * 1024 * 1024 }));
    const artifactSha = sha256(fs.readFileSync(tgz));
    fs.writeFileSync(`${tgz}.sha256`, `${artifactSha}  ${path.basename(tgz)}\n`);
    const manifest = { ...release, ARTIFACT: path.basename(tgz), ARTIFACT_SHA256: artifactSha };
    fs.writeFileSync(path.join(path.resolve(a.out), `${name}.manifest.json`), `${JSON.stringify(manifest, null, 2)}\n`);
    process.stdout.write([`SOURCE_SHA=${sha}`, `ARTIFACT=${path.basename(tgz)}`, `ARTIFACT_VERSION=${version}`, `ARTIFACT_SHA256=${artifactSha}`].join('\n') + '\n');
  } finally {
    fs.rmSync(work, { recursive: true, force: true });
  }
}

main();
