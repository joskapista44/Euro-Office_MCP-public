'use strict';

/**
 * M12.2 public-tree derivation (scripts/derive-public.cjs): RC -> git archive -> allowlist -> redaction -> fail-closed
 * gates -> PUBLIC-DERIVATION.json. Run on throw-away repositories with a synthetic policy, offline.
 *
 * The forbidden and secret samples are assembled at run time (and the synthetic policy forbids a made-up marker), so
 * this file itself passes the real public gates it is published through.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const { execFileSync, spawnSync } = require('node:child_process');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

const DERIVE = path.join(__dirname, '..', '..', 'scripts', 'derive-public.cjs');
const made = [];
const tmp = () => { const d = fs.mkdtempSync(path.join(os.tmpdir(), 'eo-derive-test-')); made.push(d); return d; };
test.after(() => { for (const d of made) fs.rmSync(d, { recursive: true, force: true }); });
const git = (cwd, ...a) => execFileSync('git', a, { cwd, encoding: 'utf8', env: { ...process.env, GIT_AUTHOR_DATE: '2026-10-12T10:00:00Z', GIT_COMMITTER_DATE: '2026-10-12T10:00:00Z' } });

const MARK = ['acme', 'internal'].join('-'); // the synthetic "internal reference"
const SECRET = ['gh', 'p_', 'a'.repeat(36)].join('');

function policy(over = {}) {
  return {
    include: ['src/', 'docs/', 'payload/', 'package.json'],
    exclude: [{ path: 'docs/private/', reason: 'test' }],
    map: [],
    runtimeVerbatim: ['src/', 'package.json'],
    redactions: [{ id: 'mark-ticket', pattern: `${MARK} #\\d+`, flags: 'g', replace: 'a ticket' }],
    pinnedVerbatim: [{ paths: ['payload/'], names: ['block.js'], tolerated: [`${MARK} #1`], reason: 'test' }],
    forbidden: [{ id: 'mark', pattern: `\\b${MARK}\\b` }],
    secretShapes: [{ id: 'github-token', pattern: 'gh[pousr]_[A-Za-z0-9]{30,}' }],
    binaryAllowlist: [],
    licences: [{ path: 'src/', spdx: 'Apache-2.0' }, { path: 'docs/', spdx: 'Apache-2.0' }, { path: 'payload/', spdx: 'Apache-2.0' }, { path: 'package.json', spdx: 'Apache-2.0' }],
    ...over,
  };
}

/** A committed repo: files = { relpath: content | {symlink} }, plus the policy at release/public/policy.json. */
function repoWith(files, pol = policy()) {
  const r = tmp();
  git(r, 'init', '-q');
  git(r, 'config', 'user.email', 't@test.invalid'); git(r, 'config', 'user.name', 't');
  const all = { 'release/public/policy.json': JSON.stringify(pol, null, 2), 'package.json': '{"name":"x","version":"1.0.0"}\n',
    'src/a.cjs': 'module.exports = 1;\n', 'docs/guide.md': `see ${MARK} #12 for details\n`, 'docs/private/notes.md': `${MARK} only here\n`,
    'payload/block.js': `var x = "${MARK} #1";\n`, ...files };
  for (const [p, c] of Object.entries(all)) {
    if (c === undefined) continue;
    fs.mkdirSync(path.dirname(path.join(r, p)), { recursive: true });
    if (c && c.symlink) fs.symlinkSync(c.symlink, path.join(r, p));
    else fs.writeFileSync(path.join(r, p), c);
  }
  git(r, 'add', '-A'); git(r, 'commit', '-q', '-m', 'rc');
  return { repo: r, sha: git(r, 'rev-parse', 'HEAD').trim() };
}

const derive = (repo, out) => spawnSync(process.execPath, [DERIVE, '--rc', 'HEAD', '--out', out, '--repo', repo], { encoding: 'utf8' });
const kv = s => Object.fromEntries(s.trim().split('\n').map(l => [l.slice(0, l.indexOf('=')), l.slice(l.indexOf('=') + 1)]));
const expectRefused = (r, out, reason) => {
  assert.equal(r.status, 2, `${reason}: expected a refusal, got ${r.status}\n${r.stdout}${r.stderr}`);
  assert.match(r.stderr, new RegExp(reason));
  assert.ok(!fs.existsSync(out) || fs.readdirSync(out).length === 0, `${reason}: nothing may be written on a refusal`);
};

test('GREEN: allowlist, redaction outside the runtime tree, pinned payload verbatim, manifest; deterministic', () => {
  const { repo, sha } = repoWith({});
  const a = tmp(); const b = tmp();
  const ra = derive(repo, path.join(a, 'pub'));
  assert.equal(ra.status, 0, ra.stderr);
  const rb = derive(repo, path.join(b, 'pub'));
  assert.equal(rb.status, 0, rb.stderr);
  assert.equal(kv(ra.stdout).PUBLIC_TREE_SHA256, kv(rb.stdout).PUBLIC_TREE_SHA256, 'two derivations of one RC are identical');
  const pub = path.join(a, 'pub');
  assert.equal(fs.readFileSync(path.join(pub, 'docs/guide.md'), 'utf8'), 'see a ticket for details\n');
  assert.ok(!fs.existsSync(path.join(pub, 'docs/private')), 'an excluded path is not published');
  assert.ok(!fs.existsSync(path.join(pub, 'release')), 'a path outside the allowlist is not published');
  assert.equal(fs.readFileSync(path.join(pub, 'payload/block.js'), 'utf8'), `var x = "${MARK} #1";\n`, 'pinned bytes unchanged');
  const m = JSON.parse(fs.readFileSync(path.join(pub, 'PUBLIC-DERIVATION.json'), 'utf8'));
  assert.equal(m.SOURCE_COMMIT, sha);
  assert.equal(m.RULE_HITS['mark-ticket'], 1);
  assert.deepEqual(m.ENTRIES.filter(e => e.transformed).map(e => e.path), ['docs/guide.md']);
  assert.ok(m.ENTRIES.every(e => e.licence === 'Apache-2.0'));
  for (const g of Object.values(m.GATES)) assert.equal(g, 'PASS');
});

test('RED: a forbidden reference no rule redacts', () => {
  const { repo } = repoWith({ 'docs/leak.md': `owner is ${MARK}\n` });
  const out = path.join(tmp(), 'pub');
  expectRefused(derive(repo, out), out, 'FORBIDDEN: docs/leak.md');
});

test('RED: the runtime tree is never rewritten -- a reference a rule COULD redact still fails there', () => {
  const { repo } = repoWith({ 'src/a.cjs': `// see ${MARK} #7\nmodule.exports = 1;\n` });
  const out = path.join(tmp(), 'pub');
  expectRefused(derive(repo, out), out, 'FORBIDDEN: src/a.cjs');
});

test('RED: a credential-shaped string', () => {
  const { repo } = repoWith({ 'docs/setup.md': `token ${SECRET}\n` });
  const out = path.join(tmp(), 'pub');
  expectRefused(derive(repo, out), out, 'SECRET: docs/setup.md');
});

test('RED: a symlink inside the published tree', () => {
  const { repo } = repoWith({ 'docs/link.md': { symlink: '../package.json' } });
  const out = path.join(tmp(), 'pub');
  expectRefused(derive(repo, out), out, 'SYMLINK: docs/link.md');
});

test('RED: a binary that is not allowlisted, and one whose hash differs from the allowlist', () => {
  const bin = Buffer.from([0x50, 0x4b, 0x00, 0x01]);
  const { repo } = repoWith({ 'docs/fixture.bin': bin });
  const out = path.join(tmp(), 'pub');
  expectRefused(derive(repo, out), out, 'BINARY: not allowlisted: docs/fixture.bin');
  const pinned = repoWith({ 'docs/fixture.bin': bin }, policy({ binaryAllowlist: [{ path: 'docs/fixture.bin', sha256: '0'.repeat(64) }] }));
  const out2 = path.join(tmp(), 'pub');
  expectRefused(derive(pinned.repo, out2), out2, 'BINARY: sha256 mismatch: docs/fixture.bin');
});

test('RED: a file without a licence mapping, and an undecided licence', () => {
  const { repo } = repoWith({}, policy({ licences: [{ path: 'src/', spdx: 'Apache-2.0' }] }));
  const out = path.join(tmp(), 'pub');
  expectRefused(derive(repo, out), out, 'LICENCE: no licence mapping');
  const pending = repoWith({}, policy({ licences: policy().licences.map(l => l.path === 'docs/' ? { ...l, spdx: 'PENDING OWNER DECISION' } : l) }));
  const out2 = path.join(tmp(), 'pub');
  expectRefused(derive(pending.repo, out2), out2, 'LICENCE: undecided licence');
  const twice = repoWith({}, policy({ licences: [...policy().licences, { path: 'docs/', spdx: 'MIT' }] }));
  const out3 = path.join(tmp(), 'pub');
  expectRefused(derive(twice.repo, out3), out3, 'POLICY_INVALID: two licence entries for docs/');
});

test('RED: a pinned file tolerates only its listed tokens', () => {
  const { repo } = repoWith({ 'payload/block.js': `var x = "${MARK} #1"; var y = "${MARK} #2";\n` });
  const out = path.join(tmp(), 'pub');
  expectRefused(derive(repo, out), out, 'FORBIDDEN: payload/block.js');
});

test('RED: a file byte-identical to a never-publish file, wherever it is placed', () => {
  const secretDoc = 'historical text that must never be published\n';
  const sha = require('node:crypto').createHash('sha256').update(secretDoc).digest('hex');
  const { repo } = repoWith({ 'docs/innocent-name.md': secretDoc }, policy({ forbiddenFiles: [{ sha256: sha, reason: 'test' }] }));
  const out = path.join(tmp(), 'pub');
  expectRefused(derive(repo, out), out, 'FORBIDDEN_FILE: docs/innocent-name.md');
});

test('RED: a policy entry that matches nothing, a missing policy, a non-empty output', () => {
  const { repo } = repoWith({}, policy({ include: [...policy().include, 'nothing-here/'] }));
  const out = path.join(tmp(), 'pub');
  expectRefused(derive(repo, out), out, 'POLICY: include matches nothing');
  const bare = repoWith({ 'release/public/policy.json': undefined });
  const out2 = path.join(tmp(), 'pub');
  expectRefused(derive(bare.repo, out2), out2, 'NO_POLICY');
  const full = tmp(); fs.writeFileSync(path.join(full, 'x'), 'x');
  const r = derive(repoWith({}).repo, full);
  assert.equal(r.status, 2); assert.match(r.stderr, /OUT_NOT_EMPTY/);
});

// ---- U2 (internal tracker): coverage of the `map` branch of publicPathFor()/derive() ---------------------------------
// The real policy.json's own map rule (release/public/root/ -> "") was previously untested here.

test('U2a: a map rule {from: "dir/", to: ""} maps that whole directory onto the public root', () => {
  const pol = policy({
    map: [{ from: 'pubroot/', to: '' }],
    include: [...policy().include, 'pubroot/'],
    licences: [...policy().licences, { path: 'index.md', spdx: 'Apache-2.0' }],
  });
  const { repo } = repoWith({ 'pubroot/index.md': 'hello root\n' }, pol);
  const out = path.join(tmp(), 'pub');
  const r = derive(repo, out);
  assert.equal(r.status, 0, r.stderr);
  assert.equal(fs.readFileSync(path.join(out, 'index.md'), 'utf8'), 'hello root\n', 'mapped to the public root, not under pubroot/');
  assert.ok(!fs.existsSync(path.join(out, 'pubroot')), 'the source directory name itself is not published');
});

test('U2b: the licence is looked up by the PUBLIC (mapped) path, not the private source path', () => {
  const pol = policy({
    map: [{ from: 'internal-src/', to: 'published/' }],
    include: [...policy().include, 'internal-src/'],
    licences: [...policy().licences, { path: 'published/', spdx: 'Apache-2.0' }],
  });
  assert.equal(pol.licences.some((l) => l.path.startsWith('internal-src')), false, 'sanity: no licence entry names the private source path at all');
  const { repo } = repoWith({ 'internal-src/a.md': 'content\n' }, pol);
  const out = path.join(tmp(), 'pub');
  const r = derive(repo, out);
  assert.equal(r.status, 0, r.stderr);
  const manifest = JSON.parse(fs.readFileSync(path.join(out, 'PUBLIC-DERIVATION.json'), 'utf8'));
  assert.equal(manifest.ENTRIES.find((e) => e.path === 'published/a.md').licence, 'Apache-2.0');
});

test('U2c: a map rule whose source matches no file anywhere -> a named POLICY gate error', () => {
  const pol = policy({ map: [{ from: 'ghost-dir/', to: '' }] });
  const { repo } = repoWith({}, pol);
  const out = path.join(tmp(), 'pub');
  expectRefused(derive(repo, out), out, 'POLICY: map source missing: ghost-dir/');
});

test('U2d: two different sources (one mapped, one not) resolving to the same public path -> a named POLICY gate error', () => {
  const pol = policy({ map: [{ from: 'legacy/dup.md', to: 'docs/guide.md' }] });
  const { repo } = repoWith({ 'legacy/dup.md': 'dup content\n' }, pol); // 'docs/guide.md' already exists via repoWith's own base fixture
  const out = path.join(tmp(), 'pub');
  expectRefused(derive(repo, out), out, 'POLICY: two sources map to docs/guide.md');
});
