#!/usr/bin/env node
'use strict';

/**
 * Euro-Office_MCP public-tree derivation (M12.2; owner decisions D1-D6, owner channel 2026-10-12).
 *
 *   node scripts/derive-public.cjs --rc <commit> --out <dir> [--repo <dir>]
 *
 * The private canonical source is the only development authority. This script turns ONE committed RC into the public
 * tree: clean `git archive` -> path selection (policy allowlist) -> deterministic redaction -> fail-closed gates ->
 * public-derivation.json. The policy is read from the RC itself (release/public/policy.json), so the output is a pure
 * function of the RC: the same commit always derives the same bytes.
 *
 * The RUNTIME tree (policy.runtimeVerbatim: bin/, src/, config/, package files) is published byte-identical to the RC:
 * no redaction rule touches it, so the public artifact runs exactly the canonical code. A forbidden match there is a
 * canonical-source defect and fails the derivation; it is fixed in the canonical source, never rewritten here.
 *
 * Gates (any failure = exit 2, NOTHING is written to --out):
 *   FORBIDDEN    no internal reference (people, rooms, trackers, hosts, paths) survives redaction
 *   SECRET       no credential-shaped string
 *   SYMLINK      no symlink at all in the public tree
 *   BINARY       every binary file is allowlisted by path + sha256
 *   LICENCE      every file maps to a decided licence; an undecided one is refused, never defaulted
 *   POLICY       every include/map source exists; a redaction rule that never fires is reported (dead rule)
 *   FORBIDDEN_FILE  no published file is byte-identical to a listed never-publish file (policy.forbiddenFiles, by hash)
 * Nothing here publishes, pushes or creates a repository.
 */

const { execFileSync } = require('node:child_process');
const crypto = require('node:crypto');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

const POLICY_PATH = 'release/public/policy.json';
const MANIFEST_NAME = 'PUBLIC-DERIVATION.json';

function fail(code, msg) { process.stderr.write(`derive-public: ${msg}\n`); process.exit(code); }
const sha256 = buf => crypto.createHash('sha256').update(buf).digest('hex');

function parseArgs(argv) {
  const a = { repo: process.cwd() };
  for (let i = 0; i < argv.length; i += 1) {
    if (argv[i] === '--rc') a.rc = argv[++i];
    else if (argv[i] === '--out') a.out = argv[++i];
    else if (argv[i] === '--repo') a.repo = argv[++i];
    else fail(64, `unknown argument: ${argv[i]}`);
  }
  if (!a.rc || !a.out) fail(64, 'usage: derive-public.cjs --rc <commit> --out <dir> [--repo <dir>]');
  return a;
}

/** Every file under root as a sorted list of POSIX relative paths (symlinks reported, not followed). */
function listTree(root) {
  const files = [];
  const links = [];
  const walk = rel => {
    for (const e of fs.readdirSync(path.join(root, rel), { withFileTypes: true })) {
      const r = rel ? `${rel}/${e.name}` : e.name;
      if (e.isSymbolicLink()) links.push(r);
      else if (e.isDirectory()) walk(r);
      else if (e.isFile()) files.push(r);
    }
  };
  walk('');
  return { files: files.sort(), links: links.sort() };
}

const isBinary = buf => buf.includes(0);
/** A pinned file is published byte-for-byte (its hash is evidence); only its listed tokens are tolerated by FORBIDDEN. */
const pinnedFor = (p, pol) => pol.pinnedVerbatim.find(x => x.paths.some(pre => underPrefix(p, pre)) && x.names.includes(p.split('/').pop())) || null;
const underPrefix = (p, prefix) => prefix.endsWith('/') ? p.startsWith(prefix) : p === prefix;

/** policy -> compiled regexes; a malformed policy refuses the derivation. */
function compilePolicy(policy) {
  const need = ['include', 'exclude', 'map', 'runtimeVerbatim', 'redactions', 'pinnedVerbatim', 'forbidden', 'secretShapes', 'binaryAllowlist', 'licences'];
  for (const k of need) if (!(k in policy)) throw new Error(`policy is missing "${k}"`);
  const rx = (s, f) => new RegExp(s, f);
  const seen = new Set();
  for (const l of policy.licences) { if (seen.has(l.path)) throw new Error(`two licence entries for ${l.path}`); seen.add(l.path); }
  return {
    include: policy.include, exclude: policy.exclude, map: policy.map,
    runtimeVerbatim: policy.runtimeVerbatim,
    redactions: policy.redactions.map(r => ({ ...r, re: rx(r.pattern, r.flags || 'g') })),
    forbidden: policy.forbidden.map(r => ({ ...r, re: rx(r.pattern, r.flags || 'gi') })),
    secretShapes: policy.secretShapes.map(r => ({ ...r, re: rx(r.pattern, 'g') })),
    binaryAllowlist: policy.binaryAllowlist,
    licences: policy.licences,
    pinnedVerbatim: policy.pinnedVerbatim,
    forbiddenFiles: policy.forbiddenFiles || [],
  };
}

/** The public path for a source path, or null when it is not published. */
function publicPathFor(p, pol) {
  for (const m of pol.map) {
    if (underPrefix(p, m.from)) return m.from.endsWith('/') ? m.to + p.slice(m.from.length) : m.to;
  }
  if (!pol.include.some(pre => underPrefix(p, pre))) return null;
  if (pol.exclude.some(x => underPrefix(p, x.path))) return null;
  return p;
}

function licenceFor(pub, pol) {
  let best = null;
  for (const l of pol.licences) if (underPrefix(pub, l.path) && (!best || l.path.length > best.path.length)) best = l;
  return best;
}

function derive(srcRoot, pol) {
  const errors = [];
  const { files, links } = listTree(srcRoot);
  for (const l of links) if (publicPathFor(l, pol) !== null) errors.push(`SYMLINK: ${l}`);
  for (const m of pol.map) if (!files.some(f => underPrefix(f, m.from))) errors.push(`POLICY: map source missing: ${m.from}`);
  for (const pre of pol.include) if (!files.some(f => underPrefix(f, pre))) errors.push(`POLICY: include matches nothing: ${pre}`);

  const ruleHits = Object.fromEntries(pol.redactions.map(r => [r.id, 0]));
  const out = new Map(); // public path -> { buf, mode, source }
  for (const f of files) {
    const pub = publicPathFor(f, pol);
    if (pub === null) continue;
    if (out.has(pub)) { errors.push(`POLICY: two sources map to ${pub}`); continue; }
    if (pub === MANIFEST_NAME) { errors.push(`POLICY: ${MANIFEST_NAME} is reserved for the derivation manifest`); continue; }
    let buf = fs.readFileSync(path.join(srcRoot, f));
    const mode = (fs.statSync(path.join(srcRoot, f)).mode & 0o111) ? 0o755 : 0o644;
    if (isBinary(buf)) {
      const allow = pol.binaryAllowlist.find(b => b.path === pub);
      if (!allow) errors.push(`BINARY: not allowlisted: ${pub}`);
      else if (allow.sha256 !== sha256(buf)) errors.push(`BINARY: sha256 mismatch: ${pub}`);
    } else if (!pinnedFor(pub, pol) && !pol.runtimeVerbatim.some(pre => underPrefix(pub, pre))) {
      let text = buf.toString('utf8');
      for (const r of pol.redactions) {
        if (r.files && !r.files.some(pre => underPrefix(pub, pre))) continue;
        const n = (text.match(r.re) || []).length;
        if (n) { ruleHits[r.id] += n; text = text.replace(r.re, r.replace); }
      }
      buf = Buffer.from(text, 'utf8');
    }
    out.set(pub, { buf, mode, source: f });
  }

  const forbiddenHits = [];
  const secretHits = [];
  for (const [pub, { buf }] of out) {
    if (isBinary(buf)) continue;
    let text = buf.toString('utf8');
    const pin = pinnedFor(pub, pol);
    if (pin) for (const tok of pin.tolerated) text = text.split(tok).join('');
    for (const r of pol.forbidden) {
      const n = (text.match(r.re) || []).length;
      if (n) forbiddenHits.push({ path: pub, rule: r.id, count: n, sample: text.match(new RegExp(`.{0,40}${r.re.source}.{0,40}`, r.re.flags.replace('g', '')))?.[0] });
    }
    for (const r of pol.secretShapes) if (r.re.test(text)) { secretHits.push({ path: pub, rule: r.id }); r.re.lastIndex = 0; }
  }
  for (const [pub, { buf }] of out) {
    const hit = pol.forbiddenFiles.find(f => f.sha256 === sha256(buf));
    if (hit) errors.push(`FORBIDDEN_FILE: ${pub} is byte-identical to a never-publish file (${hit.reason})`);
  }
  for (const h of forbiddenHits) errors.push(`FORBIDDEN: ${h.path} [${h.rule}] x${h.count}: ${JSON.stringify(h.sample)}`);
  for (const h of secretHits) errors.push(`SECRET: ${h.path} [${h.rule}]`);

  const licenceOf = {};
  for (const pub of out.keys()) {
    const l = licenceFor(pub, pol);
    if (!l) errors.push(`LICENCE: no licence mapping: ${pub}`);
    else if (!/^[A-Za-z0-9.+-]+$/.test(l.spdx)) errors.push(`LICENCE: undecided licence "${l.spdx}": ${pub}`);
    else licenceOf[pub] = l.spdx;
  }
  const deadRules = Object.entries(ruleHits).filter(([, n]) => n === 0).map(([id]) => id);
  return { errors, out, ruleHits, deadRules, licenceOf, excluded: files.filter(f => publicPathFor(f, pol) === null).length };
}

function main() {
  const a = parseArgs(process.argv.slice(2));
  const repo = path.resolve(a.repo);
  const outDir = path.resolve(a.out);
  if (fs.existsSync(outDir) && fs.readdirSync(outDir).length) fail(2, `OUT_NOT_EMPTY: ${outDir}`);
  const git = args => execFileSync('git', ['-C', repo, ...args], { encoding: null, maxBuffer: 1024 * 1024 * 1024 });
  let sha;
  try { sha = git(['rev-parse', '--verify', `${a.rc}^{commit}`]).toString().trim(); } catch { fail(2, `RC_NOT_FOUND: ${a.rc}`); }

  const work = fs.mkdtempSync(path.join(os.tmpdir(), 'eo-derive-'));
  const failClean = (code, msg) => { fs.rmSync(work, { recursive: true, force: true }); fail(code, msg); };
  try {
    const src = path.join(work, 'src');
    fs.mkdirSync(src);
    fs.writeFileSync(path.join(work, 'source.tar'), git(['archive', '--format=tar', sha]));
    execFileSync('tar', ['-xf', path.join(work, 'source.tar'), '-C', src]);
    const policyFile = path.join(src, POLICY_PATH);
    if (!fs.existsSync(policyFile)) failClean(2, `NO_POLICY: ${POLICY_PATH} is not in ${sha}`);
    const policyBuf = fs.readFileSync(policyFile);
    let pol;
    try { pol = compilePolicy(JSON.parse(policyBuf.toString('utf8'))); } catch (e) { failClean(2, `POLICY_INVALID: ${e.message}`); }

    const r = derive(src, pol);
    if (r.errors.length) {
      failClean(2, `GATE_FAIL (${r.errors.length}):\n${r.errors.slice(0, 2000).join('\n')}${r.errors.length > 2000 ? `\n... ${r.errors.length - 2000} more` : ''}`);
    }

    const entries = [...r.out.keys()].sort().map(p => ({ path: p, mode: r.out.get(p).mode.toString(8), sha256: sha256(r.out.get(p).buf), source: r.out.get(p).source,
      transformed: sha256(r.out.get(p).buf) !== sha256(fs.readFileSync(path.join(src, r.out.get(p).source))), licence: r.licenceOf[p] }));
    const runtimeChanged = entries.filter(e => e.transformed && pol.runtimeVerbatim.some(pre => underPrefix(e.path, pre)));
    if (runtimeChanged.length) failClean(2, `RUNTIME_NOT_VERBATIM: ${runtimeChanged.map(e => e.path).join(', ')}`);
    const treeDigest = sha256(entries.map(e => `${e.path}\0${e.mode}\0${e.sha256}\n`).join(''));
    const manifest = {
      PRODUCT: 'Euro-Office_MCP', SOURCE_COMMIT: sha, POLICY: POLICY_PATH, POLICY_SHA256: sha256(policyBuf),
      PUBLIC_TREE_SHA256: treeDigest, FILES: entries.length, EXCLUDED_SOURCE_FILES: r.excluded,
      GATES: { FORBIDDEN: 'PASS', SECRET: 'PASS', SYMLINK: 'PASS', BINARY: 'PASS', LICENCE: 'PASS', POLICY: 'PASS', RUNTIME_VERBATIM: 'PASS', FORBIDDEN_FILE: 'PASS' },
      TRANSFORMED_FILES: entries.filter(e => e.transformed).length,
      RULE_HITS: r.ruleHits, DEAD_RULES: r.deadRules, ENTRIES: entries,
    };

    const stage = path.join(work, 'out');
    for (const e of entries) {
      const dst = path.join(stage, e.path);
      fs.mkdirSync(path.dirname(dst), { recursive: true });
      fs.writeFileSync(dst, r.out.get(e.path).buf);
      fs.chmodSync(dst, r.out.get(e.path).mode);
    }
    // the manifest travels WITH the public tree (provenance) but is outside PUBLIC_TREE_SHA256 (it carries that digest)
    fs.writeFileSync(path.join(stage, MANIFEST_NAME), `${JSON.stringify(manifest, null, 2)}\n`);
    fs.mkdirSync(outDir, { recursive: true });
    for (const name of fs.readdirSync(stage)) fs.cpSync(path.join(stage, name), path.join(outDir, name), { recursive: true });
    const manifestPath = path.join(outDir, MANIFEST_NAME);
    process.stdout.write([`SOURCE_SHA=${sha}`, `PUBLIC_TREE_SHA256=${treeDigest}`, `FILES=${entries.length}`,
      `DEAD_RULES=${r.deadRules.join(',') || 'none'}`, `MANIFEST=${manifestPath}`].join('\n') + '\n');
  } finally {
    fs.rmSync(work, { recursive: true, force: true });
  }
}

if (require.main === module) main();
module.exports = { compilePolicy, derive, publicPathFor };
