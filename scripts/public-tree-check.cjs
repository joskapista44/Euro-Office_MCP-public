#!/usr/bin/env node
'use strict';

/**
 * Public-repository CI gate (M12, owner decision D6).
 *
 *   node scripts/public-tree-check.cjs [--root <dir>]
 *
 * The public tree is never edited by hand: it is the output of the private derivation, whose gates (forbidden
 * references, secrets, symlinks, binaries, licences) ran on exactly these bytes and are recorded in
 * PUBLIC-DERIVATION.json. This check proves the checkout IS that output -- every file present, nothing extra, same
 * bytes, same mode -- and independently scans it for generic credential shapes and symlinks. Exit 0 PASS | 2 FAIL.
 */

const crypto = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');

const MANIFEST = 'PUBLIC-DERIVATION.json';
const SKIP_DIRS = new Set(['.git', 'node_modules']);
const SECRET_SHAPES = [
  ['github-token', /gh[pousr]_[A-Za-z0-9]{30,}|github_pat_[A-Za-z0-9_]{30,}/],
  ['private-key', /-----BEGIN [A-Z ]*PRIVATE KEY-----/],
  ['aws-key', /AKIA[0-9A-Z]{16}/],
  ['slack-token', /xox[bpas]-[0-9A-Za-z-]{10,}/],
  ['jwt', /eyJ[A-Za-z0-9_-]{15,}\.eyJ[A-Za-z0-9_-]{15,}\.[A-Za-z0-9_-]{10,}/],
];

function walk(root, rel = '', out = { files: [], links: [] }) {
  for (const d of fs.readdirSync(path.join(root, rel), { withFileTypes: true })) {
    const p = rel ? `${rel}/${d.name}` : d.name;
    if (d.isSymbolicLink()) out.links.push(p);
    else if (d.isDirectory()) { if (!SKIP_DIRS.has(d.name) || rel) walk(root, p, out); }
    else if (d.isFile()) out.files.push(p);
  }
  return out;
}

function check(root) {
  const errors = [];
  let manifest;
  try { manifest = JSON.parse(fs.readFileSync(path.join(root, MANIFEST), 'utf8')); } catch (e) { return [`MANIFEST: ${MANIFEST} unreadable (${e.message})`]; }
  const expected = new Map((manifest.ENTRIES || []).map(e => [e.path, e]));
  if (expected.size === 0) errors.push('MANIFEST: no ENTRIES');
  const { files, links } = walk(root);
  for (const l of links) errors.push(`SYMLINK: ${l}`);
  for (const f of files) {
    if (f === MANIFEST) continue;
    const e = expected.get(f);
    const buf = fs.readFileSync(path.join(root, f));
    if (!e) errors.push(`EXTRA: ${f} is not in ${MANIFEST}`);
    else {
      if (crypto.createHash('sha256').update(buf).digest('hex') !== e.sha256) errors.push(`CHANGED: ${f}`);
      const mode = (fs.statSync(path.join(root, f)).mode & 0o111) ? '755' : '644';
      if (mode !== e.mode) errors.push(`MODE: ${f} is ${mode}, derived ${e.mode}`);
      expected.delete(f);
    }
    const text = buf.toString('latin1');
    for (const [id, re] of SECRET_SHAPES) if (re.test(text)) errors.push(`SECRET: ${id} shape in ${f}`);
  }
  for (const f of expected.keys()) errors.push(`MISSING: ${f}`);
  return errors;
}

if (require.main === module) {
  const args = process.argv.slice(2);
  const i = args.indexOf('--root');
  const root = path.resolve(i === -1 ? '.' : args[i + 1] || '.');
  const errors = check(root);
  if (errors.length) { process.stderr.write(`PUBLIC_TREE_CHECK FAIL (${errors.length})\n${errors.join('\n')}\n`); process.exitCode = 2; }
  else process.stdout.write('PUBLIC_TREE_CHECK PASS\n');
}

module.exports = { check };
