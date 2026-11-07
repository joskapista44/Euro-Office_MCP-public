'use strict';

/**
 * GH #365 option C, U1 (card an internal branch, kioszto zero): the /29-veryhidden CANDIDATE
 * (documentserver-patches/9.3.5.0/euro-office-bridge-v1/rev29-veryhidden-candidate/bridge-block.js,
 * NOT installed anywhere) CARRIES EVERY /28 HANDLER BYTE-IDENTICAL, with EXACTLY ONE NEW
 * capability added: "sheet.veryHidden.read/v1". No existing capability is removed or modified.
 * Reuses the SAME capsOf()/spanFrom()/digestOf() extraction the sibling /28 warrant test
 * (ds935-bridge28-table-autofilter-criteria-write-handler-identity.test.cjs) already established.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');
const { execFileSync } = require('node:child_process');

const DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.5.0', 'euro-office-bridge-v1');
const REV28_DIR = path.join(DIR, 'rev28-table-autofilter-criteria-candidate');
const REV28 = path.join(REV28_DIR, 'bridge-block.js');
const REV29_DIR = path.join(DIR, 'rev29-veryhidden-candidate');
const REV29 = path.join(REV29_DIR, 'bridge-block.js');
const BUILD_SCRIPT = path.join(REV29_DIR, 'build-rev29-veryhidden.py');

const NEW_CAPABILITY = 'sheet.veryHidden.read/v1';

const sha = text => crypto.createHash('sha256').update(text).digest('hex');

/** the index of the `}` closing the `{` at `open`; strings ('..', "..", `..`) and escapes are skipped. -1 if unbalanced. */
function closingBrace(src, open) {
  let depth = 0; let quote = null;
  for (let k = open; k < src.length; k += 1) {
    const c = src[k];
    if (quote) { if (c === '\\') { k += 1; continue; } if (c === quote) quote = null; continue; }
    if (c === '"' || c === '\'' || c === '`') quote = c;
    else if (c === '{') depth += 1;
    else if (c === '}') { depth -= 1; if (depth === 0) return k; }
  }
  return -1;
}
function spanFrom(src, head) {
  const i = src.indexOf(head);
  if (i < 0) throw new Error(`not found: ${head}`);
  const open = src.indexOf('{', i + head.length - 1);
  const close = closingBrace(src, open);
  if (open < 0 || close < 0) throw new Error(`unbalanced: ${head}`);
  return src.slice(i, close + 1);
}
function dependenciesOf(src, text, seen = new Map()) {
  for (const m of text.matchAll(/\b([A-Za-z_]\w*)\s*\(/g)) {
    const name = m[1];
    if (seen.has(name) || ['function', 'if', 'for', 'while', 'switch', 'catch', 'return', 'typeof'].includes(name) || !src.includes(`function ${name}(`)) continue;
    const body = spanFrom(src, `function ${name}(`); seen.set(name, body); dependenciesOf(src, body, seen);
  }
  return seen;
}
const capsOf = src => [...new Set([...src.matchAll(/"([A-Za-z.]+\/v\d+)": function \(args\) \{/g)].map(m => m[1]))].sort();
const digestOf = (src, cap) => {
  const h = spanFrom(src, `"${cap}": function (args) {`);
  const deps = [...dependenciesOf(src, h).entries()].sort(([a], [b]) => (a < b ? -1 : 1)).map(([n, b]) => `${n}:${b}`);
  return sha([h, ...deps].join('\n'));
};

let s28; let s29;
test.before(() => {
  s28 = fs.readFileSync(REV28, 'utf8');
  s29 = fs.readFileSync(REV29, 'utf8');
});

test('/29 carries EVERY /28 capability PLUS exactly one new one -- nothing removed', () => {
  const c28 = capsOf(s28); const c29 = capsOf(s29);
  assert.ok(c28.length >= 40, `rev28 handlers: ${c28.length}`);
  assert.deepEqual(c29.filter(c => !c28.includes(c)), [NEW_CAPABILITY], 'exactly one new capability on /29');
  assert.deepEqual(c28.filter(c => !c29.includes(c)), [], 'nothing dropped on /29');
});

test('WARRANT: every /28 handler AND the helpers it calls are byte-identical on /29', () => {
  const differing = capsOf(s28).filter(cap => digestOf(s28, cap) !== digestOf(s29, cap));
  assert.deepEqual(differing, [], 'no existing /28 handler differs on /29');
  const inst = s28.slice(0, s28.indexOf('Api.GetEuroOfficeBridge = function () {'));
  assert.ok(inst.length > 100, 'the /28 block starts with the layout installer');
  assert.ok(s29.startsWith(inst), 'the /29 block starts with the SAME installer, byte for byte');
});

test('RED CONTROL: one changed byte in a /29 copy of an UNRELATED /28 handler breaks the warrant', () => {
  const cap = 'sheet.protection.write/v1';
  const h = spanFrom(s29, `"${cap}": function (args) {`);
  const mutated = s29.replace(h, h.replace('return', 'return '));
  assert.notEqual(mutated, s29, 'the mutation applied');
  assert.notEqual(digestOf(mutated, cap), digestOf(s28, cap), cap);
});

test('the new veryHidden handler is brace-balanced and extractable, reads both model fields defensively (never a bare property read), and never guesses a third state', () => {
  const h = spanFrom(s29, `"${NEW_CAPABILITY}": function (args) {`);
  assert.match(h, /reason: "bad-args"/);
  assert.match(h, /reason: "sheet-not-found"/);
  assert.match(h, /reason: "worksheet-missing"/);
  assert.match(h, /reason: "threw"/);
  for (const prop of ['m.bHidden', 'm.bEoVeryHidden']) {
    assert.match(h, new RegExp(`try \\{[^;]*${prop.replace('.', '\\.')}`), `${prop} must be read inside a try block`);
  }
  // the ONLY happy-path shape: a single boolean field, derived from BOTH model fields AND-ed
  // together -- never a bare pass-through of bHidden alone (that would not distinguish veryHidden
  // from plain hidden, the exact defect this card exists to fix) and never a third state/string.
  assert.match(h, /value: \{ veryHidden: h === true && vh === true \} /);
  assert.doesNotMatch(h, /value: \{ veryHidden: h \} /, 'must never report bHidden alone as veryHidden');
});

test('the new manifest entry exists exactly once, names the companion engine REPLACE it depends on, and the embedded patchVersion marker advanced to /29', () => {
  const entryCount = (s29.match(new RegExp(`\\{ id: "${NEW_CAPABILITY.replace(/[/.]/g, '\\$&')}"`, 'g')) || []).length;
  assert.equal(entryCount, 1);
  const entry = spanFrom(s29, `{ id: "${NEW_CAPABILITY}"`);
  assert.match(entry, /bEoVeryHidden/);
  assert.match(entry, /never throws/);
  assert.equal((s29.match(/patchVersion: "euro-office-bridge-v1\/9\.3\.5\.0\/29"/g) || []).length, 1);
  assert.equal((s29.match(/patchVersion: "euro-office-bridge-v1\/9\.3\.5\.0\/28"/g) || []).length, 0, 'the old /28 marker must not remain');
});

test('the build script is deterministic: running it again (via --check) reproduces the SAME output_sha256 already on disk', () => {
  const out = execFileSync('python3', [BUILD_SCRIPT, '--check'], { encoding: 'utf8' });
  const { output_sha256: fresh } = JSON.parse(out);
  assert.equal(fresh, sha(s29));
});

test('manifest-mechanism-proof.json names the real base_sha256 (the actual rev28-table-autofilter-criteria-candidate/bridge-block.js on disk) and the real output_sha256', () => {
  const proof = JSON.parse(fs.readFileSync(path.join(REV29_DIR, 'manifest-mechanism-proof.json'), 'utf8'));
  assert.equal(proof.base_sha256, sha(s28));
  assert.equal(proof.output_sha256, sha(s29));
  assert.deepEqual(proof.new_capabilities, [NEW_CAPABILITY]);
});
