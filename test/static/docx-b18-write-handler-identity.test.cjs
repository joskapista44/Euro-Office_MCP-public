/**
 * M11.5d.4 sibling (#15, DS v9.3.5 porting) -- THE WARRANT FOR THE DOCX B18-WRITE RUNTIME LIST, previously untested.
 * The B18-WRITE Bridge capability (docx.paragraph.numbering.remove/v1: remove_list, outdent at level 0) is written
 * only on the runtimes in DOCX_B18_BRIDGE_WRITE_RUNTIME_ENTRIES (word /2, /3). A list entry is justified by ONE
 * thing: the handler, and the helpers it calls, are byte-identical to the first listed revision's. This test
 * extracts them from EVERY listed revision's bridge-block.js (a JS-aware brace match, strings skipped) and
 * requires the sha256 to be the same -- and equal to a pinned value, so that a change to ALL revisions at once is
 * a deliberate, visible edit too. It also extends the check to the DS v9.3.5 bridge-block.js added by the #15
 * porting unit (word /3's own line comment -- "9.3.5.0/3 (#15, DS v9.3.5 porting): byte-identical to /3").
 * The XLSX sibling (single-capability WRITE list, same shape): test/static/b03-write-handler-identity.test.cjs.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');

const DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.4.60', 'euro-office-bridge-word-v1');
const CAP = 'docx.paragraph.numbering.remove/v1';
const sha = text => crypto.createHash('sha256').update(text).digest('hex');
/** the pinned sha256 of (handler + the helpers it calls) -- the bytes every listed revision carries (computed from rev2, rev3); a deliberate edit to change it */
const PINNED = '5c5b18805fc579ec1db7e16c94edef0100d0e13c1e3357357cc295d5f7c3adb5';

/** the index of the `}` closing the `{` at `open`; strings ('..', "..", `..`) and escapes are skipped. -1 if unbalanced. */
function closingBrace(src, open) {
  let depth = 0; let quote = null;
  for (let k = open; k < src.length; k += 1) {
    const c = src[k];
    if (quote) { if (c === '\\') k += 1; else if (c === quote) quote = null; continue; }
    if (c === '"' || c === '\'' || c === '`') quote = c;
    else if (c === '{') depth += 1;
    else if (c === '}') { depth -= 1; if (depth === 0) return k; }
  }
  return -1;
}
/** the text from `head` (inclusive) to the brace that closes the first `{` at/after head's last character; throws if absent */
function spanFrom(src, head) {
  const i = src.indexOf(head);
  if (i < 0) throw new Error(`not found: ${head}`);
  const open = src.indexOf('{', i + head.length - 1);
  const close = closingBrace(src, open);
  if (open < 0 || close < 0) throw new Error(`unbalanced: ${head}`);
  return src.slice(i, close + 1);
}
const handlerOf = src => spanFrom(src, `"${CAP}": function (args) {`);
const helperOf = (src, name) => spanFrom(src, `function ${name}(`);
/** every bridge-local function the handler calls, to a fixed point */
function dependenciesOf(src, text, seen = new Map()) {
  for (const m of text.matchAll(/\b([A-Za-z_]\w*)\s*\(/g)) {
    const name = m[1];
    if (seen.has(name) || ['function', 'if', 'for', 'while', 'switch', 'catch', 'return', 'typeof'].includes(name) || !src.includes(`function ${name}(`)) continue;
    const body = helperOf(src, name); seen.set(name, body); dependenciesOf(src, body, seen);
  }
  return seen;
}
const warrantOf = src => { const h = handlerOf(src); const deps = dependenciesOf(src, h); return { handler: h, deps: [...deps.entries()].sort(([a], [b]) => (a < b ? -1 : 1)).map(([n, b]) => `${n}:${b}`) }; };
const digestOf = src => { const w = warrantOf(src); return sha([w.handler, ...w.deps].join('\n')); };
const manifestEntryOf = src => {
  const i = src.indexOf(`{ id: "${CAP}"`);
  if (i < 0) throw new Error('manifest entry not found');
  return src.slice(i, closingBrace(src, i) + 1);
};
const blockOf = rev => fs.readFileSync(path.join(DIR, `rev${rev}-candidate`, 'bridge-block.js'), 'utf8');

// DS9346 (owner D5=B, VALTAS, docs/DS9346-RUNTIME-RETIREMENT.md): core.DOCX_B18_BRIDGE_WRITE_RUNTIME_ENTRIES
// now carries ONE entry (9.3.5.0/3) -- the HISTORICAL 9.3.4.60 rev2/rev3 chain this file audits is a frozen,
// on-disk fact (MEASURED, grep '"docx.paragraph.numbering.remove/v1": function' across rev2..3-candidate)
// independent of what is currently accepted. The 9.3.5.0 copy is checked separately below, against the
// pinned digest.
const listed = [2, 3];

test('the listed revisions are exactly /2, /3, and each carries the handler', () => {
  assert.deepEqual(listed, [2, 3]);
  for (const rev of listed) {
    assert.ok(fs.existsSync(path.join(DIR, `rev${rev}-candidate`, 'bridge-block.js')), `rev${rev}-candidate/bridge-block.js`);
    const h = handlerOf(blockOf(rev));
    assert.ok(h.length > 500 && h.includes('"level-mismatch"') && h.includes('RemoveNumPr'), `rev${rev} handler shape`);
  }
});

test('WARRANT: the handler AND the helpers it calls are byte-identical on EVERY listed revision (sha256 equal)', () => {
  const digests = Object.fromEntries(listed.map(rev => [rev, digestOf(blockOf(rev))]));
  const distinct = new Set(Object.values(digests));
  assert.equal(distinct.size, 1, `the handler differs between listed revisions: ${JSON.stringify(digests)}`);
  assert.match([...distinct][0], /^[0-9a-f]{64}$/);
  assert.equal([...distinct][0], PINNED);
});

test('WARRANT: the capability MANIFEST entry (what the client checks) is byte-identical on every listed revision', () => {
  const entries = Object.fromEntries(listed.map(rev => [rev, sha(manifestEntryOf(blockOf(rev)))]));
  assert.equal(new Set(Object.values(entries)).size, 1, JSON.stringify(entries));
  const entry = manifestEntryOf(blockOf(listed[0]));
  assert.ok(entry.includes('kind: "WRITE"') && entry.includes('TARGETED'), entry.slice(0, 200));
});

// DS9346: the "every ACCEPTED revision from /2 up carries the handler" staleness guard no longer has subject
// matter -- ACCEPTED_RUNTIMES.docx is now the single 9.3.5.0/5 entry (GH #292 S4 D-state); the WARRANT tests above and the 9.3.5.0
// check below already cover the live list directly.

test('NEGATIVE: /1 (the revision BEFORE the numbering-removal capability) has no handler, so it is NOT listed', () => {
  const v1 = fs.readFileSync(path.join(DIR, 'bridge-block.js'), 'utf8');
  assert.equal(v1.includes(CAP), false);
  assert.throws(() => handlerOf(v1), /not found/);
  assert.equal(listed.includes(1), false);
});

test('RED CONTROL: ONE changed byte in a handler copy changes the digest -- the warrant can go red', () => {
  const src = blockOf(listed[0]);
  const real = digestOf(src);
  const mutated = src.replace('"level-mismatch"', '"level-mismatchX"');
  assert.notEqual(mutated, src, 'the mutation applied');
  assert.notEqual(digestOf(mutated), real);
  assert.throws(() => digestOf('Api.X = function () {}'), /not found/);
});

test('RED CONTROL: a listed revision whose handler differs makes the cross-revision comparison fail', () => {
  const a = digestOf(blockOf(listed[0]));
  const b = digestOf(blockOf(listed[1]).replace('"level-mismatch"', '"level-mismatchX"'));
  assert.notEqual(a, b);
});

test('WARRANT (DS v9.3.5 porting, #15): the 9.3.5.0 bridge-block.js carries the SAME handler (sha256 equal to PINNED)', () => {
  const src935 = fs.readFileSync(path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.5.0', 'euro-office-bridge-word-v1', 'bridge-block.js'), 'utf8');
  assert.equal(digestOf(src935), PINNED);
});

test('RED CONTROL (DS v9.3.5 porting, #15): a changed byte in the 9.3.5.0 handler copy no longer matches PINNED', () => {
  const src935 = fs.readFileSync(path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.5.0', 'euro-office-bridge-word-v1', 'bridge-block.js'), 'utf8');
  const mutated = src935.replace('"level-mismatch"', '"level-mismatchX"');
  assert.notEqual(mutated, src935, 'the mutation applied');
  assert.notEqual(digestOf(mutated), PINNED);
});

test('WARRANT (GH #292 S4 COMPAT): the word /4 candidate (9.3.5.0/rev4-candidate) carries the SAME handler (sha256 equal to PINNED) -- why /4 is on this list', () => {
  const src4 = fs.readFileSync(path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.5.0', 'euro-office-bridge-word-v1', 'rev4-candidate', 'bridge-block.js'), 'utf8');
  assert.equal(digestOf(src4), PINNED);
});
