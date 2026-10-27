'use strict';

/**
 * L-05 -- THE WARRANT FOR THE B03-WRITE RUNTIME LIST. The B03 Bridge WRITE capability (worksheet.copy/v1: copy_sheet) is written only on the
 * runtimes in XLSX_B03_BRIDGE_WRITE_RUNTIME_ENTRIES. The list was /8 only and then went STALE: /9 .. /12 are accepted for every other path, the
 * copy handler on them is the same bytes, and nobody added them -- so the prod DS (/12) answered FAIL_CLOSED 'xlsx-copy-bridge-write-runtime-not-proven'
 * for EVERY copy_sheet (L-05 LIVE measurement) while nothing was actually missing. A list entry is justified by ONE thing: the handler, and the
 * helpers it calls, are byte-identical to the first listed revision's. This test extracts them from EVERY listed revision's bridge-block.js
 * (a JS-aware brace match, strings skipped) and requires the sha256 to be the same -- and equal to a pinned value, so that a change to ALL revisions
 * at once is a deliberate, visible edit too. If a future revision differs, this goes red: the revision then does NOT join the list (or the handler
 * change is reviewed first). The same test pins the capability's MANIFEST entry (the id / kind / contract the client checks).
 * The B05 sibling: test/static/b05-write-handler-identity.test.cjs (L-03/2).
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');

const DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.4.60', 'euro-office-bridge-v1');
const CAP = 'worksheet.copy/v1';
const sha = text => crypto.createHash('sha256').update(text).digest('hex');
/** the pinned sha256 of (handler + the helpers it calls) -- the bytes every listed revision carries (computed from rev8..rev12); a deliberate edit to change it */
const PINNED = '648dc936299326f5feff1b210c6a735da577d5957e3a36fd8a77a8acbec34237';

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
/** every bridge-local function the handler calls, to a fixed point (here: resolveCell) */
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

// DS9346 (owner D5=B, VALTAS, docs/DS9346-RUNTIME-RETIREMENT.md): core.XLSX_B03_BRIDGE_WRITE_RUNTIME_ENTRIES now
// carries ONE entry (9.3.5.0/16) -- the HISTORICAL 9.3.4.60 rev8..rev16 chain this file audits is a frozen,
// on-disk fact (MEASURED, grep '"worksheet.copy/v1": function' across rev8..rev16-candidate) independent of
// what is currently accepted.
const listed = [8, 9, 10, 11, 12, 13, 14, 15, 16];

test('the listed revisions are exactly the ones with a bridge-block.js on disk, and each one carries the handler', () => {
  assert.ok(listed.length >= 5, `listed: ${listed}`);
  for (const rev of listed) {
    assert.ok(fs.existsSync(path.join(DIR, `rev${rev}-candidate`, 'bridge-block.js')), `rev${rev}-candidate/bridge-block.js`);
    const h = handlerOf(blockOf(rev));
    assert.ok(h.length > 1000 && h.includes('"bad-args"') && h.includes('"invalid-name"') && h.includes('asc_copyWorksheet'), `rev${rev} handler shape`);
  }
});

test('WARRANT: the handler AND the helpers it calls are byte-identical on EVERY listed revision (sha256 equal)', () => {
  const digests = Object.fromEntries(listed.map(rev => [rev, digestOf(blockOf(rev))]));
  const distinct = new Set(Object.values(digests));
  assert.equal(distinct.size, 1, `the handler differs between listed revisions: ${JSON.stringify(digests)}`);
  // the helper closure is whatever the handler really calls -- a handler that grows a dependency is a visible change here (pinned below with the digest)
  assert.ok(Array.isArray(warrantOf(blockOf(listed[0])).deps));
  assert.match([...distinct][0], /^[0-9a-f]{64}$/);
  assert.equal([...distinct][0], PINNED);
});

test('WARRANT: the capability MANIFEST entry (what the client checks) is byte-identical on every listed revision', () => {
  const entries = Object.fromEntries(listed.map(rev => [rev, sha(manifestEntryOf(blockOf(rev)))]));
  assert.equal(new Set(Object.values(entries)).size, 1, JSON.stringify(entries));
  const entry = manifestEntryOf(blockOf(listed[0]));
  assert.ok(entry.includes('kind: "WRITE"') && entry.includes('TARGETED'), entry.slice(0, 200));
});

// DS9346: the "every ACCEPTED revision from /8 up carries the handler" staleness guard no longer has subject
// matter -- ACCEPTED_RUNTIMES.xlsx is now the single 9.3.5.0/16 entry; the WARRANT tests above and the 9.3.5.0
// check below already cover the live list directly.

test('NEGATIVE: /5, /6 and /7 (the revisions BEFORE the copy capability) have no handler, so they are NOT listed', () => {
  const v5 = fs.readFileSync(path.join(DIR, 'bridge-block.js'), 'utf8');
  assert.equal(v5.includes(CAP), false);
  assert.throws(() => handlerOf(v5), /not found/);
  for (const rev of [6, 7]) assert.equal(blockOf(rev).includes(`"${CAP}": function (args)`), false, `rev${rev} carries no copy handler`);
  for (const rev of [5, 6, 7]) assert.equal(listed.includes(rev), false, `rev${rev}`);
});

test('RED CONTROL: ONE changed byte in a handler copy changes the digest -- the warrant can go red', () => {
  const src = blockOf(listed[0]);
  const real = digestOf(src);
  const mutated = src.replace('nm.length > 31', 'nm.length > 32');
  assert.notEqual(mutated, src, 'the mutation applied');
  assert.notEqual(digestOf(mutated), real);
  // and a changed byte anywhere in the handler's own text (the API list the handler gates on) is caught the same way
  assert.notEqual(digestOf(src.replace('"asc_registerCallback"', '"asc_registerCallbackX"')), real);
  // an unbalanced / missing handler is an error, never a silent "equal"
  assert.throws(() => digestOf('Api.X = function () {}'), /not found/);
});

test('RED CONTROL: a listed revision whose handler differs makes the cross-revision comparison fail', () => {
  const a = digestOf(blockOf(listed[0]));
  const b = digestOf(blockOf(listed[1]).replace('"workbook-protected"', '"workbook-protected-x"'));
  assert.notEqual(a, b);
});

test('WARRANT (DS v9.3.5 porting, #15): the 9.3.5.0 bridge-block.js carries the SAME handler (sha256 equal to PINNED)', () => {
  const src935 = fs.readFileSync(path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.5.0', 'euro-office-bridge-v1', 'bridge-block.js'), 'utf8');
  assert.equal(digestOf(src935), PINNED);
});

test('RED CONTROL (DS v9.3.5 porting, #15): a changed byte in the 9.3.5.0 handler copy no longer matches PINNED', () => {
  const src935 = fs.readFileSync(path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.5.0', 'euro-office-bridge-v1', 'bridge-block.js'), 'utf8');
  const mutated = src935.replace('nm.length > 31', 'nm.length > 32');
  assert.notEqual(mutated, src935, 'the mutation applied');
  assert.notEqual(digestOf(mutated), PINNED);
});
