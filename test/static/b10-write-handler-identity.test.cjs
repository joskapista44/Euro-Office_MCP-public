'use strict';

/**
 * GH #138 U4 (a maintainer's review, #ref) -- THE WARRANT FOR THE B10-WRITE RUNTIME LIST. The B10 Bridge WRITE
 * capability (sheet.printSettings.write/v1: set_print_fit/set_header_footer/set_page_break/
 * set_first_page_number/set_print_centering/set_print_scale/set_paper_size/set_header_footer_flag) was written
 * only on XLSX_B10_BRIDGE_WRITE_RUNTIME_ENTRIES = [/7, /8] -- STALE in the SAME way L-05 found for B03, X-O05/3
 * for B22 and X-O07/3 for B23: /9..15 are accepted for every other path, the printSettings.write handler on
 * them is the same bytes, and nobody widened the list -- so the prod DS (/15) would have answered FAIL_CLOSED
 * 'xlsx-print-bridge-write-runtime-not-proven' for EVERY one of the 8 print-family writes this release adds.
 * A list entry is justified by ONE thing: the handler, and the helpers it calls, are byte-identical to the
 * first listed revision's. This test extracts them from EVERY listed revision's bridge-block.js (a JS-aware
 * brace match, strings skipped) and requires the sha256 to be the same -- and equal to a pinned value, so that
 * a change to ALL revisions at once is a deliberate, visible edit too. If a future revision differs, this goes
 * red: the revision then does NOT join the list (or the handler change is reviewed first).
 * The B03 sibling: test/static/b03-write-handler-identity.test.cjs (L-05). The B22/B23 siblings (the SAME
 * staleness class, found 2026-10-15): test/static/b22-write-handler-identity.test.cjs,
 * test/static/b23-write-handler-identity.test.cjs.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');

const DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.4.60', 'euro-office-bridge-v1');
const CAP = 'sheet.printSettings.write/v1';
const sha = text => crypto.createHash('sha256').update(text).digest('hex');
/** the pinned sha256 of (handler + the helpers it calls) -- the bytes every listed revision carries (computed from rev7..rev15); a deliberate edit to change it */
const PINNED = '4a0723ee645ba4662801218c685e47664e081835982ad580c4f33ab42e0792ea';

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

// DS9346 (owner D5=B, VALTAS, docs/DS9346-RUNTIME-RETIREMENT.md): core.XLSX_B10_BRIDGE_WRITE_RUNTIME_ENTRIES now
// carries ONE entry (9.3.5.0/16) -- the HISTORICAL 9.3.4.60 rev7..rev16 chain this file audits is a frozen,
// on-disk fact (MEASURED, grep '"sheet.printSettings.write/v1": function' across rev7..rev16-candidate)
// independent of what is currently accepted. The 9.3.5.0 copy is checked separately below, against the
// pinned digest.
const listed = [7, 8, 9, 10, 11, 12, 13, 14, 15, 16];

test('the listed revisions are exactly the ones with a bridge-block.js on disk, and each one carries the handler', () => {
  assert.deepEqual(listed, [7, 8, 9, 10, 11, 12, 13, 14, 15, 16]);
  for (const rev of listed) {
    assert.ok(fs.existsSync(path.join(DIR, `rev${rev}-candidate`, 'bridge-block.js')), `rev${rev}-candidate/bridge-block.js`);
    const h = handlerOf(blockOf(rev));
    assert.ok(h.length > 200 && h.includes('"bad-args"') && h.includes('"unknown-setting"') && h.includes('headerFooterFlag'), `rev${rev} handler shape`);
  }
});

test('WARRANT: the handler AND the helpers it calls are byte-identical on EVERY listed revision (sha256 equal)', () => {
  const digests = Object.fromEntries(listed.map(rev => [rev, digestOf(blockOf(rev))]));
  const distinct = new Set(Object.values(digests));
  assert.equal(distinct.size, 1, `the handler differs between listed revisions: ${JSON.stringify(digests)}`);
  assert.ok(Array.isArray(warrantOf(blockOf(listed[0])).deps));
  assert.match([...distinct][0], /^[0-9a-f]{64}$/);
  assert.equal([...distinct][0], PINNED);
});

test('WARRANT: the capability MANIFEST entry (what the client checks) is byte-identical on every listed revision', () => {
  const entries = Object.fromEntries(listed.map(rev => [rev, sha(manifestEntryOf(blockOf(rev)))]));
  assert.equal(new Set(Object.values(entries)).size, 1, JSON.stringify(entries));
  const entry = manifestEntryOf(blockOf(listed[0]));
  assert.ok(entry.includes('kind: "WRITE"') && entry.includes('"TARGETED"'), entry.slice(0, 200));
});

// DS9346: the "every ACCEPTED revision from /7 up carries the handler" staleness guard no longer has subject
// matter -- ACCEPTED_RUNTIMES.xlsx is now the single 9.3.5.0/16 entry; the WARRANT tests above and the 9.3.5.0
// check below already cover the live list directly.

test('NEGATIVE: rev6 (before the print write capability existed) has no handler, so it is NOT listed', () => {
  assert.equal(blockOf(6).includes(`"${CAP}": function (args)`), false, 'rev6 carries no printSettings.write handler');
  assert.equal(listed.includes(6), false);
});

test('RED CONTROL: ONE changed byte in a handler copy changes the digest -- the warrant can go red', () => {
  const src = blockOf(listed[0]);
  const real = digestOf(src);
  const h = handlerOf(src);
  const i = src.indexOf(h);
  const spliceWithin = (needle, replacement) => {
    const mh = h.replace(needle, replacement);
    assert.notEqual(mh, h, `mutation target not found in the handler span: ${needle}`);
    return src.slice(0, i) + mh + src.slice(i + h.length);
  };
  const mutated = spliceWithin('"unknown-setting"', '"unknown-settingX"');
  assert.notEqual(mutated, src, 'the mutation applied');
  assert.notEqual(digestOf(mutated), real);
  assert.notEqual(digestOf(spliceWithin('headerFooterFlag', 'headerFooterFlagX')), real);
  // an unbalanced / missing handler is an error, never a silent "equal"
  assert.throws(() => digestOf('Api.X = function () {}'), /not found/);
});

test('RED CONTROL: a listed revision whose handler differs makes the cross-revision comparison fail', () => {
  const a = digestOf(blockOf(listed[0]));
  const src1 = blockOf(listed[1]);
  const h1 = handlerOf(src1);
  const i1 = src1.indexOf(h1);
  const mh1 = h1.replace('"bad-args"', '"bad-args-x"');
  assert.notEqual(mh1, h1, 'mutation target not found in the handler span');
  const mutated1 = src1.slice(0, i1) + mh1 + src1.slice(i1 + h1.length);
  const b = digestOf(mutated1);
  assert.notEqual(a, b);
});

test('REGRESSION GUARD: the pre-fix /7-/8-only list would refuse /9..15 even though they carry the same handler -- this is the bug zero found before it ever shipped (2026-10-17)', () => {
  const staleList = [7, 8];
  for (const rev of [9, 10, 11, 12, 13, 14, 15]) {
    assert.equal(staleList.includes(rev), false, `rev${rev} wrongly refused by the stale /7-/8-only list`);
    assert.ok(blockOf(rev).includes(`"${CAP}": function (args)`), `rev${rev} actually carries the printSettings.write handler`);
    assert.ok(listed.includes(rev), `the FIXED list accepts rev${rev}`);
  }
});

test('WARRANT (DS v9.3.5 porting, #15): the 9.3.5.0 bridge-block.js carries the SAME handler (sha256 equal to PINNED)', () => {
  const src935 = fs.readFileSync(path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.5.0', 'euro-office-bridge-v1', 'bridge-block.js'), 'utf8');
  assert.equal(digestOf(src935), PINNED);
});

test('RED CONTROL (DS v9.3.5 porting, #15): a changed byte in the 9.3.5.0 handler copy no longer matches PINNED', () => {
  const src935 = fs.readFileSync(path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.5.0', 'euro-office-bridge-v1', 'bridge-block.js'), 'utf8');
  const h = handlerOf(src935);
  const i = src935.indexOf(h);
  const mh = h.replace('"unknown-setting"', '"unknown-settingX"');
  assert.notEqual(mh, h, 'mutation target not found in the handler span');
  const mutated = src935.slice(0, i) + mh + src935.slice(i + h.length);
  assert.notEqual(mutated, src935, 'the mutation applied');
  assert.notEqual(digestOf(mutated), PINNED);
});
