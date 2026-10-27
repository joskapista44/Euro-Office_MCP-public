/**
 * X-O05/5-C sibling (#15, DS v9.3.5 porting) -- THE WARRANT FOR THE B24-WRITE RUNTIME LIST, previously untested.
 * The B24 Bridge WRITE capabilities (chart.series.marker.write/v1 -- symbol/size -- and chart.series.lineStyle.write/v1
 * -- dash/smooth, served through one shared runtime list, XLSX_B24_BRIDGE_WRITE_RUNTIME_ENTRIES) exist on /15 and /16
 * (the runtime-identity.cjs comment above the list: "/16 carries the B24-WRITE handler forward byte-identical --
 * fixed PREEMPTIVELY (not found by a test yet)"). This test proves that claim, and extends it to the DS v9.3.5
 * bridge-block.js added by the #15 porting unit: the handler, and the helpers it calls, are byte-identical on
 * EVERY listed revision AND on the 9.3.5.0 copy (a JS-aware brace match, strings skipped; sha256 equal, pinned).
 * If a future revision differs, this goes red: the revision then does NOT join the list (or the handler change
 * is reviewed first).
 * The B22 sibling (same shared-list-of-two-capabilities shape): test/static/b22-write-handler-identity.test.cjs (X-O05/3).
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');

const DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.4.60', 'euro-office-bridge-v1');
const MARKER = 'chart.series.marker.write/v1';
const LINE = 'chart.series.lineStyle.write/v1';
const sha = text => crypto.createHash('sha256').update(text).digest('hex');
/** the pinned sha256 of (handler + the helpers it calls) for chart.series.marker.write/v1 -- the bytes every listed revision carries (computed from rev15, rev16); a deliberate edit to change it */
const PINNED_MARKER = 'b835e1350603fbe37e0515f665ca45240cdbaec067d3b761ffbe3a528a87724e';
/** the pinned sha256 of (handler + the helpers it calls) for chart.series.lineStyle.write/v1 */
const PINNED_LINE = '76aeda87628d9ae035768ba708fc773b9f063321f99f7d0beb6b3b987f8adb67';

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
const handlerOf = (src, cap) => spanFrom(src, `"${cap}": function (args) {`);
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
const warrantOf = (src, cap) => { const h = handlerOf(src, cap); const deps = dependenciesOf(src, h); return { handler: h, deps: [...deps.entries()].sort(([a], [b]) => (a < b ? -1 : 1)).map(([n, b]) => `${n}:${b}`) }; };
const digestOf = (src, cap) => { const w = warrantOf(src, cap); return sha([w.handler, ...w.deps].join('\n')); };
const manifestEntryOf = (src, cap) => {
  const i = src.indexOf(`{ id: "${cap}"`);
  if (i < 0) throw new Error('manifest entry not found: ' + cap);
  return src.slice(i, closingBrace(src, i) + 1);
};
const blockOf = rev => fs.readFileSync(path.join(DIR, `rev${rev}-candidate`, 'bridge-block.js'), 'utf8');

// DS9346 (owner D5=B, VALTAS, docs/DS9346-RUNTIME-RETIREMENT.md): core.XLSX_B24_BRIDGE_WRITE_RUNTIME_ENTRIES
// now carries ONE entry (9.3.5.0/16) -- the HISTORICAL 9.3.4.60 rev15/rev16 is a frozen, on-disk fact
// (MEASURED, grep '"chart.series.marker.write/v1": function' across rev6..16-candidate: present only on
// rev15/rev16) independent of what is currently accepted. The 9.3.5.0 copy is checked separately below,
// against the pinned digests.
const listed = [15, 16];

test('the listed revisions are exactly /15, /16, and each carries both handlers', () => {
  assert.deepEqual(listed, [15, 16]);
  for (const rev of listed) {
    assert.ok(fs.existsSync(path.join(DIR, `rev${rev}-candidate`, 'bridge-block.js')), `rev${rev}-candidate/bridge-block.js`);
    const src = blockOf(rev);
    assert.ok(handlerOf(src, MARKER).length > 200, `rev${rev} marker handler shape`);
    assert.ok(handlerOf(src, LINE).length > 200, `rev${rev} lineStyle handler shape`);
  }
});

test('WARRANT: the marker handler AND the helpers it calls are byte-identical on EVERY listed revision (sha256 equal, pinned)', () => {
  const digests = Object.fromEntries(listed.map(rev => [rev, digestOf(blockOf(rev), MARKER)]));
  const distinct = new Set(Object.values(digests));
  assert.equal(distinct.size, 1, `the marker handler differs between listed revisions: ${JSON.stringify(digests)}`);
  assert.match([...distinct][0], /^[0-9a-f]{64}$/);
  assert.equal([...distinct][0], PINNED_MARKER);
});

test('WARRANT: the lineStyle handler AND the helpers it calls are byte-identical on EVERY listed revision (sha256 equal, pinned)', () => {
  const digests = Object.fromEntries(listed.map(rev => [rev, digestOf(blockOf(rev), LINE)]));
  const distinct = new Set(Object.values(digests));
  assert.equal(distinct.size, 1, `the lineStyle handler differs between listed revisions: ${JSON.stringify(digests)}`);
  assert.match([...distinct][0], /^[0-9a-f]{64}$/);
  assert.equal([...distinct][0], PINNED_LINE);
});

test("WARRANT: both capabilities' MANIFEST entries (what the client checks) are byte-identical on every listed revision", () => {
  for (const cap of [MARKER, LINE]) {
    const entries = Object.fromEntries(listed.map(rev => [rev, sha(manifestEntryOf(blockOf(rev), cap))]));
    assert.equal(new Set(Object.values(entries)).size, 1, `${cap}: ${JSON.stringify(entries)}`);
    const entry = manifestEntryOf(blockOf(listed[0]), cap);
    assert.ok(entry.includes('kind: "WRITE"') && entry.includes('TARGETED'), entry.slice(0, 200));
  }
});

// DS9346: the "every ACCEPTED revision from /15 up carries either handler" staleness guard no longer has
// subject matter -- ACCEPTED_RUNTIMES.xlsx is now the single 9.3.5.0/16 entry; the WARRANT tests above and
// the 9.3.5.0 check below already cover the live list directly.

test('NEGATIVE: /6 .. /14 (before the marker/lineStyle capabilities) have no handler, so they are NOT listed', () => {
  for (const rev of [6, 7, 8, 9, 10, 11, 12, 13, 14]) {
    const src = blockOf(rev);
    assert.equal(src.includes(MARKER), false, `rev${rev} carries no marker handler`);
    assert.equal(src.includes(LINE), false, `rev${rev} carries no lineStyle handler`);
    assert.equal(listed.includes(rev), false, `rev${rev}`);
  }
});

test('RED CONTROL: ONE changed byte in a marker handler copy changes the digest -- the warrant can go red', () => {
  const src = blockOf(listed[0]);
  const real = digestOf(src, MARKER);
  const h = handlerOf(src, MARKER);
  const i = src.indexOf(h);
  const mh = h.replace('"invalid-symbol"', '"invalid-symbolX"');
  assert.notEqual(mh, h, 'mutation target not found in the handler span');
  const mutated = src.slice(0, i) + mh + src.slice(i + h.length);
  assert.notEqual(mutated, src, 'the mutation applied');
  assert.notEqual(digestOf(mutated, MARKER), real);
  assert.throws(() => digestOf('Api.X = function () {}', MARKER), /not found/);
});

test('RED CONTROL: a listed revision whose handler differs makes the cross-revision comparison fail', () => {
  const a = digestOf(blockOf(listed[0]), LINE);
  const src1 = blockOf(listed[1]);
  const h1 = handlerOf(src1, LINE);
  const i1 = src1.indexOf(h1);
  const mh1 = h1.replace('"invalid-dash"', '"invalid-dashX"');
  assert.notEqual(mh1, h1, 'mutation target not found in the handler span');
  const mutated1 = src1.slice(0, i1) + mh1 + src1.slice(i1 + h1.length);
  const b = digestOf(mutated1, LINE);
  assert.notEqual(a, b);
});

test('WARRANT (DS v9.3.5 porting, #15): the 9.3.5.0 bridge-block.js carries the SAME handlers for both capabilities (sha256 equal to PINNED_MARKER/PINNED_LINE)', () => {
  const src935 = fs.readFileSync(path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.5.0', 'euro-office-bridge-v1', 'bridge-block.js'), 'utf8');
  assert.equal(digestOf(src935, MARKER), PINNED_MARKER);
  assert.equal(digestOf(src935, LINE), PINNED_LINE);
});

test('RED CONTROL (DS v9.3.5 porting, #15): a changed byte in the 9.3.5.0 marker handler copy no longer matches PINNED_MARKER', () => {
  const src935 = fs.readFileSync(path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.5.0', 'euro-office-bridge-v1', 'bridge-block.js'), 'utf8');
  const h = handlerOf(src935, MARKER);
  const i = src935.indexOf(h);
  const mh = h.replace('"invalid-symbol"', '"invalid-symbolX"');
  assert.notEqual(mh, h, 'mutation target not found in the handler span');
  const mutated = src935.slice(0, i) + mh + src935.slice(i + h.length);
  assert.notEqual(mutated, src935, 'the mutation applied');
  assert.notEqual(digestOf(mutated, MARKER), PINNED_MARKER);
});
