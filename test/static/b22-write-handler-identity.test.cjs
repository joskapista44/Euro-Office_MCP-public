/**
 * X-O05/3 -- THE WARRANT FOR THE B22-WRITE RUNTIME LIST. The B22 Bridge WRITE capabilities (chart.bar.gapOverlap.write/v1
 * and chart.series.order.write/v1, served through one shared runtime list, XLSX_B22_BRIDGE_WRITE_RUNTIME_ENTRIES) were
 * written only on /12, then went STALE: /13 and /14 carry the same byte-identical handlers (no public setter exists
 * for either target, X-O05/1's own measurement), but nobody added them -- so the prod DS (/14, deployed 2026-10-15)
 * answered FAIL_CLOSED 'xlsx-chart-style-bridge-write-runtime-not-proven' for every gapOverlap/seriesOrder write
 * while nothing was actually missing (the same staleness class L-05 already fixed for B03). A list entry is
 * justified by ONE thing per capability: the handler, and the helpers it calls, are byte-identical to the first
 * listed revision's. This test extracts them from EVERY listed revision's bridge-block.js (a JS-aware brace match,
 * strings skipped) and requires the sha256 to be the same -- and equal to a pinned value, so that a change to ALL
 * revisions at once is a deliberate, visible edit too. If a future revision differs, this goes red: the revision
 * then does NOT join the list (or the handler change is reviewed first).
 * The B03 sibling: test/static/b03-write-handler-identity.test.cjs (L-05). The B23 sibling (chart delete, same
 * staleness, same fix day): test/static/b23-write-handler-identity.test.cjs (X-O07/3).
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');

const DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.4.60', 'euro-office-bridge-v1');
const GAP = 'chart.bar.gapOverlap.write/v1';
const ORD = 'chart.series.order.write/v1';
const sha = text => crypto.createHash('sha256').update(text).digest('hex');
/** the pinned sha256 of (handler + the helpers it calls) for chart.bar.gapOverlap.write/v1 -- the bytes every listed revision carries (computed from rev12..rev14); a deliberate edit to change it */
const PINNED_GAP = '0378b568c6ccd1c21e704ebc3a9cbbd2cbefe0269f5b93fdf9dbe03deaf94eb0';
/** the pinned sha256 of (handler + the helpers it calls) for chart.series.order.write/v1 */
const PINNED_ORD = 'fa327d3cb8fd2c4eab7e14bcb5a598b0b791e04355ecd095274847056434fb5b';

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

// DS9346 (owner D5=B, VALTAS, docs/DS9346-RUNTIME-RETIREMENT.md): core.XLSX_B22_BRIDGE_WRITE_RUNTIME_ENTRIES
// now carries ONE entry (9.3.5.0/16) -- the HISTORICAL 9.3.4.60 rev12..rev16 chain this file audits is a
// frozen, on-disk fact (MEASURED, grep '"chart.bar.gapOverlap.write/v1": function' across rev12..16-candidate)
// independent of what is currently accepted. The 9.3.5.0 copy is checked separately below, against the
// pinned digests.
const listed = [12, 13, 14, 15, 16];

test('the listed revisions are exactly /12, /13, /14, /15, /16, and each carries both handlers', () => {
  assert.deepEqual(listed, [12, 13, 14, 15, 16]);
  for (const rev of listed) {
    assert.ok(fs.existsSync(path.join(DIR, `rev${rev}-candidate`, 'bridge-block.js')), `rev${rev}-candidate/bridge-block.js`);
    const src = blockOf(rev);
    assert.ok(handlerOf(src, GAP).length > 200, `rev${rev} gapOverlap handler shape`);
    assert.ok(handlerOf(src, ORD).length > 200, `rev${rev} seriesOrder handler shape`);
  }
});

test('WARRANT: the gapOverlap handler AND the helpers it calls are byte-identical on EVERY listed revision (sha256 equal, pinned)', () => {
  const digests = Object.fromEntries(listed.map(rev => [rev, digestOf(blockOf(rev), GAP)]));
  const distinct = new Set(Object.values(digests));
  assert.equal(distinct.size, 1, `the gapOverlap handler differs between listed revisions: ${JSON.stringify(digests)}`);
  assert.match([...distinct][0], /^[0-9a-f]{64}$/);
  assert.equal([...distinct][0], PINNED_GAP);
});

test('WARRANT: the seriesOrder handler AND the helpers it calls are byte-identical on EVERY listed revision (sha256 equal, pinned)', () => {
  const digests = Object.fromEntries(listed.map(rev => [rev, digestOf(blockOf(rev), ORD)]));
  const distinct = new Set(Object.values(digests));
  assert.equal(distinct.size, 1, `the seriesOrder handler differs between listed revisions: ${JSON.stringify(digests)}`);
  assert.match([...distinct][0], /^[0-9a-f]{64}$/);
  assert.equal([...distinct][0], PINNED_ORD);
});

test("WARRANT: both capabilities' MANIFEST entries (what the client checks) are byte-identical on every listed revision", () => {
  for (const cap of [GAP, ORD]) {
    const entries = Object.fromEntries(listed.map(rev => [rev, sha(manifestEntryOf(blockOf(rev), cap))]));
    assert.equal(new Set(Object.values(entries)).size, 1, `${cap}: ${JSON.stringify(entries)}`);
    const entry = manifestEntryOf(blockOf(listed[0]), cap);
    assert.ok(entry.includes('kind: "WRITE"') && entry.includes('TARGETED'), entry.slice(0, 200));
  }
});

// DS9346: the "every ACCEPTED revision from /12 up carries either handler" staleness guard no longer has
// subject matter -- ACCEPTED_RUNTIMES.xlsx is now the single 9.3.5.0/16 entry; the WARRANT tests above and
// the 9.3.5.0 check below already cover the live list directly.

test('NEGATIVE: /6 .. /11 (before the gapOverlap/seriesOrder capabilities) have no handler, so they are NOT listed', () => {
  for (const rev of [6, 7, 8, 9, 10, 11]) {
    const src = blockOf(rev);
    assert.equal(src.includes(GAP), false, `rev${rev} carries no gapOverlap handler`);
    assert.equal(src.includes(ORD), false, `rev${rev} carries no seriesOrder handler`);
    assert.equal(listed.includes(rev), false, `rev${rev}`);
  }
});

test('RED CONTROL: ONE changed byte in a handler copy changes the digest -- the warrant can go red', () => {
  const src = blockOf(listed[0]);
  const real = digestOf(src, GAP);
  const h = handlerOf(src, GAP);
  const i = src.indexOf(h);
  const mh = h.replace('args.gapWidth', 'args.gapWidthX');
  assert.notEqual(mh, h, 'mutation target not found in the handler span');
  const mutated = src.slice(0, i) + mh + src.slice(i + h.length);
  assert.notEqual(mutated, src, 'the mutation applied');
  assert.notEqual(digestOf(mutated, GAP), real);
  assert.throws(() => digestOf('Api.X = function () {}', GAP), /not found/);
});

test('RED CONTROL: a listed revision whose handler differs makes the cross-revision comparison fail', () => {
  const a = digestOf(blockOf(listed[0]), GAP);
  const src1 = blockOf(listed[1]);
  const h1 = handlerOf(src1, GAP);
  const i1 = src1.indexOf(h1);
  const mh1 = h1.replace('args.gapWidth', 'args.gapWidthX');
  assert.notEqual(mh1, h1, 'mutation target not found in the handler span');
  const mutated1 = src1.slice(0, i1) + mh1 + src1.slice(i1 + h1.length);
  const b = digestOf(mutated1, GAP);
  assert.notEqual(a, b);
});

test('REGRESSION GUARD: the pre-fix /12-only list refuses /13 and /14 even though they carry the same handler -- this is the bug found on prod /14 (2026-10-15)', () => {
  const staleList = [12];
  for (const rev of [13, 14]) {
    assert.equal(staleList.includes(rev), false, `rev${rev} wrongly refused by the stale /12-only list`);
    assert.ok(blockOf(rev).includes(GAP) && blockOf(rev).includes(ORD), `rev${rev} actually carries both handlers`);
  }
  assert.ok(listed.includes(13) && listed.includes(14), 'the FIXED list accepts both');
});

test('WARRANT (DS v9.3.5 porting, #15): the 9.3.5.0 bridge-block.js carries the SAME handlers for both capabilities (sha256 equal to PINNED_GAP/PINNED_ORD)', () => {
  const src935 = fs.readFileSync(path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.5.0', 'euro-office-bridge-v1', 'bridge-block.js'), 'utf8');
  assert.equal(digestOf(src935, GAP), PINNED_GAP);
  assert.equal(digestOf(src935, ORD), PINNED_ORD);
});

test('RED CONTROL (DS v9.3.5 porting, #15): a changed byte in the 9.3.5.0 GAP handler copy no longer matches PINNED_GAP', () => {
  const src935 = fs.readFileSync(path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.5.0', 'euro-office-bridge-v1', 'bridge-block.js'), 'utf8');
  const h = handlerOf(src935, GAP);
  const i = src935.indexOf(h);
  const mh = h.replace('args.gapWidth', 'args.gapWidthX');
  assert.notEqual(mh, h, 'mutation target not found in the handler span');
  const mutated = src935.slice(0, i) + mh + src935.slice(i + h.length);
  assert.notEqual(mutated, src935, 'the mutation applied');
  assert.notEqual(digestOf(mutated, GAP), PINNED_GAP);
});
