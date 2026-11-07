'use strict';

/**
 * THE WARRANT FOR /18 IN THE SIX XLSX BRIDGE-WRITE RUNTIME LISTS (B03/B05/B10/B22/B23/B24). A list entry is justified by
 * ONE thing: the handler, and every bridge-local helper it calls, are byte-identical to the already-listed revision's.
 * /18 (GH #300 L1 / #298 D) was meant to add ONLY sheet.tables.read/v1 + sheet.hyperlinks.read/v1 -- this test proves it
 * handler by handler: EVERY handler the /17 candidate carries (+ its helper closure) has the same sha256 in the /18
 * candidate, and /18 carries exactly those two handlers more. Without /18 in the six lists a production DS /17 -> /18
 * apply turns all six Bridge-WRITE families FAIL_CLOSED ('...-bridge-write-runtime-not-proven').
 * If a future /18 edit changes any pre-existing handler, this goes red: /18 then leaves the lists (or the change is reviewed).
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');

const core = require('../../src/core/runtime-identity.cjs');
const adapter = require('../../src/adapters/xlsx/runtime.cjs');

const DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.5.0', 'euro-office-bridge-v1');
const blockOf = rev => fs.readFileSync(path.join(DIR, `rev${rev}-candidate`, 'bridge-block.js'), 'utf8');
const sha = text => crypto.createHash('sha256').update(text).digest('hex');

/** the capability of each list -- the handlers the six Bridge-WRITE gates guard */
const WRITE_CAPS = Object.freeze({
  B03: ['sheet.columnWidth.canonicalize/v1'],
  B05: ['range.alignmentExt.write/v1'],
  B10: ['sheet.printSettings.write/v1'],
  B22: ['chart.bar.gapOverlap.write/v1', 'chart.series.order.write/v1'],
  B23: ['chart.drawing.delete.write/v1'],
  B24: ['chart.series.marker.write/v1', 'chart.series.lineStyle.write/v1'],
});
const NEW_IN_18 = ['sheet.hyperlinks.read/v1', 'sheet.tables.read/v1'];

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
/** every bridge-local function the text calls, to a fixed point */
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

test('/18 carries every /17 handler plus EXACTLY the two new READ handlers', () => {
  const c17 = capsOf(blockOf(17)); const c18 = capsOf(blockOf(18));
  assert.ok(c17.length >= 30, `rev17 handlers: ${c17.length}`);
  assert.deepEqual(c18.filter(c => !c17.includes(c)), NEW_IN_18);
  assert.deepEqual(c17.filter(c => !c18.includes(c)), []);
  for (const caps of Object.values(WRITE_CAPS)) for (const cap of caps) assert.ok(c17.includes(cap), cap);
});

test('WARRANT: every /17 handler AND the helpers it calls are byte-identical on /18 (sha256 equal)', () => {
  const s17 = blockOf(17); const s18 = blockOf(18);
  const differing = capsOf(s17).filter(cap => digestOf(s17, cap) !== digestOf(s18, cap));
  assert.deepEqual(differing, []);
});

test('the six Bridge-WRITE lists carry /16 + /17 + /18 (then /19, see ds935-bridge19-write-handler-identity), core and adapter alike', () => {
  const expected = ['9.3.5|euro-office-bridge-v1/9.3.5.0/16', '9.3.5|euro-office-bridge-v1/9.3.5.0/17', '9.3.5|euro-office-bridge-v1/9.3.5.0/18'];
  for (const k of Object.keys(WRITE_CAPS)) {
    assert.deepEqual([...core[`XLSX_${k}_BRIDGE_WRITE_RUNTIME_IDS`]].slice(0, 3), expected, `core ${k}`);
    assert.deepEqual([...adapter[`XLSX_${k}_BRIDGE_WRITE_RUNTIME_IDS`]].slice(0, 3), expected, `adapter ${k}`);
  }
});

test('RED CONTROL: one changed byte in a /18 write handler, or in a helper it calls, breaks the warrant', () => {
  const s17 = blockOf(17); const s18 = blockOf(18);
  const cap = WRITE_CAPS.B05[0];
  const h = spanFrom(s18, `"${cap}": function (args) {`);
  const mutated = s18.replace(h, h.replace('return', 'return '));
  assert.notEqual(mutated, s18, 'the handler mutation applied');
  assert.notEqual(digestOf(mutated, cap), digestOf(s17, cap));
  const helper = spanFrom(s18, 'function resolveCell(');
  const mutatedHelper = s18.replace(helper, helper.replace('Api.GetSheet(sheet)', 'Api.GetSheet(sheet + "")'));
  assert.notEqual(mutatedHelper, s18, 'the helper mutation applied');
  assert.notEqual(digestOf(mutatedHelper, cap), digestOf(s17, cap));
  assert.throws(() => digestOf('Api.X = function () {}', cap), /not found/);
});
