'use strict';

/**
 * THE WARRANT FOR /19 IN THE SIX XLSX BRIDGE-WRITE RUNTIME LISTS (B03/B05/B10/B22/B23/B24) AND THE GENERAL ACCEPTED LIST. A list entry is justified by
 * ONE thing: the handler, and every bridge-local helper it calls, are byte-identical to the already-listed revision's.
 * /19 (GH #332-#335) was meant to add ONLY sheet.table.write/v1 -- this test proves it handler by handler: EVERY handler the /18
 * candidate carries (+ its helper closure) has the same sha256 in the /19 candidate, and /19 carries exactly that one handler
 * more. Without /19 in the six lists a production DS /18 -> /19 apply turns all six Bridge-WRITE families FAIL_CLOSED
 * ('...-bridge-write-runtime-not-proven'). If a future /19 edit changes any pre-existing handler, this goes red.
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
const NEW_IN_19 = ['sheet.table.write/v1'];

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

test('/19 carries every /18 handler plus EXACTLY the one new WRITE handler', () => {
  const c18 = capsOf(blockOf(18)); const c19 = capsOf(blockOf(19));
  assert.ok(c18.length >= 32, `rev18 handlers: ${c18.length}`);
  assert.deepEqual(c19.filter(c => !c18.includes(c)), NEW_IN_19);
  assert.deepEqual(c18.filter(c => !c19.includes(c)), []);
  for (const caps of Object.values(WRITE_CAPS)) for (const cap of caps) assert.ok(c18.includes(cap), cap);
});

test('WARRANT: every /18 handler AND the helpers it calls are byte-identical on /19 (sha256 equal)', () => {
  const s18 = blockOf(18); const s19 = blockOf(19);
  const differing = capsOf(s18).filter(cap => digestOf(s18, cap) !== digestOf(s19, cap));
  assert.deepEqual(differing, []);
});

test('the six Bridge-WRITE lists carry /16 + /17 + /18 + /19, core and adapter alike; the NEW capability list starts at /19 (the /21 extension: ds935-bridge21-write-handler-identity)', () => {
  const expected = ['9.3.5|euro-office-bridge-v1/9.3.5.0/16', '9.3.5|euro-office-bridge-v1/9.3.5.0/17', '9.3.5|euro-office-bridge-v1/9.3.5.0/18', '9.3.5|euro-office-bridge-v1/9.3.5.0/19'];
  for (const k of Object.keys(WRITE_CAPS)) {
    assert.deepEqual([...core[`XLSX_${k}_BRIDGE_WRITE_RUNTIME_IDS`]].slice(0, 4), expected, `core ${k}`);
    assert.deepEqual([...adapter[`XLSX_${k}_BRIDGE_WRITE_RUNTIME_IDS`]].slice(0, 4), expected, `adapter ${k}`);
  }
  assert.deepEqual([...core.XLSX_B25_BRIDGE_WRITE_RUNTIME_IDS][0], expected[3]);
  assert.deepEqual([...adapter.XLSX_B25_BRIDGE_WRITE_RUNTIME_IDS][0], expected[3]);
});

test('RED CONTROL: one changed byte in a /19 write handler, or in a helper it calls, breaks the warrant', () => {
  const s17 = blockOf(18); const s18 = blockOf(19);
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
