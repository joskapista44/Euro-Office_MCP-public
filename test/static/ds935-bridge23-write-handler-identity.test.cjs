'use strict';

/**
 * THE WARRANT FOR /23 IN THE NINE XLSX BRIDGE-WRITE RUNTIME LISTS (B03/B05/B10/B22/B23/B24 + B25 table + B26 print layout + B27 protection).
 * A list entry is justified by ONE thing: the handler, and every bridge-local helper it calls, are byte-identical to the already-listed
 * revision's. /23 (GH #356, #357) = /22 + sheet.printPageCount.read/v1, no installer, no write. Without /23 in these lists a production
 * DS /22 -> /23 apply turns every Bridge-WRITE family FAIL_CLOSED (the /21 incident, measured LIVE 2026-10-31).
 * If a /23 edit changes any /22 handler -- or the layout installer -- this goes red.
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

/** the capability of each list -- the handlers the Bridge-WRITE gates guard */
const WRITE_CAPS = Object.freeze({
  B03: ['sheet.columnWidth.canonicalize/v1'],
  B05: ['range.alignmentExt.write/v1'],
  B10: ['sheet.printSettings.write/v1'],
  B22: ['chart.bar.gapOverlap.write/v1', 'chart.series.order.write/v1'],
  B23: ['chart.drawing.delete.write/v1'],
  B24: ['chart.series.marker.write/v1', 'chart.series.lineStyle.write/v1'],
  B25: ['sheet.table.write/v1'],
  B26: ['sheet.printSettings.write/v2'],
  B27: ['sheet.protection.write/v1', 'sheet.cellProtection.write/v1'],
});
const NEW_IN_23 = ['sheet.printPageCount.read/v1'];

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



test('/23 carries every /22 handler plus EXACTLY the one new READ handler', () => {
  const c22 = capsOf(blockOf(22)); const c23 = capsOf(blockOf(23));
  assert.ok(c22.length >= 39, `rev22 handlers: ${c22.length}`);
  assert.deepEqual(c23.filter(c => !c22.includes(c)), NEW_IN_23);
  assert.deepEqual(c22.filter(c => !c23.includes(c)), []);
  for (const caps of Object.values(WRITE_CAPS)) for (const cap of caps) assert.ok(c22.includes(cap), cap);
});

test('WARRANT: every /22 handler AND the helpers it calls are byte-identical on /23 (sha256 equal); the layout installer too', () => {
  const s22 = blockOf(22); const s23 = blockOf(23);
  const differing = capsOf(s22).filter(cap => digestOf(s22, cap) !== digestOf(s23, cap));
  assert.deepEqual(differing, []);
  const inst = s22.slice(0, s22.indexOf('Api.GetEuroOfficeBridge = function () {'));
  assert.ok(inst.length > 100, 'the /22 block starts with the layout installer');
  assert.ok(s23.startsWith(inst), 'the /23 block starts with the SAME installer, byte for byte');
});

test('the nine Bridge-WRITE lists carry /23 (+ /24 + /28), core and adapter alike', () => {
  const r = n => `9.3.5|euro-office-bridge-v1/9.3.5.0/${n}`;
  const r28 = '9.3.5|euro-office-bridge-v1/9.3.5.0/28';
  const r29 = '9.3.5|euro-office-bridge-v1/9.3.5.0/29'; // GH #365 option C U1 (internal tracker)
  // /24 (GH #338-#342 remainder) carries every /23 handler byte-identically except the one named exception
  // (ds935-bridge24-write-handler-identity.test.cjs), so the lists grow by it too
  // /28 (GH #361/#363 pre-install readiness gate, internal tracker): the SAME byte-identity warrant, see
  // ds935-bridge28-table-autofilter-criteria-write-handler-identity.test.cjs.
  const seven = [16, 17, 18, 19, 21, 22, 23, 24, 25, 26, 27].map(r).concat(r28, r29);
  for (const k of ['B03', 'B05', 'B10', 'B22', 'B23', 'B24']) {
    assert.deepEqual([...core[`XLSX_${k}_BRIDGE_WRITE_RUNTIME_IDS`]], seven, `core ${k}`);
    assert.deepEqual([...adapter[`XLSX_${k}_BRIDGE_WRITE_RUNTIME_IDS`]], seven, `adapter ${k}`);
  }
  for (const [k, want] of [['B25', [19, 21, 22, 23, 24, 25, 26, 27]], ['B26', [21, 22, 23, 24, 25, 26, 27]], ['B27', [22, 23, 24, 25, 26, 27]], ['B28', [24, 25, 26, 27]]]) {
    assert.deepEqual([...core[`XLSX_${k}_BRIDGE_WRITE_RUNTIME_IDS`]], want.map(r).concat(r28, r29), `core ${k}`);
    assert.deepEqual([...adapter[`XLSX_${k}_BRIDGE_WRITE_RUNTIME_IDS`]], want.map(r).concat(r28, r29), `adapter ${k}`);
  }
  assert.deepEqual([...core.XLSX_RUNTIME_IDS], seven);
  assert.deepEqual([...adapter.XLSX_RUNTIME_IDS], seven);
});

test('RED CONTROL: one changed byte in a /23 copy of a /22 write handler breaks the warrant', () => {
  const s22 = blockOf(22); const s23 = blockOf(23);
  for (const cap of [WRITE_CAPS.B26[0], WRITE_CAPS.B27[0]]) {
    const h = spanFrom(s23, `"${cap}": function (args) {`);
    const mutated = s23.replace(h, h.replace('return', 'return '));
    assert.notEqual(mutated, s23, `the ${cap} mutation applied`);
    assert.notEqual(digestOf(mutated, cap), digestOf(s22, cap), cap);
  }
});
