'use strict';

/**
 * THE WARRANT FOR /22 IN THE EIGHT XLSX BRIDGE-WRITE RUNTIME LISTS (B03/B05/B10/B22/B23/B24 + B25 table write + B26 print layout write).
 * A list entry is justified by ONE thing: the handler, and every bridge-local helper it calls, are byte-identical to the already-listed
 * revision's. /22 (GH #338-#342) = /21 + sheet.protection.write/v1 + sheet.cellProtection.write/v1, no installer. Without /22 in these
 * lists a production DS /21 -> /22 apply turns every Bridge-WRITE family FAIL_CLOSED (the /21 incident, measured LIVE 2026-10-31).
 * If a /22 edit changes any /21 handler -- or the /21 layout installer -- this goes red.
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
});
const NEW_IN_22 = ['sheet.cellProtection.write/v1', 'sheet.protection.write/v1'];

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



test('/22 carries every /21 handler plus EXACTLY the two new WRITE handlers', () => {
  const c21 = capsOf(blockOf(21)); const c22 = capsOf(blockOf(22));
  assert.ok(c21.length >= 37, `rev21 handlers: ${c21.length}`);
  assert.deepEqual(c22.filter(c => !c21.includes(c)), NEW_IN_22);
  assert.deepEqual(c21.filter(c => !c22.includes(c)), []);
  for (const caps of Object.values(WRITE_CAPS)) for (const cap of caps) assert.ok(c21.includes(cap), cap);
});

test('WARRANT: every /21 handler AND the helpers it calls are byte-identical on /22 (sha256 equal); the layout installer too', () => {
  const s21 = blockOf(21); const s22 = blockOf(22);
  const differing = capsOf(s21).filter(cap => digestOf(s21, cap) !== digestOf(s22, cap));
  assert.deepEqual(differing, []);
  const inst = s21.slice(0, s21.indexOf('Api.GetEuroOfficeBridge = function () {'));
  assert.ok(inst.length > 100, 'the /21 block starts with its layout installer');
  assert.ok(s22.startsWith(inst), 'the /22 block starts with the SAME installer, byte for byte');
});

test('the nine Bridge-WRITE lists carry /22 (+ /23 + /24 + /28), core and adapter alike; the protection write list is /22 (+ /23 + /24 + /28)', () => {
  const r = n => `9.3.5|euro-office-bridge-v1/9.3.5.0/${n}`;
  const r28 = '9.3.5|euro-office-bridge-v1/9.3.5.0/28';
  const r29 = '9.3.5|euro-office-bridge-v1/9.3.5.0/29'; // GH #365 option C U1 (internal tracker)
  // /23 (GH #356, #357) and /24 (GH #338-#342 remainder) each carry every prior handler byte-identically
  // (ds935-bridge23-write-handler-identity.test.cjs, ds935-bridge24-write-handler-identity.test.cjs), so the lists grow by both
  // /28 (GH #361/#363 pre-install readiness gate, internal tracker): the SAME byte-identity warrant, see
  // ds935-bridge28-table-autofilter-criteria-write-handler-identity.test.cjs.
  const six = [16, 17, 18, 19, 21, 22, 23, 24, 25, 26, 27].map(r).concat(r28, r29);
  for (const k of ['B03', 'B05', 'B10', 'B22', 'B23', 'B24']) {
    assert.deepEqual([...core[`XLSX_${k}_BRIDGE_WRITE_RUNTIME_IDS`]], six, `core ${k}`);
    assert.deepEqual([...adapter[`XLSX_${k}_BRIDGE_WRITE_RUNTIME_IDS`]], six, `adapter ${k}`);
  }
  for (const [k, want] of [['B25', [19, 21, 22, 23, 24, 25, 26, 27]], ['B26', [21, 22, 23, 24, 25, 26, 27]], ['B27', [22, 23, 24, 25, 26, 27]], ['B28', [24, 25, 26, 27]]]) {
    assert.deepEqual([...core[`XLSX_${k}_BRIDGE_WRITE_RUNTIME_IDS`]], want.map(r).concat(r28, r29), `core ${k}`);
    assert.deepEqual([...adapter[`XLSX_${k}_BRIDGE_WRITE_RUNTIME_IDS`]], want.map(r).concat(r28, r29), `adapter ${k}`);
  }
  assert.deepEqual([...core.XLSX_RUNTIME_IDS], six);
  assert.deepEqual([...adapter.XLSX_RUNTIME_IDS], six);
});

test('RED CONTROL: one changed byte in a /22 copy of a /21 write handler breaks the warrant', () => {
  const s21 = blockOf(21); const s22 = blockOf(22);
  for (const cap of [WRITE_CAPS.B25[0], WRITE_CAPS.B26[0]]) {
    const h = spanFrom(s22, `"${cap}": function (args) {`);
    const mutated = s22.replace(h, h.replace('return', 'return '));
    assert.notEqual(mutated, s22, `the ${cap} mutation applied`);
    assert.notEqual(digestOf(mutated, cap), digestOf(s21, cap), cap);
  }
});
