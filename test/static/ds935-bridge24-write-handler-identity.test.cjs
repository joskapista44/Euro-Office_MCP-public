'use strict';

/**
 * /24 (GH #338-#342 remainder, workbook.protection.read/v1 + .write/v1; GH #335 remainder, sheet.table.write/v1
 * totalsRow value:false) CARRIES EVERY /23 HANDLER, BYTE-IDENTICAL, WITH EXACTLY ONE NAMED EXCEPTION:
 * sheet.table.write/v1, whose totalsRow branch is the deliberate change this unit makes. Every OTHER handler
 * -- and the layout installer -- must be untouched.
 *
 * The product-side runtime-id widening (the nine Bridge-WRITE lists + the new B28 list + XLSX_RUNTIME_24_ID,
 * the LATER, release-sequenced unit the design document's own "W2 release sequencing" section names) landed
 * in a follow-up unit -- the list-carrying test below is that unit's own addition to THIS file, not the
 * candidate-build unit's.
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

const NEW_IN_24 = ['workbook.protection.read/v1', 'workbook.protection.write/v1'];
const CHANGED_IN_24 = 'sheet.table.write/v1';

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

test('/24 carries every /23 handler plus EXACTLY the two new capabilities', () => {
  const c23 = capsOf(blockOf(23)); const c24 = capsOf(blockOf(24));
  assert.ok(c23.length >= 40, `rev23 handlers: ${c23.length}`);
  assert.deepEqual(c24.filter(c => !c23.includes(c)).sort(), [...NEW_IN_24].sort());
  assert.deepEqual(c23.filter(c => !c24.includes(c)), []);
});

test('WARRANT: every /23 handler AND the helpers it calls are byte-identical on /24, EXCEPT the one named exception', () => {
  const s23 = blockOf(23); const s24 = blockOf(24);
  const differing = capsOf(s23).filter(cap => digestOf(s23, cap) !== digestOf(s24, cap));
  assert.deepEqual(differing, [CHANGED_IN_24], 'the ONLY handler allowed to differ is sheet.table.write/v1 -- its totalsRow branch');
  const inst = s23.slice(0, s23.indexOf('Api.GetEuroOfficeBridge = function () {'));
  assert.ok(inst.length > 100, 'the /23 block starts with the layout installer');
  assert.ok(s24.startsWith(inst), 'the /24 block starts with the SAME installer, byte for byte');
});

test('RED CONTROL: one changed byte in a /24 copy of an UNRELATED /23 handler breaks the warrant', () => {
  const s23 = blockOf(23); const s24 = blockOf(24);
  const cap = 'sheet.protection.write/v1';
  const h = spanFrom(s24, `"${cap}": function (args) {`);
  const mutated = s24.replace(h, h.replace('return', 'return '));
  assert.notEqual(mutated, s24, 'the mutation applied');
  assert.notEqual(digestOf(mutated, cap), digestOf(s23, cap), cap);
});

test('RED CONTROL (the other direction): if sheet.table.write/v1 had NOT changed, the warrant above would wrongly pass it as identical -- prove the real /24 differs there', () => {
  const s23 = blockOf(23); const s24 = blockOf(24);
  assert.notEqual(digestOf(s23, CHANGED_IN_24), digestOf(s24, CHANGED_IN_24), 'sheet.table.write/v1 is NOT byte-identical -- it is the one deliberate change');
});

test('the nine Bridge-WRITE lists carry /24 + /28, core and adapter alike; a NEW B28 list is /24 ONLY (workbook.protection.read/v1 + .write/v1 do not exist on /16..23)', () => {
  const r = n => `9.3.5|euro-office-bridge-v1/9.3.5.0/${n}`;
  const r28 = '9.3.5|euro-office-bridge-v1/9.3.5.0/28';
  const r29 = '9.3.5|euro-office-bridge-v1/9.3.5.0/29'; // GH #365 option C U1 (internal tracker)
  // /28 (GH #361/#363 pre-install readiness gate, internal tracker): the SAME byte-identity warrant, see
  // ds935-bridge28-table-autofilter-criteria-write-handler-identity.test.cjs.
  const eight = [16, 17, 18, 19, 21, 22, 23, 24, 25, 26, 27].map(r).concat(r28, r29);
  for (const k of ['B03', 'B05', 'B10', 'B22', 'B23', 'B24']) {
    assert.deepEqual([...core[`XLSX_${k}_BRIDGE_WRITE_RUNTIME_IDS`]], eight, `core ${k}`);
    assert.deepEqual([...adapter[`XLSX_${k}_BRIDGE_WRITE_RUNTIME_IDS`]], eight, `adapter ${k}`);
  }
  for (const [k, want] of [['B25', [19, 21, 22, 23, 24, 25, 26, 27]], ['B26', [21, 22, 23, 24, 25, 26, 27]], ['B27', [22, 23, 24, 25, 26, 27]], ['B28', [24, 25, 26, 27]]]) {
    assert.deepEqual([...core[`XLSX_${k}_BRIDGE_WRITE_RUNTIME_IDS`]], want.map(r).concat(r28, r29), `core ${k}`);
    assert.deepEqual([...adapter[`XLSX_${k}_BRIDGE_WRITE_RUNTIME_IDS`]], want.map(r).concat(r28, r29), `adapter ${k}`);
  }
  assert.deepEqual([...core.XLSX_RUNTIME_IDS], eight);
  assert.deepEqual([...adapter.XLSX_RUNTIME_IDS], eight);
  assert.deepEqual(core.XLSX_B28_BRIDGE_WRITE_RUNTIME_ENTRIES.map(core.runtimeIdOf), [...core.XLSX_B28_BRIDGE_WRITE_RUNTIME_IDS]);
});

test('RED CONTROL: the new B28 list does NOT widen backwards -- /23 alone is refused by it', () => {
  assert.equal([...adapter.XLSX_B28_BRIDGE_WRITE_RUNTIME_IDS].includes('9.3.5|euro-office-bridge-v1/9.3.5.0/23'), false);
  assert.equal([...core.XLSX_B28_BRIDGE_WRITE_RUNTIME_IDS].includes('9.3.5|euro-office-bridge-v1/9.3.5.0/23'), false);
});
