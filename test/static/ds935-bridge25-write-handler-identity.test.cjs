'use strict';

/**
 * /25 (W2 regression fix: workbook.protection.read/v1's `v === null` shape now answers {lockStructure:false}
 * instead of the generic protection-flag-not-boolean refusal /24 gave it) CARRIES EVERY /24 HANDLER,
 * BYTE-IDENTICAL, WITH EXACTLY ONE NAMED EXCEPTION: workbook.protection.read/v1 itself, whose null branch is the
 * deliberate change this unit makes. Every OTHER handler -- including workbook.protection.write/v1, and the
 * layout installer -- must be untouched. NO capability is added or removed.
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

const CHANGED_IN_25 = 'workbook.protection.read/v1';

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

test('/25 carries EXACTLY the same capability set as /24 -- no capability added or removed', () => {
  const c24 = capsOf(blockOf(24)); const c25 = capsOf(blockOf(25));
  assert.ok(c24.length >= 40, `rev24 handlers: ${c24.length}`);
  assert.deepEqual(c25.filter(c => !c24.includes(c)), [], 'nothing new on /25');
  assert.deepEqual(c24.filter(c => !c25.includes(c)), [], 'nothing dropped on /25');
});

test('WARRANT: every /24 handler AND the helpers it calls are byte-identical on /25, EXCEPT the one named exception', () => {
  const s24 = blockOf(24); const s25 = blockOf(25);
  const differing = capsOf(s24).filter(cap => digestOf(s24, cap) !== digestOf(s25, cap));
  assert.deepEqual(differing, [CHANGED_IN_25], 'the ONLY handler allowed to differ is workbook.protection.read/v1 -- its null branch');
  const inst = s24.slice(0, s24.indexOf('Api.GetEuroOfficeBridge = function () {'));
  assert.ok(inst.length > 100, 'the /24 block starts with the layout installer');
  assert.ok(s25.startsWith(inst), 'the /25 block starts with the SAME installer, byte for byte');
});

test('workbook.protection.write/v1 -- the OTHER capability this unit touches no code in -- is untouched on /25', () => {
  const s24 = blockOf(24); const s25 = blockOf(25);
  assert.equal(digestOf(s24, 'workbook.protection.write/v1'), digestOf(s25, 'workbook.protection.write/v1'));
});

test('RED CONTROL: one changed byte in a /25 copy of an UNRELATED /24 handler breaks the warrant', () => {
  const s24 = blockOf(24); const s25 = blockOf(25);
  const cap = 'sheet.protection.write/v1';
  const h = spanFrom(s25, `"${cap}": function (args) {`);
  const mutated = s25.replace(h, h.replace('return', 'return '));
  assert.notEqual(mutated, s25, 'the mutation applied');
  assert.notEqual(digestOf(mutated, cap), digestOf(s24, cap), cap);
});

test('RED CONTROL (the other direction): if workbook.protection.read/v1 had NOT changed, the warrant above would wrongly pass it as identical -- prove the real /25 differs there', () => {
  const s24 = blockOf(24); const s25 = blockOf(25);
  assert.notEqual(digestOf(s24, CHANGED_IN_25), digestOf(s25, CHANGED_IN_25), 'workbook.protection.read/v1 is NOT byte-identical -- it is the one deliberate change');
});

test('the null branch answers {lockStructure:false}, every other non-boolean value is unchanged (fail-closed)', () => {
  const h = spanFrom(blockOf(25), `"${CHANGED_IN_25}": function (args) {`);
  assert.match(h, /if \(v === null\) return \{ ok: true, value: \{ lockStructure: false \} \};/);
  assert.match(h, /if \(typeof v !== "boolean"\) return \{ ok: false, reason: "protection-flag-not-boolean" \};/, 'the generic non-boolean refusal survives, for every OTHER shape');
});

test('the ten Bridge-WRITE lists carry /25 + /28, core and adapter alike; B28 (workbook.protection.read/v1 + .write/v1) widens from [/24] to [/24, /25] -- and (review fix, f0451f47) once more to [/24, /25, /26]', () => {
  const r = n => `9.3.5|euro-office-bridge-v1/9.3.5.0/${n}`;
  const r28 = '9.3.5|euro-office-bridge-v1/9.3.5.0/28';
  const r29 = '9.3.5|euro-office-bridge-v1/9.3.5.0/29'; // GH #365 option C U1 (internal tracker)
  // /28 (GH #361/#363 pre-install readiness gate, internal tracker): the SAME byte-identity warrant, see
  // ds935-bridge28-table-autofilter-criteria-write-handler-identity.test.cjs.
  const ten = [16, 17, 18, 19, 21, 22, 23, 24, 25, 26, 27].map(r).concat(r28, r29);
  for (const k of ['B03', 'B05', 'B10', 'B22', 'B23', 'B24']) {
    assert.deepEqual([...core[`XLSX_${k}_BRIDGE_WRITE_RUNTIME_IDS`]], ten, `core ${k}`);
    assert.deepEqual([...adapter[`XLSX_${k}_BRIDGE_WRITE_RUNTIME_IDS`]], ten, `adapter ${k}`);
  }
  for (const [k, want] of [['B25', [19, 21, 22, 23, 24, 25, 26, 27]], ['B26', [21, 22, 23, 24, 25, 26, 27]], ['B27', [22, 23, 24, 25, 26, 27]], ['B28', [24, 25, 26, 27]]]) {
    assert.deepEqual([...core[`XLSX_${k}_BRIDGE_WRITE_RUNTIME_IDS`]], want.map(r).concat(r28, r29), `core ${k}`);
    assert.deepEqual([...adapter[`XLSX_${k}_BRIDGE_WRITE_RUNTIME_IDS`]], want.map(r).concat(r28, r29), `adapter ${k}`);
  }
  assert.deepEqual([...core.XLSX_RUNTIME_IDS], ten);
  assert.deepEqual([...adapter.XLSX_RUNTIME_IDS], ten);
  assert.deepEqual([...core.ACCEPTED_RUNTIMES.xlsx.map(core.runtimeIdOf)], ten);
});

test('RED CONTROL: B28 does not widen backwards -- /23 alone is still refused by it', () => {
  assert.equal([...adapter.XLSX_B28_BRIDGE_WRITE_RUNTIME_IDS].includes('9.3.5|euro-office-bridge-v1/9.3.5.0/23'), false);
  assert.equal([...core.XLSX_B28_BRIDGE_WRITE_RUNTIME_IDS].includes('9.3.5|euro-office-bridge-v1/9.3.5.0/23'), false);
});

test('the totals-off-accepting runtime list (table-props.cjs TOTALS_OFF_BRIDGE_RUNTIMES) widens from [/24] to [/24, /25] -- and (review fix, f0451f47: sheet.table.write/v1\'s totalsRow branch outside the touched guard) once more to [/24, /25, /26]', () => {
  const { XLSX_RUNTIME_24_ID, XLSX_RUNTIME_25_ID, XLSX_RUNTIME_26_ID, XLSX_RUNTIME_27_ID, XLSX_RUNTIME_28_ID, XLSX_RUNTIME_29_ID } = adapter;
  const tp = require('../../src/adapters/xlsx/families/table-props.cjs');
  // GH #346: + /27 (sheet.table.write/v1 byte-identical, warrant in ds935-bridge27-write-handler-identity.test.cjs)
  // GH #361/#363 pre-install readiness gate (internal tracker): + /28 (SAME warrant, ds935-bridge28-table-
  // autofilter-criteria-write-handler-identity.test.cjs)
  assert.deepEqual([...tp.TOTALS_OFF_BRIDGE_RUNTIMES], [XLSX_RUNTIME_24_ID, XLSX_RUNTIME_25_ID, XLSX_RUNTIME_26_ID, XLSX_RUNTIME_27_ID, XLSX_RUNTIME_28_ID, XLSX_RUNTIME_29_ID]);
});

test('RED CONTROL: TOTALS_OFF_BRIDGE_RUNTIMES does not widen backwards -- /23 alone is still refused by it', () => {
  const tp = require('../../src/adapters/xlsx/families/table-props.cjs');
  assert.equal([...tp.TOTALS_OFF_BRIDGE_RUNTIMES].includes('9.3.5|euro-office-bridge-v1/9.3.5.0/23'), false);
});
