'use strict';

/**
 * GH #346 (/27 candidate): /27 CARRIES EVERY /26 HANDLER, BYTE-IDENTICAL, WITH EXACTLY ONE NAMED CHANGE --
 * chart.series.style.read/v1 (the dataPoint branch answers effective + duplicated, unit 1) -- and EXACTLY ONE
 * NEW capability, chart.series.dataPoint.style.write/v1 (unit 2). This is the warrant every capability-specific
 * runtime list that is NOT a Bridge-WRITE family list needs before it may carry /27 -- found missing on the
 * the production host test DS (Bridge /27) family regression: table-props.cjs TOTALS_OFF_BRIDGE_RUNTIMES stopped at /26, so
 * set_table_property totalsRow:false was refused as xlsx-table-totals-off-bridge-runtime-not-proven on /27.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');

const DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.5.0', 'euro-office-bridge-v1');
const blockOf = rev => fs.readFileSync(path.join(DIR, `rev${rev}-candidate`, 'bridge-block.js'), 'utf8');
const sha = text => crypto.createHash('sha256').update(text).digest('hex');

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

const CHANGED_IN_27 = 'chart.series.style.read/v1';
const NEW_IN_27 = 'chart.series.dataPoint.style.write/v1';

test('/27 = the /26 capability set plus exactly one new capability, none dropped', () => {
  const c26 = capsOf(blockOf(26)); const c27 = capsOf(blockOf(27));
  assert.ok(c26.length >= 40, `rev26 handlers: ${c26.length}`);
  assert.deepEqual(c27.filter(c => !c26.includes(c)), [NEW_IN_27], 'the only new capability');
  assert.deepEqual(c26.filter(c => !c27.includes(c)), [], 'nothing dropped on /27');
});

test('WARRANT: every /26 handler AND the helpers it calls are byte-identical on /27, EXCEPT chart.series.style.read/v1', () => {
  const s26 = blockOf(26); const s27 = blockOf(27);
  const differing = capsOf(s26).filter(cap => digestOf(s26, cap) !== digestOf(s27, cap));
  assert.deepEqual(differing, [CHANGED_IN_27]);
  const inst = s26.slice(0, s26.indexOf('Api.GetEuroOfficeBridge = function () {'));
  assert.ok(inst.length > 100, 'the /26 block starts with the layout installer');
  assert.ok(s27.startsWith(inst), 'the /27 block starts with the SAME installer, byte for byte');
});

test('sheet.table.write/v1 (the totalsRow value:false branch) is byte-identical on /27, so TOTALS_OFF_BRIDGE_RUNTIMES carries /27', () => {
  const s26 = blockOf(26); const s27 = blockOf(27);
  assert.equal(digestOf(s26, 'sheet.table.write/v1'), digestOf(s27, 'sheet.table.write/v1'));
  const tp = require('../../src/adapters/xlsx/families/table-props.cjs');
  const r = n => `9.3.5|euro-office-bridge-v1/9.3.5.0/${n}`;
  const r28 = '9.3.5|euro-office-bridge-v1/9.3.5.0/28';
  const r29 = '9.3.5|euro-office-bridge-v1/9.3.5.0/29'; // GH #365 option C U1 (internal tracker)
  // + /28 (GH #361/#363 pre-install readiness gate, internal tracker): SAME byte-identity warrant, see
  // ds935-bridge28-table-autofilter-criteria-write-handler-identity.test.cjs.
  assert.deepEqual([...tp.TOTALS_OFF_BRIDGE_RUNTIMES], [24, 25, 26, 27].map(r).concat(r28, r29));
});

test('RED CONTROL: one changed byte in a /27 copy of sheet.table.write/v1 breaks the warrant', () => {
  const s26 = blockOf(26); const s27 = blockOf(27);
  const cap = 'sheet.table.write/v1';
  const h = spanFrom(s27, `"${cap}": function (args) {`);
  const mutated = s27.replace(h, h.replace('return', 'return '));
  assert.notEqual(mutated, s27, 'the mutation applied');
  assert.notEqual(digestOf(mutated, cap), digestOf(s26, cap));
});

test('RED CONTROL (other direction): chart.series.style.read/v1 really differs on /27', () => {
  assert.notEqual(digestOf(blockOf(26), CHANGED_IN_27), digestOf(blockOf(27), CHANGED_IN_27));
});
