'use strict';

/**
 * THE WARRANT FOR /21 IN THE SEVEN XLSX BRIDGE-WRITE RUNTIME LISTS (B03/B05/B10/B22/B23/B24 + B25 table write). A list entry is
 * justified by ONE thing: the handler, and every bridge-local helper it calls, are byte-identical to the already-listed revision's.
 * /21 (GH #344/#346/#354/#355) = /19 + the /20 READs (printSettings.read/v2, autoFilter.colorFilter.read/v1, chart.series.style.read/v1)
 * + sheet.printSettings.write/v2 and its load-time Layout installer (types 16-18; every other Layout type goes to the ORIGINAL
 * UndoRedo). Without /21 in these lists a production DS /19 -> /21 apply turns all seven Bridge-WRITE families FAIL_CLOSED
 * ('...-bridge-write-runtime-not-proven') -- measured LIVE on the production pair, 2026-10-31. If a /21 edit changes any /19
 * handler, this goes red.
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
});
const WRITE_CAPS_21 = Object.freeze(Object.assign({}, WRITE_CAPS, { B25: ['sheet.table.write/v1'] }));
const NEW_IN_21 = ['chart.series.style.read/v1', 'sheet.autoFilter.colorFilter.read/v1', 'sheet.printSettings.read/v2', 'sheet.printSettings.write/v2'];

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


test('/21 carries every /19 handler plus EXACTLY the /20 READs and the one new WRITE handler', () => {
  const c19 = capsOf(blockOf(19)); const c21 = capsOf(blockOf(21));
  assert.ok(c19.length >= 33, `rev19 handlers: ${c19.length}`);
  assert.deepEqual(c21.filter(c => !c19.includes(c)), NEW_IN_21);
  assert.deepEqual(c19.filter(c => !c21.includes(c)), []);
  for (const caps of Object.values(WRITE_CAPS_21)) for (const cap of caps) assert.ok(c19.includes(cap), cap);
});

test('WARRANT: every /19 handler AND the helpers it calls are byte-identical on /21 (sha256 equal)', () => {
  const s19 = blockOf(19); const s21 = blockOf(21);
  const differing = capsOf(s19).filter(cap => digestOf(s19, cap) !== digestOf(s21, cap));
  assert.deepEqual(differing, []);
});

test('the Layout installer delegates every type except 16-18 to the ORIGINAL UndoRedo (B10 print settings v1 keeps its path)', () => {
  const s21 = blockOf(21);
  assert.ok(s21.includes('if (Type !== 16 && Type !== 17 && Type !== 18) return orig.call(this, Type, Data, nSheetId, bUndo);'));
});

test('the six Bridge-WRITE lists carry /16 + /17 + /18 + /19 + /21 (+ /22 + /23 + /24 + /28), core and adapter alike; the table write list is /19 + /21 (+ /22 + /23 + /24 + /28)', () => {
  const r = n => `9.3.5|euro-office-bridge-v1/9.3.5.0/${n}`;
  const r28 = '9.3.5|euro-office-bridge-v1/9.3.5.0/28';
  const r29 = '9.3.5|euro-office-bridge-v1/9.3.5.0/29'; // GH #365 option C U1 (internal tracker)
  // /22 (GH #338-#342) and /24 (GH #338-#342 remainder) each carry every prior handler byte-identically
  // (ds935-bridge22-write-handler-identity.test.cjs, ds935-bridge24-write-handler-identity.test.cjs), so the lists grow by both
  // /28 (GH #361/#363 pre-install readiness gate, internal tracker): the SAME byte-identity warrant, see
  // ds935-bridge28-table-autofilter-criteria-write-handler-identity.test.cjs.
  const expected = [16, 17, 18, 19, 21, 22, 23, 24, 25, 26, 27].map(r).concat(r28, r29);
  for (const k of Object.keys(WRITE_CAPS)) {
    assert.deepEqual([...core[`XLSX_${k}_BRIDGE_WRITE_RUNTIME_IDS`]], expected, `core ${k}`);
    assert.deepEqual([...adapter[`XLSX_${k}_BRIDGE_WRITE_RUNTIME_IDS`]], expected, `adapter ${k}`);
  }
  assert.deepEqual([...core.XLSX_B25_BRIDGE_WRITE_RUNTIME_IDS], [r(19), r(21), r(22), r(23), r(24), r(25), r(26), r(27), r28, r29]);
  assert.deepEqual([...adapter.XLSX_B25_BRIDGE_WRITE_RUNTIME_IDS], [r(19), r(21), r(22), r(23), r(24), r(25), r(26), r(27), r28, r29]);
});

test('RED CONTROL: one changed byte in a /21 write handler, or in a helper it calls, breaks the warrant', () => {
  const s19 = blockOf(19); const s21 = blockOf(21);
  const cap = WRITE_CAPS_21.B25[0];
  const h = spanFrom(s21, `"${cap}": function (args) {`);
  const mutated = s21.replace(h, h.replace('return', 'return '));
  assert.notEqual(mutated, s21, 'the handler mutation applied');
  assert.notEqual(digestOf(mutated, cap), digestOf(s19, cap));
  const helper = spanFrom(s21, 'function resolveCell(');
  const capB05 = WRITE_CAPS.B05[0];
  const mutatedHelper = s21.replace(helper, helper.replace('Api.GetSheet(sheet)', 'Api.GetSheet(sheet + "")'));
  assert.notEqual(mutatedHelper, s21, 'the helper mutation applied');
  assert.notEqual(digestOf(mutatedHelper, capB05), digestOf(s19, capB05));
});
