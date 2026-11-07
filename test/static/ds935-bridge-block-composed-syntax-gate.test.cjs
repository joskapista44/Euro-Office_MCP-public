'use strict';

/**
 * GH #361/#363 URGENT fix (kártya 54ab1854, zero lelete 2026-11-06 07:02): a /28 jelölt
 * bridge-block.js-ét a the production host-en betöltő Document Server eltört -- a kiszolgált cell sdk-all.js
 * `node --check`-je `SyntaxError: Unexpected token ':'` (a /27-es prod referencia, e39658a4,
 * ugyanott rc=0). GYÖKÉR: build-rev28-table-autofilter-criteria.py CAPS_END_ANCHOR a handlereket
 * a READ-objektum (EO_BRIDGE_CAPS) SAJÁT záró `};`-je UTÁN szúrta be -- a meglévő "warrant" teszt
 * (ds935-bridge28-table-autofilter-criteria-write-handler-identity.test.cjs) ezt NEM látja, mert
 * az NEVESÍTETT handler-span-eket parszolja ki és ellenőrzi, az ÖSSZEÁLLÍTOTT FÁJLT SOHA.
 *
 * EZ A LAP AZ ÖSSZEÁLLÍTOTT-BLOKK SZINTAXIS-KAPU: minden /25-/28 jelöltre (a kártya saját
 * hatóköre -- rev25/rev26/rev27/rev28-table-autofilter-criteria) ellenőrzi, hogy a tárolt
 * bridge-block.js parszolható JS-e, egy minimális, a valódi beillesztési kontextust (install_
 * bridge.py plan_result(), ~80. sor: a blokk egy vesszővel-elválasztott kifejezés-sorozat
 * (comma-expression), ami a célfájlban FOLYTATÓDIK -- záró vesszővel, önmagában NEM teljes JS)
 * leutánzó burokban.
 *
 * A BUROK: tömb-literál (`[` + blokk + `]`). Egy záróvesszős kifejezés-sorozat PONTOSAN egy
 * tömb elemlistájaként valós -- a burok NEM takar el valódi törést: a `vm.Script` konstruktor
 * CSAK FORDÍT (compile), SOHA nem futtat, tehát a blokk saját, csak-DS-ben létező globáljai
 * (Api, window, AscCommonExcel, ...) sosem dereferálódnak -- a teszt kizárólag a SZINTAXIST méri.
 * RED/GREEN (lásd lent): az ÉRINTETLEN /27-es alap ÉS a javított kimenet is burokban ZÖLD; a
 * REKONSTRUÁLT, EREDETI HIBÁS anchor-elhelyezés burokban UGYANAZZAL a "Unexpected token ':'"
 * hibával PIROS -- a burok tehát nem manufaktúrál se hamis zöldet, se hamis pirosat.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const BRIDGE_V1_DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.5.0', 'euro-office-bridge-v1');

/** vm.Script COMPILES only, never runs -- the DS-only globals the fragment references are never
 * dereferenced. Wrapped in `[`/`]` (see module header): a trailing-comma expression sequence is
 * exactly an array literal's element list. Returns {ok:true} or {ok:false, detail: the first
 * SyntaxError line, never node's leading `[location]`/source-context lines}. */
function composedBlockIsValidJs(jsText, filename) {
  try {
    new vm.Script(`[${jsText}]`, { filename });
    return { ok: true };
  } catch (e) {
    return { ok: false, detail: String((e && e.message) || e) };
  }
}

/** the indices from spanFrom in the sibling warrant test, duplicated minimally here (one small
 * helper, not worth importing a whole test file for) -- same brace-balanced, quote-aware span
 * extraction, used only by the RED-control test below to reconstruct the pre-fix shape. */
function closingBrace(src, open) {
  let depth = 0; let quote = null;
  for (let k = open; k < src.length; k += 1) {
    const c = src[k];
    if (quote) { if (c === '\\') { k += 1; continue; } if (c === quote) quote = null; continue; }
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

// The kártya's own stated scope: /25 through /28 -- the candidates this fleet has actually
// built in the euro-office-bridge-v1 family (NOT bridge-word-v1, NOT 9.3.4.60, NOT the xlnm
// print-titles REPLACE-only candidate, which carries no bridge-block.js of its own).
// EXTENDED for GH #365 option C U1 (card an internal branch): /29-veryhidden-candidate carries its
// own bridge-block.js (/28 + sheet.veryHidden.read/v1) -- this gate caught the /28 CAPS_END_ANCHOR
// bug once already; the same append-point mistake (landing a new handler before, not after, the
// PREVIOUS entry's own closing brace) was in fact made and caught HERE, locally, while building
// /29 (build-rev29-veryhidden.py's own comment on CAPS_END_PREFIX/CAPS_END_SUFFIX names it).
const CANDIDATES = Object.freeze([
  { label: 'rev25', file: path.join(BRIDGE_V1_DIR, 'rev25-candidate', 'bridge-block.js') },
  { label: 'rev26', file: path.join(BRIDGE_V1_DIR, 'rev26-candidate', 'bridge-block.js') },
  { label: 'rev27', file: path.join(BRIDGE_V1_DIR, 'rev27-candidate', 'bridge-block.js') },
  { label: 'rev28-table-autofilter-criteria', file: path.join(BRIDGE_V1_DIR, 'rev28-table-autofilter-criteria-candidate', 'bridge-block.js') },
  { label: 'rev29-veryhidden', file: path.join(BRIDGE_V1_DIR, 'rev29-veryhidden-candidate', 'bridge-block.js') },
]);

for (const { label, file } of CANDIDATES) {
  test(`composed-block syntax gate: ${label}/bridge-block.js parses cleanly, wrapped in its minimal real-insertion shell`, () => {
    assert.ok(fs.existsSync(file), `missing: ${file}`);
    const src = fs.readFileSync(file, 'utf8');
    const r = composedBlockIsValidJs(src, file);
    assert.equal(r.ok, true, r.detail);
  });
}

// POZ. KONTROLL (a kártya saját kikötése): a /27-es, régóta élő, telepített alap burokban ZÖLD --
// a burok maga nem ad hamis zöldet egy olyan fájlra, amit SOHA senki nem mért meg eddig így.
test('POSITIVE CONTROL: rev27-candidate/bridge-block.js (the long-landed base every later candidate derives from) is clean wrapped', () => {
  const src = fs.readFileSync(path.join(BRIDGE_V1_DIR, 'rev27-candidate', 'bridge-block.js'), 'utf8');
  assert.equal(composedBlockIsValidJs(src, 'rev27-candidate/bridge-block.js').ok, true);
});

// NEG. KONTROLL (a kártya saját kikötése): a RÉGI, HIBÁS anchor-elhelyezéssel rekonstruált blokk
// burokban PIROS, és UGYANAZZAL a hibával, amit a the production host-es élő kiesés adott -- a kapu tehát valóban
// a GYÖKÉR-okot méri, nem egy tetszőleges, máshonnan jövő pirosat.
test('NEGATIVE CONTROL: reconstructing the ORIGINAL BUGGY CAPS_END_ANCHOR placement (handlers landing AFTER EO_BRIDGE_CAPS\'s own closing `};`) breaks the gate with the SAME "Unexpected token \':\'" the the production host outage gave', () => {
  const rev27 = fs.readFileSync(path.join(BRIDGE_V1_DIR, 'rev27-candidate', 'bridge-block.js'), 'utf8');
  const rev28Fixed = fs.readFileSync(path.join(BRIDGE_V1_DIR, 'rev28-table-autofilter-criteria-candidate', 'bridge-block.js'), 'utf8');
  const OLD_BUGGY_ANCHOR = ' var EO_BRIDGE_WRITES = {';
  assert.equal(rev27.split(OLD_BUGGY_ANCHOR).length - 1, 1, 'the old anchor text must still be unique on the rev27 base');
  const handler1 = `${spanFrom(rev28Fixed, '"sheet.table.autoFilter.criteria.read/v1": function (args) {')},`;
  const handler2 = `${spanFrom(rev28Fixed, '"sheet.cells.compiledStyle.read/v1": function (args) {')},`;
  const reconstructedBuggy = rev27.replace(OLD_BUGGY_ANCHOR, ` ${handler1} ${handler2}${OLD_BUGGY_ANCHOR}`);
  const r = composedBlockIsValidJs(reconstructedBuggy, 'reconstructed-buggy-rev28');
  assert.equal(r.ok, false, 'the old anchor placement must still be a syntax error after reconstruction');
  assert.match(r.detail, /Unexpected token ':'/);
});
