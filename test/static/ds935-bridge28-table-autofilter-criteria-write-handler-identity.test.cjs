'use strict';

/**
 * GH #361 route probe, route (c) + GH #363 2nd unit: the /28 CANDIDATE (documentserver-patches/9.3.5.0/
 * euro-office-bridge-v1/rev28-table-autofilter-criteria-candidate/bridge-block.js, NOT installed anywhere)
 * CARRIES EVERY /27 HANDLER BYTE-IDENTICAL, with EXACTLY TWO NEW capabilities added: "sheet.table.
 * autoFilter.criteria.read/v1" (GH #361) and, EXTENDING the same candidate (card an internal branch
 * gh363-cf-effective-fill-bridge-candidate-20261105), "sheet.cells.compiledStyle.read/v1" (GH #363). No
 * existing capability is removed or modified. Reuses the SAME capsOf()/spanFrom()/digestOf() extraction the
 * landed ds935-bridge25-write-handler-identity.test.cjs already established for this exact kind of warrant.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');
const { execFileSync } = require('node:child_process');

const DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.5.0', 'euro-office-bridge-v1');
const REV27 = path.join(DIR, 'rev27-candidate', 'bridge-block.js');
const REV28_DIR = path.join(DIR, 'rev28-table-autofilter-criteria-candidate');
const REV28 = path.join(REV28_DIR, 'bridge-block.js');
const BUILD_SCRIPT = path.join(REV28_DIR, 'build-rev28-table-autofilter-criteria.py');

const NEW_CAPABILITY = 'sheet.table.autoFilter.criteria.read/v1';
const NEW_CAPABILITY_2 = 'sheet.cells.compiledStyle.read/v1';

const sha = text => crypto.createHash('sha256').update(text).digest('hex');

/** the index of the `}` closing the `{` at `open`; strings ('..', "..", `..`) and escapes are skipped. -1 if unbalanced. */
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

let s27; let s28;
test.before(() => {
  s27 = fs.readFileSync(REV27, 'utf8');
  s28 = fs.readFileSync(REV28, 'utf8');
});

test('/28 carries EVERY /27 capability PLUS exactly two new ones -- nothing removed', () => {
  const c27 = capsOf(s27); const c28 = capsOf(s28);
  assert.ok(c27.length >= 40, `rev27 handlers: ${c27.length}`);
  assert.deepEqual(c28.filter(c => !c27.includes(c)), [NEW_CAPABILITY, NEW_CAPABILITY_2].sort(), 'exactly two new capabilities on /28');
  assert.deepEqual(c27.filter(c => !c28.includes(c)), [], 'nothing dropped on /28');
});

test('WARRANT: every /27 handler AND the helpers it calls are byte-identical on /28', () => {
  const differing = capsOf(s27).filter(cap => digestOf(s27, cap) !== digestOf(s28, cap));
  assert.deepEqual(differing, [], 'no existing /27 handler differs on /28');
  const inst = s27.slice(0, s27.indexOf('Api.GetEuroOfficeBridge = function () {'));
  assert.ok(inst.length > 100, 'the /27 block starts with the layout installer');
  assert.ok(s28.startsWith(inst), 'the /28 block starts with the SAME installer, byte for byte');
});

test('RED CONTROL: one changed byte in a /28 copy of an UNRELATED /27 handler breaks the warrant', () => {
  const cap = 'sheet.protection.write/v1';
  const h = spanFrom(s28, `"${cap}": function (args) {`);
  const mutated = s28.replace(h, h.replace('return', 'return '));
  assert.notEqual(mutated, s28, 'the mutation applied');
  assert.notEqual(digestOf(mutated, cap), digestOf(s27, cap), cap);
});

test('the new handler is brace-balanced and extractable, and names every unverified-shape gap by name', () => {
  const h = spanFrom(s28, `"${NEW_CAPABILITY}": function (args) {`);
  // GH #361 card an internal branch, a maintainer's own
  // measured shapes (#ref LIVE dump, #ref targeted dump): per-column criteria ARE now read -- the
  // blanket "fields are unverified" gap is retired, replaced by named per-column gaps for exactly the
  // sub-fields that remain unmeasured (only one value-filter column and one custom-filter column with one
  // condition have ever been observed).
  assert.match(h, /"internal-shape-unverified:ref"/);
  assert.doesNotMatch(h, /"internal-shape-unverified:fields"/, 'the old blanket fields-gap is retired -- per-column gaps replace it');
  for (const gap of ['blank', 'dates', 'colorFilter', 'dynamicFilter', 'top10', 'and-true', 'column-kind']) {
    assert.match(h, new RegExp(`"internal-shape-unverified:${gap}"`), `missing named gap: ${gap}`);
  }
  assert.match(h, /present: false, range: null, fields: null/, 'no AutoFilter on the table is a clean, non-gap result');
  // every per-column engine-object property read uses the SAME try/catch-per-access idiom this handler's
  // pre-existing reads already use for AutoFilter/Ref/TableParts/DisplayName -- never a bare read outside one
  for (const prop of ['col.ColId', 'filtersObj.Values', 'filtersObj.Blank', 'filtersObj.Dates', 'customObj.And', 'customObj.CustomFilters']) {
    assert.match(h, new RegExp(`try \\{[^;]*${prop.replace('.', '\\.')}`), `${prop} must be read inside a try block`);
  }
  // the operator-code -> name lookup never assumes the enum exists (a maintainer's own instruction: NEVER guess the
  // mapping -- read it from the runtime if available, else raw code + operatorName null)
  assert.match(h, /typeof Asc !== "undefined"/, 'the Asc enum lookup checks existence before use');
  assert.match(h, /\? opNames\[opCode\] : null/, 'the no-enum/no-match fallback is an explicit null, never an invented name');
});

test('the new manifest entry exists exactly once, is honest about what it does NOT measure, and the embedded patchVersion marker advanced', () => {
  const entryCount = (s28.match(new RegExp(`\\{ id: "${NEW_CAPABILITY.replace(/[/.]/g, '\\$&')}"`, 'g')) || []).length;
  assert.equal(entryCount, 1);
  const entry = spanFrom(s28, `{ id: "${NEW_CAPABILITY}"`);
  assert.match(entry, /operatorName ONLY when a runtime enum resolves it else null/);
  assert.match(entry, /named unverified array/);
  // a maintainer's review (internal tracker, komment #ref): the marker originally carried this candidate's own
  // directory suffix, unlike every earlier candidate's clean "/NN" -- fixed at the build script's source.
  assert.equal((s28.match(/patchVersion: "euro-office-bridge-v1\/9\.3\.5\.0\/28"/g) || []).length, 1);
  assert.equal((s28.match(/patchVersion: "euro-office-bridge-v1\/9\.3\.5\.0\/27"/g) || []).length, 0, 'the old /27 marker must not remain');
});

// GH #363 2nd unit (card an internal branch): the SECOND new
// capability, "sheet.cells.compiledStyle.read/v1", EXTENDING this same /28 candidate -- not a new rev29.
test('the compiledStyle handler is brace-balanced and extractable, gates its own preconditions, and never guesses a fill', () => {
  const h = spanFrom(s28, `"${NEW_CAPABILITY_2}": function (args) {`);
  // bad-args / 200-cell cap / sheet-not-found / worksheet-missing / getCompiledStyle-not-a-function -- every
  // precondition is its own named, gated reason, never a silent pass-through.
  assert.match(h, /reason: "bad-args"/);
  assert.match(h, /args\.cells\.length > 200/, 'defense in depth: the handler caps at 200 cells too, not only read-tools.cjs\'s pre-dispatch check');
  assert.match(h, /reason: "too-many-cells"/);
  assert.match(h, /reason: "sheet-not-found"/);
  assert.match(h, /reason: "worksheet-missing"/);
  assert.match(h, /reason: "getCompiledStyle-not-a-function"/);
  assert.match(h, /reason: "bad-cell-address"/);
  assert.match(h, /reason: "threw"/);
  // ONLY the fill is ever read off the compiled style -- never a second field (font/border/numFmt/...).
  assert.match(h, /compiled: \{ fill: eoCompiledStyleFillOfStyle\(style\) \}/);
  assert.doesNotMatch(h, /\.getFont\(|\.getBorder\(|\.getNumFormat\(/, 'no field other than fill is ever read off the compiled style');
  // the measured fallback order, lifted from the 2nd unit's own LIVE-proven probe (colorFromFillObject):
  // getColor() first, then getSolidFill(), never the reverse and never a third guess.
  const getColorIdx = h.indexOf('fillObj.getColor');
  const getSolidFillIdx = h.indexOf('fillObj.getSolidFill');
  assert.ok(getColorIdx > 0 && getSolidFillIdx > getColorIdx, 'getColor() must be tried before getSolidFill()');
  // every internal-model property/method access is inside its own try/catch -- the SAME discipline the
  // first handler (and the 2nd unit's own probe) already uses, never a bare read that could throw unguarded.
  for (const expr of ['ws.worksheet', 'style.getFill()', 'fillObj.getColor()', 'fillObj.getSolidFill()', 'c.getRgb()', 'm.getCompiledStyle(rc.row, rc.col)']) {
    assert.match(h, new RegExp(`try \\{[^;]*${expr.replace(/[().]/g, '\\$&')}`), `${expr} must be read inside a try block`);
  }
});

test('eoCompiledStyleA1ToRowCol (inlined in the compiledStyle handler, prefixed to avoid a name collision -- see HANDLER_ENTRY_2\'s own comment) matches the 2nd unit\'s own LIVE-proven formula -- B7 -> {row:6,col:1}, B2 -> {row:1,col:1}', () => {
  const h = spanFrom(s28, `"${NEW_CAPABILITY_2}": function (args) {`);
  const fnSrc = spanFrom(h, 'function eoCompiledStyleA1ToRowCol(a1) {');
  // eslint-disable-next-line no-new-func
  const a1ToRowCol = new Function(`${fnSrc}\nreturn eoCompiledStyleA1ToRowCol;`)();
  assert.deepEqual(a1ToRowCol('B7'), { row: 6, col: 1 });
  assert.deepEqual(a1ToRowCol('B2'), { row: 1, col: 1 });
  assert.deepEqual(a1ToRowCol('A1'), { row: 0, col: 0 });
  assert.equal(a1ToRowCol('not-a-cell'), null);
});

test('the compiledStyle handler\'s own helper names (eoCompiledStyle-prefixed) never collide with an existing file-wide function name -- the exact bug a first attempt (bare colourOf/a1ToRowCol/fillOfStyle) hit, corrupting sheet.autoFilter.colorFilter.read/v1\'s unrelated digest', () => {
  for (const name of ['eoCompiledStyleA1ToRowCol', 'eoCompiledStyleColourOf', 'eoCompiledStyleFillOfStyle']) {
    assert.equal((s27.match(new RegExp(`function ${name}\\(`, 'g')) || []).length, 0, `${name} must not already exist on /27`);
  }
  const unrelatedCap = 'sheet.autoFilter.colorFilter.read/v1';
  assert.equal(digestOf(s27, unrelatedCap), digestOf(s28, unrelatedCap), 'an unrelated handler\'s digest must be unaffected by the new helper names');
});

test('the compiledStyle manifest entry exists exactly once and is honest about scope (fill only, nothing else attempted)', () => {
  const entryCount = (s28.match(new RegExp(`\\{ id: "${NEW_CAPABILITY_2.replace(/[/.]/g, '\\$&')}"`, 'g')) || []).length;
  assert.equal(entryCount, 1);
  const entry = spanFrom(s28, `{ id: "${NEW_CAPABILITY_2}"`);
  assert.match(entry, /ONLY the fill is read/);
  assert.match(entry, /every other compiled-style field.*NOT attempted/);
  assert.match(entry, /compiled:null, never a guessed fill/);
});

test('the build script is deterministic: running it again (via --check) reproduces the SAME output_sha256 already on disk', () => {
  const out = execFileSync('python3', [BUILD_SCRIPT, '--check'], { encoding: 'utf8' });
  const { output_sha256: fresh } = JSON.parse(out);
  assert.equal(fresh, sha(s28));
});

test('manifest-mechanism-proof.json names the real base_sha256 (the actual rev27-candidate/bridge-block.js on disk) and the real output_sha256', () => {
  const proof = JSON.parse(fs.readFileSync(path.join(REV28_DIR, 'manifest-mechanism-proof.json'), 'utf8'));
  assert.equal(proof.base_sha256, sha(s27));
  assert.equal(proof.output_sha256, sha(s28));
  assert.deepEqual(proof.new_capabilities, [NEW_CAPABILITY, NEW_CAPABILITY_2]);
});

// GH #361/#363 SURGOS fix (a maintainer's lelete, 2026-11-06): the original CAPS_END_ANCHOR landed the two new
// handler entries AFTER EO_BRIDGE_CAPS's own closing `};`, not before it -- a string-literal key ("x": ...)
// as a bare top-level statement, which is a SyntaxError the warrant tests above never catch (they extract
// and check named spans, never parse the whole composed file). The composed-block syntax gate for THIS
// (array-wrapped `vm.Script`, RED/GREEN/POSITIVE CONTROL, across /25-/28) lives in its own dedicated file:
// ds935-bridge-block-composed-syntax-gate.test.cjs -- not duplicated here. (A first attempt duplicated it
// here via `execFileSync('node', ['--check'], {input: ...})`: for this ~190KB, single-line fragment,
// execFileSync's captured `e.stderr` measured SHORTER than the SAME command's output captured by a plain
// shell redirect -- 146176 vs 191281 chars -- and silently LOST the trailing "SyntaxError: ..." line the
// shell capture kept. The exit code/thrown-ness was still correct; only the TEXT assertion on `e.stderr`
// was unreliable. `vm.Script` (in-process, no subprocess pipe) does not have this failure mode.)
