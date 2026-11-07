'use strict';

/**
 * GH #365 option C, U1 (card an internal branch): the THREE engine REPLACE steps
 * (old1..3.txt/new1..3.txt next to plan-rev29-veryhidden.py) never run against the real
 * sdk-all.js in this agent's environment -- same situation as every earlier REPLACE-carrying
 * candidate in this family (plan-rev27-xlnm-fix.py: "this agent has no access to that file"). The
 * card's own MI A KESZ names the substitute proof: a SYNTHETIC minimal JS bundle containing the
 * three `old` texts VERBATIM (read from the same old1..3.txt files the plan script itself reads,
 * never retyped), the REPLACE applied exactly once each (asserting uniqueness, the SAME discipline
 * install_bridge.py's plan_result() enforces for real), the result a clean `node --check`
 * (`vm.Script`, in-process -- the composed-block syntax gate's own convention, chosen to avoid the
 * execFileSync('node', ['--check']) stderr-truncation failure mode this fleet already hit once),
 * and R1/R2/R3's OWN BEHAVIOUR exercised in a `vm.Context` (load -> model state -> save round
 * trip), with a NEGATIVE CONTROL: the SAME harness with R2 omitted must demonstrably fold a
 * veryHidden sheet back to plain hidden on save -- the exact silent-demotion defect docs/GH365-
 * VERYHIDDEN-SHEET-STATE-SOURCE-CONTRACT-20261105.md's "B fázis" already measured LIVE, now
 * reproduced mechanically so the fix's OWN necessity is provable offline, not just the fix's shape.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.5.0', 'euro-office-bridge-v1', 'rev29-veryhidden-candidate');

const read = name => fs.readFileSync(path.join(DIR, name), 'utf8');
const OLD1 = read('old1.txt'); const NEW1 = read('new1.txt');
const OLD2 = read('old2.txt'); const NEW2 = read('new2.txt');
const OLD3 = read('old3.txt'); const NEW3 = read('new3.txt');

/** install_bridge.py's own plan_result() rule, duplicated minimally for this offline proof: each
 * `old` must occur EXACTLY ONCE, else refused -- never a silent no-op on a missing or ambiguous
 * match. Returns the replaced string or throws. */
function replaceOnce(src, oldText, newText) {
  const count = src.split(oldText).length - 1;
  assert.equal(count, 1, `old text must occur exactly once (found ${count}): ${oldText.slice(0, 60)}...`);
  return src.replace(oldText, newText);
}

/**
 * A minimal synthetic engine bundle: EVisibleType, a binary-reader-shaped switch statement, a
 * binary-writer-shaped ternary, and Worksheet.prototype.setHidden -- each containing one of the
 * card's three `old` texts VERBATIM, wired to a tiny in-memory model/byte-stream stub so the
 * REPLACE'd behaviour can actually be exercised (load a byte -> model flags; model flags -> save a
 * byte), without depending on anything DS-only. `applyR2` lets the negative-control test build the
 * SAME harness with ONLY R1 (never R2) applied, to demonstrate R2's own necessity.
 */
function buildSyntheticBundle({ applyR1, applyR2, applyR3 }) {
  let readerSwitch = OLD1;
  if (applyR1) readerSwitch = replaceOnce(readerSwitch, OLD1, NEW1);
  let writerTernary = OLD2;
  if (applyR2) writerTernary = replaceOnce(writerTernary, OLD2, NEW2);
  let setHiddenDecl = OLD3;
  if (applyR3) setHiddenDecl = replaceOnce(setHiddenDecl, OLD3, NEW3);

  // the harness around each fragment: a fake memory stream (ReadByte/WriteByte), a fake
  // "oWorksheet" the reader mutates, and Worksheet.prototype wired so OLD3's own declaration head
  // (`Worksheet.prototype.setHidden=function(hidden){`) is a REAL function body once closed below.
  return `
    var EVisibleType = { visibleHidden: 0, visibleVeryHidden: 1, visibleVisible: 2 };
    function Memory(bytes) { this._bytes = bytes || []; this._pos = 0; this.written = []; }
    Memory.prototype.ReadByte = function () { return this._bytes[this._pos++]; };
    Memory.prototype.WriteByte = function (b) { this.written.push(b); };
    function Worksheet() { this.bHidden = false; this.bEoVeryHidden = false; }
    ${setHiddenDecl} this.bHidden = !!hidden; };
    // binary READER: reads one state byte off \`this.memory\` into a fresh oWorksheet, the SAME
    // shape as the card's R1 target (a switch over the read byte).
    function readWorksheetVisibility(memory) {
      var oWorksheet = new Worksheet();
      var stateByte = memory.ReadByte();
      switch (stateByte) {
        ${readerSwitch}
        case EVisibleType.visibleVisible: oWorksheet.bHidden = !1; break;
      }
      return oWorksheet;
    }
    // binary WRITER: the card's own old2/new2 text is a COMPLETE statement-expression (a ternary
    // whose TWO BRANCHES are themselves the \`this.memory.WriteByte(...)\` calls, never a value fed
    // into an outer WriteByte) -- so it is spliced in AS a statement, with \`this\` bound to the
    // {memory} holder and \`ws\` in scope as a parameter, the same shape as the card's R2 target.
    function writeWorksheetVisibility(holder, ws) {
      (function (ws) {
        ${writerTernary};
      }).call(holder, ws);
    }
    ({ Worksheet: Worksheet, readWorksheetVisibility: readWorksheetVisibility, writeWorksheetVisibility: writeWorksheetVisibility, Memory: Memory, EVisibleType: EVisibleType });
  `;
}

/** compiles+runs the synthetic bundle in a fresh vm.Context, returning its exported object. */
function loadBundle(src) {
  const ctx = {};
  vm.createContext(ctx);
  return vm.runInContext(src, ctx, { filename: 'synthetic-bundle.js' });
}

test('R1 + R2 + R3 applied: a round trip (load veryHidden byte -> model -> save) preserves the veryHidden byte', () => {
  const api = loadBundle(buildSyntheticBundle({ applyR1: true, applyR2: true, applyR3: true }));
  const mem = new api.Memory([api.EVisibleType.visibleVeryHidden]);
  const ws = api.readWorksheetVisibility(mem);
  assert.equal(ws.bHidden, true, 'bHidden must be true for a veryHidden-loaded sheet');
  assert.equal(ws.bEoVeryHidden, true, 'R1 must set the sibling field for visibleVeryHidden');
  const outMem = new api.Memory();
  api.writeWorksheetVisibility({ memory: outMem }, ws);
  assert.deepEqual(Array.from(outMem.written), [api.EVisibleType.visibleVeryHidden], 'R2 must write the veryHidden byte back, not visibleHidden');
});

test('R1 + R2 + R3 applied: a round trip for a PLAIN hidden sheet is byte-identical to before (no veryHidden bleed)', () => {
  const api = loadBundle(buildSyntheticBundle({ applyR1: true, applyR2: true, applyR3: true }));
  const mem = new api.Memory([api.EVisibleType.visibleHidden]);
  const ws = api.readWorksheetVisibility(mem);
  assert.equal(ws.bHidden, true);
  assert.equal(ws.bEoVeryHidden, false, 'a plain hidden sheet must NOT get the veryHidden sibling flag');
  const outMem = new api.Memory();
  api.writeWorksheetVisibility({ memory: outMem }, ws);
  assert.deepEqual(Array.from(outMem.written), [api.EVisibleType.visibleHidden]);
});

test('a visible sheet is untouched by R1/R2: loads false/false, saves visibleVisible', () => {
  const api = loadBundle(buildSyntheticBundle({ applyR1: true, applyR2: true, applyR3: true }));
  const mem = new api.Memory([api.EVisibleType.visibleVisible]);
  const ws = api.readWorksheetVisibility(mem);
  assert.equal(ws.bHidden, false);
  assert.equal(ws.bEoVeryHidden, false);
  const outMem = new api.Memory();
  api.writeWorksheetVisibility({ memory: outMem }, ws);
  assert.deepEqual(Array.from(outMem.written), [api.EVisibleType.visibleVisible]);
});

test('R3: un-hiding a veryHidden sheet (setHidden(false)) clears the sibling field -- a LATER re-hide (setHidden(true)) saves as plain hidden, not veryHidden (the card\'s own named, documented limit: undo loses veryHidden, matching the LIVE-measured B-fázis silent demotion\'s shape, now via the public API instead of a save-without-write)', () => {
  const api = loadBundle(buildSyntheticBundle({ applyR1: true, applyR2: true, applyR3: true }));
  const mem = new api.Memory([api.EVisibleType.visibleVeryHidden]);
  const ws = api.readWorksheetVisibility(mem);
  assert.equal(ws.bEoVeryHidden, true);
  ws.setHidden(false); // reveal
  assert.equal(ws.bEoVeryHidden, false, 'R3 must clear the sibling field on reveal');
  assert.equal(ws.bHidden, false);
  ws.setHidden(true); // re-hide, same session
  assert.equal(ws.bEoVeryHidden, false, 'setHidden(true) alone never sets the sibling field back -- only the binary reader (R1) does, on a FRESH load');
  const outMem = new api.Memory();
  api.writeWorksheetVisibility({ memory: outMem }, ws);
  assert.deepEqual(Array.from(outMem.written), [api.EVisibleType.visibleHidden], 'a re-hide after a reveal saves as plain hidden, never veryHidden -- documented limit, not a bug');
});

test('without R3: a veryHidden sheet revealed then re-hidden in the SAME session incorrectly saves as veryHidden again (demonstrates why R3 is needed, not merely cosmetic)', () => {
  const api = loadBundle(buildSyntheticBundle({ applyR1: true, applyR2: true, applyR3: false }));
  const mem = new api.Memory([api.EVisibleType.visibleVeryHidden]);
  const ws = api.readWorksheetVisibility(mem);
  ws.setHidden(false);
  ws.setHidden(true);
  const outMem = new api.Memory();
  api.writeWorksheetVisibility({ memory: outMem }, ws);
  assert.deepEqual(Array.from(outMem.written), [api.EVisibleType.visibleVeryHidden], 'WITHOUT R3 the stale sibling field survives a reveal, proving R3 is load-bearing, not decorative');
});

test('NEGATIVE CONTROL: without R2, a veryHidden sheet (R1 + R3 applied) still folds back to plain hidden on save -- R2 alone fixes the silent demotion the card\'s own source-contract doc measured LIVE', () => {
  const api = loadBundle(buildSyntheticBundle({ applyR1: true, applyR2: false, applyR3: true }));
  const mem = new api.Memory([api.EVisibleType.visibleVeryHidden]);
  const ws = api.readWorksheetVisibility(mem);
  assert.equal(ws.bHidden, true);
  assert.equal(ws.bEoVeryHidden, true, 'R1 alone DOES set the sibling field on load');
  const outMem = new api.Memory();
  api.writeWorksheetVisibility({ memory: outMem }, ws);
  assert.deepEqual(Array.from(outMem.written), [api.EVisibleType.visibleHidden], 'WITHOUT R2 the writer never consults the sibling field -- the veryHidden byte is silently demoted to hidden on save, reproducing the LIVE-measured B-fázis defect mechanically');
});

test('NEGATIVE CONTROL 2: without R1, the sibling field is never set at all -- loading a veryHidden byte looks identical to loading a plain hidden one', () => {
  const api = loadBundle(buildSyntheticBundle({ applyR1: false, applyR2: true, applyR3: true }));
  const veryHiddenWs = api.readWorksheetVisibility(new api.Memory([api.EVisibleType.visibleVeryHidden]));
  const plainHiddenWs = api.readWorksheetVisibility(new api.Memory([api.EVisibleType.visibleHidden]));
  assert.deepEqual(veryHiddenWs, plainHiddenWs, 'without R1 the two loaded models are indistinguishable -- the exact COLLAPSED_TO_HIDDEN finding this card exists to fix');
});

test('the three REPLACE fragments compose into syntactically valid JS (vm.Script compile-only, DS-only globals never dereferenced) -- RED/GREEN against the real old-text shape', () => {
  const fullSrc = buildSyntheticBundle({ applyR1: true, applyR2: true, applyR3: true });
  assert.doesNotThrow(() => new vm.Script(fullSrc, { filename: 'synthetic-bundle-fixed.js' }), 'fixed synthetic bundle must compile cleanly');
  const unfixedSrc = buildSyntheticBundle({ applyR1: false, applyR2: false, applyR3: false });
  assert.doesNotThrow(() => new vm.Script(unfixedSrc, { filename: 'synthetic-bundle-unpatched.js' }), 'the UNPATCHED synthetic bundle (old texts, verbatim) must ALSO compile cleanly -- a POSITIVE CONTROL proving this harness does not manufacture a syntax break of its own');
});

test('each old text occurs EXACTLY ONCE in the real rev29-veryhidden-candidate/old*.txt-sourced fragment -- the SAME uniqueness install_bridge.py enforces at apply time, proven here on the ACTUAL file contents, not retyped copies', () => {
  for (const [old, label] of [[OLD1, 'R1'], [OLD2, 'R2'], [OLD3, 'R3']]) {
    assert.ok(old.length > 10, `${label} old text must be non-trivial`);
    const rebuilt = buildSyntheticBundle({ applyR1: false, applyR2: false, applyR3: false });
    assert.equal(rebuilt.split(old).length - 1, 1, `${label}'s own old text must occur exactly once in the unpatched synthetic bundle`);
  }
});
