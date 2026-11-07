'use strict';

/**
 * f0451f47 -- the euro-office-bridge-v1 /26 CANDIDATE for DS 9.3.5 (documentserver-patches/9.3.5.0/
 * euro-office-bridge-v1/rev26-candidate/): worksheet.copy/v1 and sheet.table.write/v1's own workbook-protection
 * guard is moved from `e.asc_isProtectedWorkbook() !== false` to the SAME getter workbook.protection.read/v1
 * already reads (wbModel.getWorkbookProtection(c_oAscWorkbookProtectType.lockStructure)) -- not a new API.
 *
 * THE LELET this fixes (a maintainer, the production host /25 window, DS card #ref): on a workbook carrying a bare, attribute-less
 * <workbookProtection/> element, copy_sheet refused while the other four sheet-structure intents applied (the
 * /25 read-handler fix does not reach these two WRITE handlers' own, separate guard). Offline: the REAL candidate
 * blocks are evaluated against fake engines in the measured shape (asc_isProtectedWorkbook and
 * wbModel.getWorkbookProtection returning DIFFERENT values on the same fixture -- the two-API disagreement the
 * card names as the root cause). This proves the HANDLER LOGIC -- NOT that the fix clears the incident on a real
 * DS (that LIVE run, test/live/xlsx-f0451f47-worksheet-copy-protection-probe.live.cjs, is a maintainer's own lane).
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { spawnSync } = require('node:child_process');

const DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.5.0', 'euro-office-bridge-v1');

// tracker an internal branch -- SAME fix as ds935-bridge25's own
// isolatedBuildRun(), same reasoning (see that file's comment): this test's build step used to write
// rev26-candidate/bridge-block.js IN PLACE in the shared repo tree while reading rev25-candidate/bridge-block.js
// as a pinned-sha dependency -- and ds935-bridge25's OWN build step rewrites that exact rev25 file concurrently
// under full-suite execution. Isolate both directories into a per-call temp copy so nothing is shared.
function isolatedBuildRun() {
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'ds935-bridge26-'));
  const rev25Dir = path.join(tmp, 'rev25-candidate');
  const rev26Dir = path.join(tmp, 'rev26-candidate');
  fs.mkdirSync(rev25Dir);
  fs.mkdirSync(rev26Dir);
  fs.copyFileSync(path.join(DIR, 'rev25-candidate', 'bridge-block.js'), path.join(rev25Dir, 'bridge-block.js'));
  fs.copyFileSync(path.join(DIR, 'rev26-candidate', 'build-rev26.py'), path.join(rev26Dir, 'build-rev26.py'));
  const script = path.join(rev26Dir, 'build-rev26.py');
  const run = () => spawnSync('python3', ['-I', script], { cwd: rev26Dir, encoding: 'utf8' });
  return { tmp, rev26Dir, run };
}
const V25 = fs.readFileSync(path.join(DIR, 'rev25-candidate', 'bridge-block.js'), 'utf8');
const V26 = fs.readFileSync(path.join(DIR, 'rev26-candidate', 'bridge-block.js'), 'utf8');
const CP = 'worksheet.copy/v1';
const TW = 'sheet.table.write/v1';
const sha = t => crypto.createHash('sha256').update(t).digest('hex');
const reason = r => (r && r.ok === false ? r.reason : `OK:${JSON.stringify(r && r.value)}`);

const FORBIDDEN = ['Object.keys', 'Object.values', 'Object.entries', 'getOwnPropertyNames', 'for (var k in', 'for (var p in', ' in args',
  '//', 'hashValue', 'saltValue', 'spinCount', 'algorithmName', 'password', 'temporaryPassword', 'PasswordXL', 'asc_setSheet', 'asc_setPassword'];

// ---- the seven anchored substitutions, duplicated here (not imported) so this test does not trust the builder's
//      own bookkeeping -- the same discipline as the /25 test file re-deriving its own assertions. ---------------

const ANCHORS = [
  { old: 'patchVersion: "euro-office-bridge-v1/9.3.5.0/25"', new: 'patchVersion: "euro-office-bridge-v1/9.3.5.0/26"' },
  {
    old: '["asc_copyWorksheet", "asc_getWorksheetsCount", "asc_getWorksheetName", "asc_isProtectedWorkbook", "asc_registerCallback", "asc_unregisterCallback", "canEdit"]',
    new: '["asc_copyWorksheet", "asc_getWorksheetsCount", "asc_getWorksheetName", "asc_registerCallback", "asc_unregisterCallback", "canEdit"]',
  },
  {
    old: 'if (e.asc_isProtectedWorkbook() !== false) return { ok: false, reason: "workbook-protected" }; var before = names();',
    new: null, // asserted structurally below (contains workbook-protection-unreadable, var before = names() tail), not string-for-string
  },
  {
    old: 'reasons: bad-args, unsupported-property, invalid-name, editor-missing, edit-not-allowed, workbook-locked, workbook-protected, name-taken, source-not-found, source-ambiguous, <api>-missing, copy-threw, copy-error, copy-not-observed, copy-unexpected (B03-WRITE, /8)" },',
    new: null,
  },
  {
    old: '[CALLS[prop], "asc_getWorksheetsCount", "asc_getWorksheetName", "asc_showWorksheet", "asc_isProtectedWorkbook", "asc_registerCallback", "asc_unregisterCallback", "canEdit"]',
    new: '[CALLS[prop], "asc_getWorksheetsCount", "asc_getWorksheetName", "asc_showWorksheet", "asc_registerCallback", "asc_unregisterCallback", "canEdit"]',
  },
  {
    old: 'if (e.asc_isProtectedWorkbook() !== false) return { ok: false, reason: "workbook-protected" }; var sheetNames = function () {',
    new: null,
  },
  { old: null, new: null }, // table reasons list, asserted structurally below
];

test('V25 does not contain asc_isProtectedWorkbook ZERO times (sanity: the OLD guard really is there to remove)', () => {
  assert.equal((V25.match(/asc_isProtectedWorkbook/g) || []).length, 4, 'two call sites x two mentions each (API array + guard call)');
});

test('V26: asc_isProtectedWorkbook is referenced NOWHERE', () => {
  assert.equal((V26.match(/asc_isProtectedWorkbook/g) || []).length, 0);
});

for (const a of ANCHORS) {
  if (a.old && a.new) {
    test(`anchor present exactly once in each block: ${a.old.slice(0, 50)}...`, () => {
      assert.equal(V25.split(a.old).length - 1, 1, 'old anchor must be unique in V25');
      assert.equal(V26.split(a.old).length - 1, 0, 'old anchor must be GONE from V26');
      assert.equal(V26.split(a.new).length - 1, 1, 'new anchor must be unique in V26');
    });
  }
}

test('the new guard reasons list gains exactly "workbook-protected" + "workbook-protection-unreadable" (nothing else renamed) for BOTH handlers', () => {
  for (const [name, old26] of [
    ['worksheet.copy/v1', 'reasons: bad-args, unsupported-property, invalid-name, editor-missing, edit-not-allowed, workbook-locked, workbook-protected, workbook-protection-unreadable, name-taken, source-not-found, source-ambiguous, <api>-missing, copy-threw, copy-error, copy-not-observed, copy-unexpected (B03-WRITE, /8, f0451f47/26)" },'],
    ['sheet.table.write/v1', 'workbook-locked, workbook-protected, workbook-protection-unreadable, sheet-not-found, sheet-ambiguous, tables-read-failed:<reason>'],
  ]) {
    assert.ok(V26.includes(old26), `${name} reasons list must carry the new text verbatim`);
  }
});

test('BYTE-IDENTITY: V25 and V26 are IDENTICAL once the seven anchors are masked out (nothing else changed)', () => {
  const PATCHVER_OLD = 'patchVersion: "euro-office-bridge-v1/9.3.5.0/25"';
  const PATCHVER_NEW = 'patchVersion: "euro-office-bridge-v1/9.3.5.0/26"';
  const API_COPY_OLD = '["asc_copyWorksheet", "asc_getWorksheetsCount", "asc_getWorksheetName", "asc_isProtectedWorkbook", "asc_registerCallback", "asc_unregisterCallback", "canEdit"]';
  const API_COPY_NEW = '["asc_copyWorksheet", "asc_getWorksheetsCount", "asc_getWorksheetName", "asc_registerCallback", "asc_unregisterCallback", "canEdit"]';
  const GUARD_COPY_OLD = 'if (e.asc_isProtectedWorkbook() !== false) return { ok: false, reason: "workbook-protected" }; var before = names();';
  const REASONS_COPY_OLD = 'reasons: bad-args, unsupported-property, invalid-name, editor-missing, edit-not-allowed, workbook-locked, workbook-protected, name-taken, source-not-found, source-ambiguous, <api>-missing, copy-threw, copy-error, copy-not-observed, copy-unexpected (B03-WRITE, /8)" },';
  const API_TABLE_OLD = '[CALLS[prop], "asc_getWorksheetsCount", "asc_getWorksheetName", "asc_showWorksheet", "asc_isProtectedWorkbook", "asc_registerCallback", "asc_unregisterCallback", "canEdit"]';
  const API_TABLE_NEW = '[CALLS[prop], "asc_getWorksheetsCount", "asc_getWorksheetName", "asc_showWorksheet", "asc_registerCallback", "asc_unregisterCallback", "canEdit"]';
  const GUARD_TABLE_OLD = 'if (e.asc_isProtectedWorkbook() !== false) return { ok: false, reason: "workbook-protected" }; var sheetNames = function () {';
  const REASONS_TABLE_OLD = 'reasons: bad-args, unsupported-property, style-unknown, flag-unknown, value-not-boolean, name-invalid, ref-unparsable, editor-missing, <api>-missing, c_oAscChangeFilterOptions-missing, c_oAscChangeTableStyleInfo-missing, edit-not-allowed, workbook-locked, workbook-protected, sheet-not-found, sheet-ambiguous, tables-read-failed:<reason>, table-not-found, table-ambiguous, already-at-value, table-style-info-missing, header-row-missing, name-taken, range-unreadable, ref-moves-origin, ref-too-small, ref-overlaps, <call>-threw, <call>-error, table-count-changed, other-table-changed, no-effect, <property>-unexpected, target-collateral-change (GH #332-#335, /19)" },';

  const PLACEHOLDER = '\u0000MASK\u0000';
  let masked25 = V25
    .replace(PATCHVER_OLD, PLACEHOLDER)
    .replace(API_COPY_OLD, PLACEHOLDER)
    .replace(GUARD_COPY_OLD, PLACEHOLDER)
    .replace(REASONS_COPY_OLD, PLACEHOLDER)
    .replace(API_TABLE_OLD, PLACEHOLDER)
    .replace(GUARD_TABLE_OLD, PLACEHOLDER)
    .replace(REASONS_TABLE_OLD, PLACEHOLDER);
  // V26's replacement text for each anchor varies in length (the guard grows); mask by RANGE instead of by exact
  // string so this test does not re-encode the new guard body (that would just be the builder talking to itself).
  let masked26 = V26;
  const V26_MARKERS = [
    [PATCHVER_NEW, PATCHVER_NEW],
    [API_COPY_NEW, API_COPY_NEW],
    ['var wpWb; try { wpWb = e.wbModel;', 'var before = names();'],
    ['reasons: bad-args, unsupported-property, invalid-name,', '(B03-WRITE, /8, f0451f47/26)" },'],
    [API_TABLE_NEW, API_TABLE_NEW],
    ['var wpWb; try { wpWb = e.wbModel;', 'var sheetNames = function () {'],
    ['reasons: bad-args, unsupported-property, style-unknown,', '(GH #332-#335, /19, f0451f47/26)" },'],
  ];
  for (const [startTok, endTok] of V26_MARKERS) {
    const start = masked26.indexOf(startTok);
    assert.ok(start >= 0, `start token not found: ${startTok}`);
    const end = masked26.indexOf(endTok, start) + endTok.length;
    assert.ok(end > start, `end token not found after start: ${endTok}`);
    masked26 = masked26.slice(0, start) + PLACEHOLDER + masked26.slice(end);
  }
  assert.equal(masked26, masked25, 'every byte outside the seven anchors must be unchanged between /25 and /26');
});

// ---- the build --------------------------------------------------------------------------------------------

test('the /26 block = the /25 block + exactly the seven anchored substitutions, and the builder is deterministic', () => {
  const { tmp, rev26Dir, run } = isolatedBuildRun();
  try {
    const r1 = run();
    assert.equal(r1.status, 0, r1.stderr);
    const r2 = run();
    assert.equal(r2.stdout, r1.stdout, 'the build is byte-for-byte reproducible across two runs');
    assert.equal(fs.readFileSync(path.join(rev26Dir, 'bridge-block.js'), 'utf8'), V26, 'the committed block equals what the builder just produced');
    assert.equal(fs.readFileSync(path.join(rev26Dir, 'old.txt'), 'utf8'), V25);
    assert.equal(fs.readFileSync(path.join(rev26Dir, 'new.txt'), 'utf8'), V26);
  } finally {
    fs.rmSync(tmp, { recursive: true, force: true });
  }
});

test('both blocks parse the same way (the fragment trailing comma stripped)', () => {
  for (const t of [V25, V26]) assert.doesNotThrow(() => new Function('Api', 'Asc', 'window', 'AscCommon', t.replace(/\}\s*,\s*$/, '}')));
});

test('no forbidden word (no object enumeration, no line comment, no credential field name) anywhere in the two new guard stretches', () => {
  const idxCopy = V26.indexOf('var wpWb; try { wpWb = e.wbModel;');
  const idxTable = V26.lastIndexOf('var wpWb; try { wpWb = e.wbModel;');
  for (const idx of [idxCopy, idxTable]) {
    const stretch = V26.slice(idx, idx + 700);
    for (const w of FORBIDDEN) assert.ok(!stretch.includes(w), `${w} found near ${idx}`);
  }
});

test('manifest: /26, the SAME capabilities as /25 (nothing added, nothing removed), patchVersion bumped', () => {
  const man = load(engineCopy()).manifest();
  assert.equal(man.patchVersion, 'euro-office-bridge-v1/9.3.5.0/26');
  const ids25 = load(engineCopy(), V25).manifest().capabilities.map(c => c.id).sort();
  const ids26 = man.capabilities.map(c => c.id).sort();
  assert.deepEqual(ids26, ids25);
  assert.ok(ids26.includes(CP)); assert.ok(ids26.includes(TW));
});

// ---- the fake worksheet.copy/v1 engine (both protection APIs independently controllable) -------------------

function engineCopy({ list = ['Src', 'Other'], lockStructure = false, wbMissing = false, typeMissing = false, protThrows = false,
  hasIsProtectedWorkbook = true, isProtectedWorkbookReturns = false, copyMode = 'append' } = {}) {
  const names = list.slice();
  const wb = wbMissing ? undefined : {
    getWorkbookProtection(t) {
      if (protThrows) throw new Error('boom');
      if (t !== 0) throw new Error('bad type');
      return lockStructure;
    },
  };
  const e = {
    wbModel: wb,
    canEdit: () => true,
    collaborativeEditing: { getGlobalLock: () => 0 },
    asc_getWorksheetsCount: () => names.length,
    asc_getWorksheetName: i => names[i],
    asc_registerCallback: () => {},
    asc_unregisterCallback: () => {},
    asc_copyWorksheet: (w, nm) => { if (copyMode === 'append') names.push(nm[0]); },
  };
  if (hasIsProtectedWorkbook) e.asc_isProtectedWorkbook = () => isProtectedWorkbookReturns;
  const Asc = { editor: e, c_oAscWorkbookProtectType: typeMissing ? {} : { lockStructure: 0 } };
  return { Asc, window: { Asc }, AscCommon: {} };
}

function load(e, block = V26) {
  const Api = e.Api || { GetSheet: () => null };
  new Function('Api', 'Asc', 'window', 'AscCommon', block.replace(/\}\s*,\s*$/, '}'))(Api, e.Asc, e.window, e.AscCommon);
  return Api.GetEuroOfficeBridge();
}
const copyWrite = (e, block) => load(e, block).write(CP, { sheet: 'Src', name: 'SrcCopy' });

// ---- worksheet.copy/v1 -- THE FIX itself ---------------------------------------------------------------------

test('THE FIX, reproducing the LELET exactly: wbModel says null (not locked), asc_isProtectedWorkbook says true (the measured regression on the OLD guard) -- /26 lets it through, /25 still refuses', () => {
  const e = engineCopy({ lockStructure: null, isProtectedWorkbookReturns: true });
  assert.deepEqual(copyWrite(e, V26), { ok: true, value: { applied: ['copy'], name: 'SrcCopy', index: 2 } });
  const e25 = engineCopy({ lockStructure: null, isProtectedWorkbookReturns: true });
  assert.equal(reason(copyWrite(e25, V25)), 'workbook-protected', 'RED CONTROL: the /25 (pre-fix) guard still refuses on the identical fixture -- proving the test discriminates, not just passes by construction');
});

test('/26 never calls asc_isProtectedWorkbook: an editor that lacks the method entirely still succeeds when wbModel says not-locked', () => {
  const e = engineCopy({ lockStructure: false, hasIsProtectedWorkbook: false });
  assert.deepEqual(copyWrite(e, V26), { ok: true, value: { applied: ['copy'], name: 'SrcCopy', index: 2 } });
});

test('RED CONTROL: the SAME lacks-the-method editor makes /25 fail closed with a NAMED missing-api reason (proving /25 really still depends on it)', () => {
  const e = engineCopy({ lockStructure: false, hasIsProtectedWorkbook: false });
  assert.equal(reason(copyWrite(e, V25)), 'asc_isProtectedWorkbook-missing');
});

test('lockStructure === true: refused workbook-protected on /26 (the one state that was ever meant to be refused)', () => {
  const e = engineCopy({ lockStructure: true });
  assert.equal(reason(copyWrite(e, V26)), 'workbook-protected');
});

test('lockStructure === false (explicit boolean, always worked): unchanged -- copy succeeds', () => {
  const e = engineCopy({ lockStructure: false });
  assert.deepEqual(copyWrite(e, V26), { ok: true, value: { applied: ['copy'], name: 'SrcCopy', index: 2 } });
});

test('every other non-boolean/unreadable shape on /26: fail-closed, workbook-protection-unreadable (never treated as not-locked)', () => {
  for (const params of [{ protThrows: true }, { wbMissing: true }, { typeMissing: true }, { lockStructure: 'nope' }, { lockStructure: 0 }, { lockStructure: 1 }]) {
    const e = engineCopy(params);
    assert.equal(reason(copyWrite(e, V26)), 'workbook-protection-unreadable', JSON.stringify(params));
  }
});

test('unrelated refusal reasons on /26 are unchanged from /25 (bad-args, invalid-name, name-taken, source-not-found)', () => {
  const e = () => engineCopy({ lockStructure: false });
  assert.equal(reason(load(e()).write(CP, 'nope')), 'bad-args');
  assert.equal(reason(load(e()).write(CP, { sheet: 'Src', name: "'bad" })), 'invalid-name');
  assert.equal(reason(load(e()).write(CP, { sheet: 'Src', name: 'Other' })), 'name-taken');
  assert.equal(reason(load(e()).write(CP, { sheet: 'Nope', name: 'X' })), 'source-not-found');
});

// ---- the fake sheet.table.write/v1 engine (same shapes as ds935-bridge19-table-write.test.cjs / bridge24's own
//      tableWorld, extended with the same two independently-controllable protection knobs as engineCopy above) ---

function mkTable(name, ref = 'A1:C5') {
  return { name, ref, columns: ['Item', 'Qty', 'Note'], style: { name: 'TableStyleMedium2', rowStripes: true, columnStripes: false, firstColumn: false, lastColumn: false } };
}
function engineTable({ tables = { S1: [mkTable('Table1')] }, lockStructure = false, wbMissing = false, typeMissing = false, protThrows = false,
  hasIsProtectedWorkbook = true, isProtectedWorkbookReturns = false } = {}) {
  const model = tables;
  const names = Object.keys(model);
  const find = n => { for (const s of names) for (const t of model[s]) if (t.name === n) return t; return null; };
  const part = t => ({
    DisplayName: t.name, Ref: { getAbsName: () => '$' + t.ref.replace(/([A-Z]+)(\d+)/g, '$$$1$$$2') },
    isHeaderRow: () => true, isTotalsRow: () => false, TableColumns: t.columns.map(c => ({ Name: c })),
    TableStyleInfo: { Name: t.style.name, ShowRowStripes: t.style.rowStripes, ShowColumnStripes: t.style.columnStripes, ShowFirstColumn: t.style.firstColumn, ShowLastColumn: t.style.lastColumn },
    AutoFilter: {},
  });
  const Api = { GetSheet: n => (Object.prototype.hasOwnProperty.call(model, n) ? { worksheet: { get TableParts() { return model[n].map(part); } } } : null) };
  const wb = wbMissing ? undefined : {
    getWorkbookProtection(t) {
      if (protThrows) throw new Error('boom');
      if (t !== 0) throw new Error('bad type');
      return lockStructure;
    },
  };
  const e = {
    wbModel: wb,
    canEdit: () => true, collaborativeEditing: { getGlobalLock: () => false },
    asc_getWorksheetsCount: () => names.length, asc_getWorksheetName: i => names[i], asc_showWorksheet: () => {},
    asc_getActiveWorksheetIndex: () => 0, asc_registerCallback: () => {}, asc_unregisterCallback: () => {},
    asc_changeAutoFilter(tn, opt, val) { const t = find(tn); if (opt === 2) t.style.name = val; },
  };
  if (hasIsProtectedWorkbook) e.asc_isProtectedWorkbook = () => isProtectedWorkbookReturns;
  const consts = { c_oAscChangeFilterOptions: { style: 2 }, c_oAscChangeTableStyleInfo: { rowStripes: 6, columnStripes: 3, firstColumn: 1, lastColumn: 2, rowTotal: 5 } };
  const Asc = Object.assign({ editor: e, c_oAscWorkbookProtectType: typeMissing ? {} : { lockStructure: 0 } }, consts);
  return { Api, Asc, window: { Asc }, AscCommon: {} };
}
const tableWrite = (e, block, args) => load(e, block).write(TW, Object.assign({ sheet: 'S1', table: 'Table1', property: 'style', name: 'TableStyleMedium5' }, args));

test('THE FIX (table): same repro as copy -- null/true divergence, /26 applies, /25 refuses', () => {
  const e = engineTable({ lockStructure: null, isProtectedWorkbookReturns: true });
  assert.equal(reason(tableWrite(e, V26)), 'OK:{"applied":["style"],"table":"Table1","observed":"TableStyleMedium5"}');
  const e25 = engineTable({ lockStructure: null, isProtectedWorkbookReturns: true });
  assert.equal(reason(tableWrite(e25, V25)), 'workbook-protected', 'RED CONTROL');
});

test('/26 table write never calls asc_isProtectedWorkbook', () => {
  const e = engineTable({ lockStructure: false, hasIsProtectedWorkbook: false });
  assert.equal(reason(tableWrite(e, V26)), 'OK:{"applied":["style"],"table":"Table1","observed":"TableStyleMedium5"}');
});

test('RED CONTROL (table): the lacks-the-method editor fails /25 closed, named', () => {
  const e = engineTable({ lockStructure: false, hasIsProtectedWorkbook: false });
  assert.equal(reason(tableWrite(e, V25)), 'asc_isProtectedWorkbook-missing');
});

test('lockStructure === true (table): refused workbook-protected', () => {
  const e = engineTable({ lockStructure: true });
  assert.equal(reason(tableWrite(e, V26)), 'workbook-protected');
});

test('every other non-boolean/unreadable shape (table): workbook-protection-unreadable', () => {
  for (const params of [{ protThrows: true }, { wbMissing: true }, { typeMissing: true }, { lockStructure: 3 }]) {
    const e = engineTable(params);
    assert.equal(reason(tableWrite(e, V26)), 'workbook-protection-unreadable', JSON.stringify(params));
  }
});

test('unrelated table refusal reasons unchanged on /26 (sheet-not-found, table-not-found)', () => {
  const lock = { lockStructure: false };
  assert.equal(reason(tableWrite(engineTable(lock), V26, { sheet: 'Nope' })), 'sheet-not-found');
  assert.equal(reason(tableWrite(engineTable(lock), V26, { table: 'Nope' })), 'table-not-found');
});

// ---- plan-rev26.py: proves the DERIVATION mechanism (same discipline as the /17 plan test) on a synthetic,
//      internally-consistent fake "unpatched" input -- the real prod sdk-all.js is held only by the operator
//      (maintainer), this agent cannot read it, so pinning plan-rev26.py's output against the REAL prod unpatched
//      sha (a maintainer's #ref: input 5e1e7d6df5a0 -> cell-sdk-all-rev26.js ef7428a1410a8a835427cbf23ba94ddbaf5bccefb076dc67d3b27b7200c7bf64)
//      is NOT_MEASURED here: that input is not in the repo/fixtures, and is not reproducible from a chosen sha256
//      (a preimage). maintainer's own the production host-window build (plan-rev25.py copy + two constants bumped to /26, independently
//      byte-matched against a separate substitution) already measured that value; this test measures that
//      plan-rev26.py itself -- unmodified from that pattern -- derives a correct, install_bridge.py-accepted plan,
//      and refuses (writing nothing) on any input that does not match the base plan's recorded input_sha256.
test('plan-rev26.py derives a plan the UNCHANGED install_bridge.py accepts; an unknown input is refused, nothing written', () => {
  const os = require('node:os');
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'eo-b26-plan-'));
  try {
    const root = path.join(__dirname, '..', '..', 'documentserver-patches');
    fs.cpSync(path.join(root, '9.3.5.0'), path.join(tmp, '9.3.5.0'), { recursive: true });
    fs.copyFileSync(path.join(root, '9.3.4.60', 'install_bridge.py'), path.join(tmp, 'install_bridge.py'));
    const data = Buffer.from('head;AscBuilder.Cell.init=function(x){};mid;Api.GetSheets=function(){};tail');
    const planPath = path.join(tmp, '9.3.5.0', 'fresh-install.json');
    const plan = JSON.parse(fs.readFileSync(planPath, 'utf8'));
    plan.cell.input_sha256 = sha(data);
    fs.writeFileSync(planPath, JSON.stringify(plan));
    fs.writeFileSync(path.join(tmp, 'in.js'), data);
    const helper = path.join(tmp, '9.3.5.0', 'euro-office-bridge-v1', 'rev26-candidate', 'plan-rev26.py');
    const out = path.join(tmp, 'p26.json');
    const d = spawnSync('python3', ['-I', helper, '--input', path.join(tmp, 'in.js'), '--out', out], { encoding: 'utf8' });
    assert.equal(d.status, 0, d.stderr);
    const p26 = JSON.parse(fs.readFileSync(out, 'utf8'));
    assert.equal(p26.cell.patchVersion, 'euro-office-bridge-v1/9.3.5.0/26');
    assert.equal(p26.cell.blocks.find(b => b.file.startsWith('euro-office-bridge-v1/')).sha256, sha(V26));
    fs.copyFileSync(out, path.join(tmp, '9.3.5.0', 'p26.json'));
    const c = spawnSync('python3', ['-I', path.join(tmp, 'install_bridge.py'), '--target', 'cell', '--plan', path.join(tmp, '9.3.5.0', 'p26.json'), '--input', path.join(tmp, 'in.js'), '--check'], { encoding: 'utf8' });
    assert.equal(c.status, 0, c.stderr);
    assert.match(c.stdout, /STATUS=PATCHABLE PATCH_VERSION=euro-office-bridge-v1\/9\.3\.5\.0\/26/);
    fs.writeFileSync(path.join(tmp, 'bad.js'), Buffer.concat([data, Buffer.from('x')]));
    const bad = path.join(tmp, 'bad.json');
    const n = spawnSync('python3', ['-I', helper, '--input', path.join(tmp, 'bad.js'), '--out', bad], { encoding: 'utf8' });
    assert.equal(n.status, 2);
    assert.match(n.stderr, /UNKNOWN_INPUT/);
    assert.equal(fs.existsSync(bad), false);
  } finally { fs.rmSync(tmp, { recursive: true, force: true }); }
});
