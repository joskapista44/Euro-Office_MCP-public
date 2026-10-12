'use strict';

/**
 * Static gate: no NEW direct ONLYOFFICE internal-model dependency (owner decision, owner channel 2026-10-08, point 7).
 * SOURCE KNOWN_LEGACY_SET=0 since M11.5e.4 (the B21 Bridge READ cutover removed the last source use; M11.5d.4 removed
 * B18 before it; before that: 4 = B03 B10 B18 B21; the B05-WRITE Bridge cutover removed B05; before it: 5 = B03 B05 B10
 * B18 B21; Wave B READ /4 removed B09, B13 and B05's READ half; the B15 /5 RUNTIME_HYDRATION cutover removed B15; B21
 * was added by owner decision M11.3a.1 -- an inventory defect, present since 2026-09-25), NEW_DIRECT_INTERNAL=0.
 * M11.5e.5 (B21's own D-state LIVE acceptance PASS): B21-READ = REMOVED_FROM_LEGACY -- ACCEPTED_OPERATIONAL_LEGACY_SET
 * is empty again, equal to KNOWN_LEGACY_SET (0, both sets). Rules: test/static/direct-internal-rules.cjs.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const R = require('./direct-internal-rules.cjs');

const SRC = path.join(__dirname, '..', '..', 'src');
function srcModules(dir = SRC, out = []) {
  for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
    const p = path.join(dir, e.name);
    if (e.isDirectory()) srcModules(p, out);
    else if (p.endsWith('.cjs')) out.push({ file: path.relative(SRC, p).split(path.sep).join('/'), source: fs.readFileSync(p, 'utf8') });
  }
  return out;
}

test('NEW_DIRECT_INTERNAL=0: every internal-model use in src/ is a grandfathered legacy use', () => {
  const mods = srcModules();
  assert.ok(mods.length > 30, 'the walk sees the source tree');
  const v = R.judge(mods);
  assert.deepEqual(v.newDirectInternal, [], 'a new direct internal dependency: migrate it to public/Bridge, or owner decision');
});

test('the legacy register is exact: a migration that removes uses must shrink (or delete) its entry in the same change', () => {
  assert.deepEqual(R.judge(srcModules()).staleLegacy, []);
});

test('SOURCE_LEGACY_REGISTER_COUNT=0 and frozen (M11.5e.4 B21 cutover): the register is empty, and the set may only shrink', () => {
  assert.deepEqual([...R.KNOWN_LEGACY_SET], []);
  assert.equal(Object.isFrozen(R.KNOWN_LEGACY_SET), true);
  const v = R.judge(srcModules());
  assert.deepEqual(v.unknownIds, []);
  assert.deepEqual(v.legacyIds, [...R.KNOWN_LEGACY_SET], 'every legacy id still has uses; a migrated id leaves BOTH lists');
});

test('NEG controls: the gate goes red on a new internal use, on a grown legacy use, and on an unknown id', () => {
  const base = srcModules();
  const withNewFile = [...base, { file: 'adapters/xlsx/families/new.cjs', source: "const b = `ws.worksheet.foo(); e.asc_setSomething(1);`;" }];
  assert.deepEqual(R.judge(withNewFile).newDirectInternal.map(x => x.token).sort(), ['.asc_setSomething', '.worksheet']);
  const grown = base.map(m => (m.file === 'drivers/eurooffice/xlsx-driver.cjs' ? { ...m, source: `${m.source}\nconst x = \`ws.worksheet.getIndex()\`;` } : m));
  assert.deepEqual(R.judge(grown).newDirectInternal, [{ file: 'drivers/eurooffice/xlsx-driver.cjs', token: '.worksheet', found: 1, grandfathered: 0 }]);
  assert.deepEqual(R.judge(base, [...R.LEGACY, { id: 'B99', file: 'x.cjs', token: 'AscCommon', count: 1 }]).unknownIds, ['B99']);
  // a comment naming an internal API is not a use
  assert.deepEqual(R.judge([{ file: 'core/x.cjs', source: '// the old path called ws.worksheet.asc_getPageSetup()\n/* AscCommon.History */' }]).newDirectInternal, []);
});

test('Wave B READ /4: a removed B05-READ / B09 / B13 token reintroduced into the driver goes red (grandfathered 0)', () => {
  const base = srcModules();
  const DRV = 'drivers/eurooffice/xlsx-driver.cjs';
  const back = [
    ['B05 READ', 'const ro = al.getReadingOrder(); const i = al.getIndent(); const s = al.getShrinkToFit();', ['getIndent', 'getReadingOrder', 'getShrinkToFit']],
    ['B09', 'const ps = m.PagePrintOptions.asc_getPageSetup(); ps.asc_getFitToWidth(); ps.asc_getScale();', ['.asc_getFitToWidth', '.asc_getPageSetup', '.asc_getScale', 'PagePrintOptions']],
    ['B13', 'const d = String(p.pivot.asc_getDataRef());', ['.asc_getDataRef']],
  ];
  for (const [group, code, tokens] of back) {
    const v = R.judge(base.map(m => (m.file === DRV ? { ...m, source: `${m.source}\nconst q = \`${code}\`;` } : m)));
    assert.deepEqual(v.newDirectInternal.map(x => x.token).sort(), tokens, group);
    assert.ok(v.newDirectInternal.every(x => x.file === DRV && x.grandfathered === 0), `${group}: no allowance left for it`);
  }
  // M11.5a.4: B05's WRITE half is gone too -- no B05 entry is left anywhere
  assert.equal(R.LEGACY.some(e => e.id === 'B05'), false);
  assert.equal(R.LEGACY.some(e => e.id === 'B09' || e.id === 'B13'), false);
});

test('B15 /5: the removed direct hydration reintroduced into the driver goes red (grandfathered 0, no B15 entry left)', () => {
  const DRV = 'drivers/eurooffice/xlsx-driver.cjs';
  const code = 'const m = ws.worksheet; const v = m.workbook.oApi.wb.getWorksheet(m.getIndex());';
  const v = R.judge(srcModules().map(m => (m.file === DRV ? { ...m, source: `${m.source}\nconst q = \`${code}\`;` } : m)));
  assert.deepEqual(v.newDirectInternal.map(x => [x.token, x.grandfathered]).sort(), [['.getWorksheet', 0], ['.worksheet', 0]]);
  assert.equal(R.LEGACY.some(e => e.id === 'B15'), false);
  assert.equal(R.KNOWN_LEGACY_SET.includes('B15'), false);
});

test('M11.5a.4 B05-WRITE cutover: setIndent=0 setShrinkToFit=0 in src/, and a direct setter put back goes RED (grandfathered 0)', () => {
  const base = srcModules();
  const counts = R.countInternal(base);
  const all = Object.values(counts).reduce((n, t) => n + (t.setIndent || 0) + (t.setShrinkToFit || 0), 0);
  assert.equal(all, 0, 'B05_DIRECT_INTERNAL_TOKENS=0');
  const FMT = 'adapters/xlsx/families/format.cjs';
  for (const [code, token] of [["c.range.setIndent(2);", 'setIndent'], ["c.range.setShrinkToFit(true);", 'setShrinkToFit']]) {
    const v = R.judge(base.map(m => (m.file === FMT ? { ...m, source: `${m.source}\nconst q = \`${code}\`;` } : m)));
    assert.deepEqual(v.newDirectInternal, [{ file: FMT, token, found: 1, grandfathered: 0 }], token);
  }
});

test('M11.5a.5 / M11.5b.5 / M11.5c.5 / M11.5d.5 / M11.5e.5: B05-, B10-, B03-, B18- and B21-{WRITE,READ} REMOVED_FROM_LEGACY (in neither set)', () => {
  assert.equal(R.KNOWN_LEGACY_SET.includes('B18') || R.ACCEPTED_OPERATIONAL_LEGACY_SET.includes('B18'), false, 'B18');
  assert.equal(R.KNOWN_LEGACY_SET.includes('B03') || R.ACCEPTED_OPERATIONAL_LEGACY_SET.includes('B03'), false, 'B03');
  assert.equal(R.KNOWN_LEGACY_SET.includes('B21') || R.ACCEPTED_OPERATIONAL_LEGACY_SET.includes('B21'), false, 'B21');
  for (const id of ['B05', 'B10']) assert.equal(R.KNOWN_LEGACY_SET.includes(id) || R.ACCEPTED_OPERATIONAL_LEGACY_SET.includes(id), false, id);
});

test('M11.5e.5: the B21 Bridge READ cutover reaches its OWN D-state LIVE acceptance PASS -- the legacy set is empty and SYMMETRIC (KNOWN == ACCEPTED == []), unlike the M11.5e.4 mid-migration step this closure follows', () => {
  assert.deepEqual([...R.KNOWN_LEGACY_SET], []);
  assert.equal(Object.isFrozen(R.KNOWN_LEGACY_SET), true);
  assert.deepEqual([...R.ACCEPTED_OPERATIONAL_LEGACY_SET], []);
  assert.equal(Object.isFrozen(R.ACCEPTED_OPERATIONAL_LEGACY_SET), true);
  assert.deepEqual([...R.ACCEPTED_OPERATIONAL_LEGACY_SET], [...R.KNOWN_LEGACY_SET]);
});

test('M11.5b.4: the B10 direct internals are gone from print.cjs, and a re-added print setter is a NEW direct internal (RED)', () => {
  const base = srcModules(); const P = 'adapters/xlsx/families/print.cjs';
  assert.deepEqual(R.judge(base).newDirectInternal, []);
  const back = base.map(m => (m.file === P ? { ...m, source: `${m.source}\nconst q = \`m.PagePrintOptions.asc_getPageSetup().asc_setScale(1);\`;` } : m));
  assert.ok(R.judge(back).newDirectInternal.some(x => x.file === P), 'a print setter back in the product = new direct internal');
});

test('M11.5c.4: the B03 direct internals are gone from the xlsx driver, and a re-added asc_copyWorksheet is a NEW direct internal (RED)', () => {
  const base = srcModules(); const D = 'drivers/eurooffice/xlsx-driver.cjs';
  assert.deepEqual(R.judge(base).newDirectInternal, []);
  for (const code of ['e.asc_copyWorksheet(-1, [dst], [index]);', 'e.asc_getWorksheetsCount();', "e.asc_registerCallback('asc_onError', f);"]) {
    const back = base.map(m => (m.file === D ? { ...m, source: `${m.source}\nconst q = \`${code}\`;` } : m));
    assert.ok(R.judge(back).newDirectInternal.some(x => x.file === D), code);
  }
});

test('M11.5d.4: the B18 direct internals are gone from the DOCX product, and a re-added RemoveNumPr / CParagraph is a NEW direct internal (RED)', () => {
  const base = srcModules();
  assert.deepEqual(R.judge(base).newDirectInternal, []);
  for (const F of ['adapters/docx/families/list.cjs', 'drivers/eurooffice/docx-driver.cjs', 'drivers/eurooffice/ds-bridge.cjs']) {
    for (const code of ['p.Paragraph.RemoveNumPr();', "const k = 'CParagraph';"]) {
      const back = base.map(m => (m.file === F ? { ...m, source: `${m.source}\nconst q = \`${code}\`;` } : m));
      assert.ok(R.judge(back).newDirectInternal.some(x => x.file === F), `${F}: ${code} back in the product = new direct internal`);
    }
  }
});

test('M11.5e.4: the B21 direct internals are gone from the DOCX product, and a re-added .Section / Get_TitlePage / IsEvenAndOdd -- direct or as a dynamic name string -- is a NEW direct internal (RED, grandfathered 0)', () => {
  const base = srcModules();
  assert.deepEqual(R.judge(base).newDirectInternal, []);
  for (const F of ['drivers/eurooffice/docx-driver.cjs', 'adapters/docx/families/header-footer.cjs']) {
    for (const code of ["const v = sec.Section.Get_TitlePage();", "const name = 'IsEvenAndOdd'; inner[name]();", "const x = doc.Section;"]) {
      const back = base.map(m => (m.file === F ? { ...m, source: `${m.source}\nconst q = \`${code}\`;` } : m));
      const v = R.judge(back);
      assert.ok(v.newDirectInternal.some(x => x.file === F && x.grandfathered === 0), `${F}: ${code} back in the product = new direct internal`);
    }
  }
  // negative control: public API use is not internal
  assert.deepEqual(R.judge([{ file: 'adapters/docx/families/new-family.cjs', source: 'const s = doc.GetSections(); s[0].GetHeader("default", false);' }]).newDirectInternal, []);
});
