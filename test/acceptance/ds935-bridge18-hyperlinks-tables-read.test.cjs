'use strict';

/**
 * GH #300 L1 / #298 D -- the euro-office-bridge-v1 /18 candidate for DS 9.3.5 (documentserver-patches/9.3.5.0/
 * euro-office-bridge-v1/rev18-candidate/): TWO new READ capabilities, sheet.hyperlinks.read/v1 and
 * sheet.tables.read/v1. Offline: the REAL candidate block is evaluated against a fake Api whose worksheet
 * model has the shape LIVE-measured on DS 9.3.5 (docs/evidence/gh300-gh298-engine-access-probe-20261027.json):
 * hyperlinkManager.getAll() -> [{bbox, data}], data carrying Hyperlink/Location/Tooltip/getHyperlinkType();
 * TableParts -> [{DisplayName, Ref, isHeaderRow(), isTotalsRow(), TableColumns, TableStyleInfo, AutoFilter}].
 * The "named fields only" owner condition is asserted with sentinel values planted on EVERY object the
 * handler touches, in fields it must never read.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const fs = require('node:fs');
const path = require('node:path');
const { spawnSync } = require('node:child_process');

const DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.5.0', 'euro-office-bridge-v1');
const V17 = fs.readFileSync(path.join(DIR, 'rev17-candidate', 'bridge-block.js'), 'utf8');
const V18 = fs.readFileSync(path.join(DIR, 'rev18-candidate', 'bridge-block.js'), 'utf8');
const CAP_H = 'sheet.hyperlinks.read/v1';
const CAP_T = 'sheet.tables.read/v1';
const sha = t => crypto.createHash('sha256').update(t).digest('hex');

function bridgeWith(sheets) {
  const Api = { GetSheet: n => (Object.prototype.hasOwnProperty.call(sheets, n) ? sheets[n] : null) };
  // the block is a fragment that ends with "}," (it is spliced into sdk-all.js before an anchor); the same
  // trailing-comma strip makes the shipped blocks parse, see the syntax test below
  new Function('Api', V18.replace(/\}\s*,\s*$/, '}'))(Api);
  return Api.GetEuroOfficeBridge();
}

/** a fake hyperlink record, in the LIVE-measured shape (bbox.getAbsName / data.{Hyperlink,Location,Tooltip,
 * getHyperlinkType}); `extra` plants an UNNAMED sentinel field on `data`, to prove it never leaks. */
function link(ref, { address = null, location = null, tooltip = null, kind = 1, extra } = {}) {
  const data = { Hyperlink: address, Location: location, Tooltip: tooltip, getHyperlinkType: () => kind };
  if (extra !== undefined) data._sentinel = extra;
  return { bbox: { getAbsName: () => ref }, data, _elSentinel: extra !== undefined ? extra : undefined };
}

/** a fake table part, in the LIVE-measured shape; `styleExtra`/`extra` plant UNNAMED sentinels. */
function table({ name = 'Table1', ref = '$B$10:$E$13', header = true, totals = false, columns = ['Region', 'Q1'],
  style = { Name: 'TableStyleLight9', ShowRowStripes: true, ShowColumnStripes: false, ShowFirstColumn: false, ShowLastColumn: false },
  autoFilter = {}, extra, styleExtra } = {}) {
  const t = {
    DisplayName: name,
    Ref: { getAbsName: () => ref },
    isHeaderRow: () => header,
    isTotalsRow: () => totals,
    TableColumns: columns.map(c => ({ Name: c })),
    TableStyleInfo: style ? Object.assign({}, style, styleExtra !== undefined ? { _styleSentinel: styleExtra } : {}) : null,
    AutoFilter: autoFilter,
  };
  if (extra !== undefined) t._sentinel = extra;
  return t;
}

/** a fake worksheet in the LIVE-measured shape */
function sheet({ links = [], tables = [], noManager = false, noTableParts = false, managerThrows = false, tableThrows = false } = {}) {
  const worksheet = {};
  if (!noManager) worksheet.hyperlinkManager = { getAll: () => { if (managerThrows) throw new Error('boom'); return links; } };
  if (!noTableParts) worksheet.TableParts = tableThrows ? { get length() { throw new Error('boom'); } } : tables;
  return { worksheet };
}

test('the /18 block = the /17 block + exactly the builder\'s substitutions', () => {
  const r = spawnSync('python3', ['-I', '-c', 'import importlib.util,json,sys;s=importlib.util.spec_from_file_location("b",sys.argv[1]);m=importlib.util.module_from_spec(s);s.loader.exec_module(m);print(json.dumps({"hEntry":m.HYPERLINKS_ENTRY,"tEntry":m.TABLES_ENTRY,"hImpl":m.HYPERLINKS_IMPL,"tImpl":m.TABLES_IMPL,"sha":m.V17_BLOCK_SHA}))',
    path.join(DIR, 'rev18-candidate', 'build-rev18.py')], { encoding: 'utf8' });
  assert.equal(r.status, 0, r.stderr);
  const b = JSON.parse(r.stdout);
  assert.equal(sha(V17), b.sha, 'the builder pins the /17 candidate block');
  for (const frag of [b.hEntry, b.tEntry, b.hImpl, b.tImpl]) assert.equal(V18.split(frag).length - 1, 1, `fragment occurs once: ${frag.slice(0, 40)}`);
  const back = V18.split(b.hEntry).join('').split(b.tEntry).join('').split(b.hImpl).join('').split(b.tImpl).join('')
    .replace('patchVersion: "euro-office-bridge-v1/9.3.5.0/18"', 'patchVersion: "euro-office-bridge-v1/9.3.5.0/17"');
  assert.equal(back, V17, 'nothing else moved');
  assert.equal(fs.readFileSync(path.join(DIR, 'rev18-candidate', 'old.txt'), 'utf8'), V17);
  assert.equal(fs.readFileSync(path.join(DIR, 'rev18-candidate', 'new.txt'), 'utf8'), V18);
});

test('both blocks parse the same way (the fragment trailing comma stripped)', () => {
  for (const t of [V17, V18]) assert.doesNotThrow(() => new Function('Api', t.replace(/\}\s*,\s*$/, '}')));
});

test('the manifest advertises /18 and both capabilities as non-mutating READ', () => {
  const m = bridgeWith({}).manifest();
  assert.equal(m.patchVersion, 'euro-office-bridge-v1/9.3.5.0/18');
  for (const id of [CAP_H, CAP_T]) {
    const c = m.capabilities.filter(x => x.id === id);
    assert.equal(c.length, 1, id);
    assert.equal(c[0].kind, 'READ');
    assert.equal(c[0].mutation, 'NO');
  }
});

// ---- sheet.hyperlinks.read/v1 --------------------------------------------------------------------

test('no hyperlinks on the sheet -> empty links (the NEG. KONTROLL sheet)', () => {
  const r = bridgeWith({ T2: sheet({ links: [] }) }).read(CAP_H, { sheet: 'T2' });
  assert.deepEqual(r, { ok: true, value: { links: [] } });
});

test('three links, row-then-column order, address/location/tooltip/kind as measured, $ stripped', () => {
  const s = sheet({ links: [
    link('$A$3', { address: 'a3-url', tooltip: 'tip', kind: 1 }),
    link('$C$1', { address: 'c1-url', kind: 2 }),
    link('$A$1', { address: 'a1-url', location: 'Other!B2', kind: 7 }),
  ] });
  const r = bridgeWith({ T1: s }).read(CAP_H, { sheet: 'T1' });
  assert.equal(r.ok, true);
  assert.deepEqual(r.value.links.map(l => l.ref), ['A1', 'C1', 'A3']);
  assert.deepEqual(r.value.links[0], { ref: 'A1', address: 'a1-url', location: 'Other!B2', tooltip: null, kind: 'unknown:7' });
  assert.deepEqual(r.value.links[1], { ref: 'C1', address: 'c1-url', location: null, tooltip: null, kind: 'internal' });
  assert.deepEqual(r.value.links[2], { ref: 'A3', address: 'a3-url', location: null, tooltip: 'tip', kind: 'external' });
});

test('a multi-cell hyperlink ref (a range address) sorts by its own first cell, not before everything', () => {
  const s = sheet({ links: [
    link('$A$10:$B$12', { address: 'range-url' }),
    link('$B$3', { address: 'b3-url' }),
    link('$C$5', { address: 'c5-url' }),
  ] });
  const r = bridgeWith({ T: s }).read(CAP_H, { sheet: 'T' });
  assert.equal(r.ok, true);
  assert.deepEqual(r.value.links.map(l => l.ref), ['B3', 'C5', 'A10:B12']);
});

test('an unnamed field on the hyperlink payload never reaches the output', () => {
  const s = sheet({ links: [link('$A$1', { address: 'a1-url', extra: 'SENTINEL-HLINK' })] });
  const r = bridgeWith({ T1: s }).read(CAP_H, { sheet: 'T1' });
  const out = JSON.stringify(r);
  assert.ok(!out.includes('SENTINEL-HLINK'), 'an unnamed payload field leaked');
});

test('hyperlink type-gates: a non-string address/location/tooltip, a non-number kind -- named refusals', () => {
  const b = bridgeWith({
    A: sheet({ links: [link('$A$1', { address: 5 })] }),
    L: sheet({ links: [link('$A$1', { location: 5 })] }),
    T: sheet({ links: [link('$A$1', { tooltip: 5 })] }),
    K: sheet({ links: [{ bbox: { getAbsName: () => 'A1' }, data: { Hyperlink: 'x', getHyperlinkType: () => 'not-a-number' } }] }),
  });
  assert.deepEqual(b.read(CAP_H, { sheet: 'A' }), { ok: false, reason: 'hyperlink-field-not-string:address' });
  assert.deepEqual(b.read(CAP_H, { sheet: 'L' }), { ok: false, reason: 'hyperlink-field-not-string:location' });
  assert.deepEqual(b.read(CAP_H, { sheet: 'T' }), { ok: false, reason: 'hyperlink-field-not-string:tooltip' });
  assert.deepEqual(b.read(CAP_H, { sheet: 'K' }), { ok: false, reason: 'hyperlink-kind-not-number' });
});

test('a bbox that cannot resolve to an A1 string is range-unreadable', () => {
  const s = sheet({ links: [{ bbox: { getAbsName: () => 42 }, data: { Hyperlink: 'x', getHyperlinkType: () => 1 } }] });
  assert.deepEqual(bridgeWith({ T: s }).read(CAP_H, { sheet: 'T' }), { ok: false, reason: 'range-unreadable' });
});

test('the hyperlink walk is capped: more than 10000 links is a refusal, never a truncated list', () => {
  const links = []; for (let i = 1; i <= 10001; i += 1) links.push(link('$A$' + i, { address: 'x' }));
  const r = bridgeWith({ T: sheet({ links }) }).read(CAP_H, { sheet: 'T' });
  assert.deepEqual(r, { ok: false, reason: 'links-over-cap:10000' });
});

test('hyperlinks: named refusals -- bad args, missing sheet, missing manager, a throwing getAll', () => {
  const b = bridgeWith({ P: sheet({ links: [] }), N: sheet({ noManager: true }), X: sheet({ managerThrows: true }) });
  assert.deepEqual(b.read(CAP_H, {}), { ok: false, reason: 'bad-args' });
  assert.deepEqual(b.read(CAP_H, { sheet: 'nope' }), { ok: false, reason: 'sheet-not-found' });
  assert.deepEqual(b.read(CAP_H, { sheet: 'N' }), { ok: false, reason: 'hyperlink-manager-missing' });
  assert.deepEqual(b.read(CAP_H, { sheet: 'X' }), { ok: false, reason: 'hyperlinks-threw' });
});

// ---- sheet.tables.read/v1 -------------------------------------------------------------------------

test('no tables on the sheet -> empty tables (the NEG. KONTROLL sheet)', () => {
  const r = bridgeWith({ T2: sheet({ tables: [] }) }).read(CAP_T, { sheet: 'T2' });
  assert.deepEqual(r, { ok: true, value: { tables: [] } });
});

test('two tables, all named fields as measured, $ stripped from ref, autoFilter a boolean presence flag', () => {
  const s = sheet({ tables: [
    table({ name: 'Table1', ref: '$B$10:$E$13', columns: ['Region', 'Q1', 'Q2', 'Q3'], autoFilter: {} }),
    table({ name: 'Table2', ref: '$B$16:$C$18', columns: ['Name', 'Score'], style: null, autoFilter: null }),
  ] });
  const r = bridgeWith({ T1: s }).read(CAP_T, { sheet: 'T1' });
  assert.equal(r.ok, true);
  assert.equal(r.value.tables.length, 2);
  assert.deepEqual(r.value.tables[0], { name: 'Table1', ref: 'B10:E13', headerRow: true, totalsRow: false, columns: ['Region', 'Q1', 'Q2', 'Q3'], style: { name: 'TableStyleLight9', rowStripes: true, columnStripes: false, firstColumn: false, lastColumn: false }, autoFilter: true });
  assert.deepEqual(r.value.tables[1], { name: 'Table2', ref: 'B16:C18', headerRow: true, totalsRow: false, columns: ['Name', 'Score'], style: null, autoFilter: false });
});

test('unnamed fields on the table object and on its style info never reach the output', () => {
  const s = sheet({ tables: [table({ extra: 'SENTINEL-TABLE', styleExtra: 'SENTINEL-STYLE' })] });
  const out = JSON.stringify(bridgeWith({ T: s }).read(CAP_T, { sheet: 'T' }));
  assert.ok(!out.includes('SENTINEL-TABLE'), 'an unnamed table field leaked');
  assert.ok(!out.includes('SENTINEL-STYLE'), 'an unnamed style field leaked');
});

test('table type-gates: a non-string name/style-name, a non-boolean headerRow/style-flag -- named refusals', () => {
  const b = bridgeWith({
    N: sheet({ tables: [table({ name: 5 })] }),
    H: sheet({ tables: [Object.assign(table({}), { isHeaderRow: () => 'yes' })] }),
    SN: sheet({ tables: [table({ style: { Name: 5, ShowRowStripes: true, ShowColumnStripes: true, ShowFirstColumn: true, ShowLastColumn: true } })] }),
    SF: sheet({ tables: [table({ style: { Name: 'x', ShowRowStripes: 'yes', ShowColumnStripes: true, ShowFirstColumn: true, ShowLastColumn: true } })] }),
    CN: sheet({ tables: [table({ columns: [] })], }),
  });
  assert.deepEqual(b.read(CAP_T, { sheet: 'N' }), { ok: false, reason: 'table-field-not-string:name' });
  assert.deepEqual(b.read(CAP_T, { sheet: 'H' }), { ok: false, reason: 'table-field-not-boolean:headerRow' });
  assert.deepEqual(b.read(CAP_T, { sheet: 'SN' }), { ok: false, reason: 'table-field-not-string:style.name' });
  assert.deepEqual(b.read(CAP_T, { sheet: 'SF' }), { ok: false, reason: 'table-field-not-boolean:style.ShowRowStripes' });
  assert.deepEqual(b.read(CAP_T, { sheet: 'CN' }).ok, true, 'zero columns is a valid (if unusual) table, not a refusal');
});

test('a throwing AutoFilter getter is a named refusal, never a guessed value', () => {
  const t = table({});
  Object.defineProperty(t, 'AutoFilter', { get() { throw new Error('boom'); }, configurable: true });
  assert.deepEqual(bridgeWith({ T: sheet({ tables: [t] }) }).read(CAP_T, { sheet: 'T' }), { ok: false, reason: 'tables-threw' });
});

test('a column whose Name is not a string is a named refusal', () => {
  const s = sheet({ tables: [Object.assign(table({}), { TableColumns: [{ Name: 'ok' }, { Name: 5 }] })] });
  assert.deepEqual(bridgeWith({ T: s }).read(CAP_T, { sheet: 'T' }), { ok: false, reason: 'table-field-not-string:columns[].Name' });
});

test('a Ref that cannot resolve to an A1 string is range-unreadable', () => {
  const s = sheet({ tables: [Object.assign(table({}), { Ref: { getAbsName: () => null } })] });
  assert.deepEqual(bridgeWith({ T: s }).read(CAP_T, { sheet: 'T' }), { ok: false, reason: 'range-unreadable' });
});

test('the table walk is capped: more than 1000 tables is a refusal, never a truncated list', () => {
  const tables = []; for (let i = 1; i <= 1001; i += 1) tables.push(table({ name: 'T' + i, ref: '$A$1:$B$2' }));
  assert.deepEqual(bridgeWith({ T: sheet({ tables }) }).read(CAP_T, { sheet: 'T' }), { ok: false, reason: 'tables-over-cap:1000' });
});

test('the column walk is capped: more than 16384 columns is a refusal, never a truncated list', () => {
  const columns = []; for (let i = 1; i <= 16385; i += 1) columns.push('C' + i);
  assert.deepEqual(bridgeWith({ T: sheet({ tables: [table({ columns })] }) }).read(CAP_T, { sheet: 'T' }), { ok: false, reason: 'columns-over-cap:16384' });
});

test('tables: named refusals -- bad args, missing sheet, missing TableParts, a throwing TableParts access', () => {
  const b = bridgeWith({ P: sheet({ tables: [] }), N: sheet({ noTableParts: true }), X: sheet({ tableThrows: true }) });
  assert.deepEqual(b.read(CAP_T, {}), { ok: false, reason: 'bad-args' });
  assert.deepEqual(b.read(CAP_T, { sheet: 'nope' }), { ok: false, reason: 'sheet-not-found' });
  assert.deepEqual(b.read(CAP_T, { sheet: 'N' }), { ok: false, reason: 'table-parts-missing' });
  assert.deepEqual(b.read(CAP_T, { sheet: 'X' }), { ok: false, reason: 'tables-threw' });
});

// ---- shared: read only, no model mutation ----------------------------------------------------------

test('both handlers call no setter: the fake model is unchanged after every read', () => {
  const links = [link('$A$1', { address: 'a1-url' })];
  const tables = [table({})];
  const s = sheet({ links, tables });
  const before = JSON.stringify({ links: s.worksheet.hyperlinkManager.getAll().map(l => ({ ref: l.bbox.getAbsName(), data: l.data })), tables: s.worksheet.TableParts });
  bridgeWith({ P: s }).read(CAP_H, { sheet: 'P' });
  bridgeWith({ P: s }).read(CAP_T, { sheet: 'P' });
  const after = JSON.stringify({ links: s.worksheet.hyperlinkManager.getAll().map(l => ({ ref: l.bbox.getAbsName(), data: l.data })), tables: s.worksheet.TableParts });
  assert.equal(after, before);
});

test('plan-rev18.py derives a plan the UNCHANGED install_bridge.py accepts; an unknown input is refused, nothing written', () => {
  const os = require('node:os');
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'eo-b18-plan-'));
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
    const helper = path.join(tmp, '9.3.5.0', 'euro-office-bridge-v1', 'rev18-candidate', 'plan-rev18.py');
    const out = path.join(tmp, 'p18.json');
    const d = spawnSync('python3', ['-I', helper, '--input', path.join(tmp, 'in.js'), '--out', out], { encoding: 'utf8' });
    assert.equal(d.status, 0, d.stderr);
    const p18 = JSON.parse(fs.readFileSync(out, 'utf8'));
    assert.equal(p18.cell.patchVersion, 'euro-office-bridge-v1/9.3.5.0/18');
    assert.equal(p18.cell.blocks.find(b => b.file.startsWith('euro-office-bridge-v1/')).sha256, sha(V18));
    fs.copyFileSync(out, path.join(tmp, '9.3.5.0', 'p18.json'));
    const c = spawnSync('python3', ['-I', path.join(tmp, 'install_bridge.py'), '--target', 'cell', '--plan', path.join(tmp, '9.3.5.0', 'p18.json'), '--input', path.join(tmp, 'in.js'), '--check'], { encoding: 'utf8' });
    assert.equal(c.status, 0, c.stderr);
    assert.match(c.stdout, /STATUS=PATCHABLE PATCH_VERSION=euro-office-bridge-v1\/9\.3\.5\.0\/18/);
    fs.writeFileSync(path.join(tmp, 'bad.js'), Buffer.concat([data, Buffer.from('x')]));
    const bad = path.join(tmp, 'bad.json');
    const n = spawnSync('python3', ['-I', helper, '--input', path.join(tmp, 'bad.js'), '--out', bad], { encoding: 'utf8' });
    assert.equal(n.status, 2);
    assert.match(n.stderr, /UNKNOWN_INPUT/);
    assert.equal(fs.existsSync(bad), false);
  } finally { fs.rmSync(tmp, { recursive: true, force: true }); }
});
