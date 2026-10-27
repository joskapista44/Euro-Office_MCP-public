'use strict';

// Owner decision (B) (the owner, owner channel 2026-10-08, internal tracker): the DS runtime identity is the served
// bundle's own bridge manifest, read through ONE canonical handshake; the editor URL's '/9.3.4-<hash>/web-apps/'
// segment is an ephemeral DS cache key (a restart regenerated it on unchanged content) and is never a gate.
//
// DS9346 (owner D5=B, VALTAS, docs/DS9346-RUNTIME-RETIREMENT.md): the entire 9.3.4.60 runtime family (/5 through
// /16 xlsx, /1 through /3 docx) is RETIRED -- ACCEPTED_RUNTIMES carries ONE entry per format now (9.3.5.0/16
// xlsx, 9.3.5.0/3 docx). Every test below that used to assert "this is the accepted runtime, that candidate
// isn't yet" now asserts the single accepted identity directly; every test whose ENTIRE subject was the
// 9.3.4.60-era progressive-widening narrative (candidate N joins the list, the off-list example moves to N+1)
// is retired along with it -- a single-entry list has no "safely unclaimed number" to chase.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { RUNTIMES, XLSX_RUNTIME_17, ACCEPTED_RUNTIMES, XLSX_RUNTIME_ID, XLSX_RUNTIME_IDS, DOCX_RUNTIME_ID, DOCX_RUNTIME_IDS, runtimeIdOf, cacheKeyOf, HANDSHAKE_BODY, classifyHandshake } = require('../../src/core/runtime-identity.cjs');
const XLSX_RUNTIME_17_ID = runtimeIdOf(XLSX_RUNTIME_17);

// manifest()'s default IS the accepted xlsx identity now -- positive-control tests need no override;
// a test that needs the (now entirely retired) 9.3.4.60 family overrides documentServerVersion/patchVersion
// explicitly, exactly as it already did for every OTHER non-default case.
const manifest = (over = {}) => ({
  bridgeVersion: '1.0.0', documentServerVersion: '9.3.5', documentServerBuild: '9.3.5-9475e9f12a1905e2ada1fbc4bb87d01b',
  patchVersion: 'euro-office-bridge-v1/9.3.5.0/16',
  capabilities: [{ id: 'range.validations.read/v1', version: 1, kind: 'READ', format: 'XLSX', mutation: 'NO', semanticContract: 'x' }],
  ...over,
});
// eslint-disable-next-line no-new-func
const handshake = Api => new Function('Api', HANDSHAKE_BODY)(Api);
const apiWith = m => ({ GetEuroOfficeBridge: () => ({ manifest: () => m, read: () => ({ ok: true, value: null }) }) });

test('the accepted runtimes: DS 9.3.5, cell bridge /16, word bridge /5 -- keyed by (version, patchVersion)', () => {
  assert.equal(XLSX_RUNTIME_ID, '9.3.5|euro-office-bridge-v1/9.3.5.0/16');
  assert.equal(DOCX_RUNTIME_ID, '9.3.5|euro-office-bridge-word-v1/9.3.5.0/5');
  assert.equal(runtimeIdOf(RUNTIMES.xlsx), XLSX_RUNTIME_ID);
});

test('PASS: the handshake on the served manifest yields the runtime id; documentServerBuild is not part of it', () => {
  const c = classifyHandshake(handshake(apiWith(manifest())), RUNTIMES.xlsx);
  assert.equal(c.ok, true);
  assert.equal(c.runtimeId, XLSX_RUNTIME_ID);
  assert.equal('documentServerBuild' in c.manifest, false, 'the handshake does not even carry the patch-time cache key');
  assert.deepEqual(c.manifest.capabilities, [{ id: 'range.validations.read/v1', version: 1, kind: 'READ', mutation: 'NO' }]);
});

test('CACHE-KEY CHANGE REGRESSION: the URL cache key changes, the manifest is the same -> the SAME runtime id (PASS)', () => {
  const before = 'https://eo/9.3.5-9475e9f12a1905e2ada1fbc4bb87d01b/web-apps/apps/spreadsheeteditor/main/index.html';
  const after = 'https://eo/9.3.5-94bc577ecccd9583be1b130bd5a6b407/web-apps/apps/spreadsheeteditor/main/index.html';
  assert.notEqual(cacheKeyOf(before), cacheKeyOf(after));
  const a = classifyHandshake(handshake(apiWith(manifest())), RUNTIMES.xlsx);
  const b = classifyHandshake(handshake(apiWith(manifest({ documentServerBuild: cacheKeyOf(after) }))), RUNTIMES.xlsx);
  assert.equal(a.ok && b.ok, true);
  assert.equal(a.runtimeId, b.runtimeId);
});

test('cacheKeyOf: diagnostic only -- the key or null', () => {
  assert.equal(cacheKeyOf('https://eo/9.3.5-94bc577ecccd9583be1b130bd5a6b407/web-apps/x'), '9.3.5-94bc577ecccd9583be1b130bd5a6b407');
  assert.equal(cacheKeyOf('https://eo/web-apps/x'), null);
  assert.equal(cacheKeyOf(undefined), null);
});

test('FAIL_CLOSED: missing bridge / manifest, malformed manifest, wrong bridge major, wrong runtime, wrong patch', () => {
  const cases = [
    [handshake({}), 'BRIDGE_UNAVAILABLE'],
    [handshake({ GetEuroOfficeBridge: () => { throw new Error('x'); } }), 'BRIDGE_UNAVAILABLE'],
    [handshake({ GetEuroOfficeBridge: () => ({ manifest: () => { throw new Error('x'); } }) }), 'BRIDGE_UNAVAILABLE'],
    [handshake(apiWith(null)), 'MANIFEST_MISSING'],
    [handshake(apiWith(manifest({ bridgeVersion: undefined }))), 'MANIFEST_MALFORMED'],
    [handshake(apiWith(manifest({ documentServerVersion: 934 }))), 'MANIFEST_MALFORMED'],
    [handshake(apiWith(manifest({ patchVersion: '' }))), 'MANIFEST_MALFORMED'],
    [handshake(apiWith(manifest({ capabilities: 'all' }))), 'MANIFEST_MALFORMED'],
    [handshake(apiWith(manifest({ bridgeVersion: '2.0.0' }))), 'BRIDGE_VERSION_MISMATCH'],
    // DS9346: 9.3.4 is now the UNACCEPTED version (every 9.3.4.60 entry retired) -- the inverse of
    // this test's pre-retirement shape, same reason code (the version itself is not in the list).
    [handshake(apiWith(manifest({ documentServerVersion: '9.3.4' }))), 'RUNTIME_VERSION_MISMATCH'],
    [handshake(apiWith(manifest({ patchVersion: 'euro-office-bridge-v1/9.3.4.60/2' }))), 'PATCH_VERSION_MISMATCH'],
    [handshake(apiWith(manifest({ patchVersion: 'euro-office-bridge-word-v1/9.3.4.60/1' }))), 'PATCH_VERSION_MISMATCH'],
    [null, 'MANIFEST_MALFORMED'],
    [{ ok: 'yes' }, 'MANIFEST_MALFORMED'],
    [{ ok: true }, 'MANIFEST_MALFORMED'],
  ];
  for (const [r, reason] of cases) assert.deepEqual(classifyHandshake(r, RUNTIMES.xlsx), { ok: false, reason }, JSON.stringify(r));
});

test('the word runtime is its own identity: the cell manifest does not pass as the word runtime, and back', () => {
  assert.equal(classifyHandshake(handshake(apiWith(manifest())), RUNTIMES.docx).reason, 'PATCH_VERSION_MISMATCH');
  const word = manifest({ patchVersion: 'euro-office-bridge-word-v1/9.3.5.0/5' });
  assert.equal(classifyHandshake(handshake(apiWith(word)), RUNTIMES.docx).runtimeId, DOCX_RUNTIME_ID);
});

test('HANDSHAKE_BODY reads the manifest only -- it never calls bridge.read or writes', () => {
  let reads = 0;
  handshake({ GetEuroOfficeBridge: () => ({ manifest: () => manifest(), read: () => { reads += 1; } }) });
  assert.equal(reads, 0);
  assert.equal(/\.read\(|Set[A-Z]|Add[A-Z]|Delete/.test(HANDSHAKE_BODY), false);
});

// LEGACY PIN INVENTORY GUARD: no cache-key literal may stand in product code -- the ephemeral key must not come back
// as a compatibility authority (a comment may still name a measured key as history).
test('STATIC GUARD: no 9.3.4-<hash> cache-key literal in any src/ string or code token', () => {
  const root = path.join(__dirname, '../../src');
  const files = [];
  const walk = d => { for (const e of fs.readdirSync(d, { withFileTypes: true })) { const p = path.join(d, e.name); if (e.isDirectory()) walk(p); else if (p.endsWith('.cjs') || p.endsWith('.js')) files.push(p); } };
  walk(root);
  assert.ok(files.length > 20, 'the walk sees the source tree');
  const offenders = [];
  for (const f of files) {
    fs.readFileSync(f, 'utf8').split('\n').forEach((line, i) => {
      const code = line.replace(/\/\/.*$/, '').replace(/^\s*\*.*$/, '').replace(/\/\*.*?\*\//g, '');
      if (/['"`]\d+\.\d+\.\d+-[0-9a-f]{8,}['"`]/.test(code)) offenders.push(`${path.relative(root, f)}:${i + 1}`);
    });
  }
  assert.deepEqual(offenders, []);
});

test('STATIC GUARD: every runtime-gated list in src is built from a core/runtime-identity.cjs RUNTIME_ID (DS9346: each now a SINGLE 9.3.5.0 entry)', () => {
  const root = path.join(__dirname, '../../src');
  const lists = [];
  const walk = d => { for (const e of fs.readdirSync(d, { withFileTypes: true })) { const p = path.join(d, e.name); if (e.isDirectory()) walk(p); else if (p.endsWith('.cjs')) fs.readFileSync(p, 'utf8').split('\n').forEach((l, i) => { if (/^const \w+_RUNTIMES = (Object\.freeze\(\[|DOCX_RUNTIME_IDS|DOCX_B18_BRIDGE_WRITE_RUNTIME_IDS|DOCX_B21_BRIDGE_READ_RUNTIME_IDS|XLSX_RUNTIME_IDS|XLSX_B05_BRIDGE_WRITE_RUNTIME_IDS|XLSX_B10_BRIDGE_WRITE_RUNTIME_IDS|XLSX_B03_BRIDGE_WRITE_RUNTIME_IDS|XLSX_B22_BRIDGE_WRITE_RUNTIME_IDS|XLSX_B23_BRIDGE_WRITE_RUNTIME_IDS)/.test(l)) lists.push([`${path.relative(root, p)}:${i + 1}`, l]); }); } };
  walk(root);
  assert.ok(lists.length >= 10, `found ${lists.length} runtime lists`);
  for (const [where, l] of lists) assert.match(l, /= (DOCX_RUNTIME_IDS|DOCX_B18_BRIDGE_WRITE_RUNTIME_IDS|DOCX_B21_BRIDGE_READ_RUNTIME_IDS|XLSX_RUNTIME_IDS|XLSX_B05_BRIDGE_WRITE_RUNTIME_IDS|XLSX_B10_BRIDGE_WRITE_RUNTIME_IDS|XLSX_B03_BRIDGE_WRITE_RUNTIME_IDS|XLSX_B22_BRIDGE_WRITE_RUNTIME_IDS|XLSX_B23_BRIDGE_WRITE_RUNTIME_IDS);$/, where);
  // the B05/B10/B03/B22/B23 Bridge WRITE lists (adapter + driver) are each the ONLY declarers of their own list name
  assert.deepEqual(lists.filter(([, l]) => /XLSX_B05_BRIDGE_WRITE_RUNTIME_IDS;$/.test(l)).map(([w]) => w.replace(/:\d+$/, '')).sort(), ['adapters/xlsx/families/format.cjs', 'drivers/eurooffice/xlsx-driver.cjs']);
  assert.deepEqual(lists.filter(([, l]) => /XLSX_B10_BRIDGE_WRITE_RUNTIME_IDS;$/.test(l)).map(([w]) => w.replace(/:\d+$/, '')).sort(), ['adapters/xlsx/families/print.cjs', 'drivers/eurooffice/xlsx-driver.cjs']);
  assert.deepEqual(lists.filter(([, l]) => /XLSX_B03_BRIDGE_WRITE_RUNTIME_IDS;$/.test(l)).map(([w]) => w.replace(/:\d+$/, '')).sort(), ['adapters/xlsx/families/sheets.cjs', 'drivers/eurooffice/xlsx-driver.cjs']);
  assert.deepEqual(lists.filter(([, l]) => /XLSX_B22_BRIDGE_WRITE_RUNTIME_IDS;$/.test(l)).map(([w]) => w.replace(/:\d+$/, '')).sort(), ['adapters/xlsx/families/charts.cjs', 'drivers/eurooffice/xlsx-driver.cjs']);
  assert.deepEqual(lists.filter(([, l]) => /XLSX_B23_BRIDGE_WRITE_RUNTIME_IDS;$/.test(l)).map(([w]) => w.replace(/:\d+$/, '')).sort(), ['adapters/xlsx/families/charts.cjs', 'drivers/eurooffice/xlsx-driver.cjs']);
});

test('the XLSX adapter\'s own runtime copy (it imports nothing outside its tree) equals the core identity -- every Bridge WRITE list now carries the SAME two 9.3.5.0 entries (/16 + /17, GH #326/#327 compatibility-first)', () => {
  assert.equal(require('../../src/adapters/xlsx/runtime.cjs').XLSX_RUNTIME_ID, XLSX_RUNTIME_ID);
  assert.deepEqual([...require('../../src/adapters/xlsx/runtime.cjs').XLSX_RUNTIME_IDS], [...XLSX_RUNTIME_IDS]);
  const { PIVOT_STATE_BRIDGE_RUNTIMES } = require('../../src/adapters/xlsx/pivot-source.cjs');
  for (const l of [PIVOT_STATE_BRIDGE_RUNTIMES]) assert.deepEqual([...l], [...XLSX_RUNTIME_IDS]);

  const core = require('../../src/core/runtime-identity.cjs');
  for (const [coreKey, adapterModule, adapterKey, driverKey] of [
    ['XLSX_B03_BRIDGE_WRITE_RUNTIME_IDS', 'families/sheets.cjs', 'B03_BRIDGE_WRITE_RUNTIMES', 'B03_BRIDGE_WRITE_RUNTIMES'],
    ['XLSX_B10_BRIDGE_WRITE_RUNTIME_IDS', 'families/print.cjs', 'B10_BRIDGE_WRITE_RUNTIMES', 'B10_BRIDGE_WRITE_RUNTIMES'],
    ['XLSX_B05_BRIDGE_WRITE_RUNTIME_IDS', 'families/format.cjs', 'B05_BRIDGE_WRITE_RUNTIMES', 'B05_BRIDGE_WRITE_RUNTIMES'],
    ['XLSX_B22_BRIDGE_WRITE_RUNTIME_IDS', 'families/charts.cjs', 'B22_BRIDGE_WRITE_RUNTIMES', 'B22_BRIDGE_WRITE_RUNTIMES'],
    ['XLSX_B23_BRIDGE_WRITE_RUNTIME_IDS', 'families/charts.cjs', 'B23_BRIDGE_WRITE_RUNTIMES', 'B23_BRIDGE_WRITE_RUNTIMES'],
  ]) {
    assert.deepEqual([...core[coreKey]], [XLSX_RUNTIME_ID, XLSX_RUNTIME_17_ID], coreKey);
    const adapterRuntimeKey = coreKey.replace('XLSX_', 'XLSX_');
    assert.deepEqual([...require('../../src/adapters/xlsx/runtime.cjs')[adapterRuntimeKey]], [...core[coreKey]], `adapters/xlsx/runtime.cjs ${adapterRuntimeKey}`);
    assert.deepEqual([...require(`../../src/adapters/xlsx/${adapterModule}`)[adapterKey]], [...core[coreKey]], `adapters/xlsx/${adapterModule} ${adapterKey}`);
    assert.deepEqual([...require('../../src/drivers/eurooffice/xlsx-driver.cjs')[driverKey]], [...core[coreKey]], `xlsx-driver.cjs ${driverKey}`);
  }
  assert.deepEqual(core.XLSX_B03_BRIDGE_WRITE_RUNTIME_ENTRIES.map(core.runtimeIdOf), [...core.XLSX_B03_BRIDGE_WRITE_RUNTIME_IDS]);
  assert.deepEqual(core.XLSX_B10_BRIDGE_WRITE_RUNTIME_ENTRIES.map(core.runtimeIdOf), [...core.XLSX_B10_BRIDGE_WRITE_RUNTIME_IDS]);
  assert.deepEqual(core.XLSX_B05_BRIDGE_WRITE_RUNTIME_ENTRIES.map(core.runtimeIdOf), [...core.XLSX_B05_BRIDGE_WRITE_RUNTIME_IDS]);
  assert.deepEqual(core.XLSX_B22_BRIDGE_WRITE_RUNTIME_ENTRIES.map(core.runtimeIdOf), [...core.XLSX_B22_BRIDGE_WRITE_RUNTIME_IDS]);
  assert.deepEqual(core.XLSX_B23_BRIDGE_WRITE_RUNTIME_ENTRIES.map(core.runtimeIdOf), [...core.XLSX_B23_BRIDGE_WRITE_RUNTIME_IDS]);
});

test('ACCEPTED LIST (GH #326/#327: 9.3.5.0/16 + /17 xlsx, compatibility-first; docx 9.3.5.0/5 only since the GH #292 S4 D-state; no range/wildcard); every 9.3.4.60 revision, suffix, other DS / family -> FAIL_CLOSED', () => {
  assert.deepEqual([...XLSX_RUNTIME_IDS], ['9.3.5|euro-office-bridge-v1/9.3.5.0/16', XLSX_RUNTIME_17_ID]);
  // POSITIVE: /17 is now accepted too (GH #326/#327), the same pair/runtimeId shape /16 already proves above
  assert.deepEqual([classifyHandshake(handshake(apiWith(manifest({ patchVersion: 'euro-office-bridge-v1/9.3.5.0/17' }))), ACCEPTED_RUNTIMES.xlsx).ok, classifyHandshake(handshake(apiWith(manifest({ patchVersion: 'euro-office-bridge-v1/9.3.5.0/17' }))), ACCEPTED_RUNTIMES.xlsx).runtimeId], [true, XLSX_RUNTIME_17_ID]);
  // GH #292 S4 D-state (owner GO 2026-10-24): the word /3 and /4 compat entries are retired with their DS rollback points; /5 alone
  assert.deepEqual(ACCEPTED_RUNTIMES.docx.map(runtimeIdOf), [DOCX_RUNTIME_ID]);
  assert.deepEqual([...DOCX_RUNTIME_IDS], ACCEPTED_RUNTIMES.docx.map(runtimeIdOf));

  // POSITIVE: the one accepted pair, each family
  assert.deepEqual([classifyHandshake(handshake(apiWith(manifest({ patchVersion: 'euro-office-bridge-word-v1/9.3.5.0/5' }))), ACCEPTED_RUNTIMES.docx).ok, classifyHandshake(handshake(apiWith(manifest({ patchVersion: 'euro-office-bridge-word-v1/9.3.5.0/5' }))), ACCEPTED_RUNTIMES.docx).runtimeId], [true, DOCX_RUNTIME_ID]);
  // NEG (D-state): the retired word /3 and /4 refuse as a patch mismatch on the accepted 9.3.5
  for (const old of ['3', '4']) assert.deepEqual(classifyHandshake(handshake(apiWith(manifest({ patchVersion: `euro-office-bridge-word-v1/9.3.5.0/${old}` }))), ACCEPTED_RUNTIMES.docx), { ok: false, reason: 'PATCH_VERSION_MISMATCH' }, old);
  assert.deepEqual([classifyHandshake(handshake(apiWith(manifest())), ACCEPTED_RUNTIMES.xlsx).ok, classifyHandshake(handshake(apiWith(manifest())), ACCEPTED_RUNTIMES.xlsx).runtimeId], [true, XLSX_RUNTIME_ID]);

  // NEG: every retired 9.3.4.60 revision (both families) now refuses one check EARLIER (RUNTIME_VERSION_MISMATCH --
  // 9.3.4 itself has no accepted entry at all, not merely a patch mismatch on an otherwise-accepted version)
  for (const pv of ['euro-office-bridge-word-v1/9.3.4.60/1', 'euro-office-bridge-word-v1/9.3.4.60/2', 'euro-office-bridge-word-v1/9.3.4.60/3',
    'euro-office-bridge-v1/9.3.4.60/5', 'euro-office-bridge-v1/9.3.4.60/6', 'euro-office-bridge-v1/9.3.4.60/16'])
    assert.deepEqual(classifyHandshake(handshake(apiWith(manifest({ documentServerVersion: '9.3.4', patchVersion: pv }))), ACCEPTED_RUNTIMES.docx), { ok: false, reason: 'RUNTIME_VERSION_MISMATCH' }, pv);

  // NEG: documentServerVersion 9.3.5 (now accepted) with any OTHER patchVersion still refuses (PATCH_VERSION_MISMATCH:
  // the version itself has an accepted entry, but none of these pairs it correctly)
  for (const pv of ['euro-office-bridge-v1/9.3.5.0/15', 'euro-office-bridge-v1/9.3.5.0/18', 'euro-office-bridge-word-v1/9.3.5.0/2', 'euro-office-bridge-word-v1/9.3.5.0/4',
    'euro-office-bridge-v1/9.3.4.60/16', ' euro-office-bridge-v1/9.3.5.0/16', 'euro-office-bridge-v2/9.3.5.0/16'])
    assert.deepEqual(classifyHandshake(handshake(apiWith(manifest({ patchVersion: pv }))), ACCEPTED_RUNTIMES.xlsx), { ok: false, reason: 'PATCH_VERSION_MISMATCH' }, pv);

  // the single-runtime call form is unchanged (DOCX + every older caller): a 9.3.4.60 patch under RUNTIMES.xlsx refuses too
  assert.equal(classifyHandshake(handshake(apiWith(manifest({ documentServerVersion: '9.3.4', patchVersion: 'euro-office-bridge-v1/9.3.4.60/6' }))), RUNTIMES.xlsx).reason, 'RUNTIME_VERSION_MISMATCH');
});
