'use strict';

// Owner decision (B) (the owner, owner channel 2026-10-08, internal tracker): the DS runtime identity is the served
// bundle's own bridge manifest, read through ONE canonical handshake; the editor URL's '/9.3.4-<hash>/web-apps/'
// segment is an ephemeral DS cache key (a restart regenerated it on unchanged content) and is never a gate.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { RUNTIMES, ACCEPTED_RUNTIMES, XLSX_RUNTIME_ID, XLSX_RUNTIME_IDS, DOCX_RUNTIME_ID, DOCX_RUNTIME_IDS, runtimeIdOf, cacheKeyOf, HANDSHAKE_BODY, classifyHandshake } = require('../../src/core/runtime-identity.cjs');

const manifest = (over = {}) => ({
  bridgeVersion: '1.0.0', documentServerVersion: '9.3.4', documentServerBuild: '9.3.4-9475e9f12a1905e2ada1fbc4bb87d01b',
  patchVersion: 'euro-office-bridge-v1/9.3.4.60/5',
  capabilities: [{ id: 'range.validations.read/v1', version: 1, kind: 'READ', format: 'XLSX', mutation: 'NO', semanticContract: 'x' }],
  ...over,
});
// eslint-disable-next-line no-new-func
const handshake = Api => new Function('Api', HANDSHAKE_BODY)(Api);
const apiWith = m => ({ GetEuroOfficeBridge: () => ({ manifest: () => m, read: () => ({ ok: true, value: null }) }) });

test('the accepted runtimes: DS 9.3.4, cell bridge /3, word bridge /1 -- keyed by (version, patchVersion)', () => {
  assert.equal(XLSX_RUNTIME_ID, '9.3.4|euro-office-bridge-v1/9.3.4.60/5');
  assert.equal(DOCX_RUNTIME_ID, '9.3.4|euro-office-bridge-word-v1/9.3.4.60/1');
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
  const before = 'https://eo/9.3.4-9475e9f12a1905e2ada1fbc4bb87d01b/web-apps/apps/spreadsheeteditor/main/index.html';
  const after = 'https://eo/9.3.4-94bc577ecccd9583be1b130bd5a6b407/web-apps/apps/spreadsheeteditor/main/index.html';
  assert.notEqual(cacheKeyOf(before), cacheKeyOf(after));
  const a = classifyHandshake(handshake(apiWith(manifest())), RUNTIMES.xlsx);
  const b = classifyHandshake(handshake(apiWith(manifest({ documentServerBuild: cacheKeyOf(after) }))), RUNTIMES.xlsx);
  assert.equal(a.ok && b.ok, true);
  assert.equal(a.runtimeId, b.runtimeId);
});

test('cacheKeyOf: diagnostic only -- the key or null', () => {
  assert.equal(cacheKeyOf('https://eo/9.3.4-94bc577ecccd9583be1b130bd5a6b407/web-apps/x'), '9.3.4-94bc577ecccd9583be1b130bd5a6b407');
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
    [handshake(apiWith(manifest({ documentServerVersion: '9.3.5' }))), 'RUNTIME_VERSION_MISMATCH'],
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
  const word = manifest({ patchVersion: 'euro-office-bridge-word-v1/9.3.4.60/1' });
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

test('STATIC GUARD: every runtime-gated list in src is built from a core/runtime-identity.cjs RUNTIME_ID (XLSX: the explicit /5+/6 list)', () => {
  const root = path.join(__dirname, '../../src');
  const lists = [];
  const walk = d => { for (const e of fs.readdirSync(d, { withFileTypes: true })) { const p = path.join(d, e.name); if (e.isDirectory()) walk(p); else if (p.endsWith('.cjs')) fs.readFileSync(p, 'utf8').split('\n').forEach((l, i) => { if (/^const \w+_RUNTIMES = (Object\.freeze\(\[|DOCX_RUNTIME_IDS|DOCX_B18_BRIDGE_WRITE_RUNTIME_IDS|DOCX_B21_BRIDGE_READ_RUNTIME_IDS|XLSX_RUNTIME_IDS|XLSX_B05_BRIDGE_WRITE_RUNTIME_IDS|XLSX_B10_BRIDGE_WRITE_RUNTIME_IDS|XLSX_B03_BRIDGE_WRITE_RUNTIME_IDS)/.test(l)) lists.push([`${path.relative(root, p)}:${i + 1}`, l]); }); } };
  walk(root);
  // 11 before Wave B READ /4; it removed the driver's ALIGN_EXT / PRINT / PIVOT read gates (now Bridge capabilities);
  // the B15 /5 cutover removed EO_INTERNAL_SHEET_HYDRATION_RUNTIMES (the bridge hydration client gates itself)
  // M11.5a.2: the XLSX gates take XLSX_RUNTIME_IDS (the explicit [/5, /6] list, pinned below); M11.5d.1: DOCX takes DOCX_RUNTIME_IDS ([/1, /2])
  // M11.5e.4: B21 cut over from DOCX_RUNTIME_IDS to the narrower DOCX_B21_BRIDGE_READ_RUNTIME_IDS (word /3 only)
  assert.ok(lists.length >= 10, `found ${lists.length} runtime lists`);
  for (const [where, l] of lists) assert.match(l, /= (DOCX_RUNTIME_IDS|DOCX_B18_BRIDGE_WRITE_RUNTIME_IDS|DOCX_B21_BRIDGE_READ_RUNTIME_IDS|XLSX_RUNTIME_IDS|XLSX_B05_BRIDGE_WRITE_RUNTIME_IDS|XLSX_B10_BRIDGE_WRITE_RUNTIME_IDS|XLSX_B03_BRIDGE_WRITE_RUNTIME_IDS);$/, where);
  // M11.5a.4: the B05 Bridge WRITE lists (adapter + driver) are the ONLY ones on the /6-only list
  assert.deepEqual(lists.filter(([, l]) => /XLSX_B05_BRIDGE_WRITE_RUNTIME_IDS;$/.test(l)).map(([w]) => w.replace(/:\d+$/, '')).sort(), ['adapters/xlsx/families/format.cjs', 'drivers/eurooffice/xlsx-driver.cjs']);
  // M11.5b.4: the B10 Bridge WRITE lists (adapter + driver) are the ONLY ones on the /7-only list
  assert.deepEqual(lists.filter(([, l]) => /XLSX_B10_BRIDGE_WRITE_RUNTIME_IDS;$/.test(l)).map(([w]) => w.replace(/:\d+$/, '')).sort(), ['adapters/xlsx/families/print.cjs', 'drivers/eurooffice/xlsx-driver.cjs']);
  // M11.5c.4: the B03 Bridge WRITE lists (adapter + driver) are the ONLY ones on the /8-only list
  assert.deepEqual(lists.filter(([, l]) => /XLSX_B03_BRIDGE_WRITE_RUNTIME_IDS;$/.test(l)).map(([w]) => w.replace(/:\d+$/, '')).sort(), ['adapters/xlsx/families/sheets.cjs', 'drivers/eurooffice/xlsx-driver.cjs']);
});

test('the XLSX adapter\'s own runtime copy (it imports nothing outside its tree) equals the core identity', () => {
  assert.equal(require('../../src/adapters/xlsx/runtime.cjs').XLSX_RUNTIME_ID, XLSX_RUNTIME_ID);
  assert.deepEqual([...require('../../src/adapters/xlsx/runtime.cjs').XLSX_RUNTIME_IDS], [...XLSX_RUNTIME_IDS]);
  const { PIVOT_STATE_BRIDGE_RUNTIMES } = require('../../src/adapters/xlsx/pivot-source.cjs');
  for (const l of [PIVOT_STATE_BRIDGE_RUNTIMES]) assert.deepEqual([...l], [...XLSX_RUNTIME_IDS]);
  // M11.5c.4: B03-WRITE is /8-only (Bridge WRITE), the adapter copy == the core value == the driver list
  const c3 = require('../../src/core/runtime-identity.cjs');
  assert.deepEqual([...c3.XLSX_B03_BRIDGE_WRITE_RUNTIME_IDS], ['9.3.4|euro-office-bridge-v1/9.3.4.60/8']);
  assert.deepEqual([...require('../../src/adapters/xlsx/runtime.cjs').XLSX_B03_BRIDGE_WRITE_RUNTIME_IDS], [...c3.XLSX_B03_BRIDGE_WRITE_RUNTIME_IDS]);
  assert.deepEqual([...require('../../src/adapters/xlsx/families/sheets.cjs').B03_BRIDGE_WRITE_RUNTIMES], [...c3.XLSX_B03_BRIDGE_WRITE_RUNTIME_IDS]);
  assert.deepEqual([...require('../../src/drivers/eurooffice/xlsx-driver.cjs').B03_BRIDGE_WRITE_RUNTIMES], [...c3.XLSX_B03_BRIDGE_WRITE_RUNTIME_IDS]);
  assert.deepEqual(c3.XLSX_B03_BRIDGE_WRITE_RUNTIME_ENTRIES.map(c3.runtimeIdOf), [...c3.XLSX_B03_BRIDGE_WRITE_RUNTIME_IDS]);
  // M11.5b.4: B10-WRITE is /7-only (Bridge WRITE), the adapter copy == the core value == the driver list
  const c = require('../../src/core/runtime-identity.cjs');
  assert.deepEqual([...c.XLSX_B10_BRIDGE_WRITE_RUNTIME_IDS], ['9.3.4|euro-office-bridge-v1/9.3.4.60/7', '9.3.4|euro-office-bridge-v1/9.3.4.60/8']);
  assert.deepEqual([...require('../../src/adapters/xlsx/runtime.cjs').XLSX_B10_BRIDGE_WRITE_RUNTIME_IDS], [...c.XLSX_B10_BRIDGE_WRITE_RUNTIME_IDS]);
  assert.deepEqual([...require('../../src/adapters/xlsx/families/print.cjs').B10_BRIDGE_WRITE_RUNTIMES], [...c.XLSX_B10_BRIDGE_WRITE_RUNTIME_IDS]);
  assert.deepEqual([...require('../../src/drivers/eurooffice/xlsx-driver.cjs').B10_BRIDGE_WRITE_RUNTIMES], [...c.XLSX_B10_BRIDGE_WRITE_RUNTIME_IDS]);
  assert.deepEqual(c.XLSX_B10_BRIDGE_WRITE_RUNTIME_ENTRIES.map(c.runtimeIdOf), [...c.XLSX_B10_BRIDGE_WRITE_RUNTIME_IDS]);
  // M11.5a.4 (DS5_POLICY=B05_ONLY_FAIL_CLOSED): B05-WRITE alone is narrowed (/6 + /7 since M11.5b.1, never /5), the adapter copy == the core value
  const core = require('../../src/core/runtime-identity.cjs');
  assert.deepEqual([...core.XLSX_B05_BRIDGE_WRITE_RUNTIME_IDS], ['9.3.4|euro-office-bridge-v1/9.3.4.60/6', '9.3.4|euro-office-bridge-v1/9.3.4.60/7', '9.3.4|euro-office-bridge-v1/9.3.4.60/8']);
  assert.deepEqual([...require('../../src/adapters/xlsx/runtime.cjs').XLSX_B05_BRIDGE_WRITE_RUNTIME_IDS], [...core.XLSX_B05_BRIDGE_WRITE_RUNTIME_IDS]);
  assert.deepEqual([...require('../../src/adapters/xlsx/families/format.cjs').B05_BRIDGE_WRITE_RUNTIMES], [...core.XLSX_B05_BRIDGE_WRITE_RUNTIME_IDS]);
  assert.deepEqual([...require('../../src/drivers/eurooffice/xlsx-driver.cjs').B05_BRIDGE_WRITE_RUNTIMES], [...core.XLSX_B05_BRIDGE_WRITE_RUNTIME_IDS]);
  assert.deepEqual(core.XLSX_B05_BRIDGE_WRITE_RUNTIME_ENTRIES.map(core.runtimeIdOf), [...core.XLSX_B05_BRIDGE_WRITE_RUNTIME_IDS]);
});

test('M11.5a.2 / M11.5b.1 / M11.5c.1 ACCEPTED LIST: exactly /5, /6, /7 and /8 (explicit, no range/wildcard); /4, /9, suffixes, other DS / family -> FAIL_CLOSED', () => {
  assert.deepEqual([...XLSX_RUNTIME_IDS], ['9.3.4|euro-office-bridge-v1/9.3.4.60/5', '9.3.4|euro-office-bridge-v1/9.3.4.60/6', '9.3.4|euro-office-bridge-v1/9.3.4.60/7', '9.3.4|euro-office-bridge-v1/9.3.4.60/8']);
  // M11.5e.2: DOCX = the explicit Word list /1 + /2 + /3 (the primary /1 first)
  assert.deepEqual(ACCEPTED_RUNTIMES.docx.map(runtimeIdOf), [DOCX_RUNTIME_ID, '9.3.4|euro-office-bridge-word-v1/9.3.4.60/2', '9.3.4|euro-office-bridge-word-v1/9.3.4.60/3']);
  assert.deepEqual([...DOCX_RUNTIME_IDS], ACCEPTED_RUNTIMES.docx.map(runtimeIdOf));
  for (const n of ['1', '2', '3']) {
    const c = classifyHandshake(handshake(apiWith(manifest({ patchVersion: `euro-office-bridge-word-v1/9.3.4.60/${n}` }))), ACCEPTED_RUNTIMES.docx);
    assert.deepEqual([c.ok, c.runtimeId], [true, `9.3.4|euro-office-bridge-word-v1/9.3.4.60/${n}`], `word /${n}`);
  }
  for (const pv of ['euro-office-bridge-word-v1/9.3.4.60/0', 'euro-office-bridge-word-v1/9.3.4.60/4', 'euro-office-bridge-word-v1/9.3.4.60/3x', 'euro-office-bridge-word-v1/9.3.4.60/12', 'euro-office-bridge-v1/9.3.4.60/2', 'euro-office-bridge-v1/9.3.4.60/8'])
    assert.deepEqual(classifyHandshake(handshake(apiWith(manifest({ patchVersion: pv }))), ACCEPTED_RUNTIMES.docx), { ok: false, reason: 'PATCH_VERSION_MISMATCH' }, pv);
  const hs = pv => handshake(apiWith(manifest({ patchVersion: pv })));
  for (const n of ['5', '6', '7']) {
    const c = classifyHandshake(hs(`euro-office-bridge-v1/9.3.4.60/${n}`), ACCEPTED_RUNTIMES.xlsx);
    assert.deepEqual([c.ok, c.runtimeId], [true, `9.3.4|euro-office-bridge-v1/9.3.4.60/${n}`], `/${n}: the MATCHED entry's own id`);
  }
  for (const pv of ['euro-office-bridge-v1/9.3.4.60/4', 'euro-office-bridge-v1/9.3.4.60/9', 'euro-office-bridge-v1/9.3.4.60/7x', 'euro-office-bridge-v1/9.3.4.60/8x', 'euro-office-bridge-v1/9.3.4.60/1', 'euro-office-bridge-v1/9.3.4.60/6x',
    'euro-office-bridge-v1/9.3.4.60/56', 'euro-office-bridge-v1/9.3.4.60/', 'euro-office-bridge-v1/9.3.4.61/5', 'euro-office-bridge-v2/9.3.4.60/6', 'euro-office-bridge-word-v1/9.3.4.60/1', ' euro-office-bridge-v1/9.3.4.60/6'])
    assert.deepEqual(classifyHandshake(hs(pv), ACCEPTED_RUNTIMES.xlsx), { ok: false, reason: 'PATCH_VERSION_MISMATCH' }, pv);
  assert.deepEqual(classifyHandshake(handshake(apiWith(manifest({ documentServerVersion: '9.3.5', patchVersion: 'euro-office-bridge-v1/9.3.4.60/6' }))), ACCEPTED_RUNTIMES.xlsx), { ok: false, reason: 'RUNTIME_VERSION_MISMATCH' });
  // the single-runtime call form is unchanged (DOCX + every older caller)
  assert.equal(classifyHandshake(hs('euro-office-bridge-v1/9.3.4.60/6'), RUNTIMES.xlsx).reason, 'PATCH_VERSION_MISMATCH');
});
