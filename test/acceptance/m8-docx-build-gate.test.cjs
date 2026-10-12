'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { OPERATION_STATUS } = require('../../src/core/execution/result.cjs');
const { docxAdapter } = require('../../src/adapters/docx/adapter.cjs');
const driverModule = require('../../src/drivers/eurooffice/docx-driver.cjs');
const { createStore, createFakeDocxDriver } = require('../fakes/fake-docx-driver.cjs');

/**
 * M8 DOCX internal build gate (owner decision A, owner channel, 2026-10-05) -- NARROWED by the
 * W-B1920-C1 word-bridge cutover. Of the three original internal dependencies, only
 * CParagraph.RemoveNumPr (WRITE) still goes through this build-gate mechanism: BUILD VERIFIED (the
 * driver's per-dependency list) AND CAPABILITY VERIFIED (present, expected shape); on an unknown /
 * unparsable build the affected operation is refused BEFORE the write (0 dispatch, 0 write, 0
 * command, no default). The other two (list-level format READ, table grid READ) no longer use this
 * mechanism at all -- they are cut over to the Word bridge (docx.numbering.levelFormat.read/v1,
 * docx.table.effectiveGrid.read/v1), fail-closed at the driver's own read (null on any bridge
 * failure, no internal fallback); their family-level "unavailable" refusal is covered by
 * test/acceptance/docx-m5-gate78.test.cjs (script.levelFormatsUnreadable / script.gridUnreadable),
 * which this file no longer duplicates. The fake uses the REAL driver's gate (internalGateOf /
 * the B18 / B21 runtime gates); script.runtime simulates another runtime. M11.5d.4: B18 (the numbering removal) left this
 * internal mechanism -- it runs only through the word /2 Bridge WRITE, refused PREWRITE on every other runtime.
 */

// since owner decision (B) (internal tracker) the gate key is the served Word bridge manifest's runtime identity
const { DOCX_RUNTIME_ID: VERIFIED, DOCX_RUNTIME_IDS, cacheKeyOf } = require('../../src/core/runtime-identity.cjs');
const VERIFIED_V2 = '9.3.4|euro-office-bridge-word-v1/9.3.4.60/2'; // M11.5d.1: the compat Word /2
const VERIFIED_V3 = '9.3.4|euro-office-bridge-word-v1/9.3.4.60/3'; // M11.5e.2: the compat Word /3 (B21 Bridge READ)
const UNKNOWN = '9.3.4|euro-office-bridge-word-v1/9.3.4.60/0';
const PLAIN = { bold: false, italic: false, underline: false, strike: false, fontFamily: 'Arial', fontSize: 22, language: 'hu-HU' };
const A = 'Alpha item.';
const E1 = 'Existing one.';
const E2 = 'Existing two.';
const Z = 'Closing paragraph.';
const listDoc = () => [
  { text: 'Intro heading', level: 1, props: { ...PLAIN, bold: true } },
  { text: A, props: { ...PLAIN } },
  { text: E1, props: { ...PLAIN }, list: { type: 'bullet', level: 0, numId: 'e1' } },
  { text: E2, props: { ...PLAIN }, list: { type: 'bullet', level: 0, numId: 'e1' } },
  { text: Z, props: { ...PLAIN } },
];
const cell = text => ({ span: 1, paras: [{ text, props: { ...PLAIN } }] });
const grid = () => ({ rows: [0, 1, 2].map(r => [0, 1, 2].map(c => cell(`G r${r}c${c}`))), gridCalc: [14.586, 14.586, 18.9] });
const tableStore = () => createStore([{ text: 'Merge heading', level: 1, props: { ...PLAIN, bold: true } }], { doc: { tables: [grid()] } });

async function run(store, task, script = {}) {
  let t = 0;
  const driver = createFakeDocxDriver(store, script);
  const options = { driver, fileId: 'd', now: () => (t += 1), ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, docxAdapter, createExecutor(task));
  if (result.execution.operationAccounting) assert.equal(result.execution.operationAccounting.balanced, true);
  return { result, driver, m: result.metrics, x: result.execution };
}
const task = (...steps) => ({ id: 'm8-docx-build-gate', steps });
const reason = (x, i) => { const e = x.appliedOperations[i].evidence; return (e && e.evidence && e.evidence.reason) || (e && e.reason); };
const status = (x, i) => x.appliedOperations[i].status;

// M11.5d.4: B18 left the internal build gate -- the numbering removal runs ONLY through the word /2 Bridge WRITE; every
// other runtime (word /1 included) refuses it PREWRITE with the named reason, never a direct internal fallback.
const B18_REFUSAL = 'docx-numbering-remove-bridge-write-runtime-not-proven';
const USERS = [
  ['B18 numbering removal (Bridge WRITE)', () => createStore(listDoc()), [{ intent: 'remove_list', selector: { text: E1 } }], 0],
];

for (const [name, store, steps, at] of USERS) {
  test(`B18 runtime gate (POSITIVE): ${name} on word /2 -> APPLIED`, async () => {
    const { x } = await run(store(), task(...steps), { runtime: VERIFIED_V2 });
    assert.ok(x.appliedOperations.every(e => e.status === OPERATION_STATUS.APPLIED), JSON.stringify(x.appliedOperations.map(e => [e.status, e.evidence])));
  });

  for (const [label, build] of [['word /1 (no Bridge WRITE)', VERIFIED], ['unknown build', UNKNOWN], ['unparsable build (null)', null]]) {
    test(`B18 runtime gate (NEGATIVE): ${name} on ${label} -> refused PREWRITE, 0 dispatch / 0 write / 0 command for the step`, async () => {
      const s = store();
      const durableBefore = JSON.stringify(s.durable);
      const { x, driver, m } = await run(s, task(...steps), { runtime: build });
      assert.equal(status(x, at), OPERATION_STATUS.FAILED);
      assert.equal(reason(x, at), B18_REFUSAL);
      assert.equal(driver.calls.dispatch, at, 'only the steps before it dispatched');
      assert.equal(m.writes, at);
      assert.equal(m.commands, at);
      assert.equal(driver.calls.refusedInternal || 0, 0, 'the adapter refused first; the driver second line was not even reached');
      if (at === 0) assert.equal(JSON.stringify(s.durable), durableBefore, 'nothing written');
      else assert.notEqual(durableBefore, null);
    });
  }
}

test('B18 runtime gate (NEGATIVE): the removal via outdent at level 0 -- the Bridge WRITE alone not proven -> refused PREWRITE, 0 dispatch for the step', async () => {
  const steps = [{ intent: 'set_list_level', selector: { text: E2 }, level: 0 }, { intent: 'outdent_list_item', selector: { text: E2 } }];
  const ok = await run(createStore(listDoc()), task(...steps), { runtime: VERIFIED_V2 });
  assert.ok(ok.x.appliedOperations.every(e => [OPERATION_STATUS.APPLIED, OPERATION_STATUS.SATISFIED].includes(e.status)), JSON.stringify(ok.x.appliedOperations.map(e => [e.status, e.evidence])));
  const { x, driver } = await run(createStore(listDoc()), task(...steps), { runtime: VERIFIED_V2, gateOverride: { numberingRemoveBridge: false } });
  assert.equal(status(x, 1), OPERATION_STATUS.FAILED);
  assert.equal(reason(x, 1), B18_REFUSAL);
  assert.equal(driver.calls.refusedInternal || 0, 0);
  assert.ok(driver.calls.dispatch <= 1, 'no dispatch for the removal step');
});

test('B18 driver second line: a removal command reaching dispatch off word /2 is refused BEFORE anything is sent (no internal fallback)', async () => {
  const d = driverModule.createEuroOfficeDocxDriver({ url: 'http://unused.invalid', user: 'u', pass: 'p', loadPlaywright: () => ({ ok: false }) });
  const r = await d.dispatch(null, { kind: 'callCommand', body: 'return 1;', meta: { family: 'list', listAction: 'remove_list', bridgeNumberingRemove: true } });
  assert.deepEqual([r.ok, r.dispatched, r.outcome], [false, false, 'b18-bridge-write-runtime-not-proven']);
  for (const rt of [VERIFIED, UNKNOWN, null]) assert.equal(driverModule.internalWriteRefusal({ meta: { bridgeNumberingRemove: true } }, rt).outcome, 'b18-bridge-write-runtime-not-proven', String(rt));
  assert.equal(driverModule.internalWriteRefusal({ meta: { bridgeNumberingRemove: true } }, VERIFIED_V2), null, 'word /2: let through');
  assert.equal(driverModule.internalWriteRefusal({ meta: {} }, VERIFIED), null, 'a non-removal command is not touched');
});

// The list-level-format / table-grid "unavailable" family refusal is covered by
// test/acceptance/docx-m5-gate78.test.cjs (script.levelFormatsUnreadable / script.gridUnreadable) --
// those two no longer run through THIS build-gate mechanism, so not duplicated here.

test('build gate does NOT block DOCX operations that do not use the three dependencies (unknown build)', async () => {
  const listOk = await run(createStore(listDoc()), task(
    { intent: 'apply_list', selector: { text: A }, type: 'bullet', level: 0 },
    { intent: 'format_text', selector: { text: Z }, format: { bold: true } },
  ), { runtime: UNKNOWN });
  assert.ok(listOk.x.appliedOperations.every(e => e.status === OPERATION_STATUS.APPLIED), JSON.stringify(listOk.x.appliedOperations.map(e => [e.status, e.evidence])));
  // a vertical merge and its ROW-only split (1 column: no grid read), on the same unknown build
  const MERGE_V = { intent: 'merge_table_cells', selector: { tableCell: { text: 'G r0c0', tableAnchor: 'G r2c2' } }, endSelector: { tableCell: { text: 'G r1c0' } }, terminalMergedText: 'G r0c0 G r1c0', terminalMergedSpan: { rows: 2, cols: 1 } };
  const SPLIT_V = { intent: 'split_table_cell', selector: { tableCell: { text: 'G r0c0 G r1c0', tableAnchor: 'G r2c2' } }, splitRows: 2, splitCols: 1 };
  const tableOk = await run(tableStore(), task(MERGE_V, SPLIT_V), { runtime: UNKNOWN });
  assert.ok(tableOk.x.appliedOperations.length === 2 && tableOk.x.appliedOperations.every(e => e.status === OPERATION_STATUS.APPLIED), JSON.stringify([tableOk.x.stopDetail, tableOk.x.appliedOperations.map(e => [e.status, e.evidence])]));
});

test('driver second line: dispatch refuses a command naming an internal dependency on an unverified build, before anything is sent', async () => {
  const d = driverModule.createEuroOfficeDocxDriver({ url: 'http://unused.invalid', user: 'u', pass: 'p', loadPlaywright: () => ({ ok: false }) });
  for (const internal of [['CParagraph.RemoveNumPr'], ['Some.UnknownDependency']]) {
    const r = await d.dispatch(null, { kind: 'callCommand', body: 'return 1;', meta: { internal } });
    assert.deepEqual([r.ok, r.dispatched, r.outcome], [false, false, 'internal-docx-runtime-not-verified'], JSON.stringify(internal));
  }
});

test('driver gate: the gate key is the Word manifest runtime id (the URL cache key is diagnostic only); B18 = the word /2 /3 Bridge WRITE list; B21 = the word /3 Bridge READ list only', () => {
  assert.equal(cacheKeyOf('https://eurooffice.example.org/9.3.4-94bc577ecccd9583be1b130bd5a6b407/web-apps/apps/documenteditor/main/index.html?x=1'), '9.3.4-94bc577ecccd9583be1b130bd5a6b407');
  assert.equal(VERIFIED, '9.3.4|euro-office-bridge-word-v1/9.3.4.60/1');
  assert.deepEqual([...driverModule.B18_BRIDGE_WRITE_RUNTIMES], [VERIFIED_V2, VERIFIED_V3]);
  assert.ok(Object.isFrozen(driverModule.B18_BRIDGE_WRITE_RUNTIMES));
  assert.deepEqual([...DOCX_RUNTIME_IDS], [VERIFIED, VERIFIED_V2, VERIFIED_V3]);
  assert.equal(driverModule.internalGateOf(VERIFIED_V2).numberingRemoveBridge, true);
  assert.equal(driverModule.internalGateOf(VERIFIED_V3).numberingRemoveBridge, true, 'M11.5e.2: word /3 carries the B18 Bridge WRITE byte-identically');
  const ok = driverModule.internalGateOf(VERIFIED);
  assert.deepEqual(Object.keys(ok).sort(), ['numberingRemoveBridge', 'runtime', 'sectionVariant']);
  assert.equal(ok.numberingRemoveBridge, false, 'word /1 has no Bridge WRITE: the removal is refused there');
  assert.equal(driverModule.internalGateOf(UNKNOWN).numberingRemoveBridge, false);
  // M11.5e.4: B21 cut over to the word /3 Bridge READ ONLY -- narrower than B18's list (word /1, /2 have no capability)
  assert.deepEqual([...driverModule.B21_BRIDGE_READ_RUNTIMES], [VERIFIED_V3]);
  assert.ok(Object.isFrozen(driverModule.B21_BRIDGE_READ_RUNTIMES));
  assert.equal(ok.sectionVariant, false, 'word /1 has no Bridge READ: the variant re-check is refused there');
  assert.equal(driverModule.internalGateOf(VERIFIED_V2).sectionVariant, false, 'word /2 has no B21 Bridge READ either');
  assert.equal(driverModule.internalGateOf(VERIFIED_V3).sectionVariant, true);
  assert.equal(driverModule.internalGateOf(UNKNOWN).sectionVariant, false);
  assert.equal(driverModule.applyInternalGate, undefined, 'M11.5d.4: the RemoveNumPr presence read and its gate are gone');
});

test('driver gate: propertyQuery() gates the Word bridge reads on the stable Word runtime (no cache key), and the two removed internal reads are gone from it', () => {
  const { BRIDGE_CAPABILITIES } = require('../../src/drivers/eurooffice/ds-bridge.cjs');
  const { ACCEPTED_RUNTIMES } = require('../../src/core/runtime-identity.cjs');
  assert.equal(driverModule.PROPERTY_QUERY, driverModule.propertyQuery());
  assert.ok(driverModule.PROPERTY_QUERY.includes(JSON.stringify(ACCEPTED_RUNTIMES.docx)), 'the accepted Word runtimes (/1 + /2 + /3) are threaded into __eoBridgeRead');
  assert.equal(/9\.3\.4-[0-9a-f]{8,}/.test(driverModule.PROPERTY_QUERY), false, 'no cache key in the query');
  assert.doesNotMatch(driverModule.PROPERTY_QUERY, /\.GetLvl\(/);
  assert.doesNotMatch(driverModule.PROPERTY_QUERY, /GetFormat/);
  assert.doesNotMatch(driverModule.PROPERTY_QUERY, /TableGridCalc/);
  // NEGATIVE CONTROL: the GONE-regexes above must be able to match something, or they would pass vacuously.
  assert.match('inner.GetLvl(lk).GetFormat()', /\.GetLvl\(/);
  assert.match('inner.GetLvl(lk).GetFormat()', /GetFormat/);
  assert.match('t.Table.TableGridCalc', /TableGridCalc/);
  // M11.5d.4: the B18 presence read (el.Paragraph.RemoveNumPr typeof) is GONE from the read query (NEGATIVE CONTROL: the regex matches it)
  assert.doesNotMatch(driverModule.PROPERTY_QUERY, /RemoveNumPr/);
  assert.match('typeof el.Paragraph.RemoveNumPr', /RemoveNumPr/);
  assert.ok(driverModule.PROPERTY_QUERY.includes(BRIDGE_CAPABILITIES.DOCX_LIST_LEVEL_FORMAT_READ.id));
  assert.ok(driverModule.PROPERTY_QUERY.includes(BRIDGE_CAPABILITIES.DOCX_TABLE_GRID_READ.id));
  // M11.5e.4: the B21 direct internal getters are GONE from the read query; the Bridge READ capability id IS there
  assert.doesNotMatch(driverModule.PROPERTY_QUERY, /Get_TitlePage|IsEvenAndOdd/);
  assert.match('sec.Section.Get_TitlePage(); sec.Section.IsEvenAndOdd();', /Get_TitlePage|IsEvenAndOdd/);
  assert.ok(driverModule.PROPERTY_QUERY.includes(BRIDGE_CAPABILITIES.DOCX_SECTION_VARIANT_STATE_READ.id));
  assert.equal(driverModule.propertyQuery({ sectionVariantState: false }).includes(BRIDGE_CAPABILITIES.DOCX_SECTION_VARIANT_STATE_READ.id), false, 'opt-out/unverified: the capability id is not even named');
});
