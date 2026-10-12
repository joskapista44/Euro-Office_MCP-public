'use strict';

/**
 * Static gate: the product MCP exposes EXACTLY { product_status, xlsx_read, docx_read, xlsx_execute_task } (owner decisions
 * M11.3b.2 and M11.4b.1, owner channel). Nothing else may appear; `docx_execute_task` stays WITHHELD and no routed
 * write capability is released. xlsx_execute_task is served only while the WRITE release gate holds (M11.4a.1/2/3 CLOSED,
 * the exact six-intent allowlist, DOCX_WRITE_ACCEPTED empty). Changing the exposed set is a release decision: this gate
 * must be changed in the same diff, on purpose.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const { EXPOSED_CAPABILITY_TOOLS, TOOL_ROWS, exposedTools, capabilityInventory, writeReleaseGate, WRITE_RELEASE_PRECONDITIONS, ACCEPTED_XLSX_WRITE_ALLOWLIST } = require('../../src/server/tools.cjs');
const { buildServerTools } = require('../../src/server/server-tools.cjs');
const W = require('../../src/server/write-tools.cjs');

const EXPECTED_EXPOSED = Object.freeze(['product_status', 'xlsx_read', 'docx_read', 'xlsx_execute_task']);
const EXPECTED_WITHHELD = Object.freeze(['docx_execute_task']);
const EXPECTED_CAPABILITY_TOOLS = Object.freeze(['xlsx_read', 'docx_read', 'xlsx_execute_task']);
const SIX = Object.freeze(['write_range', 'insert_rows', 'delete_rows', 'rename_sheet', 'unmerge_range', 'set_chart_title']);

/** judge({ exposed, capabilityList, rows }) -> list of violations (empty = exact). Pure, so the NEG controls can drive it. */
function judge({ exposed, capabilityList, rows }) {
  const v = [];
  const extra = exposed.filter(t => !EXPECTED_EXPOSED.includes(t));
  const missing = EXPECTED_EXPOSED.filter(t => !exposed.includes(t));
  if (extra.length) v.push(`extra exposed: ${extra.join(',')}`);
  if (missing.length) v.push(`missing exposed: ${missing.join(',')}`);
  if (new Set(exposed).size !== exposed.length) v.push('duplicate exposed tool');
  for (const t of capabilityList) if (!EXPECTED_CAPABILITY_TOOLS.includes(t)) v.push(`capability list names ${t}`);
  for (const r of rows) {
    if (r.READ_WRITE !== 'WRITE') continue;
    const released = r.EXPOSE_IN_PRODUCT_MCP !== 'NO' || capabilityList.includes(r.TOOL);
    if (released && r.TOOL !== 'xlsx_execute_task') v.push(`WRITE tool released: ${r.TOOL}`);
  }
  return v;
}

const live = () => ({ exposed: exposedTools(), capabilityList: [...EXPOSED_CAPABILITY_TOOLS], rows: TOOL_ROWS });

test('EXACT_TOOLSET: exposedTools() = product_status, xlsx_read, docx_read, xlsx_execute_task -- nothing else', () => {
  assert.deepEqual(exposedTools(), [...EXPECTED_EXPOSED]);
  assert.deepEqual(judge(live()), []);
});

test('DOCX_WRITE_WITHHELD: docx_execute_task and every routed write capability are not exposed', () => {
  assert.deepEqual(TOOL_ROWS.filter(r => !exposedTools().includes(r.TOOL)).map(r => r.TOOL), [...EXPECTED_WITHHELD]);
  assert.ok(capabilityInventory().every(r => r.EXPOSE_IN_PRODUCT_MCP === 'NO'), 'no routed write capability is released');
  assert.deepEqual([...EXPOSED_CAPABILITY_TOOLS], [...EXPECTED_CAPABILITY_TOOLS]);
  assert.equal(Object.isFrozen(EXPOSED_CAPABILITY_TOOLS), true);
  assert.deepEqual([...W.DOCX_WRITE_ALLOWLIST], []);
});

test('WRITE_RELEASE_GATE: M11.4a.1/2/3 CLOSED, the handler allowlist is EXACTLY the accepted six, DOCX_WRITE_ACCEPTED empty', () => {
  assert.deepEqual({ ...WRITE_RELEASE_PRECONDITIONS }, { 'M11.4a.1': 'CLOSED', 'M11.4a.2': 'CLOSED', 'M11.4a.3': 'CLOSED' });
  assert.deepEqual([...ACCEPTED_XLSX_WRITE_ALLOWLIST], [...SIX]);
  assert.deepEqual([...W.XLSX_WRITE_ALLOWLIST].sort(), [...SIX].sort());
  assert.deepEqual(writeReleaseGate(), { ok: true, violations: [] });
});

test('the served registry is the exact set (what tools/list carries)', () => {
  const config = { schemaVersion: 2, nextcloud: { url: 'https://nc.example' }, broker: { url: 'http://127.0.0.1:9/api/euro-broker' } };
  assert.deepEqual(Object.keys(buildServerTools({ config, configIdentity: 'fixture' }).tools), [...EXPECTED_EXPOSED]);
});

test('NEG controls: an extra tool, a missing tool, a released DOCX WRITE tool or a write capability turns the gate red', () => {
  const base = live();
  assert.notDeepEqual(judge({ ...base, exposed: [...base.exposed, 'docx_execute_task'] }), []);
  assert.notDeepEqual(judge({ ...base, exposed: base.exposed.filter(t => t !== 'docx_read') }), []);
  assert.notDeepEqual(judge({ ...base, exposed: base.exposed.filter(t => t !== 'xlsx_execute_task') }), []);
  assert.notDeepEqual(judge({ ...base, capabilityList: [...base.capabilityList, 'docx_execute_task'] }), []);
  assert.notDeepEqual(judge({ ...base, rows: base.rows.map(r => (r.TOOL === 'docx_execute_task' ? { ...r, EXPOSE_IN_PRODUCT_MCP: 'YES' } : r)) }), []);
  assert.notDeepEqual(judge({ ...base, exposed: [...base.exposed, 'xlsx:write_range'] }), []);
});

test('NEG controls: the WRITE release gate fails closed on an open milestone, a widened / narrowed allowlist or a DOCX intent', () => {
  const docx = [];
  const bad = [
    { preconditions: { ...WRITE_RELEASE_PRECONDITIONS, 'M11.4a.3': 'ACTIVE' }, xlsxAllowlist: SIX, docxAllowlist: docx },
    { preconditions: { 'M11.4a.1': 'CLOSED', 'M11.4a.2': 'CLOSED' }, xlsxAllowlist: SIX, docxAllowlist: docx },
    { xlsxAllowlist: [...SIX, 'merge_range'], docxAllowlist: docx },
    { xlsxAllowlist: SIX.slice(1), docxAllowlist: docx },
    { xlsxAllowlist: [...SIX.slice(0, 5), 'insert_rows'], docxAllowlist: docx }, // six entries, one duplicate, one missing
    { xlsxAllowlist: SIX, docxAllowlist: ['insert_paragraph'] },
  ];
  for (const b of bad) assert.equal(writeReleaseGate(b).ok, false, JSON.stringify(b));
  assert.equal(writeReleaseGate({ xlsxAllowlist: [...SIX].reverse(), docxAllowlist: docx }).ok, true, 'order does not matter, membership does');
  // and the served surface obeys it: a failing gate withholds xlsx_execute_task, the READ release is untouched
  for (const b of bad) assert.deepEqual(exposedTools({ writeGate: b }), ['product_status', 'xlsx_read', 'docx_read'], JSON.stringify(b));
});

module.exports = { judge, EXPECTED_EXPOSED, EXPECTED_WITHHELD };
