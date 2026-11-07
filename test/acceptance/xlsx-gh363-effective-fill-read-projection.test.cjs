'use strict';

// GH #363 2nd unit (card an internal branch): offline checks
// for the read_effective_fill projection (src/adapters/xlsx/families/effective-fill-read.cjs) and its
// read-tools.cjs wiring -- no DS/MCP involved. The ENGINE side is simulated via `inv.compiledStyleReads`,
// the SAME shape the driver's own bridgeAxis()-backed compiledStyleReads axis returns (xlsx-driver.cjs).

const test = require('node:test');
const assert = require('node:assert/strict');

const { projectEffectiveFillRead, hexOfInt } = require('../../src/adapters/xlsx/families/effective-fill-read.cjs');
const { normalizeInventory } = require('../../src/adapters/xlsx/inventory.cjs');
const rt = require('../../src/server/read-tools.cjs');

const inv = (sheets, reads) => ({ sheets: sheets.map(name => ({ name })), compiledStyleReads: reads });

test('hexOfInt(): a measured fill colour -> #RRGGBB, masked to 24 bits, uppercase, zero-padded', () => {
  assert.equal(hexOfInt(16711680), '#FF0000');
  assert.equal(hexOfInt(0), '#000000');
  assert.equal(hexOfInt(255), '#0000FF');
  assert.equal(hexOfInt(null), null);
  assert.equal(hexOfInt(undefined), null);
  assert.equal(hexOfInt('16711680'), null, 'a string is never coerced -- only a real number is a usable value');
  assert.equal(hexOfInt(Number.NaN), null);
});

test('projectEffectiveFillRead(): zero\'s own EXACT LIVE shapes (docs/evidence/gh363-cf-effective-readback-live-20261105/internal-route-probe-v2-measured.json) -- a firing cell reads the rule\'s own hex fill, a non-firing cell reads null', () => {
  const a = { sheet: 'S', range: 'B2:B9' };
  const read = [
    { cell: 'B7', compiled: { fill: 16711680 } },
    { cell: 'B2', compiled: null },
  ];
  const out = projectEffectiveFillRead(inv(['S'], { 'S!B2:B9': read }), a);
  assert.deepEqual(out, { sheet: 'S', range: 'B2:B9', cells: [
    { cell: 'B7', fill: '#FF0000' },
    { cell: 'B2', fill: null },
  ] });
});

test('projectEffectiveFillRead(): a pre-/28 runtime degrades the WHOLE axis to the named CAPABILITY_MISSING string, never a per-cell guess', () => {
  const a = { sheet: 'S', range: 'B2:B9' };
  const out = projectEffectiveFillRead(inv(['S'], { 'S!B2:B9': 'UNOBSERVABLE:CAPABILITY_MISSING' }), a);
  assert.deepEqual(out, { sheet: 'S', range: 'B2:B9', fill: 'UNOBSERVABLE:CAPABILITY_MISSING' });
});

test('projectEffectiveFillRead(): not read / malformed-read / sheet-not-found are each their own named reason, never silently defaulted', () => {
  const a = { sheet: 'S', range: 'B2:B9' };
  assert.deepEqual(projectEffectiveFillRead(inv(['S'], undefined), a), { sheet: 'S', range: 'B2:B9', fill: 'UNOBSERVABLE:not-read' });
  assert.deepEqual(projectEffectiveFillRead(inv(['S'], {}), a), { sheet: 'S', range: 'B2:B9', fill: 'UNOBSERVABLE:not-read' });
  assert.deepEqual(projectEffectiveFillRead(inv(['S'], { 'S!B2:B9': { not: 'an array' } }), a), { sheet: 'S', range: 'B2:B9', fill: 'UNOBSERVABLE:malformed-read' });
  assert.deepEqual(projectEffectiveFillRead(inv(['OTHER'], {}), a), { sheet: 'S', range: 'B2:B9', error: 'SHEET_NOT_FOUND' });
});

test('projectEffectiveFillRead(): a per-cell read failure (bad address, threw) carries its OWN named reason, never folded into fill:null without explanation', () => {
  const a = { sheet: 'S', range: 'A1:A2' };
  const read = [
    { cell: 'A1', compiled: null, reason: 'bad-cell-address' },
    { cell: 'A2', compiled: null, reason: 'threw' },
  ];
  const out = projectEffectiveFillRead(inv(['S'], { 'S!A1:A2': read }), a);
  assert.deepEqual(out.cells, [
    { cell: 'A1', fill: null, reason: 'UNOBSERVABLE:bad-cell-address' },
    { cell: 'A2', fill: null, reason: 'UNOBSERVABLE:threw' },
  ]);
});

test('projectEffectiveFillRead(): a malformed per-cell entry is its own named reason, never crashes the whole read', () => {
  const a = { sheet: 'S', range: 'A1:A1' };
  const out = projectEffectiveFillRead(inv(['S'], { 'S!A1:A1': [null, 42, { cell: 'A1' }] }), a);
  assert.deepEqual(out.cells, [
    { cell: undefined, fill: null, reason: 'UNOBSERVABLE:malformed-entry' },
    { cell: undefined, fill: null, reason: 'UNOBSERVABLE:malformed-entry' },
    { cell: 'A1', fill: null },
  ]);
});

test('read-tools.cjs: read_effective_fill is ACCEPTED, requires sheet+range, is in the xlsx input schema', () => {
  const ok = rt.validateReadRequest('xlsx', { fileId: '1', operation: 'read_effective_fill', sheet: 'S', range: 'b2:b9' });
  assert.equal(ok.ok, true);
  assert.equal(ok.args.range, 'B2:B9');
  assert.equal(rt.validateReadRequest('xlsx', { fileId: '1', operation: 'read_effective_fill', sheet: 'S' }).code, 'INVALID_ARGS');
  assert.ok(rt.inputSchema('xlsx').properties.operation.enum.includes('read_effective_fill'));
  assert.equal(rt.OPERATIONS.xlsx.read_effective_fill.status, rt.STATUS.ACCEPTED);
});

test('read-tools.cjs: read_effective_fill is bounded to 200 cells, STRICTER than read_range\'s general MAX_READ_CELLS -- the LIMIT refusal names the operation', () => {
  const ok201 = rt.validateReadRequest('xlsx', { fileId: '1', operation: 'read_effective_fill', sheet: 'S', range: 'A1:J21' }); // 10*21 = 210 cells
  assert.equal(ok201.code, 'LIMIT');
  assert.match(ok201.reason, /read_effective_fill/);
  assert.match(ok201.reason, /200/);
  const ok200 = rt.validateReadRequest('xlsx', { fileId: '1', operation: 'read_effective_fill', sheet: 'S', range: 'A1:J20' }); // 10*20 = 200 cells
  assert.equal(ok200.ok, true, 'exactly 200 cells must be accepted, not refused');
  // the SAME 210-cell range is fine for a DIFFERENT operation (read_range's own, much larger bound) --
  // this check is read_effective_fill-specific, not a general range-size rule.
  const other = rt.validateReadRequest('xlsx', { fileId: '1', operation: 'read_range', sheet: 'S', range: 'A1:J21' });
  assert.equal(other.ok, true);
});

test('read-tools.cjs: xlsxRequest(\'read_effective_fill\', ...) expands the range into the EXACT cell list (cellsOf), carried alongside the original sheet+range for the projection\'s own lookup key', () => {
  const req = rt.xlsxRequest('read_effective_fill', { sheet: 'S', range: 'A1:B2' });
  assert.deepEqual(req, { windows: [], compiledStyleRanges: [{ sheet: 'S', range: 'A1:B2', cells: ['A1', 'B1', 'A2', 'B2'] }] });
});

test('read-tools.cjs: XLSX_PROJECTIONS wires read_effective_fill to projectEffectiveFillRead, not a copy', () => {
  assert.equal(rt.XLSX_PROJECTIONS.read_effective_fill, projectEffectiveFillRead);
});

// SEAM TEST (a maintainer's review, internal tracker, komment #ref): every test above hand-builds `inv` directly
// (the `inv()` helper above), bypassing src/adapters/xlsx/inventory.cjs's own normalizeInventory() --
// the REAL product path (read-tools.cjs readLive -> normalizeInventory -> projection). Before this fix,
// normalizeInventory() dropped payload.compiledStyleReads entirely (it was never added to the key list),
// so inv.compiledStyleReads was ALWAYS undefined on the real path -- read_effective_fill answered
// 'UNOBSERVABLE:not-read' on EVERY runtime, /28 included, regardless of whether the capability was
// present. The hand-built `inv()` tests above could not see this: they construct the post-normalization
// shape directly, skipping the seam that was broken.
test('SEAM: normalizeInventory() -> projectEffectiveFillRead -- the REAL product path, a raw driver payload through, not a hand-built inv', () => {
  const rawPayload = (compiledStyleReads) => ({ ok: true, sheets: [{ name: 'S', index: 0, used: 'A1:B2' }], cells: [], truncated: false, unsupported: [], ...(compiledStyleReads !== undefined ? { compiledStyleReads } : {}) });

  // (a) /28-shape: a measured pair, one cell fires (hex fill), one does not (null)
  const measured = normalizeInventory(rawPayload({ 'S!A1:A2': [{ cell: 'A1', compiled: { fill: 16711680 } }, { cell: 'A2', compiled: null }] }));
  assert.deepEqual(projectEffectiveFillRead(measured, { sheet: 'S', range: 'A1:A2' }), { sheet: 'S', range: 'A1:A2', cells: [{ cell: 'A1', fill: '#FF0000' }, { cell: 'A2', fill: null }] });

  // (b) pre-/28 runtime: the whole-axis degrade string, never a per-cell guess
  const degraded = normalizeInventory(rawPayload({ 'S!A1:A2': 'UNOBSERVABLE:CAPABILITY_MISSING' }));
  assert.deepEqual(projectEffectiveFillRead(degraded, { sheet: 'S', range: 'A1:A2' }), { sheet: 'S', range: 'A1:A2', fill: 'UNOBSERVABLE:CAPABILITY_MISSING' });

  // (c) not requested at all: the dedicated not-read reason, never silently defaulted or confused with (b)
  const notAsked = normalizeInventory(rawPayload(undefined));
  assert.deepEqual(projectEffectiveFillRead(notAsked, { sheet: 'S', range: 'A1:A2' }), { sheet: 'S', range: 'A1:A2', fill: 'UNOBSERVABLE:not-read' });
});

test('RED CONTROL: before the fix, normalizeInventory() dropped compiledStyleReads -- reproduced here on a LOCAL copy of the pre-fix normalization, not the real module', () => {
  // A minimal stand-in for the pre-fix normalizeInventory(): everything the seam test above needs,
  // EXCEPT compiledStyleReads is never copied from the payload -- the exact regression zero found.
  const preFixNormalize = payload => ({ sheets: payload.sheets, compiledStyleReads: undefined });
  const rawPayload = { ok: true, sheets: [{ name: 'S', index: 0, used: 'A1:B2' }], cells: [], truncated: false, unsupported: [], compiledStyleReads: { 'S!A1:A2': [{ cell: 'A1', compiled: { fill: 16711680 } }, { cell: 'A2', compiled: null }] } };
  const preFixInv = preFixNormalize(rawPayload);
  const got = projectEffectiveFillRead(preFixInv, { sheet: 'S', range: 'A1:A2' });
  assert.deepEqual(got, { sheet: 'S', range: 'A1:A2', fill: 'UNOBSERVABLE:not-read' }, 'the pre-fix shape answers not-read even though the payload DID carry a measured read -- this is the bug, reproduced');
  assert.notDeepEqual(got, projectEffectiveFillRead(normalizeInventory(rawPayload), { sheet: 'S', range: 'A1:A2' }), 'the REAL (fixed) normalizeInventory must answer differently on the SAME raw payload');
});
