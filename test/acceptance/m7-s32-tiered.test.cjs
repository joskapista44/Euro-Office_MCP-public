'use strict';

// M7 S32 slice 1 (owner GO 2026-10-03, owner channel): the tiered observation for write_range / clear_range above 5000
// used cells. Contract: docs/evidence/M7-S32-TIERED-CONTRACT.md. The editor bodies run here against an in-memory
// spreadsheet stand-in (shape only -- the LIVE harness proves the editor behaviour).

const test = require('node:test');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const { readQuery, usedProbeQuery, digestQuery, tierOf, tieredPayload, withLimitNote, SHA256_SRC, MAX_TIERED_CELLS } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { normalizeInventory } = require('../../src/adapters/xlsx/inventory.cjs');
const { liveScope } = require('../../src/adapters/xlsx/adapter.cjs');
const { parseRange, cellsOf, addr } = require('../../src/adapters/xlsx/a1.cjs');
const cells = require('../../src/adapters/xlsx/families/cells.cjs');
const clear = require('../../src/adapters/xlsx/families/clear.cjs');
const { installBridge, withPatch, BRIDGE_RUNTIME_ID } = require('../fakes/ds-bridge-runtime.cjs');

// ---- an in-memory workbook --------------------------------------------------------------------------------------

/** book: { sheets: [{ name, cells: Map(a -> { raw, bold?, nf?, fill? }) }] }; raw: number | text | '=formula'. */
function book(spec) {
  return { sheets: Object.entries(spec).map(([name, entries]) => ({ name, cells: new Map(Object.entries(entries).map(([a, raw]) => [a, { raw, bold: false, nf: 'General', fill: null }])) })) };
}
const clone = b => ({ sheets: b.sheets.map(s => ({ name: s.name, cells: new Map([...s.cells].map(([a, c]) => [a, { ...c }])) })) });
const set = (b, sheet, a, raw) => { const s = b.sheets.find(x => x.name === sheet); const c = s.cells.get(a) || { bold: false, nf: 'General', fill: null }; s.cells.set(a, { ...c, raw }); return b; };
const empty = raw => raw === undefined || raw === '';

function usedOf(s) {
  const gs = [...s.cells].filter(([, c]) => !empty(c.raw)).map(([a]) => parseRange(a));
  if (!gs.length) return null;
  return `${addr(Math.min(...gs.map(g => g.c1)), Math.min(...gs.map(g => g.r1)))}:${addr(Math.max(...gs.map(g => g.c1)), Math.max(...gs.map(g => g.r1)))}`;
}

/** The measured getter shapes: GetValue is text (numbers too), GetFormula returns the value text of a non-formula cell. */
function apiOf(b) {
  const cellOf = (s, a) => {
    const c = s.cells.get(a) || { raw: undefined, bold: false, nf: 'General', fill: null };
    const f = typeof c.raw === 'string' && c.raw.startsWith('=') ? c.raw : '';
    const ref = f && /^=([A-Z]+\d+)$/.exec(f);
    const refRaw = ref ? (s.cells.get(ref[1]) || {}).raw : undefined;
    const v = f ? (ref ? (empty(refRaw) ? '0' : String(refRaw)) : '7') : empty(c.raw) ? '' : String(c.raw);
    return { GetAddress: () => a, GetValue: () => v, GetFormula: () => f || v, MergeArea: { GetAddress: () => a }, _num: typeof c.raw === 'number' || !!f, _count: c.count,
      GetNumberFormat: () => c.nf, GetFillColor: () => (c.fill === null ? null : { GetRGB: () => c.fill }), range: { getFont: () => ({ getBold: () => c.bold }) } };
  };
  // `range` = the first cell's model range: the DS Bridge reads a single-cell GetRange(a).range (B01 / B02 cutover)
  const rangeOf = (s, a1) => ({ GetAddress: () => a1, range: cellOf(s, cellsOf(parseRange(a1))[0]).range, ForEach: fn => { for (const a of cellsOf(parseRange(a1))) fn(cellOf(s, a)); },
    SetValue: v => { for (const a of cellsOf(parseRange(a1))) { const c = s.cells.get(a) || { bold: false, nf: 'General', fill: null }; s.cells.set(a, { ...c, raw: v }); } } });
  const ws = s => ({ GetName: () => s.name, GetIndex: () => b.sheets.indexOf(s), GetUsedRange: () => (usedOf(s) ? rangeOf(s, usedOf(s)) : null), GetRange: a1 => rangeOf(s, a1) });
  const api = { GetSheets: () => b.sheets.map(ws), GetSheet: n => { const s = b.sheets.find(x => x.name === n); return s ? ws(s) : null; }, WorksheetFunction: { COUNT: cell => { if (cell._count === 'THROW') throw new Error('COUNT failed'); return cell._count !== undefined ? cell._count.value : (cell._num ? 1 : 0); } } };
  installBridge(api); // the REAL served DS Bridge block (bold / italic go through it since the B01 / B02 cutover)
  return api;
}
const run = (body, b) => new Function('Api', body)(apiOf(b));

/** What the driver's readLive returns for `request` on `b`: the probe, the tier decision, then the matching read. */
function observe(b, request) {
  const probe = run(usedProbeQuery(), b);
  const tier = tierOf(probe.sheets, request);
  const opts = { formatSheets: request.formatSheets || [], objectSheets: request.objectSheets || [] };
  if (tier.tier === 'tiered') {
    const payload = tieredPayload(run(readQuery(request.windows, { ...opts, usedRanges: false }), b), run(digestQuery(request.windows), b), tier, { identity: BRIDGE_RUNTIME_ID, cacheKey: null, fileId: null });
    return { tier, primary: normalizeInventory(payload), payload };
  }
  const p = run(readQuery(request.windows, opts), b);
  return { tier, primary: normalizeInventory(p), payload: p };
}

/** A 1000-row x 10-column sheet (10 000 cells): text, numbers and a formula column. */
function bigBook() {
  const spec = {};
  for (let r = 1; r <= 1000; r += 1) {
    Object.assign(spec, { [`A${r}`]: `T${r}`, [`B${r}`]: r, [`C${r}`]: r + 0.5, [`D${r}`]: `D${r}`, [`E${r}`]: `=B${r}+C${r}`, [`F${r}`]: `F${r}`, [`G${r}`]: 7 * r, [`H${r}`]: `H${r}`, [`I${r}`]: 3 * r, [`J${r}`]: `J${r}` });
  }
  return book({ S: spec, Other: { A1: 'x', B2: 5 } });
}

const rowsOf = (from, to, fn) => Array.from({ length: to - from + 1 }, (_, i) => fn(from + i));
const writeTask = (...ranges) => ({ id: 't', steps: ranges.map(([id, range, mark]) => { const g = parseRange(range); return { intent: 'write_range', id, sheet: 'S', range, values: rowsOf(g.r1, g.r2, r => Array.from({ length: g.cols }, (_, k) => `${mark}${r}x${k}`)) }; }) });
const opOf = (task, i) => require('../../src/adapters/xlsx/adapter.cjs').planOperations(task).operations[i];
/** Applies a planned write_range op to the book (what the editor would do). */
const apply = (b, op) => { for (const c of op.cells) set(b, op.selector.sheet, c.a, c.formula !== undefined ? c.formula : c.v); return b; };
const l3 = (op, before, after) => cells.writePreservation(op, { primary: before.primary }, { primary: after.primary });

// ---- 1. below 5000: unchanged ------------------------------------------------------------------------------------

test('readQuery is byte-identical for the existing query shapes (sha256, recorded here) -- baseline re-recorded at the DS Bridge Wave A cutover', () => {
  const SHAPES = {
    empty: [[], {}],
    windows: [[{ sheet: 'Data', range: 'A1:B2' }, { sheet: 'S32P', range: 'A1:J100', format: true }], {}],
    fmtBold: [[{ sheet: 'Data', range: 'A1:B2', format: true }], {}],
    style: [[{ sheet: 'Data', range: 'A1', style: true }], { styleSheets: ['Data'], formatSheets: ['Data'] }],
    deps: [[], { objectSheets: ['*'], conditional: true }],
    layout: [[], { layout: [{ sheet: 'Data', maxCol: 3, maxRow: 2 }] }],
    panes: [[], { panes: true }],
    rules: [[], { validationRanges: [{ sheet: 'D', range: 'A1' }], cfRanges: [{ sheet: 'D', range: 'A1' }], filterSheets: ['*'] }],
    objects: [[], { chartSheets: ['*'], pageSheets: ['*'], pagePerSheet: true, printSheets: ['*'], colBreaks: true, drawingSheets: ['*'] }],
  };
  // Re-recorded at the DS Bridge Wave A cutover (owner GO, internal tracker point 8; LIVE parity 9/9 #ref). The
  // a65cd4b baseline was kept through a chain of carve-outs (S4 edge, X-F05, X-F07a, print scale, shape text, column
  // canon, X-O04/05, B01/B02 font). The cutover moves the S4 edge block itself into the bridge and removes five per-shape
  // gate flags (internalBold / internalAppearance / internalPanes / internalFilter / internalColBreaks), so that chain
  // cannot be carried any further. Instead the cutover diff was proven shape by shape against the pre-cutover driver
  // (1b002d8): every changed line is inside a migrated block (B04 align/border + S4 edge, B05 gate re-nesting, B06 pane,
  // B07 filter field, B12 column breaks, B14/B16 drawing anchor/text, B17 width canon, the bridgeAxis helper) --
  // docs/evidence/ds-bridge-wave-a-cutover-readquery-diff-20261006.txt. From here on these hashes are the baseline.
  // Re-recorded once more at the PR #39 merge check (2026-10-07): the recorded set predated a one-line COMMENT rewording
  // inside the query ("NOT gated by internalAppearance" -> "NOT read through the DS Bridge"); reverting that single line
  // reproduces all nine previous hashes, so the query is otherwise byte-identical.
  // Re-recorded at the runtime-identity change (owner decision (B), owner channel 2026-10-08, internal tracker): the
  // bridge client now takes the expected (documentServerVersion, patchVersion) instead of a URL cache key, checks the
  // manifest shape / capability id / capability version separately, and the align-ext / print reasons say runtime.
  // Proven shape by shape against ceece069 (whose hashes reproduce the set above): every changed line is in the bridge
  // client, the embedded expected runtime, the reason rename or one comment --
  // docs/evidence/ds-runtime-identity-readquery-diff-20261008.txt.
  // Re-recorded at the Wave B /4 runtime bump (2026-10-08, design docs/DS-BRIDGE-WAVE-B-READ-DESIGN-20261008.md): the ONLY
  // changed lines are the embedded expected patchVersion /3 -> /4 in the two bridge-client calls --
  // docs/evidence/wave-b-runtime-v4-readquery-diff-20261008.txt.
  // Re-recorded at the Wave B READ driver cutover (B05 READ / B09 / B13 -> range.alignmentExt / sheet.printSettings /
  // pivot.state .read/v1, patch family /4): every changed hunk is inside a migrated block (the B05 block of addStyle, the
  // pivot reader -- now generated in every shape, no internalPivot flag -- and its map, the print-read comment and body,
  // breaksOfModel gone into the bridge) -- docs/evidence/wave-b-read-cutover-readquery-diff-20261008.txt.
  // Re-recorded at the B15 /5 runtime bump (2026-10-09, owner GO owner channel): the ONLY changed lines are the
  // embedded expected patchVersion /4 -> /5 in the two bridge-client calls (2 per shape, all nine shapes) --
  // docs/evidence/b15-runtime-v5-readquery-diff-20261009.txt.
  // Re-recorded at the M11.5a.2 dual-runtime compatibility (owner GO owner channel): the ONLY changed lines are
  // the bridge client's two runtime compares (one accepted runtime -> an explicit list, exact match) and the embedded
  // expected runtime /5 -> the explicit list [/5, /6] in the two client calls (-4 +5 lines per shape, all nine shapes; the
  // previous nine hashes reproduce from 1d31d8d) -- docs/evidence/m11-5a2-dual-runtime-readquery-diff-20261010.txt.
  // Re-recorded at the M11.5b.1 /7 compatibility (M11.5 continuous mandate, internal tracker): the ONLY changed
  // lines are the embedded accepted list [/5, /6] -> [/5, /6, /7] in the two client calls (-2 +2 per shape, all nine;
  // the previous nine hashes reproduce from e7e5b44) -- docs/evidence/m11-5b1-v7-compat-readquery-diff-20261011.txt.
  // Re-recorded at the M11.5c.1 /8 compatibility (same mandate): the ONLY changed lines are the embedded accepted list
  // [/5, /6, /7] -> [/5, /6, /7, /8] in the two client calls (-2 +2 per shape, all nine; the previous nine hashes
  // reproduce from c21d104) -- docs/evidence/m11-5c1-v8-compat-readquery-diff-20261011.txt.
  // Re-recorded at the M12.2 public-release comment hygiene (owner GO M12, internal tracker): internal references were
  // removed from product comments, and some of those comments sit inside the generated query text. Every changed line
  // of the nine generated queries is a comment line (54 lines, 0 code lines; the previous nine hashes reproduce from
  // 1d724eb) -- docs/evidence/m12-2-comment-hygiene-readquery-diff-20261012.txt.
  const AT_CUTOVER = {
    empty: '5b6090c13b5e8bace4193d67ea4802935537f31d6001d726db580bf9ea57cb78', windows: '80189bd0733f45c12e26f40f4cf6ddb283e7691e2337f992a0e5a9cafee84008', fmtBold: '1fd9c2af554ba01d704f155e466fba040d2c47e5ed8d2adf066998fefc6b704f', style: 'cad4789e54ab244e176ff7989f81a080835d98cb8ab675be1395f57e95c8dc2a', deps: '5b0dc294d49838f9e2e0ab4c0b4dd203de89781bcd9e42cef5bd368263a5979e', layout: '7bae2af9a481b8251f7bc98712964a640e8041a5ae05f37348bcb9b7c6ac60a3', panes: '3ccd17f76113b4e04e437af6a884757f7b489fddcc5d5bd846ec8621ead53a99', rules: '1bd220e1a593a909b605f1a037c9d25416e9af6e33c5c949971c1581d28c0a41', objects: '20cbd40038f67060a22da905d692607c23c04df8715bb05ee3f1d3575afb4b6a'
  };
  // the migrated direct internals are gone from EVERY shape (no fallback can hide in a shape that does not ask for it)
  // (call / access forms only: the driver's own comments still name what the bridge replaced)
  const GONE = /readInternalFontFlag|\.getBold\(|\.getItalic\(|\.getBorderFull\(|\.getAlignHorizontal\(|\.ColId\b|\.topLeftFrozenCell|breaksOfModel|\.asc_getDataRef|\.getReadingOrder\(|\.getShrinkToFit\(|PagePrintOptions|\.getDocContent\(|\.drawingBase\b|\.charCountToModelColWidth\(/;
  for (const [k, [w, o]] of Object.entries(SHAPES)) {
    const q = readQuery(w, o);
    assert.doesNotMatch(q, GONE, `${k}: no migrated direct internal`);
    assert.equal(crypto.createHash('sha256').update(q).digest('hex'), AT_CUTOVER[k], k);
  }
});

test('a <= 5000-cell request is the FULL tier and returns the full read (every used-range cell, no digest)', () => {
  const b = book({ S: Object.fromEntries(rowsOf(1, 400, r => [`A${r}`, `T${r}`])) });
  const task = writeTask(['w1', 'B1:B10', 'n']);
  const o = observe(b, liveScope(task));
  assert.equal(o.tier.tier, 'full');
  assert.equal(o.primary.tier, 'full');
  assert.equal(o.primary.digests, null);
  assert.equal(Object.keys(o.primary.cells).length, 410, 'the used range (400) + the window (10)');
});

// ---- 2. the tier decision ------------------------------------------------------------------------------------------

test('tier decision: tiered only for a marked write/clear task; every other case refused with its reason', () => {
  const big = [{ name: 'S', used: 'A1:J5000' }];
  const w = [{ sheet: 'S', range: 'A1:J100' }];
  assert.equal(tierOf(big, { windows: w, tiered: true }).tier, 'tiered');
  assert.equal(tierOf([{ name: 'S', used: 'A1:J400' }], { windows: w, tiered: true }).tier, 'full');
  const why = req => tierOf(big, req).reason;
  assert.equal(why({ windows: w }), 'family-not-migrated-to-the-tiered-observation');
  // slice 2: the dependency readers (with the T3 formula index) and an AutoFilter read are allowed in a tiered request
  assert.equal(tierOf(big, { windows: w, tiered: true, dependents: true }).tier, 'tiered');
  assert.equal(tierOf(big, { windows: w, tiered: true, dependents: true, filterSheets: ['S'] }).tier, 'tiered');
  assert.equal(why({ windows: w, tiered: true, dependents: true, objectSheets: ['S'] }), 'tiered-observation-reads-windows-only: objectSheets');
  assert.equal(why({ windows: w, tiered: true, styleSheets: ['S'] }), 'tiered-observation-reads-windows-only: styleSheets');
  assert.equal(why({ windows: [{ sheet: 'S', range: 'A1', style: true }], tiered: true }), 'tiered-observation-reads-windows-only: window-style');
  assert.equal(why({ windows: rowsOf(0, 5, i => ({ sheet: 'S', range: `A${i * 100 + 1}:J${i * 100 + 100}` })), tiered: true }), 'task-windows-over-5000-cells');
  assert.equal(tierOf([{ name: 'S', used: 'A1:J5001' }], { windows: w, tiered: true }).reason, `workbook-over-${MAX_TIERED_CELLS}-cells`);
});

test('the bound is stable for a task: a far write that GROWS the used range keeps the same tier before and after', () => {
  const b = book({ S: { A1: 'a', B2: 2 } });
  const task = writeTask(['w1', 'Z1000', 'n']);
  const req = liveScope(task);
  const before = observe(b, req);
  const after = observe(apply(clone(b), opOf(task, 0)), req);
  assert.deepEqual([before.tier.tier, after.tier.tier], ['tiered', 'tiered']);
  assert.equal(before.tier.bound, after.tier.bound);
  assert.equal(l3(opOf(task, 0), before, after).ok, true, 'used-range growth by the own target is not a change outside');
});

test('above 5000 a non-eligible request is the old full read (truncated) plus an explicit IMPLEMENTATION_LIMIT entry', () => {
  const b = bigBook();
  const task = { id: 't', steps: [{ intent: 'merge_range', id: 'm1', sheet: 'S', range: 'L1:M1' }] };
  const o = observe(b, liveScope(task));
  assert.equal(o.tier.tier, 'refused');
  assert.equal(o.primary.complete, false);
  assert.equal(o.primary.truncated, true);
  // readLive appends the note with this very function
  const note = withLimitNote([], o.tier)[0].reason;
  assert.match(note, /^IMPLEMENTATION_LIMIT:family-not-migrated-to-the-tiered-observation \(bound \d+\)$/);
});

// ---- 3. the digest and SHA-256 -------------------------------------------------------------------------------------

test('the in-editor SHA-256 equals Node crypto (ASCII, accents, astral, empty, > one block)', () => {
  const sha = new Function(`${SHA256_SRC}; return __sha256;`)();
  for (const s of ['', 'abc', 'árvíztűrő tükörfúrógép', '𝔘𝔫𝔦😀', 'x'.repeat(1000), JSON.stringify([['A1', 'text', 'T1', '']])]) {
    assert.equal(sha(s), crypto.createHash('sha256').update(s, 'utf8').digest('hex'), JSON.stringify(s).slice(0, 30));
  }
});

test('digest: non-blank cells outside the windows only; a formula keeps its text, not its value; the hash is SHA-256 of the canonical', () => {
  const b = book({ S: { A1: 'in', B1: 1, A2: 'out', B2: 2, C2: '=A2', D2: '' }, Other: {} });
  const g = run(digestQuery([{ sheet: 'S', range: 'A1:B1' }], { canonical: true }), b);
  assert.equal(g.ok, true);
  const s = g.digests.find(x => x.sheet === 'S');
  assert.deepEqual(JSON.parse(s.canonical), [['A2', 'text', 'out', ''], ['B2', 'number', '2', ''], ['C2', 'formula', '', '=A2']]);
  assert.equal(s.hash, crypto.createHash('sha256').update(s.canonical, 'utf8').digest('hex'));
  assert.deepEqual(s.excluded, ['A1:B1']);
  assert.deepEqual(g.digests.find(x => x.sheet === 'Other'), { sheet: 'Other', used: null, excluded: [], count: 0, hash: crypto.createHash('sha256').update('[]').digest('hex'), canonical: '[]' });
});

test('the probe and the digest bodies call nothing but getters, ForEach and COUNT', () => {
  const MUTATING = /\.(Add\w*|Delete|Remove\w*|Insert\w*|Set\w*|Merge|UnMerge|Split\w*|Paste\w*|Replace\w*|Move\w*|Apply\w*|Clear\w*|Copy\w*|AutoFit|Sort|SetValue|ShowAllData)\s*\(/;
  for (const q of [usedProbeQuery(), digestQuery([{ sheet: "x'); Api.GetSheet('S').GetRange('A1').SetValue(1); ('", range: 'A1' }])]) {
    const code = q.replace(/const rects = \[[\s\S]*?\];/, 'const rects = [];');
    assert.doesNotMatch(code, MUTATING);
    for (const c of [...code.matchAll(/\.([A-Z]\w*)\s*\(/g)].map(m => m[1])) assert.match(c, /^(Get\w*|ForEach|COUNT)$/, c);
  }
});

// ---- 4. per-step L3 on the tiered observation -----------------------------------------------------------------------

/** Two-step task on the 10 000-cell book: step 1 writes L1:L10, step 2 writes L20:L29 (both outside the used range). */
function scenario() {
  const b = bigBook();
  const task = writeTask(['w1', 'L1:L10', 'a'], ['w2', 'L20:L29', 'b']);
  const req = liveScope(task);
  const before = observe(b, req);
  const after = apply(clone(b), opOf(task, 0));
  return { b, task, req, before, after, op: opOf(task, 0) };
}

test('tiered PASS: step 1 changed only its own target; every other window and the rest unchanged', () => {
  const s = scenario();
  assert.equal(s.before.tier.tier, 'tiered');
  const r = l3(s.op, s.before, observe(s.after, s.req));
  assert.equal(r.ok, true, JSON.stringify(r.evidence));
  assert.equal(r.evidence.tier, 'tiered');
  assert.equal(r.evidence.compared, 10, 'the OTHER window (step 2) is compared cell by cell');
});

test('OWNER POINT 1 -- negative control: step 1 corrupts a LATER step\'s target -> step 1\'s L3 FAILS (cell by cell)', () => {
  const s = scenario();
  set(s.after, 'S', 'L25', 'corrupted');
  const r = l3(s.op, s.before, observe(s.after, s.req));
  assert.equal(r.ok, false);
  assert.deepEqual(r.evidence.changed.map(c => c.cell), ['S!L25']);
});

const OUTSIDE = {
  'value changed': a => set(a, 'S', 'B500', 999),
  'number -> text with the same value': a => set(a, 'S', 'B500', '500'),
  'formula text changed, same value': a => set(a, 'S', 'E500', '=C500+B500'),
  'two cells swapped': a => { set(a, 'S', 'A500', 'T501'); set(a, 'S', 'A501', 'T500'); },
  'a cell appeared': a => set(a, 'S', 'K700', 'new'),
  'a cell disappeared': a => set(a, 'S', 'D300', ''),
  'a change on another sheet': a => set(a, 'Other', 'A1', 'y'),
};
for (const [name, mutate] of Object.entries(OUTSIDE)) {
  test(`T2 mutant outside the windows -> L3 FAIL: ${name}`, () => {
    const s = scenario();
    mutate(s.after);
    const r = l3(s.op, s.before, observe(s.after, s.req));
    assert.equal(r.ok, false, name);
    assert.equal(r.evidence.digest.changed.length, 1);
    assert.equal(r.evidence.digest.changed[0].change, 'content-outside-windows');
  });
}

test('RED without the digest: the same outside mutant is INVISIBLE to the cell-level L3 on a window-only read', () => {
  const s = scenario();
  set(s.after, 'S', 'B500', 999);
  const strip = o => ({ primary: { ...o.primary, tier: 'full', digests: null } });
  const r = cells.verifyPreservation(s.op, strip(s.before), strip(observe(s.after, s.req)));
  assert.equal(r.ok, true, 'the digest is load-bearing: the window cells alone pass');
});

test('a misaligned exclusion window, a missing digest, a tier change: all fail closed', () => {
  const s = scenario();
  const a = observe(s.after, s.req);
  const shifted = { ...s.req, windows: s.req.windows.map((w, i) => (i === 1 ? { ...w, range: 'L21:L30' } : w)) };
  const mis = l3(s.op, s.before, { primary: { ...a.primary, digests: observe(s.after, shifted).primary.digests } });
  assert.equal(mis.ok, false);
  assert.equal(mis.evidence.digest.changed[0].change, 'excluded-windows-differ');
  const noDigest = normalizeInventory({ ...a.payload, digests: a.payload.digests.filter(d => d.sheet !== 'Other') });
  assert.equal(noDigest.complete, false);
  assert.equal(l3(s.op, s.before, { primary: noDigest }).evidence.reason, 'xlsx-preservation-observation-incomplete');
  const full = { primary: { ...a.primary, tier: 'full', digests: null } };
  assert.equal(l3(s.op, s.before, full).evidence.reason, 'xlsx-preservation-tier-changed');
});

test('used-range SHRINK by a clear of the edge rows is not a change outside the windows', () => {
  const b = bigBook();
  const task = { id: 't', steps: [{ intent: 'clear_range', id: 'c1', sheet: 'S', range: 'A991:J1000' }] };
  const req = liveScope(task);
  const before = observe(b, req);
  const after = clone(b);
  for (const a of cellsOf(parseRange('A991:J1000'))) set(after, 'S', a, '');
  const op = require('../../src/adapters/xlsx/adapter.cjs').planOperations(task).operations[0];
  const r = clear.verifyPreservation(op, { primary: before.primary }, { primary: observe(after, req).primary });
  assert.equal(before.tier.tier, 'tiered');
  assert.equal(r.ok, true, JSON.stringify(r.evidence));
});

// ---- 5. clear_range keeps its own checks on the tiered observation ---------------------------------------------------

test('OWNER POINT 2: clear_range L1 (format / fill) and L3 (bold) on the tiered observation -- unchanged contract', () => {
  const b = bigBook();
  const task = { id: 't', steps: [{ intent: 'clear_range', id: 'c1', sheet: 'S', range: 'A10:B10' }] };
  const req = liveScope(task);
  const op = require('../../src/adapters/xlsx/adapter.cjs').planOperations(task).operations[0];
  const s = b.sheets[0];
  s.cells.set('A10', { ...s.cells.get('A10'), bold: true, nf: '0.00', fill: 255 });
  const before = observe(b, req);
  assert.equal(before.tier.tier, 'tiered');
  assert.equal(clear.classify(before.primary, op).state, 'APPLY', 'the gate reads format and bold in T1');
  const good = clone(b);
  set(good, 'S', 'A10', ''); set(good, 'S', 'B10', '');
  const okAfter = observe(good, req);
  assert.equal(clear.verifyPrimitive(op, { primary: before.primary }, { primary: okAfter.primary }).ok, true);
  assert.equal(clear.verifyPreservation(op, { primary: before.primary }, { primary: okAfter.primary }).ok, true);
  const lostBold = clone(good);
  lostBold.sheets[0].cells.set('A10', { ...lostBold.sheets[0].cells.get('A10'), bold: false });
  const rb = clear.verifyPreservation(op, { primary: before.primary }, { primary: observe(lostBold, req).primary });
  assert.equal(rb.ok, false, 'the bold check is NOT replaced by the digest');
  assert.equal(rb.evidence.targetBold.changed[0].cell, 'A10');
  const lostFormat = clone(good);
  lostFormat.sheets[0].cells.set('A10', { ...lostFormat.sheets[0].cells.get('A10'), nf: 'General' });
  assert.equal(clear.verifyPrimitive(op, { primary: before.primary }, { primary: observe(lostFormat, req).primary }).ok, false, 'format kept is still L1');
});

test('write_range gates on the tiered observation: classification, write-over-formula refusal, SATISFIED retry', () => {
  const b = bigBook();
  const over = writeTask(['w1', 'E5:E6', 'x']);
  const o = observe(b, liveScope(over));
  assert.equal(o.tier.tier, 'tiered');
  assert.equal(cells.classify(o.primary, opOf(over, 0)).evidence.reason, 'xlsx-write-over-formula-refused');
  const task = writeTask(['w1', 'L1:L10', 'a']);
  const done = apply(clone(b), opOf(task, 0));
  assert.equal(cells.classify(observe(done, liveScope(task)).primary, opOf(task, 0)).state, 'SATISFIED');
});

test('a formula OUTSIDE the windows whose VALUE follows a written target is not a change (its text is what L3 keeps)', () => {
  const b = bigBook();
  set(b, 'S', 'M5', '=L5');
  const task = writeTask(['w1', 'L1:L10', 'a']);
  const req = liveScope(task);
  const before = observe(b, req);
  const after = observe(apply(clone(b), opOf(task, 0)), req);
  const v = o => run(readQuery([{ sheet: 'S', range: 'M5' }], { usedRanges: false }), o).cells[0].v;
  assert.notEqual(v(b), v(apply(clone(b), opOf(task, 0))), 'positive control: the formula value really moved');
  assert.equal(l3(opOf(task, 0), before, after).ok, true);
});

// ---- 6. end to end through withSession + the executor, on a stand-in driver --------------------------------------

const { withSession } = require('../../src/core/session/lifecycle.cjs');
const { createExecutor } = require('../../src/core/execution/executor.cjs');
const { xlsxAdapter } = require('../../src/adapters/xlsx/adapter.cjs');

/**
 * The XLSX driver's readLive / dispatch on the in-memory book: the REAL tier decision, query bodies and tiered payload;
 * the refusal note exactly as readLive appends it. `inject(dispatchNo, book)` changes the book after a dispatch (a
 * foreign edit between two observations). forceSave / close follow the fake driver's confirmation contract.
 */
function standInDriver(b, { inject = null } = {}) {
  const calls = { dispatch: 0, readLive: 0, forceSave: 0 };
  const tiers = [];
  let confirmedAt = null;
  return {
    calls, tiers,
    async open() { return { ok: true, id: 'stand-in' }; },
    async close() { return { ok: true, confirmation: confirmedAt === null ? null : (calls.dispatch > confirmedAt ? { state: 'unconfirmed', reason: 'superseded-by-a-newer-save' } : { state: 'confirmed', reason: 'forcesave-confirmed' }) }; },
    async forceSave() { calls.forceSave += 1; confirmedAt = calls.dispatch; return { ok: true, confirmed: true, forceSave: { state: 'confirmed', reason: 'forcesave-confirmed' } }; },
    async canSave() { return { ok: true, canSave: false }; },
    async readPackage() { return { ok: true, parts: {} }; },
    async readLive(_h, request = {}) {
      calls.readLive += 1;
      const o = observe(b, request);
      tiers.push(o.tier.tier);
      const payload = { ...o.payload };
      if (o.tier.tier === 'refused' && payload.truncated) payload.unsupported = withLimitNote(payload.unsupported, o.tier);
      return { ok: true, authority: 'LIVE_READ', payload };
    },
    async dispatch(_h, command) {
      calls.dispatch += 1;
      const r = run(command.body, b);
      if (inject) inject(calls.dispatch, b);
      return { ok: !!(r && r.ok), dispatched: true };
    },
  };
}

async function session(b, task, opts) {
  const driver = standInDriver(b, opts);
  const options = { driver, fileId: 'stand-in', ready: { wait: async () => {} }, persistence: { wait: async () => {} } };
  const result = await withSession(options, xlsxAdapter, createExecutor(task, { observation: { live: liveScope(task) } }));
  return { result, driver };
}

test('E2E 10k: a two-step write, every observation tiered; ok + FORCESAVE_CONFIRMED; the reopened retry is SATISFIED 0 / 0', async () => {
  const b = bigBook();
  const task = writeTask(['w1', 'L1:L10', 'a'], ['w2', 'L20:L29', 'b']);
  const { result, driver } = await session(b, task);
  assert.equal(result.ok, true, JSON.stringify(result.execution.stopDetail || result.outcome));
  assert.equal(result.persistenceState, 'FORCESAVE_CONFIRMED');
  assert.deepEqual([...new Set(driver.tiers)], ['tiered']);
  assert.deepEqual(result.execution.preservationVerification.map(p => p.verdict), ['PASS', 'PASS']);
  assert.equal(b.sheets[0].cells.get('L25').raw, 'b25x0', 'the write landed');
  const retry = await session(b, task);
  assert.equal(retry.result.ok, true);
  assert.deepEqual([retry.result.metrics.writes, retry.result.metrics.commands], [0, 0]);
  assert.equal(retry.result.persistenceState, 'NOT_REQUIRED');
});

test('E2E L3 negative: a foreign change OUTSIDE the windows after step 1 -> step 1 not APPLIED, no step 2 dispatch, task not ok', async () => {
  const b = bigBook();
  const task = writeTask(['w1', 'L1:L10', 'a'], ['w2', 'L20:L29', 'b']);
  const { result, driver } = await session(b, task, { inject: (n, bk) => { if (n === 1) set(bk, 'S', 'B700', -1); } });
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 1, 'the stop rule: nothing after the failed step');
  assert.notEqual(result.execution.appliedOperations[0].status, 'APPLIED');
  assert.equal(result.execution.preservationVerification[0].verdict === 'PASS', false);
});

test('E2E OWNER POINT 1: step 1 corrupting step 2\'s target -> step 1\'s L3 fails, step 2 never runs', async () => {
  const b = bigBook();
  const task = writeTask(['w1', 'L1:L10', 'a'], ['w2', 'L20:L29', 'b']);
  const { result, driver } = await session(b, task, { inject: (n, bk) => { if (n === 1) set(bk, 'S', 'L22', 'corrupted'); } });
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 1);
  assert.notEqual(result.execution.appliedOperations[0].status, 'APPLIED');
});

test('E2E clear 10k: 1000 cells cleared, ok; the retry SATISFIED 0 / 0', async () => {
  const b = bigBook();
  const task = { id: 't', steps: [{ intent: 'clear_range', id: 'c1', sheet: 'S', range: 'A101:J200' }] };
  const { result, driver } = await session(b, task);
  assert.equal(result.ok, true, JSON.stringify(result.execution.stopDetail || result.outcome));
  assert.deepEqual([...new Set(driver.tiers)], ['tiered']);
  assert.equal(b.sheets[0].cells.get('A150').raw, '');
  const retry = await session(b, task);
  assert.equal(retry.result.ok, true);
  assert.deepEqual([retry.result.metrics.writes, retry.result.metrics.commands], [0, 0]);
});

test('E2E: another family above 5000 is refused BEFORE any write, the refusal named as an IMPLEMENTATION_LIMIT', async () => {
  const b = bigBook();
  const task = { id: 't', steps: [{ intent: 'merge_range', id: 'm1', sheet: 'S', range: 'L1:M1' }] };
  const { result, driver } = await session(b, task);
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
  assert.match(JSON.stringify(result.execution.appliedOperations[0].evidence), /IMPLEMENTATION_LIMIT:family-not-migrated-to-the-tiered-observation/);
});

test('E2E below 5000: the FULL tier, no digest, the same results as before', async () => {
  const b = book({ S: Object.fromEntries(rowsOf(1, 300, r => [`A${r}`, `T${r}`])) });
  const task = writeTask(['w1', 'B1:B10', 'a']);
  const { result, driver } = await session(b, task);
  assert.equal(result.ok, true);
  assert.deepEqual([...new Set(driver.tiers)], ['full']);
  assert.equal(result.execution.preservationVerification[0].evidence.evidence.tier, undefined, 'no tiered evidence below the ceiling');
});

// ---- 7. the digest's type rests on a CALIBRATED COUNT (owner review of e5878f3, owner channel) -------------------

/** Overrides COUNT for one cell: { value } (any JS value, undefined included) or 'THROW'. */
const setCount = (b, sheet, a, count) => { const s = b.sheets.find(x => x.name === sheet); s.cells.set(a, { ...s.cells.get(a), count }); return b; };
const digestRow = (b, windows) => run(digestQuery(windows, { canonical: true }), b);

test('POSITIVE CONTROL: COUNT 1 -> number, COUNT 0 -> text, the observation complete', () => {
  const b = book({ S: { A1: 'in', A2: 'text', B2: 42 } });
  const g = digestRow(b, [{ sheet: 'S', range: 'A1' }]);
  assert.deepEqual(g.unsupported, []);
  assert.deepEqual(JSON.parse(g.digests[0].canonical), [['A2', 'text', 'text', ''], ['B2', 'number', '42', '']]);
});

const UNCALIBRATED = { null: { value: null }, undefined: { value: undefined }, 2: { value: 2 }, "'1'": { value: '1' }, true: { value: true }, NaN: { value: NaN }, '-1': { value: -1 }, throws: 'THROW' };
for (const [name, count] of Object.entries(UNCALIBRATED)) {
  test(`COUNT ${name} on a non-formula cell outside the windows -> UNOBSERVABLE, never text; the observation incomplete`, () => {
    const b = setCount(book({ S: { A1: 'in', A2: 'text', B2: 42 } }), 'S', 'A2', count);
    const g = digestRow(b, [{ sheet: 'S', range: 'A1' }]);
    assert.equal(g.unsupported.length, 1, JSON.stringify(g.unsupported));
    assert.match(g.unsupported[0].field, /^S!A2/);
    assert.ok(!JSON.parse(g.digests[0].canonical).some(e => e[0] === 'A2'), 'no entry at all -- not a text entry');
    const inv = normalizeInventory(tieredPayload({ sheets: [{ name: 'S', index: 0, used: 'A1:B2' }], cells: [], truncated: false, unsupported: [], objects: {} }, g, { bound: 6000 }, {}));
    assert.equal(inv.complete, false);
  });
}

test('a formula cell needs no COUNT for its type: an uncalibrated COUNT there is not a gap', () => {
  const b = setCount(book({ S: { A1: 'in', C2: '=B2' } }), 'S', 'C2', { value: null });
  const g = digestRow(b, [{ sheet: 'S', range: 'A1' }]);
  assert.deepEqual(g.unsupported, []);
  assert.deepEqual(JSON.parse(g.digests[0].canonical), [['C2', 'formula', '', '=B2']]);
});

test('BEFORE the write: an uncalibrated COUNT outside the windows -> the write is refused, 0 dispatch', async () => {
  const b = setCount(bigBook(), 'S', 'B700', { value: 2 });
  const task = writeTask(['w1', 'L1:L10', 'a']);
  const o = observe(b, liveScope(task));
  assert.equal(o.tier.tier, 'tiered');
  assert.equal(o.primary.complete, false);
  assert.equal(cells.classify(o.primary, opOf(task, 0)).evidence.reason, 'xlsx-observation-unreadable-getters');
  const { result, driver } = await session(b, task);
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 0);
});

test('AFTER the write: the type read lost (COUNT turns uncalibrated) -> step 1 is NOT APPLIED, no further dispatch', async () => {
  const b = bigBook();
  const task = writeTask(['w1', 'L1:L10', 'a'], ['w2', 'L20:L29', 'b']);
  const { result, driver } = await session(b, task, { inject: (n, bk) => { if (n === 1) setCount(bk, 'S', 'B700', { value: undefined }); } });
  assert.equal(result.ok, false);
  assert.equal(driver.calls.dispatch, 1);
  assert.notEqual(result.execution.appliedOperations[0].status, 'APPLIED');
  assert.equal(result.execution.preservationVerification[0].evidence.evidence.reason, 'xlsx-preservation-observation-incomplete');
});

test('POSITIVE CONTROL for the same scenario: every COUNT calibrated -> both steps APPLIED, ok', async () => {
  const b = bigBook();
  const task = writeTask(['w1', 'L1:L10', 'a'], ['w2', 'L20:L29', 'b']);
  const { result } = await session(b, task);
  assert.equal(result.ok, true);
  assert.deepEqual(result.execution.appliedOperations.map(o => o.status), ['APPLIED', 'APPLIED']);
});
