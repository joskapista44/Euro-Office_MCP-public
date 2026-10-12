'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { readQuery } = require('../../src/drivers/eurooffice/xlsx-driver.cjs');
const { installBridge } = require('../fakes/ds-bridge-runtime.cjs');

/**
 * M7 S4, after the owner's review (owner channel, 2026-10-03): the engine's getBorderFull() asks the top neighbour only
 * when the (0-based) row index is > 1 and the left neighbour only when the column index is > 1 (cell SDK 9.3.4). The S4
 * edge probe MEASURED it (run m7-s4-edge-probe1-20261003): the top shared edge of row 2 and the left shared edge of
 * column B never showed the neighbour's side, in session and reopened; asking the neighbour's own getBorderFull() for
 * exactly those two sides matched the true shared edge on every cell. The driver's style read now does that -- the
 * same approved build-gated getter, no new dependency. A neighbour read that fails makes the cell's border unobservable.
 * B04 cutover (internal tracker point 8): the read now runs INSIDE the DS Bridge's range.appearance.read/v1 (the REAL served
 * block is installed on the stand-in); the bridge reads each cell through GetRange(<address>).range, so only addresses
 * that are NOT window cells count as neighbour reads.
 */

const prop = (s, c) => ({ s, c: c === null ? null : { getRgb: () => c } });
const EMPTY = prop(0, null);
const borderOf = (o = {}) => ({ t: EMPTY, b: EMPTY, l: EMPTY, r: EMPTY, d: EMPTY, dd: false, du: false, ...o });

/**
 * A one-sheet stand-in: `engine` = each cell's engine-side getBorderFull (the skip included); the neighbour reads go through
 * the PUBLIC Api.GetSheet(..).GetRange(address).range.getBorderFull(), recorded in `calls`.
 */
function read(engine, opts = {}) {
  const calls = [];
  const list = Object.keys(engine);
  const mk = a => ({
    GetAddress: () => a, GetValue: () => '', GetFormula: () => '', MergeArea: { GetAddress: () => a }, GetNumberFormat: () => 'General', GetFillColor: () => null, GetWrapText: () => false,
    GetCharacters: () => ({ GetFont: () => ({ GetBold: () => null, GetItalic: () => null, GetName: () => 'Calibri', GetSize: () => 11, GetColor: () => ({ GetRGB: () => 0 }) }) }),
    range: { getFont: () => ({ getBold: () => false }), getAlign: () => ({ getAlignHorizontal: () => 0, getAlignVertical: () => 0 }), getBorderFull: () => engine[a] || borderOf() },
  });
  const rangeOf = arg => {
    if (arg === 'WINDOW') return { GetAddress: () => 'X', ForEach: fn => list.forEach(a => fn(mk(a))) };
    if (list.includes(arg)) return mk(arg); // the bridge's own-cell read
    calls.push(arg);
    if (opts.neighbourThrows) throw new Error('model');
    return { range: { getBorderFull: () => (opts.model || {})[arg] || borderOf() } };
  };
  const wsApi = { GetName: () => 'S', GetIndex: () => 0, GetUsedRange: () => null, GetRange: rangeOf };
  const Api = { GetSheets: () => [wsApi], GetSheet: n => (n === 'S' ? wsApi : null), WorksheetFunction: { COUNT: () => 0 } };
  installBridge(Api);
  const out = new Function('Api', readQuery([{ sheet: 'S', range: 'WINDOW', style: true }], {}))(Api).cells;
  return { cells: Object.fromEntries(out.map(c => [c.a, c])), calls };
}

test('row 2 top / column B left: the neighbour\'s facing side is read (the engine skips it); other positions untouched', () => {
  // the engine-side reads: B2 and C2 / B3 see nothing of the neighbours on those two sides (the measured skip)
  const engine = { B2: borderOf(), C2: borderOf(), B3: borderOf(), C3: borderOf({ t: prop(12, 16711680) }) };
  const model = { B1: borderOf({ b: prop(13, null) }), C1: borderOf({ b: prop(12, 16711680) }), A2: borderOf({ r: prop(13, null) }), A3: borderOf({ r: prop(6, 32768) }) };
  const { cells, calls } = read(engine, { model });
  const by = a => cells[a];
  assert.deepEqual(by('B2').border.t, { s: 13, c: null });
  assert.deepEqual(by('B2').border.l, { s: 13, c: null });
  assert.deepEqual(by('C2').border.t, { s: 12, c: 16711680 });
  assert.deepEqual(by('C2').border.l, { s: 0, c: null }, 'column C: the engine already reads the left neighbour, nothing added');
  assert.deepEqual(by('B3').border.l, { s: 6, c: 32768 });
  assert.deepEqual(by('B3').border.t, { s: 0, c: null }, 'row 3: the engine already reads the top neighbour, nothing added');
  assert.deepEqual(by('C3').border.t, { s: 12, c: 16711680 }, 'an own side wins; no neighbour asked');
  assert.deepEqual([...calls].sort(), ['A2', 'A3', 'B1', 'C1'], 'only the skipped positions ask the neighbour, by its public address');
});

test('an own side in row 2 / column B is kept; the neighbour is not asked for it', () => {
  const { cells, calls } = read({ B2: borderOf({ t: prop(7, 255), l: prop(7, 255) }) }, { model: { B1: borderOf({ b: prop(13, null) }) } });
  assert.deepEqual(cells.B2.border.t, { s: 7, c: 255 });
  assert.deepEqual(cells.B2.border.l, { s: 7, c: 255 });
  assert.equal(calls.length, 0);
});

test('fail-closed: the neighbour read throws -> that cell\'s border is UNOBSERVABLE, not an empty side', () => {
  const { cells } = read({ B2: borderOf(), D5: borderOf() }, { neighbourThrows: true });
  assert.match(String(cells.B2.border), /^UNOBSERVABLE:/);
  assert.equal(typeof cells.D5.border, 'object', 'a cell that needs no neighbour read is unaffected');
});
