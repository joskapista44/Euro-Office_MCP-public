'use strict';

/**
 * L-02/1 (owner GO 2026-10-14, inventory row L-02): format_range format.font.name (string) and format.font.size (number) on the PRODUCT
 * surface. The family already planned, wrote (PUBLIC SetFontName / SetFontSize) and verified them (format.cjs FONT_VALUES, read on the
 * public GetCharacters(1, len).GetFont()); only the product schema held them back. MEASURED earlier (m7-s22-probe1, session + two reopens):
 * a fractional size and an uninstalled font name round-trip byte-exact. Offline, the product WRITE tool with the fake XLSX editor; the LIVE
 * proof is test/live/release-l02-font-product.live.cjs (RED on the pre-release build, GREEN on the rehearsal build).
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const W = require('../../src/server/write-tools.cjs');
const { xlsxAdapter } = require('../../src/adapters/xlsx/adapter.cjs');
const format = require('../../src/adapters/xlsx/families/format.cjs');
const { harness, req, createStore, book } = require('../fakes/write-wiring.cjs');

const font = (range, f, id = 'f') => ({ intent: 'format_range', id, sheet: 'Budget', range, format: { font: f } });
const durable = (h, a) => h.store.durable.find(x => x.name === 'Budget').cells[a];
const statuses = r => r.operations.map(o => o.status);
const validate = f => W.validateWriteRequest(req([font('A2', f)]));

test('L-02/1: name and size are released font keys; the other font keys and the key set of format are unchanged', () => {
  assert.equal(validate({ name: 'Courier New' }).ok, true);
  assert.equal(validate({ size: 10.5 }).ok, true);
  assert.equal(validate({ name: 'Arial', size: 14, bold: true, italic: false, color: '#112233' }).ok, true);
  assert.deepEqual([...W.FORMAT_RANGE_RELEASED_KEYS], ['numberFormat', 'border', 'fill', 'font', 'align', 'wrap']);
});

test('L-02/1: name, size, both -- each APPLIED + read back through the public reader; a retry is SATISFIED with 0 dispatch', async () => {
  const h = harness({ store: createStore(book()) });
  for (const f of [{ name: 'Courier New' }, { size: 14 }, { size: 10.5 }, { name: 'ZzNoSuchFont', size: 12.25 }]) {
    const r = await h.tool.handler(req([font('A2', f)]));
    assert.equal(r.ok, true, `${J(f)}: ${J(r).slice(0, 600)}`);
    assert.deepEqual(r.operations.map(o => [o.status, o.l1, o.l3]), [['APPLIED', 'PASS', 'PASS']], J(f));
    for (const [k, v] of Object.entries(f)) assert.equal(durable(h, 'A2').font[k], v, `${J(f)} ${k}`);
    const d0 = h.calls().dispatch;
    const again = await h.tool.handler(req([font('A2', f)]));
    assert.deepEqual([again.ok, statuses(again), h.calls().dispatch - d0], [true, ['SATISFIED'], 0], `retry ${J(f)}`);
  }
});

test('L-02/1: the size bounds [1, 409] are released and the family-bound values just outside are refused before dispatch', async () => {
  const h = harness({ store: createStore(book()) });
  for (const size of [1, 409, 0.5 + 0.5, 11.3]) assert.equal((await h.tool.handler(req([font('A3', { size })]))).ok, true, String(size));
  const d0 = h.calls().dispatch;
  for (const size of [0, 0.99, 409.01, 410, -1, '12', null, true, NaN]) {
    const r = await h.tool.handler(req([font('A3', { size })]));
    assert.deepEqual([r.ok, r.code, r.dispatched], [false, 'REFUSED_SHAPE', false], `size ${J(size)}`);
  }
  assert.equal(h.calls().dispatch, d0);
});

test('L-02/1: font names -- plain, accented, a name the machine does not have, 31 characters (the MEASURED xlsx limit) pass; empty / padded / control character / 32 and 100 characters are refused', async () => {
  const h = harness({ store: createStore(book()) });
  for (const name of ['Arial', 'Times New Roman', 'Árvíztűrő Tükörfúrógép', 'ZzNoSuchFont_x', "Brush Script's", 'x'.repeat(31), 'Á'.repeat(31)]) {
    assert.equal(validate({ name }).ok, true, J(name));
  }
  const d0 = h.calls().dispatch;
  for (const name of ['', ' Arial', 'Arial ', 'A\nB', 'A\tB', 'A\u0000B', 'x'.repeat(32), 'Á'.repeat(32), 'x'.repeat(100), 1, null]) {
    const r = await h.tool.handler(req([font('A3', { name })]));
    assert.deepEqual([r.ok, r.code, r.dispatched], [false, 'REFUSED_SHAPE', false], `name ${J(name)}`);
  }
  assert.equal(h.calls().dispatch, d0);
});

test('L-02/1: an empty font object is refused by name, and the message lists the released keys', async () => {
  const h = harness({ store: createStore(book()) });
  const r = await h.tool.handler(req([font('A2', {})]));
  assert.deepEqual([r.ok, r.code, r.dispatched], [false, 'REFUSED_SHAPE', false]);
  assert.match(r.reason, /at least one of color, bold, italic, name, size, underline, strikeout/);
  const unknown = await h.tool.handler(req([font('A2', { subscript: true })]));
  assert.deepEqual([unknown.ok, unknown.code], [false, 'UNKNOWN_FIELD']);
});

test('L-02/1: the neighbouring axes survive a name + size change (bold, italic, colour, wrap, number format, align, the value; an untouched neighbour is byte-equal)', async () => {
  const h = harness({ store: createStore(book()) });
  for (const step of [
    font('B2', { bold: true }, 'b'), font('B2', { color: '#3A7F2C' }, 'c'),
    { intent: 'format_range', id: 'w', sheet: 'Budget', range: 'B2', format: { wrap: true } },
    { intent: 'format_range', id: 'n', sheet: 'Budget', range: 'B2', format: { numberFormat: '0.00' } },
    { intent: 'format_range', id: 'a', sheet: 'Budget', range: 'B2', format: { align: { horizontal: 'center' } } },
  ]) assert.equal((await h.tool.handler(req([step]))).ok, true, step.id);
  const before = { ...durable(h, 'B2') };
  const neighbour = { ...durable(h, 'B3') };
  const r = await h.tool.handler(req([font('B2', { name: 'Courier New', size: 16 })]));
  assert.equal(r.ok, true, J(r).slice(0, 600));
  const after = durable(h, 'B2');
  assert.deepEqual([after.font.name, after.font.size], ['Courier New', 16]);
  assert.deepEqual([after.v, after.nf, after.wrap, after.bold, after.align, after.font.color], [before.v, before.nf, before.wrap, before.bold, before.align, before.font.color]);
  assert.deepEqual(durable(h, 'B3'), neighbour);
});

test('L-02/1: a range changes every cell and nothing outside; the lowering carries a hostile name as a JSON string literal, never as code', () => {
  const p = xlsxAdapter.planOperations({ id: 't', steps: [font('A2:A4', { name: 'Ar"ial\\x\'); evil(); //', size: 12.5 })] });
  assert.equal(p.ok, true);
  const body = format.lower(p.operations[0]).body;
  assert.ok(body.includes('r.SetFontName("Ar\\"ial\\\\x\'); evil(); //");'), body.split('\n').filter(l => /SetFont/.test(l)).join('|'));
  assert.doesNotThrow(() => new Function('Api', `return (async () => { ${body} })();`), 'the lowered command still parses');
});

test('L-02/1: a range write reaches A2:A4 and not A5', async () => {
  const h = harness({ store: createStore(book()) });
  const outside = { ...durable(h, 'A5') };
  const r = await h.tool.handler(req([font('A2:A4', { name: 'Courier New', size: 13 })]));
  assert.equal(r.ok, true, J(r).slice(0, 600));
  for (const a of ['A2', 'A3', 'A4']) assert.deepEqual([durable(h, a).font.name, durable(h, a).font.size], ['Courier New', 13], a);
  assert.deepEqual(durable(h, 'A5'), outside);
});

test('L-02/1 RED controls: a write that does not take effect / only reaches the first cell / leaks to a neighbour / changes the number format too is NOT APPLIED', async () => {
  for (const [fault, extra, range] of [['ignore', {}, 'A2'], ['partial', {}, 'A2:A4'], ['format-leaks', { leakCell: 'A5' }, 'A2'], ['format-extra-key', {}, 'A2']]) {
    const h = harness({ store: createStore(book()), script: { fault, ...extra } });
    const r = await h.tool.handler(req([font(range, { name: 'Courier New', size: 15 })]));
    assert.notEqual(statuses(r)[0], 'APPLIED', `${fault}: ${J(r).slice(0, 400)}`);
  }
});

test('L-02/1 (LIVE r1): the 31-character name limit holds at the FAMILY planner too -- a longer name is refused before any write, not cut after the reopen', () => {
  const plan = name => xlsxAdapter.planOperations({ id: 't', steps: [font('A2', { name })] });
  assert.equal(plan('x'.repeat(31)).ok, true);
  for (const n of ['x'.repeat(32), 'x'.repeat(100)]) {
    const p = plan(n);
    assert.equal(p.ok, false, String(n.length));
    assert.match(p.detail, /at most 31 characters/);
  }
});

function J(x) { return JSON.stringify(x); }
