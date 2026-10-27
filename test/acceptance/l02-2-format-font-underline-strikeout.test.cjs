'use strict';

/**
 * L-02/2 (owner GO 2026-10-14, inventory row L-02): format_range format.font.underline (single | singleAccounting | double | doubleAccounting |
 * none) and format.font.strikeout (boolean) on the PRODUCT surface. The family already planned, wrote (PUBLIC range SetUnderline /
 * SetStrikeout) and verified them (format-font-effects.cjs, X-F05; the product path was LIVE-accepted by m7-f05-live1 on a number cell and a
 * formula cell); only the product schema held them back. subscript / superscript stay unreleased. Offline, the product WRITE tool with the fake
 * XLSX editor; the LIVE proof is test/live/release-l02-2-underline-strike-product.live.cjs (RED on the pre-release build, GREEN on the rehearsal).
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const W = require('../../src/server/write-tools.cjs');
const { UNDERLINE } = require('../../src/adapters/xlsx/families/format-font-effects.cjs');
const { harness, req, createStore, book } = require('../fakes/write-wiring.cjs');

const font = (range, f, id = 'f') => ({ intent: 'format_range', id, sheet: 'Budget', range, format: { font: f } });
const durable = (h, a) => h.store.durable.find(x => x.name === 'Budget').cells[a];
const statuses = r => r.operations.map(o => o.status);
const J = x => JSON.stringify(x);
const validate = f => W.validateWriteRequest(req([font('A2', f)]));

test('L-02/2: the five underline types and the strikeout boolean are released; subscript / superscript are NOT; the format key set is unchanged', () => {
  for (const underline of Object.keys(UNDERLINE)) assert.equal(validate({ underline }).ok, true, underline);
  assert.equal(validate({ strikeout: true }).ok, true);
  assert.equal(validate({ strikeout: false }).ok, true);
  assert.equal(validate({ underline: 'double', strikeout: true, name: 'Arial', size: 12, bold: true }).ok, true);
  for (const f of [{ subscript: true }, { superscript: false }, { strike: true }]) assert.deepEqual([validate(f).ok, validate(f).code], [false, 'UNKNOWN_FIELD'], J(f));
  assert.deepEqual([...W.FORMAT_RANGE_RELEASED_KEYS], ['numberFormat', 'border', 'fill', 'font', 'align', 'wrap']);
});

test('L-02/2: every underline type -- APPLIED + read back through the public reader (xlUnderlineStyle<Type>); a retry is SATISFIED with 0 dispatch', async () => {
  const h = harness({ store: createStore(book()) });
  for (const [underline, read] of Object.entries(UNDERLINE).filter(([k]) => k !== 'none')) {
    const r = await h.tool.handler(req([font('A2', { underline })]));
    assert.equal(r.ok, true, `${underline}: ${J(r).slice(0, 600)}`);
    assert.deepEqual(r.operations.map(o => [o.status, o.l1, o.l3]), [['APPLIED', 'PASS', 'PASS']], underline);
    assert.equal(durable(h, 'A2').font.underline, read, underline);
    const d0 = h.calls().dispatch;
    const again = await h.tool.handler(req([font('A2', { underline })]));
    assert.deepEqual([again.ok, statuses(again), h.calls().dispatch - d0], [true, ['SATISFIED'], 0], `retry ${underline}`);
  }
});

test('L-02/2: underline "none" removes an underline that was set (single -> none), and none on a never-underlined cell is SATISFIED', async () => {
  const h = harness({ store: createStore(book()) });
  assert.equal((await h.tool.handler(req([font('A3', { underline: 'single' })]))).ok, true);
  const off = await h.tool.handler(req([font('A3', { underline: 'none' })]));
  assert.deepEqual([off.ok, statuses(off)], [true, ['APPLIED']], J(off).slice(0, 500));
  assert.equal(durable(h, 'A3').font.underline, UNDERLINE.none);
  const never = await h.tool.handler(req([font('A4', { underline: 'none' })]));
  assert.equal(never.ok, true, J(never).slice(0, 500));
});

test('L-02/2: strikeout true, then false -- APPLIED each (false reads null, the bold/italic ambiguity), retry SATISFIED with 0 dispatch', async () => {
  const h = harness({ store: createStore(book()) });
  for (const value of [true, false]) {
    const r = await h.tool.handler(req([font('A2', { strikeout: value })]));
    assert.equal(r.ok, true, `${value}: ${J(r).slice(0, 600)}`);
    assert.deepEqual(statuses(r), ['APPLIED'], String(value));
    assert.equal(durable(h, 'A2').font.strike, value ? true : null, String(value));
    const d0 = h.calls().dispatch;
    const again = await h.tool.handler(req([font('A2', { strikeout: value })]));
    assert.deepEqual([again.ok, statuses(again), h.calls().dispatch - d0], [true, ['SATISFIED'], 0], `retry ${value}`);
  }
});

test('L-02/2: a NUMBER cell and a FORMULA cell take underline and strikeout (the cell types the S21 probes and m7-f05-live1 measured)', async () => {
  const h = harness({ store: createStore(book()) });
  assert.equal((await h.tool.handler(req([{ intent: 'write_range', id: 'fx', sheet: 'Budget', range: 'F1', formulas: [['=1+1']] }]))).ok, true);
  const num = await h.tool.handler(req([font('B2', { underline: 'double' })]));
  assert.equal(num.ok, true, J(num).slice(0, 500));
  assert.equal(durable(h, 'B2').font.underline, UNDERLINE.double);
  const fx = await h.tool.handler(req([font('F1', { strikeout: true })]));
  assert.equal(fx.ok, true, J(fx).slice(0, 500));
  assert.equal(durable(h, 'F1').font.strike, true);
});

test('L-02/2: the neighbouring font axes survive (name, size, bold, italic, colour) -- and wrap, number format, the value; an untouched neighbour is byte-equal', async () => {
  const h = harness({ store: createStore(book()) });
  for (const f of [{ name: 'Courier New' }, { size: 16 }, { bold: true }, { italic: true }, { color: '#3A7F2C' }]) assert.equal((await h.tool.handler(req([font('A2', f)]))).ok, true, J(f));
  const before = { ...durable(h, 'A2'), font: { ...durable(h, 'A2').font } };
  const neighbour = { ...durable(h, 'A3') };
  const r = await h.tool.handler(req([font('A2', { underline: 'single' })]));
  assert.equal(r.ok, true, J(r).slice(0, 600));
  const r2 = await h.tool.handler(req([font('A2', { strikeout: true })]));
  assert.equal(r2.ok, true, J(r2).slice(0, 600));
  const after = durable(h, 'A2');
  for (const k of ['name', 'size', 'bold', 'italic', 'color']) assert.equal(J(after.font[k]), J(before.font[k]), k);
  assert.deepEqual([after.v, after.wrap, after.nf], [before.v, before.wrap, before.nf]);
  assert.deepEqual([after.font.underline, after.font.strike], [UNDERLINE.single, true]);
  assert.deepEqual(durable(h, 'A3'), neighbour);
});

test('L-02/2: an unknown underline type, a non-boolean strikeout and a non-string underline are refused BEFORE dispatch (0 dispatch)', async () => {
  const h = harness({ store: createStore(book()) });
  for (const f of [{ underline: 'dotted' }, { underline: 'Single' }, { underline: '' }, { underline: true }, { underline: 1 }, { underline: null }, { underline: 'toString' },
    { strikeout: 'true' }, { strikeout: 1 }, { strikeout: null }]) {
    const r = await h.tool.handler(req([font('A2', f)]));
    assert.deepEqual([r.ok, r.code, r.dispatched], [false, 'REFUSED_SHAPE', false], J(f));
  }
  assert.equal(h.calls().dispatch, 0);
});

test('L-02/2: a range changes every cell and nothing outside', async () => {
  const h = harness({ store: createStore(book()) });
  const outside = { ...durable(h, 'A5') };
  const r = await h.tool.handler(req([font('A2:A4', { underline: 'double' })]));
  assert.equal(r.ok, true, J(r).slice(0, 600));
  for (const a of ['A2', 'A3', 'A4']) assert.equal(durable(h, a).font.underline, UNDERLINE.double, a);
  assert.deepEqual(durable(h, 'A5'), outside);
});

test('L-02/2 RED controls: a write that does not take effect / only reaches the first cell / leaks to a neighbour / changes the number format too is NOT APPLIED', async () => {
  for (const [fault, extra, range] of [['ignore', {}, 'A2'], ['partial', {}, 'A2:A4'], ['format-leaks', { leakCell: 'A5' }, 'A2'], ['format-extra-key', {}, 'A2']]) {
    for (const f of [{ underline: 'double' }, { strikeout: true }]) {
      const h = harness({ store: createStore(book()), script: { fault, ...extra } });
      const r = await h.tool.handler(req([font(range, f)]));
      assert.notEqual(statuses(r)[0], 'APPLIED', `${fault} ${J(f)}: ${J(r).slice(0, 300)}`);
    }
  }
});
