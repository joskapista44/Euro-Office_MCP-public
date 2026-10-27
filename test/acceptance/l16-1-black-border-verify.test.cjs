'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const format = require('../../src/adapters/xlsx/families/format.cjs');
const { planOperations } = require('../../src/adapters/xlsx/adapter.cjs');
const { key } = require('../../src/adapters/xlsx/inventory.cjs');

/**
 * L-16/1 LIVE (prod DS /15): fill 'none' / border 'remove' on a cell with a BLACK diagonal answered VERIFY_CANNOT_PROVE_TERMINAL_STATE although the write had
 * landed (the retry was SATISFIED, a reopen showed fill cleared / Top gone / diagonal kept). The cause: the composite re-issues a kept black border explicitly,
 * which reads c: 0 in the writing session, while the fresh read BEFORE the write had c: null (black reads 0 in a session and null after a reopen) -- the L1/L3
 * drift comparison saw `border.d` {s:7,c:null} -> {s:7,c:0} as a leaked change. The cells below have the exact shape of the LIVE read (DIAG_BEFORE / DIAG_AFTER).
 */

const cell = (border, fill) => ({
  merge: 'D10', nf: 'General', fill, bold: false, italic: false, wrap: false, rot: 0, readingOrder: 0, indent: 0, shrink: false, f: '', v: 10,
  font: { bold: null, italic: null, name: 'Calibri', size: 11, color: 0, underline: 'xlUnderlineStyleNone', strike: null, sub: null, sup: null },
  align: { h: null, v: 0 },
  border: { t: { s: 0, c: 0 }, b: { s: 0, c: 0 }, l: { s: 0, c: 0 }, r: { s: 0, c: 0 }, d: { s: 7, c: null }, dd: true, du: false, ...border },
});
const obs = (a10, extra = {}) => ({ primary: { cells: { [key('B', 'D10')]: a10, ...extra } } });
const planned = format => planOperations({ id: 't', steps: [{ intent: 'format_range', sheet: 'B', range: 'D10:D10', format }] }).operations[0];

const FILL = 3832776;
const wasFill = obs(cell({}, FILL));

test('fill none: a kept BLACK diagonal that reads 0 after the restore (null before) is NOT drift', () => {
  const op = planned({ fill: 'none' });
  const r = format.verifyPrimitive(op, wasFill, obs(cell({ d: { s: 7, c: 0 } }, 'S:No Fill')));
  assert.equal(r.ok, true, JSON.stringify(r.evidence));
});

test('fill none: an unchanged read (null stays null) and a kept black SIDE with the same null -> 0 are fine too', () => {
  const op = planned({ fill: 'none' });
  assert.equal(format.verifyPrimitive(op, wasFill, obs(cell({}, 'S:No Fill'))).ok, true);
  assert.equal(format.verifyPrimitive(op, obs(cell({ l: { s: 7, c: null } }, FILL)), obs(cell({ l: { s: 7, c: 0 } }, 'S:No Fill'))).ok, true);
});

test('NEG: the tolerance is only null == black on the SAME style -- another colour, another style, a flipped direction, a lost diagonal are still drift', () => {
  const op = planned({ fill: 'none' });
  const drift = (border, was = wasFill) => format.verifyPrimitive(op, was, obs(cell(border, 'S:No Fill')));
  assert.equal(drift({ d: { s: 7, c: 255 } }).ok, false);                   // black -> red
  assert.equal(drift({ d: { s: 12, c: 0 } }).ok, false);                    // thin -> medium
  assert.equal(drift({ dd: false, du: true }).ok, false);                   // direction flipped
  assert.equal(drift({ d: { s: 0, c: 0 }, dd: false }).ok, false);          // diagonal lost
  assert.equal(drift({ d: { s: 7, c: 0 } }, obs(cell({ d: { s: 7, c: 255 } }, FILL))).ok, false);   // red -> black (the other direction of the tolerance)
});

test('NEG: the tolerance is scoped to the restoring ops -- a plain format op still counts a null -> 0 border colour change as drift', () => {
  const op = planned({ font: { bold: true } });
  const was = obs(cell({}, FILL));
  const now = obs({ ...cell({ d: { s: 7, c: 0 } }, FILL), bold: true, font: { ...cell({}, FILL).font, bold: true } });
  const r = format.verifyPrimitive(op, was, now);
  assert.equal(r.ok, false);
  assert.ok(r.evidence.otherKeysChanged.some(x => x.keys.includes('border.d')), JSON.stringify(r.evidence));
});

test('border remove (Top) on a cell with a black diagonal: the restored diagonal reading 0 is not drift (L1; L3 compares the other cells through the same driftOf)', () => {
  const op = planned({ border: { mode: 'remove', sides: ['top'] } });
  const d9 = (border = {}) => ({ ...cell({ d: { s: 0, c: 0 }, dd: false, ...border }, 'S:No Fill'), merge: 'D9', v: 0 });
  const was = obs(cell({ t: { s: 7, c: 255 } }, 'S:No Fill'), { [key('B', 'D9')]: d9({ b: { s: 7, c: 255 } }) });
  const now = obs(cell({ d: { s: 7, c: 0 } }, 'S:No Fill'), { [key('B', 'D9')]: d9() });
  const r = format.verifyPrimitive(op, was, now);
  assert.equal(r.ok, true, JSON.stringify(r.evidence));
});
