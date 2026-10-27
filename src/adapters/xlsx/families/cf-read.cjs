'use strict';

const { CF_ENUMS } = require('./cf.cjs');

/**
 * L-08 unit 2: the PROJECTION behind the `read_conditional_formats` read tool -- a pure field-select over `inv.cfReads` (the driver's own
 * in-session reads, keyed '<sheet>!<range>'), shaped as the MIRROR of the add_conditional_format rule schema (cf.cjs): the same field names,
 * the same enum words, '#RRGGBB' colours. No new bridge call. LIVE basis: the L-08 1st-unit gap measurement (l08-r1: driver and api read all 56 writable rows).
 *
 * What the engine returns that the write schema has no word for is NEVER turned into a value: it is named.
 *   - a whole type the writer cannot create        -> { type: 'unsupported', engineType, reason }
 *   - one field value outside the write enum        -> the rule stays, `unrepresentable: [{ field, engine, reason }]`
 *   - a whole range that could not be read          -> { rules: null, reason: 'UNOBSERVABLE:...' }
 * "No Fill" (the engine's string for a rule without a fill) is normalized in ONE place (fillOf) to null, the write-side no-fill shape.
 */

const { OPS1, OPS2, DATABAR_FILL, DATABAR_DIRECTION, DATABAR_AXIS, ICON_SETS } = CF_ENUMS;
const invert = o => Object.freeze(Object.fromEntries(Object.entries(o).map(([k, v]) => [v, k])));
const OP_OF = invert({ ...OPS1, ...OPS2 });
const FILL_OF = invert(DATABAR_FILL);
const DIRECTION_OF = invert(DATABAR_DIRECTION);
const AXIS_OF = invert(DATABAR_AXIS);
const NO_FILL = 'No Fill';

const hexOf = n => '#' + (n >>> 0).toString(16).padStart(6, '0').toUpperCase().slice(-6);
const isColor = n => Number.isInteger(n) && n >= 0 && n <= 0xFFFFFF;

/** The ONE normalization of a fill: null (no fill) | '#RRGGBB' | { unrepresentable }. */
function fillOf(v) {
  if (v === null || v === undefined || v === NO_FILL) return { value: null };
  if (isColor(v)) return { value: hexOf(v) };
  return { unrepresentable: v };
}

const bad = (field, engine, reason) => ({ field, engine, reason });
const withBad = (rule, list) => (list.length ? { ...rule, unrepresentable: list } : rule);

function cellValue(e) {
  const op = OP_OF[e.operator];
  if (!op) return { type: 'unsupported', engineType: e.type, reason: `NO_WRITE_PATH: operator ${JSON.stringify(e.operator)} is not writable` };
  const f = fillOf(e.fill);
  const rule = { type: 'cellValue', operator: op, formula1: e.formula1 };
  if (e.formula2 !== '' && e.formula2 !== null && e.formula2 !== undefined) rule.formula2 = e.formula2;
  rule.fill = f.unrepresentable === undefined ? f.value : null;
  return withBad(rule, f.unrepresentable === undefined ? [] : [bad('fill', f.unrepresentable, 'not a colour')]);
}

function expression(e) {
  const f = fillOf(e.fill);
  const rule = { type: 'expression', formula: `=${e.formula1}`, fill: f.unrepresentable === undefined ? f.value : null };
  return withBad(rule, f.unrepresentable === undefined ? [] : [bad('fill', f.unrepresentable, 'not a colour')]);
}

function aboveAverage(e) {
  const list = [];
  if (typeof e.aboveBelow !== 'boolean') list.push(bad('above', e.aboveBelow, 'not a boolean'));
  if (e.numStdDev !== 0 && e.numStdDev !== 1) list.push(bad('numStdDev', e.numStdDev, 'the writable values are 0 and 1'));
  return withBad({ type: 'aboveAverage', above: e.aboveBelow === true, numStdDev: e.numStdDev === 1 ? 1 : 0 }, list);
}

function top10(e) {
  const list = [];
  if (!Number.isInteger(e.rank)) list.push(bad('rank', e.rank, 'not an integer'));
  if (typeof e.percent !== 'boolean') list.push(bad('percent', e.percent, 'not a boolean'));
  if (e.topBottom !== 'xlTop10Top' && e.topBottom !== 'xlTop10Bottom') list.push(bad('bottom', e.topBottom, 'neither xlTop10Top nor xlTop10Bottom'));
  return withBad({ type: 'top10', rank: e.rank, percent: e.percent === true, bottom: e.topBottom === 'xlTop10Bottom' }, list);
}

function unique(e) {
  if (e.dupeUnique === 'xlDuplicate') return { type: 'unsupported', engineType: e.type, reason: 'NO_WRITE_PATH: the duplicate-values variant is not writable' };
  return { type: 'unique' };
}

function colorScale(e) {
  let stops = null;
  try { stops = typeof e.stopColors === 'string' ? JSON.parse(e.stopColors) : null; } catch (err) { stops = null; }
  if (!Array.isArray(stops)) return { type: 'colorScale', unrepresentable: [bad('colors', e.stopColors === undefined ? null : e.stopColors, 'stop colours not readable')] };
  if (stops.length !== 3) return { type: 'unsupported', engineType: e.type, reason: `NO_WRITE_PATH: only the 3-stop colour scale is writable (this one has ${stops.length})` };
  if (!stops.every(isColor)) return { type: 'colorScale', unrepresentable: [bad('colors', stops, 'a stop colour is not an RGB integer')] };
  return { type: 'colorScale', colors: stops.map(hexOf) };
}

function databar(e) {
  const list = [];
  const rule = { type: 'databar' };
  if (e.barColor !== undefined) { if (isColor(e.barColor)) rule.barColor = hexOf(e.barColor); else list.push(bad('barColor', e.barColor, 'not an RGB integer')); }
  if (e.barFillType !== undefined) { if (FILL_OF[e.barFillType]) rule.barFillType = FILL_OF[e.barFillType]; else list.push(bad('barFillType', e.barFillType, 'not a writable fill type')); }
  if (e.direction !== undefined) { if (DIRECTION_OF[e.direction]) rule.direction = DIRECTION_OF[e.direction]; else list.push(bad('direction', e.direction, 'not a writable direction (leftToRight is not writable)')); }
  if (e.showValue !== undefined) { if (typeof e.showValue === 'boolean') rule.showValue = e.showValue; else list.push(bad('showValue', e.showValue, 'not a boolean')); }
  if (e.axisPosition !== undefined) { if (AXIS_OF[e.axisPosition]) rule.axisPosition = AXIS_OF[e.axisPosition]; else list.push(bad('axisPosition', e.axisPosition, 'not a writable axis position (midpoint is not writable)')); }
  return withBad(rule, list);
}

function iconSet(e) {
  const list = [];
  const rule = { type: 'iconSet' };
  if (e.iconSet !== undefined) { if (ICON_SETS.includes(e.iconSet)) rule.iconSet = e.iconSet; else list.push(bad('iconSet', e.iconSet, e.iconSet === 'xl3Arrows' ? 'xl3Arrows is not writable (the engine stores it as xl3TrafficLights2)' : 'not a writable icon set')); }
  if (e.reverseOrder !== undefined) { if (typeof e.reverseOrder === 'boolean') rule.reverseOrder = e.reverseOrder; else list.push(bad('reverseOrder', e.reverseOrder, 'not a boolean')); }
  if (e.showIconOnly !== undefined) { if (typeof e.showIconOnly === 'boolean') rule.showIconOnly = e.showIconOnly; else list.push(bad('showIconOnly', e.showIconOnly, 'not a boolean')); }
  return withBad(rule, list);
}

const BY_TYPE = Object.freeze({ xlCellValue: cellValue, xlExpression: expression, xlAboveAverageCondition: aboveAverage, xlTop10: top10, xlUniqueValues: unique, xlColorScale: colorScale, xlDatabar: databar, xlIconSets: iconSet });

/** One driver entry -> one rule of the write schema, plus where it applies. */
function ruleOf(e) {
  const f = Object.hasOwn(BY_TYPE, e.type) ? BY_TYPE[e.type] : null;
  const rule = f ? f(e) : { type: 'unsupported', engineType: e.type, reason: 'NO_WRITE_PATH: rule type is not writable' };
  return { ...rule, appliesTo: e.appliesTo };
}

/** inv.cfReads['<sheet>!<range>'] -> the tool result. A string is a whole-range gap, never guessed as "no rules". */
function projectCfRead(inv, a) {
  const r = inv.cfReads ? inv.cfReads[`${a.sheet}!${a.range}`] : undefined;
  if (r === undefined) return { sheet: a.sheet, range: a.range, rules: null, reason: 'UNOBSERVABLE:not-read' };
  if (typeof r === 'string') return { sheet: a.sheet, range: a.range, rules: null, reason: r };
  return { sheet: a.sheet, range: a.range, rules: r.map(ruleOf) };
}

module.exports = { projectCfRead, ruleOf, fillOf };
