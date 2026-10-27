'use strict';

const { parseRange, rangeText } = require('../a1.cjs');
const cells = require('./cells.cjs');
const { nameUses } = require('../dependencies.cjs');
const formula = require('../formula.cjs');

/**
 * XLSX `cf` family: add_conditional_format / delete_conditional_format (M6 wave 6).
 *
 * The reader is PUBLIC and range-level (the source's measured getters, B-D34): every rule the range reports, with
 * type, operator, formulas, fill and appliesTo. The whole-sheet reader of wave 3 stays as it is (it measures only
 * xlExpression, for the sheet-graph families).
 * Audit: B-D34 MIGRATE (xlCellValue + xlGreater + fill); B-D35 between / B-D36 other types NOT_MEASURED; B-D34's
 * defect: the source matched only the fields a request named -- a rule with the same condition but ANOTHER fill
 * "satisfied" the request.
 *
 * Contract (the measured shape; W6 probe C1-C4): rule = { type: 'cellValue', operator: 'greater', formula1, fill: '#RRGGBB' }.
 * Identity = type + operator + formula1 + fill + appliesTo == the range. add: SATISFIED when exactly one identical
 * rule is on the range; refused when a rule with the same condition but another fill is there (no partial identity)
 * or more than one identical. delete: SATISFIED when none; APPLY when exactly one -- the command re-identifies it by
 * its getters at dispatch (never by a stale index); refused when more than one.
 * L1 the identical count (1 / 0) and the rule count on the range +1 / -1; L3 every other rule the range reports and
 * every cell. Retry SATISFIED 0/0.
 *
 * M7 R18 (owner M7 GO 2026-10-03; tier D probe docs/evidence/M7-D-PROBE.md, every getter read back exactly in session
 * and after a reopen): type 'cellValue' with operator greater | less | equal (formula1) or between | notBetween
 * (formula1, formula2), fill optional (absent = the engine's 'No Fill', read as null); type 'expression' with a
 * formula ('=...', read back without the '=') and a fill. Identity adds formula2. The other types (top10, unique,
 * scales, data bars, icon sets) stay refused.
 *
 * M7 X-O01 (the S29/S29b raw-API probes, LIVE MEASURED): three MORE types, each with its OWN identity
 * (no operator/formula1/formula2/fill -- the driver's cfReads no longer calls those getters on these types, since
 * they throw; CF_IDENTITY_FIELDS below is the single source of truth for both the reader-shape assumption and the
 * identity comparison):
 *   type 'aboveAverage' { above: bool, numStdDev?: 0 | 1 } -- AddAboveAverage + SetAboveBelow + SetNumStdDev.
 *   type 'top10' { rank: 1..1000, percent: bool, bottom: bool } -- AddTop10 + SetRank + SetPercent + SetTopBottom.
 *   type 'unique' -- ONLY the xlUnique shape (AddUniqueValues, its own default; SetDupeUnique is never called).
 *     A request for the duplicate variant (`dupe: true`) is refused at PLAN time: the S29b probe measured that
 *     SetDupeUnique('xlDuplicate') makes the rule's own GetType() read back empty -- unidentifiable, not merely
 *     "a different value" -- so this family never attempts that write.
 * Colour scales, data bars and icon sets' Add/Get surface is public (M7 X-O01 enum-probe LIVE
 * MEASURED: ApiFormatConditions.AddColorScale/AddDatabar/AddIconSetCondition all exist and create a rule; the
 * driver's reader already reports their type + appliesTo + the few scalar fields that probe found read-only-safe
 * (xlsx-driver.cjs CF_SCALAR_FIELDS: databar barColor; icon sets iconSet + showIconOnly; colour scale none)). What
 * is NOT measured is the CRITERIA objects' own setter shape (GetColorScaleCriteria / GetIconCriteria) and the exact
 * enum strings for every appearance setter the probe only named (BarFillType, Direction, MinPointType, AxisPosition,
 * IconSet, ...) -- guessing those now would repeat the W-08 CreateColorFromRGB mistake at three times the scope.
 * X-O01 (LIVE phase 1 + 2 probes, 2026-10-12, docs/evidence/xo01-cf-appearance-probe-phase*-20261012.json): the
 * appearance setters ARE measured now -- databar barColor / barFillType / direction / showValue / axisPosition, icon set
 * iconSet (19 of 20 names; xl3Arrows refused) / reverseOrder / showIconOnly, colour scale stop colours -- and an
 * appearance key a rule asks for is part of its identity (databarRule / iconSetRule / colorScaleRule below). Without
 * appearance keys the rule is still the engine's OWN DEFAULT, exactly as before: { type: 'colorScale' } -> AddColorScale(3) (the one size-arg the probe exercised and confirmed
 * works); { type: 'databar' } -> AddDatabar(); { type: 'iconSet' } -> AddIconSetCondition(). Identity is type +
 * appliesTo only (CF_IDENTITY_FIELDS[...] = [], same shape as 'unique' -- no field to compare, since none of the
 * default appearance was ever set by this family). A later unit MAY widen these with a measured probe for the
 * criteria/enum shapes; until then the create call itself, and whether it actually persists past a reopen, is this
 * unit's own first LIVE acceptance (a maintainer's #ref verdict: "a pontos parameter-alak ... az implementacio elso LIVE
 * elfogadasa").
 */

const INTENTS = Object.freeze(['add_conditional_format', 'delete_conditional_format']);
const HEX = /^#[0-9A-Fa-f]{6}$/;

// M7 X-O01: the rule types with their OWN scalar identity (no fill, no operator / formula).
const SCALAR_RULES = Object.freeze({
  aboveAverage: r => aboveAverageRule(r),
  top10: r => top10Rule(r),
  unique: r => uniqueRule(r),
  colorScale: r => colorScaleRule(r),
  databar: r => databarRule(r),
  iconSet: r => iconSetRule(r),
});

// X-O01 appearance (LIVE phase 1 + 2, docs/evidence/xo01-cf-appearance-probe-phase*-20261012.json): ONLY the values a
// probe SET and read back after a reopen, or the engine default it read (set explicitly in the acceptance).
const DATABAR_FILL = Object.freeze({ solid: 'xlDataBarFillSolid', gradient: 'xlDataBarFillGradient' });
const DATABAR_DIRECTION = Object.freeze({ context: 'xlContext', rightToLeft: 'xlRTL' }); // xlLTR: not measured
const DATABAR_AXIS = Object.freeze({ automatic: 'xlDataBarAxisAutomatic', none: 'xlDataBarAxisNone' }); // midpoint: not measured
// SetIconSet round-trip, 19 of 20 names; 'xl3Arrows' reads back 'xl3TrafficLights2' (an engine mapping defect) -> refused
const ICON_SETS = Object.freeze(['xl3ArrowsGray', 'xl3Flags', 'xl3TrafficLights1', 'xl3TrafficLights2', 'xl3Signs', 'xl3Symbols', 'xl3Symbols2', 'xl4Arrows', 'xl4ArrowsGray', 'xl4RedToBlack', 'xl4CRV', 'xl4TrafficLights', 'xl5Arrows', 'xl5ArrowsGray', 'xl5CRV', 'xl5Quarters', 'xl3Stars', 'xl3Triangles', 'xl5Boxes']);
const intOfHex = h => parseInt(h.slice(1), 16);
const rgbOfHex = h => [1, 3, 5].map(i => parseInt(h.slice(i, i + 2), 16));

/** The appearance keys a rule asked for are its identity (the engine default -- nothing asked -- keeps the type-only one). */
function databarRule(r) {
  const allowed = ['type', 'barColor', 'barFillType', 'direction', 'showValue', 'axisPosition'];
  if (Object.keys(r).some(k => !allowed.includes(k))) return { problem: `a databar rule is { type: 'databar', ${allowed.slice(1).join('?, ')}? }` };
  const out = { type: 'xlDatabar', idKeys: [] };
  if ('barColor' in r) { if (typeof r.barColor !== 'string' || !HEX.test(r.barColor)) return { problem: "rule.barColor must be '#RRGGBB'" }; out.barColor = intOfHex(r.barColor); out.barRgb = rgbOfHex(r.barColor); out.idKeys.push('barColor'); }
  if ('barFillType' in r) { if (!Object.hasOwn(DATABAR_FILL, r.barFillType)) return { problem: `rule.barFillType must be one of ${Object.keys(DATABAR_FILL).join(' | ')}` }; out.barFillType = DATABAR_FILL[r.barFillType]; out.idKeys.push('barFillType'); }
  if ('direction' in r) { if (!Object.hasOwn(DATABAR_DIRECTION, r.direction)) return { problem: `rule.direction must be one of ${Object.keys(DATABAR_DIRECTION).join(' | ')} (leftToRight is not measured)` }; out.direction = DATABAR_DIRECTION[r.direction]; out.idKeys.push('direction'); }
  if ('showValue' in r) { if (typeof r.showValue !== 'boolean') return { problem: 'rule.showValue must be true or false' }; out.showValue = r.showValue; out.idKeys.push('showValue'); }
  if ('axisPosition' in r) { if (!Object.hasOwn(DATABAR_AXIS, r.axisPosition)) return { problem: `rule.axisPosition must be one of ${Object.keys(DATABAR_AXIS).join(' | ')} (midpoint is not measured)` }; out.axisPosition = DATABAR_AXIS[r.axisPosition]; out.idKeys.push('axisPosition'); }
  return out;
}

function iconSetRule(r) {
  const allowed = ['type', 'iconSet', 'reverseOrder', 'showIconOnly'];
  if (Object.keys(r).some(k => !allowed.includes(k))) return { problem: "an iconSet rule is { type: 'iconSet', iconSet?, reverseOrder?, showIconOnly? } (criteria thresholds are not supported)" };
  const out = { type: 'xlIconSets', idKeys: [] };
  if ('iconSet' in r) {
    if (r.iconSet === 'xl3Arrows') return { problem: "rule.iconSet 'xl3Arrows' is refused: the engine stores it as 'xl3TrafficLights2' (measured)" };
    if (!ICON_SETS.includes(r.iconSet)) return { problem: `rule.iconSet must be one of ${ICON_SETS.join(' | ')}` };
    out.iconSet = r.iconSet; out.idKeys.push('iconSet');
  }
  if ('reverseOrder' in r) { if (typeof r.reverseOrder !== 'boolean') return { problem: 'rule.reverseOrder must be true or false' }; out.reverseOrder = r.reverseOrder; out.idKeys.push('reverseOrder'); }
  if ('showIconOnly' in r) { if (typeof r.showIconOnly !== 'boolean') return { problem: 'rule.showIconOnly must be true or false' }; out.showIconOnly = r.showIconOnly; out.idKeys.push('showIconOnly'); }
  return out;
}

function colorScaleRule(r) {
  if (Object.keys(r).some(k => !['type', 'colors'].includes(k))) return { problem: "a colorScale rule is { type: 'colorScale', colors? } -- colors: three '#RRGGBB' stop colours (low, mid, high); stop types / values are not supported" };
  const out = { type: 'xlColorScale', idKeys: [] };
  if ('colors' in r) {
    if (!Array.isArray(r.colors) || r.colors.length !== 3 || !r.colors.every(c => typeof c === 'string' && HEX.test(c))) return { problem: "rule.colors must be three '#RRGGBB' colours (the 3-stop scale)" };
    out.stopColors = JSON.stringify(r.colors.map(intOfHex)); out.stopRgbs = r.colors.map(rgbOfHex); out.idKeys.push('stopColors');
  }
  return out;
}

function plan(step, index) {
  const at = `step ${index}: ${step.intent}`;
  const unknown = Object.keys(step).filter(k => !['intent', 'id', 'sheet', 'range', 'rule'].includes(k));
  if (unknown.length) return { ok: false, detail: `${at} does not support field(s): ${unknown.join(', ')}` };
  if (typeof step.sheet !== 'string' || step.sheet === '') return { ok: false, detail: `${at} requires a sheet name` };
  const g = parseRange(step.range);
  if (!g) return { ok: false, detail: `${at} requires a plain A1 range` };
  const r = step.rule;
  const rule = ruleOf(r);
  if (rule.problem) return { ok: false, detail: `${at}: ${rule.problem}` };
  return {
    ok: true,
    operation: {
      id: step.id || `s${index}`,
      intent: step.intent,
      family: 'cf',
      // a text condition may name a defined name: a step deleting it clashes (selector.uses, W9 composition)
      selector: Object.freeze({ sheet: step.sheet, range: rangeText(g), uses: usesOf(step) }),
      rule: Object.freeze(rule),
      cells: Object.freeze([]),
      preservationMode: 'PRIMITIVE_INTEGRAL',
    },
  };
}

const OPS1 = Object.freeze({ greater: 'xlGreater', less: 'xlLess', equal: 'xlEqual' });
const OPS2 = Object.freeze({ between: 'xlBetween', notBetween: 'xlNotBetween' });
const operand = x => (typeof x === 'string' && x !== '' && !/^=/.test(x)) || Number.isFinite(x);

/** The rule a step asks for, in the editor's terms (fill as the RGB integer, null = no fill) -- or { problem }. */
function ruleOf(r) {
  if (!r || typeof r !== 'object' || Array.isArray(r)) return { problem: 'rule must be an object' };
  if (Object.hasOwn(SCALAR_RULES, r.type)) return SCALAR_RULES[r.type](r);
  const fill = fillOf(r);
  if (fill.problem) return fill;
  if (r.type === 'expression') return expressionRule(r, fill);
  return r.type === 'cellValue' ? cellValueRule(r, fill) : { problem: "rule.type must be 'cellValue', 'expression', 'aboveAverage', 'top10', 'unique', 'colorScale', 'databar' or 'iconSet' (the duplicate-values variant of 'unique' is not supported; colorScale/databar/iconSet take only their measured appearance keys)" };
}

/** M7 X-O01: no fill, no operator/formula -- each has its OWN scalar identity (CF_IDENTITY_FIELDS). */
function aboveAverageRule(r) {
  if (Object.keys(r).some(k => !['type', 'above', 'numStdDev'].includes(k))) return { problem: "an aboveAverage rule is { type: 'aboveAverage', above, numStdDev? }" };
  if (typeof r.above !== 'boolean') return { problem: 'rule.above must be true or false' };
  const numStdDev = 'numStdDev' in r ? r.numStdDev : 0;
  if (numStdDev !== 0 && numStdDev !== 1) return { problem: 'rule.numStdDev must be 0 or 1 (the measured values)' };
  return { type: 'xlAboveAverageCondition', aboveBelow: r.above, numStdDev };
}

function top10Rule(r) {
  if (Object.keys(r).some(k => !['type', 'rank', 'percent', 'bottom'].includes(k))) return { problem: "a top10 rule is { type: 'top10', rank, percent, bottom }" };
  if (!Number.isInteger(r.rank) || r.rank < 1 || r.rank > 1000) return { problem: 'rule.rank must be an integer in [1, 1000]' };
  if (typeof r.percent !== 'boolean') return { problem: 'rule.percent must be true or false' };
  if (typeof r.bottom !== 'boolean') return { problem: 'rule.bottom must be true or false' };
  return { type: 'xlTop10', rank: r.rank, percent: r.percent, topBottom: r.bottom ? 'xlTop10Bottom' : 'xlTop10Top' };
}

/** CSAK the xlUnique shape (the S29b probe's own LIVE measurement: the duplicate variant breaks GetType()). */
function uniqueRule(r) {
  if (Object.keys(r).some(k => !['type', 'dupe'].includes(k))) return { problem: "a unique rule is { type: 'unique', dupe? }" };
  if ('dupe' in r && r.dupe !== false) return { problem: "rule.dupe must be false (or absent) -- the duplicate variant is measured unidentifiable (GetType() reads empty after SetDupeUnique('xlDuplicate'))" };
  return { type: 'xlUniqueValues' };
}


function expressionRule(r, fill) {
  if (Object.keys(r).some(k => !['type', 'formula', 'fill'].includes(k))) return { problem: "an expression rule is { type: 'expression', formula, fill }" };
  if (typeof r.formula !== 'string' || formula.canonical(r.formula) === null || r.formula.length > 1024) return { problem: "rule.formula must be a formula starting with '='" };
  if (fill.fill === null) return { problem: 'an expression rule needs a fill (the measured shape)' };
  return { type: 'xlExpression', operator: '', formula1: r.formula.slice(1), formula2: '', ...fill };
}

function cellValueRule(r, fill) {
  if (Object.keys(r).some(k => !['type', 'operator', 'formula1', 'formula2', 'fill'].includes(k))) return { problem: 'rule must be { type, operator, formula1, formula2?, fill? }' };
  if (!operand(r.formula1)) return { problem: "rule.formula1 must be a number or a non-empty text without '='" };
  if (OPS1[r.operator]) {
    if ('formula2' in r) return { problem: `operator ${r.operator} takes formula1 only` };
    return { type: 'xlCellValue', operator: OPS1[r.operator], formula1: String(r.formula1), formula2: '', ...fill };
  }
  if (!OPS2[r.operator]) return { problem: `operator must be one of ${[...Object.keys(OPS1), ...Object.keys(OPS2)].join(' | ')}` };
  if (!operand(r.formula2)) return { problem: `operator ${r.operator} needs formula2 (a number or a text without '=')` };
  return { type: 'xlCellValue', operator: OPS2[r.operator], formula1: String(r.formula1), formula2: String(r.formula2), ...fill };
}

function fillOf(r) {
  if (!('fill' in r)) return { fill: null, rgb: null };
  if (typeof r.fill !== 'string' || !HEX.test(r.fill)) return { problem: "rule.fill must be '#RRGGBB'" };
  const [rr, gg, bb] = [1, 3, 5].map(i => parseInt(r.fill.slice(i, i + 2), 16));
  return { fill: rr * 65536 + gg * 256 + bb, rgb: Object.freeze([rr, gg, bb]) };
}

const f1 = x => String(x === null || x === undefined ? '' : x).replace(/^=/, '');
function usesOf(step) {
  const r = step.rule || {};
  const t = typeof r.formula === 'string' ? r.formula.slice(1) : (typeof r.formula1 === 'string' ? r.formula1 : null);
  if (t === null) return Object.freeze([]);
  const u = nameUses(`=${t}`);
  return Object.freeze(u.problem ? ['name:*'] : u.names.map(n => `name:${n.toLowerCase()}`));
}

/** M7 X-O01: each new type's OWN scalar identity fields (no fill) -- the single source of truth for the identity
 * check below AND for `lower()`'s dispatch-time re-identification (CF_IDENTITY_GETTERS maps the same keys to their
 * public getter names). */
const CF_IDENTITY_FIELDS = Object.freeze({
  xlAboveAverageCondition: ['aboveBelow', 'numStdDev'],
  xlTop10: ['rank', 'percent', 'topBottom'],
  xlUniqueValues: [],
  xlColorScale: [],
  xlDatabar: [],
  xlIconSets: [],
});
const CF_IDENTITY_GETTERS = Object.freeze({ aboveBelow: 'GetAboveBelow', numStdDev: 'GetNumStdDev', rank: 'GetRank', percent: 'GetPercent', topBottom: 'GetTopBottom',
  barFillType: 'GetBarFillType', direction: 'GetDirection', showValue: 'GetShowValue', axisPosition: 'GetAxisPosition', iconSet: 'GetIconSet', reverseOrder: 'GetReverseOrder', showIconOnly: 'GetShowIconOnly' });
// X-O01: the identity keys that are not one scalar getter -- a JS expression over the rule object `x` (dispatch-time)
const CF_IDENTITY_EXPRS = Object.freeze({
  barColor: 'rgbOf(x.GetBarColor())',
  stopColors: 'JSON.stringify((x.GetColorScaleCriteria() || []).map(c => rgbOf(c.GetColor())))',
});

const sameText = (a, b, expr) => (expr ? formula.sameFormula(`=${f1(a)}`, `=${b}`) : f1(a) === b);
function sameCondition(x, r) {
  if (x.type !== r.type) return false;
  const own = r.idKeys ? [...CF_IDENTITY_FIELDS[r.type], ...r.idKeys] : CF_IDENTITY_FIELDS[r.type];
  if (own) return own.every(k => x[k] === r[k]);
  return (r.type === 'xlExpression' || x.operator === r.operator) && sameText(x.formula1, r.formula1, r.type === 'xlExpression') && f1(x.formula2) === r.formula2;
}
function identical(x, r, range) {
  const appliesOk = String(x.appliesTo || '').replace(/\$/g, '') === range;
  if (CF_IDENTITY_FIELDS[r.type]) return sameCondition(x, r) && appliesOk;
  return sameCondition(x, r) && x.fill === r.fill && appliesOk;
}

function gate(p, op) {
  if (!p || p.format !== 'xlsx') return { ok: false, reason: 'xlsx-observation-missing' };
  if (!p.complete) return { ok: false, reason: p.truncated ? 'xlsx-observation-truncated' : 'xlsx-observation-unreadable-getters', unsupported: p.unsupported };
  const s = p.sheets.filter(x => x.name === op.selector.sheet);
  if (s.length !== 1) return { ok: false, reason: s.length > 1 ? 'xlsx-sheet-ambiguous' : 'xlsx-sheet-not-found', available: p.sheets.map(x => x.name) };
  const list = p.cfReads ? p.cfReads[`${op.selector.sheet}!${op.selector.range}`] : undefined;
  if (!Array.isArray(list)) return { ok: false, reason: 'xlsx-cf-read-unobservable', why: list === undefined ? 'not-read' : list };
  return { ok: true, list, same: list.filter(x => identical(x, op.rule, op.selector.range)), near: list.filter(x => sameCondition(x, op.rule) && !identical(x, op.rule, op.selector.range)) };
}

function classify(p, op) {
  const g = gate(p, op);
  if (!g.ok) return { state: 'FAIL_CLOSED', evidence: g };
  const where = { sheet: op.selector.sheet, range: op.selector.range, rule: op.rule };
  if (g.same.length > 1) return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-cf-rule-identity-ambiguous', count: g.same.length } };
  if (op.intent === 'add_conditional_format') {
    if (g.same.length === 1) return { state: 'SATISFIED', evidence: { ...where, proof: 'fresh-live-the-identical-rule-on-the-range' } };
    if (g.near.length) return { state: 'FAIL_CLOSED', evidence: { reason: 'xlsx-cf-same-condition-other-fill-or-range', near: g.near.slice(0, 10) } };
    return { state: 'APPLY', evidence: where };
  }
  if (g.same.length === 0) return { state: 'SATISFIED', evidence: { ...where, proof: 'fresh-live-no-identical-rule' } };
  return { state: 'APPLY', evidence: where };
}

function resolve(p, selector, ctx = {}) {
  const g = gate(p, ctx.operation);
  if (!g.ok) return { ok: false, target: null, selector, ambiguity: g.reason, detail: g };
  return { ok: true, target: Object.freeze({ sheet: selector.sheet, range: selector.range }) };
}

/** add: the public Add[/AboveAverage/Top10/UniqueValues] + its own setters. delete: re-identify the ONE identical
 * rule by its getters (type-branched -- CF_IDENTITY_FIELDS, never GetOperator / GetFormula1 / GetFormula2 /
 * GetFillColor on a type that doesn't have them), then Delete. */
function lower(op) {
  const { sheet, range } = op.selector;
  const r = op.rule;
  const body = op.intent === 'add_conditional_format' ? addBody(r, range) : deleteBody(r, range);
  const J = JSON.stringify;
  return {
    kind: 'callCommand',
    body: `
try {
  const name = ${J(sheet)};
  const ws = Api.GetSheet(name);
  if (!ws || ws.GetName() !== name) return { ok: false, stage: 'sheet-identity' };
  ${body}
  return { ok: true };
} catch (err) { return { ok: false, stage: 'threw', error: String(err) }; }
`,
    meta: Object.freeze({ family: 'cf', action: op.intent, sheet, range, rule: r }),
  };
}

const DEFAULT_ONLY_ADD = Object.freeze({ xlColorScale: 'AddColorScale(3)', xlDatabar: 'AddDatabar()', xlIconSets: 'AddIconSetCondition()' });

function addBody(r, range) {
  const J = JSON.stringify;
  if (r.type === 'xlAboveAverageCondition') return `const c = ws.GetRange(${J(range)}).GetFormatConditions().AddAboveAverage();
  c.SetAboveBelow(${J(r.aboveBelow)});
  c.SetNumStdDev(${J(r.numStdDev)});`;
  if (r.type === 'xlTop10') return `const c = ws.GetRange(${J(range)}).GetFormatConditions().AddTop10();
  c.SetRank(${J(r.rank)});
  c.SetPercent(${J(r.percent)});
  c.SetTopBottom(${J(r.topBottom)});`;
  if (r.type === 'xlUniqueValues') return `const c = ws.GetRange(${J(range)}).GetFormatConditions().AddUniqueValues();`;
  // M7 X-O01: default-only creation (no appearance setters -- the criteria/enum shapes are NEM-MERT, see the
  // family doc comment). AddColorScale(3) is the one size-arg a maintainer's #ref enum probe exercised and confirmed.
  if (DEFAULT_ONLY_ADD[r.type]) return `const c = ws.GetRange(${J(range)}).GetFormatConditions().${DEFAULT_ONLY_ADD[r.type]};${appearanceSetters(r)}`;
  return `const c = ws.GetRange(${J(range)}).GetFormatConditions().Add(${addArgs(r)});${r.rgb ? `
  c.SetFillColor(Api.CreateColorFromRGB(${r.rgb.join(', ')}));` : ''}`;
}

/** X-O01: the measured appearance setters, only for the keys the rule asked for (none -> the engine default, as before). */
function appearanceSetters(r) {
  const J = JSON.stringify;
  const k = new Set(r.idKeys || []);
  const out = [];
  if (k.has('barColor')) out.push(`c.SetBarColor(Api.CreateColorFromRGB(${r.barRgb.join(', ')}));`);
  if (k.has('barFillType')) out.push(`c.SetBarFillType(${J(r.barFillType)});`);
  if (k.has('direction')) out.push(`c.SetDirection(${J(r.direction)});`);
  if (k.has('showValue')) out.push(`c.SetShowValue(${J(r.showValue)});`);
  if (k.has('axisPosition')) out.push(`c.SetAxisPosition(${J(r.axisPosition)});`);
  if (k.has('iconSet')) out.push(`c.SetIconSet(${J(r.iconSet)});`);
  if (k.has('reverseOrder')) out.push(`c.SetReverseOrder(${J(r.reverseOrder)});`);
  if (k.has('showIconOnly')) out.push(`c.SetShowIconOnly(${J(r.showIconOnly)});`);
  if (k.has('stopColors')) out.push(`const stops = c.GetColorScaleCriteria();\n  ${r.stopRgbs.map((rgb, i) => `stops[${i}].SetColor(Api.CreateColorFromRGB(${rgb.join(', ')}));`).join('\n  ')}`);
  return out.length ? `\n  ${out.join('\n  ')}` : '';
}

function deleteBody(r, range) {
  const J = JSON.stringify;
  const own = CF_IDENTITY_FIELDS[r.type] && [...CF_IDENTITY_FIELDS[r.type], ...(r.idKeys || [])];
  const sameExpr = own
    ? `x.GetType() === ${J(r.type)}${own.map(k => ` && ${CF_IDENTITY_EXPRS[k] || `x.${CF_IDENTITY_GETTERS[k]}()`} === ${J(r[k])}`).join('')}
      && String(x.GetAppliesTo().GetAddress(false, false, 'xlA1', false)) === ${J(range)}`
    : `x.GetType() === ${J(r.type)} && ${r.type === 'xlExpression' ? 'true' : `x.GetOperator() === ${J(r.operator)}`} && String(x.GetFormula1()).replace(/^=/, '') === ${J(r.formula1)}${r.formula2 ? ` && String(x.GetFormula2()).replace(/^=/, '') === ${J(r.formula2)}` : ''}
      && (fill && typeof fill.GetRGB === 'function' ? fill.GetRGB() : null) === ${r.fill} && String(x.GetAppliesTo().GetAddress(false, false, 'xlA1', false)) === ${J(range)}`;
  return `const fc = ws.GetRange(${J(range)}).GetFormatConditions();
  const rgbOf = v => (v && typeof v.GetRGB === 'function' ? v.GetRGB() : null);
  const hits = [];
  for (let i = 1; i <= fc.GetCount(); i += 1) {
    const x = fc.GetItem(i);
    ${own ? '' : 'const fill = x.GetFillColor();\n    '}const same = ${sameExpr};
    if (same) hits.push(x);
  }
  if (hits.length !== 1) return { ok: false, stage: 'identity-at-dispatch', hits: hits.length };
  hits[0].Delete();`;
}

/** The measured Add calls: (type, operator, formula1[, formula2]) for cellValue; ('xlExpression', null, '=...'). */
function addArgs(r) {
  const J = JSON.stringify;
  if (r.type === 'xlExpression') return `${J(r.type)}, null, ${J(`=${r.formula1}`)}`;
  return [r.type, r.operator, r.formula1, ...(r.formula2 ? [r.formula2] : [])].map(x => J(x)).join(', ');
}

function verifyPrimitive(op, before, after) {
  const b = gate(before && before.primary, op);
  const a = gate(after && after.primary, op);
  if (!b.ok || !a.ok) return { ok: false, evidence: { reason: 'xlsx-cf-read-unobservable', before: b.ok || b.reason, after: a.ok || a.reason } };
  const add = op.intent === 'add_conditional_format';
  const ok = a.same.length === (add ? 1 : 0) && a.list.length === b.list.length + (add ? 1 : -1);
  return { ok, evidence: { sheet: op.selector.sheet, range: op.selector.range, identical: a.same.length, countBefore: b.list.length, countAfter: a.list.length, proof: 'fresh-live-range-rules' } };
}

/** L3: every OTHER rule the range reports unchanged (as a multiset), every cell. */
function verifyPreservation(op, before, after) {
  const rest = cells.verifyPreservation(op, before, after);
  if (!rest.ok && rest.evidence && rest.evidence.reason) return rest;
  const b = gate(before.primary, op);
  const a = gate(after.primary, op);
  const others = g => g.list.filter(x => !identical(x, op.rule, op.selector.range)).map(x => JSON.stringify(x)).sort();
  const kept = b.ok && a.ok && JSON.stringify(others(b)) === JSON.stringify(others(a));
  // X-L13/6: once tiered, report preservationScope FULL -- cfReads is keyed by the exact requested
  // (sheet,range) pair (xlsx-driver.cjs's own `cfAsked` loop), never a used-range scan, and the check above
  // was ALREADY scoped to the op's own range even below 45000 -- tiering reduces nothing. Below 45000 unchanged.
  const tiered = before.primary.tier === 'tiered';
  return { ok: rest.ok && kept, evidence: { ...rest.evidence, otherRulesKept: kept, ...(tiered ? { preservationScope: 'FULL' } : {}) } };
}

function wholeTaskChecks(op, finalObservation) {
  const g = gate(finalObservation && finalObservation.primary, op);
  const ok = g.ok && g.same.length === (op.intent === 'add_conditional_format' ? 1 : 0);
  return [{ name: `${op.selector.sheet}!${op.selector.range} ${op.intent === 'add_conditional_format' ? 'carries' : 'does not carry'} the rule`, ok, detail: ok ? null : { reason: g.reason || 'count' } }];
}

const windowsOf = () => [];
const liveExtras = op => ({ cfRanges: [{ sheet: op.selector.sheet, range: op.selector.range }] });

// the write-side enums, exported so the READ projection (cf-read.cjs) inverts the SAME tables instead of a copy
const CF_ENUMS = Object.freeze({ OPS1, OPS2, DATABAR_FILL, DATABAR_DIRECTION, DATABAR_AXIS, ICON_SETS });

module.exports = { CF_ENUMS, INTENTS, plan, classify, resolve, lower, verifyPrimitive, verifyPreservation, wholeTaskChecks, windowsOf, liveExtras };
