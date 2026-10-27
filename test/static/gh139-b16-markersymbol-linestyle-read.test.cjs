'use strict';

/**
 * GH #139 B /16 (card an internal branch, kiosztó zero; design
 * docs/GH139-B16-CHART-STYLE-READ-DESIGN.md; LIVE MEASURED by a maintainer, card comment #ref --
 * write+session+reopen round trip, S1==S2 on every chart, 10/10 symbol + 11/11 dash token
 * bijective) -- the two new Bridge READ capabilities:
 *   chart.series.markerSymbol.read/v1  { symbol, size }
 *   chart.series.lineStyle.read/v1     { dash, smooth }
 *
 * This is the B11/B22-style offline WARRANT: a JS-aware brace match (strings skipped, the SAME
 * `closingBrace`/`spanFrom` test/static/b22-write-handler-identity.test.cjs already uses) extracts
 * each handler from the candidate block, never re-typed by hand, then runs it against a fake Api
 * built from a maintainer's OWN measured fixture rows (card #ref b16e-1.out), never a value this test
 * invented. No LIVE run, no Document Server -- this agent (maintainer) has no euro-broker access
 * (owner decision #ref, one principal; design doc section 0).
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { BRIDGE_CAPABILITIES } = require('../../src/drivers/eurooffice/ds-bridge.cjs');
const core = require('../../src/core/runtime-identity.cjs');

const DIR = path.join(__dirname, '..', '..', 'documentserver-patches', '9.3.4.60', 'euro-office-bridge-v1');
const SYMBOL_ID = 'chart.series.markerSymbol.read/v1';
const LINE_ID = 'chart.series.lineStyle.read/v1';
const blockOf = rev => fs.readFileSync(path.join(DIR, `rev${rev}-candidate`, 'bridge-block.js'), 'utf8');

/** the index of the `}` closing the `{` at `open`; strings ('..', "..", `..`) and escapes are skipped. */
function closingBrace(src, open) {
  let depth = 0; let quote = null;
  for (let k = open; k < src.length; k += 1) {
    const c = src[k];
    if (quote) { if (c === '\\') k += 1; else if (c === quote) quote = null; continue; }
    if (c === '"' || c === '\'' || c === '`') quote = c;
    else if (c === '{') depth += 1;
    else if (c === '}') { depth -= 1; if (depth === 0) return k; }
  }
  return -1;
}
function spanFrom(src, head) {
  const i = src.indexOf(head);
  if (i < 0) throw new Error('not found: ' + head);
  const open = src.indexOf('{', i + head.length - 1);
  const close = closingBrace(src, open);
  if (open < 0 || close < 0) throw new Error('unbalanced: ' + head);
  return src.slice(i, close + 1);
}
const handlerOf = (src, cap) => spanFrom(src, `"${cap}": function (args) {`);
const manifestEntryOf = (src, cap) => spanFrom(src, `{ id: "${cap}"`);
function toFn(span) {
  const body = span.slice(span.indexOf('function'));
  return new Function('return (' + body + ')')();
}

function makeApi(seriesByChart) {
  return {
    GetSheet(name) {
      if (name !== 'S') return null;
      return { GetAllCharts: () => Object.keys(seriesByChart).map(chartName => ({
        GetName: () => chartName,
        Chart: { chart: { plotArea: { charts: [{ series: seriesByChart[chartName] }] } } },
      })) };
    },
  };
}

test('both capability ids are declared in ds-bridge.cjs (readable by the driver)', () => {
  assert.equal(BRIDGE_CAPABILITIES.MARKER_SYMBOL_READ.id, SYMBOL_ID);
  assert.equal(BRIDGE_CAPABILITIES.LINE_STYLE_READ.id, LINE_ID);
  assert.equal(BRIDGE_CAPABILITIES.MARKER_SYMBOL_READ.valueType, 'object');
  assert.equal(BRIDGE_CAPABILITIES.LINE_STYLE_READ.valueType, 'object');
});

test('runtime-identity: the accepted list (DS9346 collapsed it to the single 9.3.5.0/16 entry -- the historical /15, /16 9.3.4.60 chain this test used to name is retired; GH #326/#327 widens again to /16 + /17)', () => {
  const ids = core.XLSX_RUNTIME_IDS;
  assert.deepEqual([...ids], ['9.3.5|euro-office-bridge-v1/9.3.5.0/16', '9.3.5|euro-office-bridge-v1/9.3.5.0/17']);
});

test('rev16-candidate exists and carries both NEW handlers; rev15 and earlier carry NEITHER (never guessed as already-served)', () => {
  const src16 = blockOf(16);
  assert.ok(src16.includes(`"${SYMBOL_ID}": function (args)`));
  assert.ok(src16.includes(`"${LINE_ID}": function (args)`));
  for (const rev of [6, 7, 8, 9, 10, 11, 12, 13, 14, 15]) {
    const src = blockOf(rev);
    // the bare id can legitimately appear in PROSE (rev15's own lineStyle.write/v1 manifest text
    // forward-references "a fresh chart.series.lineStyle.read/v1-shaped read ... not yet built, a
    // separate later capability") -- the handler/manifest-entry patterns are the real signal.
    assert.equal(src.includes(`"${SYMBOL_ID}": function (args)`), false, `rev${rev} must not have a ${SYMBOL_ID} handler`);
    assert.equal(src.includes(`{ id: "${SYMBOL_ID}"`), false, `rev${rev} must not advertise a ${SYMBOL_ID} manifest entry`);
    assert.equal(src.includes(`"${LINE_ID}": function (args)`), false, `rev${rev} must not have a ${LINE_ID} handler`);
    assert.equal(src.includes(`{ id: "${LINE_ID}"`), false, `rev${rev} must not advertise a ${LINE_ID} manifest entry`);
  }
});

test('manifest entries: kind READ, mutation NO, patchVersion /16 only', () => {
  const src = blockOf(16);
  for (const id of [SYMBOL_ID, LINE_ID]) {
    const entry = manifestEntryOf(src, id);
    assert.ok(entry.includes('kind: "READ"'), entry.slice(0, 120));
    assert.ok(entry.includes('mutation: "NO"'), entry.slice(0, 120));
    assert.ok(entry.includes('GH #139 B /16'), 'cites its own origin, not a guessed provenance');
  }
  assert.ok(src.includes('patchVersion: "euro-office-bridge-v1/9.3.4.60/16"'));
});

// ---- behavioural WARRANT: the extracted handlers, run against a maintainer's own measured rows ----
// (card #ref, b16e-1.out TABLE: symbol circle=0/dash=1/diamond=2/dot=3/none=4/plus=6/square=7/
// star=8/triangle=9/x=10, token 5 unmeasured; dash dash=0/dashDot=1/dot=2/lgDash=3/lgDashDot=4/
// lgDashDotDot=5/solid=6/sysDash=7/sysDashDot=8/sysDashDotDot=9/sysDot=10, 11/11 bijective; smooth
// true/false exactly, bar: field UNDEFINED.)
test('markerSymbol.read/v1: the full measured symbol table round-trips, token 5 is named; NO marker object fails closed, named (zero mechanism-probe #ref)', () => {
  const src = blockOf(16);
  const read = toFn(handlerOf(src, SYMBOL_ID));
  const TABLE = { circle: 0, dash: 1, diamond: 2, dot: 3, none: 4, plus: 6, square: 7, star: 8, triangle: 9, x: 10 };
  for (const [name, token] of Object.entries(TABLE)) {
    global.Api = makeApi({ c: [{ marker: { symbol: token, size: 7 } }] });
    assert.deepEqual(read({ sheet: 'S', chart: 'c', series: 0 }), { ok: true, value: { symbol: name, size: 7 } }, name);
  }
  global.Api = makeApi({ c: [{ marker: { symbol: 5, size: null } }] });
  assert.deepEqual(read({ sheet: 'S', chart: 'c', series: 0 }), { ok: true, value: { symbol: 'UNOBSERVABLE:symbol-token-unmeasured:5', size: null } });
  // a maintainer's #ref: a {symbol:null,size:null} with ok:true, for a series with NO marker object at
  // all (bar/pie), was byte-identical to a written-but-unset lineNormal series (ok:true too) --
  // the two states must not look the same. A series with NO marker object now fails closed, named.
  global.Api = makeApi({ c: [{}] });
  assert.deepEqual(read({ sheet: 'S', chart: 'c', series: 0 }), { ok: false, reason: 'no-marker-object' }, 'no marker object: fails closed, named, never a null-shaped ok:true');
});

test('lineStyle.read/v1: the full measured dash table round-trips (NOT DASH_ENUM declaration order -- solid is code 6), smooth is named when the field is not a boolean', () => {
  const src = blockOf(16);
  const read = toFn(handlerOf(src, LINE_ID));
  const TABLE = { dash: 0, dashDot: 1, dot: 2, lgDash: 3, lgDashDot: 4, lgDashDotDot: 5, solid: 6, sysDash: 7, sysDashDot: 8, sysDashDotDot: 9, sysDot: 10 };
  for (const [name, code] of Object.entries(TABLE)) {
    global.Api = makeApi({ c: [{ spPr: { ln: { prstDash: code } }, smooth: true }] });
    assert.deepEqual(read({ sheet: 'S', chart: 'c', series: 0 }), { ok: true, value: { dash: name, smooth: true } }, name);
  }
  // a maintainer's own NEG control: a bar series never carries `smooth` at all -- UNDEFINED, not false
  global.Api = makeApi({ bar: [{ spPr: {}, smooth: undefined }] });
  assert.deepEqual(read({ sheet: 'S', chart: 'bar', series: 0 }), { ok: true, value: { dash: null, smooth: 'UNOBSERVABLE:smooth-field-undefined' } });
  // never-written line: dash null, never defaulted to 'solid' (design decision (a))
  global.Api = makeApi({ c: [{ smooth: false }] });
  assert.deepEqual(read({ sheet: 'S', chart: 'c', series: 0 }), { ok: true, value: { dash: null, smooth: false } });
});

test('both handlers share the series-lookup prefix reasons (bad-args, sheet-not-found, chart-not-found, chart-type-not-single, series-index-out-of-range)', () => {
  const src = blockOf(16);
  for (const id of [SYMBOL_ID, LINE_ID]) {
    const read = toFn(handlerOf(src, id));
    global.Api = makeApi({});
    assert.equal(read({ sheet: 'S', chart: 'x' }).reason, 'bad-args', id);
    assert.equal(read({ sheet: 'Nope', chart: 'x', series: 0 }).reason, 'sheet-not-found', id);
    assert.equal(read({ sheet: 'S', chart: 'nope', series: 0 }).reason, 'chart-not-found', id);
    global.Api = makeApi({ c: [{ marker: { symbol: 0 } }] });
    assert.equal(read({ sheet: 'S', chart: 'c', series: 9 }).reason, 'series-index-out-of-range', id);
  }
});

test('driver wiring (xlsx-driver.cjs): markerSymbol scoped to the SAME types as markerStyle (scatter+lineNormal); lineStyle scoped to bar+lineNormal', () => {
  const driverSrc = fs.readFileSync(path.join(__dirname, '..', '..', 'src', 'drivers', 'eurooffice', 'xlsx-driver.cjs'), 'utf8');
  assert.ok(driverSrc.includes('BRIDGE_CAPABILITIES.MARKER_SYMBOL_READ'));
  assert.ok(driverSrc.includes('BRIDGE_CAPABILITIES.LINE_STYLE_READ'));
  const markerBlock = driverSrc.slice(driverSrc.indexOf("type === 'scatter' || type === 'lineNormal'"), driverSrc.indexOf('MARKER_SYMBOL_READ') + 40);
  assert.ok(markerBlock.includes('s.markerSymbol ='));
  const lineBlock = driverSrc.slice(driverSrc.indexOf("type === 'bar' || type === 'lineNormal'"), driverSrc.indexOf('LINE_STYLE_READ') + 20);
  assert.ok(lineBlock.includes('s.lineStyle ='));
});
