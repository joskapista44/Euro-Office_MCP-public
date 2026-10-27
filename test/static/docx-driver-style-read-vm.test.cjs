'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const { PROPERTY_QUERY, buildTargetedStyleQuery } = require('../../src/drivers/eurooffice/docx-driver.cjs');

/**
 * W-10 (style DEFINITIONS) -- the MANDATORY vm-test (the "W04-5 lesson" #ref). a maintainer's LIVE
 * run1 found the FIRST version of this block (getter-based: GetBold/GetFontSize/
 * GetColor/GetJc/GetIndLeft/GetSpacingAfter on the style's own GetTextPr()/GetParaPr()) never
 * read anything live -- every round came back verify-uncertain. The fix zero named: the style's
 * ToJSON() itself carries the readable rPr/pPr sub-fields, confirmed live AFTER A REOPEN
 * (W10-1c evidence R2, docs/evidence/docx-w10-1c-style-write-run2-20261007.json, + the W10-2
 * SetColor/SetSpacingAfter evidence):
 *   rPr.b = true | rPr.sz = <half-points> | rPr.color = { auto, r, g, b }
 *   pPr.jc = 'center' | pPr.ind.left = <twips> | pPr.spacing.after = <twips>
 *
 * This test runs the ACTUAL `PROPERTY_QUERY` string (imported from the driver, not retyped) via
 * `vm`, against a hand-built fake `Api`/`doc` whose style's `ToJSON()` returns a STRING with
 * EXACTLY this measured shape (a maintainer's point 2, #ref) -- RED on the old getter-based driver
 * (asserted explicitly below, against a copy of the removed code), GREEN on the fixed one.
 *
 * CORRECTED (LIVE run2): `rPr.sz` is in HALF-POINTS (the write-argument unit), but
 * the task-facing `text.size` field is PT -- the driver must divide by 2 on the way out, the
 * SAME way `indentLeft`/`spacingAfter` already divide by 20 (twips -> pt). The fixture below
 * uses a maintainer's own measured example (a requested size:26(pt) wrote rPr.sz=52(half-pt)) so the
 * expected value here is directly traceable to the LIVE finding, not an arbitrary number.
 */

function makeFakeApi({ styleToJsonIsString = true, tojsonPayload } = {}) {
  const defaultPayload = {
    name: 'W10VmStyle', basedOn: null, styleType: 'paragraph',
    rPr: { b: true, sz: 52, color: { auto: false, r: 34, g: 85, b: 170 } }, // 52 half-pt = 26pt (a maintainer's exact measured example, #ref)
    pPr: { jc: 'center', ind: { left: 240 }, spacing: { after: 120 } },
  };
  const payload = tojsonPayload === undefined ? defaultPayload : tojsonPayload;
  // GetBold/GetFontSize/GetColor/GetJc/GetIndLeft/GetSpacingAfter are DELIBERATELY ABSENT from
  // this fake TextPr/ParaPr -- the whole point of this fixture is to prove the driver no longer
  // depends on them (the LIVE-disproven getter surface).
  const styleObj = {
    GetName: () => 'W10VmStyle',
    ToJSON: () => (styleToJsonIsString ? JSON.stringify(payload) : payload),
    GetTextPr: () => ({}),
    GetParaPr: () => ({}),
  };
  const normalPayload = { name: 'Normal', basedOn: null, styleType: 'paragraph', rPr: {}, pPr: {} };
  const normalObj = {
    GetName: () => 'Normal',
    ToJSON: () => (styleToJsonIsString ? JSON.stringify(normalPayload) : normalPayload),
    GetTextPr: () => ({}),
    GetParaPr: () => ({}),
  };
  const allStyles = [styleObj, normalObj];
  const doc = {
    GetElementsCount: () => 0,
    GetElement: () => null,
    Search: () => [],
    GetAllBookmarksNames: () => [],
    GetBookmarkRange: () => null,
    GetFootnotesFirstParagraphs: () => [],
    GetEndNotesFirstParagraphs: () => [],
    GetCurrentFootEndnote: () => null,
    GetAllTables: () => [],
    GetSections: () => [],
    GetDefaultTextPr: () => null,
    GetAllStyles: () => allStyles,
    GetStyle: n => allStyles.find(s => s.GetName() === n) || null,
  };
  return { GetDocument: () => doc };
}

function runPropertyQuery(fakeApi) {
  const sandbox = { Api: fakeApi, result: undefined };
  vm.createContext(sandbox);
  const raw = vm.runInContext(`(function(){ ${PROPERTY_QUERY} })()`, sandbox);
  return JSON.parse(JSON.stringify(raw));
}

test('vm: the REAL PROPERTY_QUERY reads a style via its ToJSON() (a STRING) -- the measured rPr/pPr shape, never the disproven getters', () => {
  const r = runPropertyQuery(makeFakeApi({ styleToJsonIsString: true }));
  assert.equal(r.error, undefined, `PROPERTY_QUERY threw: ${r.error}`);
  const st = r.document.styles.W10VmStyle;
  assert.ok(st, 'the style must be keyed by its NAME, not a styleId');
  assert.deepEqual(st.identity, { name: 'W10VmStyle', basedOn: null, next: null, link: null, styleType: 'paragraph', basedOnName: null, nextName: null, linkName: null });
  assert.deepEqual(st.text, { bold: true, size: 26, color: '2255AA', italic: 'inherited', underline: 'inherited', strike: 'inherited', fontFamily: 'inherited' }); // GH #292: the further keys, absent here
  assert.deepEqual(st.paragraph, { align: 'center', indentLeft: 12, spacingAfter: 6, spacingBefore: 'inherited', indentRight: 'inherited', indentFirstLine: 'inherited', keepNext: 'inherited', keepLines: 'inherited', pageBreakBefore: 'inherited' }); // 240/20, 120/20 twips->pt
});

test('vm: the SAME fixture with ToJSON returning a plain OBJECT (not a string) reads identically -- the W07-6 guard accepts either shape', () => {
  const r = runPropertyQuery(makeFakeApi({ styleToJsonIsString: false }));
  assert.deepEqual(r.document.styles.W10VmStyle.text, { bold: true, size: 26, color: '2255AA', italic: 'inherited', underline: 'inherited', strike: 'inherited', fontFamily: 'inherited' });
});

test('vm: color.auto === true reads "auto", never a guessed hex', () => {
  const r = runPropertyQuery(makeFakeApi({ tojsonPayload: { name: 'W10VmStyle', basedOn: null, styleType: 'paragraph', rPr: { color: { auto: true, r: 0, g: 0, b: 0 } }, pPr: {} } }));
  assert.equal(r.document.styles.W10VmStyle.text.color, 'auto');
});

test('vm, NEGATIVE: a key ABSENT from rPr/pPr reads "inherited" -- never a guessed default, never UNOBSERVABLE (a read that succeeded is not a read failure)', () => {
  const r = runPropertyQuery(makeFakeApi({ tojsonPayload: { name: 'W10VmStyle', basedOn: null, styleType: 'paragraph', rPr: {}, pPr: {} } }));
  assert.deepEqual(r.document.styles.W10VmStyle.text, { bold: 'inherited', size: 'inherited', color: 'inherited', italic: 'inherited', underline: 'inherited', strike: 'inherited', fontFamily: 'inherited' });
  assert.deepEqual(r.document.styles.W10VmStyle.paragraph, { align: 'inherited', indentLeft: 'inherited', spacingAfter: 'inherited', spacingBefore: 'inherited', indentRight: 'inherited', indentFirstLine: 'inherited', keepNext: 'inherited', keepLines: 'inherited', pageBreakBefore: 'inherited' });
});

test('vm, NEGATIVE: rPr/pPr themselves absent from ToJSON reads "inherited" too (not a read failure -- the TOP-level shape still parsed)', () => {
  const r = runPropertyQuery(makeFakeApi({ tojsonPayload: { name: 'W10VmStyle', basedOn: null, styleType: 'paragraph' } }));
  assert.deepEqual(r.document.styles.W10VmStyle.text, { bold: 'inherited', size: 'inherited', color: 'inherited', italic: 'inherited', underline: 'inherited', strike: 'inherited', fontFamily: 'inherited' });
});

test('vm, NEGATIVE: a style with no ToJSON method at all reads every field UNOBSERVABLE (a genuine read failure, never "inherited")', () => {
  const api = makeFakeApi();
  const doc = api.GetDocument();
  const st = doc.GetStyle('W10VmStyle');
  delete st.ToJSON;
  const r = runPropertyQuery(api);
  assert.equal(r.error, undefined);
  const fields = r.document.styles.W10VmStyle;
  for (const v of Object.values(fields.text)) assert.match(v, /^UNOBSERVABLE:/);
  for (const v of Object.values(fields.paragraph)) assert.match(v, /^UNOBSERVABLE:/);
  assert.equal(fields.identity, null);
});

test('vm, NEGATIVE: ToJSON returns a non-JSON string (parse throws) reads every field UNOBSERVABLE, never throws out of the query', () => {
  const api = makeFakeApi();
  const doc = api.GetDocument();
  const st = doc.GetStyle('W10VmStyle');
  st.ToJSON = () => 'not valid json {{{';
  const r = runPropertyQuery(api);
  assert.equal(r.error, undefined);
  assert.match(r.document.styles.W10VmStyle.text.bold, /^UNOBSERVABLE:/);
});

test('vm, NEGATIVE: doc.GetAllStyles absent reads an empty styles map, never throws', () => {
  const api = makeFakeApi();
  const doc = api.GetDocument();
  delete doc.GetAllStyles;
  const r = runPropertyQuery(api);
  assert.equal(r.error, undefined);
  assert.deepEqual(r.document.styles, {});
});

test('RED/GREEN: the OLD getter-based read (GetBold/GetFontSize/... on GetTextPr()/GetParaPr()) reads everything UNOBSERVABLE on THIS fixture -- reproducing zero\'s exact LIVE finding that the getter surface never worked', () => {
  // the fixture's GetTextPr()/GetParaPr() deliberately return {} (no Get* methods at all) --
  // exactly what a maintainer's LIVE run measured. A hand-reduced copy of the REMOVED getter logic:
  const read1s = (obj, n) => (!obj || typeof obj[n] !== 'function' ? 'UNOBSERVABLE:no-such-getter' : obj[n]());
  const api = makeFakeApi();
  const st = api.GetDocument().GetStyle('W10VmStyle');
  const tp = st.GetTextPr();
  const pp = st.GetParaPr();
  assert.equal(read1s(tp, 'GetBold'), 'UNOBSERVABLE:no-such-getter');
  assert.equal(read1s(tp, 'GetFontSize'), 'UNOBSERVABLE:no-such-getter');
  assert.equal(read1s(pp, 'GetJc'), 'UNOBSERVABLE:no-such-getter');
  // ... which is exactly why DEFINE/MODIFY came back verify-uncertain in LIVE run1: EVERY
  // requested field read UNOBSERVABLE, so verifyPrimitive's unobservableRequestedFields()
  // was never empty, no matter what was actually written.
});

test('RED/GREEN: the size unit mismatch (LIVE run2) -- a raw half-point rPr.sz compared against a pt-based request never matches', () => {
  const r = runPropertyQuery(makeFakeApi());
  const sizeNow = r.document.styles.W10VmStyle.text.size;
  assert.equal(sizeNow, 26, 'the FIXED driver must externalize pt, not the raw half-point rPr.sz');
  // the bug this reproduces: comparing the UN-halved raw value (52) against the REQUESTED pt
  // value (26) -- exactly the mismatch that made every text.size round uncertain in LIVE run1/2.
  assert.notEqual(sizeNow, 52, 'the raw half-point value must never leak into the pt-based field');
});

/**
 * W-10 (LIVE run3): doc.GetAllStyles() is MEASURED to NEVER include a style just
 * created via doc.CreateStyle(name, 'paragraph') -- in the same session, after a save, and
 * after a reopen -- even though doc.GetStyle(name) finds it every time. This fake models
 * EXACTLY that asymmetry: 'Normal' is in BOTH GetAllStyles() and GetStyle(); 'FreshCustom' is
 * ONLY in GetStyle() -- GetAllStyles() never lists it, reproducing the measured bug precisely.
 */
function makeFakeApiWithBlindCustomStyle() {
  const normalPayload = { name: 'Normal', basedOn: null, styleType: 'paragraph', rPr: { b: false }, pPr: {} };
  const normalObj = { GetName: () => 'Normal', ToJSON: () => JSON.stringify(normalPayload) };
  const customPayload = { name: 'FreshCustom', basedOn: null, styleType: 'paragraph', rPr: { b: true }, pPr: {} };
  const customObj = { GetName: () => 'FreshCustom', ToJSON: () => JSON.stringify(customPayload) };
  const doc = {
    GetElementsCount: () => 0,
    GetElement: () => null,
    Search: () => [],
    GetAllBookmarksNames: () => [],
    GetBookmarkRange: () => null,
    GetFootnotesFirstParagraphs: () => [],
    GetEndNotesFirstParagraphs: () => [],
    GetCurrentFootEndnote: () => null,
    GetAllTables: () => [],
    GetSections: () => [],
    GetDefaultTextPr: () => null,
    GetAllStyles: () => [normalObj], // 'FreshCustom' is DELIBERATELY absent here
    // W-10-7 (LIVE run4): doc.GetStyle(<unknown name>) was MEASURED to NOT return
    // null -- it falls back to the 'Normal' style object (own diagnostic: a NEG_CONTROL query
    // for a never-created name came back {found:true, identity:{name:'Normal'}}). This fixture
    // reproduces EXACTLY that fallback -- any name other than 'Normal'/'FreshCustom' resolves to
    // normalObj, not null.
    GetStyle: n => ({ Normal: normalObj, FreshCustom: customObj }[n] || normalObj),
  };
  return { GetDocument: () => doc };
}

function runTargetedQuery(fakeApi, names) {
  const sandbox = { Api: fakeApi, result: undefined };
  vm.createContext(sandbox);
  const raw = vm.runInContext(`(function(){ ${buildTargetedStyleQuery(names)} })()`, sandbox);
  return JSON.parse(JSON.stringify(raw));
}

test('vm: buildTargetedStyleQuery finds a style doc.GetAllStyles() never lists -- reproducing zero\'s exact LIVE finding, and PROPERTY_QUERY alone (via GetAllStyles) does NOT find it', () => {
  const api = makeFakeApiWithBlindCustomStyle();
  const viaPropertyQuery = runPropertyQuery(api);
  assert.equal(viaPropertyQuery.document.styles.FreshCustom, undefined, 'PROPERTY_QUERY (GetAllStyles-based) must NOT see it -- this IS the bug being reproduced');
  assert.ok(viaPropertyQuery.document.styles.Normal, 'the built-in control must still be visible via PROPERTY_QUERY');
  const viaTargeted = runTargetedQuery(api, ['FreshCustom']);
  assert.equal(viaTargeted.error, undefined, `buildTargetedStyleQuery threw: ${viaTargeted.error}`);
  assert.ok(viaTargeted.FreshCustom, 'the targeted query MUST find it via doc.GetStyle(name)');
  assert.equal(viaTargeted.FreshCustom.text.bold, true);
});

test('RED/GREEN (LIVE run4): a name that was never created reads null, never the Normal-fallback style -- doc.GetStyle(<unknown>) returning Normal must not be read as "found"', () => {
  const api = makeFakeApiWithBlindCustomStyle();
  const r = runTargetedQuery(api, ['DoesNotExist']);
  assert.equal(r.error, undefined);
  assert.equal(r.DoesNotExist, null, 'without the GetName()===name guard this would read the Normal style data instead of null');
});

test('RED/GREEN: reverting the GetName()===name guard reproduces the exact bug -- a never-created name resolves to the Normal style, exactly as LIVE run4 measured', () => {
  // a hand-reduced copy of the PRE-FIX styleOf (bare `if (!st) return null;` on the raw GetStyle result)
  const sandbox = { Api: makeFakeApiWithBlindCustomStyle(), result: undefined };
  vm.createContext(sandbox);
  const preFixQuery = `
try {
  const doc = Api.GetDocument();
  const styleOf = (name) => {
    const st = doc.GetStyle(name);
    if (!st) return null;
    const tj = JSON.parse(st.ToJSON());
    return { name, identity: { name: tj.name, basedOn: tj.basedOn, styleType: tj.styleType } };
  };
  return { DoesNotExist: styleOf('DoesNotExist') };
} catch (err) { return { error: String(err) }; }
`;
  const raw = vm.runInContext(`(function(){ ${preFixQuery} })()`, sandbox);
  const r = JSON.parse(JSON.stringify(raw));
  assert.equal(r.error, undefined);
  assert.equal(r.DoesNotExist.identity.name, 'Normal', 'reproducing the measured bug: the pre-fix read treats the Normal fallback as the requested style');
});

test('vm: buildTargetedStyleQuery is injection-safe for a name containing a quote or backtick', () => {
  const api = makeFakeApiWithBlindCustomStyle();
  for (const weirdName of ["Weird'Name", 'Weird"Name', 'Weird`Name']) {
    assert.doesNotThrow(() => new vm.Script(`(function(){${buildTargetedStyleQuery([weirdName])}})`), weirdName);
    const r = runTargetedQuery(api, [weirdName]);
    assert.equal(r.error, undefined, weirdName);
    assert.equal(r[weirdName], null, weirdName); // not found, but no syntax break either
  }
});

// GH #292 S6 (LIVE, docs/evidence/gh292-style-definition-probe-20261021.json): a style ParaPr's ToJSON reads SetJc('left')
// as 'start', SetJc('right') as 'end', SetJc('both') as 'both' (the saved w:jc is left / right / both). The task-facing
// align value is left | right | center | justify -- both reads normalize to it; an unknown raw value passes through.
test('GH #292 S6: both style reads normalize the measured jc names (start / end / both) to left / right / justify; center and an unknown value pass through', () => {
  for (const [raw, want] of [['start', 'left'], ['end', 'right'], ['both', 'justify'], ['center', 'center'], ['left', 'left'], ['right', 'right'], ['distribute', 'distribute']]) {
    const api = makeFakeApi({ tojsonPayload: { name: 'W10VmStyle', basedOn: null, styleType: 'paragraph', rPr: {}, pPr: { jc: raw } } });
    assert.equal(runPropertyQuery(api).document.styles.W10VmStyle.paragraph.align, want, `PROPERTY_QUERY ${raw}`);
    assert.equal(runTargetedQuery(api, ['W10VmStyle']).W10VmStyle.paragraph.align, want, `buildTargetedStyleQuery ${raw}`);
  }
});

// GH #292 S1 / S2: the measured ToJSON shapes of the further style keys (docs/evidence/gh292-style-definition-probe-20261021.json)
test('GH #292 S1 / S2: both style reads read italic / underline / strike / font and space-before / right / first-line (hanging) indent / keep / page-break keys', () => {
  const payload = { name: 'W10VmStyle', basedOn: null, styleType: 'paragraph',
    rPr: { i: true, u: true, strike: true, rFonts: { ascii: 'Courier New', cs: 'Courier New', eastAsia: 'Courier New', hAnsi: 'Courier New' } },
    pPr: { spacing: { before: 240, beforeAutoSpacing: 'off' }, ind: { right: 720, firstLine: -360 }, keepNext: true, keepLines: true, pageBreakBefore: true } };
  for (const read of [api => runPropertyQuery(api).document.styles.W10VmStyle, api => runTargetedQuery(api, ['W10VmStyle']).W10VmStyle]) {
    const st = read(makeFakeApi({ tojsonPayload: payload }));
    assert.deepEqual(st.text, { bold: 'inherited', size: 'inherited', color: 'inherited', italic: true, underline: true, strike: true, fontFamily: 'Courier New' });
    assert.deepEqual(st.paragraph, { align: 'inherited', indentLeft: 'inherited', spacingAfter: 'inherited', spacingBefore: 12, indentRight: 36, indentFirstLine: -18, keepNext: true, keepLines: true, pageBreakBefore: true });
  }
});

// GH #292 S4 (LIVE, docs/evidence/gh292s4-next-link-internal-probe-20261022.json): Heading 1's ToJSON carries next '886' and link '837' -- styleIds,
// resolved to names like basedOn; an id no read saw is named UNOBSERVABLE, never guessed
test('GH #292 S4: both style reads resolve next / link styleIds to names; an unseen id reads UNOBSERVABLE; absent reads null', () => {
  const mk = (name, payload) => ({ GetName: () => name, ToJSON: () => JSON.stringify(payload), GetTextPr: () => ({}), GetParaPr: () => ({}) });
  const styles = [
    mk('Heading 1', { name: 'Heading 1', styleId: '827', basedOn: '886', next: '886', link: '837', styleType: 'paragraphStyle' }),
    mk('Normal', { name: 'Normal', styleId: '886', basedOn: null, styleType: 'paragraphStyle' }),
    mk('Heading 1 Char', { name: 'Heading 1 Char', styleId: '837', basedOn: null, link: '827', styleType: 'characterStyle' }),
    mk('Orphan', { name: 'Orphan', styleId: '900', basedOn: null, next: '999', styleType: 'paragraphStyle' }),
  ];
  const doc = { GetElementsCount: () => 0, GetElement: () => null, Search: () => [], GetBookmarkRange: () => null, GetFootnotesFirstParagraphs: () => [],
    GetEndNotesFirstParagraphs: () => [], GetCurrentFootEndnote: () => null, GetAllTables: () => [], GetSections: () => [], GetDefaultTextPr: () => null,
    GetAllStyles: () => styles, GetStyle: n => styles.find(s => s.GetName() === n) || null };
  const api = { GetDocument: () => doc };
  const pick = id => ({ next: id.next, link: id.link, nextName: id.nextName, linkName: id.linkName });
  for (const [label, read] of [['PROPERTY_QUERY', n => runPropertyQuery(api).document.styles[n]], ['targeted', n => runTargetedQuery(api, [n])[n]]]) {
    assert.deepEqual(pick(read('Heading 1').identity), { next: '886', link: '837', nextName: 'Normal', linkName: 'Heading 1 Char' }, label);
    assert.deepEqual(pick(read('Heading 1 Char').identity), { next: null, link: '827', nextName: null, linkName: 'Heading 1' }, label);
    assert.deepEqual(pick(read('Normal').identity), { next: null, link: null, nextName: null, linkName: null }, label);
    assert.equal(read('Orphan').identity.nextName, 'UNOBSERVABLE:next-style-not-enumerated', label);
  }
});
