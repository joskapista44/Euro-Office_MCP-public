'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');

const { analyze, resolveRef, nameUses, refOnTarget } = require('../../src/adapters/xlsx/dependencies.cjs');
const formula = require('../../src/adapters/xlsx/formula.cjs');

// A #REF! is dropped from dependency COLLECTION as a whole token (the live part of the formula is still read), and the
// ONE named guard refOnTarget refuses a #REF! hosted on the sheets a write changes -- nothing else.
const ctx = { sheets: ['Data', 'Other'], names: [] };
const raw = extra => ({
  names: [], validations: { Data: [], Other: [] }, conditional: { Data: [], Other: [] },
  objects: { Data: { charts: [], pivots: [], names: [] }, Other: { charts: [], pivots: [], names: [] } }, ...extra,
});

test('token drop: every #REF! shape goes WHOLE, and no trailing cell is left to be misread as a same-sheet reference', () => {
  for (const f of ['=#REF!', '= #REF!*2', '=#REF!A1+1', '=#REF!$A$1:$B$2', '=SUM(#REF!A:A)', '=#REF!3:4']) {
    assert.deepEqual(resolveRef(f, 'Other', ctx), { sheets: [] }, f);
    assert.deepEqual(nameUses(f), { names: [] }, f);
  }
  // 'Data!#REF!': the #REF! goes, the sheet prefix still names Data -- kept as a (conservative) dependency, not dropped
  assert.deepEqual(resolveRef('=Data!#REF!', 'Other', ctx), { sheets: ['Data'] });
  assert.deepEqual(nameUses('=Data!#REF!'), { names: [] });
  assert.deepEqual(resolveRef('=Data!A1+#REF!', 'Other', ctx), { sheets: ['Data'] }, 'the live reference beside a #REF! is collected');
  assert.deepEqual(formula.sheetRefs('=#REF!A1+Data!B2').names, ['Data'], "'#REF!' is never a sheet named REF");
  // the token drop does not hide what is still unresolvable
  assert.deepEqual(resolveRef('=#REF+1', 'Other', ctx), { problem: 'broken-reference' }, 'a #REF without the ! is not the dropped token');
  assert.deepEqual(resolveRef('=Gone!A1+#REF!', 'Other', ctx), { problem: 'missing-sheet:Gone' });
});

test('analyze: a #REF! anywhere keeps the dependencies COMPLETE, and broken[] names where each one sits', () => {
  const cells = { 'Data\u0000C1': { f: '= #REF!*2' }, 'Other\u0000A1': { f: '=Data!B2+#REF!' } };
  const a = analyze(['Data', 'Other'], cells, raw({
    names: [{ name: 'Gone', ref: '#REF!$A$1', scopeSheet: 'Other' }, { name: 'WGone', ref: '#REF!', scopeSheet: null }],
    validations: { Data: [{ type: 'list', ranges: [{ r1: 0, c1: 4, r2: 0, c2: 4 }], formula1: '=#REF!', formula2: null }], Other: [] },
  }));
  assert.equal(a.dependencies.complete, true, JSON.stringify(a.dependencies.unobservable));
  const where = a.dependencies.broken.map(b => `${b.kind}:${b.host}`).sort();
  assert.deepEqual(where, ['formula:Data', 'formula:Other', 'name:Other', 'name:null', 'validation:Data']);
  assert.ok(a.dependencies.refs.some(r => r.kind === 'formula' && r.host === 'Other' && r.sheets.includes('Data')), 'the live Data!B2 is a dependency');
});

test('refOnTarget: refuses only a #REF! hosted on a target sheet, named apart from the reader-gap refusals', () => {
  const deps = { broken: [{ kind: 'formula', host: 'Data', name: 'C1', text: '= #REF!*2' }, { kind: 'name', host: null, name: 'WGone', text: '#REF!' }] };
  const r = refOnTarget(deps, ['data']);
  assert.equal(r.reason, 'xlsx-ref-on-target-unmeasured');
  assert.deepEqual(r.refs.map(x => x.name), ['C1']);
  assert.equal(refOnTarget(deps, ['Other']), null, 'a #REF! on ANOTHER sheet does not block');
  assert.equal(refOnTarget(deps, []), null, 'a workbook-level name sits on no sheet');
  assert.equal(refOnTarget({ broken: [] }, ['Data']), null);
  assert.equal(refOnTarget({}, ['Data']).refs, 'not-read', 'an analysis without broken[] is not read: fail-closed');
});
