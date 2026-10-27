'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const { commandBody } = require('../../src/adapters/docx/families/style-definition.cjs');

/**
 * GH #292 S6 (LIVE, docs/evidence/gh292-style-definition-probe-20261021.json): on a style ParaPr, SetJc('justify') stores
 * NOTHING -- the justify value is 'both'. The released define_style / modify_style accepted align 'justify' and sent
 * SetJc('justify') (docs/evidence/gh292-style-align-installed-release14-BUG-20261021.json: the style was created without
 * a jc). This runs the REAL command body in vm against a fake style and records the SetJc argument per requested align.
 */
function setJcArgFor(align, mode = 'define') {
  const calls = [];
  const style = { GetName: () => 'S', GetTextPr: () => ({}), GetParaPr: () => ({ SetJc: v => calls.push(v) }) };
  const doc = { GetStyle: () => (mode === 'define' ? { GetName: () => 'Normal' } : style), CreateStyle: () => style };
  const sandbox = { Api: { GetDocument: () => doc } };
  vm.createContext(sandbox);
  vm.runInContext(`(function(){ ${commandBody({ family: 'styleDefinition', action: 'styleDefinition', mode, name: 'S', fields: { paragraph: { align } } })} })()`, sandbox);
  return calls;
}

test('GH #292 S6: justify is written as SetJc("both"); left / right / center unchanged -- define and modify alike', () => {
  for (const mode of ['define', 'modify']) {
    assert.deepEqual(setJcArgFor('justify', mode), ['both'], mode);
    for (const a of ['left', 'right', 'center']) assert.deepEqual(setJcArgFor(a, mode), [a], `${mode} ${a}`);
  }
});
