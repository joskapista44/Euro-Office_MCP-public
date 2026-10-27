'use strict';

/**
 * v1.1.0 (B4): the public capability list (CAPABILITIES.md) must name exactly the read operations and write intents
 * the exposed tools offer. The truth is the exposed tools' own input schemas (what `tools/list` returns), not a
 * second hand-kept list. The doc lives at release/public/root/CAPABILITIES.md in the canonical tree and at the root of
 * the derived public tree; the test runs in both.
 */
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { buildServerTools } = require('../../src/server/server-tools.cjs');

const ROOT = path.join(__dirname, '..', '..');
const DOC = ['release/public/root/CAPABILITIES.md', 'CAPABILITIES.md'].map(rel => path.join(ROOT, rel)).find(p => fs.existsSync(p));

const { tools } = buildServerTools({ config: { broker: { url: 'http://127.0.0.1:1' }, playwright: { path: '/nonexistent' } }, configIdentity: 'test' });
const readOps = name => tools[name].inputSchema.properties.operation.enum;
const writeIntents = name => tools[name].inputSchema.properties.task.properties.steps.items.oneOf.map(i => i.properties.intent.const);
const FROM_SCHEMA = {
  xlsx_read: readOps('xlsx_read'),
  docx_read: readOps('docx_read'),
  xlsx_execute_task: writeIntents('xlsx_execute_task'),
  docx_execute_task: writeIntents('docx_execute_task'),
};

/** { tool: { count, names } } from "## `tool` -- N ..." headings and their "- `name`" bullets. */
function sections(text) {
  const out = {};
  let cur = null;
  for (const line of text.split('\n')) {
    const h = /^## `([a-z_]+)` -- (\d+) /.exec(line);
    if (h) { cur = out[h[1]] = { count: Number(h[2]), names: [] }; continue; }
    if (line.startsWith('## ')) { cur = null; continue; }
    const b = /^- `([a-z0-9_]+)`\s*$/.exec(line);
    if (b && cur) cur.names.push(b[1]);
  }
  return out;
}

const mismatches = (doc, truth) => {
  const errs = [];
  for (const [tool, names] of Object.entries(truth)) {
    const s = doc[tool];
    if (!s) { errs.push(`${tool}: no section`); continue; }
    const missing = names.filter(n => !s.names.includes(n));
    const extra = s.names.filter(n => !names.includes(n));
    if (missing.length) errs.push(`${tool}: missing ${missing.join(', ')}`);
    if (extra.length) errs.push(`${tool}: not offered ${extra.join(', ')}`);
    if (new Set(s.names).size !== s.names.length) errs.push(`${tool}: a name listed twice`);
    if (s.count !== names.length) errs.push(`${tool}: heading says ${s.count}, the schema offers ${names.length}`);
  }
  return errs;
};

test('the exposed tools are exactly the four capability tools plus product_status', () => {
  assert.deepEqual(Object.keys(tools).sort(), ['docx_execute_task', 'docx_read', 'product_status', 'xlsx_execute_task', 'xlsx_read']);
});

test('CAPABILITIES.md names exactly the read operations and write intents of the exposed tool schemas', () => {
  assert.ok(DOC, 'CAPABILITIES.md not found (release/public/root/ or the tree root)');
  assert.deepEqual(mismatches(sections(fs.readFileSync(DOC, 'utf8')), FROM_SCHEMA), []);
});

test('RED CONTROL: a doc with one intent dropped, one invented, or a wrong count fails (synthetic copies, not the real file)', () => {
  const real = fs.readFileSync(DOC, 'utf8');
  const victim = FROM_SCHEMA.docx_execute_task[0];
  assert.match(mismatches(sections(real.replace(`- \`${victim}\`\n`, '')), FROM_SCHEMA).join('|'), new RegExp(`missing ${victim}`));
  assert.match(mismatches(sections(real.replace(`- \`${victim}\`\n`, `- \`${victim}\`\n- \`invented_intent_q7\`\n`)), FROM_SCHEMA).join('|'), /not offered invented_intent_q7/);
  const n = FROM_SCHEMA.xlsx_read.length;
  assert.match(mismatches(sections(real.replace(`## \`xlsx_read\` -- ${n} `, `## \`xlsx_read\` -- ${n + 1} `)), FROM_SCHEMA).join('|'), /heading says/);
});
