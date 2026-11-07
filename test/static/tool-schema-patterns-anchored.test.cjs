'use strict';

/**
 * Static gate: every `pattern` in every served tool inputSchema starts with '^' and ends with '$'. llama.cpp turns tool
 * schemas into a sampling grammar and refuses any other pattern ("JSON schema conversion failed: Pattern must start with
 * '^' and end with '$'", HTTP 400 on every request) -- 6 bare '^=' formula patterns made the server unusable there.
 * The anchored formula pattern must accept exactly what '^=' accepted, under the server's own check (RegExp, 'u').
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const { buildServerTools } = require('../../src/server/server-tools.cjs');
const W = require('../../src/server/write-tools.cjs');

const config = { schemaVersion: 2, nextcloud: { url: 'https://nc.example' }, broker: { url: 'http://127.0.0.1:9/api/euro-broker' } };

function patternsOf(node, path, out) {
  if (node && typeof node === 'object') {
    if (typeof node.pattern === 'string') out.push({ path, pattern: node.pattern });
    for (const [k, v] of Object.entries(node)) patternsOf(v, `${path}/${k}`, out);
  }
  return out;
}

test('every served tool-schema pattern is anchored at both ends', () => {
  const { tools } = buildServerTools({ config, configIdentity: 'fixture' });
  const all = Object.entries(tools).flatMap(([name, t]) => patternsOf(t.inputSchema, name, []));
  assert.ok(all.length > 100, `walked ${all.length} patterns`); // positive control: the walk reaches the step schemas
  assert.ok(all.some(p => p.path.includes('formula')), 'the formula patterns are in the walk');
  const bare = all.filter(p => !(p.pattern.startsWith('^') && p.pattern.endsWith('$')));
  assert.deepEqual(bare, []);
});

test('the anchored formula pattern accepts exactly what the bare ^= accepted', () => {
  const formula = W.STEP_SCHEMAS.write_array_formula.properties.formula.pattern;
  assert.notEqual(formula, '^=');
  const next = new RegExp(formula, 'u');
  const prev = new RegExp('^=', 'u');
  const samples = ['=', '=A1', '=SUM(A1:A3)', '="a\nb"', '=A1\r\n+B1', '=X Y', '=X ', '=\u{1F600}', '=' + 'A'.repeat(8191),
    '=' + 'x\n'.repeat(4000), '', 'A1', ' =A1', '\n=A1', '+A1', '==', '=\u0000'];
  for (const s of samples) assert.equal(next.test(s), prev.test(s), JSON.stringify(s.slice(0, 20)));
});

// The second llama.cpp stage: once anchored, the pattern is turned into a GBNF grammar, and that parser refuses an
// escape it does not know ("Failed to initialize samplers: failed to parse grammar"). The subset kept here: no group
// modifier ((?...)), no escape outside a [...] class, and inside a class only \uXXXX \n \r \t \\ \[ \].
function grammarUnsafe(pattern) {
  if (pattern.includes('(?')) return 'group modifier';
  let inClass = false;
  for (let i = 0; i < pattern.length; i++) {
    const c = pattern[i];
    if (c === '\\') {
      const n = pattern[i + 1];
      if (!inClass) return `escape \\${n} outside a class`;
      if (n === 'u') { if (!/^[0-9a-fA-F]{4}$/.test(pattern.slice(i + 2, i + 6))) return 'short \\u'; i += 5; continue; }
      if (!'nrt\\[]'.includes(n)) return `escape \\${n} inside a class`;
      i++; continue;
    }
    if (c === '[' && !inClass) inClass = true;
    else if (c === ']' && inClass) inClass = false;
  }
  return null;
}

test('every served tool-schema pattern stays inside the grammar-safe subset', () => {
  const { tools } = buildServerTools({ config, configIdentity: 'fixture' });
  const all = Object.entries(tools).flatMap(([name, t]) => patternsOf(t.inputSchema, name, []));
  const bad = all.map(p => ({ ...p, why: grammarUnsafe(p.pattern) })).filter(p => p.why);
  assert.deepEqual(bad, []);
  // negative control: the checker does flag each construct it exists for
  for (const p of ['^(?!=).+$', '^\\d$', '^\\$?a$', '^[\\s]$', '^a\\.b$']) assert.ok(grammarUnsafe(p), p);
  assert.equal(grammarUnsafe('^[^\\u0000-\\u001f\\[\\]\\\\]+$'), null);
});

test('the docx link target uses the shared, defined pattern (no undefined pattern from a load cycle)', () => {
  const D = require('../../src/server/docx-write-tools.cjs');
  const found = patternsOf(D.inputSchema ? D.inputSchema() : {}, 'docx', []).filter(p => p.pattern === W.LINK_URL_PATTERN);
  assert.equal(typeof W.LINK_URL_PATTERN, 'string');
  assert.ok(found.length >= 1, `docx inputSchema carries LINK_URL_PATTERN (${found.length})`);
});

test('each rewritten pattern accepts exactly what its escaped original accepted (RegExp, u)', () => {
  const rw = [
    ['^(?!=).+$', '^[^=\\n\\r\\u2028\\u2029].*$', ['a', '=a', 'a=', '', '\n', 'a\nb', ' ', 'x ', '\u{1F600}', ' ']],
    ['^([01]\\d|2[0-3]):[0-5]\\d(:[0-5]\\d)?$', '^([01][0-9]|2[0-3]):[0-5][0-9](:[0-5][0-9])?$', ['00:00', '23:59:59', '24:00', '9:00', '12:3', '١٢:٣٤']],
    ['^\\d{4}-\\d{2}-\\d{2}$', '^[0-9]{4}-[0-9]{2}-[0-9]{2}$', ['2026-11-01', '2026-1-01', '٢٠٢٦-١١-٠١', '20261101']],
    ['^[<>]-?[0-9]+(\\.[0-9]+)?$', '^[<>]-?[0-9]+([.][0-9]+)?$', ['>1', '<-1.5', '>1x5', '>1.', '=1']],
    ['^data:image\\/png;base64,[A-Za-z0-9+/]+=*$', '^data:image/png;base64,[A-Za-z0-9+/]+=*$', ['data:image/png;base64,QUJD', 'data:image/png;base64,', 'data:imageXpng;base64,QQ==']],
    ['^\\$?[A-Za-z]{1,3}\\$?[1-9][0-9]{0,6}$', '^[$]?[A-Za-z]{1,3}[$]?[1-9][0-9]{0,6}$', ['A1', '$A$1', 'AB$12', '$$A1', 'A0']],
  ];
  for (const [prev, next, samples] of rw) {
    const a = new RegExp(prev, 'u'); const b = new RegExp(next, 'u');
    for (const s of samples) assert.equal(b.test(s), a.test(s), `${next} on ${JSON.stringify(s)}`);
  }
  // the \s spelling: every code point classified the same as \s
  const ws = new RegExp(`^[${W.WS_NON_C0}]$`, 'u');
  for (let cp = 0; cp <= 0x10ffff; cp++) {
    if (cp >= 0xd800 && cp <= 0xdfff) continue;
    const ch = String.fromCodePoint(cp);
    if (cp < 0x20) continue; // C0 is covered by the \u0000-\u001f range next to it
    assert.equal(ws.test(ch), /^\s$/u.test(ch), `U+${cp.toString(16)}`);
  }
  // and the two whole patterns against their originals on edge strings
  const oldUrl = new RegExp('^(https?://[^\\s/?#\\u0000-\\u001f]+[^\\s\\u0000-\\u001f]*|mailto:[^\\s\\u0000-\\u001f]+)$', 'u');
  const newUrl = new RegExp(W.LINK_URL_PATTERN, 'u');
  for (const s of ['https://a.b/c?d#e', 'http://x', 'https:// x', 'https://a b', 'mailto:a@b', 'mailto:a b', 'https://a　', 'ftp://a', 'https://a\tb'])
    assert.equal(newUrl.test(s), oldUrl.test(s), JSON.stringify(s));
  const oldFont = new RegExp('^[^\\u0000-\\u001f\\u007f\\s]([^\\u0000-\\u001f\\u007f]{0,29}[^\\u0000-\\u001f\\u007f\\s])?$', 'u');
  const { tools } = buildServerTools({ config, configIdentity: 'fixture' });
  const fontPat = Object.values(tools).flatMap(t => patternsOf(t.inputSchema, '', [])).map(p => p.pattern).find(p => p.startsWith('^[^\\u0000-\\u001f\\u007f '));
  assert.ok(fontPat, 'FONT_NAME pattern found');
  const newFont = new RegExp(fontPat, 'u');
  for (const s of ['Arial', ' Arial', 'Arial ', 'Times New Roman', ' X', 'X　', 'A', '', 'Ab\u007f'])
    assert.equal(newFont.test(s), oldFont.test(s), JSON.stringify(s));
});
