'use strict';

/**
 * Product MCP entrypoint acceptance (owner decision, owner channel 2026-10-08, point 16 A-L).
 * Offline: a test config, no production credential, no editor. The DS handshake is exercised through the canonical
 * classifier with injected handshake results (H); the live probe is refused until the credential is provisioned.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const { spawn, spawnSync } = require('node:child_process');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');

const ROOT = path.join(__dirname, '..', '..');
const BIN = path.join(ROOT, 'bin', 'euro-office-mcp.cjs');
const { productStatus, dsSection } = require('../../src/server/product-status.cjs');
const { releaseIdentity } = require('../../src/server/identity.cjs');
const { validateConfig } = require('../../src/server/config.cjs');
const { TOOL_ROWS, capabilityInventory, exposedTools, exposureGate } = require('../../src/server/tools.cjs');
const { redact } = require('../../src/server/log.cjs');

// every temp dir is removed when the file's tests end (the suite runs often; /tmp is shared by the fleet)
const made = [];
const tmp = () => { const d = fs.mkdtempSync(path.join(os.tmpdir(), 'eo-mcp-test-')); made.push(d); return d; };
test.after(() => { for (const d of made) fs.rmSync(d, { recursive: true, force: true }); });
const writeConfig = (dir, obj) => { const f = path.join(dir, 'config.json'); fs.writeFileSync(f, typeof obj === 'string' ? obj : JSON.stringify(obj)); return f; };
// M11.2: schemaVersion 2 -- a non-secret broker ROUTE, no credential reference (the superseded `credential` key is refused)
const GOOD = Object.freeze({ schemaVersion: 2, nextcloud: { url: 'https://nc.test.invalid' }, broker: { url: 'http://127.0.0.1:9/api/euro-broker' }, log: { level: 'debug' } });
const NO_CRED = Object.freeze({ schemaVersion: 2, nextcloud: { url: 'https://nc.test.invalid' } });
const INIT = { jsonrpc: '2.0', id: 1, method: 'initialize', params: { protocolVersion: '2025-06-18', capabilities: {}, clientInfo: { name: 'test', version: '0' } } };

/** Spawn the server; `send` lines, collect stdout lines; resolves when the process exits. */
function serve(configFile, { env = {}, lines = [], endInput = true, signal = null, waitForLines = 0 } = {}) {
  return new Promise(resolve => {
    const state = tmp();
    const p = spawn(process.execPath, [BIN, '--config', configFile], { env: { ...process.env, XDG_STATE_HOME: state, ...env }, stdio: ['pipe', 'pipe', 'pipe'] });
    let out = ''; let err = '';
    p.stdout.on('data', d => {
      out += d;
      if (signal && out.split('\n').filter(Boolean).length >= waitForLines) { p.kill(signal); }
    });
    p.stderr.on('data', d => { err += d; });
    for (const l of lines) p.stdin.write(`${typeof l === 'string' ? l : JSON.stringify(l)}\n`);
    if (endInput) p.stdin.end();
    const timer = setTimeout(() => p.kill('SIGKILL'), 15000);
    p.on('exit', (code, sig) => { clearTimeout(timer); resolve({ code, sig, out, err, state, lines: out.split('\n').filter(Boolean) }); });
  });
}

test('A) --identity: the product identity lines; a source checkout says UNRELEASED, never a guessed hash', () => {
  const dir = tmp();
  const r = spawnSync(process.execPath, [BIN, '--identity', '--config', writeConfig(dir, GOOD)], { encoding: 'utf8', env: { ...process.env, XDG_STATE_HOME: dir } });
  assert.equal(r.status, 0, r.stderr);
  const kv = Object.fromEntries(r.stdout.trim().split('\n').map(l => [l.slice(0, l.indexOf('=')), l.slice(l.indexOf('=') + 1)]));
  assert.equal(kv.PRODUCT, 'Euro-Office_MCP');
  assert.equal(kv.SOURCE_COMMIT, 'UNRELEASED');
  assert.equal(kv.ARTIFACT_VERSION, 'UNRELEASED');
  assert.equal(kv.ARTIFACT_SHA256, 'NOT_RECORDED');
  assert.equal(kv.CONFIG_SCHEMA_VERSION, '2');
  assert.match(kv.CONFIG_IDENTITY, /^[0-9a-f]{64}$/);
  for (const k of ['DS_DOCUMENT_SERVER_VERSION', 'DS_PATCH_VERSION', 'DS_BRIDGE_CAPABILITIES']) assert.match(kv[k], /^UNAVAILABLE\(LIVE_PROBE_NOT_WIRED\)$/);
  assert.equal(kv.EXPOSURE_GATE, 'CAPABILITIES_RELEASED'); // M11.3b.2: exactly the two READ tools
});

test('A) an installed release: release.json (builder) + artifact.json (installer) are read, not computed', () => {
  const dir = tmp();
  const sha = 'a'.repeat(40); const art = 'b'.repeat(64);
  fs.writeFileSync(path.join(dir, 'release.json'), JSON.stringify({ PRODUCT: 'Euro-Office_MCP', SOURCE_COMMIT: sha, ARTIFACT_VERSION: `0.1.0-m1-${sha.slice(0, 12)}` }));
  fs.writeFileSync(path.join(dir, 'artifact.json'), JSON.stringify({ ARTIFACT_SHA256: art }));
  assert.deepEqual(releaseIdentity(dir), { PRODUCT: 'Euro-Office_MCP', RELEASED: true, SOURCE_COMMIT: sha, ARTIFACT_VERSION: `0.1.0-m1-${sha.slice(0, 12)}`, ARTIFACT_SHA256: art });
  fs.writeFileSync(path.join(dir, 'release.json'), JSON.stringify({ PRODUCT: 'Other', SOURCE_COMMIT: sha }));
  assert.equal(releaseIdentity(dir).SOURCE_COMMIT, 'UNRELEASED', 'a foreign release.json is not this product');
});

test('B-D) clean stdio startup, initialize, tools/list: exactly product_status + the two READ tools + xlsx_execute_task (M11.4b.1); stdout carries only protocol', async () => {
  const dir = tmp();
  const r = await serve(writeConfig(dir, GOOD), { lines: [INIT, { jsonrpc: '2.0', method: 'notifications/initialized' }, { jsonrpc: '2.0', id: 2, method: 'tools/list' }] });
  assert.equal(r.code, 0, r.err);
  assert.equal(r.lines.length, 2, 'one answer per request, nothing for the notification');
  const [init, list] = r.lines.map(l => JSON.parse(l));
  assert.equal(init.id, 1);
  assert.equal(init.result.protocolVersion, '2025-06-18');
  assert.equal(init.result.serverInfo.name, 'euro-office-product');
  assert.deepEqual(list.result.tools.map(t => t.name), ['product_status', 'xlsx_read', 'docx_read', 'xlsx_execute_task', 'docx_execute_task']);
  assert.match(r.err, /"msg":"starting"/, 'the log goes to stderr');
});

test('C) an unknown protocol version is answered with the server\'s own latest, not echoed', async () => {
  const dir = tmp();
  const r = await serve(writeConfig(dir, GOOD), { lines: [{ ...INIT, params: { ...INIT.params, protocolVersion: '1999-01-01' } }] });
  assert.equal(JSON.parse(r.lines[0]).result.protocolVersion, '2025-06-18');
});

test('E) product_status: identity, artifact, config schema, DS, exposure gate, dependencies -- and no secret', async () => {
  const dir = tmp();
  const r = await serve(writeConfig(dir, GOOD), { lines: [INIT, { jsonrpc: '2.0', id: 3, method: 'tools/call', params: { name: 'product_status', arguments: {} } }] });
  const res = JSON.parse(r.lines[1]).result;
  assert.equal(res.isError, false);
  const s = res.structuredContent;
  assert.equal(s.product, 'Euro-Office_MCP');
  assert.equal(s.sourceCommit, 'UNRELEASED');
  assert.deepEqual(Object.keys(s.artifact).sort(), ['released', 'sha256', 'version']);
  assert.equal(s.config.schemaVersion, 2);
  assert.equal(s.config.credentialModel, 'REQUESTER_OWN_VIA_BROKER');
  assert.equal(s.config.brokerRoute, 'CONFIGURED');
  assert.equal(s.ds.status, 'FAIL_CLOSED');
  assert.equal(s.exposureGate.state, 'CAPABILITIES_RELEASED');
  assert.deepEqual(s.tools.exposed, ['product_status', 'xlsx_read', 'docx_read', 'xlsx_execute_task', 'docx_execute_task']);
  assert.deepEqual(s.tools.withheld, []);
  assert.equal(s.exposureGate.writeReleaseGate, 'PASS');
  assert.ok(s.dependencies.node && s.dependencies.playwright);
  const text = JSON.stringify(res);
  for (const leak of ['127.0.0.1:9', 'api/euro-broker', 'keyName']) assert.equal(text.includes(leak), false, `no ${leak} in the answer`);
});

test('F) malformed config -> fail-closed startup: exit 78, nothing on stdout', async () => {
  const cases = [
    '{not json',
    JSON.stringify({ ...GOOD, schemaVersion: 1 }),
    JSON.stringify({ ...GOOD, credential: { keyName: 'k_ok', brokerUrl: 'http://127.0.0.1:9/' } }),
    JSON.stringify({ ...GOOD, broker: { url: 'https://broker.test.invalid/api/euro-broker' } }),
    JSON.stringify({ ...GOOD, extra: 1 }),
    JSON.stringify({ ...GOOD, nextcloud: { url: 'http://plain.invalid' } }),
    JSON.stringify({ ...GOOD, broker: { url: 'http://127.0.0.1:9/api/euro-broker', password: 'x' } }),
    JSON.stringify({ ...GOOD, nextcloud: { url: 'https://nc.invalid', apiToken: 'x' } }),
    JSON.stringify({ ...GOOD, policy: { exposeAll: true } }),
  ];
  for (const c of cases) {
    const dir = tmp();
    const r = await serve(writeConfig(dir, c), { lines: [INIT] });
    assert.equal(r.code, 78, c);
    assert.equal(r.out, '', `stdout empty: ${c}`);
    assert.match(r.err, /CONFIG_INVALID/);
  }
  const missing = await serve('/nonexistent/euro-office-mcp.json', { lines: [INIT] });
  assert.equal(missing.code, 78);
});

test('F) the validator refuses every secret-shaped key at any depth (no credential-reference exception since M11.2)', () => {
  for (const k of ['password', 'secret', 'token', 'apiKey', 'Authorization', 'cookie', 'privateKey']) {
    assert.equal(validateConfig({ ...GOOD, documentServer: { [k]: 'x' } }).ok, false, k);
  }
  assert.equal(validateConfig(GOOD).ok, true);
  assert.equal(validateConfig(NO_CRED).ok, true);
  assert.equal(validateConfig({ ...GOOD, broker: { url: 'http://127.0.0.1:9/api/euro-broker', keyName: 'k_ok' } }).ok, false, 'keyName');
});

test('G) missing broker route -> startup ok, every editor operation refused (DS probe fail-closed)', async () => {
  const s = await productStatus({ config: NO_CRED, configIdentity: 'x' });
  assert.equal(s.config.brokerRoute, 'MISSING');
  assert.deepEqual([s.ds.status, s.ds.reason], ['FAIL_CLOSED', 'BROKER_ROUTE_MISSING']);
  assert.deepEqual([(await productStatus({ config: GOOD, configIdentity: 'x' })).ds.reason], ['LIVE_PROBE_NOT_WIRED']);
  const dir = tmp();
  const r = await serve(writeConfig(dir, NO_CRED), { lines: [INIT, { jsonrpc: '2.0', id: 2, method: 'tools/call', params: { name: 'product_status' } }] });
  assert.equal(r.code, 0);
  assert.equal(JSON.parse(r.lines[1]).result.structuredContent.ds.reason, 'BROKER_ROUTE_MISSING');
});

test('H) DS handshake mismatch -> FAIL_CLOSED through the one canonical classifier; a matching manifest -> PASS', async () => {
  // DS9346 (owner D5=B, VALTAS, docs/DS9346-RUNTIME-RETIREMENT.md): the default manifest now carries the
  // ACCEPTED 9.3.5.0/16 identity -- the historical 9.3.4.60/5 baseline is retired.
  const manifest = (over = {}) => ({ ok: true, manifest: { bridgeVersion: '1.0.0', documentServerVersion: '9.3.5', patchVersion: 'euro-office-bridge-v1/9.3.5.0/16', capabilities: [{ id: 'textStyle.bold.read/v1', version: 1, kind: 'READ', mutation: 'NO' }], ...over } });
  const probe = h => () => ({ handshake: h });
  const cases = [
    [manifest({ patchVersion: 'euro-office-bridge-v1/9.3.4.60/2' }), 'PATCH_VERSION_MISMATCH'],
    // an engine version nobody has ever accepted (9.3.6) keeps testing the RUNTIME_VERSION_MISMATCH path.
    [manifest({ documentServerVersion: '9.3.6' }), 'RUNTIME_VERSION_MISMATCH'],
    [manifest({ bridgeVersion: '2.0.0' }), 'BRIDGE_VERSION_MISMATCH'],
    [{ ok: true, manifest: { bridgeVersion: '1.0.0' } }, 'MANIFEST_MALFORMED'],
    [{ ok: false, reason: 'MANIFEST_MISSING' }, 'MANIFEST_MISSING'],
    [null, 'MANIFEST_MALFORMED'],
  ];
  for (const [h, reason] of cases) {
    const s = await productStatus({ config: GOOD, configIdentity: 'x', dsProbe: probe(h) });
    assert.deepEqual([s.ds.status, s.ds.reason], ['FAIL_CLOSED', reason], JSON.stringify(h));
    assert.match(s.health, /^DEGRADED/);
  }
  const ok = await productStatus({ config: GOOD, configIdentity: 'x', dsProbe: probe(manifest()) });
  assert.equal(ok.ds.status, 'PASS');
  assert.equal(ok.ds.runtimeId, '9.3.5|euro-office-bridge-v1/9.3.5.0/16');
  assert.deepEqual(ok.ds.capabilities, ['textStyle.bold.read/v1@1']);
  const threw = await productStatus({ config: GOOD, configIdentity: 'x', dsProbe: () => { throw new Error('boom'); } });
  assert.deepEqual([threw.ds.status, threw.ds.reason], ['FAIL_CLOSED', 'PROBE_THREW'], 'a throwing probe never becomes PASS');
  assert.deepEqual(dsSection({ refused: 'X' }).status, 'FAIL_CLOSED');
});

test('I) stdout pollution = 0: every stdout line is a JSON-RPC 2.0 message, a bad line gets a JSON error', async () => {
  const dir = tmp();
  const r = await serve(writeConfig(dir, GOOD), { lines: [INIT, 'garbage{', { jsonrpc: '2.0', id: 9, method: 'nope' }, { jsonrpc: '2.0', id: 10, method: 'ping' }, { jsonrpc: '2.0', id: 11, method: 'tools/call', params: { name: 'missing_tool' } }] });
  assert.equal(r.code, 0);
  const msgs = r.lines.map(l => JSON.parse(l));
  for (const m of msgs) assert.equal(m.jsonrpc, '2.0');
  assert.equal(msgs.find(m => m.error && m.error.code === -32700).id, null);
  assert.equal(msgs.find(m => m.id === 9).error.code, -32601);
  assert.deepEqual(msgs.find(m => m.id === 10).result, {});
  assert.equal(msgs.find(m => m.id === 11).error.code, -32602);
});

test('I) static: nothing in src/server or bin writes to stdout except the protocol writer', () => {
  const files = [...fs.readdirSync(path.join(ROOT, 'src', 'server')).map(f => path.join(ROOT, 'src', 'server', f)), BIN];
  for (const f of files) {
    const code = fs.readFileSync(f, 'utf8').replace(/\/\*[\s\S]*?\*\//g, '').replace(/(^|[^:])\/\/.*$/gm, '$1');
    assert.equal(/console\.(log|info|debug|table|dir)\(/.test(code), false, `${path.basename(f)}: no console output`);
    const stdoutWrites = (code.match(/process\.stdout\.write\(/g) || []).length;
    if (f === BIN) assert.equal(stdoutWrites, 2, 'bin: only --identity and --check-config print to stdout (they do not serve)');
    else assert.equal(stdoutWrites, 0, `${path.basename(f)}: no direct stdout write`);
  }
});

test('J) graceful EOF: stdin closed -> clean exit 0 after answering', async () => {
  const dir = tmp();
  const r = await serve(writeConfig(dir, GOOD), { lines: [INIT] });
  assert.equal(r.code, 0);
  assert.match(r.err, /"reason":"stdin-eof"/);
});

test('K) SIGTERM: stop reading, exit 0', async () => {
  const dir = tmp();
  const r = await serve(writeConfig(dir, GOOD), { lines: [INIT], endInput: false, signal: 'SIGTERM', waitForLines: 1 });
  assert.equal(r.code, 0, `exit ${r.code} ${r.sig}`);
  assert.match(r.err, /"reason":"SIGTERM"/);
});

test('L) two agent UIDs: independent configs and state dirs, no cross-write', async () => {
  const a = tmp(); const b = tmp();
  const ra = await serve(writeConfig(a, GOOD), { lines: [INIT], env: { XDG_STATE_HOME: path.join(a, 'state') } });
  const rb = await serve(writeConfig(b, NO_CRED), { lines: [INIT], env: { XDG_STATE_HOME: path.join(b, 'state') } });
  assert.equal(ra.code, 0); assert.equal(rb.code, 0);
  const logOf = d => { const dir = path.join(d, 'state', 'euro-office-mcp'); const files = fs.readdirSync(dir); assert.equal(files.length, 1); return { text: fs.readFileSync(path.join(dir, files[0]), 'utf8'), mode: fs.statSync(dir).mode & 0o777 }; };
  const la = logOf(a); const lb = logOf(b);
  assert.equal(la.mode, 0o700); assert.equal(lb.mode, 0o700);
  const idA = JSON.parse(la.text.split('\n')[0]).f.configIdentity;
  const idB = JSON.parse(lb.text.split('\n')[0]).f.configIdentity;
  assert.notEqual(idA, idB);
  assert.equal(la.text.includes(idB), false); assert.equal(lb.text.includes(idA), false);
});

test('logging: secret-shaped keys and credential-shaped strings are redacted, long strings truncated', () => {
  const r = redact({ password: 'p', nested: { apiToken: 't' }, note: 'Authorization: Bearer abcdefghijklmnop', big: 'x'.repeat(1000) });
  assert.equal(r.password, '[REDACTED]'); assert.equal(r.nested.apiToken, '[REDACTED]');
  assert.equal(r.note.includes('abcdefghijklmnop'), false);
  assert.ok(r.big.length < 400);
});

test('tool inventory: derived from the adapters; product_status + the two READ tools + the two WRITE tools exposed (GH #228 proposed state); only the accepted DOCX intents are released routed capabilities', () => {
  const caps = capabilityInventory();
  assert.equal(caps.filter(r => r.FORMAT === 'DOCX').length, 38); // GH #268c: + set_image_size, set_image_wrap (routed + coded, NOT released); // GH #262: capabilityInventory() now also covers the 5 NEW_INTENTS (was 28, M4-migrated only -- a pre-existing gap this unit's own set_page_setup release exposed)
  assert.equal(caps.filter(r => r.FORMAT === 'XLSX').length, 49);
  assert.deepEqual(caps.filter(r => r.EXPOSE_IN_PRODUCT_MCP !== 'NO').map(r => r.TOOL).sort(), ['docx:add_bookmark', 'docx:add_comment_on_match', 'docx:add_endnote', 'docx:add_footnote', 'docx:add_link', 'docx:apply_list', 'docx:define_style', 'docx:delete_paragraph', 'docx:delete_table_column', 'docx:delete_table_row', 'docx:edit_footer', 'docx:edit_header', 'docx:format_table', 'docx:format_text', 'docx:indent_list_item', 'docx:insert_paragraph_after', 'docx:insert_paragraph_before', 'docx:insert_table_column_after', 'docx:insert_table_column_before', 'docx:insert_table_row_after', 'docx:insert_table_row_before', 'docx:merge_table_cells', 'docx:modify_style', 'docx:move_paragraph_after', 'docx:move_paragraph_before', 'docx:outdent_list_item', 'docx:remove_list', 'docx:replace_table_cell', 'docx:reply_comment', 'docx:resolve_comment', 'docx:rewrite_paragraph', 'docx:set_list_level', 'docx:set_list_type', 'docx:set_numbering', 'docx:set_page_setup', 'docx:split_table_cell', 'docx:set_image_size', 'docx:set_image_wrap'].sort());
  assert.deepEqual(caps.filter(r => r.EXPOSE_IN_PRODUCT_MCP !== 'NO').map(r => r.TOOL).sort(), ['docx:add_bookmark', 'docx:add_comment_on_match', 'docx:add_endnote', 'docx:add_footnote', 'docx:add_link', 'docx:apply_list', 'docx:define_style', 'docx:delete_paragraph', 'docx:delete_table_column', 'docx:delete_table_row', 'docx:edit_footer', 'docx:edit_header', 'docx:format_table', 'docx:format_text', 'docx:indent_list_item', 'docx:insert_paragraph_after', 'docx:insert_paragraph_before', 'docx:insert_table_column_after', 'docx:insert_table_column_before', 'docx:insert_table_row_after', 'docx:insert_table_row_before', 'docx:merge_table_cells', 'docx:modify_style', 'docx:move_paragraph_after', 'docx:move_paragraph_before', 'docx:outdent_list_item', 'docx:remove_list', 'docx:replace_table_cell', 'docx:reply_comment', 'docx:resolve_comment', 'docx:rewrite_paragraph', 'docx:set_list_level', 'docx:set_list_type', 'docx:set_numbering', 'docx:set_page_setup', 'docx:split_table_cell', 'docx:set_image_size', 'docx:set_image_wrap'].sort());
  assert.deepEqual(caps.filter(r => r.EXPOSE_IN_PRODUCT_MCP !== 'NO').map(r => r.TOOL).sort(), ['docx:add_bookmark', 'docx:add_comment_on_match', 'docx:add_endnote', 'docx:add_footnote', 'docx:add_link', 'docx:apply_list', 'docx:define_style', 'docx:delete_paragraph', 'docx:delete_table_column', 'docx:delete_table_row', 'docx:edit_footer', 'docx:edit_header', 'docx:format_table', 'docx:format_text', 'docx:indent_list_item', 'docx:insert_paragraph_after', 'docx:insert_paragraph_before', 'docx:insert_table_column_after', 'docx:insert_table_column_before', 'docx:insert_table_row_after', 'docx:insert_table_row_before', 'docx:merge_table_cells', 'docx:modify_style', 'docx:move_paragraph_after', 'docx:move_paragraph_before', 'docx:outdent_list_item', 'docx:remove_list', 'docx:replace_table_cell', 'docx:reply_comment', 'docx:resolve_comment', 'docx:rewrite_paragraph', 'docx:set_list_level', 'docx:set_list_type', 'docx:set_numbering', 'docx:set_page_setup', 'docx:split_table_cell', 'docx:set_image_size', 'docx:set_image_wrap'].sort());
  assert.deepEqual(caps.filter(r => r.EXPOSE_IN_PRODUCT_MCP !== 'NO').map(r => r.TOOL).sort(), ['docx:add_bookmark', 'docx:add_comment_on_match', 'docx:add_endnote', 'docx:add_footnote', 'docx:add_link', 'docx:apply_list', 'docx:define_style', 'docx:delete_paragraph', 'docx:delete_table_column', 'docx:delete_table_row', 'docx:edit_footer', 'docx:edit_header', 'docx:format_table', 'docx:format_text', 'docx:indent_list_item', 'docx:insert_paragraph_after', 'docx:insert_paragraph_before', 'docx:insert_table_column_after', 'docx:insert_table_column_before', 'docx:insert_table_row_after', 'docx:insert_table_row_before', 'docx:merge_table_cells', 'docx:modify_style', 'docx:move_paragraph_after', 'docx:move_paragraph_before', 'docx:outdent_list_item', 'docx:remove_list', 'docx:replace_table_cell', 'docx:reply_comment', 'docx:resolve_comment', 'docx:rewrite_paragraph', 'docx:set_list_level', 'docx:set_list_type', 'docx:set_numbering', 'docx:set_page_setup', 'docx:split_table_cell', 'docx:set_image_size', 'docx:set_image_wrap'].sort());
  assert.deepEqual(caps.filter(r => r.EXPOSE_IN_PRODUCT_MCP !== 'NO').map(r => r.TOOL).sort(), ['docx:add_bookmark', 'docx:add_comment_on_match', 'docx:add_endnote', 'docx:add_footnote', 'docx:add_link', 'docx:apply_list', 'docx:define_style', 'docx:delete_paragraph', 'docx:delete_table_column', 'docx:delete_table_row', 'docx:edit_footer', 'docx:edit_header', 'docx:format_table', 'docx:format_text', 'docx:indent_list_item', 'docx:insert_paragraph_after', 'docx:insert_paragraph_before', 'docx:insert_table_column_after', 'docx:insert_table_column_before', 'docx:insert_table_row_after', 'docx:insert_table_row_before', 'docx:merge_table_cells', 'docx:modify_style', 'docx:move_paragraph_after', 'docx:move_paragraph_before', 'docx:outdent_list_item', 'docx:remove_list', 'docx:replace_table_cell', 'docx:reply_comment', 'docx:resolve_comment', 'docx:rewrite_paragraph', 'docx:set_list_level', 'docx:set_list_type', 'docx:set_numbering', 'docx:set_page_setup', 'docx:split_table_cell', 'docx:set_image_size', 'docx:set_image_wrap'].sort());
  assert.deepEqual(caps.filter(r => r.EXPOSE_IN_PRODUCT_MCP !== 'NO').map(r => r.TOOL).sort(), ['docx:add_bookmark', 'docx:add_comment_on_match', 'docx:add_endnote', 'docx:add_footnote', 'docx:add_link', 'docx:apply_list', 'docx:define_style', 'docx:delete_paragraph', 'docx:delete_table_column', 'docx:delete_table_row', 'docx:edit_footer', 'docx:edit_header', 'docx:format_table', 'docx:format_text', 'docx:indent_list_item', 'docx:insert_paragraph_after', 'docx:insert_paragraph_before', 'docx:insert_table_column_after', 'docx:insert_table_column_before', 'docx:insert_table_row_after', 'docx:insert_table_row_before', 'docx:merge_table_cells', 'docx:modify_style', 'docx:move_paragraph_after', 'docx:move_paragraph_before', 'docx:outdent_list_item', 'docx:remove_list', 'docx:replace_table_cell', 'docx:reply_comment', 'docx:resolve_comment', 'docx:rewrite_paragraph', 'docx:set_list_level', 'docx:set_list_type', 'docx:set_numbering', 'docx:set_page_setup', 'docx:split_table_cell', 'docx:set_image_size', 'docx:set_image_wrap'].sort());
  assert.deepEqual(caps.filter(r => r.EXPOSE_IN_PRODUCT_MCP !== 'NO').map(r => r.TOOL).sort(), ['docx:add_bookmark', 'docx:add_comment_on_match', 'docx:add_endnote', 'docx:add_footnote', 'docx:add_link', 'docx:apply_list', 'docx:define_style', 'docx:delete_paragraph', 'docx:delete_table_column', 'docx:delete_table_row', 'docx:edit_footer', 'docx:edit_header', 'docx:format_table', 'docx:format_text', 'docx:indent_list_item', 'docx:insert_paragraph_after', 'docx:insert_paragraph_before', 'docx:insert_table_column_after', 'docx:insert_table_column_before', 'docx:insert_table_row_after', 'docx:insert_table_row_before', 'docx:merge_table_cells', 'docx:modify_style', 'docx:move_paragraph_after', 'docx:move_paragraph_before', 'docx:outdent_list_item', 'docx:remove_list', 'docx:replace_table_cell', 'docx:reply_comment', 'docx:resolve_comment', 'docx:rewrite_paragraph', 'docx:set_list_level', 'docx:set_list_type', 'docx:set_numbering', 'docx:set_page_setup', 'docx:split_table_cell', 'docx:set_image_size', 'docx:set_image_wrap'].sort());
  for (const r of [...TOOL_ROWS, ...caps]) for (const k of ['TOOL', 'SOURCE_MODULE', 'FORMAT', 'READ_WRITE', 'AUTH_REQUIRED', 'CONFIG_REQUIRED', 'CURRENT_ACCEPTANCE', 'EXPOSE_IN_PRODUCT_MCP', 'REASON']) assert.ok(r[k], `${r.TOOL}.${k}`);
  assert.deepEqual(exposedTools(), ['product_status', 'xlsx_read', 'docx_read', 'xlsx_execute_task', 'docx_execute_task']);
  assert.deepEqual(exposureGate().releasedCapabilityTools, ['xlsx_read', 'docx_read', 'xlsx_execute_task', 'docx_execute_task']);
  assert.deepEqual(exposureGate().writeReleasePreconditions, { 'M11.4a.1': 'CLOSED', 'M11.4a.2': 'CLOSED', 'M11.4a.3': 'CLOSED' });
  assert.deepEqual(exposureGate().releasePreconditions, { REQUESTER_AUTHORIZATION: 'PASS', CREDENTIAL_PATH: 'PASS' });
  // NEG: an intent the adapter does not route never appears
  assert.equal(caps.some(r => r.TOOL === 'xlsx:autofit_range'), false);
});

test('the inventory document is the generator output (no drift between doc and code)', () => {
  const gen = require('node:child_process').execFileSync(process.execPath, [path.join(ROOT, 'scripts', 'gen-tool-inventory.cjs')], { encoding: 'utf8' });
  assert.equal(fs.readFileSync(path.join(ROOT, 'docs', 'PRODUCT-MCP-TOOL-INVENTORY-20261008.md'), 'utf8'), gen);
});
