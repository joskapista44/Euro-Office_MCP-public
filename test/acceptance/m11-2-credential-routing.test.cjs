'use strict';

/**
 * M11.2 -- requester-own credential routing, OFFLINE acceptance (owner GO, owner channel 2026-10-09).
 *
 * The product runtime gets an editor config ONLY through the broker config source: the agent's own token goes to the
 * broker op `product_ocs_config`, which uses the requester's OWN Nextcloud credential. A fake broker (fetch seam)
 * stands in for the maintainer broker; a fake Playwright stands in for the editor. Proven here:
 *   - zero / maintainer requester fixtures: each request carries ONLY its own token; the body carries no identity;
 *   - cross-credential selection is impossible from the product side; no retry, no fallback, one request per call;
 *   - every broker answer maps onto one controlled class (no throw, no secret in any outcome);
 *   - an edit-capable config is ACCEPTED for READ, and the READ path (open + readLive + close) dispatches nothing and
 *     saves nothing;
 *   - the product runtime never constructs the legacy Basic-auth config source.
 */

const test = require('node:test');
const { after } = test;
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const os = require('node:os');

const { HANDSHAKE_BODY } = require('../../src/core/runtime-identity.cjs');
const { ERRORS, validateTarget, createBrokerConfigSource } = require('../../src/server/broker-config-source.cjs');
const { ConfigSourceError } = require('../../src/drivers/eurooffice/config-source.cjs');
const docx = require('../../src/drivers/eurooffice/docx-driver.cjs');
const { exposedTools, TOOL_ROWS } = require('../../src/server/tools.cjs');

const ROOT = path.join(__dirname, '..', '..');
const BROKER = 'http://127.0.0.1:9/api/euro-broker';
const TOKENS = Object.freeze({ zero: 'tok-zero-FIXTURE-0000', maintainer: 'tok-maintainer-FIXTURE-1111' });
const NC_UID_BY_TOKEN = Object.freeze({ [TOKENS.zero]: 'zero', [TOKENS.maintainer]: 'maintainer' });
const FILE = '1302148';

/** A config with the measured LIVE key structure (product_ocs_config, 2026-10-09) and synthetic values. */
function editConfig(ncUid, fileId = FILE) {
  return {
    document: { fileType: 'docx', key: `KEY-${fileId}`, permissions: { edit: true, protect: true, changeHistory: true }, title: 'fixture.docx',
      url: 'https://ds.example/download', referenceData: { fileKey: fileId, instanceId: 'instfixture' }, info: { favorite: false } },
    documentType: 'word',
    editorConfig: { lang: 'hu', region: 'hu-HU', callbackUrl: 'https://nc.example/callback', user: { id: `instfixture_${ncUid}`, name: ncUid, image: '' },
      createUrl: 'https://nc.example/create', customization: { goback: { url: 'https://nc.example/' }, compactHeader: true, feedback: false, uiTheme: 'default' } },
    documentServerUrl: 'https://ds.example/', iat: 1, exp: 2, token: 'ds-jwt-FIXTURE',
  };
}

/** A fake broker: answers by the bearer token (as the real one: requester = token owner). Records every request. */
function fakeBroker(answer = null) {
  const calls = [];
  const fetchImpl = async (url, init) => {
    const bearer = String(init.headers.Authorization || '').replace(/^Bearer /, '');
    const body = JSON.parse(init.body);
    calls.push({ url, bearer, body });
    if (answer) return answer({ bearer, body });
    const uid = NC_UID_BY_TOKEN[bearer];
    if (!uid) return { status: 401, json: async () => ({ error: 'unauthorized' }) };
    return { status: 200, json: async () => ({ ok: true, status: 200, config: editConfig(uid, body.fileId) }) };
  };
  return { calls, fetchImpl };
}

const source = (who, broker, extra = {}) => createBrokerConfigSource({ brokerUrl: BROKER, readToken: () => TOKENS[who], fetchImpl: broker.fetchImpl, ...extra });
const reply = (status, json) => () => ({ status, json: async () => json });

// ---- identity binding ---------------------------------------------------------------------------------------------

test('ZERO_REQUESTER -> ZERO_CONFIG, ELLIOT_REQUESTER -> ELLIOT_CONFIG: each request carries ONLY its own token', async () => {
  for (const who of ['zero', 'maintainer']) {
    const b = fakeBroker();
    const r = await source(who, b).requestEditorConfig(FILE, 'READ');
    assert.equal(r.ok, true, JSON.stringify(r));
    assert.equal(r.config.editorConfig.user.id, `instfixture_${who}`);
    assert.equal(b.calls.length, 1);
    assert.equal(b.calls[0].bearer, TOKENS[who]);
    assert.equal(b.calls[0].url, BROKER);
  }
});

test('the broker request body is exactly {op, instance, fileId, requestedMode}: no user, uid, requester or credential', async () => {
  const b = fakeBroker();
  await source('zero', b).requestEditorConfig(FILE, 'WRITE');
  assert.deepEqual(Object.keys(b.calls[0].body).sort(), ['fileId', 'instance', 'op', 'requestedMode']);
  assert.deepEqual(b.calls[0].body, { op: 'product_ocs_config', instance: 'production', fileId: FILE, requestedMode: 'WRITE' });
});

test('CROSS_CREDENTIAL=IMPOSSIBLE from the product side: identity-bearing tool arguments are refused before any broker call', () => {
  for (const k of ['user', 'uid', 'requester', 'principal', 'nextcloudUser', 'as_user', 'password', 'app_password', 'credential', 'agent_id']) {
    const r = validateTarget({ fileId: FILE, [k]: 'maintainer' }, 'READ');
    assert.deepEqual([r.ok, r.code], [false, ERRORS.BROKER_REQUEST_INVALID], k);
    const nested = validateTarget({ fileId: FILE, options: { [k]: 'maintainer' } }, 'READ');
    assert.equal(nested.ok, false, `nested ${k}`);
  }
  assert.deepEqual(validateTarget({ fileId: FILE }, 'READ'), { ok: true, fileId: FILE, operation: 'READ' });
  assert.equal(validateTarget({ fileId: FILE }, 'ADMIN').ok, false);
});

test('CROSS_CREDENTIAL=IMPOSSIBLE: a caller-supplied user/credential option has no effect -- only the own token is sent', async () => {
  const b = fakeBroker();
  const s = createBrokerConfigSource({ brokerUrl: BROKER, readToken: () => TOKENS.zero, fetchImpl: b.fetchImpl, user: 'maintainer', nextcloudUser: 'maintainer', credential: 'x' });
  const r = await s.requestEditorConfig(FILE, 'READ');
  assert.equal(r.config.editorConfig.user.id, 'instfixture_zero');
  assert.equal(b.calls[0].bearer, TOKENS.zero);
  assert.equal(JSON.stringify(b.calls[0].body).includes('maintainer'), false);
});

test('cross-requester RESPONSE identity: not verifiable in the product (documented, no false proof)', async () => {
  // The product has no trustworthy expected Nextcloud uid (the uid mapping and the instance prefix are broker-side and
  // can differ from the agent name, e.g. principal alice -> NC uid alice). The binding token -> own credential is the
  // broker's (LIVE-proven, M11.1b). What the product CAN verify is the target binding: a config for another fileId fails.
  const b = fakeBroker(({ body }) => ({ status: 200, json: async () => ({ ok: true, status: 200, config: editConfig('zero', String(Number(body.fileId) + 1)) }) }));
  const r = await source('zero', b).requestEditorConfig(FILE, 'READ');
  assert.deepEqual([r.ok, r.code], [false, ERRORS.BROKER_RESPONSE_INVALID]);
  assert.match(r.detail, /not bound to the requested fileId/);
});

// ---- controlled errors, no fallback --------------------------------------------------------------------------------

const MAPPING = [
  ['broker 400 BAD_BODY', reply(400, { code: 'PRODUCT_OCS_CONFIG_BAD_BODY' }), ERRORS.BROKER_REQUEST_INVALID],
  ['broker 401 (token not accepted)', reply(401, { code: 'AUTHZ_DENY_NO_AUTH' }), ERRORS.CONTROLLED_ACCESS_ERROR],
  ['principal not allowed (euro gate)', reply(403, { code: 'EURO_BROKER_DENY_PRINCIPAL' }), ERRORS.CONTROLLED_ACCESS_ERROR],
  ['principal not allowed (product gate)', reply(403, { code: 'PRODUCT_OCS_CONFIG_DENY_PRINCIPAL' }), ERRORS.CONTROLLED_ACCESS_ERROR],
  ['MISSING_MAPPING', reply(403, { code: 'AUTHZ_DENY_NO_NC_IDENTITY' }), ERRORS.CONTROLLED_ACCESS_ERROR],
  ['MISSING_CREDENTIAL', reply(403, { code: 'PRODUCT_OCS_CONFIG_DENY_NO_REQUESTER_SECRET' }), ERRORS.CONTROLLED_ACCESS_ERROR],
  ['credential use refused', reply(403, { code: 'AUTHZ_DENY_USE' }), ERRORS.CONTROLLED_ACCESS_ERROR],
  ['credential unavailable', reply(500, { code: 'PRODUCT_OCS_CONFIG_REQUESTER_CREDENTIAL_UNAVAILABLE' }), ERRORS.CONTROLLED_ACCESS_ERROR],
  ['DENY_AUTHZ (G2: missing file == not allowed)', reply(403, { code: 'PRODUCT_OCS_CONFIG_DENY_AUTHZ' }), ERRORS.ACCESS_DENIED],
  ['DENY_LEVEL', reply(403, { code: 'PRODUCT_OCS_CONFIG_DENY_LEVEL' }), ERRORS.ACCESS_DENIED],
  ['WRITE answered with a view config (target read-only)', reply(403, { code: 'PRODUCT_OCS_CONFIG_DENY_CONFIG_BELOW_REQUEST' }), ERRORS.ACCESS_DENIED],
  ['G1 CONFIG_EXCEEDS_VERDICT', reply(403, { code: 'PRODUCT_OCS_CONFIG_DENY_CONFIG_EXCEEDS_VERDICT' }), ERRORS.EXTERNAL_POLICY_DENIED],
  ['zuul2 unavailable', reply(503, { code: 'PRODUCT_ZUUL2_UNAVAILABLE' }), ERRORS.BROKER_UNAVAILABLE],
  ['upstream failed', reply(502, { code: 'PRODUCT_OCS_CONFIG_UPSTREAM_FAILED' }), ERRORS.BROKER_UNAVAILABLE],
  ['broker 500 without code', reply(500, {}), ERRORS.BROKER_UNAVAILABLE],
  ['broker unreachable', () => { throw new TypeError('fetch failed'); }, ERRORS.BROKER_UNAVAILABLE],
  ['broker timeout', () => { const e = new Error('timed out'); e.name = 'TimeoutError'; throw e; }, ERRORS.BROKER_UNAVAILABLE],
  ['200 without ok', reply(200, { config: editConfig('zero') }), ERRORS.BROKER_RESPONSE_INVALID],
  ['200 ok without config', reply(200, { ok: true, status: 200 }), ERRORS.BROKER_RESPONSE_INVALID],
  ['200 non-JSON body', () => ({ status: 200, json: async () => { throw new SyntaxError('x'); } }), ERRORS.BROKER_RESPONSE_INVALID],
  ['200 config without document.key', reply(200, { ok: true, config: { ...editConfig('zero'), document: { referenceData: { fileKey: FILE } } } }), ERRORS.BROKER_RESPONSE_INVALID],
  ['unknown 4xx code', reply(418, { code: 'SOMETHING_NEW' }), ERRORS.BROKER_RESPONSE_INVALID],
];

test('every broker answer maps onto exactly one controlled class; one request, no retry, no other mode or credential', async () => {
  for (const [name, answer, expected] of MAPPING) {
    for (const operation of ['READ', 'WRITE']) {
      const b = fakeBroker(answer);
      const r = await source('zero', b).requestEditorConfig(FILE, operation);
      assert.deepEqual([r.ok, r.code], [false, expected], `${name} / ${operation}`);
      assert.equal(b.calls.length, 1, `${name}: exactly one broker request (no fallback)`);
      assert.equal(b.calls[0].body.requestedMode, operation, `${name}: the mode is never escalated`);
      assert.equal(b.calls[0].bearer, TOKENS.zero, `${name}: never another credential`);
    }
  }
});

test('G1_FALLBACK=NONE: READ refused with CONFIG_EXCEEDS_VERDICT -> EXTERNAL_POLICY_DENIED, no WRITE retry, no Basic auth', async () => {
  const realFetch = globalThis.fetch;
  let direct = 0;
  globalThis.fetch = async () => { direct += 1; throw new Error('a direct fetch must not happen'); };
  try {
    const b = fakeBroker(reply(403, { code: 'PRODUCT_OCS_CONFIG_DENY_CONFIG_EXCEEDS_VERDICT' }));
    const s = source('zero', b);
    await assert.rejects(s.getEditorConfig({ fileId: FILE, operation: 'READ' }), e => e instanceof ConfigSourceError && e.code === ERRORS.EXTERNAL_POLICY_DENIED);
    assert.deepEqual(b.calls.map(c => c.body.requestedMode), ['READ']);
    assert.equal(direct, 0);
  } finally { globalThis.fetch = realFetch; }
});

test('missing agent token -> CONTROLLED_ACCESS_ERROR before any broker call; bad local input -> BROKER_REQUEST_INVALID', async () => {
  const b = fakeBroker();
  const noToken = createBrokerConfigSource({ brokerUrl: BROKER, readToken: () => { throw new Error('ENOENT /home/x/.agent-token'); }, fetchImpl: b.fetchImpl });
  const r = await noToken.requestEditorConfig(FILE, 'READ');
  assert.deepEqual([r.ok, r.code], [false, ERRORS.CONTROLLED_ACCESS_ERROR]);
  assert.equal(r.detail.includes('/home/x'), false, 'no path in the outcome');
  for (const [s, f, op] of [[source('zero', b), 'abc', 'READ'], [source('zero', b), FILE, 'DELETE'],
    [createBrokerConfigSource({ brokerUrl: 'https://broker.example.invalid/api/euro-broker', readToken: () => TOKENS.zero, fetchImpl: b.fetchImpl }), FILE, 'READ']]) {
    const x = await s.requestEditorConfig(f, op);
    assert.deepEqual([x.ok, x.code], [false, ERRORS.BROKER_REQUEST_INVALID]);
  }
  assert.equal(b.calls.length, 0);
});

test('SECRET_LEAK=NO: the agent token appears in no outcome, refusal detail or thrown error', async () => {
  const outcomes = [];
  for (const [, answer] of MAPPING) {
    const s = source('zero', fakeBroker(answer));
    outcomes.push(await s.requestEditorConfig(FILE, 'READ'));
    outcomes.push(await s.getEditorConfig({ fileId: FILE, operation: 'READ' }).catch(e => ({ message: e.message, code: e.code, detail: e.detail })));
  }
  outcomes.push(await source('zero', fakeBroker()).requestEditorConfig(FILE, 'READ'));
  const text = JSON.stringify(outcomes);
  for (const t of Object.values(TOKENS)) assert.equal(text.includes(t), false);
  assert.equal(/password|app_?password/i.test(text), false);
});

// ---- READ with an edit-capable config: accepted, and non-mutating ---------------------------------------------------

const frame = msg => '42' + JSON.stringify(['message', msg]);
// the three generated property-query bodies (B21: default | not requested | runtime not verified)
const PROPERTY_QUERIES = Object.freeze({ [docx.propertyQuery()]: 'default',
  [docx.propertyQuery({ sectionVariantState: false, sectionVariantUnavailable: 'not-requested' })]: 'opt-out',
  [docx.propertyQuery({ sectionVariantState: false, sectionVariantUnavailable: 'runtime-not-verified' })]: 'gated' });

function fakeWs(url) {
  return { url: () => url, handlers: {}, on(ev, cb) { this.handlers[ev] = cb; }, emit(ev, payload) { this.handlers[ev] && this.handlers[ev]({ payload }); } };
}

function fakeEditor() {
  const log = [];
  const ctl = { wsHandler: null, ws: null };
  const frameObj = {
    url: () => 'https://ds.example/9.3.4-x/web-apps/apps/documenteditor/main/index.html',
    evaluate: async (fn, arg) => {
      if (arg && arg.body === HANDSHAKE_BODY) { log.push('handshake'); return { ok: true, outcome: 'command-completed', result: null }; }
      // the READ path's own property query is a read (callCommand without a write); anything else would be a write command
      if (arg && PROPERTY_QUERIES[arg.body]) { log.push('read-query'); log.push(`read-query:${PROPERTY_QUERIES[arg.body]}`); return { ok: true, outcome: 'command-completed', result: { paragraphs: [], document: {} } }; }
      if (arg && arg.body) { log.push('command'); return { ok: true, outcome: 'command-completed', result: null }; }
      const src = String(fn);
      if (src.includes('asc_isDocumentCanSave')) { log.push('canSave'); return { ok: true, canSave: false }; }
      if (src.includes('e.forceSave()')) { log.push('forceSave'); return { ok: true, sent: true }; }
      if (src.includes('pluginMethod_GetFileHTML() ||')) { log.push('readLive'); return '<p>fixture</p>'; }
      return true;
    },
  };
  const page = {
    on(ev, cb) { if (ev === 'websocket') ctl.wsHandler = cb; }, goto: async () => {}, waitForFunction: async () => {}, waitForTimeout: async () => {},
    frames: () => [frameObj],
    evaluate: async fn => { if (String(fn).includes('requestClose')) log.push('requestClose'); return { ok: true, detail: 'requestClose-called' }; },
  };
  let launches = 0;
  const browser = { newPage: async () => page, close: async () => { log.push('browser-close'); if (ctl.ws && ctl.ws.handlers.close) ctl.ws.handlers.close(); } };
  // the editor <-> DS socket of the session (the config's DS host + document key), authenticated, no save frames
  const connect = fileId => {
    ctl.ws = fakeWs(`wss://ds.example/9.3.4-x/doc/42/c/?shardkey=KEY-${fileId}&EIO=4&transport=websocket`);
    ctl.wsHandler(ctl.ws);
    ctl.ws.emit('framereceived', frame({ type: 'auth', result: 1 }));
  };
  return { log, connect, launches: () => launches, loadPlaywright: () => ({ ok: true, pw: { chromium: { launch: async () => { launches += 1; return browser; } } } }) };
}

const opened = [];
after(async () => { for (const d of opened) await d.close().catch(() => {}); });

test('READ_WITH_EDIT_CONFIG=SAFE: an edit-capable config is accepted; open + readLive + close dispatch nothing and save nothing', async () => {
  const realFetch = globalThis.fetch;
  let direct = 0;
  globalThis.fetch = async () => { direct += 1; throw new Error('the product path must not fetch Nextcloud directly'); };
  try {
    const b = fakeBroker();
    const ed = fakeEditor();
    const d = docx.createEuroOfficeDocxDriver({ configSource: source('zero', b), loadPlaywright: ed.loadPlaywright, pollMs: 5, closeAckTimeoutMs: 200 });
    opened.push(d);
    const o = await d.open({ fileId: FILE, operation: 'READ' });
    assert.equal(o.ok, true, o.detail);
    assert.equal(b.calls.length, 1);
    assert.equal(b.calls[0].body.requestedMode, 'READ');
    ed.connect(FILE);
    const r = await d.readLive();
    assert.equal(r.ok, true, r.detail);
    const c = await d.close();
    opened.pop();
    assert.equal(c.ok, true, c.detail);
    // DOCUMENT_MUTATION=0 / WRITE=0: no write command was dispatched; HISTORY_MUTATION=0 / SAVE_REQUIRED=NO: no forceSave,
    // and the close confirmed that no save traffic happened
    assert.equal(ed.log.includes('read-query'), true, 'the READ path read the live document');
    assert.equal(ed.log.includes('command'), false);
    assert.equal(ed.log.includes('forceSave'), false);
    assert.match(String(c.detail), /no-save-traffic/);
    assert.equal(direct, 0, 'no direct Nextcloud/Basic-auth fetch');
  } finally { globalThis.fetch = realFetch; }
});

test('a refused config never starts the editor: the driver open fails with the controlled class, no browser launch, no Basic-auth fallback', async () => {
  const realFetch = globalThis.fetch;
  let direct = 0;
  globalThis.fetch = async () => { direct += 1; throw new Error('BROKER -> DIRECT_BASIC_AUTH fallback must not happen'); };
  try {
  for (const [answer, expected] of [[reply(403, { code: 'PRODUCT_OCS_CONFIG_DENY_AUTHZ' }), ERRORS.ACCESS_DENIED],
    [reply(403, { code: 'PRODUCT_OCS_CONFIG_DENY_CONFIG_EXCEEDS_VERDICT' }), ERRORS.EXTERNAL_POLICY_DENIED],
    [() => { throw new TypeError('fetch failed'); }, ERRORS.BROKER_UNAVAILABLE]]) {
    const ed = fakeEditor();
    const d = docx.createEuroOfficeDocxDriver({ configSource: source('zero', fakeBroker(answer)), loadPlaywright: ed.loadPlaywright, pollMs: 5 });
    const o = await d.open({ fileId: FILE, operation: 'READ' });
    assert.deepEqual([o.ok, o.code], [false, expected]);
    assert.equal(ed.launches(), 0);
  }
  assert.equal(direct, 0, 'BROKER_TO_BASIC_FALLBACK=NONE');
  } finally { globalThis.fetch = realFetch; }
});

test('BROKER_TO_BASIC_FALLBACK=NONE: an explicit configSource is never replaced by the Basic-auth source, even with url/user/pass present', () => {
  const { resolveConfigSource } = require('../../src/drivers/eurooffice/config-source.cjs');
  const s = source('zero', fakeBroker());
  assert.equal(resolveConfigSource({ configSource: s, url: 'https://nc.example', user: 'u', pass: 'p' }), s);
  assert.throws(() => resolveConfigSource({ configSource: { notASource: true }, url: 'https://nc.example', user: 'u', pass: 'p' }), TypeError);
  // the broker source module has no route to Nextcloud or to the legacy source at all
  // code only: the module's own comments name these paths to say they are NOT taken
  const text = fs.readFileSync(path.join(ROOT, 'src', 'server', 'broker-config-source.cjs'), 'utf8')
    .replace(/\/\*[\s\S]*?\*\//g, '').split('\n').map(l => l.replace(/\/\/.*$/, '')).join('\n');
  assert.equal(/createBasicAuthConfigSource|resolveConfigSource|\/ocs\/|Basic /.test(text), false);
});

test('no production path constructs a driver or the Basic-auth source: the 79 user/pass constructions are test/harness only', () => {
  const walk = d => fs.readdirSync(d, { withFileTypes: true }).flatMap(e => (e.isDirectory() ? walk(path.join(d, e.name)) : [path.join(d, e.name)]));
  const prod = [...walk(path.join(ROOT, 'src')), ...walk(path.join(ROOT, 'bin'))].filter(f => f.endsWith('.cjs') || f.endsWith('.js'));
  const owners = new Set(['src/drivers/eurooffice/docx-driver.cjs', 'src/drivers/eurooffice/xlsx-driver.cjs', 'src/drivers/eurooffice/config-source.cjs']);
  // M11.3a: the ONE product-side driver constructor (the READ tools) -- checked by its own, narrower rule below
  const READ_TOOLS = 'src/server/read-tools.cjs';
  for (const f of prod) {
    const rel = path.relative(ROOT, f);
    if (owners.has(rel)) continue;
    const src = fs.readFileSync(f, 'utf8');
    assert.equal(/createBasicAuthConfigSource\(|resolveConfigSource\(/.test(src), false, rel);
    if (rel !== READ_TOOLS) assert.equal(/createEuroOffice(Docx|Xlsx)Driver\(/.test(src), false, rel);
  }
});

test('M11.3a: ALLOWED_PRODUCT_DRIVER_CONSTRUCTION_SITES=1, exactly src/server/read-tools.cjs', () => {
  const walk = d => fs.readdirSync(d, { withFileTypes: true }).flatMap(e => (e.isDirectory() ? walk(path.join(d, e.name)) : [path.join(d, e.name)]));
  const owners = new Set(['src/drivers/eurooffice/docx-driver.cjs', 'src/drivers/eurooffice/xlsx-driver.cjs', 'src/drivers/eurooffice/config-source.cjs']);
  const sites = [...walk(path.join(ROOT, 'src')), ...walk(path.join(ROOT, 'bin'))].filter(f => /\.c?js$/.test(f))
    .map(f => path.relative(ROOT, f)).filter(rel => !owners.has(rel))
    .filter(rel => /createEuroOffice(Docx|Xlsx)Driver\(/.test(fs.readFileSync(path.join(ROOT, rel), 'utf8')));
  assert.deepEqual(sites, ['src/server/read-tools.cjs']);
});

test('M11.3a: the READ tools construct a driver ONLY with the broker configSource -- never url/user/pass', () => {
  const code = fs.readFileSync(path.join(ROOT, 'src', 'server', 'read-tools.cjs'), 'utf8')
    .replace(/\/\*[\s\S]*?\*\//g, '').split('\n').map(l => l.replace(/\/\/.*$/, '')).join('\n');
  const calls = code.match(/createEuroOffice(Docx|Xlsx)Driver\([^)]*\)/g) || [];
  assert.equal(calls.length, 2, 'one constructor per format');
  for (const c of calls) assert.match(c, /^createEuroOffice(Docx|Xlsx)Driver\(\{ configSource, loadPlaywright \}\)$/, c);
  assert.match(code, /configSource\.kind !== 'BROKER'/, 'a non-broker source is refused before the constructor');
  assert.equal(/\b(url|user|pass)\s*:/.test(code), false, 'no url/user/pass option anywhere in the READ tools');
  const { productDriver } = require('../../src/server/read-tools.cjs');
  assert.throws(() => productDriver('docx', { configSource: { kind: 'BASIC_AUTH', getEditorConfig: async () => ({}) } }), TypeError);
  assert.throws(() => productDriver('xlsx', { configSource: null }), TypeError);
});

// ---- product runtime: broker source only; tools unchanged ----------------------------------------------------------

const PRODUCT_RUNTIME_FILES = () => [path.join(ROOT, 'bin', 'euro-office-mcp.cjs'),
  ...fs.readdirSync(path.join(ROOT, 'src', 'server')).map(f => path.join(ROOT, 'src', 'server', f))];

test('DIRECT_BASIC_AUTH_PRODUCT_CALLS=0: the product runtime never builds the legacy Basic-auth source or passes user/pass', () => {
  for (const f of PRODUCT_RUNTIME_FILES()) {
    const src = fs.readFileSync(f, 'utf8');
    assert.equal(/createBasicAuthConfigSource|resolveConfigSource\(/.test(src), false, `${path.relative(ROOT, f)}: legacy source`);
    assert.equal(/'Basic '|"Basic "|`Basic /.test(src), false, `${path.relative(ROOT, f)}: Basic auth`);
    assert.equal(/\bpass\s*:/.test(src), false, `${path.relative(ROOT, f)}: a pass option`);
  }
});

// M11.2 left the exposure unchanged; M11.3b.2 released the two READ tools; M11.4b.1 (owner, owner channel) adds
// xlsx_execute_task -- docx_execute_task stays withheld (exact-set guard: test/static/tool-exposure-exact-set-guard.test.cjs)
test('tool exposure: product_status + the two READ tools + xlsx_execute_task exposed, docx_execute_task withheld', () => {
  assert.deepEqual(exposedTools(), ['product_status', 'xlsx_read', 'docx_read', 'xlsx_execute_task']);
  assert.deepEqual(TOOL_ROWS.filter(r => r.EXPOSE_IN_PRODUCT_MCP !== 'YES').map(r => r.TOOL).sort(), ['docx_execute_task']);
});

test('M12: the own token is read from EURO_OFFICE_MCP_TOKEN_PATH at call time and only travels as the Bearer header', async () => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'eo-token-path-'));
  try {
    const file = path.join(dir, 'token');
    fs.writeFileSync(file, 'own-token-123\n');
    let seen = null;
    const fetchImpl = async (url, init) => { seen = init; return { status: 200, json: async () => ({ ok: true, status: 200, config: { document: { key: 'k', referenceData: { fileKey: '42' } }, documentServerUrl: 'https://ds.example.org/' } }) }; };
    const src = createBrokerConfigSource({ brokerUrl: BROKER, env: { EURO_OFFICE_MCP_TOKEN_PATH: file }, fetchImpl });
    const r = await src.requestEditorConfig('42', 'READ');
    assert.equal(r.ok, true);
    assert.equal(seen.headers.Authorization, 'Bearer own-token-123');
    assert.doesNotMatch(seen.body, /own-token-123/);
  } finally { fs.rmSync(dir, { recursive: true, force: true }); }
});
