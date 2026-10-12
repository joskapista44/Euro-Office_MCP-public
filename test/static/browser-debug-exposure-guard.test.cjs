'use strict';

/**
 * Static gate: the browser runtime keeps NO debug exposure (owner decision M11.3b.1 closure, owner channel 2026-10-10).
 * Canonical runtime: Playwright `chromium.launch()` with no options -> Chromium runs with --remote-debugging-pipe (CDP over
 * inherited fds, unreachable for any other process/UID). Debug TCP port: NONE. A launch option, a debugging flag, a
 * CDP connect/launchServer path or a persistent profile reintroduced into src/ turns this gate red.
 * Evidence of the live runtime (debug_flags=[--remote-debugging-pipe], no port, cross-UID): docs/M11-3B1-LOOPBACK-HARDENING.md.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const SRC = path.join(__dirname, '..', '..', 'src');
// the exact register of browser launch sites: a new one must be added here on purpose (and stay option-less)
const LAUNCH_SITES = Object.freeze(['drivers/eurooffice/docx-driver.cjs', 'drivers/eurooffice/xlsx-driver.cjs']);
const FORBIDDEN = Object.freeze([/remote-debugging-(port|address)/i, /remoteDebugging/i, /connectOverCDP/, /launchServer/,
  /launchPersistentContext/, /chromium\.connect\b/, /\bdevtools\s*:/, /\bwsEndpoint\b/, /--inspect\b/]);

const stripComments = src => src.replace(/\/\*[\s\S]*?\*\//g, '').replace(/(^|[^:'"`])\/\/.*$/gm, '$1');

function srcModules(dir = SRC, out = []) {
  for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
    const p = path.join(dir, e.name);
    if (e.isDirectory()) srcModules(p, out);
    else if (p.endsWith('.cjs')) out.push({ file: path.relative(SRC, p).split(path.sep).join('/'), source: fs.readFileSync(p, 'utf8') });
  }
  return out;
}

/** judge(mods) -> { forbidden: [{file, rule}], launches: [{file, args}], unregistered: [file] } */
function judge(mods) {
  const forbidden = [];
  const launches = [];
  for (const m of mods) {
    const code = stripComments(m.source);
    for (const r of FORBIDDEN) if (r.test(code)) forbidden.push({ file: m.file, rule: String(r) });
    for (const x of code.matchAll(/\bchromium\s*\.\s*launch\s*\(([^)]*)\)/g)) launches.push({ file: m.file, args: x[1].trim() });
  }
  const unregistered = [...new Set(launches.map(l => l.file))].filter(f => !LAUNCH_SITES.includes(f));
  return { forbidden, launches, unregistered };
}

test('DEBUG_PORT_GUARD: no debugging port/address, CDP connect, launchServer or persistent profile in src/', () => {
  const mods = srcModules();
  assert.ok(mods.length > 30, 'the walk sees the source tree');
  assert.deepEqual(judge(mods).forbidden, []);
});

test('every browser launch is option-less (Playwright default = --remote-debugging-pipe) and in the exact register', () => {
  const v = judge(srcModules());
  assert.deepEqual(v.launches, LAUNCH_SITES.map(file => ({ file, args: '' })));
  assert.deepEqual(v.unregistered, []);
});

test('NEG controls: the gate goes red on a debug port, a launch option, a CDP path and an unregistered launch site', () => {
  const base = srcModules();
  const put = (file, extra) => base.map(m => (m.file === file ? { ...m, source: `${m.source}\n${extra}` } : m));
  const D = LAUNCH_SITES[0];
  assert.equal(judge(put(D, "const a = ['--remote-debugging-port=9222'];")).forbidden.length, 1);
  assert.equal(judge(put(D, "const a = ['--remote-debugging-address=0.0.0.0'];")).forbidden.length, 1);
  assert.equal(judge(put(D, 'const b = await pw.chromium.connectOverCDP(u);')).forbidden.length, 1);
  assert.equal(judge(put(D, 'const s = await pw.chromium.launchServer();')).forbidden.length, 1);
  assert.equal(judge(put(D, 'const c = await pw.chromium.launchPersistentContext(dir);')).forbidden.length, 1);
  const withArgs = judge(put(D, "const b2 = await pw.chromium.launch({ args: ['--no-sandbox'] });"));
  assert.notDeepEqual(withArgs.launches, LAUNCH_SITES.map(file => ({ file, args: '' })));
  const extraSite = judge([...base, { file: 'server/new-runtime.cjs', source: 'const b = await pw.chromium.launch();' }]);
  assert.deepEqual(extraSite.unregistered, ['server/new-runtime.cjs']);
  // a comment naming the flag is not a use
  assert.deepEqual(judge([{ file: 'x.cjs', source: '// never pass --remote-debugging-port\n/* connectOverCDP is forbidden */' }]).forbidden, []);
});

module.exports = { judge, LAUNCH_SITES };
