'use strict';

/**
 * M11.3b.2 -- the product browser runtime (owner decision, owner channel 2026-10-10).
 *
 * The product loads `playwright-core` (exact dependency, lockfile) and uses the browsers the host provides at the
 * configured `playwright.path` (production: /opt/ms-playwright). Pinned and checked before any broker call or launch:
 *   playwright-core 1.62.1 | chromium-headless-shell revision 1234 | the revision's executable present and executable.
 * BROWSER_DOWNLOAD=NO: nothing here installs or downloads a browser (playwright-core has no install script and its
 * launch never downloads); a missing or incompatible runtime is the controlled error BROWSER_RUNTIME_UNAVAILABLE.
 * The launch itself stays in the drivers and stays option-less (M11.3b.1 model: --remote-debugging-pipe, no debug port).
 */

const fs = require('node:fs');
const path = require('node:path');

const PINNED = Object.freeze({ playwrightCore: '1.62.1', browser: 'chromium-headless-shell', revision: '1234' });
const EXECUTABLE = Object.freeze(['chromium_headless_shell-1234', 'chrome-headless-shell-linux64', 'chrome-headless-shell']);
const UNAVAILABLE = 'BROWSER_RUNTIME_UNAVAILABLE';

const unavailable = detail => Object.freeze({ ok: false, code: UNAVAILABLE, detail });

/** The installed playwright-core package files (its `exports` do not expose them to require). */
function defaultReadPackageFile(name) {
  return JSON.parse(fs.readFileSync(path.join(path.dirname(require.resolve('playwright-core')), name), 'utf8'));
}

/** The installed playwright-core identity (version + pinned browser revision), read from its own package files. */
function playwrightCoreIdentity(readPackageFile) {
  try {
    const version = readPackageFile('package.json').version;
    const browsers = readPackageFile('browsers.json').browsers || [];
    const shell = browsers.find(b => b.name === PINNED.browser);
    return { version, revision: shell ? String(shell.revision) : null };
  } catch (e) {
    return { error: `playwright-core not loadable: ${String((e && e.code) || (e && e.message) || e)}` };
  }
}

function isExecutable(fsImpl, file) {
  try { fsImpl.accessSync(file, fs.constants.X_OK); return true; } catch { return false; }
}

function requirePlaywright(requireImpl) {
  try { return { ok: true, pw: requireImpl('playwright-core') }; } catch (e) { return unavailable(`playwright-core not loadable: ${String((e && e.code) || e)}`); }
}

/**
 * createBrowserRuntime({ browsersPath }) -> { check(), load() }
 *   check() -> { ok:true, executable } | { ok:false, code:'BROWSER_RUNTIME_UNAVAILABLE', detail }  (no side effect)
 *   load()  -> { ok:true, pw } | the same refusal  (the drivers' loadPlaywright seam)
 * `readPackageFile` / `requireImpl` / `fsImpl` are test seams.
 */
function createBrowserRuntime({ browsersPath, readPackageFile = defaultReadPackageFile, requireImpl = require, fsImpl = fs } = {}) {
  function check() {
    const executable = typeof browsersPath === 'string' && path.isAbsolute(browsersPath) ? path.join(browsersPath, ...EXECUTABLE) : null;
    const id = executable ? playwrightCoreIdentity(readPackageFile) : {};
    // the first failing condition, in order: configured dir -> loadable -> pinned version -> pinned revision -> installed
    const problem = [
      [!executable, 'playwright.path is not configured (absolute browsers directory)'],
      [id.error, id.error],
      [id.version !== PINNED.playwrightCore, `playwright-core ${id.version} is not the pinned ${PINNED.playwrightCore}`],
      [id.revision !== PINNED.revision, `playwright-core expects ${PINNED.browser} ${id.revision}, pinned ${PINNED.revision}`],
      [!isExecutable(fsImpl, executable), `${PINNED.browser} ${PINNED.revision} is not installed under playwright.path`],
    ].find(([bad]) => bad);
    return problem ? unavailable(problem[1]) : { ok: true, executable };
  }

  function load() {
    const c = check();
    if (!c.ok) return c;
    // the browsers directory must be set before playwright-core is first loaded (it reads the variable at load time)
    process.env.PLAYWRIGHT_BROWSERS_PATH = browsersPath;
    process.env.PLAYWRIGHT_SKIP_BROWSER_DOWNLOAD = '1';
    return requirePlaywright(requireImpl);
  }

  return { check, load };
}

module.exports = { createBrowserRuntime, PINNED, EXECUTABLE, UNAVAILABLE };
