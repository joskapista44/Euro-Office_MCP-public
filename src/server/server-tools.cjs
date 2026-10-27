'use strict';

/**
 * The product MCP tool registry (M11.3a). Every product tool is BUILT here -- `product_status`, the READ tools
 * `xlsx_read` / `docx_read` (released in M11.3b.2), the RELEASED WRITE tool `xlsx_execute_task` (M11.4b.1) and
 * the WITHHELD WRITE tool `docx_execute_task` (built in GH #219 3, same shape `xlsx_execute_task` had before
 * M11.4b.1) -- but the server is handed ONLY the exposed subset (tools.cjs `exposedTools()`, the exposure gate).
 * A withheld tool is therefore neither listed by `tools/list` nor callable by `tools/call` ("Unknown tool");
 * releasing it is the gate's decision (EXPOSED_CAPABILITY_TOOLS + preconditions), not this file's.
 */

const { productStatus } = require('./product-status.cjs');
const { exposedTools } = require('./tools.cjs');
const { createBrokerConfigSource } = require('./broker-config-source.cjs');
const { createReadTool } = require('./read-tools.cjs');
const { createSessionGate } = require('./session-gate.cjs');
const { createSessionRegistry, sessionReuseConfig } = require('./session-registry.cjs');
const { createWriteTool } = require('./write-tools.cjs');
const { createDocxWriteTool } = require('./docx-write-tools.cjs');
const os = require('node:os');

/**
 * GH #219 3: DOCX_SESSION_REUSE (default OFF) -- a SEPARATE switch from EURO_MCP_SESSION_REUSE.
 * docx-driver.cjs gained the GH #100 U1/U3 surface in GH #220 1b (release/resetForNextCall/health/
 * documentKey, mirroring xlsx-driver.cjs), so wiring docx onto the SAME `sessions` registry
 * xlsx already uses is now structurally safe -- but it is NOT the same decision as "the registry
 * is enabled": `EURO_MCP_SESSION_REUSE=1` is already a LIVE, owner-GO'd production switch for
 * xlsx (docs/GH100-SESSION-REUSE-DESIGN.md section 6); turning docx reuse on as an unannounced
 * side effect of that flag would change `docx_read`'s (already EXPOSED, already LIVE) behaviour
 * without its own decision. `EURO_MCP_DOCX_SESSION_REUSE` must ALSO be '1' -- either switch OFF
 * (the default) leaves docx byte-identical to before this unit: no take()/hold(), no extra
 * config fetch, every call opens and closes its own session.
 */
const DOCX_SESSION_REUSE_ENV = 'EURO_MCP_DOCX_SESSION_REUSE';

/**
 * buildServerTools({ config, configIdentity, readTool, writeTool }) -> { registry, tools, gate, sessions }. `readTool` / `writeTool`: test seam options.
 * M11.3b.2: ONE session gate per process, shared by xlsx_read and docx_read (MAX_ACTIVE_SESSIONS_PER_PROCESS=1); the
 * browser runtime comes from config `playwright.path`.
 */
function buildServerTools({ config, configIdentity, readTool = {}, writeTool = {}, log = null, env = process.env } = {}) {
  const b = (config && config.broker) || {};
  const configSource = createBrokerConfigSource({ brokerUrl: b.url, instance: b.instance, timeoutMs: b.timeoutMs });
  const gate = readTool.gate || createSessionGate();
  // GH #100 U2: the warm-session registry. OFF unless the switch AND its three limits are set; nothing holds or takes a session before U4.
  const sessions = createSessionRegistry(sessionReuseConfig(env), { log });
  const browsersPath = config && config.playwright ? config.playwright.path : undefined;
  // GH #100 U4b: what a warm WRITE session is bound to. One process = one requester (the broker derives it from the process's own token), so the
  // principal is the process uid; the signature pins the broker route, the instance and the open mode.
  const principal = typeof process.getuid === 'function' ? String(process.getuid()) : os.userInfo().username;
  const reuse = {
    registry: sessions,
    principal,
    configSignature: `${b.instance || 'production'}|${b.url || ''}|WRITE`,
  };
  // GH #219 3: docx's OWN reuse object, on the SAME shared registry (`sessions`) -- but gated by its
  // OWN switch (see DOCX_SESSION_REUSE_ENV above), and with a configSignature that carries `|docx|`,
  // a format DISCRIMINATOR the xlsx signature above does not need and does not get (docs/
  // GH220-DOCX-WRITE-LIFECYCLE-DESIGN.md section 4: `sessionKey` itself has no format component --
  // a Nextcloud fileId is unique per file regardless of format, so the key is DERIVED-safe without
  // one; this discriminator is a zero-cost defensive margin, not a correctness requirement, added
  // here because this is the first unit that actually shares `sessions` across both formats).
  const docxReuseOn = sessions.enabled === true && env[DOCX_SESSION_REUSE_ENV] === '1';
  const docxReuse = docxReuseOn ? { registry: sessions, principal, configSignature: `${b.instance || 'production'}|${b.url || ''}|docx|WRITE` } : null;
  const registry = {
    product_status: {
      description: 'Read-only Euro-Office_MCP product status: identity, artifact, config schema, DS Bridge handshake, lifecycle gate, dependencies. Returns no secrets.',
      inputSchema: { type: 'object', properties: {}, additionalProperties: false },
      handler: () => productStatus({ config, configIdentity }),
    },
    // GH #100 U4c: the SAME reuse object as xlsx_execute_task below -- the same key (principal +
    // configSignature), so a read and a write on the same file contend for the ONE held slot
    // instead of evicting each other (docs/GH100-U4C-READ-REUSE-DESIGN.md sec. 3).
    xlsx_read: createReadTool('xlsx', { configSource, browsersPath, reuse, ...readTool, gate }),
    // GH #219 3: docx_read now ALSO carries a reuse object -- `docxReuse`, which is `null` unless
    // BOTH switches are on (see DOCX_SESSION_REUSE_ENV above). With either OFF (today's default),
    // `reuse: null` is byte-identical to this tool's shape before this unit: no take()/hold(), no
    // extra config fetch. docx_read and docx_execute_task share THIS SAME `docxReuse` object (the
    // SAME key), the GH #100 U4c read/write contention design, now extended to docx.
    docx_read: createReadTool('docx', { configSource, browsersPath, reuse: docxReuse, ...readTool, gate }),
    // M11.4a.1 handler + schema, M11.4a.2 canonical executor wiring (the SAME session gate as the READ tools) -- built but
    // WITHHELD: not in exposedTools(), so neither listed nor callable. There is no DOCX WRITE handler (DOCX_WRITE_ACCEPTED is empty).
    // GH #76: each write's phase timings go to the per-UID log (numbers only), so a slow call can be broken down afterwards
    xlsx_execute_task: createWriteTool('xlsx', { configSource, browsersPath, reuse, ...(log ? { onTiming: t => log.info('xlsx_execute_task timing', t) } : {}), ...writeTool, gate }),
    // GH #219 3: docx_execute_task BUILT here (the handler object exists, fully wired -- the SAME
    // session gate, broker configSource and now, if both switches are on, the SAME warm-session
    // registry as docx_read) but WITHHELD exactly the way xlsx_execute_task was before M11.4b.1:
    // `registry.docx_execute_task` exists, but `tools.cjs`'s EXPOSED_CAPABILITY_TOOLS does not name
    // it, so `tools` (the exposed subset below) never includes it -- tools/list omits it, tools/call
    // answers "Unknown tool". Releasing it is `tools.cjs`'s own decision (test/static/
    // tool-exposure-exact-set-guard.test.cjs), NOT this unit's -- unchanged by this diff.
    docx_execute_task: createDocxWriteTool({ configSource, browsersPath, reuse: docxReuse, ...writeTool, gate }),
  };
  const tools = Object.fromEntries(exposedTools().filter(n => registry[n]).map(n => [n, registry[n]]));
  return { registry, tools, gate, sessions };
}

module.exports = { buildServerTools };
