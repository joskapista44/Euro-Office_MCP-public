#!/usr/bin/env node
'use strict';

/**
 * Euro-Office_MCP product entrypoint (owner decision, owner channel 2026-10-08) -- stdio MCP, no daemon, no port.
 *
 *   euro-office-mcp [--config <path>]              serve MCP on stdin/stdout
 *   euro-office-mcp --identity [--config <path>]   print the product/runtime identity (KEY=VALUE) and exit
 *   euro-office-mcp --check-config [--config <p>]  validate the config and exit
 *
 * Config: --config, else $EURO_OFFICE_MCP_CONFIG, else /etc/euro-office-mcp/config.json.
 * Exit codes: 0 ok | 64 usage | 78 configuration invalid (fail-closed startup: nothing is served).
 */

const { DEFAULT_CONFIG_PATH, loadConfig } = require('../src/server/config.cjs');
const { createLogger } = require('../src/server/log.cjs');
const { createMcpServer } = require('../src/server/mcp.cjs');
const { productStatus, identityLines } = require('../src/server/product-status.cjs');
const { buildServerTools } = require('../src/server/server-tools.cjs');
const { releaseIdentity } = require('../src/server/identity.cjs');

const EX_USAGE = 64;
const EX_CONFIG = 78;
// M11.3b.2 bounded shutdown (owner decision, owner channel 2026-10-10): never wait forever for an active session
const SHUTDOWN_TIMEOUT_MS = 10000;
const SHUTDOWN_GRACE_MS = 2000;

function parseArgs(argv) {
  const out = { mode: 'serve', config: null };
  for (let i = 0; i < argv.length; i += 1) {
    const a = argv[i];
    if (a === '--identity') out.mode = 'identity';
    else if (a === '--check-config') out.mode = 'check-config';
    else if (a === '--config') { out.config = argv[i + 1]; i += 1; if (!out.config) return { error: '--config needs a path' }; }
    else return { error: `unknown argument: ${a}` };
  }
  return out;
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  if (args.error) { process.stderr.write(`euro-office-mcp: ${args.error}\n`); return EX_USAGE; }
  const file = args.config || process.env.EURO_OFFICE_MCP_CONFIG || DEFAULT_CONFIG_PATH;
  const cfg = loadConfig(file);

  if (args.mode === 'check-config') {
    if (cfg.ok) { process.stdout.write(`CONFIG_OK ${cfg.identity}\n`); return 0; }
    process.stderr.write(`CONFIG_INVALID ${JSON.stringify(cfg.errors)}\n`);
    return EX_CONFIG;
  }

  if (args.mode === 'identity') {
    const status = await productStatus({ config: cfg.ok ? cfg.config : null, configIdentity: cfg.ok ? cfg.identity : null });
    const lines = identityLines(status);
    if (!cfg.ok) lines[lines.findIndex(l => l.startsWith('CONFIG_IDENTITY='))] = 'CONFIG_IDENTITY=INVALID';
    process.stdout.write(`${lines.join('\n')}\n`);
    if (!cfg.ok) { process.stderr.write(`CONFIG_INVALID ${JSON.stringify(cfg.errors)}\n`); return EX_CONFIG; }
    return 0;
  }

  // serve: an invalid config is a fail-closed startup -- nothing reaches stdout, the reason goes to stderr
  if (!cfg.ok) { process.stderr.write(`euro-office-mcp: CONFIG_INVALID ${JSON.stringify(cfg.errors)} (${file})\n`); return EX_CONFIG; }
  const log = createLogger({ level: (cfg.config.log && cfg.config.log.level) || 'info' });
  const rel = releaseIdentity();
  log.info('starting', { product: rel.PRODUCT, sourceCommit: rel.SOURCE_COMMIT, artifactVersion: rel.ARTIFACT_VERSION, configIdentity: cfg.identity, pid: process.pid });

  // every tool is built; only the exposure gate's subset is served (M11.3b.2: product_status + xlsx_read + docx_read)
  const { tools, gate } = buildServerTools({ config: cfg.config, configIdentity: cfg.identity });
  const server = createMcpServer({ tools, info: { version: rel.ARTIFACT_VERSION }, log });
  let shuttingDown = false;
  for (const sig of ['SIGTERM', 'SIGINT']) {
    process.on(sig, () => {
      if (shuttingDown) return;
      shuttingDown = true;
      server.stop(sig);
      // M11.3b.2 bounded shutdown: the active editor session is closed through its driver's own close(), bounded by
      // SHUTDOWN_TIMEOUT_MS; the hard deadline exits even if a call never settles (the browser child dies with its pipe)
      setTimeout(() => { log.error('shutdown deadline reached, exiting', { timeoutMs: SHUTDOWN_TIMEOUT_MS }); process.exit(0); }, SHUTDOWN_TIMEOUT_MS + SHUTDOWN_GRACE_MS).unref();
      gate.shutdown({ timeoutMs: SHUTDOWN_TIMEOUT_MS }).then(r => log.info('session cleanup', r));
    });
  }
  const reason = await server.start();
  log.info('stopped', { reason });
  return 0;
}

main().then(code => { process.exitCode = code; }, err => {
  process.stderr.write(`euro-office-mcp: fatal ${String(err && err.message || err)}\n`);
  process.exitCode = 1;
});
