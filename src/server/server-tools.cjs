'use strict';

/**
 * The product MCP tool registry (M11.3a). Every product tool is BUILT here -- `product_status` and the READ tools
 * `xlsx_read` / `docx_read` (released in M11.3b.2) and the WITHHELD WRITE tool `xlsx_execute_task` (M11.4a.1) -- but the server is handed ONLY the exposed subset (tools.cjs `exposedTools()`,
 * the exposure gate). A withheld tool is therefore neither listed by `tools/list` nor callable by `tools/call`
 * ("Unknown tool"); releasing it is the gate's decision (EXPOSED_CAPABILITY_TOOLS + preconditions), not this file's.
 */

const { productStatus } = require('./product-status.cjs');
const { exposedTools } = require('./tools.cjs');
const { createBrokerConfigSource } = require('./broker-config-source.cjs');
const { createReadTool } = require('./read-tools.cjs');
const { createSessionGate } = require('./session-gate.cjs');
const { createWriteTool } = require('./write-tools.cjs');

/**
 * buildServerTools({ config, configIdentity, readTool, writeTool }) -> { registry, tools, gate }. `readTool` / `writeTool`: test seam options.
 * M11.3b.2: ONE session gate per process, shared by xlsx_read and docx_read (MAX_ACTIVE_SESSIONS_PER_PROCESS=1); the
 * browser runtime comes from config `playwright.path`.
 */
function buildServerTools({ config, configIdentity, readTool = {}, writeTool = {} } = {}) {
  const b = (config && config.broker) || {};
  const configSource = createBrokerConfigSource({ brokerUrl: b.url, instance: b.instance, timeoutMs: b.timeoutMs });
  const gate = readTool.gate || createSessionGate();
  const browsersPath = config && config.playwright ? config.playwright.path : undefined;
  const registry = {
    product_status: {
      description: 'Read-only Euro-Office_MCP product status: identity, artifact, config schema, DS Bridge handshake, lifecycle gate, dependencies. Returns no secrets.',
      inputSchema: { type: 'object', properties: {}, additionalProperties: false },
      handler: () => productStatus({ config, configIdentity }),
    },
    xlsx_read: createReadTool('xlsx', { configSource, browsersPath, ...readTool, gate }),
    docx_read: createReadTool('docx', { configSource, browsersPath, ...readTool, gate }),
    // M11.4a.1 handler + schema, M11.4a.2 canonical executor wiring (the SAME session gate as the READ tools) -- built but
    // WITHHELD: not in exposedTools(), so neither listed nor callable. There is no DOCX WRITE handler (DOCX_WRITE_ACCEPTED is empty).
    xlsx_execute_task: createWriteTool('xlsx', { configSource, browsersPath, ...writeTool, gate }),
  };
  const tools = Object.fromEntries(exposedTools().filter(n => registry[n]).map(n => [n, registry[n]]));
  return { registry, tools, gate };
}

module.exports = { buildServerTools };
