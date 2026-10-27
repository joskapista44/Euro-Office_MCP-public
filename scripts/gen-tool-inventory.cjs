#!/usr/bin/env node
'use strict';

// Regenerates docs/PRODUCT-MCP-TOOL-INVENTORY-20261008.md from the code (src/server/tools.cjs):
//   node scripts/gen-tool-inventory.cjs > docs/PRODUCT-MCP-TOOL-INVENTORY-20261008.md
const { TOOL_ROWS, capabilityInventory, exposureGate, exposedTools } = require('../src/server/tools.cjs');

const caps = capabilityInventory();
const cols = ['TOOL', 'SOURCE_MODULE', 'FORMAT', 'READ_WRITE', 'AUTH_REQUIRED', 'CONFIG_REQUIRED', 'CURRENT_ACCEPTANCE', 'EXPOSE_IN_PRODUCT_MCP', 'REASON'];
const tbl = rows => [`| ${cols.join(' | ')} |`, `|${cols.map(() => '---').join('|')}|`, ...rows.map(r => `| ${cols.map(c => String(r[c]).replace(/\|/g, '/')).join(' | ')} |`)].join('\n');
const g = exposureGate();
const exposed = TOOL_ROWS.filter(r => r.EXPOSE_IN_PRODUCT_MCP === 'YES').length;
process.stdout.write(`# Product MCP tool inventory (2026-10-08)

Owner decision: the owner, owner channel 2026-10-08, point 4. Card: \`an internal branch\`.

**Generated from the code** (\`src/server/tools.cjs\`) by \`node scripts/gen-tool-inventory.cjs\`. The capability rows
are the intents the canonical adapters actually route (\`planOperations\`), not a copy of the Talk-bridge tool list. The
acceptance test \`test/acceptance/product-mcp-entrypoint.test.cjs\` pins the counts and the exposure, and fails if this
file drifts from the code.

\`\`\`text
TOOL_INVENTORY_TOTAL=${TOOL_ROWS.length} MCP tools + ${caps.length} capabilities (DOCX ${caps.filter(r => r.FORMAT === 'DOCX').length}, XLSX ${caps.filter(r => r.FORMAT === 'XLSX').length})
TOOLS_EXPOSED=${exposed} (${exposedTools().join(', ')})
TOOLS_WITHHELD=${TOOL_ROWS.length - exposed} MCP tools + ${caps.filter(r => r.EXPOSE_IN_PRODUCT_MCP !== 'YES').length} of ${caps.length} capabilities
EXPOSURE_GATE=${g.state} (released capability tools: ${g.releasedCapabilityTools.length})
\`\`\`

## What is exposed, and why everything else is withheld

1. **Exposure is an explicit list** (\`EXPOSED_CAPABILITY_TOOLS\` in \`src/server/tools.cjs\`) plus release preconditions;
   there is no allow-all and no config switch: \`config.policy\` with any key is refused at startup.
2. **M11.3b.2 (owner decision, owner channel 2026-10-10)** released exactly the two READ tools \`xlsx_read\` /
   \`docx_read\`: accepted read operations only (pre-dispatch refusal of PARTIAL / NOT_MEASURED / INTERNAL / WRITE /
   unknown), the requester-own broker route (M11.2), one editor session per process. The exact exposed set is pinned by
   \`test/static/tool-exposure-exact-set-guard.test.cjs\`.
3. **M11.4b.1 (owner decision, owner channel)** adds the XLSX WRITE tool \`xlsx_execute_task\`: exactly the six
   accepted intents, exposed only while the WRITE release gate holds (M11.4a.1/2/3 CLOSED, the exact XLSX allowlist).
4. **GH #228 (PROPOSED STATE, owner decision pending, \`docs/GH228-DOCX-WRITE-RELEASE-GATE-PROPOSAL.md\`)** adds the DOCX WRITE
   tool \`docx_execute_task\` with the accepted DOCX allowlist (\`rewrite_paragraph\`; GH #233 adds \`insert_paragraph_before\` / \`insert_paragraph_after\` (\`docs/GH233-PARAGRAPH-INSERT-PRODUCT-DECISION.md\`), GH #234 adds \`delete_paragraph\` (\`docs/GH234-PARAGRAPH-DELETE-PRODUCT-DECISION.md\`), GH #235 adds \`move_paragraph_before\` / \`move_paragraph_after\` (\`docs/GH235-PARAGRAPH-MOVE-PRODUCT-DECISION.md\`), GH #236 adds \`apply_list\` / \`set_list_level\` / \`set_list_type\` (\`docs/GH236-LIST-APPLY-TYPE-LEVEL-PRODUCT-DECISION.md\`), GH #237 adds \`indent_list_item\` / \`outdent_list_item\` / \`remove_list\` (\`docs/GH237-LIST-INDENT-OUTDENT-REMOVE-PRODUCT-DECISION.md\`)). The WRITE release gate is **per format**: the XLSX
   tool waits only on the XLSX allowlist, the DOCX tool only on the DOCX allowlist. **Every other routed write capability stays withheld.**

## MCP tools

${tbl(TOOL_ROWS)}

## Capabilities (one row per routed intent)

${tbl(caps)}
`);
