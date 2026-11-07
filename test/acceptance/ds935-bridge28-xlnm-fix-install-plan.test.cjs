'use strict';

/**
 * internal tracker: documentserver-patches/9.3.5.0/euro-office-bridge-v1/rev28-table-autofilter-criteria-
 * candidate/plan-rev28-xlnm-fix.py derives ONE install plan carrying BOTH the /28 Bridge candidate
 * (GH #361+#363) AND the GH #362 `_xlnm.Print_Titles` save-fix REPLACE step, in a SINGLE input_sha256 ->
 * output_sha256 chain -- so a single DS outage installs all three, instead of the /28 candidate silently
 * dropping the #362 fix that today lives ONLY in the separate /27 plan (plan-rev27-xlnm-fix.py).
 *
 * This agent has no access to the real unpatched sdk-all.js (plan-rev28-table-autofilter-criteria.py and
 * plan-rev27-xlnm-fix.py both state the same constraint about themselves), so this suite is entirely
 * OFFLINE/SYNTHETIC: it copies the REAL script (plan-rev28-xlnm-fix.py), the REAL old.txt/new.txt, and the
 * REAL install_bridge.py into a disposable tmp tree that mirrors documentserver-patches/9.3.5.0's own
 * relative layout -- so `Path(__file__).resolve().parent`-based path resolution inside the script behaves
 * exactly as it would in the real tree, but against a small SYNTHETIC fresh-install.json + bundle this
 * suite controls. No DS, no python package beyond the stdlib the scripts already use.
 *
 * RED CONTROL (the card's own MI A KESZ): a COPY of the script with the replace-step line removed from
 * derive() must refuse to apply the fix -- proving this suite's own PASS is not a tautology (see "RED
 * CONTROL" test below).
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { execFileSync } = require('node:child_process');

const REPO_ROOT = path.join(__dirname, '..', '..');
const REAL_DIR = path.join(REPO_ROOT, 'documentserver-patches', '9.3.5.0', 'euro-office-bridge-v1', 'rev28-table-autofilter-criteria-candidate');
const REAL_SCRIPT = path.join(REAL_DIR, 'plan-rev28-xlnm-fix.py');
const REAL_XLNM_DIR = path.join(REPO_ROOT, 'documentserver-patches', '9.3.5.0', 'euro-office-bridge-v1', 'rev27-xlnm-print-titles-candidate');
const REAL_INSTALL_BRIDGE = path.join(REPO_ROOT, 'documentserver-patches', '9.3.4.60', 'install_bridge.py');

const sha = b => crypto.createHash('sha256').update(b).digest('hex');

/** Builds a disposable tmp tree mirroring documentserver-patches/9.3.5.0's own relative layout, with a
 * SYNTHETIC fresh-install.json (two insert blocks, like the real one: a chart-read block at its own
 * anchor, a bridge block at the default anchor) and a SYNTHETIC bundle containing the real old.txt text
 * once, the chart anchor once, and the default anchor once. Returns paths needed to invoke the script. */
function buildTree(scriptSrc) {
  const d = fs.mkdtempSync(path.join(os.tmpdir(), 'plan28xlnm-'));
  const p935 = path.join(d, '9.3.5.0');
  const pBridgeV1 = path.join(p935, 'euro-office-bridge-v1');
  const pRev28 = path.join(pBridgeV1, 'rev28-table-autofilter-criteria-candidate');
  const pRev27xlnm = path.join(pBridgeV1, 'rev27-xlnm-print-titles-candidate');
  const p934 = path.join(d, '9.3.4.60');
  for (const p of [pRev28, pRev27xlnm, p934]) fs.mkdirSync(p, { recursive: true });

  // the REAL old.txt/new.txt -- not retyped, copied verbatim from their real location, so the fix text this
  // test exercises is the SAME text the real candidate carries.
  const oldText = fs.readFileSync(path.join(REAL_XLNM_DIR, 'old.txt'));
  const newText = fs.readFileSync(path.join(REAL_XLNM_DIR, 'new.txt'));
  fs.writeFileSync(path.join(pRev27xlnm, 'old.txt'), oldText);
  fs.writeFileSync(path.join(pRev27xlnm, 'new.txt'), newText);

  // the REAL install_bridge.py, for the script's own cross-check import.
  fs.copyFileSync(REAL_INSTALL_BRIDGE, path.join(p934, 'install_bridge.py'));

  // a synthetic /28 bridge block -- content is irrelevant to this test, only its bytes must be stable.
  const bridge28Bytes = Buffer.from('BRIDGE28_BLOCK();');
  fs.writeFileSync(path.join(pRev28, 'bridge-block.js'), bridge28Bytes);

  const chartAnchor = 'AscBuilder.Cell.init=function(';
  const defaultAnchor = 'Api.GetSheets=function(){';
  const chartBytes = Buffer.from('CHART_READ_BLOCK();');
  fs.mkdirSync(path.join(p935, 'euro-office-chart-read-v1'), { recursive: true });
  fs.writeFileSync(path.join(p935, 'euro-office-chart-read-v1', 'chart-read-block.js'), chartBytes);

  const bundle = Buffer.from(`x;${oldText.toString('utf8')}y;${chartAnchor}z;${defaultAnchor}w;`);
  fs.writeFileSync(path.join(p935, 'fresh-install.json'), JSON.stringify({
    cell: {
      file: 'sdkjs/cell/sdk-all.js',
      input_sha256: sha(bundle),
      anchor: defaultAnchor,
      blocks: [
        { file: 'euro-office-chart-read-v1/chart-read-block.js', sha256: sha(chartBytes), layer: 'euro-office-chart-read-v1/9.3.5.0/1', anchor: chartAnchor },
        { file: 'euro-office-bridge-v1/bridge-block.js', sha256: sha(Buffer.from('ignored-placeholder')), layer: 'euro-office-bridge-v1/9.3.5.0/16' },
      ],
      output_sha256: 'deadbeef'.repeat(8),
      patchVersion: 'euro-office-bridge-v1/9.3.5.0/16',
    },
  }, null, 2));

  const scriptPath = path.join(pRev28, 'plan-rev28-xlnm-fix.py');
  fs.writeFileSync(scriptPath, scriptSrc);

  const inputPath = path.join(d, 'in.js');
  fs.writeFileSync(inputPath, bundle);
  const outPath = path.join(d, 'out.json');

  return { d, scriptPath, inputPath, outPath, bundle, oldText, newText, bridge28Bytes, chartBytes, defaultAnchor, chartAnchor };
}

function run(scriptPath, inputPath, outPath) {
  try {
    const stdout = execFileSync('python3', [scriptPath, '--input', inputPath, '--out', outPath], { encoding: 'utf8' });
    return { status: 0, stdout, stderr: '' };
  } catch (e) {
    return { status: e.status, stdout: e.stdout ? e.stdout.toString() : '', stderr: e.stderr ? e.stderr.toString() : String(e) };
  }
}

test('plan-rev28-xlnm-fix.py: the derived plan carries the /28 bridge block, the ONE replace step FIRST, and the pre-existing chart-read insert untouched -- one input_sha256 -> output_sha256 chain, cross-checked against install_bridge.py', () => {
  const scriptSrc = fs.readFileSync(REAL_SCRIPT, 'utf8');
  const { scriptPath, inputPath, outPath, bundle, oldText, newText, bridge28Bytes, chartBytes, defaultAnchor, chartAnchor } = buildTree(scriptSrc);
  const r = run(scriptPath, inputPath, outPath);
  assert.equal(r.status, 0, `stdout=${r.stdout} stderr=${r.stderr}`);

  const printed = JSON.parse(r.stdout);
  assert.equal(printed.patchVersion, 'euro-office-bridge-v1/9.3.5.0/28');

  const plan = JSON.parse(fs.readFileSync(outPath, 'utf8'));
  const blocks = plan.cell.blocks;
  assert.equal(blocks.length, 3, 'replace step + chart-read insert + /28 bridge insert');
  assert.deepEqual(blocks[0], { old: oldText.toString('utf8'), new: newText.toString('utf8') }, 'the replace step is FIRST in plan order');
  assert.equal(blocks[1].file, 'euro-office-chart-read-v1/chart-read-block.js', 'the pre-existing chart-read insert is untouched');
  assert.equal(blocks[2].file, 'euro-office-bridge-v1/rev28-table-autofilter-criteria-candidate/bridge-block.js');
  assert.equal(blocks[2].layer, 'euro-office-bridge-v1/9.3.5.0/28');
  assert.equal(plan.cell.patchVersion, 'euro-office-bridge-v1/9.3.5.0/28');

  // the actual output bytes: replace applied, THEN both inserts, in anchor order.
  const expectedOut = Buffer.concat([
    Buffer.from(`x;${newText.toString('utf8')}y;`),
    chartBytes, Buffer.from(`${chartAnchor}z;`),
    bridge28Bytes, Buffer.from(`${defaultAnchor}w;`),
  ]);
  assert.equal(plan.cell.output_sha256, sha(expectedOut), 'output_sha256 matches the actual replace+insert result');
  assert.equal(sha(bundle), plan.cell.input_sha256);
});

test('plan-rev28-xlnm-fix.py: REFUSES (REPLACE_OLD_NOT_UNIQUE) and writes NOTHING when the fix text is absent from the bundle', () => {
  const scriptSrc = fs.readFileSync(REAL_SCRIPT, 'utf8');
  const { scriptPath, inputPath, outPath, defaultAnchor, chartAnchor } = buildTree(scriptSrc);
  // overwrite the bundle+plan with one that has NO old.txt occurrence at all.
  const d = path.dirname(inputPath);
  const bundleNoFix = Buffer.from(`x;y;${chartAnchor}z;${defaultAnchor}w;`);
  fs.writeFileSync(inputPath, bundleNoFix);
  const plan935 = JSON.parse(fs.readFileSync(path.join(d, '9.3.5.0', 'fresh-install.json'), 'utf8'));
  plan935.cell.input_sha256 = sha(bundleNoFix);
  fs.writeFileSync(path.join(d, '9.3.5.0', 'fresh-install.json'), JSON.stringify(plan935));

  const r = run(scriptPath, inputPath, outPath);
  assert.equal(r.status, 2, r.stdout + r.stderr);
  assert.match(r.stderr, /REPLACE_OLD_NOT_UNIQUE/);
  assert.equal(fs.existsSync(outPath), false, 'nothing written on refusal');
});

test('RED CONTROL: plan-rev28-xlnm-fix.py REFUSES (cross-check RESULT_MISMATCH) if the replace-prepend line is removed from derive() -- proving this is not a silently-passing no-op mutation', () => {
  // a naive removal of the one line that prepends the replace entry to cell["blocks"] does NOT silently
  // produce an insert-only plan: transform() still unconditionally folds the replace into the output bytes
  // (it receives old_text/new_text directly, not read back from cell["blocks"]), so the resulting bytes and
  // the now-replace-less plan.blocks disagree -- and install_bridge.py's own plan_result() cross-check (which
  // DOES derive its replace steps from plan["blocks"]) catches the mismatch and refuses. This is the FIRST,
  // cheaper way a naive mutation surfaces: the script never even gets a chance to write a plan missing the fix.
  const realSrc = fs.readFileSync(REAL_SCRIPT, 'utf8');
  const marker = '    cell["blocks"] = [{"old": old_text.decode("utf-8"), "new": new_text.decode("utf-8")}] + cell["blocks"]';
  assert.ok(realSrc.includes(marker), 'the real script must still carry this exact line (fixture assumption)');
  const brokenSrc = realSrc.replace(marker, '    pass  # RED CONTROL: replace step removed');
  assert.notEqual(brokenSrc, realSrc);

  const { scriptPath, inputPath, outPath } = buildTree(brokenSrc);
  const r = run(scriptPath, inputPath, outPath);
  assert.equal(r.status, 2, `expected the cross-check to refuse: stdout=${r.stdout} stderr=${r.stderr}`);
  assert.match(r.stderr, /plan_result\(\) disagrees \(status=RESULT_MISMATCH\)/);
  assert.equal(fs.existsSync(outPath), false, 'nothing written when the cross-check disagrees');
});

test('RED CONTROL, SECOND FORM (the card\'s own wording -- "take the replace step out of the plan"): the /28-only sibling script (plan-rev28-table-autofilter-criteria.py, no xlnm awareness at all) derives a plan with NO replace step from the SAME bundle -- this is today\'s real regression, which plan-rev28-xlnm-fix.py exists to close', () => {
  const siblingSrc = fs.readFileSync(path.join(REAL_DIR, 'plan-rev28-table-autofilter-criteria.py'), 'utf8');
  assert.ok(/def derive\(base, data, read_block\):/.test(siblingSrc), 'fixture assumption: the sibling script has no replace-step parameters at all');

  const { d, bundle } = buildTree(fs.readFileSync(REAL_SCRIPT, 'utf8'));
  const siblingPath = path.join(d, '9.3.5.0', 'euro-office-bridge-v1', 'rev28-table-autofilter-criteria-candidate', 'plan-rev28-table-autofilter-criteria.py');
  fs.writeFileSync(siblingPath, siblingSrc);
  const inputPath = path.join(d, 'in2.js');
  fs.writeFileSync(inputPath, bundle);
  const outPath = path.join(d, 'out2.json');

  const r = run(siblingPath, inputPath, outPath);
  assert.equal(r.status, 0, `sibling script should succeed (it never attempted the #362 fix): stdout=${r.stdout} stderr=${r.stderr}`);
  const plan = JSON.parse(fs.readFileSync(outPath, 'utf8'));
  const hasReplace = plan.cell.blocks.some(b => 'old' in b);
  assert.equal(hasReplace, false, 'RED, by design: this sibling plan is insert-only -- installing it alone silently drops the #362 fix, exactly the outage internal tracker closes');
  // and the /28 insert itself is still correct on this sibling -- confirming the difference is PURELY the
  // missing replace step, not some other divergence between the two scripts.
  const bridgeEntry = plan.cell.blocks.find(b => b.file && b.file.includes('rev28-table-autofilter-criteria-candidate'));
  assert.ok(bridgeEntry, 'the /28 bridge block is present');
  assert.equal(bridgeEntry.layer, 'euro-office-bridge-v1/9.3.5.0/28');
});
