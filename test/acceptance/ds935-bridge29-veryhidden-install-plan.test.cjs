'use strict';

/**
 * GH #365 option C U1 (card an internal branch): documentserver-patches/9.3.5.0/euro-office-bridge-v1/
 * rev29-veryhidden-candidate/plan-rev29-veryhidden.py derives ONE install plan carrying the /29 Bridge
 * candidate (/28 + sheet.veryHidden.read/v1) AND four engine REPLACE steps -- the GH #362
 * `_xlnm.Print_Titles` save-fix FIRST (the card's own MI A KESZ item 3: "/29 blokk + a #362 REPLACE +
 * R1-R3"), then R1/R2/R3 -- in a SINGLE input_sha256 -> output_sha256 chain, same precedent as
 * plan-rev28-xlnm-fix.py (internal tracker) which combined the /28 candidate with the SAME #362 fix.
 *
 * This agent has no access to the real unpatched sdk-all.js (every plan-rev*.py script in this tree
 * states the same constraint about itself), so this suite is entirely OFFLINE/SYNTHETIC, mirroring
 * ds935-bridge28-xlnm-fix-install-plan.test.cjs's own approach: it copies the REAL script, the REAL
 * old.txt/new.txt (xlnm fix) and old1..3.txt/new1..3.txt (R1-R3), and the REAL install_bridge.py into a
 * disposable tmp tree that mirrors documentserver-patches/9.3.5.0's own relative layout -- so
 * `Path(__file__).resolve().parent`-based path resolution inside the script behaves exactly as it
 * would in the real tree, but against a small SYNTHETIC fresh-install.json + bundle this suite
 * controls.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const crypto = require('node:crypto');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const { execFileSync } = require('node:child_process');

const REPO_ROOT = path.join(__dirname, '..', '..');
const BRIDGE_V1 = path.join(REPO_ROOT, 'documentserver-patches', '9.3.5.0', 'euro-office-bridge-v1');
const REAL_DIR = path.join(BRIDGE_V1, 'rev29-veryhidden-candidate');
const REAL_SCRIPT = path.join(REAL_DIR, 'plan-rev29-veryhidden.py');
const REAL_XLNM_DIR = path.join(BRIDGE_V1, 'rev27-xlnm-print-titles-candidate');
const REAL_INSTALL_BRIDGE = path.join(REPO_ROOT, 'documentserver-patches', '9.3.4.60', 'install_bridge.py');

const sha = b => crypto.createHash('sha256').update(b).digest('hex');

const readReal = name => fs.readFileSync(path.join(REAL_DIR, name));
const REAL_OLD = [readReal('old1.txt'), readReal('old2.txt'), readReal('old3.txt')];
const REAL_NEW = [readReal('new1.txt'), readReal('new2.txt'), readReal('new3.txt')];
const REAL_XLNM_OLD = fs.readFileSync(path.join(REAL_XLNM_DIR, 'old.txt'));
const REAL_XLNM_NEW = fs.readFileSync(path.join(REAL_XLNM_DIR, 'new.txt'));

/** Builds a disposable tmp tree mirroring documentserver-patches/9.3.5.0's own relative layout, with a
 * SYNTHETIC fresh-install.json (two insert blocks, like the real one) and a SYNTHETIC bundle containing
 * the real xlnm old.txt text once, the three R1-R3 old texts once each, the chart anchor once, and the
 * default anchor once -- in an order where none of the five `old` texts is a substring of another and
 * none occurs inside another's span. */
function buildTree(scriptSrc) {
  const d = fs.mkdtempSync(path.join(os.tmpdir(), 'plan29veryhidden-'));
  const p935 = path.join(d, '9.3.5.0');
  const pBridgeV1 = path.join(p935, 'euro-office-bridge-v1');
  const pRev28 = path.join(pBridgeV1, 'rev28-table-autofilter-criteria-candidate');
  const pRev29 = path.join(pBridgeV1, 'rev29-veryhidden-candidate');
  const pRev27xlnm = path.join(pBridgeV1, 'rev27-xlnm-print-titles-candidate');
  const p934 = path.join(d, '9.3.4.60');
  for (const p of [pRev28, pRev29, pRev27xlnm, p934]) fs.mkdirSync(p, { recursive: true });

  // the REAL old.txt/new.txt (xlnm fix) and old1..3.txt/new1..3.txt (R1-R3) -- not retyped, copied
  // verbatim from their real locations, so the fix text this test exercises is the SAME text the
  // real candidates carry.
  fs.writeFileSync(path.join(pRev27xlnm, 'old.txt'), REAL_XLNM_OLD);
  fs.writeFileSync(path.join(pRev27xlnm, 'new.txt'), REAL_XLNM_NEW);
  for (let i = 0; i < 3; i += 1) {
    fs.writeFileSync(path.join(pRev29, `old${i + 1}.txt`), REAL_OLD[i]);
    fs.writeFileSync(path.join(pRev29, `new${i + 1}.txt`), REAL_NEW[i]);
  }

  // the REAL install_bridge.py, for the script's own cross-check import.
  fs.copyFileSync(REAL_INSTALL_BRIDGE, path.join(p934, 'install_bridge.py'));

  // the /28 base block is referenced by the fresh-install.json's bridge-block entry BEFORE the
  // script swaps it to /29 -- its content is irrelevant to this test, only the /29 block's bytes
  // (what the script actually inserts) must be stable.
  const bridge29Bytes = Buffer.from('BRIDGE29_VERYHIDDEN_BLOCK();');
  fs.writeFileSync(path.join(pRev29, 'bridge-block.js'), bridge29Bytes);

  const chartAnchor = 'AscBuilder.Cell.init=function(';
  const defaultAnchor = 'Api.GetSheets=function(){';
  const chartBytes = Buffer.from('CHART_READ_BLOCK();');
  fs.mkdirSync(path.join(p935, 'euro-office-chart-read-v1'), { recursive: true });
  fs.writeFileSync(path.join(p935, 'euro-office-chart-read-v1', 'chart-read-block.js'), chartBytes);

  // the bundle contains each real `old` text EXACTLY ONCE, in a plausible order (xlnm fix, then
  // R1/R2/R3 -- matching the card's own stated plan order), each separated by filler so no `old`
  // text is a substring of another's filler-adjacent span.
  const bundle = Buffer.concat([
    Buffer.from('a;'), REAL_XLNM_OLD, Buffer.from(';b;'),
    REAL_OLD[0], Buffer.from(';c;'),
    REAL_OLD[1], Buffer.from(';d;'),
    REAL_OLD[2], Buffer.from(';e;'),
    Buffer.from(`${chartAnchor}z;${defaultAnchor}w;`),
  ]);
  fs.writeFileSync(path.join(p935, 'fresh-install.json'), JSON.stringify({
    cell: {
      file: 'sdkjs/cell/sdk-all.js',
      input_sha256: sha(bundle),
      anchor: defaultAnchor,
      blocks: [
        { file: 'euro-office-chart-read-v1/chart-read-block.js', sha256: sha(chartBytes), layer: 'euro-office-chart-read-v1/9.3.5.0/1', anchor: chartAnchor },
        { file: 'euro-office-bridge-v1/rev28-table-autofilter-criteria-candidate/bridge-block.js', sha256: sha(Buffer.from('ignored-placeholder')), layer: 'euro-office-bridge-v1/9.3.5.0/28' },
      ],
      output_sha256: 'deadbeef'.repeat(8),
      patchVersion: 'euro-office-bridge-v1/9.3.5.0/28',
    },
  }, null, 2));

  const scriptPath = path.join(pRev29, 'plan-rev29-veryhidden.py');
  fs.writeFileSync(scriptPath, scriptSrc);

  const inputPath = path.join(d, 'in.js');
  fs.writeFileSync(inputPath, bundle);
  const outPath = path.join(d, 'out.json');

  return { d, scriptPath, inputPath, outPath, bundle, bridge29Bytes, chartBytes, chartAnchor, defaultAnchor };
}

function run(scriptPath, inputPath, outPath) {
  try {
    const stdout = execFileSync('python3', [scriptPath, '--input', inputPath, '--out', outPath], { encoding: 'utf8' });
    return { status: 0, stdout, stderr: '' };
  } catch (e) {
    return { status: e.status, stdout: e.stdout ? e.stdout.toString() : '', stderr: e.stderr ? e.stderr.toString() : String(e) };
  }
}

test('plan-rev29-veryhidden.py: the derived plan carries the /29 bridge block, the FOUR replace steps in order (xlnm, R1, R2, R3), and the pre-existing chart-read insert untouched -- one input_sha256 -> output_sha256 chain, cross-checked against install_bridge.py', () => {
  const scriptSrc = fs.readFileSync(REAL_SCRIPT, 'utf8');
  const { scriptPath, inputPath, outPath, bundle, bridge29Bytes, chartBytes, chartAnchor, defaultAnchor } = buildTree(scriptSrc);
  const r = run(scriptPath, inputPath, outPath);
  assert.equal(r.status, 0, `stdout=${r.stdout} stderr=${r.stderr}`);

  const printed = JSON.parse(r.stdout);
  assert.equal(printed.patchVersion, 'euro-office-bridge-v1/9.3.5.0/29');

  const plan = JSON.parse(fs.readFileSync(outPath, 'utf8'));
  const blocks = plan.cell.blocks;
  assert.equal(blocks.length, 6, '4 replace steps + chart-read insert + /29 bridge insert');
  assert.deepEqual(blocks[0], { old: REAL_XLNM_OLD.toString('utf8'), new: REAL_XLNM_NEW.toString('utf8') }, 'the xlnm (#362) replace step is FIRST, per the card\'s own stated order');
  assert.deepEqual(blocks[1], { old: REAL_OLD[0].toString('utf8'), new: REAL_NEW[0].toString('utf8') }, 'R1 is SECOND');
  assert.deepEqual(blocks[2], { old: REAL_OLD[1].toString('utf8'), new: REAL_NEW[1].toString('utf8') }, 'R2 is THIRD');
  assert.deepEqual(blocks[3], { old: REAL_OLD[2].toString('utf8'), new: REAL_NEW[2].toString('utf8') }, 'R3 is FOURTH');
  assert.equal(blocks[4].file, 'euro-office-chart-read-v1/chart-read-block.js', 'the pre-existing chart-read insert is untouched');
  assert.equal(blocks[5].file, 'euro-office-bridge-v1/rev29-veryhidden-candidate/bridge-block.js');
  assert.equal(blocks[5].layer, 'euro-office-bridge-v1/9.3.5.0/29');
  assert.equal(plan.cell.patchVersion, 'euro-office-bridge-v1/9.3.5.0/29');

  // the actual output bytes: all 4 replaces applied in order, THEN both inserts, in anchor order.
  let afterReplace = bundle.toString('utf8');
  for (const [o, n] of [[REAL_XLNM_OLD, REAL_XLNM_NEW], [REAL_OLD[0], REAL_NEW[0]], [REAL_OLD[1], REAL_NEW[1]], [REAL_OLD[2], REAL_NEW[2]]]) {
    afterReplace = afterReplace.replace(o.toString('utf8'), n.toString('utf8'));
  }
  const expectedOut = Buffer.concat([
    Buffer.from(afterReplace.slice(0, afterReplace.indexOf(chartAnchor))),
    chartBytes, Buffer.from(`${chartAnchor}${afterReplace.slice(afterReplace.indexOf(chartAnchor) + chartAnchor.length, afterReplace.indexOf(defaultAnchor))}`),
    bridge29Bytes, Buffer.from(`${defaultAnchor}${afterReplace.slice(afterReplace.indexOf(defaultAnchor) + defaultAnchor.length)}`),
  ]);
  assert.equal(plan.cell.output_sha256, sha(expectedOut), 'output_sha256 matches the actual 4x-replace+insert result');
  assert.equal(sha(bundle), plan.cell.input_sha256);
});

test('plan-rev29-veryhidden.py: REFUSES (REPLACE_OLD_NOT_UNIQUE) and writes NOTHING when R2\'s fix text is absent from the bundle', () => {
  const scriptSrc = fs.readFileSync(REAL_SCRIPT, 'utf8');
  const { scriptPath, inputPath, outPath, chartAnchor, defaultAnchor } = buildTree(scriptSrc);
  const d = path.dirname(inputPath);
  // drop R2's old text only -- the other three remain, proving the refusal is step-specific, not a
  // blanket "something is missing" check.
  const bundleNoR2 = Buffer.concat([
    Buffer.from('a;'), REAL_XLNM_OLD, Buffer.from(';b;'),
    REAL_OLD[0], Buffer.from(';c;NO_R2_HERE;d;'),
    REAL_OLD[2], Buffer.from(';e;'),
    Buffer.from(`${chartAnchor}z;${defaultAnchor}w;`),
  ]);
  fs.writeFileSync(inputPath, bundleNoR2);
  const plan935Path = path.join(d, '9.3.5.0', 'fresh-install.json');
  const plan935 = JSON.parse(fs.readFileSync(plan935Path, 'utf8'));
  plan935.cell.input_sha256 = sha(bundleNoR2);
  fs.writeFileSync(plan935Path, JSON.stringify(plan935));

  const r = run(scriptPath, inputPath, outPath);
  assert.equal(r.status, 2, r.stdout + r.stderr);
  assert.match(r.stderr, /REPLACE_OLD_NOT_UNIQUE/);
  assert.equal(fs.existsSync(outPath), false, 'nothing written on refusal');
});

test('RED CONTROL: plan-rev29-veryhidden.py REFUSES (cross-check RESULT_MISMATCH) if the replace-prepend line is removed from derive() -- proving this is not a silently-passing no-op mutation', () => {
  const realSrc = fs.readFileSync(REAL_SCRIPT, 'utf8');
  const marker = 'cell["blocks"] = replace_blocks + cell["blocks"]';
  assert.ok(realSrc.includes(marker), 'the real script must still carry this exact line (fixture assumption)');
  const brokenSrc = realSrc.replace(marker, 'pass  # RED CONTROL: replace steps removed from the plan');
  assert.notEqual(brokenSrc, realSrc);

  const { scriptPath, inputPath, outPath } = buildTree(brokenSrc);
  const r = run(scriptPath, inputPath, outPath);
  assert.equal(r.status, 2, `expected the cross-check to refuse: stdout=${r.stdout} stderr=${r.stderr}`);
  assert.match(r.stderr, /plan_result\(\) disagrees \(status=RESULT_MISMATCH\)/);
  assert.equal(fs.existsSync(outPath), false, 'nothing written when the cross-check disagrees');
});
