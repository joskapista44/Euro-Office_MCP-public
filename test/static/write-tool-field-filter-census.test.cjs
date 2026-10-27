'use strict';

/**
 * Offline census (an owner-review NEM-MERT, closed here): every xlsx/docx family's own declared
 * top-level step field x the WRITE tool's two by-name refusal sets, CREDENTIAL_FIELD and
 * ROUTE_OVERRIDE (write-tools.cjs's `checkNamesDeep`, the SAME function the tool calls -- never a
 * hand-copied pattern list). `checkNamesDeep` walks every key at every depth of the whole `args`
 * object, so a family's own, legitimate step-level field name can collide with it. Two already did
 * (X-L13/1 LIVE r1): sort_range's `key` (CREDENTIAL_FIELD) and create_pivot's `host`
 * (ROUTE_OVERRIDE) -- the owner's fix was a field rename (sort_range: `key` -> `keyColumn`;
 * create_pivot: `host` -> `pivotSheet`; the write-tool filters themselves stayed untouched), so the
 * collision set below is now empty. This file is the regression gate going forward: if a NEW
 * collision appears (a new family field, or a filter-set addition, lands on an existing field), the
 * exact-match test below fails -- the next fix (intent-aware filter vs. another field rename) is a
 * contract decision for the owner, not this file.
 *
 * No PPTX family exists yet in this codebase (checked: src/adapters/pptx does not exist).
 *
 * SCOPE, and why some families are listed as NEM-MERT below instead of censused:
 *   - A family's TOP-LEVEL step fields are in scope (this is what collides -- `args.task.steps[i].<field>`
 *     sits at the depth `checkNamesDeep` walks). Fields NESTED inside a step's own sub-object (xlsx
 *     format.cjs's `step.format.*`, docx text/list/rich/table's run- or cell-level style keys) are a
 *     SEPARATE axis: they are validated by each family's own nested-shape check, never reach
 *     `checkNamesDeep` as a bare key at `args`-depth the way a top-level field does, and are NOT
 *     censused here.
 *   - Most families declare their legitimate top-level fields as a module-level `const FIELDS`
 *     (array, or an object keyed by intent for a multi-intent family) -- extracted straight from
 *     each file's own source text below (balanced-bracket scan, never hand-retyped), so the census
 *     number comes from the code.
 *   - A few families (cf, filter, freeze, names, range, validation, style-definition) build their
 *     allowlist as a LOCAL variable instead of a module constant (often per-intent, via a ternary) --
 *     these are read by hand from the source (quoted verbatim, see MANUAL_FIELDS) and included too,
 *     so they are not silently dropped just for not following the FIELDS-constant convention.
 *   - The rest (listed in NEM_MERT_FAMILIES) have no top-level field-allowlist at all, under any
 *     name -- genuinely unmeasured for this census, not elided.
 */

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const { checkNamesDeep, ROUTE_FIELDS } = require('../../src/server/write-tools.cjs');
const { CREDENTIAL_FIELDS } = require('../../src/server/read-tools.cjs');

const XLSX_DIR = path.join(__dirname, '..', '..', 'src', 'adapters', 'xlsx', 'families');
const DOCX_DIR = path.join(__dirname, '..', '..', 'src', 'adapters', 'docx', 'families');
const PPTX_DIR = path.join(__dirname, '..', '..', 'src', 'adapters', 'pptx', 'families');

/** Balanced-bracket scan for a top-level `const <name> = <expr>;` in `src` -- never a regex copy of
 * the expression's own content, so nested references (page.cjs's `...MARGINS`) resolve to the SAME
 * file's own other `const` declarations, not a guessed value. */
function findTopLevelConst(src, name) {
  const re = new RegExp(`(?:^|\\n)const\\s+${name}\\s*=\\s*`);
  const m = re.exec(src);
  if (!m) return undefined;
  const start = m.index + m[0].length;
  let depth = 0;
  let i = start;
  let seenOpen = false;
  for (; i < src.length; i++) {
    const c = src[i];
    if (c === '(' || c === '[' || c === '{') { depth++; seenOpen = true; }
    else if (c === ')' || c === ']' || c === '}') depth--;
    else if (c === ';' && depth === 0 && seenOpen) break;
  }
  return { decl: `const ${name} = ${src.slice(start, i)};`, snippet: src.slice(start, i) };
}

function extractFields(src) {
  const fields = findTopLevelConst(src, 'FIELDS');
  if (!fields) return undefined;
  const idents = new Set((fields.snippet.match(/[A-Za-z_$][A-Za-z0-9_$]*/g) || []));
  const deps = [];
  for (const id of idents) {
    if (id === 'FIELDS' || id === 'Object' || id === 'freeze') continue;
    const d = findTopLevelConst(src, id);
    if (d) deps.push(d.decl);
  }
  return new Function(`${deps.join('\n')}\nconst FIELDS = ${fields.snippet};\nreturn FIELDS;`)();
}

// Array-shaped FIELDS families: file basename -> the REAL intent(s) that array's own unknown-field
// check guards (read from each file's own `step.intent !== '...'` / `INTENTS` declaration; an
// object-shaped FIELDS family already keys by the real intent, no mapping needed here).
const REAL_INTENTS = Object.freeze({
  cells: ['write_range'], clear: ['clear_range'], format: ['format_range'], 'format-run': ['format_run'],
  layout: ['layout_range'], merge: ['merge_range', 'unmerge_range'], 'pivot-refresh': ['refresh_pivot'],
  pivots: ['create_pivot'], 'pivots-v2': ['create_pivot'], 'shape-text': ['set_shape_text'], 'chart-point-labels': ['set_chart_point_data_label'], 'shape-rotation': ['set_shape_rotation'], 'sheet-visibility': ['set_sheet_visibility'], sort: ['sort_range'], page: ['set_page_layout'],
  structure: ['insert_rows', 'delete_rows', 'insert_columns', 'delete_columns'],
  numbering: ['set_numbering'], 'page-setup': ['set_page_setup'], 'table-format': ['format_table'], comments: ['resolve_comment'],
  'delete-cells': ['delete_cells'],
});

// Six families keep their allowlist as a LOCAL (often per-intent, conditional) variable instead of
// a module `FIELDS` const -- quoted VERBATIM from each file (filter.cjs:44, freeze.cjs:31,
// names.cjs:94+130, range.cjs:56, validation.cjs:37, cf.cjs:124, style-definition.cjs:56), so they
// are not dropped from the census just for not following that convention. 'intent'/'id' are already
// excluded below for the FIELDS-const families; they are left out here for the same reason.
const MANUAL_FIELDS = Object.freeze([
  { family: 'cf', intent: 'add_conditional_format', fields: ['sheet', 'range', 'rule'] },
  { family: 'cf', intent: 'delete_conditional_format', fields: ['sheet', 'range', 'rule'] },
  { family: 'filter', intent: 'filter_range', fields: ['sheet', 'range', 'field', 'value', 'values', 'and'] },
  { family: 'filter', intent: 'clear_filter', fields: ['sheet', 'range'] },
  { family: 'freeze', intent: 'freeze_panes', fields: ['sheet', 'rows', 'columns'] },
  { family: 'freeze', intent: 'unfreeze_panes', fields: ['sheet'] },
  { family: 'names', intent: 'set_print_area', fields: ['sheet', 'range', 'areas', 'replaces'] },
  { family: 'names', intent: 'clear_print_area', fields: ['sheet'] },
  { family: 'names', intent: 'set_print_titles', fields: ['sheet', 'rows', 'columns'] },
  { family: 'names', intent: 'clear_print_titles', fields: ['sheet'] },
  { family: 'names', intent: 'set_defined_name', fields: ['name', 'refersTo', 'scope'] },
  { family: 'names', intent: 'delete_defined_name', fields: ['name', 'scope'] },
  { family: 'range', intent: 'copy_range', fields: ['sheet', 'range', 'target', 'targetSheet'] },
  { family: 'range', intent: 'move_range', fields: ['sheet', 'range', 'target', 'targetSheet', 'expect'] },
  { family: 'validation', intent: 'set_validation', fields: ['sheet', 'range', 'type', 'operator', 'alertStyle', 'min', 'max', 'value', 'items', 'listRange', 'formula'] },
  { family: 'validation', intent: 'clear_validation', fields: ['sheet', 'range'] },
  { family: 'style-definition', intent: 'define_style', fields: ['dependsOn', 'name', 'basedOn', 'text', 'paragraph'] },
  { family: 'style-definition', intent: 'modify_style', fields: ['dependsOn', 'name', 'basedOn', 'text', 'paragraph'] },
]);

// No top-level field-allowlist found, under any name, in these families -- unmeasured for this
// census (see the file doc comment's SCOPE section; several of these DO validate a NESTED
// sub-object's own keys, which is a different axis, not this one).
const NEM_MERT_FAMILIES = Object.freeze([
  'format-align-ext', 'format-diagonal', 'format-fill-clear', 'format-font-effects', 'format-maps',
  'format-remove', 'format-rotation', // format.cjs's own nested-shape helpers, not separate families
  'formatting', 'text', 'list', 'rich', 'table', // nested run/cell style keys only (TEXT_KEYS/CELL_KEYS/PRESERVED_PROPERTIES)
  'notes', 'paragraph', 'table-column', 'table-merge-split', 'table-row', 'header-footer',
]);

function censusDir(dir) {
  if (!fs.existsSync(dir)) return { byFamily: [], missing: [] };
  const byFamily = [];
  const missing = [];
  for (const file of fs.readdirSync(dir).filter(f => f.endsWith('.cjs'))) {
    const name = path.basename(file, '.cjs');
    if (NEM_MERT_FAMILIES.includes(name)) continue;
    const src = fs.readFileSync(path.join(dir, file), 'utf8');
    let f;
    try { f = extractFields(src); } catch { f = undefined; }
    if (f === undefined) { missing.push(name); continue; }
    if (Array.isArray(f)) {
      const intents = REAL_INTENTS[name];
      assert.ok(intents, `array FIELDS in ${name}.cjs but no REAL_INTENTS entry -- update this test, not a measurement gap`);
      const entries = {};
      for (const intent of intents) entries[intent] = f;
      byFamily.push({ family: name, entries });
    } else if (f && typeof f === 'object') {
      byFamily.push({ family: name, entries: f });
    }
  }
  for (const m of MANUAL_FIELDS) {
    const existing = byFamily.find(h => h.family === m.family);
    if (existing) existing.entries[m.intent] = m.fields;
    else byFamily.push({ family: m.family, entries: { [m.intent]: m.fields } });
  }
  return { byFamily, missing };
}

function runCensus() {
  const { byFamily: xlsx, missing: xlsxMissing } = censusDir(XLSX_DIR);
  const { byFamily: docx, missing: docxMissing } = censusDir(DOCX_DIR);
  const collisions = [];
  let checked = 0;
  for (const h of [...xlsx, ...docx]) {
    for (const [intent, fieldList] of Object.entries(h.entries)) {
      for (const field of fieldList) {
        if (field === 'intent' || field === 'id') continue;
        checked += 1;
        const r = checkNamesDeep({ task: { id: 'x', steps: [{ intent, [field]: 'probe-value' }] } }, 'args');
        if (r) collisions.push({ family: h.family, intent, field, code: r.code });
      }
    }
  }
  return { collisions, checked, missing: [...xlsxMissing, ...docxMissing] };
}

test('no PPTX family exists yet (if this ever flips to true, this census needs a PPTX pass added)', () => {
  assert.equal(fs.existsSync(PPTX_DIR), false);
});

test('write-tool field-filter census: every xlsx/docx family\'s own top-level step field x CREDENTIAL_FIELD/ROUTE_OVERRIDE, via the real checkNamesDeep', () => {
  const { collisions, checked, missing } = runCensus();

  // sort_range's key -> keyColumn and create_pivot's host -> pivotSheet (the owner's field-rename fix for the two
  // collisions this census originally found) -- confirm the OLD names are gone from the census's own field lists
  // (the rename actually landed in the source this test reads, not just in the filter's blind spot).
  assert.ok(!collisions.some(c => c.intent === 'sort_range' && c.field === 'key'), 'sort_range must no longer declare a bare "key" field');
  assert.ok(!collisions.some(c => c.intent === 'create_pivot' && c.field === 'host'), 'create_pivot must no longer declare a bare "host" field');

  // NEG. KONTROLL: a deliberately harmless, made-up field name, through the SAME real function.
  assert.equal(!!checkNamesDeep({ task: { id: 'x', steps: [{ intent: 'write_range', zsirafMezo: 'probe-value' }] } }, 'args'), false);

  // The REGRESSION GATE: the set is now EMPTY. A new entry here is a NEW finding -- triage it with
  // the owner before updating this list; it is not this test's job to silently grow.
  const got = collisions.map(c => `${c.intent}.${c.field}=${c.code}`).sort();
  assert.deepEqual(got, []);

  assert.ok(checked > 150, `sanity floor on how many (intent,field) pairs got checked, got ${checked}`);
  assert.ok(missing.length > 0, 'the NEM_MERT_FAMILIES list should correspond to real misses, not an empty set');
});

test('CREDENTIAL_FIELDS / ROUTE_FIELDS are read live from the product modules, not copied here', () => {
  assert.ok(CREDENTIAL_FIELDS.has('key'));
  assert.ok(ROUTE_FIELDS.has('host'));
});
