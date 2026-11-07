'use strict';

/**
 * Formula text for the XLSX `cells` family (M6 wave 2). Pure.
 *
 * The editor does not store formula text as submitted; it re-renders it. MEASURED (wave-2 probe set a,
 * run m6-w2p-1790582574594-971df7, docs/evidence/m6-w2-formula-probe-a-20260928.json):
 *   '=J1*3' -> '= J1*3'   (a '= ' prefix on every formula)
 *   '=j1*4' -> '= J1*4'   (cell references upper-cased)
 *   '=J1 * 5' -> '= J1*5' (no spaces around operators)
 * So identity is compared on a canonical form -- but ONLY where it was measured. String literals
 * ("...") and quoted sheet names ('...'!) are kept VERBATIM (owner rule: a general rule must not touch
 * their case or spaces). Set b (run m6-w2b-1790582946957-e87e73) measured literals unchanged, and a
 * quoted name that needs no quotes read back unquoted; docs/evidence/M6-W2-FORMULA-PROBE.md.
 */

/** Tokens: 'str' = a string literal incl. quotes, 'qname' = a quoted sheet name incl. quotes, 'other'. */
function tokenize(formula) {
  const s = String(formula);
  const out = [];
  let other = '';
  const flush = () => { if (other) { out.push({ kind: 'other', text: other }); other = ''; } };
  for (let i = 0; i < s.length;) {
    const ch = s[i];
    if (ch === '"' || ch === "'") {
      let j = i + 1;
      for (;;) {
        if (j >= s.length) return null; // unterminated
        if (s[j] === ch) { if (s[j + 1] === ch) { j += 2; continue; } break; }
        j += 1;
      }
      flush();
      out.push({ kind: ch === '"' ? 'str' : 'qname', text: s.slice(i, j + 1) });
      i = j + 1;
    } else { other += ch; i += 1; }
  }
  flush();
  return out;
}

/** A sheet name that needs no quotes: an identifier that is not itself cell-reference shaped. */
const bareSheetName = n => /^[A-Za-z_\u00C0-\uFFFF][\w.\u00C0-\uFFFF]*$/.test(n) && !/^[A-Za-z]{1,3}\d+$/.test(n);

/**
 * The measured canonical form: '=' prefix without spaces; text outside literals upper-cased and
 * space-free; a quoted sheet name that needs no quotes compared unquoted (set b: ='Monthly_Plan'!B16
 * reads back = Monthly_Plan!B16); a name that NEEDS quotes stays verbatim (not measured -> a
 * re-rendering there fails closed).
 */
function canonical(formula) {
  return render(formula, null);
}

/**
 * canonical(), with every reference to sheet `from` rewritten to `to` (bare or quoted, outside string
 * literals). The expected text of a formula after a rename: the wave-3 probe measured that a referencing
 * formula follows SetName (run m6-w3p-1790584404615-f856e6: '= M6W3_f856e6!A1*2' -> '= M6W3b_f856e6!A1*2').
 */
const canonicalRenamed = (formula, from, to) => render(formula, { from: String(from).toUpperCase(), to: String(to).toUpperCase(), raw: String(to) });

/**
 * M7 R26c (tier D probe, docs/evidence/M7-D-PROBE.md): a name that needs quotes ('Kiadás 1') is written back quoted and
 * VERBATIM after a rename (= 'Kiadás x'!A1+1, session = reopen); a bare-able name upper-cased, unquoted (wave 3).
 */
const renamedRef = rename => (bareSheetName(rename.raw) ? rename.to : `'${rename.raw.replace(/'/g, "''")}'`);

function render(formula, rename) {
  const s = String(formula);
  if (!s.startsWith('=')) return null;
  const toks = tokenize(s.slice(1));
  if (!toks) return null;
  const re = rename && new RegExp(`(^|[^A-Z0-9_.\\u00C0-\\uFFFF])${rename.from.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}!`, 'g');
  return '=' + toks.map((t, i) => {
    if (t.kind === 'other') {
      const up = t.text.replace(/\s+/g, '').toUpperCase();
      return re ? up.replace(re, `$1${renamedRef(rename)}!`) : up;
    }
    const next = toks[i + 1];
    if (t.kind === 'qname' && next && next.kind === 'other' && next.text.startsWith('!')) {
      const inner = t.text.slice(1, -1).replace(/''/g, "'");
      if (rename && inner.toUpperCase() === rename.from) return renamedRef(rename);
      if (bareSheetName(inner)) return inner.toUpperCase();
    }
    return t.text;
  }).join('');
}

/** Same formula after the editor's measured re-rendering? */
const sameFormula = (submitted, readBack) => {
  const a = canonical(submitted);
  const b = canonical(readBack);
  return a !== null && a === b;
};

/**
 * GH #298 C: a [book] reference outside literals -- '[...]' directly followed by an (optional) sheet / name and '!'
 * ([Book2.xlsx]Data!A1, [1]Sheet1!A1, [book]!Name). Any OTHER '[' is a structured (table) reference: Table1[Qty],
 * Table1[[#This Row],[Qty]], [@Qty] -- the #298 matrix ST-09 measured the old "every '[' is a book" rule refusing them.
 */
const BOOK_REF = /\[[^[\]]*\][\w.\u00C0-\uFFFF]*!/;

/**
 * Sheet names a formula refers to: a quoted name ('My Sheet'!A1, '' unescaped) or a bare name
 * directly before '!' outside literals. `external` is true for a [book] reference; `structured` for a structured
 * (table) reference -- it names no sheet, and which sheet its table lives on is not read (no table reader, #298 A).
 */
function sheetRefs(formula) {
  const toks = tokenize(String(formula).replace(/^=/, ''));
  if (!toks) return { ok: false, names: [], external: false, structured: false };
  const names = [];
  let external = false;
  let structured = false;
  toks.forEach((t, i) => {
    const next = toks[i + 1];
    if (t.kind === 'qname' && next && next.kind === 'other' && next.text.startsWith('!')) {
      const inner = t.text.slice(1, -1).replace(/''/g, "'");
      if (/^\[/.test(inner)) external = true;
      names.push(inner.replace(/^\[[^\]]*\]/, ''));
    }
    if (t.kind === 'other') {
      if (BOOK_REF.test(t.text)) external = true;
      else if (/\[/.test(t.text)) structured = true;
      // '#REF!' is an error value, never a sheet named REF (a broken reference: dependencies.cjs drops the whole token)
      for (const m of t.text.matchAll(/([A-Za-z_À-￿][\w.À-￿]*)!/g)) if (t.text[m.index - 1] !== '#') names.push(m[1]);
    }
  });
  return { ok: true, names: [...new Set(names)], external, structured };
}

const escapeRx = s => String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/**
 * GH #332/#333 (owner decision A): `formula` with every STRUCTURED (table) reference to `from` renamed
 * to `to` (`From[` -> `To[`, case-insensitive, word-boundary on the left -- the same boundary
 * `table-structure.cjs`'s own `structuredRefs()` detection uses), OUTSIDE string literals and quoted
 * sheet names (kept VERBATIM, via tokenize() -- the same protection canonicalRenamed gives sheet names).
 * Unlike canonical()/canonicalRenamed(), this does NOT re-render the rest of the formula (no upper-
 * casing, no whitespace removal outside the match): the engine re-renders on its own read-back
 * (file header comment above), so a minimal, targeted substitution is enough and safer to reason about.
 *
 * Returns `null` for a non-formula string (no leading '=') or an unparsable one (unterminated
 * literal). `{ formula, changed, sawInLiteral }`:
 *   changed       true iff at least one `From[` occurrence outside a literal was rewritten
 *   sawInLiteral  true iff `From[` ALSO occurs inside a string literal or quoted sheet name (e.g.
 *                 INDIRECT("From[Qty]")) -- the rewrite cannot reach it; the CALLER treats this as
 *                 unsafe to rename at all, not as a partial success
 * A ref with no match anywhere in the 'other' tokens (e.g. an in-table `[@Qty]` relative reference,
 * which never embeds the table name) comes back `changed: false` -- nothing needed rewriting there.
 */
function tableRenameRewrite(formula, from, to) {
  const s = String(formula);
  if (!s.startsWith('=')) return null;
  const toks = tokenize(s.slice(1));
  if (!toks) return null;
  const FROM = escapeRx(String(from));
  const refRe = () => new RegExp(`(^|[^A-Za-z0-9_.\\u00C0-\\uFFFF])${FROM}(\\s*\\[)`, 'gi');
  const litRe = () => new RegExp(`${FROM}\\s*\\[`, 'i');
  let changed = false;
  let sawInLiteral = false;
  const rebuilt = toks.map(t => {
    if (t.kind !== 'other') {
      if (litRe().test(t.text)) sawInLiteral = true;
      return t.text;
    }
    if (!refRe().test(t.text)) return t.text;
    changed = true;
    return t.text.replace(refRe(), (_, pre) => `${pre}${to}[`);
  }).join('');
  return { formula: changed ? '=' + rebuilt : s, changed, sawInLiteral };
}

module.exports = { tokenize, canonical, canonicalRenamed, sameFormula, sheetRefs, bareSheetName, tableRenameRewrite };
