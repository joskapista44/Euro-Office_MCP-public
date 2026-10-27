#!/usr/bin/env python3
"""
GH #293 L7 (owner GO B, 2026-10-22: the list-scoped numbering change, only on a proven, complete affected set) -- OFFLINE build of the
euro-office-bridge-word-v1 patchVersion /5 CANDIDATE on DS 9.3.5.

/5 = /4 + exactly TWO new capabilities:
  docx.numbering.listMembers.read/v1   (kind READ, mutation NO)
  docx.numbering.listDetach.write/v1   (kind WRITE, mutation TARGETED)
Measured behind them (test/live/docx-gh293-l7-detach-probe.live.cjs, docx-gh293-l7-internals-probe.live.cjs;
docs/evidence/gh293-l7-detach-probe-run2-clone-20261022.json, gh293-l7-internals-probe-run1-20261022.json,
gh293-l7-internals-probe-run2-released-table-cell-comember-20261022.json):
  - the PUBLIC read the product has is blind to a list item inside a table cell (the released set_numbering changed one silently);
    the editor's own CDocument.GetAllParagraphsByNumbering(NumPr(numId, null)) sees every paragraph of a numbering, the table cell's too;
  - CNumbering.GetAllNumsByAbstractNum(<the abstractNum OBJECT>) lists every w:num on one definition; CNum.LvlOverride[l].GetStartOverride()
    answers the level's startOverride (null slot = none, also after ClearAllLvlOverride);
  - CNum.Copy() makes a new w:num on a NEW abstractNum equal on all levels incl. the overrides, and Paragraph.SetNumPr(newNumId, lvl)
    repoints an item: a whole list detached that way keeps every label; DS continues a count ACROSS two w:num on one abstractNum when the
    later one has no startOverride (so a detach is only label-preserving when no count can flow between the lists).
The READ gives the WHOLE affected set (every member of the target's numbering, document-wide) and the other lists on the same definition
with their overrides and order; the WRITE re-reads, compares with what the caller saw, re-checks the isolation rule, then detaches.
Every /4 capability stays byte-identical.

Deterministic: the /5 block is derived from the /4 candidate block (../rev4-candidate/bridge-block.js, sha pinned) by FOUR exact
substitutions, each asserted once.

Usage:
  build-rev5.py                          -> writes bridge-block.js, old.txt, new.txt next to this file
  build-rev5.py --sdk <served /4 word sdk-all.js> [--out <file>]
                                         -> also writes manifest.json (whole-block replace, /4 output sha -> /5 sha)
This script never writes a Document Server file.
"""
import hashlib, json, sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
V4_FILE = HERE.parent / "rev4-candidate" / "bridge-block.js"
V4_FILE_SHA = "9ae2b09e78d02028cb7821664141e27aa6d231395238fffc8bedbe7104a757d3"
V4_SDK_SHA = "46f2873d1e10df07bb9a4a651fc3ca2b147bcd68f0c68da1b4c9bd1f8e51af98"   # = rev4-candidate/manifest.json patched_sha256 (served /4)
V4_SDK_SIZE = 18564885

READ_REASONS = ("bad-args, unsupported-property, invalid-level, document-unavailable, search-threw, target-not-found, target-ambiguous, "
                "paragraph-unreadable, getNumbering-threw, level-mismatch, internal-missing, numpr-unreadable, abstract-unreadable, nums-threw, "
                "nums-unreadable, paragraphs-unreadable, override-unreadable, member-unreadable, own-num-not-listed")
WRITE_REASONS = (READ_REASONS + ", not-numbered, members-mismatch, others-mismatch, nothing-to-detach, order-unreadable, not-isolatable, "
                 "copy-threw, set-threw, write-not-observed")

READ_ENTRY = (
    '{ id: "docx.numbering.listMembers.read/v1", version: 1, kind: "READ", format: "DOCX", mutation: "NO", '
    'semanticContract: "one list item by its exact text (the public Search, label excluded, exactly one hit) and its level: the WHOLE '
    'affected set of its numbering -- every paragraph on the same w:num document-wide (the editor\'s numbering collection, table cells '
    'included; the public read does not see those), each {text, level} in document order -- plus every OTHER w:num on the same abstract '
    'definition with its members, the levels its startOverride covers, and whether it has members before / after the target list; the '
    'target\'s own override levels, a per-level signature (format, level text, start, suffix, alignment) of each w:num, and the order flags '
    '{targetFirst, contiguous} (null when the editor cannot order them); not numbered = {present:false}; any unreadable piece fails '
    'closed with a reason, never a partial set; ids are session-local (comparable within one session only); reasons: ' + READ_REASONS +
    ' (GH #293 L7, word /5)" }, '
)
WRITE_ENTRY = (
    '{ id: "docx.numbering.listDetach.write/v1", version: 1, kind: "WRITE", format: "DOCX", mutation: "TARGETED", '
    'semanticContract: "one list by one of its items {text, level, members:[{text, level}] as the caller read them, others: the number '
    'of other w:num on the definition}: re-reads docx.numbering.listMembers.read/v1, refuses when the members or the other lists differ '
    'from what the caller saw, refuses when nothing shares the definition (nothing-to-detach, no mutation), and refuses unless the count is '
    'isolated: the list is contiguous among the definition\'s items, no other list has items both before and after it, a list after it has '
    'a startOverride on every level it uses, and the list itself is first or has a startOverride on every level it uses; then CNum.Copy() '
    '(a new w:num on a NEW abstract definition, equal on every level incl. the overrides) and Paragraph.SetNumPr on EVERY member; re-read: '
    'every member on the new w:num, none left on the old; the result is an operation receipt {applied:[list-detached], before:{numId}, '
    'moved} and NOT a verification -- the caller verifies with a fresh read; no fallback; served only through write(); reasons: ' +
    WRITE_REASONS + ' (GH #293 L7, word /5)" }, '
)

HELPER = (
    'function eoListMembers(args) { '
    'if (!args || typeof args !== "object" || typeof args.text !== "string" || args.text === "") return { ok: false, reason: "bad-args" }; '
    'if (!(typeof args.level === "number" && Math.floor(args.level) === args.level && args.level >= 0 && args.level <= 8)) return { ok: false, reason: "invalid-level" }; '
    'var doc; try { doc = Api.GetDocument(); } catch (e) { return { ok: false, reason: "document-unavailable" }; } '
    'if (!doc || typeof doc.Search !== "function") return { ok: false, reason: "document-unavailable" }; '
    'var hits; try { hits = (doc.Search(args.text, { MatchCase: true }) || []).filter(function (r) { return String(r.GetText({ Numbering: false }) || "") === args.text; }); } catch (e) { return { ok: false, reason: "search-threw" }; } '
    'if (hits.length === 0) return { ok: false, reason: "target-not-found" }; if (hits.length !== 1) return { ok: false, reason: "target-ambiguous", count: hits.length }; '
    'var p; try { p = hits[0].GetParagraph(0); } catch (e) { p = null; } if (!p) return { ok: false, reason: "paragraph-unreadable" }; '
    'var cur; try { cur = typeof p.GetNumbering === "function" ? p.GetNumbering() : null; } catch (e) { return { ok: false, reason: "getNumbering-threw" }; } '
    'if (!cur) return { ok: true, value: { present: false } }; '
    'var lvl = typeof cur.GetLevelIndex === "function" ? cur.GetLevelIndex() : null; if (lvl !== args.level) return { ok: false, reason: "level-mismatch", level: Number.isInteger(lvl) ? lvl : null }; '
    'var ip = p.Paragraph, ld = doc.Document; '
    'if (!ip || typeof ip.GetNumPr !== "function" || typeof ip.SetNumPr !== "function" || !ld || typeof ld.GetNumbering !== "function" || typeof ld.GetAllParagraphsByNumbering !== "function") return { ok: false, reason: "internal-missing" }; '
    'if (typeof AscWord === "undefined" || !AscWord || typeof AscWord.NumPr !== "function" || typeof ApiParagraph !== "function") return { ok: false, reason: "internal-missing" }; '
    'var N = ld.GetNumbering(); if (!N || typeof N.GetNum !== "function" || typeof N.GetAbstractNum !== "function" || typeof N.GetAllNumsByAbstractNum !== "function") return { ok: false, reason: "internal-missing" }; '
    'var np; try { np = ip.GetNumPr(); } catch (e) { np = null; } if (!np || np.NumId === undefined || np.NumId === null) return { ok: false, reason: "numpr-unreadable" }; '
    'var num = N.GetNum(np.NumId); if (!num || typeof num.GetAbstractNumId !== "function" || typeof num.GetId !== "function" || typeof num.Copy !== "function" || typeof num.GetLvl !== "function") return { ok: false, reason: "internal-missing" }; '
    'var abs = N.GetAbstractNum(num.GetAbstractNumId()); if (!abs) return { ok: false, reason: "abstract-unreadable" }; '
    'function textOf(q) { var t; try { t = new ApiParagraph(q).GetText({ Numbering: false }); } catch (e) { return null; } return typeof t === "string" ? t.replace(/[\\r\\n]+$/, "") : null; } '
    'function starts(n) { var r = []; for (var l = 0; l < 9; l += 1) { var o = n.LvlOverride ? n.LvlOverride[l] : null; var s = null; if (o) { try { s = typeof o.GetStartOverride === "function" ? o.GetStartOverride() : o.StartOverride; } catch (e) { return null; } } if (typeof s === "number" && Math.floor(s) === s && s >= 0) r.push(l); } return r; } '
    'function sig(n) { var r = []; for (var l = 0; l < 9; l += 1) { var v; try { v = n.GetLvl(l); } catch (e) { return null; } if (!v) { r.push(null); continue; } var t = null; try { var a = typeof v.GetLvlText === "function" ? v.GetLvlText() : null; if (Array.isArray(a)) t = a.map(function (x) { return String(x.Type) + ":" + String(x.Value); }).join("|"); } catch (e) { t = null; } var f = null, st = null; try { f = typeof v.GetFormat === "function" ? v.GetFormat() : v.Format; st = typeof v.GetStart === "function" ? v.GetStart() : v.Start; } catch (e) { return null; } r.push([f === undefined ? null : f, t, st === undefined ? null : st, v.Suff === undefined ? null : v.Suff, v.Jc === undefined ? null : v.Jc]); } return r; } '
    'function paras(id) { var a; try { a = ld.GetAllParagraphsByNumbering(new AscWord.NumPr(id, null)); } catch (e) { return null; } if (!Array.isArray(a)) return null; a = a.slice(); if (typeof AscWord.sortByDocumentPosition === "function") { try { AscWord.sortByDocumentPosition(a); } catch (e) { return null; } } return a; } '
    'var nums; try { nums = N.GetAllNumsByAbstractNum(abs); } catch (e) { return { ok: false, reason: "nums-threw" }; } if (!Array.isArray(nums)) return { ok: false, reason: "nums-unreadable" }; '
    'var ownId = String(np.NumId), self = null, others = [], ownParas = null, all = []; '
    'for (var i = 0; i < nums.length; i += 1) { var n = nums[i]; var id = String(n.GetId()); var ps = paras(n.GetId()); if (ps === null) return { ok: false, reason: "paragraphs-unreadable" }; var stl = starts(n); if (stl === null) return { ok: false, reason: "override-unreadable" }; '
    'var mem = []; for (var j = 0; j < ps.length; j += 1) { var q = ps[j], qt = textOf(q), qn; try { qn = q.GetNumPr(); } catch (e) { qn = null; } if (qt === null || !qn) return { ok: false, reason: "member-unreadable" }; mem.push({ text: qt, level: qn.Lvl }); all.push(q); } '
    'var g = { numId: id, starts: stl, members: mem, sig: sig(n), paras: ps }; if (g.sig === null) return { ok: false, reason: "member-unreadable" }; if (id === ownId) { self = g; ownParas = ps; } else others.push(g); } '
    'if (!self) return { ok: false, reason: "own-num-not-listed" }; '
    'var order = null; if (typeof AscWord.sortByDocumentPosition === "function") { try { AscWord.sortByDocumentPosition(all); var first = -1, last = -1; for (var k = 0; k < all.length; k += 1) { if (ownParas.indexOf(all[k]) >= 0) { if (first < 0) first = k; last = k; } } var contiguous = true; for (var k2 = first; k2 <= last; k2 += 1) { if (ownParas.indexOf(all[k2]) < 0) contiguous = false; } var targetFirst = true; '
    'for (var oi = 0; oi < others.length; oi += 1) { var o = others[oi]; o.before = false; o.after = false; for (var k3 = 0; k3 < o.paras.length; k3 += 1) { var at = all.indexOf(o.paras[k3]); if (at < first) { o.before = true; targetFirst = false; } if (at > last) o.after = true; } } order = { targetFirst: targetFirst, contiguous: contiguous }; } catch (e) { order = null; } } '
    'var outOthers = []; for (var oj = 0; oj < others.length; oj += 1) { var oo = others[oj]; outOthers.push({ numId: oo.numId, starts: oo.starts, members: oo.members, sig: oo.sig, before: order ? oo.before : null, after: order ? oo.after : null }); } '
    'return { ok: true, value: { present: true, level: lvl, numId: ownId, abstractId: String(num.GetAbstractNumId()), starts: self.starts, members: self.members, sig: self.sig, others: outOthers, order: order }, internal: { N: N, num: num, ownParas: ownParas, ownId: ownId } }; '
    '} '
)

READ_IMPL = (
    '"docx.numbering.listMembers.read/v1": function (args) { '
    'if (args && typeof args === "object") { var keys = Object.keys(args); for (var i = 0; i < keys.length; i += 1) { if (keys[i] !== "text" && keys[i] !== "level") return { ok: false, reason: "unsupported-property" }; } } '
    'var m = eoListMembers(args); if (!m.ok) return m; return { ok: true, value: m.value }; '
    '}, '
)

WRITE_IMPL = (
    '"docx.numbering.listDetach.write/v1": function (args) { '
    'if (!args || typeof args !== "object" || !Array.isArray(args.members) || typeof args.others !== "number") return { ok: false, reason: "bad-args" }; '
    'var keys = Object.keys(args); for (var i = 0; i < keys.length; i += 1) { if (keys[i] !== "text" && keys[i] !== "level" && keys[i] !== "members" && keys[i] !== "others") return { ok: false, reason: "unsupported-property" }; } '
    'var m = eoListMembers({ text: args.text, level: args.level }); if (!m.ok) return m; if (!m.value.present) return { ok: false, reason: "not-numbered" }; '
    'var v = m.value; if (JSON.stringify(v.members) !== JSON.stringify(args.members)) return { ok: false, reason: "members-mismatch", members: v.members }; '
    'if (v.others.length !== args.others) return { ok: false, reason: "others-mismatch", others: v.others.length }; '
    'if (v.others.length === 0) return { ok: false, reason: "nothing-to-detach" }; '
    'if (!v.order) return { ok: false, reason: "order-unreadable" }; '
    'function covers(stl, mem) { for (var a = 0; a < mem.length; a += 1) { if (stl.indexOf(mem[a].level) < 0) return false; } return true; } '
    'var iso = v.order.contiguous && (v.order.targetFirst || covers(v.starts, v.members)); '
    'for (var oi = 0; oi < v.others.length; oi += 1) { var o = v.others[oi]; if (o.before && o.after) iso = false; if (o.after && !covers(o.starts, o.members)) iso = false; } '
    'if (!iso) return { ok: false, reason: "not-isolatable", order: v.order }; '
    'var copy; try { copy = m.internal.num.Copy(); } catch (e) { return { ok: false, reason: "copy-threw", error: String(e).slice(0, 200) }; } '
    'if (!copy || typeof copy.GetId !== "function") return { ok: false, reason: "copy-threw" }; '
    'var newId = copy.GetId(), ps = m.internal.ownParas; '
    'try { for (var j = 0; j < ps.length; j += 1) { var q = ps[j]; q.SetNumPr(newId, q.GetNumPr().Lvl); } } catch (e) { return { ok: false, reason: "set-threw", error: String(e).slice(0, 200), before: { numId: m.internal.ownId } }; } '
    'var okAll = true; try { for (var k = 0; k < ps.length; k += 1) { var qn = ps[k].GetNumPr(); if (!qn || String(qn.NumId) !== String(newId)) okAll = false; } } catch (e) { okAll = false; } '
    'if (!okAll || String(copy.GetAbstractNumId()) === v.abstractId) return { ok: false, reason: "write-not-observed", before: { numId: m.internal.ownId } }; '
    'return { ok: true, value: { applied: ["list-detached"], before: { numId: m.internal.ownId }, moved: ps.length } }; '
    '}, '
)

# (old, new) -- each `old` must occur EXACTLY once in the /4 block
SUBS = [
    ('patchVersion: "euro-office-bridge-word-v1/9.3.5.0/4"', 'patchVersion: "euro-office-bridge-word-v1/9.3.5.0/5"'),
    ('(GH #292 S4, word /4)" }, ], };', '(GH #292 S4, word /4)" }, ' + READ_ENTRY + WRITE_ENTRY + '], };'),
    ('value >= 0; } var EO_BRIDGE_CAPS = { ', 'value >= 0; } ' + HELPER + 'var EO_BRIDGE_CAPS = { ' + READ_IMPL),
    ('var EO_BRIDGE_WRITES = { ', 'var EO_BRIDGE_WRITES = { ' + WRITE_IMPL),
]
MARKERS = ['"docx.numbering.listMembers.read/v1": function', '"docx.numbering.listDetach.write/v1": function', 'function eoListMembers(']


def sha(b):
    return hashlib.sha256(b).hexdigest()


def build_block():
    raw = V4_FILE.read_bytes()
    if sha(raw) != V4_FILE_SHA: sys.exit(f"REFUSE: /4 block file sha {sha(raw)} != pinned; NO OUTPUT")
    v4 = raw.decode("utf-8").rstrip()
    if not v4.endswith(","): sys.exit("REFUSE: /4 block file does not end with the separator comma; NO OUTPUT")
    v4 = v4[:-1]
    t = v4
    if any(t.count(m) for m in MARKERS): sys.exit("REFUSE: a /5 marker already in the /4 block; NO OUTPUT")
    for old, new in SUBS:
        n = t.count(old)
        if n != 1: sys.exit(f"REFUSE: substitution anchor occurs {n}x (want 1): {old[:60]!r}; NO OUTPUT")
        t = t.replace(old, new, 1)
    for m in MARKERS:
        if t.count(m) != 1: sys.exit(f"REFUSE: marker {m!r} count {t.count(m)}; NO OUTPUT")
    return v4.encode("utf-8"), t.encode("utf-8")


def main():
    v4, v5 = build_block()
    out = {"v4_block_sha": sha(v4), "v4_block_size": len(v4), "v5_block_sha": sha(v5), "v5_block_size": len(v5)}
    sdk = None
    if "--sdk" in sys.argv:
        sdk = Path(sys.argv[sys.argv.index("--sdk") + 1]).read_bytes()
        if sha(sdk) != V4_SDK_SHA or len(sdk) != V4_SDK_SIZE:
            sys.exit(f"REFUSE: sdk input sha {sha(sdk)} / {len(sdk)} is not the /4 output; NO OUTPUT")
        text = sdk.decode("utf-8")
        old, new = v4.decode("utf-8"), v5.decode("utf-8")
        if text.count(old) != 1: sys.exit("REFUSE: /4 block not exactly once in the sdk input; NO OUTPUT")
        if any(text.count(m) for m in MARKERS): sys.exit("REFUSE: a /5 marker already present in the input; NO OUTPUT")
        patched = text.replace(old, new, 1).encode("utf-8")
        out.update({"v4_sdk_sha": sha(sdk), "v4_sdk_size": len(sdk), "v5_sdk_sha": sha(patched), "v5_sdk_size": len(patched)})
        if "--out" in sys.argv: Path(sys.argv[sys.argv.index("--out") + 1]).write_bytes(patched)
    (HERE / "bridge-block.js").write_bytes(v5 + b",")   # the shipped file form: block + separator comma, no newline
    (HERE / "old.txt").write_bytes(v4)
    (HERE / "new.txt").write_bytes(v5)
    base = {"format": "sdk-all.js", "name": "euro-office-bridge-word-v1", "runtime": "9.3.5.0", "target": "sdk-all.js",
            "getters": MARKERS[:2],
            "operations": [{"id": "impl-bridge-entrypoint-word-v5-gh293-list-scope", "old": v4.decode("utf-8"), "new": v5.decode("utf-8")}]}
    if sdk is not None:
        man = dict(base, original_sha256=out["v4_sdk_sha"], patched_sha256=out["v5_sdk_sha"], original_size=out["v4_sdk_size"], patched_size=out["v5_sdk_size"],
                   _patched_sha256_NOTE="computed by build-rev5.py from the SERVED /4 word sdk-all.js (sha = rev4-candidate/manifest.json patched_sha256); CANDIDATE, NOT applied to any Document Server (DS_PATCH_APPLY=HOLD).")
        (HERE / "manifest.json").write_text(json.dumps(man, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
    print(json.dumps(out))


if __name__ == "__main__":
    main()
