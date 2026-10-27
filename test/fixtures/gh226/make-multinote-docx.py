#!/usr/bin/env python3
"""GH #226 (read_notes full text, the NEXT_CHAIN boundary measurement): a DOCX with THREE footnotes and THREE
endnotes of DIFFERENT paragraph counts (footnotes 1/2/3, endnotes 3/1/2), written as raw OOXML (the SAME minimal, dependency-free technique
test/fixtures/gh225/make-multilevel-docx.py uses -- no WRITE path exists for footnotes/endnotes in
this product, exactly the gap that made raw-OOXML the right tool there). The structure follows the
standard ECMA-376 footnote/endnote shape (footnote ids 0/1 reserved for separator/
continuationSeparator, same as a real Word-produced document); it has NOT been confirmed against
the real DS 9.3.5 editor's own open/parse step -- that is this harness's own first, and genuinely
open, question (docs/GH226-NOTES-READ-DECISION.md section 2).

NOTES is the single source of truth: test/live/docx-notes-chain-boundary.live.cjs imports it
(after checking that regenerating gives the committed .docx byte-for-byte).
Usage: make-multinote-docx.py <out.docx>"""
import sys, zipfile

# per kind, per note (in reference order): the paragraph texts. Every text is UNIQUE, so a boundary overrun is visible by text.
FOOT_NOTES = [
    ['ZQN foot A p1'],
    ['ZQN foot B p1', 'ZQN foot B p2'],
    ['ZQN foot C p1', 'ZQN foot C p2', 'ZQN foot C p3'],
]
END_NOTES = [
    ['ZQN end A p1', 'ZQN end A p2', 'ZQN end A p3'],
    ['ZQN end B p1'],
    ['ZQN end C p1', 'ZQN end C p2'],
]
BODY_TEXTS = ['ZQN body one', 'ZQN body two', 'ZQN body three']

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
CT_NS = 'http://schemas.openxmlformats.org/package/2006/content-types'
PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships'


def note_part(tag, notes, first_id):
    """One <w:footnotes>/<w:endnotes> part: the two reserved notes (separator/continuationSeparator,
    id 0/1 -- the standard Word-produced shape) plus ONE custom note per entry of `notes`, ids
    first_id.., each with its paragraphs as separate <w:p> elements (the first carries the
    note-reference mark, every paragraph a plain, unstyled run)."""
    mark = 'footnoteRef' if tag == 'footnote' else 'endnoteRef'
    customs = ''
    for k, paras in enumerate(notes):
        first = ('<w:p><w:r><w:%s/></w:r><w:r><w:t xml:space="preserve"> %s</w:t></w:r></w:p>'
                 % (mark, paras[0]))
        rest = ''.join('<w:p><w:r><w:t>%s</w:t></w:r></w:p>' % p for p in paras[1:])
        customs += '<w:%s w:id="%d">%s%s</w:%s>' % (tag, first_id + k, first, rest, tag)
    reserved = (
        '<w:%s w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r></w:p></w:%s>'
        '<w:%s w:type="continuationSeparator" w:id="1"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:%s>'
        % (tag, tag, tag, tag)
    )
    plural = 'footnotes' if tag == 'footnote' else 'endnotes'
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:%s xmlns:w="%s">%s%s</w:%s>'
            % (plural, W, reserved, customs, plural))


def main(out_path):
    first_id = 2
    footnotes_xml = note_part('footnote', FOOT_NOTES, first_id)
    endnotes_xml = note_part('endnote', END_NOTES, first_id)

    # one body paragraph per note index, carrying that index's footnote AND endnote reference
    body = ''.join(
        '<w:p><w:r><w:t xml:space="preserve">%s</w:t></w:r>'
        '<w:r><w:footnoteReference w:id="%d"/></w:r>'
        '<w:r><w:endnoteReference w:id="%d"/></w:r></w:p>'
        % (BODY_TEXTS[k], first_id + k, first_id + k)
        for k in range(len(FOOT_NOTES))
    )
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="%s"><w:body>%s'
           '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>' % (W, body))

    content_types = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="%s">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>'
        '<Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>'
        '</Types>' % CT_NS
    )
    pkg_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="%s">'
        '<Relationship Id="rId1" Type="%s/officeDocument" Target="word/document.xml"/>'
        '</Relationships>' % (PKG_REL_NS, R_NS)
    )
    doc_rels = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="%s">'
        '<Relationship Id="rId1" Type="%s/footnotes" Target="footnotes.xml"/>'
        '<Relationship Id="rId2" Type="%s/endnotes" Target="endnotes.xml"/>'
        '</Relationships>' % (PKG_REL_NS, R_NS, R_NS)
    )
    parts = {
        '[Content_Types].xml': content_types,
        '_rels/.rels': pkg_rels,
        'word/_rels/document.xml.rels': doc_rels,
        'word/document.xml': doc,
        'word/footnotes.xml': footnotes_xml,
        'word/endnotes.xml': endnotes_xml,
    }
    with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            info = zipfile.ZipInfo(name, date_time=(2026, 10, 19, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, data)


if __name__ == '__main__':
    main(sys.argv[1])
