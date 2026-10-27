#!/usr/bin/env python3
"""GH #223 multi-section acceptance: a DOCX with THREE sections whose page setup is pairwise
DISTINCT on every #223 field (size, the four margins, column count/space, start page number),
written as raw OOXML so the READ round never has to WRITE first. Section 1 is landscape. No value
equals a common default (A4 11906x16838, 1440 margins, 1 column, start 1), so an unread getter
falling back to a default cannot pass. Deterministic (fixed zip timestamps).
The SECTIONS list is the single source of truth: test/live/docx-page-setup-multisection-live.live.cjs
imports the SAME list (and first checks that regenerating gives the committed .docx byte-for-byte),
never a second hand-copied table.
Usage: make-multisection-docx.py <out.docx>"""
import sys, zipfile

SECTIONS = [
    {'w': 11000, 'h': 14000, 'orient': None, 'top': 1102, 'right': 1203, 'bottom': 1304, 'left': 1001, 'cols': 2, 'space': 720, 'start': 7},
    {'w': 15000, 'h': 10000, 'orient': 'landscape', 'top': 901, 'right': 802, 'bottom': 703, 'left': 604, 'cols': 3, 'space': 360, 'start': 21},
    {'w': 12240, 'h': 15840, 'orient': None, 'top': 2002, 'right': 1903, 'bottom': 1804, 'left': 1705, 'cols': 4, 'space': 425, 'start': 3},
]

def sect_pr(s):
    orient = ' w:orient="%s"' % s['orient'] if s['orient'] else ''
    return ('<w:sectPr><w:pgSz w:w="%d" w:h="%d"%s/>'
            '<w:pgMar w:top="%d" w:right="%d" w:bottom="%d" w:left="%d" w:header="708" w:footer="708" w:gutter="0"/>'
            '<w:pgNumType w:start="%d"/><w:cols w:num="%d" w:space="%d"/></w:sectPr>'
            % (s['w'], s['h'], orient, s['top'], s['right'], s['bottom'], s['left'], s['start'], s['cols'], s['space']))

def para(text, sect=''):
    ppr = '<w:pPr>%s</w:pPr>' % sect if sect else ''
    return '<w:p>%s<w:r><w:t>%s</w:t></w:r></w:p>' % (ppr, text)

body = ''
for i, s in enumerate(SECTIONS):
    body += para('ZQR GH223 section %d body' % i)
    # every section but the last ends on a paragraph-level sectPr; the last is the body-level one
    if i < len(SECTIONS) - 1:
        body += para('ZQR GH223 section %d end' % i, sect_pr(s))
body += sect_pr(SECTIONS[-1])
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="%s"><w:body>%s'
       '</w:body></w:document>' % (W, body))
parts = {
    '[Content_Types].xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
    '_rels/.rels': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
    'word/document.xml': doc,
}
if __name__ == '__main__':
    with zipfile.ZipFile(sys.argv[1], 'w', zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            info = zipfile.ZipInfo(name, date_time=(2026, 10, 19, 0, 0, 0)); info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, data)
