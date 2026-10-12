#!/usr/bin/env python3
"""M5 known-limits row 14: a DOCX whose tables have UNEVEN LOGICAL WIDTH -- a shape the public
Euro-Office API cannot create (W4c G1-G5 always produce a colspan grid), so it is written as raw
OOXML. Deterministic (fixed zip timestamps). Two tables:
  R (row 14 case ii): EQUAL physical cell count (3 per row), UNEQUAL logical width --
     row 0: three span-1 cells + gridAfter 1 (width 3) | rows 1-2: one gridSpan-2 cell (width 4)
  S (row 14 case i):  UNEQUAL physical count -- row 0 three cells + gridAfter 1 | row 1 four cells
Usage: make-ragged-grid-docx.py <out.docx>"""
import sys, zipfile

def cell(text, span=1):
    tc = '<w:tcPr><w:gridSpan w:val="%d"/></w:tcPr>' % span if span > 1 else ''
    return '<w:tc>%s<w:p><w:r><w:t>%s</w:t></w:r></w:p></w:tc>' % (tc, text)

def row(cells, after=0):
    tr = '<w:trPr><w:gridAfter w:val="%d"/></w:trPr>' % after if after else ''
    return '<w:tr>%s%s</w:tr>' % (tr, ''.join(cell(t, s) for t, s in cells))

def table(rows):
    grid = ''.join('<w:gridCol w:w="2000"/>' for _ in range(4))
    return ('<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/></w:tblPr><w:tblGrid>%s</w:tblGrid>%s</w:tbl>'
            % (grid, ''.join(row(c, a) for c, a in rows)))

def para(text):
    return '<w:p><w:r><w:t>%s</w:t></w:r></w:p>' % text

R = table([([('R a0', 1), ('R a1', 1), ('R a2', 1)], 1),
           ([('R b0', 1), ('R b1', 2), ('R b2', 1)], 0),
           ([('R c0', 1), ('R c1', 2), ('R anchor', 1)], 0)])
S = table([([('S a0', 1), ('S a1', 1), ('S a2', 1)], 1),
           ([('S b0', 1), ('S b1', 1), ('S b2', 1), ('S anchor', 1)], 0)])
body = para('ZQR ragged start') + R + para('ZQR ragged between') + S + para('ZQR ragged end')
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="%s"><w:body>%s'
       '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>' % (W, body))
parts = {
    '[Content_Types].xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
    '_rels/.rels': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
    'word/document.xml': doc,
}
with zipfile.ZipFile(sys.argv[1], 'w', zipfile.ZIP_DEFLATED) as z:
    for name, data in parts.items():
        info = zipfile.ZipInfo(name, date_time=(2026, 9, 26, 0, 0, 0)); info.compress_type = zipfile.ZIP_DEFLATED
        z.writestr(info, data)
