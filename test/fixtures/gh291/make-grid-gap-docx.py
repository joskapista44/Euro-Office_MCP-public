#!/usr/bin/env python3
"""GH #291 H4 (table shape matrix): tables whose short rows leave a GRID GAP -- a shape the public Euro-Office API cannot
create, so it is written as raw OOXML (as test/fixtures/m5/make-ragged-grid-docx.py). Deterministic (fixed zip timestamps).
Four 4-column tables:
  G1  row 0: three cells + gridAfter 1                 | row 1: four cells
  G2  row 0: two cells + gridAfter 2                   | row 1: four cells
  G3  row 0: gridBefore 1 + three cells                | row 1: four cells
  G4  row 0: three cells + gridAfter 1                 | row 1: b0, a gridSpan-2 b1, anchor
Usage: make-grid-gap-docx.py <out.docx>"""
import sys, zipfile

def cell(text, span=1):
    tc = '<w:tcPr><w:gridSpan w:val="%d"/></w:tcPr>' % span if span > 1 else ''
    return '<w:tc>%s<w:p><w:r><w:t>%s</w:t></w:r></w:p></w:tc>' % (tc, text)

def row(cells, before=0, after=0):
    pr = ('<w:gridBefore w:val="%d"/>' % before if before else '') + ('<w:gridAfter w:val="%d"/>' % after if after else '')
    return '<w:tr>%s%s</w:tr>' % ('<w:trPr>%s</w:trPr>' % pr if pr else '', ''.join(cell(t, s) for t, s in cells))

def table(rows):
    grid = ''.join('<w:gridCol w:w="2000"/>' for _ in range(4))
    return '<w:tbl><w:tblPr><w:tblW w:w="8000" w:type="dxa"/></w:tblPr><w:tblGrid>%s</w:tblGrid>%s</w:tbl>' % (grid, ''.join(rows))

def para(text):
    return '<w:p><w:r><w:t>%s</w:t></w:r></w:p>' % text

full = lambda tag: row([('%s b0' % tag, 1), ('%s b1' % tag, 1), ('%s b2' % tag, 1), ('%s anchor' % tag, 1)])
G1 = table([row([('G1 a0', 1), ('G1 a1', 1), ('G1 a2', 1)], after=1), full('G1')])
G2 = table([row([('G2 a0', 1), ('G2 a1', 1)], after=2), full('G2')])
G3 = table([row([('G3 a1', 1), ('G3 a2', 1), ('G3 a3', 1)], before=1), full('G3')])
G4 = table([row([('G4 a0', 1), ('G4 a1', 1), ('G4 a2', 1)], after=1), row([('G4 b0', 1), ('G4 b1', 2), ('G4 anchor', 1)])])
body = para('GH291 H4 start') + ''.join(t + para('GH291 H4 after %s' % n) for n, t in (('G1', G1), ('G2', G2), ('G3', G3), ('G4', G4)))
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
        info = zipfile.ZipInfo(name, date_time=(2026, 10, 22, 0, 0, 0)); info.compress_type = zipfile.ZIP_DEFLATED
        z.writestr(info, data)
