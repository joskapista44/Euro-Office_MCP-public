#!/usr/bin/env python3
"""GH #225 multilevel acceptance: a DOCX with ONE numbering (numId 1) used at TWO levels, plus a
second, distinct numbering (numId 2) and a plain control paragraph, written as raw OOXML so the
READ round never has to WRITE first (the public API has no measured level>0 SetNumPr path).
Level formats differ on every field the product READ reports (numFmt, lvlText, suff, lvlJc,
start), and level 1 starts at a non-default value. Deterministic (fixed zip timestamps).
The PARAS / LEVELS lists are the single source of truth: test/live/docx-numbering-multilevel-live.live.cjs
imports them (after checking that regenerating gives the committed .docx byte-for-byte).
Usage: make-multilevel-docx.py <out.docx>"""
import sys, zipfile

# abstractNum 0 (used by numId 1): two levels; abstractNum 1 (numId 2): one level
LEVELS = {
    0: [{'ilvl': 0, 'numFmt': 'decimal', 'lvlText': '%1.', 'suff': 'tab', 'lvlJc': 'left', 'start': 2},
        {'ilvl': 1, 'numFmt': 'lowerLetter', 'lvlText': '%2)', 'suff': 'space', 'lvlJc': 'right', 'start': 4}],
    1: [{'ilvl': 0, 'numFmt': 'upperRoman', 'lvlText': '%1-', 'suff': 'nothing', 'lvlJc': 'center', 'start': 7}],
}
NUMS = {1: 0, 2: 1}  # numId -> abstractNumId
PARAS = [
    {'text': 'ZQR GH225 ML item A level0', 'numId': 1, 'ilvl': 0},
    {'text': 'ZQR GH225 ML item B level1', 'numId': 1, 'ilvl': 1},
    {'text': 'ZQR GH225 ML item C level0', 'numId': 1, 'ilvl': 0},
    {'text': 'ZQR GH225 ML item D other numbering', 'numId': 2, 'ilvl': 0},
    {'text': 'ZQR GH225 ML control', 'numId': None, 'ilvl': None},
]

def lvl(l):
    return ('<w:lvl w:ilvl="%d"><w:start w:val="%d"/><w:numFmt w:val="%s"/><w:suff w:val="%s"/>'
            '<w:lvlText w:val="%s"/><w:lvlJc w:val="%s"/><w:pPr><w:ind w:left="%d" w:hanging="360"/></w:pPr></w:lvl>'
            % (l['ilvl'], l['start'], l['numFmt'], l['suff'], l['lvlText'], l['lvlJc'], 720 * (l['ilvl'] + 1)))

def para(p):
    ppr = ('<w:pPr><w:numPr><w:ilvl w:val="%d"/><w:numId w:val="%d"/></w:numPr></w:pPr>' % (p['ilvl'], p['numId'])) if p['numId'] else ''
    return '<w:p>%s<w:r><w:t>%s</w:t></w:r></w:p>' % (ppr, p['text'])

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
numbering = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="%s">' % W
             + ''.join('<w:abstractNum w:abstractNumId="%d"><w:multiLevelType w:val="%s"/>%s</w:abstractNum>'
                       % (a, 'multilevel' if len(ls) > 1 else 'singleLevel', ''.join(lvl(l) for l in ls)) for a, ls in LEVELS.items())
             + ''.join('<w:num w:numId="%d"><w:abstractNumId w:val="%d"/></w:num>' % (n, a) for n, a in NUMS.items())
             + '</w:numbering>')
doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="%s"><w:body>%s'
       '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:body></w:document>' % (W, ''.join(para(p) for p in PARAS)))
parts = {
    '[Content_Types].xml': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>',
    '_rels/.rels': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
    'word/_rels/document.xml.rels': '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/></Relationships>',
    'word/document.xml': doc,
    'word/numbering.xml': numbering,
}
if __name__ == '__main__':
    with zipfile.ZipFile(sys.argv[1], 'w', zipfile.ZIP_DEFLATED) as z:
        for name, data in parts.items():
            info = zipfile.ZipInfo(name, date_time=(2026, 10, 19, 0, 0, 0)); info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, data)
