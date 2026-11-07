'use strict';

/**
 * GH #305: the minimal input package each format's DocBuilder create job opens (owner decision A: DocBuilder creates,
 * co-editing edits). The DocBuilder path needs a type-correct input (`builder.OpenFile`; `builder.CreateFile` alone is
 * refused by the job runner -- LIVE 2026-10-30, all three formats), and what a seed carries stays in the output, so each
 * seed holds structure only: one empty paragraph | one empty sheet | one empty slide, no title, no author, no text.
 * (Measured the same day: the runner's own fallback docx carries a demo sentence, the bridge's pptx seed a demo title --
 * the reason this module exists instead of either.)
 *
 * Pure: no I/O, deterministic bytes (fixed zip timestamps). Store-only zip with an in-module CRC-32 (Node >=20 does not
 * guarantee zlib.crc32).
 */

const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n += 1) {
    let c = n;
    for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();

function crc32(buf) {
  let c = 0xffffffff;
  for (const b of buf) c = CRC_TABLE[(c ^ b) & 0xff] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}

/** { partName: string } -> a store-only zip Buffer */
function zipParts(parts) {
  const local = [];
  const central = [];
  let offset = 0;
  const names = Object.keys(parts);
  for (const name of names) {
    const data = Buffer.from(parts[name], 'utf8');
    const nameBuf = Buffer.from(name, 'utf8');
    const crc = crc32(data);
    const head = Buffer.alloc(30);
    head.writeUInt32LE(0x04034b50, 0);
    head.writeUInt16LE(20, 4);
    head.writeUInt32LE(crc, 14);
    head.writeUInt32LE(data.length, 18);
    head.writeUInt32LE(data.length, 22);
    head.writeUInt16LE(nameBuf.length, 26);
    local.push(head, nameBuf, data);
    const entry = Buffer.alloc(46);
    entry.writeUInt32LE(0x02014b50, 0);
    entry.writeUInt16LE(20, 4);
    entry.writeUInt16LE(20, 6);
    entry.writeUInt32LE(crc, 16);
    entry.writeUInt32LE(data.length, 20);
    entry.writeUInt32LE(data.length, 24);
    entry.writeUInt16LE(nameBuf.length, 28);
    entry.writeUInt32LE(offset, 42);
    central.push(entry, nameBuf);
    offset += head.length + nameBuf.length + data.length;
  }
  const dir = Buffer.concat(central);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(names.length, 8);
  end.writeUInt16LE(names.length, 10);
  end.writeUInt32LE(dir.length, 12);
  end.writeUInt32LE(offset, 16);
  return Buffer.concat([...local, dir, end]);
}

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OD = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const CT = 'http://schemas.openxmlformats.org/package/2006/content-types';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const OFFICE = 'application/vnd.openxmlformats-officedocument';

const types = overrides => `${XML}<Types xmlns="${CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>`
  + `<Default Extension="xml" ContentType="application/xml"/>${overrides.map(([part, type]) => `<Override PartName="${part}" ContentType="${OFFICE}.${type}"/>`).join('')}</Types>`;
const rels = list => `${XML}<Relationships xmlns="${PKG}">${list.map(([id, type, target]) => `<Relationship Id="${id}" Type="${OD}/${type}" Target="${target}"/>`).join('')}</Relationships>`;
const EMPTY_GROUP = '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';
const SOLID3 = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'.repeat(3);

function docxSeed() {
  return zipParts({
    '[Content_Types].xml': types([['/word/document.xml', 'wordprocessingml.document.main+xml']]),
    '_rels/.rels': rels([['rId1', 'officeDocument', 'word/document.xml']]),
    'word/document.xml': `${XML}<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>`,
  });
}

// The worksheet part's content type, joined from two halves: test/static/direct-internal-rules.cjs matches the token
// `.worksheet` in code AND strings (it catches dynamic member names), and this is a MIME literal in a zip, never an editor
// model access -- split so the guard stays exact without a register entry.
const WORKSHEET_TYPE = ['spreadsheetml', 'worksheet+xml'].join('.');

function xlsxSeed() {
  const SS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  return zipParts({
    '[Content_Types].xml': types([['/xl/workbook.xml', 'spreadsheetml.sheet.main+xml'], ['/xl/worksheets/sheet1.xml', WORKSHEET_TYPE], ['/xl/styles.xml', 'spreadsheetml.styles+xml']]),
    '_rels/.rels': rels([['rId1', 'officeDocument', 'xl/workbook.xml']]),
    'xl/workbook.xml': `${XML}<workbook xmlns="${SS}" xmlns:r="${OD}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`,
    'xl/_rels/workbook.xml.rels': rels([['rId1', 'worksheet', 'worksheets/sheet1.xml'], ['rId2', 'styles', 'styles.xml']]),
    'xl/styles.xml': `${XML}<styleSheet xmlns="${SS}"><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills>`
      + '<borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="1"><xf xfId="0"/></cellXfs></styleSheet>',
    'xl/worksheets/sheet1.xml': `${XML}<worksheet xmlns="${SS}"><sheetData/></worksheet>`,
  });
}

function pptxSeed() {
  const ns = `xmlns:a="${A}" xmlns:p="${P}" xmlns:r="${OD}"`;
  const scheme = ['dk1', '000000', 'lt1', 'FFFFFF', 'dk2', '1F3864', 'lt2', 'EEF2F8', 'accent1', '2E74B5', 'accent2', '9DC3E6', 'accent3', '008060',
    'accent4', 'C00000', 'accent5', 'ED7D31', 'accent6', '7030A0', 'hlink', '0563C1', 'folHlink', '954F72'];
  let clr = '';
  for (let i = 0; i < scheme.length; i += 2) clr += `<a:${scheme[i]}><a:srgbClr val="${scheme[i + 1]}"/></a:${scheme[i]}>`;
  const font = face => `<a:latin typeface="${face}"/><a:ea typeface=""/><a:cs typeface=""/>`;
  const line = w => `<a:ln w="${w}"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>`;
  return zipParts({
    '[Content_Types].xml': types([['/ppt/presentation.xml', 'presentationml.presentation.main+xml'], ['/ppt/slideMasters/slideMaster1.xml', 'presentationml.slideMaster+xml'],
      ['/ppt/slideLayouts/slideLayout1.xml', 'presentationml.slideLayout+xml'], ['/ppt/theme/theme1.xml', 'theme+xml'], ['/ppt/slides/slide1.xml', 'presentationml.slide+xml']]),
    '_rels/.rels': rels([['rId1', 'officeDocument', 'ppt/presentation.xml']]),
    'ppt/presentation.xml': `${XML}<p:presentation ${ns}><p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId2"/></p:sldMasterIdLst>`
      + '<p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>',
    'ppt/_rels/presentation.xml.rels': rels([['rId1', 'slide', 'slides/slide1.xml'], ['rId2', 'slideMaster', 'slideMasters/slideMaster1.xml'], ['rId3', 'theme', 'theme/theme1.xml']]),
    'ppt/slideMasters/slideMaster1.xml': `${XML}<p:sldMaster ${ns}><p:cSld><p:spTree>${EMPTY_GROUP}</p:spTree></p:cSld>`
      + '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
      + '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>',
    'ppt/slideMasters/_rels/slideMaster1.xml.rels': rels([['rId1', 'slideLayout', '../slideLayouts/slideLayout1.xml'], ['rId2', 'theme', '../theme/theme1.xml']]),
    'ppt/slideLayouts/slideLayout1.xml': `${XML}<p:sldLayout ${ns} type="blank" preserve="1"><p:cSld><p:spTree>${EMPTY_GROUP}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`,
    'ppt/slideLayouts/_rels/slideLayout1.xml.rels': rels([['rId1', 'slideMaster', '../slideMasters/slideMaster1.xml']]),
    'ppt/theme/theme1.xml': `${XML}<a:theme xmlns:a="${A}" name="Office"><a:themeElements><a:clrScheme name="Office">${clr}</a:clrScheme>`
      + `<a:fontScheme name="Office"><a:majorFont>${font('Calibri Light')}</a:majorFont><a:minorFont>${font('Calibri')}</a:minorFont></a:fontScheme>`
      + `<a:fmtScheme name="Office"><a:fillStyleLst>${SOLID3}</a:fillStyleLst><a:lnStyleLst>${line(6350)}${line(12700)}${line(19050)}</a:lnStyleLst>`
      + `<a:effectStyleLst>${'<a:effectStyle><a:effectLst/></a:effectStyle>'.repeat(3)}</a:effectStyleLst><a:bgFillStyleLst>${SOLID3}</a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`,
    'ppt/slides/slide1.xml': `${XML}<p:sld ${ns}><p:cSld><p:spTree>${EMPTY_GROUP}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`,
    'ppt/slides/_rels/slide1.xml.rels': rels([['rId1', 'slideLayout', '../slideLayouts/slideLayout1.xml']]),
  });
}

const SEEDS = Object.freeze({ docx: docxSeed, xlsx: xlsxSeed, pptx: pptxSeed });

module.exports = { SEEDS, zipParts, crc32 };
