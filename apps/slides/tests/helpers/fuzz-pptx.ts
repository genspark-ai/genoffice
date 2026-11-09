/** A seeded generator of varied-but-valid PPTX packages for the hostile-input
 *  fuzz suite — the slides sibling of the docs suite's helpers/fuzz-docx. The
 *  composition vocabulary: one to three text-only slides with CJK, latin and
 *  emoji runs, rich run properties, empty paragraphs, two slide sizes. Parts a
 *  real producer always writes stay present; the corpus keeps everything the
 *  open pipeline reads well-formed so corruption is the only variable.
 *  Same seed -> same bytes, so a failing seed reproduces exactly. */
import JSZip from 'jszip'

/** mulberry32: small, fast, fully deterministic across runs and platforms. */
export function rng(seed: number): () => number {
  let a = seed >>> 0
  return () => {
    a |= 0
    a = (a + 0x6d2b79f5) | 0
    let t = Math.imul(a ^ (a >>> 15), 1 | a)
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296
  }
}

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
const CJK = ['季度回顾', '营收明细', '附录甲', '风险登记', '第三季度业绩']
const LATIN = ['Quarterly review', 'Appendix A', 'risk register', 'milestone 12', 'draft v2']
const EMOJI = ['🎯', '📈', '✅']

function escapeXml(text: string): string {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

function runXml(text: string, r: () => number): string {
  const props: string[] = []
  if (r() < 0.4) props.push('b="1"')
  if (r() < 0.3) props.push('i="1"')
  if (r() < 0.3) props.push(`sz="${[1400, 1800, 2400, 3200][Math.floor(r() * 4)]}"`)
  if (r() < 0.2) props.push('u="sng"')
  const rPr = props.length > 0 ? `<a:rPr ${props.join(' ')}/>` : ''
  return `<a:r>${rPr}<a:t>${escapeXml(text)}</a:t></a:r>`
}

function textBody(r: () => number, paragraphCount: number): string {
  const paragraphs: string[] = []
  for (let p = 0; p < paragraphCount; p++) {
    if (p > 0 && r() < 0.15) {
      paragraphs.push('<a:p><a:endParaRPr/></a:p>')
      continue
    }
    const runs: string[] = []
    const runCount = 1 + Math.floor(r() * 3)
    for (let i = 0; i < runCount; i++) {
      const roll = r()
      const text =
        roll < 0.45
          ? CJK[Math.floor(r() * CJK.length)]!
          : roll < 0.9
            ? LATIN[Math.floor(r() * LATIN.length)]!
            : EMOJI[Math.floor(r() * EMOJI.length)]!
      runs.push(runXml(text, r))
    }
    paragraphs.push(`<a:p>${runs.join('')}</a:p>`)
  }
  return `<p:txBody><a:bodyPr/><a:lstStyle/>${paragraphs.join('')}</p:txBody>`
}

function slideXml(r: () => number): string {
  const shapes: string[] = []
  const shapeCount = 1 + Math.floor(r() * 4)
  for (let s = 0; s < shapeCount; s++) {
    const x = Math.floor(r() * 8000000)
    const y = Math.floor(r() * 4000000)
    shapes.push(
      '<p:sp><p:nvSpPr><p:cNvPr id="' +
        (s + 2) +
        '" name="TextBox ' +
        (s + 1) +
        '"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>' +
        '<p:spPr><a:xfrm><a:off x="' +
        x +
        '" y="' +
        y +
        '"/><a:ext cx="' +
        (3000000 + Math.floor(r() * 2000000)) +
        '" cy="' +
        (600000 + Math.floor(r() * 900000)) +
        '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>' +
        textBody(r, 1 + Math.floor(r() * 3)) +
        '</p:sp>',
    )
  }
  return (
    XML_DECL +
    '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
    '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
    '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
    shapes.join('') +
    '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'
  )
}

/** Build one varied-but-valid deck. The master/layout/theme are fixed minimal
 *  parts; the slides carry the variety. */
export async function buildFuzzPptx(seed: number): Promise<Uint8Array> {
  const r = rng(seed * 6151 + 29)
  const zip = new JSZip()
  const slideCount = 1 + Math.floor(r() * 3)
  // 16:9 and 4:3 surfaces, in EMU
  const slideSize = r() < 0.6 ? [12192000, 6858000] : [9144000, 6858000]

  zip.file(
    '[Content_Types].xml',
    `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>' +
      '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>' +
      '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>' +
      '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' +
      Array.from(
        { length: slideCount },
        (_v, i) =>
          `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`,
      ).join('') +
      '</Types>',
  )
  zip.file(
    '_rels/.rels',
    `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>' +
      '</Relationships>',
  )
  const slideRels = Array.from(
    { length: slideCount },
    (_v, i) =>
      `<Relationship Id="rId${i + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`,
  ).join('')
  zip.file(
    'ppt/presentation.xml',
    `${XML_DECL}<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ` +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>' +
      '<p:sldIdLst>' +
      Array.from(
        { length: slideCount },
        (_v, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 2}"/>`,
      ).join('') +
      '</p:sldIdLst>' +
      `<p:sldSz cx="${slideSize[0]}" cy="${slideSize[1]}"/>` +
      '<p:notesSz cx="6858000" cy="9144000"/></p:presentation>',
  )
  zip.file(
    'ppt/_rels/presentation.xml.rels',
    `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>' +
      slideRels +
      '<Relationship Id="rId' +
      (slideCount + 2) +
      '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>' +
      '</Relationships>',
  )
  zip.file(
    'ppt/slideMasters/slideMaster1.xml',
    `${XML_DECL}<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ` +
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
      '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
      '</p:spTree></p:cSld>' +
      '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>' +
      '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>',
  )
  zip.file(
    'ppt/slideMasters/_rels/slideMaster1.xml.rels',
    `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>' +
      '</Relationships>',
  )
  zip.file(
    'ppt/slideLayouts/slideLayout1.xml',
    `${XML_DECL}<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ` +
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
      'type="blank" preserve="1">' +
      '<p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
      '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
      '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>',
  )
  zip.file(
    'ppt/slideLayouts/_rels/slideLayout1.xml.rels',
    `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>' +
      '</Relationships>',
  )
  zip.file(
    'ppt/theme/theme1.xml',
    `${XML_DECL}<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Fuzz">` +
      '<a:themeElements><a:clrScheme name="Fuzz"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>' +
      '<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>' +
      '<a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>' +
      '<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>' +
      '<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>' +
      '<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>' +
      '<a:fontScheme name="Fuzz"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>' +
      '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>' +
      '<a:fmtScheme name="Fuzz"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
      '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>' +
      '<a:lnStyleLst><a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>' +
      '<a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>' +
      '<a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>' +
      '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle>' +
      '<a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>' +
      '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
      '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>' +
      '</a:fmtScheme></a:themeElements></a:theme>',
  )
  for (let s = 0; s < slideCount; s++) {
    zip.file(`ppt/slides/slide${s + 1}.xml`, slideXml(r))
  }
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}
