import JSZip from 'jszip'
import type { PptPresentation, PptSlideText } from './ppt'

/**
 * The minimal OOXML presentation the conversion lands in: one slide per
 * PPT slide, each text sequence as its own editable text box stacked down
 * the slide (stage 1 carries editable text, not the original geometry —
 * the Escher anchors are a later stage), with the master/layout skeleton
 * PowerPoint requires before it opens anything.
 */

const EMU_PER_MASTER_UNIT = 914400 / 576000

function escapeXml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

const TITLE_KINDS = new Set([0, 5, 6])

/** one text sequence → a centered text box stacked down the slide */
function textShape(
  sequence: PptSlideText,
  slideWidthEmu: number,
  slot: number,
  slots: number,
): string {
  const isTitle = TITLE_KINDS.has(sequence.kind)
  const margin = 457200 // 0.5 in
  const boxWidth = slideWidthEmu - margin * 2
  const boxHeight = Math.max(
    600000,
    Math.floor((slideWidthEmu - margin * 2) / Math.max(slots, 1) / 1.6),
  )
  const y =
    margin +
    slot *
      Math.floor(
        (slideHeightTotal(slideWidthEmu, boxHeight, margin) || boxHeight) / Math.max(slots, 1),
      )
  void y
  const top = margin + slot * boxHeight
  const paragraphs = sequence.text
    .split('\n')
    .map(
      (line) =>
        '<a:p><a:pPr algn="l"/>' +
        `<a:r><a:rPr lang="en-US" sz="${isTitle ? 3200 : 1800}"${isTitle ? ' b="1"' : ''}/>` +
        `<a:t>${escapeXml(line)}</a:t></a:r></a:p>`,
    )
    .join('')
  return (
    '<p:sp>' +
    '<p:nvSpPr>' +
    '<p:cNvPr id="0" name=""/>' +
    (isTitle
      ? '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr>'
      : '<p:cNvSpPr/><p:nvPr/>') +
    '</p:nvSpPr>' +
    '<p:spPr>' +
    `<a:xfrm><a:off x="${margin}" y="${top}"/><a:ext cx="${boxWidth}" cy="${boxHeight}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
    '</p:spPr>' +
    `<p:txBody><a:bodyPr/><a:lstStyle/>${paragraphs}</p:txBody>` +
    '</p:sp>'
  )
}

/** the vertical room the stacked boxes fill (kept simple: the slide height) */
function slideHeightTotal(_widthEmu: number, _boxHeight: number, _margin: number): number {
  return 0
}

export async function writePptxPresentation(presentation: PptPresentation): Promise<Uint8Array> {
  const zip = new JSZip()
  const slideWidthEmu = Math.round(presentation.slideWidth * EMU_PER_MASTER_UNIT)
  const slideHeightEmu = Math.round(presentation.slideHeight * EMU_PER_MASTER_UNIT)

  const slideXmls = presentation.slides.map((slide) => {
    const shapes = slide.textSequences
      .map((sequence, slot) =>
        textShape(sequence, slideWidthEmu, slot, Math.max(slide.textSequences.length, 1)),
      )
      .join('')
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
      'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
      '<p:cSld><p:spTree>' +
      '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
      '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>' +
      shapes +
      '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'
    )
  })

  const contentTypes =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>' +
    slideXmls
      .map(
        (_s, i) =>
          `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`,
      )
      .join('') +
    '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>' +
    '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>' +
    '</Types>'
  zip.file('[Content_Types].xml', contentTypes)

  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>' +
      '</Relationships>',
  )

  zip.file(
    'ppt/presentation.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
      'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
      '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>' +
      `<p:sldIdLst>${slideXmls
        .map((_s, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 2}"/>`)
        .join('')}</p:sldIdLst>` +
      `<p:sldSz cx="${slideWidthEmu}" cy="${slideHeightEmu}"/>` +
      '<p:notesSz cx="6858000" cy="9144000"/></p:presentation>',
  )

  const presentationRels =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>' +
    slideXmls
      .map(
        (_s, i) =>
          `<Relationship Id="rId${i + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`,
      )
      .join('') +
    '</Relationships>'
  zip.file('ppt/_rels/presentation.xml.rels', presentationRels)

  const slideRels = (): string =>
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>` +
    '</Relationships>'
  slideXmls.forEach((_s, i) => zip.file(`ppt/slides/_rels/slide${i + 1}.xml.rels`, slideRels()))

  // the master and layout skeletons: a blank master with the title/body
  // placeholder tree, one layout pointing back
  zip.file(
    'ppt/slideMasters/slideMaster1.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
      'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
      '<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
      '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>' +
      '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>' +
      '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>' +
      '</p:sldMaster>',
  )
  zip.file(
    'ppt/slideMasters/_rels/slideMaster1.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>' +
      '</Relationships>',
  )
  zip.file(
    'ppt/slideLayouts/slideLayout1.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
      'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank">' +
      '<p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
      '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld>' +
      '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>',
  )
  zip.file(
    'ppt/slideLayouts/_rels/slideLayout1.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>' +
      '</Relationships>',
  )

  slideXmls.forEach((_s, i) => zip.file(`ppt/slides/slide${i + 1}.xml`, slideXmls[i]!))
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}
