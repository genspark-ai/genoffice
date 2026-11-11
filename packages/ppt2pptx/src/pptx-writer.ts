import JSZip from 'jszip'
import type { PptPresentation, PptSlide, PptSlideText } from './ppt'
import { FSP, OPT, SHAPE_PRESETS, type EscherRect, type PptImage, type PptShape } from './escher'
import { applyCharacterRuns, type TextCharFormat, type TextStyle } from './text-style'

/**
 * The OOXML presentation the conversion lands in. Stage 1 stacked every text
 * sequence down the slide; stage 3 uses the Escher anchors when a drawing
 * carries them, applies StyleTextProp runs to the text, emits embedded PNG /
 * JPEG blips as <p:pic> shapes and reproduces the main master's shapes. A deck
 * without anchors still takes the stage-1 fallback, so no text is ever lost.
 */

/** PowerPoint measures the slide surface in 1/576-inch master units. */
const EMU_PER_MASTER_UNIT = 914400 / 576

const EMPTY_STYLE: TextStyle = { paragraphs: [], characters: [] }

function xmlLegal(ch: string): boolean {
  const c = ch.codePointAt(0)!
  return (c >= 0x20 && c !== 0xfffe && c !== 0xffff) || c === 0x9 || c === 0xa || c === 0xd
}

function escapeXml(text: string): string {
  return [...text]
    .filter(xmlLegal)
    .join('')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
}

const TITLE_KINDS = new Set([0, 6]) // 5 is the subtitle

/** OEPlaceholderAtom.placeholderId -> DrawingML placeholder type (MS-PPT spec) */
const PLACEHOLDER_TYPES: Record<number, string> = {
  1: 'title',
  3: 'ctrTitle',
  13: 'title',
  15: 'ctrTitle',
  2: 'body',
  14: 'body',
  18: 'body',
  4: 'subTitle',
  16: 'subTitle',
  7: 'dt',
  8: 'sldNum',
  9: 'ftr',
  10: 'hdr',
  5: 'sldImg',
  11: 'sldImg',
  12: 'body',
  19: 'obj',
  20: 'obj',
  21: 'obj',
  22: 'obj',
  23: 'obj',
  24: 'obj',
}

interface TextDefaults {
  sizePt: number | null
  titleSizePt: number | null
  typeface: string | null
}

function textDefaultsOf(presentation: PptPresentation): TextDefaults {
  const levels = presentation.masterStyles.levels
  const withSize = levels.filter((level) => level.sizePt !== null)
  const title = withSize.find((level) => level.textType === 0) ?? withSize[0]
  const body = withSize.find((level) => level.textType === 1) ?? withSize[0]
  const fontRef = title?.fontRef ?? body?.fontRef ?? null
  const typefaces = presentation.masterStyles.typefaces
  return {
    sizePt: body?.sizePt ?? null,
    titleSizePt: title?.sizePt ?? null,
    typeface: (fontRef !== null ? typefaces[fontRef] : undefined) ?? typefaces[0] ?? null,
  }
}

interface Box {
  x: number
  y: number
  cx: number
  cy: number
}

function boxOf(rect: EscherRect): Box {
  const x = Math.round(rect.left * EMU_PER_MASTER_UNIT)
  const y = Math.round(rect.top * EMU_PER_MASTER_UNIT)
  return {
    x,
    y,
    cx: Math.max(1, Math.round((rect.right - rect.left) * EMU_PER_MASTER_UNIT)),
    cy: Math.max(1, Math.round((rect.bottom - rect.top) * EMU_PER_MASTER_UNIT)),
  }
}

/** stage-1 fallback for text with no anchor */
function stackedBox(slideWidthEmu: number, slot: number, slots: number): Box {
  const margin = 457200 // 0.5 in
  const height = Math.max(
    600000,
    Math.floor((slideWidthEmu - margin * 2) / Math.max(slots, 1) / 1.6),
  )
  return { x: margin, y: margin + slot * height, cx: slideWidthEmu - margin * 2, cy: height }
}

function colorHex(value: number | undefined): string | null {
  if (value === undefined || (value & 0x19000000) !== 0) return null
  const channel = (shift: number): string => ((value >> shift) & 0xff).toString(16).padStart(2, '0')
  return (channel(0) + channel(8) + channel(16)).toUpperCase()
}

function xfrmXml(box: Box, shape: PptShape): string {
  const flipH = (shape.flags & FSP.flipH) !== 0 ? ' flipH="1"' : ''
  const flipV = (shape.flags & FSP.flipV) !== 0 ? ' flipV="1"' : ''
  // DrawingML rot is in 1/60000 of a degree, clockwise
  const rot = shape.rotation !== 0 ? ` rot="${Math.round(shape.rotation * 60000)}"` : ''
  return `<a:xfrm${rot}${flipH}${flipV}><a:off x="${box.x}" y="${box.y}"/><a:ext cx="${box.cx}" cy="${box.cy}"/></a:xfrm>`
}

function shapeFillXml(shape: PptShape): string {
  const fill = colorHex(shape.props.get(OPT.fillColor))
  if (!fill) return ''
  return `<a:solidFill><a:srgbClr val="${fill}"/></a:solidFill>`
}

function shapeLineXml(shape: PptShape): string {
  const line = colorHex(shape.props.get(OPT.lineColor))
  const width = shape.props.get(OPT.lineWidth)
  if (!line && width === undefined) return ''
  return (
    `<a:ln${width !== undefined ? ` w="${Math.max(0, Math.min(width, 1_000_000))}"` : ''}>` +
    (line ? `<a:solidFill><a:srgbClr val="${line}"/></a:solidFill>` : '') +
    '</a:ln>'
  )
}

function runXml(
  text: string,
  format: TextCharFormat | null,
  isTitle: boolean,
  defaults: TextDefaults,
): string {
  const fallbackPt = isTitle ? (defaults.titleSizePt ?? 32) : (defaults.sizePt ?? 18)
  const size = (format?.sizePt ?? fallbackPt) * 100
  const attrs = [
    'lang="en-US"',
    `sz="${size}"`,
    ...(format?.bold ? ['b="1"'] : []),
    ...(format?.italic ? ['i="1"'] : []),
    ...(format?.underline ? ['u="sng"'] : []),
  ].join(' ')
  const fill = format?.color ? `<a:solidFill><a:srgbClr val="${format.color}"/></a:solidFill>` : ''
  const latin = defaults.typeface ? `<a:latin typeface="${escapeXml(defaults.typeface)}"/>` : ''
  return `<a:r><a:rPr ${attrs}>${fill}${latin}</a:rPr><a:t>${escapeXml(text)}</a:t></a:r>`
}

/** Split text into paragraphs, keeping each character run's formatting. */
function paragraphsXml(
  text: string,
  style: TextStyle,
  isTitle: boolean,
  align: 'l' | 'ctr' | 'r' | 'just' | null,
  defaults: TextDefaults,
): string {
  const runs = applyCharacterRuns(text, style.characters)
  const paragraphs: Array<Array<{ text: string; format: TextCharFormat | null }>> = [[]]
  for (const run of runs) {
    const parts = run.text.split('\n')
    parts.forEach((part, index) => {
      if (index > 0) paragraphs.push([])
      if (part !== '') paragraphs[paragraphs.length - 1]!.push({ text: part, format: run.format })
    })
  }
  const pPr = `<a:pPr algn="${align ?? 'l'}"/>`
  return paragraphs
    .map(
      (paragraph) =>
        `<a:p>${pPr}${paragraph.map((r) => runXml(r.text, r.format, isTitle, defaults)).join('')}</a:p>`,
    )
    .join('')
}

interface ShapeText {
  kind: number
  text: string
  style: TextStyle
  align: 'l' | 'ctr' | 'r' | 'just' | null
}

function spXml(
  shape: PptShape,
  text: ShapeText | null,
  box: Box,
  id: number,
  defaults: TextDefaults,
): string {
  const isTitle = text !== null && TITLE_KINDS.has(text.kind)
  const preset = SHAPE_PRESETS[shape.shapeType] ?? 'rect'
  const body = text ? paragraphsXml(text.text, text.style, isTitle, text.align, defaults) : '<a:p/>'
  return (
    '<p:sp>' +
    '<p:nvSpPr>' +
    `<p:cNvPr id="${id}" name=""/>` +
    (() => {
      const phType = shape.placeholderId
        ? (PLACEHOLDER_TYPES[shape.placeholderId] ?? (isTitle ? 'title' : null))
        : isTitle
          ? 'title'
          : null
      return phType
        ? `<p:cNvSpPr>${isTitle ? '<a:spLocks noGrp="1"/>' : ''}</p:cNvSpPr><p:nvPr><p:ph type="${phType}"/></p:nvPr>`
        : '<p:cNvSpPr/><p:nvPr/>'
    })() +
    '</p:nvSpPr>' +
    `<p:spPr>${xfrmXml(box, shape)}<a:prstGeom prst="${preset}"><a:avLst/></a:prstGeom>` +
    shapeFillXml(shape) +
    shapeLineXml(shape) +
    '</p:spPr>' +
    `<p:txBody><a:bodyPr/><a:lstStyle/>${body}</p:txBody>` +
    '</p:sp>'
  )
}

function picXml(shape: PptShape, box: Box, id: number, relId: string): string {
  return (
    '<p:pic>' +
    '<p:nvPicPr>' +
    `<p:cNvPr id="${id}" name=""/>` +
    '<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/>' +
    '</p:nvPicPr>' +
    `<p:blipFill><a:blip r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
    `<p:spPr>${xfrmXml(box, shape)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>` +
    '</p:pic>'
  )
}

/** One slide's spTree shapes plus its image relationships. */
function slideShapes(
  slide: PptSlide,
  slideWidthEmu: number,
  images: Array<PptImage | null>,
  defaults: TextDefaults,
): { xml: string; rels: Array<{ id: string; imageIndex: number }> } {
  const rels: Array<{ id: string; imageIndex: number }> = []
  const parts: string[] = []
  let id = 2 // the spTree's group shape is 1
  let slot = 0

  const shapes = slide.shapes
  if (shapes.length === 0) {
    const slots = Math.max(slide.textSequences.length, 1)
    for (const sequence of slide.textSequences) {
      parts.push(
        spXml(
          {
            spid: 0,
            rotation: 0,
            placeholderId: null,
            groupBounds: null,
            shapeType: 202,
            flags: 0,
            rect: null,
            props: new Map(),
            text: null,
            blipIndex: null,
          },
          { kind: sequence.kind, text: sequence.text, style: EMPTY_STYLE, align: null },
          stackedBox(slideWidthEmu, slot++, slots),
          id++,
          defaults,
        ),
      )
    }
    return { xml: parts.join(''), rels }
  }

  // Outline text (SlideListWithText) has no atom inside its ClientTextbox;
  // give those shapes their text in outline order, then stack any remainder.
  const used = new Set(
    shapes.map((shape) => shape.text?.text).filter((text): text is string => !!text),
  )
  const leftover = slide.textSequences.filter((sequence) => !used.has(sequence.text))
  let outlineAt = 0
  const emit = (shape: PptShape, text: ShapeText | null, box: Box): void => {
    const image = shape.blipIndex === null ? null : (images[shape.blipIndex - 1] ?? null)
    if (image) {
      const relId = `rId${rels.length + 2}`
      rels.push({ id: relId, imageIndex: shape.blipIndex! - 1 })
      parts.push(picXml(shape, box, id++, relId))
    } else {
      parts.push(spXml(shape, text, box, id++, defaults))
    }
  }

  for (const shape of shapes) {
    let text: ShapeText | null = shape.text
    if (!text && outlineAt < leftover.length) {
      const sequence: PptSlideText = leftover[outlineAt++]!
      text = { kind: sequence.kind, text: sequence.text, style: EMPTY_STYLE, align: null }
    }
    const box = shape.rect
      ? boxOf(shape.rect)
      : stackedBox(slideWidthEmu, slot++, Math.max(shapes.length, 1))
    emit(shape, text, box)
  }
  while (outlineAt < leftover.length) {
    const sequence = leftover[outlineAt++]!
    emit(
      {
        spid: 0,
        rotation: 0,
        placeholderId: null,
        groupBounds: null,
        shapeType: 202,
        flags: 0,
        rect: null,
        props: new Map(),
        text: null,
        blipIndex: null,
      },
      { kind: sequence.kind, text: sequence.text, style: EMPTY_STYLE, align: null },
      stackedBox(slideWidthEmu, slot++, Math.max(shapes.length + leftover.length, 1)),
    )
  }
  return { xml: parts.join(''), rels }
}

const spTreeHead =
  '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'

export async function writePptxPresentation(presentation: PptPresentation): Promise<Uint8Array> {
  const zip = new JSZip()
  const slideWidthEmu = Math.round(presentation.slideWidth * EMU_PER_MASTER_UNIT)
  const slideHeightEmu = Math.round(presentation.slideHeight * EMU_PER_MASTER_UNIT)
  const images = presentation.images

  // Media parts, named by their 1-based blip index
  images.forEach((image, index) => {
    if (image) zip.file(`ppt/media/image${index + 1}.${image.ext}`, image.bytes)
  })

  const slideRels: string[] = []
  const defaults = textDefaultsOf(presentation)
  const slideXmls = presentation.slides.map((slide, index) => {
    const { xml: shapes, rels } = slideShapes(slide, slideWidthEmu, images, defaults)
    const relations =
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>' +
      rels
        .map(
          (rel) =>
            `<Relationship Id="${rel.id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image${rel.imageIndex + 1}.${images[rel.imageIndex]!.ext}"/>`,
        )
        .join('')
    slideRels[index] =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      relations +
      '</Relationships>'
    return (
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
      'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
      `<p:cSld><p:spTree>${spTreeHead}${shapes}</p:spTree></p:cSld>` +
      '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'
    )
  })

  const imageDefaults = [
    ...new Set(images.filter((i): i is PptImage => i !== null).map((i) => i.ext)),
  ]
    .map((ext) => {
      const type =
        ext === 'png'
          ? 'image/png'
          : ext === 'emf'
            ? 'image/x-emf'
            : ext === 'wmf'
              ? 'image/x-wmf'
              : 'image/jpeg'
      return `<Default Extension="${ext}" ContentType="${type}"/>`
    })
    .join('')
  const contentTypes =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    imageDefaults +
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
  slideRels.forEach((rels, i) => zip.file(`ppt/slides/_rels/slide${i + 1}.xml.rels`, rels))

  // the main master's anchored shapes are inherited by every slide
  const masterShapes = (presentation.master?.shapes ?? [])
    .filter((shape) => shape.rect !== null && shape.blipIndex === null)
    .map((shape, index) => spXml(shape, shape.text, boxOf(shape.rect!), index + 2, defaults))
    .join('')
  zip.file(
    'ppt/slideMasters/slideMaster1.xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
      'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
      `<p:cSld><p:spTree>${spTreeHead}${masterShapes}</p:spTree></p:cSld>` +
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

  slideXmls.forEach((xml, i) => zip.file(`ppt/slides/slide${i + 1}.xml`, xml))
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}
