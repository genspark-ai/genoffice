import JSZip from 'jszip'

import type { PictureMedia } from './escher'
import {
  EMU_PER_MASTER_UNIT,
  type PptAnyShape,
  type PptFill,
  type PptLine,
  type PptPresentation,
  type PptTextBody,
} from './ppt'
import { AUTO_NUMBER_SCHEMES, type ResolvedParagraph, type ResolvedRun } from './text-style'

/**
 * The OOXML presentation the conversion lands in: one slide per PPT slide
 * with its background and shapes in z-order, one media part per distinct
 * picture, and the blank master/layout skeleton PowerPoint requires.
 */

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

const NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'

const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

const EMPTY_TREE =
  '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'

interface SlideMedia {
  /** pib → relationship id */
  rels: Map<number, string>
  relXml: string[]
}

class MediaRegistry {
  /** pib → media file name */
  readonly files = new Map<number, { name: string; media: PictureMedia }>()
  readonly extensions = new Set<string>()
  constructor(private readonly presentation: PptPresentation) {}

  /** the media file a pib maps to (one part per distinct picture across the deck) */
  file(pib: number): string | null {
    const known = this.files.get(pib)
    if (known) return known.name
    const media = this.presentation.pictures.get(pib)
    if (!media) return null
    const name = `image${this.files.size + 1}.${media.extension}`
    this.files.set(pib, { name, media })
    this.extensions.add(media.extension)
    return name
  }
}

function solidFill(color: string, alpha?: number): string {
  const alphaXml =
    alpha !== undefined && alpha < 1 ? `<a:alpha val="${Math.round(alpha * 100000)}"/>` : ''
  return `<a:solidFill><a:srgbClr val="${color}">${alphaXml}</a:srgbClr></a:solidFill>`
}

function fillXml(
  fill: PptFill | undefined,
  slideMedia: SlideMedia,
  registry: MediaRegistry,
): string {
  if (!fill || fill.kind === 'background') return '<a:noFill/>'
  if (fill.kind === 'solid') return solidFill(fill.color, fill.alpha)
  if (fill.kind === 'pattern') {
    return `<a:pattFill prst="${fill.preset}"><a:fgClr><a:srgbClr val="${fill.fg}"/></a:fgClr><a:bgClr><a:srgbClr val="${fill.bg}"/></a:bgClr></a:pattFill>`
  }
  if (fill.kind === 'gradient') {
    const stops = fill.stops
      .map(
        (stop) =>
          `<a:gs pos="${Math.round(stop.pos * 1000)}"><a:srgbClr val="${stop.color}"/></a:gs>`,
      )
      .join('')
    const pct = (value: number): number => Math.round(value * 100000)
    const radial = fill.radial
    const shade = radial
      ? `<a:path path="${radial.path}"><a:fillToRect l="${pct(radial.l)}" t="${pct(radial.t)}" r="${pct(radial.r)}" b="${pct(radial.b)}"/></a:path>`
      : `<a:lin ang="${Math.round(fill.angle * 60000)}" scaled="0"/>`
    return `<a:gradFill rotWithShape="1"><a:gsLst>${stops}</a:gsLst>${shade}</a:gradFill>`
  }
  const rId = pictureRel(fill.pib, slideMedia, registry)
  if (!rId) return '<a:noFill/>'
  const placement = fill.tile
    ? '<a:tile tx="0" ty="0" sx="100000" sy="100000" flip="none" algn="tl"/>'
    : '<a:stretch><a:fillRect/></a:stretch>'
  return `<a:blipFill><a:blip r:embed="${rId}"/>${placement}</a:blipFill>`
}

function pictureRel(pib: number, slideMedia: SlideMedia, registry: MediaRegistry): string | null {
  const known = slideMedia.rels.get(pib)
  if (known) return known
  const file = registry.file(pib)
  if (!file) return null
  const rId = `rId${slideMedia.rels.size + 2}`
  slideMedia.rels.set(pib, rId)
  slideMedia.relXml.push(
    `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${file}"/>`,
  )
  return rId
}

function lineXml(line: PptLine | undefined): string {
  if (!line) return '<a:ln><a:noFill/></a:ln>'
  const cmpd = line.compound ? ` cmpd="${line.compound}"` : ''
  const dash = line.dash ? `<a:prstDash val="${line.dash}"/>` : ''
  return `<a:ln w="${Math.max(0, Math.round(line.width))}"${cmpd}>${solidFill(line.color)}${dash}</a:ln>`
}

function xfrmXml(shape: PptAnyShape): string {
  const attrs =
    (shape.rotation ? ` rot="${Math.round(shape.rotation * 60000)}"` : '') +
    (shape.flipH ? ' flipH="1"' : '') +
    (shape.flipV ? ' flipV="1"' : '')
  return `<a:xfrm${attrs}><a:off x="${shape.x}" y="${shape.y}"/><a:ext cx="${shape.cx}" cy="${shape.cy}"/></a:xfrm>`
}

function geometryXml(shape: Extract<PptAnyShape, { type: 'shape' }>): string {
  if (shape.path) {
    const { width, height, commands } = shape.path
    const pt = ([x, y]: [number, number]): string =>
      `<a:pt x="${Math.round(x)}" y="${Math.round(y)}"/>`
    let body = ''
    for (const command of commands) {
      if (command.kind === 'move') body += `<a:moveTo>${pt(command.points[0]!)}</a:moveTo>`
      else if (command.kind === 'line')
        body += command.points.map((p) => `<a:lnTo>${pt(p)}</a:lnTo>`).join('')
      else if (command.kind === 'curve') {
        for (let index = 0; index + 3 <= command.points.length; index += 3) {
          body += `<a:cubicBezTo>${pt(command.points[index]!)}${pt(command.points[index + 1]!)}${pt(command.points[index + 2]!)}</a:cubicBezTo>`
        }
      } else body += '<a:close/>'
    }
    return `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="${Math.max(1, width)}" h="${Math.max(1, height)}">${body}</a:path></a:pathLst></a:custGeom>`
  }
  const guides = Object.entries(shape.adjust ?? {})
    .map(([name, value]) => `<a:gd name="${name}" fmla="val ${Math.round(value)}"/>`)
    .join('')
  return `<a:prstGeom prst="${shape.preset ?? 'rect'}"><a:avLst>${guides}</a:avLst></a:prstGeom>`
}

/** master units (1/8 pt) → hundredths of a point for spcPts */
const masterToHundredths = (value: number): number => Math.round((Math.abs(value) / 8) * 100)

function spacingXml(tag: string, value: number | undefined): string {
  if (value === undefined || value === 0) return ''
  if (value > 0) return `<a:${tag}><a:spcPct val="${Math.round(value * 1000)}"/></a:${tag}>`
  return `<a:${tag}><a:spcPts val="${masterToHundredths(value)}"/></a:${tag}>`
}

let fieldCounter = 0
function fieldId(): string {
  fieldCounter++
  return `{${fieldCounter.toString(16).padStart(8, '0').toUpperCase()}-0000-0000-0000-000000000000}`
}

function fieldText(run: ResolvedRun, slideNumber: number): string {
  if (run.field === 'slidenum') return String(slideNumber)
  if (run.field === 'datetime1') {
    const now = new Date()
    return `${now.getMonth() + 1}/${now.getDate()}/${now.getFullYear()}`
  }
  return run.text === '*' ? '' : run.text
}

function runXml(run: ResolvedRun, slideNumber: number): string {
  const attrs =
    ` lang="en-US" sz="${Math.round(run.size * 100)}"` +
    (run.bold ? ' b="1"' : '') +
    (run.italic ? ' i="1"' : '') +
    (run.underline ? ' u="sng"' : '') +
    (run.position ? ` baseline="${Math.round(run.position * 1000)}"` : '')
  const fonts =
    (run.font ? `<a:latin typeface="${escapeXml(run.font)}"/>` : '') +
    (run.eaFont ? `<a:ea typeface="${escapeXml(run.eaFont)}"/>` : '')
  const shadow = run.shadow
    ? `<a:effectLst><a:outerShdw blurRad="38100" dist="38100" dir="2700000" algn="tl"><a:srgbClr val="${run.shadow}"><a:alpha val="43137"/></a:srgbClr></a:outerShdw></a:effectLst>`
    : ''
  const rPr = `<a:rPr${attrs}>${solidFill(run.color)}${shadow}${fonts}</a:rPr>`
  if (run.field) {
    return `<a:fld id="${fieldId()}" type="${run.field === 'datetime1' ? 'datetime1' : run.field}">${rPr}<a:t>${escapeXml(fieldText(run, slideNumber))}</a:t></a:fld>`
  }
  // '\u000b' is PowerPoint's soft line break
  return run.text
    .split('\u000b')
    .map((piece) => (piece.length ? `<a:r>${rPr}<a:t>${escapeXml(piece)}</a:t></a:r>` : ''))
    .join(`<a:br>${rPr}</a:br>`)
}

const ALIGN = ['l', 'ctr', 'r', 'just', 'dist']
const TAB_ALIGN = ['l', 'ctr', 'r', 'dec']

function bulletMarkXml(paragraph: ResolvedParagraph): string {
  const scheme = paragraph.autoNumber && AUTO_NUMBER_SCHEMES[paragraph.autoNumber.scheme]
  if (!scheme) return `<a:buChar char="${escapeXml(paragraph.bulletChar)}"/>`
  const startAt = paragraph.autoNumber!.startAt
  return `<a:buAutoNum type="${scheme}"${startAt > 1 ? ` startAt="${startAt}"` : ''}/>`
}

function paragraphXml(paragraph: ResolvedParagraph, slideNumber: number): string {
  const marL = Math.round(paragraph.leftMargin * EMU_PER_MASTER_UNIT)
  const indent = Math.round((paragraph.indent - paragraph.leftMargin) * EMU_PER_MASTER_UNIT)
  const attrs =
    ` marL="${Math.max(0, marL)}" indent="${indent}" lvl="${paragraph.level}"` +
    ` algn="${ALIGN[paragraph.align] ?? 'l'}"`
  let bullet = '<a:buNone/>'
  if (paragraph.bullet && paragraph.runs.length > 0) {
    bullet =
      (paragraph.bulletColor
        ? `<a:buClr><a:srgbClr val="${paragraph.bulletColor}"/></a:buClr>`
        : '') +
      (paragraph.bulletSize && paragraph.bulletSize > 0
        ? `<a:buSzPct val="${paragraph.bulletSize * 1000}"/>`
        : '') +
      (paragraph.bulletFont ? `<a:buFont typeface="${escapeXml(paragraph.bulletFont)}"/>` : '') +
      bulletMarkXml(paragraph)
  }
  const tabs = paragraph.tabs
    ? `<a:tabLst>${paragraph.tabs
        .map(
          (tab) =>
            `<a:tab pos="${Math.round(tab.position * EMU_PER_MASTER_UNIT)}" algn="${TAB_ALIGN[tab.type] ?? 'l'}"/>`,
        )
        .join('')}</a:tabLst>`
    : ''
  const pPr =
    `<a:pPr${attrs}>` +
    spacingXml('lnSpc', paragraph.lineSpacing) +
    spacingXml('spcBef', paragraph.spaceBefore) +
    spacingXml('spcAft', paragraph.spaceAfter) +
    bullet +
    tabs +
    '</a:pPr>'
  const runs = paragraph.runs.map((run) => runXml(run, slideNumber)).join('')
  const endPr = `<a:endParaRPr lang="en-US" sz="${Math.round(paragraph.endSize * 100)}"/>`
  return `<a:p>${pPr}${runs}${endPr}</a:p>`
}

function textBodyXml(text: PptTextBody, slideNumber: number): string {
  const [l, t, r, b] = text.insets
  const bodyAttrs =
    ` wrap="${text.wrap ? 'square' : 'none'}" lIns="${l}" tIns="${t}" rIns="${r}" bIns="${b}"` +
    ` anchor="${text.anchor}"` +
    (text.anchorCenter ? ' anchorCtr="1"' : '') +
    (text.vertical ? ` vert="${text.vertical}"` : '')
  const autofit = text.autofit ? '<a:spAutoFit/>' : ''
  return `<p:txBody><a:bodyPr${bodyAttrs}>${autofit}</a:bodyPr><a:lstStyle/>${text.paragraphs.map((p) => paragraphXml(p, slideNumber)).join('')}</p:txBody>`
}

function shapeXml(
  shape: PptAnyShape,
  id: number,
  slideMedia: SlideMedia,
  registry: MediaRegistry,
  slideNumber: number,
): string {
  if (shape.type === 'picture') {
    const rId = pictureRel(shape.pib, slideMedia, registry)
    if (!rId) return ''
    const crop = shape.crop
      ? `<a:srcRect l="${shape.crop.l}" t="${shape.crop.t}" r="${shape.crop.r}" b="${shape.crop.b}"/>`
      : ''
    const clrChange = shape.transparent
      ? `<a:clrChange><a:clrFrom><a:srgbClr val="${shape.transparent}"/></a:clrFrom><a:clrTo><a:srgbClr val="${shape.transparent}"><a:alpha val="0"/></a:srgbClr></a:clrTo></a:clrChange>`
      : ''
    return (
      '<p:pic>' +
      `<p:nvPicPr><p:cNvPr id="${id}" name="Picture ${id}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>` +
      `<p:blipFill><a:blip r:embed="${rId}">${clrChange}</a:blip>${crop}<a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
      `<p:spPr>${xfrmXml(shape)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${lineXml(shape.line)}</p:spPr>` +
      '</p:pic>'
    )
  }
  const useBgFill = shape.fill?.kind === 'background'
  return (
    `<p:sp${useBgFill ? ' useBgFill="1"' : ''}>` +
    `<p:nvSpPr><p:cNvPr id="${id}" name="Shape ${id}"/><p:cNvSpPr${shape.text && !shape.fill && !shape.line ? ' txBox="1"' : ''}/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr>${xfrmXml(shape)}${geometryXml(shape)}${useBgFill ? '' : fillXml(shape.fill, slideMedia, registry)}${lineXml(shape.line)}</p:spPr>` +
    (shape.text ? textBodyXml(shape.text, slideNumber) : '') +
    '</p:sp>'
  )
}

const MEDIA_CONTENT_TYPES: Record<string, string> = {
  png: 'image/png',
  jpeg: 'image/jpeg',
  emf: 'image/x-emf',
  wmf: 'image/x-wmf',
  bmp: 'image/bmp',
  tiff: 'image/tiff',
}

export async function writePptxPresentation(
  presentation: PptPresentation,
): Promise<{ bytes: Uint8Array; warnings: string[] }> {
  const zip = new JSZip()
  const slideWidthEmu = Math.round(presentation.slideWidth * EMU_PER_MASTER_UNIT)
  const slideHeightEmu = Math.round(presentation.slideHeight * EMU_PER_MASTER_UNIT)
  const registry = new MediaRegistry(presentation)
  const warnings: string[] = []
  let missingPictures = 0

  presentation.slides.forEach((slide, index) => {
    const slideMedia: SlideMedia = { rels: new Map(), relXml: [] }
    const background = slide.background
      ? `<p:bg><p:bgPr>${fillXml(slide.background, slideMedia, registry)}<a:effectLst/></p:bgPr></p:bg>`
      : ''
    let nextId = 2
    const shapes = slide.shapes
      .map((shape) => {
        if (shape.type === 'picture' && !registry.file(shape.pib)) {
          missingPictures++
          return ''
        }
        return shapeXml(shape, nextId++, slideMedia, registry, index + 1)
      })
      .join('')
    const xml =
      `${XML_HEAD}<p:sld ${NS}><p:cSld>${background}<p:spTree>${EMPTY_TREE}${shapes}</p:spTree></p:cSld>` +
      '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>'
    zip.file(`ppt/slides/slide${index + 1}.xml`, xml)
    zip.file(
      `ppt/slides/_rels/slide${index + 1}.xml.rels`,
      `${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>' +
        slideMedia.relXml.join('') +
        '</Relationships>',
    )
  })
  if (missingPictures > 0)
    warnings.push(`${missingPictures} picture(s) could not be decoded and were omitted`)

  for (const { name, media } of registry.files.values()) zip.file(`ppt/media/${name}`, media.bytes)

  const slideCount = presentation.slides.length
  zip.file(
    '[Content_Types].xml',
    `${XML_HEAD}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      [...registry.extensions]
        .map(
          (ext) =>
            `<Default Extension="${ext}" ContentType="${MEDIA_CONTENT_TYPES[ext] ?? 'application/octet-stream'}"/>`,
        )
        .join('') +
      '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>' +
      Array.from(
        { length: slideCount },
        (_s, i) =>
          `<Override PartName="/ppt/slides/slide${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`,
      ).join('') +
      '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>' +
      '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>' +
      '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' +
      '</Types>',
  )

  zip.file(
    '_rels/.rels',
    `${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>' +
      '</Relationships>',
  )

  zip.file(
    'ppt/presentation.xml',
    `${XML_HEAD}<p:presentation ${NS}>` +
      '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>' +
      `<p:sldIdLst>${Array.from({ length: slideCount }, (_s, i) => `<p:sldId id="${256 + i}" r:id="rId${i + 2}"/>`).join('')}</p:sldIdLst>` +
      `<p:sldSz cx="${slideWidthEmu}" cy="${slideHeightEmu}"/>` +
      '<p:notesSz cx="6858000" cy="9144000"/></p:presentation>',
  )

  zip.file(
    'ppt/_rels/presentation.xml.rels',
    `${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>' +
      Array.from(
        { length: slideCount },
        (_s, i) =>
          `<Relationship Id="rId${i + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`,
      ).join('') +
      `<Relationship Id="rId${slideCount + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>` +
      '</Relationships>',
  )

  // the master and layout skeletons: a blank master, one blank layout pointing back, a theme
  zip.file(
    'ppt/slideMasters/slideMaster1.xml',
    `${XML_HEAD}<p:sldMaster ${NS}><p:cSld><p:spTree>${EMPTY_TREE}</p:spTree></p:cSld>` +
      '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>' +
      '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>' +
      '</p:sldMaster>',
  )
  zip.file(
    'ppt/slideMasters/_rels/slideMaster1.xml.rels',
    `${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>' +
      '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>' +
      '</Relationships>',
  )
  zip.file(
    'ppt/slideLayouts/slideLayout1.xml',
    `${XML_HEAD}<p:sldLayout ${NS} type="blank"><p:cSld name="Blank"><p:spTree>${EMPTY_TREE}</p:spTree></p:cSld>` +
      '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>',
  )
  zip.file(
    'ppt/slideLayouts/_rels/slideLayout1.xml.rels',
    `${XML_HEAD}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>' +
      '</Relationships>',
  )
  zip.file('ppt/theme/theme1.xml', THEME_XML)

  const bytes = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
  return { bytes, warnings }
}

const THEME_XML =
  `${XML_HEAD}<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements>` +
  '<a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>' +
  '<a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1>' +
  '<a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4>' +
  '<a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink>' +
  '<a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme>' +
  '<a:fontScheme name="Office"><a:majorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>' +
  '<a:minorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>' +
  '<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>' +
  '<a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="25400"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="38100"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>' +
  '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>' +
  '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>' +
  '</a:fmtScheme></a:themeElements></a:theme>'
