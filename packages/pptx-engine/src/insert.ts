/**
 * Element insertion — synthesizes a raw <p:sp> fragment and hangs it on
 * slide.elements.
 *
 * Naturally compatible with patch-based saving: a new element's
 * anchor.originalXml is the generated XML, which patchSlideXml includes when
 * splicing elements together; deletion is removal from the array.
 * Both change the spTree structure, driving a full-slide rebuild via
 * slide.structureDirty.
 */
import type { EmuRect, Paragraph, PictureElement, Slide, SlideElement, TextElement } from './types'
import { generateParagraphXml, generateXfrmXml } from './generate'
import { escapeXmlAttr } from './xml-utils'
import { relsPathFor } from './zip'
import type { OpenedPptx } from './index'
import { cleanupDeletedElementResources } from './resource-cleanup'

/**
 * 'textbox' is a special value (plain text box without prstGeom); anything else is
 * an OOXML preset geometry name (rect/roundRect/ellipse/triangle/star5/rightArrow/
 * chevron…). Presets whose polygon approximation the render layer hasn't
 * implemented fall back to a rectangle; always correct in PowerPoint.
 */
export type NewShapeKind = 'textbox' | (string & {})

export interface NewElementOptions {
  kind: NewShapeKind
  offset: EmuRect
  paragraphs?: Paragraph[]
  /** Solid shape fill (#RRGGBB); textbox has no fill by default */
  fillColor?: string
  /** Shape stroke (solid color, width in EMU) */
  stroke?: { color: string; widthEmu: number }
}

let insertCounter = 1

// ── Line / connector insertion ─────────────────────────────

/** Insertable line/connector kinds: p:cxnSp fragments with optional arrow ends */
const LINE_KINDS: Record<string, { prst: string; head?: boolean; tail?: boolean }> = {
  line: { prst: 'line' },
  lineArrow: { prst: 'straightConnector1', tail: true },
  lineArrowDouble: { prst: 'straightConnector1', head: true, tail: true },
  lineBent: { prst: 'bentConnector3' },
  lineCurved: { prst: 'curvedConnector3' },
}

export function isLineKind(kind: string): boolean {
  return Object.prototype.hasOwnProperty.call(LINE_KINDS, kind)
}

const DEFAULT_LINE_STROKE = { color: '#000000', widthEmu: 12700 }

function buildCxnSpXml(
  slide: Slide,
  opts: NewElementOptions,
  def: { prst: string; head?: boolean; tail?: boolean },
): string {
  const id = nextCNvPrId(slide)
  const name = `${
    def.prst.startsWith('bentConnector')
      ? 'Elbow Connector'
      : def.prst.startsWith('curvedConnector')
        ? 'Curved Connector'
        : 'Straight Connector'
  } ${id}`
  const o = opts.offset
  const stroke = opts.stroke ?? DEFAULT_LINE_STROKE
  const color = stroke.color.replace(/^#/, '').slice(0, 6).toUpperCase()
  const head = def.head ? '<a:headEnd type="triangle" w="med" len="med"/>' : ''
  const tail = def.tail ? '<a:tailEnd type="triangle" w="med" len="med"/>' : ''
  return (
    `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${id}" name="${escapeXmlAttr(name)}"/>` +
    '<p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>' +
    `<p:spPr><a:xfrm><a:off x="${o.x}" y="${o.y}"/><a:ext cx="${o.cx}" cy="${o.cy}"/></a:xfrm>` +
    `<a:prstGeom prst="${def.prst}"><a:avLst/></a:prstGeom>` +
    `<a:ln w="${Math.round(stroke.widthEmu)}" cap="flat">` +
    `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>${head}${tail}</a:ln>` +
    '</p:spPr></p:cxnSp>'
  )
}

/** Max cNvPr id used in the slide (including new elements); new elements take max+1 */
export function nextCNvPrId(slide: Slide): number {
  let max = 1
  const scan = (xml: string) => {
    for (const m of xml.matchAll(/<p:cNvPr\s[^>]*\bid="(\d+)"/g)) {
      max = Math.max(max, Number(m[1]))
    }
  }
  scan(slide.originalXml)
  for (const el of slide.elements) scan(el.anchor.originalXml)
  return max + 1
}

export function buildSpXml(slide: Slide, opts: NewElementOptions): string {
  const id = nextCNvPrId(slide)
  const isTextbox = opts.kind === 'textbox'
  const name = isTextbox ? `TextBox ${id}` : `Shape ${id}`
  const o = opts.offset
  const xfrm = `<a:xfrm><a:off x="${o.x}" y="${o.y}"/><a:ext cx="${o.cx}" cy="${o.cy}"/></a:xfrm>`
  // Parser convention: has txBody and no prstGeom → 'text'; textbox omits prstGeom
  const geom = isTextbox
    ? ''
    : `<a:prstGeom prst="${escapeXmlAttr(opts.kind)}"><a:avLst/></a:prstGeom>`
  const fill = opts.fillColor
    ? `<a:solidFill><a:srgbClr val="${opts.fillColor.replace(/^#/, '').slice(0, 6).toUpperCase()}"/></a:solidFill>`
    : ''
  const ln = opts.stroke
    ? `<a:ln w="${Math.round(opts.stroke.widthEmu)}"><a:solidFill><a:srgbClr val="${opts.stroke.color.replace(/^#/, '').slice(0, 6).toUpperCase()}"/></a:solidFill></a:ln>`
    : ''
  const paras = (opts.paragraphs?.length ? opts.paragraphs : [{ runs: [{ text: '' }] }])
    .map((p) => generateParagraphXml(p))
    .join('')
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${escapeXmlAttr(name)}"/>` +
    `<p:cNvSpPr${isTextbox ? ' txBox="1"' : ''}/><p:nvPr/></p:nvSpPr>` +
    `<p:spPr>${xfrm}${geom}${fill}${ln}</p:spPr>` +
    `<p:txBody><a:bodyPr wrap="square" rtlCol="0"/><a:lstStyle/>${paras}</p:txBody></p:sp>`
  )
}

/** Synthesize a new element and hang it on the slide; returns the model element (immediately usable by the render layer). */
export function addElement(slide: Slide, opts: NewElementOptions): TextElement {
  const lineDef = LINE_KINDS[opts.kind]
  if (lineDef) {
    const stroke = opts.stroke ?? DEFAULT_LINE_STROKE
    const el: TextElement = {
      id: `spnew_${(insertCounter++).toString(36)}_${Date.now().toString(36)}`,
      type: 'shape',
      anchor: {
        spIndex: slide.elements.length,
        originalXml: buildCxnSpXml(slide, opts, lineDef),
        range: [0, 0],
      },
      transform: { offset: { ...opts.offset }, rot: 0, flipH: false, flipV: false },
      presetGeometry: lineDef.prst,
      fill: { type: 'none' },
      stroke: {
        fill: { type: 'solid', color: stroke.color },
        width: Math.round(stroke.widthEmu),
        ...(lineDef.head ? { headEnd: { type: 'triangle' as const } } : {}),
        ...(lineDef.tail ? { tailEnd: { type: 'triangle' as const } } : {}),
      },
    }
    slide.elements.push(el)
    slide.structureDirty = true
    return el
  }
  const xml = buildSpXml(slide, opts)
  const el: TextElement = {
    id: `spnew_${(insertCounter++).toString(36)}_${Date.now().toString(36)}`,
    type: opts.kind === 'textbox' ? 'text' : 'shape',
    anchor: { spIndex: slide.elements.length, originalXml: xml, range: [0, 0] },
    transform: { offset: { ...opts.offset }, rot: 0, flipH: false, flipV: false },
    ...(opts.kind !== 'textbox' ? { presetGeometry: opts.kind } : {}),
    ...(opts.fillColor ? { fill: { type: 'solid' as const, color: opts.fillColor } } : {}),
    ...(opts.stroke
      ? {
          stroke: {
            fill: { type: 'solid' as const, color: opts.stroke.color },
            width: Math.round(opts.stroke.widthEmu),
          },
        }
      : {}),
    text: { paragraphs: opts.paragraphs?.length ? opts.paragraphs : [{ runs: [{ text: '' }] }] },
  }
  slide.elements.push(el)
  slide.structureDirty = true
  return el
}

// ── Table insertion (graphicFrame + a:tbl) ─────────────────────────────

export interface NewTableOptions {
  rows: number
  cols: number
  offset: EmuRect
}

/** PowerPoint's default style for new tables (Medium Style 2 - Accent 1, built-in fallback in the render layer) */
const DEFAULT_TABLE_STYLE_ID = '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}'

/**
 * Build the table graphicFrame fragment (equal-width columns / equal-height rows,
 * default built-in style, empty cells). Insertion goes through appendRawElements
 * (materialize+reparse), reusing the existing table parsing/rendering pipeline.
 */
export function buildTableXml(slide: Slide, opts: NewTableOptions): string {
  const id = nextCNvPrId(slide)
  const rows = Math.max(1, Math.floor(opts.rows))
  const cols = Math.max(1, Math.floor(opts.cols))
  const colW = Math.max(1, Math.floor(opts.offset.cx / cols))
  const rowH = Math.max(1, Math.floor(opts.offset.cy / rows))
  const grid = Array.from({ length: cols }, () => `<a:gridCol w="${colW}"/>`).join('')
  const cell = '<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody><a:tcPr/></a:tc>'
  const trs = Array.from(
    { length: rows },
    () => `<a:tr h="${rowH}">${cell.repeat(cols)}</a:tr>`,
  ).join('')
  return (
    `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="Table ${id}"/>` +
    '<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>' +
    `<p:xfrm><a:off x="${opts.offset.x}" y="${opts.offset.y}"/><a:ext cx="${opts.offset.cx}" cy="${opts.offset.cy}"/></p:xfrm>` +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">' +
    `<a:tbl><a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>${DEFAULT_TABLE_STYLE_ID}</a:tableStyleId></a:tblPr>` +
    `<a:tblGrid>${grid}</a:tblGrid>${trs}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`
  )
}

// ── Picture insertion (media part surgery) ─────────────────────────────

const IMAGE_MIME: Record<string, string> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  webp: 'image/webp',
  tif: 'image/tiff',
  tiff: 'image/tiff',
}

const IMAGE_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'

export interface NewPictureOptions {
  /** Image bytes */
  bytes: Uint8Array
  /** Lowercase extension (png/jpg/…) */
  ext: string
  offset: EmuRect
  /** cNvPr name (default Picture N; hand-drawn ink is marked with the aislides-ink prefix) */
  name?: string
  /** cNvPr descr: editor-specific payload (e.g. ink vector points), recoverable on reopen */
  descr?: string
}

/**
 * Insert a picture: media bytes into the package + [Content_Types] Default +
 * slide rels registration + synthesized <p:pic> fragment hung on slide.elements.
 * The dataUrl is generated on demand by the caller (media resolver).
 */
/**
 * Land an image into the package: media part + Content_Types Default + slide rels.
 * Returns the new relationship id and media path (shared by picture insertion /
 * shape picture fill).
 */
export function addImageMediaAndRel(
  opened: OpenedPptx,
  slide: Slide,
  bytes: Uint8Array,
  extRaw: string,
): { rid: string; mediaPath: string } | null {
  const { archive } = opened
  const ext = extRaw.toLowerCase()
  const mime = IMAGE_MIME[ext]
  if (!mime) return null

  // 1) media part: number = current max + 1
  let maxNum = 0
  for (const path of archive.entries.keys()) {
    const m = /^ppt\/media\/image(\d+)\./.exec(path)
    if (m) maxNum = Math.max(maxNum, Number(m[1]))
  }
  const mediaPath = `ppt/media/image${maxNum + 1}.${ext}`
  archive.entries.set(mediaPath, bytes)

  // 2) [Content_Types] Default (added the first time this extension appears)
  const ctPath = '[Content_Types].xml'
  const ct = archive.readText(ctPath)
  if (ct && !new RegExp(`<Default Extension="${ext}"`).test(ct)) {
    const dflt = `<Default Extension="${ext}" ContentType="${mime}"/>`
    archive.entries.set(ctPath, Buffer.from(ct.replace('</Types>', `${dflt}</Types>`), 'utf8'))
  }

  // 3) slide rels: new rId (the rels file may not exist)
  const relsPath = relsPathFor(slide.path)
  const rels =
    archive.readText(relsPath) ??
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
  let maxRid = 0
  for (const m of rels.matchAll(/Id="rId(\d+)"/g)) maxRid = Math.max(maxRid, Number(m[1]))
  const rid = `rId${maxRid + 1}`
  const relXml = `<Relationship Id="${rid}" Type="${IMAGE_REL_TYPE}" Target="../media/image${maxNum + 1}.${ext}"/>`
  archive.entries.set(
    relsPath,
    Buffer.from(rels.replace('</Relationships>', `${relXml}</Relationships>`), 'utf8'),
  )
  return { rid, mediaPath }
}

export function addPicture(
  opened: OpenedPptx,
  slide: Slide,
  opts: NewPictureOptions,
): PictureElement | null {
  const added = addImageMediaAndRel(opened, slide, opts.bytes, opts.ext)
  if (!added) return null
  const { rid, mediaPath } = added

  // 4) <p:pic> fragment
  const id = nextCNvPrId(slide)
  const name = opts.name ?? `Picture ${id}`
  const descrAttr = opts.descr ? ` descr="${escapeXmlAttr(opts.descr)}"` : ''
  const xml =
    `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${escapeXmlAttr(name)}"${descrAttr}/>` +
    '<p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
    `<p:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
    `<p:spPr>${generateXfrmXml({ offset: opts.offset, rot: 0, flipH: false, flipV: false })}` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'

  const el: PictureElement = {
    id: `picnew_${(insertCounter++).toString(36)}_${Date.now().toString(36)}`,
    type: 'picture',
    anchor: { spIndex: slide.elements.length, originalXml: xml, range: [0, 0] },
    transform: { offset: { ...opts.offset }, rot: 0, flipH: false, flipV: false },
    name,
    ...(opts.descr ? { descr: opts.descr } : {}),
    mediaRef: mediaPath,
  }
  slide.elements.push(el)
  slide.structureDirty = true
  return el
}

/**
 * Delete by element id and collect relationships/resources no longer used by any
 * retained package part. The OpenedPptx context is required because media may be
 * shared by other objects or slides.
 */
export function deleteElement(opened: OpenedPptx, slide: Slide, elementId: string): boolean {
  const idx = slide.elements.findIndex((e) => e.id === elementId)
  if (idx < 0) return false
  const removedXml = slide.elements[idx]!.anchor.originalXml
  slide.elements.splice(idx, 1)
  slide.structureDirty = true
  cleanupDeletedElementResources(opened, slide, removedXml)
  return true
}

// ── Grouping (p:grpSp) ──────────────────────────────────────────────────────

/**
 * Compute the bounding box of a set of elements (slide coordinates, EMU).
 * Ignores rotation: uses the axis-aligned bounding box of each element's offset
 * rect.
 */
export function calcBoundingBox(elements: SlideElement[]): EmuRect {
  let minX = Infinity,
    minY = Infinity,
    maxX = -Infinity,
    maxY = -Infinity
  for (const el of elements) {
    const o = el.transform.offset
    minX = Math.min(minX, o.x)
    minY = Math.min(minY, o.y)
    maxX = Math.max(maxX, o.x + o.cx)
    maxY = Math.max(maxY, o.y + o.cy)
  }
  return { x: minX, y: minY, cx: maxX - minX, cy: maxY - minY }
}

/**
 * Build the <p:grpSp> XML fragment.
 *
 * OOXML conventions (ECMA 376 §19.3.1.22):
 *  - grpSpPr/xfrm describes the group's position and size on the slide (<a:off>/<a:ext>)
 *  - grpSpPr/xfrm/chOff + chExt define the child coordinate system's origin and size
 *  - This implementation sets chOff == bbox.xy and chExt == bbox.cxcy, i.e. the child
 *    coordinate system is 1:1 with the slide's → child elements can reuse their
 *    original slide coordinates inside the group with no transform
 *  - childrenXml: concatenation of each child's raw XML fragment (passthrough
 *    children keep their original bytes)
 */
export function buildGrpSpXml(slide: Slide, bbox: EmuRect, childrenXml: string): string {
  const id = nextCNvPrId(slide)
  const name = `Group ${id}`
  const { x, y, cx, cy } = bbox
  const grpXfrm =
    `<a:xfrm>` +
    `<a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/>` +
    `<a:chOff x="${x}" y="${y}"/><a:chExt cx="${cx}" cy="${cy}"/>` +
    `</a:xfrm>`
  return (
    `<p:grpSp>` +
    `<p:nvGrpSpPr><p:cNvPr id="${id}" name="${escapeXmlAttr(name)}"/>` +
    `<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
    `<p:grpSpPr>${grpXfrm}</p:grpSpPr>` +
    childrenXml +
    `</p:grpSp>`
  )
}
