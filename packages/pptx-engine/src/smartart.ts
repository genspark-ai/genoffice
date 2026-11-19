/**
 * SmartArt (simplified) — instead of writing real diagram parts (the
 * data/layout/colors/quickStyle quartet is too heavy), generate a <p:grpSp> shape
 * group from a preset layout; it is freely editable and is just a plain shape group
 * in PowerPoint.
 *
 * Supported layouts: list (vertical list) / process (horizontal flow) / cycle /
 * hierarchy / pyramid / matrix / venn.
 */
import type { EmuRect, Slide } from './types'
import { creationIdXml, escapeXmlAttr, escapeXmlText } from './xml-utils'
import { appendRawElements, type OpenedPptx } from './index'
import { nextCNvPrId } from './insert'
import { layoutShapes, type SmartArtChildShape, type SmartArtLayout } from './smartart-layout'

export type { SmartArtLayout } from './smartart-layout'

export interface NewSmartArtOptions {
  layout: SmartArtLayout
  /** Text for each node (determines node count, 1..8) */
  items: string[]
  offset: EmuRect
}

/** Fragment for a single child shape (in child coordinate space). */
function childSpXml(id: number, s: SmartArtChildShape): string {
  const para = s.text
    ? `<a:p><a:pPr algn="ctr"/><a:r><a:rPr sz="${Math.round((s.fontSize ?? 14) * 100)}" b="1">` +
      '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr>' +
      `<a:t>${escapeXmlText(s.text)}</a:t></a:r></a:p>`
    : '<a:p/>'
  return (
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${escapeXmlAttr(`SmartShape ${id}`)}"/>` +
    '<p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
    `<p:spPr><a:xfrm><a:off x="${s.box.x}" y="${s.box.y}"/><a:ext cx="${s.box.cx}" cy="${s.box.cy}"/></a:xfrm>` +
    `<a:prstGeom prst="${s.prst}"><a:avLst/></a:prstGeom>` +
    `<a:solidFill><a:srgbClr val="${s.color}">` +
    (s.alpha != null && s.alpha < 1 ? `<a:alpha val="${Math.round(s.alpha * 100000)}"/>` : '') +
    '</a:srgbClr></a:solidFill></p:spPr>' +
    '<p:txBody><a:bodyPr wrap="square" anchor="ctr" rtlCol="0"><a:normAutofit/></a:bodyPr>' +
    `<a:lstStyle/>${para}</p:txBody></p:sp>`
  )
}

/** Build the grpSp group fragment. */
export function buildSmartArtXml(slide: Slide, opts: NewSmartArtOptions): string {
  const baseId = nextCNvPrId(slide)
  const o = opts.offset
  const shapes = layoutShapes(opts.layout, opts.items, o.cx, o.cy)
  const children = shapes.map((s, i) => childSpXml(baseId + 1 + i, s)).join('')
  return (
    `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="${baseId}" name="SmartArt ${baseId}">${creationIdXml()}</p:cNvPr>` +
    '<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
    `<p:grpSpPr><a:xfrm><a:off x="${o.x}" y="${o.y}"/><a:ext cx="${o.cx}" cy="${o.cy}"/>` +
    `<a:chOff x="0" y="0"/><a:chExt cx="${o.cx}" cy="${o.cy}"/></a:xfrm></p:grpSpPr>` +
    `${children}</p:grpSp>`
  )
}

/** Insert SmartArt (shape-group version): append + reparse, returning the new slide and element id. */
export function addSmartArt(
  opened: OpenedPptx,
  slideIndex: number,
  opts: NewSmartArtOptions,
): { slide: Slide; elementId: string } | null {
  const slide = opened.deck.slides[slideIndex]
  if (!slide || !opts.items.length) return null
  const r = appendRawElements(opened, slideIndex, [buildSmartArtXml(slide, opts)])
  return r ? { slide: r.slide, elementId: r.elementIds[r.elementIds.length - 1]! } : null
}
