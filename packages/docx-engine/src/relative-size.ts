import type { SectionSettings } from './types'

/** what a relative width is a share of (wp14:sizeRelH relativeFrom) */
export const REL_SIZE_FROM_H = [
  'margin',
  'page',
  'leftMargin',
  'rightMargin',
  'insideMargin',
  'outsideMargin',
] as const

/** what a relative height is a share of (wp14:sizeRelV relativeFrom) */
export const REL_SIZE_FROM_V = [
  'margin',
  'page',
  'topMargin',
  'bottomMargin',
  'insideMargin',
  'outsideMargin',
] as const

export type RelSizeFrom = (typeof REL_SIZE_FROM_H)[number] | (typeof REL_SIZE_FROM_V)[number]

/** a side of a drawing sized as a share of the page: percent (40 is 40%) of `from` */
export interface RelSize {
  from: RelSizeFrom
  pct: number
}

/** wp14:sizeRelH / wp14:sizeRelV of a drawing; a side Word sizes absolutely (pct 0) is absent */
export interface DrawingRelSize {
  w?: RelSize
  h?: RelSize
}

const WP14_NS = 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing'
const UNITS_PER_PCT = 1000

const relRe = (axis: 'H' | 'V') =>
  new RegExp(
    `<wp14:sizeRel${axis}\\b[^>]*\\brelativeFrom="(\\w+)"[^>]*>\\s*` +
      `<wp14:pct(?:Width|Height)>\\s*(-?\\d+)\\s*</wp14:pct(?:Width|Height)>\\s*</wp14:sizeRel${axis}>`,
  )

function sideOf(xml: string, axis: 'H' | 'V'): RelSize | undefined {
  const m = relRe(axis).exec(xml)
  const units = m ? parseInt(m[2], 10) : 0
  const known: readonly string[] = axis === 'H' ? REL_SIZE_FROM_H : REL_SIZE_FROM_V
  if (!m || !(units > 0) || !known.includes(m[1])) return undefined
  return { from: m[1] as RelSizeFrom, pct: units / UNITS_PER_PCT }
}

/** the relative size a wp:anchor / wp:inline states; undefined when both sides are absolute */
export function drawingRelSizeOf(drawingXml: string): DrawingRelSize | undefined {
  if (!drawingXml.includes('<wp14:sizeRel')) return undefined
  const w = sideOf(drawingXml, 'H')
  const h = sideOf(drawingXml, 'V')
  return w || h ? { ...(w ? { w } : {}), ...(h ? { h } : {}) } : undefined
}

function sideXml(axis: 'H' | 'V', rel: RelSize | null, declareNs: boolean): string {
  const pct = axis === 'H' ? 'pctWidth' : 'pctHeight'
  const ns = declareNs ? ` xmlns:wp14="${WP14_NS}"` : ''
  const units = rel ? Math.round(rel.pct * UNITS_PER_PCT) : 0
  return (
    `<wp14:sizeRel${axis}${ns} relativeFrom="${rel?.from ?? 'margin'}">` +
    `<wp14:${pct}>${units}</wp14:${pct}></wp14:sizeRel${axis}>`
  )
}

/**
 * One side of a drawing made relative (`rel`) or absolute (null). An absolute
 * side keeps the frame the file names with a zero share, as Word writes it; a
 * drawing that never stated a relative size is left without one.
 */
export function setDrawingRelSize(
  drawingXml: string,
  axis: 'H' | 'V',
  rel: RelSize | null,
): string {
  const existing = new RegExp(`<wp14:sizeRel${axis}\\b[\\s\\S]*?</wp14:sizeRel${axis}>`).exec(
    drawingXml,
  )
  if (existing) {
    if (!rel) {
      return drawingXml.replace(
        existing[0],
        existing[0].replace(/(<wp14:pct(?:Width|Height)>)\s*-?\d+\s*(<)/, '$10$2'),
      )
    }
    return drawingXml.replace(
      existing[0],
      sideXml(axis, rel, /^<wp14:sizeRel[HV]\b[^>]*\bxmlns:wp14=/.test(existing[0])),
    )
  }
  if (!rel) return drawingXml
  const xml = sideXml(axis, rel, true)
  // the schema puts the pair last in the anchor, the width before the height
  if (axis === 'H' && drawingXml.includes('<wp14:sizeRelV'))
    return drawingXml.replace('<wp14:sizeRelV', `${xml}<wp14:sizeRelV`)
  const close = drawingXml.lastIndexOf('</wp:anchor>')
  return close < 0 ? drawingXml : drawingXml.slice(0, close) + xml + drawingXml.slice(close)
}

const VML_FROM: Record<RelSizeFrom, string> = {
  margin: 'margin',
  page: 'page',
  leftMargin: 'left-margin-area',
  rightMargin: 'right-margin-area',
  topMargin: 'top-margin-area',
  bottomMargin: 'bottom-margin-area',
  insideMargin: 'inner-margin-area',
  outsideMargin: 'outer-margin-area',
}

/** the same side in a VML fallback's style: mso-width-percent is in tenths of a percent */
export function setVmlRelSize(style: string, axis: 'H' | 'V', rel: RelSize | null): string {
  const side = axis === 'H' ? 'width' : 'height'
  const rest = style
    .split(';')
    .filter((decl) => !new RegExp(`^\\s*mso-${side}-(?:percent|relative)\\s*:`).test(decl))
  if (rel)
    rest.push(
      `mso-${side}-percent:${Math.round(rel.pct * 10)}`,
      `mso-${side}-relative:${VML_FROM[rel.from]}`,
    )
  return rest.join(';')
}

/** the length (twips) a relative size is a share of, on a page of the section */
export function relSizeBaseTwips(
  section: SectionSettings,
  axis: 'H' | 'V',
  from: RelSizeFrom,
): number {
  const s = section
  if (axis === 'H') {
    if (from === 'page') return s.pageWidth
    if (from === 'leftMargin' || from === 'insideMargin') return s.marginLeft
    if (from === 'rightMargin' || from === 'outsideMargin') return s.marginRight
    return s.pageWidth - s.marginLeft - s.marginRight
  }
  if (from === 'page') return s.pageHeight
  if (from === 'topMargin' || from === 'insideMargin') return s.marginTop
  if (from === 'bottomMargin' || from === 'outsideMargin') return s.marginBottom
  return s.pageHeight - s.marginTop - s.marginBottom
}

/** the size (twips) a relative side works out to in the section */
export const relSizeTwips = (section: SectionSettings, axis: 'H' | 'V', rel: RelSize): number =>
  Math.max(0, (relSizeBaseTwips(section, axis, rel.from) * rel.pct) / 100)
