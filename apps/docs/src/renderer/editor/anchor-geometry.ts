/**
 * Anchor geometry: where an anchored object (wp:anchor picture or shape,
 * w:tblpPr floating table, header/footer float) sits relative to the frame it
 * is positioned against. One resolver for the X/Y origin maths that the run
 * picture, picture block, floating table, nested float, header strip and the
 * paginator's page-relative shift used to each derive on their own.
 *
 * Coordinates are CSS px. `resolveAnchor` returns offsets measured from the
 * named origin band; when the frame carries the band geometry the absolute
 * page position is filled in too, otherwise the renderer expresses the band
 * through CSS custom properties (`--doc-margin-left`, `--tblp-dy`, ...).
 *
 * Word truths (word-truths.md): relativeFrom="column" measures from the text
 * column, not the paragraph's indented box (L400); a margin-relative posOffset
 * counts from the margin top (L81); a quarter-turned picture is placed by its
 * unturned extent box and wraps by the turned bounding box (L93).
 */
import type { HfImage, HfTextBox, TableModel, TextboxDisplay } from '@genoffice/docx-engine'
import { quarterTurnInsetPx } from './image-rotation'

export const EMU_PER_PX = 9525
export const TWIPS_PER_PX = 15

export type AnchorRelH =
  | 'page'
  | 'margin'
  | 'column'
  | 'character'
  | 'leftMargin'
  | 'rightMargin'
  | 'insideMargin'
  | 'outsideMargin'
export type AnchorRelV =
  | 'page'
  | 'margin'
  | 'paragraph'
  | 'line'
  | 'topMargin'
  | 'bottomMargin'
  | 'insideMargin'
  | 'outsideMargin'
/** wp:align and w:tblpXSpec / w:tblpYSpec keywords share this enum */
export type AnchorSpec = 'left' | 'center' | 'right' | 'inside' | 'outside' | 'top' | 'bottom'
export type AnchorXAlign = 'left' | 'center' | 'right'
export type AnchorYAlign = 'top' | 'center' | 'bottom'
export type AnchorWrap = 'none' | 'behind' | 'square' | 'tight' | 'through' | 'topBottom'
export type AnchorXOrigin = 'page' | 'margin' | 'column' | 'leftMargin' | 'rightMargin'
export type AnchorYOrigin = 'page' | 'margin' | 'paragraph' | 'line' | 'topMargin' | 'bottomMargin'

export interface AnchorInput {
  /** wp:positionH relativeFrom / w:horzAnchor (text = column); null = column */
  relH?: AnchorRelH | 'text' | null
  /** wp:positionV relativeFrom / w:vertAnchor (text = paragraph); null = paragraph */
  relV?: AnchorRelV | 'text' | null
  /** wp:posOffset / w:tblpX (px); null = alignment keyword or none */
  xPx?: number | null
  yPx?: number | null
  /** wp:align / w:tblpXSpec */
  xSpec?: string | null
  /** wp:align / w:tblpYSpec */
  ySpec?: string | null
  widthPx?: number | null
  heightPx?: number | null
  rotDeg?: number | null
  wrap?: AnchorWrap | null
  /** text-wrap side for square / tight / through: the side the object hugs */
  side?: 'left' | 'right' | null
}

export interface AnchorFrame {
  page?: { w: number; h: number }
  margins?: { left: number; right: number; top: number; bottom: number }
  /** raw w:pgMar top/bottom when `margins` carries strip-pushed values: the topMargin / bottomMargin bands */
  sectMargins?: { top: number; bottom: number }
  /** text column: page x of its left edge and width (a cell's text column in cells) */
  column?: { x: number; w: number }
  paragraph?: { top?: number; indentLeft?: number; indentFirst?: number; lineTop?: number }
  cell?: { x: number; w: number }
  hfStrip?: { top: number; kind: 'header' | 'footer' }
}

export interface AnchorCtx {
  kind: 'body' | 'cell' | 'textbox' | 'hf' | 'nested'
  surface: 'editor' | 'static' | 'measure'
}

export interface ResolvedAnchor {
  xOrigin: AnchorXOrigin
  yOrigin: AnchorYOrigin
  /** posOffset from the origin band's left edge; null = alignment keyword or no position */
  x: number | null
  /** posOffset from the origin band's top edge; null = alignment keyword or no position */
  y: number | null
  xAlign: AnchorXAlign | null
  yAlign: AnchorYAlign | null
  /** true when the X band is absolute on the page (page / margin bands) */
  pageRelX: boolean
  /** true when the Y band is absolute on the page (page / margin bands) */
  pageRelY: boolean
  /** absolute page coordinates when the frame carries the band; alignment resolved against the band */
  pageX?: number
  pageY?: number
  /** extent box and the quarter-turn inset the turned bounding box adds on each side */
  box: { w: number; h: number; inset: number }
  wrap: AnchorWrap
  side: 'left' | 'right' | null
  /** flow-exclusion rectangle in page coordinates (square / tight / through / topBottom with a frame) */
  obstacle?: { left: number; top: number; right: number; bottom: number }
  /** px a page-relative Y lifts the object above the anchor line (negative), when the frame resolves it */
  lift?: number
}

const X_ALIGN: Record<string, AnchorXAlign> = {
  left: 'left',
  center: 'center',
  right: 'right',
  inside: 'left',
  outside: 'right',
}
const Y_ALIGN: Record<string, AnchorYAlign> = {
  top: 'top',
  center: 'center',
  bottom: 'bottom',
  inside: 'top',
  outside: 'bottom',
}

/** w:tblpYSpec / wp:align keyword as a vertical alignment (inside = top, outside = bottom on odd pages) */
export function yAlignOf(spec: string | null | undefined): AnchorYAlign | null {
  return spec ? (Y_ALIGN[spec] ?? null) : null
}

/** w:tblpXSpec / wp:align keyword as a horizontal alignment (inside = left, outside = right on odd pages) */
export function xAlignOf(spec: string | null | undefined): AnchorXAlign | null {
  return spec ? (X_ALIGN[spec] ?? null) : null
}

/** offset of a `size` object aligned inside a `span` band (top/left = 0, center, bottom/right) */
export function alignedOffset(
  align: AnchorXAlign | AnchorYAlign | null,
  span: number,
  size: number,
): number {
  if (align === 'center') return (span - size) / 2
  if (align === 'bottom' || align === 'right') return span - size
  return 0
}

export function xOriginOf(rel: AnchorInput['relH']): AnchorXOrigin {
  switch (rel) {
    case 'page':
    case 'margin':
    case 'leftMargin':
    case 'rightMargin':
      return rel
    case 'insideMargin':
      return 'leftMargin'
    case 'outsideMargin':
      return 'rightMargin'
    default:
      return 'column'
  }
}

export function yOriginOf(rel: AnchorInput['relV']): AnchorYOrigin {
  switch (rel) {
    case 'page':
    case 'margin':
    case 'line':
    case 'topMargin':
    case 'bottomMargin':
      return rel
    case 'insideMargin':
      return 'topMargin'
    case 'outsideMargin':
      return 'bottomMargin'
    default:
      return 'paragraph'
  }
}

/** [left, right] page x of the band a horizontal origin measures in; null when the frame lacks it */
export function anchorXBand(origin: AnchorXOrigin, frame: AnchorFrame): [number, number] | null {
  const { page, margins, column } = frame
  switch (origin) {
    case 'page':
      return page ? [0, page.w] : null
    case 'margin':
      if (page && margins) return [margins.left, page.w - margins.right]
      return column ? [column.x, column.x + column.w] : null
    case 'leftMargin':
      return margins ? [0, margins.left] : null
    case 'rightMargin':
      return page && margins ? [page.w - margins.right, page.w] : null
    default:
      return column ? [column.x, column.x + column.w] : null
  }
}

/** [top, bottom] page y of the band a vertical origin measures in; null when the frame lacks it */
export function anchorYBand(origin: AnchorYOrigin, frame: AnchorFrame): [number, number] | null {
  const { page, margins, paragraph } = frame
  const sect = frame.sectMargins ?? margins
  switch (origin) {
    case 'page':
      return [0, page?.h ?? Number.POSITIVE_INFINITY]
    case 'margin':
      if (!margins) return null
      return [margins.top, page ? page.h - margins.bottom : Number.POSITIVE_INFINITY]
    case 'topMargin':
      return sect ? [0, sect.top] : null
    case 'bottomMargin':
      return page && sect ? [page.h - sect.bottom, page.h] : null
    case 'line':
      return paragraph?.lineTop !== undefined ? [paragraph.lineTop, paragraph.lineTop] : null
    default:
      return paragraph?.top !== undefined ? [paragraph.top, paragraph.top] : null
  }
}

/**
 * Resolve an anchor against its frame. Offsets stay band-relative; absolute
 * page coordinates and the obstacle rect are filled in when the frame carries
 * the band the anchor measures in.
 */
export function resolveAnchor(
  anchor: AnchorInput,
  frame: AnchorFrame,
  _ctx: AnchorCtx,
): ResolvedAnchor {
  const xOrigin = xOriginOf(anchor.relH)
  const yOrigin = yOriginOf(anchor.relV)
  const w = Math.max(0, Number(anchor.widthPx) || 0)
  const h = Math.max(0, Number(anchor.heightPx) || 0)
  const inset = quarterTurnInsetPx(w, h, anchor.rotDeg)
  const xAlign = anchor.xPx == null ? xAlignOf(anchor.xSpec) : null
  const yAlign = anchor.yPx == null ? yAlignOf(anchor.ySpec) : null
  const out: ResolvedAnchor = {
    xOrigin,
    yOrigin,
    x: anchor.xPx ?? null,
    y: anchor.yPx ?? null,
    xAlign,
    yAlign,
    pageRelX: xOrigin !== 'column',
    pageRelY:
      yOrigin === 'page' ||
      yOrigin === 'margin' ||
      yOrigin === 'topMargin' ||
      yOrigin === 'bottomMargin',
    box: { w, h, inset },
    wrap: anchor.wrap ?? 'none',
    side: anchor.side ?? null,
  }
  const bx = anchorXBand(xOrigin, frame)
  if (bx) out.pageX = bx[0] + (out.x ?? alignedOffset(xAlign, bx[1] - bx[0], w))
  const by = anchorYBand(yOrigin, frame)
  if (by && (out.y != null || Number.isFinite(by[1]))) {
    out.pageY = by[0] + (out.y ?? alignedOffset(yAlign, by[1] - by[0], h))
  }
  if (
    out.pageX !== undefined &&
    out.pageY !== undefined &&
    out.wrap !== 'none' &&
    out.wrap !== 'behind'
  ) {
    // the turned bounding box excludes text (L93): same centre, swapped sides
    out.obstacle = {
      left: out.pageX + inset,
      top: out.pageY - inset,
      right: out.pageX + w - inset,
      bottom: out.pageY + h + inset,
    }
  }
  if (out.pageRelY && out.pageY !== undefined && frame.paragraph?.top !== undefined) {
    out.lift = Math.min(0, out.pageY - frame.paragraph.top)
  }
  return out
}

const SIDE_WRAP_RE = /^(square|tight|through)-(left|right)$/

/** anchor of a run-level picture (docInlineImage attrs) */
export function runPictureAnchor(attrs: Record<string, unknown>): AnchorInput {
  const wrap = attrs.wrap as string | null
  const side = SIDE_WRAP_RE.exec(wrap ?? '')
  return {
    relH: (attrs.relH as AnchorRelH | null) ?? null,
    relV: (attrs.relV as AnchorRelV | null) ?? null,
    xPx: attrs.offsetXEmu != null ? Number(attrs.offsetXEmu) / EMU_PER_PX : null,
    yPx: attrs.offsetYEmu != null ? Number(attrs.offsetYEmu) / EMU_PER_PX : null,
    xSpec: (attrs.posH as string | null) ?? null,
    ySpec: (attrs.posV as string | null) ?? null,
    widthPx: Number(attrs.widthPx) || 0,
    heightPx: Number(attrs.heightPx) || 0,
    rotDeg: attrs.rotDeg as number | null,
    wrap:
      wrap === 'front'
        ? 'none'
        : wrap === 'behind' || wrap === 'topBottom'
          ? wrap
          : side
            ? (side[1] as AnchorWrap)
            : null,
    side: side ? (side[2] as 'left' | 'right') : null,
  }
}

/** anchor of a picture block (docProtected image attrs) */
export function blockPictureAnchor(attrs: Record<string, unknown>): AnchorInput {
  const wrap = attrs.imageWrap as string | null
  const side = SIDE_WRAP_RE.exec(wrap ?? '')
  return {
    relH: (attrs.imageRelH as AnchorRelH | null) ?? null,
    relV: (attrs.imageRelV as AnchorRelV | null) ?? null,
    xPx: attrs.imageOffsetXEmu != null ? Number(attrs.imageOffsetXEmu) / EMU_PER_PX : null,
    yPx: attrs.imageOffsetYEmu != null ? Number(attrs.imageOffsetYEmu) / EMU_PER_PX : null,
    xSpec: (attrs.imagePosH as string | null) ?? null,
    ySpec: (attrs.imagePosV as string | null) ?? null,
    widthPx: Number(attrs.imageWidthPx) || 0,
    heightPx: Number(attrs.imageHeightPx) || 0,
    rotDeg: attrs.imageRotDeg as number | null,
    wrap:
      wrap === 'front'
        ? 'none'
        : wrap === 'behind' || wrap === 'topBottom'
          ? wrap
          : side
            ? (side[1] as AnchorWrap)
            : null,
    side: side ? (side[2] as 'left' | 'right') : null,
  }
}

/** anchor of a w:tblpPr floating table (docTable attrs or a TableModel.floatPos) */
export function tableFloatAnchor(
  pos:
    | NonNullable<TableModel['floatPos']>
    | {
        tblFloatXTwips?: unknown
        tblFloatYTwips?: unknown
        tblFloatHorzAnchor?: unknown
        tblFloatVertAnchor?: unknown
        tblFloatXSpec?: unknown
        tblFloatYSpec?: unknown
      },
  side: 'left' | 'right' | null,
  widthPx = 0,
): AnchorInput {
  const p = pos as Record<string, unknown>
  const model = 'xTwips' in p || 'yTwips' in p
  const xT = model ? p.xTwips : p.tblFloatXTwips
  const yT = model ? p.yTwips : p.tblFloatYTwips
  const xSpec = (model ? p.xSpec : p.tblFloatXSpec) as string | null | undefined
  const ySpec = (model ? p.ySpec : p.tblFloatYSpec) as string | null | undefined
  return {
    relH: ((model ? p.horzAnchor : p.tblFloatHorzAnchor) as AnchorInput['relH']) ?? null,
    relV: ((model ? p.vertAnchor : p.tblFloatVertAnchor) as AnchorInput['relV']) ?? null,
    xPx: xSpec ? null : xT != null ? (Number(xT) || 0) / TWIPS_PER_PX : null,
    yPx: ySpec ? null : yT != null ? (Number(yT) || 0) / TWIPS_PER_PX : null,
    xSpec: xSpec ?? null,
    ySpec: ySpec ?? null,
    widthPx,
    wrap: 'square',
    side,
  }
}

/** anchor of a header/footer floating image or text box */
export function hfFloatAnchor(img: HfImage | HfTextBox): AnchorInput {
  return {
    relH: img.posHRel ?? null,
    relV: img.posVRel ?? null,
    xPx: img.posXPx ?? null,
    yPx: img.posYPx ?? null,
    xSpec: img.posH ?? null,
    ySpec: img.posV ?? null,
    widthPx: img.widthPx ?? 0,
    heightPx: img.heightPx ?? 0,
    wrap:
      'behind' in img && img.behind
        ? 'behind'
        : (((img as HfImage).wrap as AnchorWrap | undefined) ?? 'none'),
  }
}

/** anchor of a floating text box / shape wrapper (offsets already resolved to the column by parse) */
export function textboxAnchor(box: TextboxDisplay): AnchorInput {
  return {
    relH: box.pageRelX ? (box.pageRelXFrom ?? 'margin') : null,
    relV: box.pageRelV ? (box.pageRelVFrom ?? 'margin') : null,
    xPx: box.offsetXEmu != null ? box.offsetXEmu / EMU_PER_PX : null,
    yPx: box.offsetYEmu != null ? box.offsetYEmu / EMU_PER_PX : null,
    widthPx: box.widthPx ?? 0,
    heightPx: box.heightPx ?? 0,
    rotDeg: box.rotDeg ?? null,
    wrap: box.behind
      ? 'behind'
      : box.wrapSides
        ? 'square'
        : box.bandTopPx !== undefined
          ? 'topBottom'
          : 'none',
  }
}

// ---- floating tables (w:tblpPr) ----

export interface TableFloatInput {
  anchor: AnchorInput
  side: 'left' | 'right'
  /** fitted table box width (px); 0 when unknown */
  widthPx: number
  /** w:*FromText gaps (px, >= 0) */
  distance: { top: number; right: number; bottom: number; left: number }
  /** w:tblpX present (a keyword X leaves it unused) */
  hasX: boolean
  /** numeric w:tblpX (px) even under a keyword; 0 when absent */
  xPx: number
  /** editor only: the pct width keeps the stylesheet's centring, and `centerMargin`
   *  is the width module's centred-box margin when it already resolved one */
  pctWidth?: boolean
  centerMargin?: string | null
  /** legacy cell-margin hang of a margin-anchored float (negative tblInd) */
  indentTwips?: number
}

export interface TableFloatPlacement {
  geom: ResolvedAnchor
  /** Y is a page/margin position resolved by the paginator (--tblp-dy), not a flow offset */
  pageRelV: boolean
  styles: string[]
  data: Record<string, string>
}

export function tableFloatDistancePx(
  distance: Partial<Record<'top' | 'right' | 'bottom' | 'left', number>> | null | undefined,
): TableFloatInput['distance'] {
  const px = (twips: unknown) => Math.max(0, (Number(twips) || 0) / TWIPS_PER_PX)
  return {
    top: px(distance?.top),
    right: px(distance?.right),
    bottom: px(distance?.bottom),
    left: px(distance?.left),
  }
}

const fx = (n: number) => n.toFixed(1)

/**
 * Placement of a floating table in its surface's CSS: the editor float (body)
 * measures X/Y against the section bands through custom properties; a float
 * nested in a cell measures against the cell.
 */
export function tableFloatCss(input: TableFloatInput, ctx: AnchorCtx): TableFloatPlacement {
  const geom = resolveAnchor(input.anchor, {}, ctx)
  const { side, widthPx, distance } = input
  const y = geom.y ?? 0
  const styles: string[] = []
  const data: Record<string, string> = {}
  if (ctx.kind === 'nested') {
    const x = geom.x ?? 0
    // cell-relative: tblpY below the cell content top, a left X from the cell edge
    styles.push(`float:${side}`)
    const top = Math.max(0, y)
    if (top > 0) {
      styles.push(`margin-top:${fx(top)}px`)
      // lines of the anchor paragraph fill the tblpY band above the table
      styles.push(`shape-outside:inset(${fx(top)}px 0 0 0)`)
    }
    // every band is the cell (page and margin alike): a numeric X offsets from
    // the cell edge, a keyword aligns in it (right = the float's own side)
    if (geom.xAlign === 'center' && widthPx > 0)
      styles.push(`margin-left:calc((100% - ${fx(widthPx)}px) / 2)`)
    else if (!geom.xAlign && x > 0 && side === 'left') styles.push(`margin-left:${fx(x)}px`)
    const textSide = side === 'left' ? 'right' : 'left'
    if (distance[textSide]) styles.push(`margin-${textSide}:${fx(distance[textSide])}px`)
    if (distance.bottom) styles.push(`margin-bottom:${fx(distance.bottom)}px`)
    return { geom, pageRelV: false, styles, data }
  }
  // w:tblpY under vertAnchor="page"/"margin" is a position on the landing page,
  // not a flow offset (a missing vertAnchor is text-relative); the paginator
  // resolves the target and applies it through --tblp-dy
  const vSpec = geom.yAlign
  const pageRelV = geom.pageRelY && (geom.y != null || vSpec !== null)
  if (pageRelV) {
    data['data-tblp-vy'] = fx(vSpec ? 0 : y)
    data['data-tblp-vanchor'] = geom.yOrigin
    if (vSpec) data['data-tblp-vspec'] = vSpec
    styles.push('margin-top:var(--tblp-dy,0px)')
    // the band above the shifted table hosts the anchor paragraph's lines
    styles.push('shape-outside:inset(var(--tblp-dy,0px) 0 0 0)')
  }
  const top = pageRelV ? 0 : y + distance.top
  if (top) styles.push(`margin-top:${fx(top)}px`)
  // Word lays the anchor paragraph's lines in the tblpY band above the table at
  // full width; a band under one 12pt line hosts nothing and keeps the plain margin
  if (!pageRelV && y >= 16) styles.push(`shape-outside:inset(${fx(y)}px 0 0 0)`)
  if (distance.bottom) styles.push(`margin-bottom:${fx(distance.bottom)}px`)
  // horzAnchor="page" measures from the paper edge: subtract the left margin
  const fromPageEdge = geom.xOrigin === 'page'
  const xSpec = input.anchor.xSpec ?? null
  const x = input.xPx
  const xExpr = fromPageEdge ? `calc(${fx(x)}px - var(--doc-margin-left,0px))` : `${fx(x)}px`
  if (side === 'left') {
    if (geom.xAlign === 'left' && fromPageEdge) {
      // left / inside keyword on the page band: the box starts at the paper edge
      styles.push('margin-left:calc(0px - var(--doc-margin-left,0px))')
    } else if (geom.xAlign === 'center' && fromPageEdge && widthPx > 0) {
      // page-band centre: the paper's middle, not the content box's
      styles.push(
        `margin-left:calc((var(--doc-content-w,100%) + var(--doc-margin-left,0px) + var(--doc-margin-right,0px) - ${fx(widthPx)}px) / 2 - var(--doc-margin-left,0px))`,
      )
    } else if (xSpec === 'center' && input.centerMargin) {
      styles.push(input.centerMargin)
    } else if (xSpec === 'center' && widthPx > 0 && !input.pctWidth) {
      styles.push(`margin-left:max(0px,calc((var(--doc-content-w,100%) - ${fx(widthPx)}px) / 2))`)
    } else if (xSpec === null && (x || fromPageEdge)) {
      // capped at the right content edge; the cap never pulls a too-wide table
      // left of the content box
      const capped =
        widthPx > 0
          ? `min(${xExpr},max(0px,calc(var(--doc-content-w,100%) - ${fx(widthPx)}px)))`
          : xExpr
      // a page-anchored X may reach into the left margin; margin anchors stop at the content edge
      const floor = fromPageEdge ? 'calc(0px - var(--doc-margin-left,0px))' : '0px'
      styles.push(`margin-left:max(${floor},${capped})`)
    } else if (xSpec === null && geom.xOrigin === 'margin' && (input.indentTwips ?? 0) < 0) {
      styles.push(`margin-left:${fx((input.indentTwips ?? 0) / TWIPS_PER_PX)}px`)
    }
    if (distance.right) styles.push(`margin-right:${fx(distance.right)}px`)
  } else {
    if (distance.left) styles.push(`margin-left:${fx(distance.left)}px`)
    if (geom.xAlign === 'right') {
      // right / outside keyword: the box ends at the band's right edge (the
      // stylesheet float already hugs the column; the page band overhangs the margin)
      if (fromPageEdge) styles.push('margin-right:calc(0px - var(--doc-margin-right,0px))')
    } else if (input.hasX && widthPx > 0) {
      styles.push(
        `margin-right:max(0px,calc(var(--doc-content-w,100%) - ${xExpr} - ${fx(widthPx)}px))`,
      )
    }
  }
  return { geom, pageRelV, styles, data }
}

/** left margin of a float flowed inline (no side room) that keeps its w:tblpX position */
export function tableFloatInlineLeftCss(attrs: Record<string, unknown>): string | null {
  if (!attrs.tblFloatSuppressed || attrs.tblFloatXSpec || attrs.tblFloatXTwips == null) return null
  const x = (Number(attrs.tblFloatXTwips) || 0) / TWIPS_PER_PX
  // margin / text (the default when w:horzAnchor is omitted) both start at the column edge
  if (attrs.tblFloatHorzAnchor !== 'page') return x > 0 ? `${fx(x)}px` : null
  return `calc(${fx(x)}px - var(--doc-margin-left,0px))`
}

/**
 * Side room (twips) a text-anchored float leaves beside itself in a column of
 * `fitTwips`, and whether the float starts at the column's left edge (within
 * ~1in of the paper edge for page anchors); null when the column is unknown.
 */
export function tableFloatSideRoom(
  pos: NonNullable<TableModel['floatPos']> | undefined,
  gridTwips: number,
  fitTwips: number | null,
): { roomTwips: number; leftEdge: boolean } | null {
  if (fitTwips == null) return null
  const floatX = pos?.xTwips ?? 0
  const span =
    (pos?.horzAnchor === 'page' ? 0 : floatX) + gridTwips + (pos?.distanceTwips?.right ?? 0)
  // a right / outside keyword hugs the far edge: the room is on the left (not tested here)
  const rightKeyword = pos?.xSpec === 'right' || pos?.xSpec === 'outside'
  return {
    roomTwips: fitTwips - span,
    leftEdge: !rightKeyword && (pos?.horzAnchor === 'page' ? floatX <= 1440 : floatX <= 720),
  }
}

// ---- paginator-facing float facts (the data-tblp-* twins, model side) ----

export interface PageRelVTarget {
  /** w:tblpY / wp:posOffset on the band (px); 0 under a keyword */
  yPx: number
  /** the paginator shifts page / margin targets; the margin-edge bands are emitted, not shifted */
  band: 'page' | 'margin' | 'topMargin' | 'bottomMargin'
  spec: AnchorYAlign | null
}

/** page-relative Y the paginator resolves into the float shift; null for a flow offset */
export function pageRelVTarget(geom: ResolvedAnchor): PageRelVTarget | null {
  if (!geom.pageRelY || (geom.y == null && geom.yAlign === null)) return null
  return {
    yPx: geom.y ?? 0,
    band: geom.yOrigin as PageRelVTarget['band'],
    spec: geom.yAlign,
  }
}

export interface TableFloatFacts {
  /** CSS-floated (w:tblpPr kept, not suppressed to inline flow) */
  floated: boolean
  pageRelV: PageRelVTarget | null
  /** suppressed text-anchored float with a negative w:tblpY: hangs this far above its anchor */
  liftPx: number
  /** text-anchored float that turns the page whole when it does not fit (probe 2026-10-13) */
  pushWhole: boolean
}

const round1 = (n: number) => Math.round(n * 10) / 10

/** what the paginator needs to know about a docTable's w:tblpPr, from its PM attrs */
export function tableFloatFacts(attrs: Record<string, unknown>, ctx: AnchorCtx): TableFloatFacts {
  const side = attrs.tblFloat === 'left' || attrs.tblFloat === 'right' ? attrs.tblFloat : null
  const floated = !attrs.tblFloatSuppressed && side !== null
  const vertAnchor = (attrs.tblFloatVertAnchor as string | null | undefined) ?? 'text'
  const yTwips = Number(attrs.tblFloatYTwips) || 0
  const textY = vertAnchor === 'text' ? yTwips : 0
  const distance = attrs.tblFloatDistance as Record<string, number> | null
  let pageRelV: PageRelVTarget | null = null
  if (floated) {
    const widthPx = Number(attrs.widthPx) || Number(attrs.tblFloatWidthPx) || 0
    pageRelV = pageRelVTarget(resolveAnchor(tableFloatAnchor(attrs, side, widthPx), {}, ctx))
  }
  return {
    floated,
    pageRelV,
    liftPx: attrs.tblFloatSuppressed && textY < 0 ? round1(-textY / TWIPS_PER_PX) : 0,
    pushWhole:
      !!attrs.tblFloatSource &&
      vertAnchor === 'text' &&
      yTwips >= 0 &&
      (Number(distance?.top) || 0) > 0,
  }
}

export interface PictureBlockFacts {
  /** in-flow wrap (square/tight/through/topBottom) placed by the paginator's float shift */
  pageRelV: PageRelVTarget | null
  /** the side-wrap band wrapper (doc-img-float img-wrap-band), not a CSS float */
  banded: boolean
}

/** what the paginator needs to know about a docProtected picture block, from its PM attrs */
export function pictureBlockFacts(
  attrs: Record<string, unknown>,
  ctx: AnchorCtx,
): PictureBlockFacts {
  const wrap = attrs.imageWrap as string | null
  const banded =
    attrs.imageBand === true &&
    /-(?:left|right)$/.test(String(wrap ?? '')) &&
    attrs.imageOffsetXEmu != null &&
    (Number(attrs.imageHeightPx) || 0) > 0
  const inFlowWrap = !!wrap && wrap !== 'front' && wrap !== 'behind' && !banded
  const pageRelV = inFlowWrap
    ? pageRelVTarget(resolveAnchor(blockPictureAnchor(attrs), {}, ctx))
    : null
  return { pageRelV, banded }
}
