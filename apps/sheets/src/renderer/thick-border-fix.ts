/**
 * Excel draws cell borders in three weights: thin 1px, medium 2px, thick
 * 3px (at 100% zoom, 96 DPI), and a thick edge is visibly heavier than a
 * medium one. Univer's border extension strokes them at 1 / 2 / 3 device
 * pixels — but it also shifts the canvas by half a device pixel so 1px
 * strokes land crisp (FIX_ONE_PIXEL_BLUR_OFFSET). The even-width medium
 * stroke is therefore centred on a pixel boundary and smears over three
 * pixel rows at half intensity — the very same three rows the 3px thick
 * stroke fills. On screen the two weights have an identical footprint and
 * differ only in darkness, so thick never reads as thicker than medium.
 *
 * The fix routes THICK edges through our own stroke: 4 device pixels,
 * pixel-aligned (the half-pixel shift is cancelled for the even width), which
 * is the smallest weight whose footprint is strictly wider than the medium
 * smear on every device pixel ratio, and it keeps Univer's device-pixel
 * convention (constant across zoom, exactly like thin, medium and the
 * gridlines), so the three weights stay ordered at every zoom level. Thin
 * and medium edges are still drawn by the stock extension, byte-for-byte.
 *
 * Patched on the `Border` extension prototype like the other Univer
 * render patches (merge-border-fix, center-continuous): the extension is
 * instantiated per render unit by Univer itself, so a prototype patch is the
 * only hook that reaches every sheet, including print/PDF contexts.
 */
import { BorderStyleTypes } from '@univerjs/core'
import type { IRange } from '@univerjs/core'
import {
  Border,
  BORDER_TYPE,
  drawLineByBorderType,
  setLineType,
  type SpreadsheetSkeleton,
  type UniverRenderingContext,
} from '@univerjs/engine-render'

/// Device-pixel width of a THICK border stroke (Excel: 1.5× medium; see the
/// module comment for why 3 device pixels does not read as thicker).
export const THICK_BORDER_DEVICE_PX = 4

/// Device-pixel stroke width per border style — Univer's built-in ladder
/// (thin 1, medium 2) with the wider thick. Exported for tests; the built-in
/// values are what the stock extension keeps drawing.
export function borderStrokeDevicePx(style: BorderStyleTypes): number {
  switch (style) {
    case BorderStyleTypes.MEDIUM:
    case BorderStyleTypes.MEDIUM_DASHED:
    case BorderStyleTypes.MEDIUM_DASH_DOT:
    case BorderStyleTypes.MEDIUM_DASH_DOT_DOT:
      return 2
    case BorderStyleTypes.THICK:
      return THICK_BORDER_DEVICE_PX
    default:
      return 1
  }
}

/// Univer's `setLineWidthByPrecision` divides by the canvas scale (device
/// pixel ratio × zoom), so a device-pixel width maps to this many CSS
/// pixels — the same at every zoom, matching the built-in strokes.
export function borderStrokeCssPx(style: BorderStyleTypes, scale: number): number {
  return borderStrokeDevicePx(style) / Math.max(scale, Number.EPSILON)
}

/// Straight edges only — diagonals keep the stock stroke.
const EDGE_TYPES: readonly string[] = [
  BORDER_TYPE.TOP,
  BORDER_TYPE.BOTTOM,
  BORDER_TYPE.LEFT,
  BORDER_TYPE.RIGHT,
]

interface BorderCacheItemLike {
  readonly type: string
  readonly style: BorderStyleTypes
  readonly color: string
}

type BorderCacheLike = Record<string, BorderCacheItemLike | Record<string, never>>

interface RenderBorderContextLike {
  readonly ctx: UniverRenderingContext
  readonly precisionScale: number
  readonly overflowCache: unknown
  readonly spreadsheetSkeleton: SpreadsheetSkeleton
  readonly diffRanges: IRange[]
}

interface BorderProtoLike {
  renderBorderByCell(
    context: RenderBorderContextLike,
    row: number,
    col: number,
    borderCacheItem: BorderCacheLike,
  ): true | undefined
  isRenderDiffRangesByRow(startRow: number, endRow: number, diffRanges?: IRange[]): boolean
  _getOverflowExclusion(overflowCache: unknown, type: string, row: number, col: number): boolean
}

function isThickEdge(
  item: BorderCacheItemLike | Record<string, never>,
): item is BorderCacheItemLike {
  return 'style' in item && item.style === BorderStyleTypes.THICK && EDGE_TYPES.includes(item.type)
}

let installed = false

export function installThickBorderFix(): void {
  if (installed) return
  installed = true
  const proto = Border.prototype as unknown as BorderProtoLike
  const original = proto.renderBorderByCell
  if (typeof original !== 'function') return
  proto.renderBorderByCell = function (
    this: BorderProtoLike,
    context: RenderBorderContextLike,
    row: number,
    col: number,
    borderCacheItem: BorderCacheLike,
  ): true | undefined {
    const thickKeys = Object.keys(borderCacheItem).filter((key) => {
      const item = borderCacheItem[key]
      return item !== undefined && isThickEdge(item)
    })
    if (thickKeys.length === 0) return original.call(this, context, row, col, borderCacheItem)
    // Same guards as the stock method: hidden rows/columns (merged cells
    // excepted) and cells outside the dirty ranges draw nothing.
    const { ctx, spreadsheetSkeleton, precisionScale, diffRanges, overflowCache } = context
    const cellInfo = spreadsheetSkeleton.getCellWithCoordByIndex(row, col, false)
    const { isMerged, mergeInfo } = cellInfo
    if (!isMerged) {
      const worksheet = spreadsheetSkeleton.worksheet
      if (!worksheet.getRowVisible(row) || !worksheet.getColVisible(col)) return true
    }
    if (!this.isRenderDiffRangesByRow(mergeInfo.startRow, mergeInfo.endRow, diffRanges)) {
      return true
    }
    const rest: BorderCacheLike = {}
    for (const key of Object.keys(borderCacheItem)) {
      if (!thickKeys.includes(key)) rest[key] = borderCacheItem[key] as BorderCacheItemLike
    }
    if (Object.keys(rest).length > 0) original.call(this, context, row, col, rest)
    const width = THICK_BORDER_DEVICE_PX
    ctx.save()
    // Even width: cancel the extension's half-pixel shift so the stroke is
    // centred on the cell edge and fills whole pixel rows on both sides.
    ctx.translateWithPrecisionRatio(-0.5, -0.5)
    setLineType(ctx, BorderStyleTypes.THICK)
    ctx.setLineWidthByPrecision(width)
    const position = {
      startX: cellInfo.startX,
      startY: cellInfo.startY,
      endX: cellInfo.endX,
      endY: cellInfo.endY,
    }
    for (const key of thickKeys) {
      const { type, color } = borderCacheItem[key] as BorderCacheItemLike
      // Text overflowing into this cell suppresses its vertical edges.
      if (this._getOverflowExclusion(overflowCache, type, row, col)) continue
      ctx.strokeStyle = color || 'rgb(0,0,0)'
      drawLineByBorderType(ctx, type as BORDER_TYPE, (width - 1) / 2 / precisionScale, position)
    }
    ctx.restore()
    return undefined
  }
}
