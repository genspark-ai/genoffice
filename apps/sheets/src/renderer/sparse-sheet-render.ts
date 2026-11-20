import { BooleanNumber, type IRange } from '@univerjs/core'
import {
  FIX_ONE_PIXEL_BLUR_OFFSET,
  getColor,
  SHEET_EXTENSION_PREFIX,
  Spreadsheet,
  SpreadsheetSkeleton,
  type UniverRenderingContext2D,
} from '@univerjs/engine-render'

const SPARSE_ROW_SPAN = 4_096
const EXTRA_BOUND = 0.4
const MAX_EXTRA_ROWS = 64

interface GridlineRenderer {
  getSkeleton(): SpreadsheetSkeleton | undefined
  forceDisableGridlines: boolean
  isPrinting: boolean
  _drawAuxiliary(ctx: UniverRenderingContext2D): void
  _clearRectangle(
    ctx: UniverRenderingContext2D,
    rows: readonly number[],
    columns: readonly number[],
    ranges: IRange[],
  ): void
  _drawGapAreas(
    ctx: UniverRenderingContext2D,
    skeleton: SpreadsheetSkeleton,
    startRow: number,
    endRow: number,
    startColumn: number,
    endColumn: number,
    startX: number,
    endX: number,
    startY: number,
    endY: number,
  ): void
}

/// Hidden rows have the same accumulated height as their previous neighbor.
/// Jump over each run of equal positions in the sheet's physical coordinates.
function* boundaryRows(
  accumulation: readonly number[],
  start: number,
  end: number,
): Iterable<number> {
  let index = start
  while (index <= end) {
    const position = accumulation[index]!
    yield index
    index += 1
    if (accumulation[index] !== position) continue
    let low = index + 1
    let high = end + 1
    while (low < high) {
      const mid = low + Math.floor((high - low) / 2)
      if (accumulation[mid]! <= position) low = mid + 1
      else high = mid
    }
    index = low
  }
}

export function* distinctBoundaries(
  accumulation: readonly number[],
  start: number,
  end: number,
): Iterable<number> {
  for (const row of boundaryRows(accumulation, start, end)) yield accumulation[row]!
}

/// Keep the sheet's original row coordinates in every render pass so that
/// ranges, selection and merge anchors still refer to the original cells.
export function visibleCellRanges(heights: readonly number[], range: IRange): IRange[] {
  const result: IRange[] = []
  const start = Math.max(0, range.startRow)
  const end = Math.min(heights.length - 1, range.endRow)
  for (const row of boundaryRows(heights, start, end)) {
    if (heights[row] === (heights[row - 1] ?? 0)) continue
    const last = result.at(-1)
    if (last && last.endRow + 1 === row) last.endRow = row
    else result.push({ ...range, startRow: row, endRow: row })
  }
  return result
}

function mergedRangeHasHeight(heights: readonly number[], range: IRange): boolean {
  return (heights[range.endRow] ?? heights.at(-1) ?? 0) > (heights[range.startRow - 1] ?? 0)
}

type ViewportInfo = Parameters<Spreadsheet['draw']>[1]
type SpreadsheetDrawInternals = Pick<GridlineRenderer, '_drawAuxiliary'> & {
  _refreshIncrementalState: boolean
}

/// The stock draw call passes a physical bounding rectangle to extensions
/// that visit every cell. A six-row viewport may physically span a million
/// rows. Pass the visible runs to those same renderers instead.
function installSparseCellDraw(): void {
  const original = Spreadsheet.prototype.draw
  Spreadsheet.prototype.draw = function (ctx, viewport) {
    const internals = this as unknown as SpreadsheetDrawInternals
    const skeleton = this.getSkeleton()
    const heights = skeleton?.rowHeightAccumulation
    if (!skeleton || !heights || this.isPrinting) return original.call(this, ctx, viewport)
    const cacheRange = skeleton.getCacheRangeByViewport(viewport, this.isPrinting)
    if (cacheRange.endRow - cacheRange.startRow < SPARSE_ROW_SPAN) {
      return original.call(this, ctx, viewport)
    }
    internals._drawAuxiliary(ctx)
    const difference = internals._refreshIncrementalState
      ? (viewport.diffBounds?.map((bound) => skeleton.getRangeByViewBound(bound)) ?? [])
      : []
    const bounds = difference.length ? difference : [cacheRange]
    const viewRanges = bounds.flatMap((range) => visibleCellRanges(heights, range))
    if (!viewRanges.length) return
    // Some extensions fall back to the full physical rectangle when no
    // diff ranges are set. These runs are also the full repaint's dirty area.
    const parentScale = this.getParentScale()
    const scene = this.getScene()
    const overflowView = difference.length
      ? viewRanges.map((range) => ({
          ...range,
          startColumn: cacheRange.startColumn,
          endColumn: cacheRange.endColumn,
        }))
      : viewRanges
    for (const extension of this.getExtensionsByOrder()) {
      const started = performance.now()
      const info = {
        viewRanges:
          extension === this.fontExtension || extension === this.borderExtension
            ? overflowView
            : viewRanges,
        checkOutOfViewBound: true,
        viewportKey: viewport.viewportKey,
        viewBound: viewport.cacheBound,
        diffBounds: viewport.diffBounds,
      }
      extension.draw(ctx, parentScale, skeleton, viewRanges, info)
      this.addRenderFrameTimeMetricToScene(
        `${SHEET_EXTENSION_PREFIX}${extension.uKey}`,
        performance.now() - started,
        scene,
      )
    }
  }
}

interface StyleSkeletonInternals {
  _setStylesCacheForOneCell(
    row: number,
    column: number,
    options: { cacheItem: { bg: boolean; border: boolean } } | { mergeRange: IRange },
  ): void
}

/// Use the same runs while building the render cache. Each original visible
/// cell still goes through Univer's own style, font, overflow and merge logic.
function installSparseStyleCache(): void {
  const external = SpreadsheetSkeleton.prototype as unknown as StyleSkeletonInternals
  if (!external._setStylesCacheForOneCell) return
  const original = SpreadsheetSkeleton.prototype.setStylesCache
  SpreadsheetSkeleton.prototype.setStylesCache = function (viewport?: ViewportInfo) {
    const skeleton = this as unknown as StyleSkeletonInternals
    const heights = this.rowHeightAccumulation
    const columns = this.columnWidthAccumulation
    if (!viewport || !heights || !columns) return original.call(this, viewport)
    const cacheRange = this.getCacheRangeByViewport(viewport)
    if (cacheRange.endRow - cacheRange.startRow < SPARSE_ROW_SPAN)
      return original.call(this, viewport)

    this.updateVisibleRange(viewport)
    const partial =
      !viewport.isDirty &&
      !viewport.isForceDirty &&
      (viewport.diffBounds?.length ||
        viewport.diffCacheBounds?.length ||
        viewport.diffX ||
        viewport.diffY)
    const styleRanges = partial
      ? viewport.shouldCacheUpdate
        ? (viewport.diffCacheBounds?.map((bound) => this.getRangeByViewBound(bound)) ?? [])
        : []
      : [cacheRange]
    for (const range of styleRanges) {
      const { startColumn, endColumn } = range
      if (endColumn < 0 || range.endRow < 0) continue
      const expandStart = Math.max(0, startColumn - 20)
      const expandEnd = Math.min(columns.length - 1, endColumn + 20)
      for (const visible of visibleCellRanges(heights, range)) {
        for (let row = visible.startRow; row <= visible.endRow; row += 1) {
          if (!this.worksheet.getRowVisible(row)) continue
          for (let col = startColumn; col <= endColumn; col += 1) {
            skeleton._setStylesCacheForOneCell(row, col, { cacheItem: { bg: true, border: true } })
          }
          for (let col = expandStart; col < endColumn; col += 1) {
            skeleton._setStylesCacheForOneCell(row, col, {
              cacheItem: { bg: false, border: false },
            })
          }
          if (endColumn === 0) continue
          for (let col = endColumn + 1; col < expandEnd; col += 1) {
            skeleton._setStylesCacheForOneCell(row, col, {
              cacheItem: { bg: false, border: false },
            })
          }
        }
      }
      for (const mergeRange of this.getCurrentRowColumnSegmentMergeData(range)) {
        if (mergedRangeHasHeight(heights, mergeRange)) {
          skeleton._setStylesCacheForOneCell(mergeRange.startRow, mergeRange.startColumn, {
            mergeRange,
          })
        }
      }
    }
    return this
  }
}

let installed = false

/// Univer strokes a horizontal gridline for every physical row in the cache
/// range, including zero-height rows. Filtering to five rows far apart can
/// therefore queue a million identical canvas strokes. Paint a line once per
/// coordinate in that viewport; reuse Univer's merge, overflow and gap code.
export function installSparseSheetRender(): void {
  if (installed) return
  const prototype = Spreadsheet.prototype as unknown as GridlineRenderer
  const original = prototype._drawAuxiliary
  if (!original || !prototype._clearRectangle || !prototype._drawGapAreas) return
  installed = true
  installSparseCellDraw()
  installSparseStyleCache()
  prototype._drawAuxiliary = function (ctx) {
    const skeleton = this.getSkeleton()
    const segment = skeleton?.rowColumnSegment
    const rows = skeleton?.rowHeightAccumulation
    const columns = skeleton?.columnWidthAccumulation
    if (
      !skeleton ||
      !segment ||
      !rows?.length ||
      !columns?.length ||
      this.isPrinting ||
      segment.endRow - segment.startRow < SPARSE_ROW_SPAN
    ) {
      original.call(this, ctx)
      return
    }
    if (skeleton.showGridlines === BooleanNumber.FALSE || this.forceDisableGridlines) return

    const extraRows = Math.min(
      MAX_EXTRA_ROWS,
      Math.ceil((segment.endRow - segment.startRow + 1) * EXTRA_BOUND),
    )
    const extraColumns = Math.ceil((segment.endColumn - segment.startColumn + 1) * EXTRA_BOUND)
    const startRow = Math.max(0, segment.startRow - extraRows)
    const endRow = Math.min(rows.length - 1, segment.endRow + extraRows)
    const startColumn = Math.max(0, segment.startColumn - extraColumns)
    const endColumn = Math.min(columns.length - 1, segment.endColumn + extraColumns)
    if (startRow > endRow || startColumn > endColumn) return
    const startX = columns[startColumn - 1] ?? 0
    const endX = columns[endColumn]!
    const startY = rows[startRow - 1] ?? 0
    const endY = rows[endRow]!

    ctx.save()
    try {
      ctx.setLineWidthByPrecision(1)
      ctx.strokeStyle =
        skeleton.gridlinesColor ?? ctx.renderConfig.gridlinesColor ?? getColor([214, 216, 219])
      ctx.translateWithPrecisionRatio(FIX_ONE_PIXEL_BLUR_OFFSET, FIX_ONE_PIXEL_BLUR_OFFSET)
      ctx.beginPath()
      const horizontal = (y: number) => {
        ctx.moveToByPrecision(startX, y)
        ctx.lineToByPrecision(endX, y)
      }
      const vertical = (x: number) => {
        ctx.moveToByPrecision(x, startY)
        ctx.lineToByPrecision(x, endY)
      }
      horizontal(startY)
      for (const y of distinctBoundaries(rows, startRow, endRow)) {
        if (y > startY) horizontal(y)
      }
      vertical(startX)
      for (const x of distinctBoundaries(columns, startColumn, endColumn)) {
        if (x > startX) vertical(x)
      }
      ctx.stroke()

      const merged = skeleton.worksheet
        .getMergedCellRange(startRow, startColumn, endRow, endColumn)
        .filter((range) => mergedRangeHasHeight(rows, range))
      this._clearRectangle(ctx, rows, columns, merged)
      this._clearRectangle(ctx, rows, columns, skeleton.overflowCache.toNativeArray())
      this._drawGapAreas(
        ctx,
        skeleton,
        startRow,
        endRow,
        startColumn,
        endColumn,
        startX,
        endX,
        startY,
        endY,
      )
    } finally {
      ctx.restore()
    }
  }
}
