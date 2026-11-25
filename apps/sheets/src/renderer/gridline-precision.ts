import { BooleanNumber } from '@univerjs/core'
import {
  FIX_ONE_PIXEL_BLUR_OFFSET,
  getColor,
  Spreadsheet,
  type SpreadsheetSkeleton,
  type UniverRenderingContext2D,
} from '@univerjs/engine-render'

const SPARSE_ROW_SPAN = 4_096
const EXTRA_BOUND = 0.4
const MAX_EXTRA_ROWS = 64
const GRIDLINE_PRECISION_LIMIT = 2 ** 23

interface GridlineRenderer {
  getSkeleton(): SpreadsheetSkeleton | undefined
  forceDisableGridlines: boolean
  isPrinting: boolean
  _drawAuxiliary(ctx: UniverRenderingContext2D): void
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

function* distinctBoundaries(
  accumulation: readonly number[],
  start: number,
  end: number,
): Iterable<number> {
  for (const row of boundaryRows(accumulation, start, end)) yield accumulation[row]!
}

let installed = false

/** Draw distant gridlines in viewport-local coordinates to preserve one-pixel strokes. */
export function installGridlinePrecision(): void {
  if (installed) return
  const prototype = Spreadsheet.prototype as unknown as GridlineRenderer
  const original = prototype._drawAuxiliary
  if (!original || !prototype._drawGapAreas) return
  installed = true
  prototype._drawAuxiliary = function (ctx) {
    const skeleton = this.getSkeleton()
    const segment = skeleton?.rowColumnSegment
    const rows = skeleton?.rowHeightAccumulation
    const columns = skeleton?.columnWidthAccumulation
    if (!skeleton || !segment || !rows?.length || !columns?.length || this.isPrinting)
      return original.call(this, ctx)
    if (skeleton.showGridlines === BooleanNumber.FALSE || this.forceDisableGridlines) return

    const rowCount = segment.endRow - segment.startRow + 1
    const sparse = rowCount > SPARSE_ROW_SPAN
    const rowPadding = Math.ceil(rowCount * EXTRA_BOUND)
    const extraRows = sparse ? Math.min(MAX_EXTRA_ROWS, rowPadding) : rowPadding
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
    const { a, d } = ctx.getTransform()
    const scaleX = Math.abs(a) || 1
    const scaleY = Math.abs(d) || 1
    if (
      Math.max(endY * Math.max(1, scaleY), endX * Math.max(1, scaleX)) < GRIDLINE_PRECISION_LIMIT
    ) {
      return original.call(this, ctx)
    }
    // The canvas rasterizer uses float32. At 2^23 it loses the half-device-
    // pixel offset of a one-pixel stroke if we send sheet coordinates.
    const originX = Math.round(startX * scaleX) / scaleX
    const originY = Math.round(startY * scaleY) / scaleY

    ctx.save()
    try {
      ctx.setLineWidthByPrecision(1)
      ctx.strokeStyle =
        skeleton.gridlinesColor ?? ctx.renderConfig.gridlinesColor ?? getColor([214, 216, 219])
      ctx.translate(originX, originY)
      ctx.translateWithPrecisionRatio(FIX_ONE_PIXEL_BLUR_OFFSET, FIX_ONE_PIXEL_BLUR_OFFSET)
      ctx.beginPath()
      const horizontal = (y: number) => {
        ctx.moveToByPrecision(startX - originX, y - originY)
        ctx.lineToByPrecision(endX - originX, y - originY)
      }
      const vertical = (x: number) => {
        ctx.moveToByPrecision(x - originX, startY - originY)
        ctx.lineToByPrecision(x - originX, endY - originY)
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
        .filter((range) => rows[range.endRow]! > (rows[range.startRow - 1] ?? 0))
      for (const range of [...merged, ...skeleton.overflowCache.toNativeArray()]) {
        const x = (columns[range.startColumn - 1] ?? 0) - originX
        const y = (rows[range.startRow - 1] ?? 0) - originY
        const width = (columns[range.endColumn] ?? columns.at(-1)!) - originX - x
        const height = (rows[range.endRow] ?? rows.at(-1)!) - originY - y
        ctx.clearRectByPrecision(x, y, width, height)
        ctx.beginPath()
        ctx.rectByPrecision(x, y, width, height)
        ctx.stroke()
      }
      ctx.translate(-originX, -originY)
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
