/**
 * Excel `alignment horizontal="centerContinuous"` (Center Across Selection):
 * the anchor cell's text centers across the run of trailing blank cells that
 * share the format, clipped at the run's edges, without merging cells.
 * Univer has no such alignment, so the loader maps it to CENTER + OVERFLOW
 * and marks the run end on the anchor's `cell.custom`; the prototype patches
 * below widen the overflow rectangle (which is also the clip box) and the
 * centering box from the anchor cell to the whole run.
 */
import { Font, SpreadsheetSkeleton } from '@univerjs/engine-render'

export const CENTER_ACROSS_END_KEY = 'centerAcrossEnd'

interface WorksheetLike {
  getCell?(
    row: number,
    column: number,
  ): { custom?: Record<string, unknown> | null } | null | undefined
}

function centerAcrossEnd(
  worksheet: WorksheetLike | null | undefined,
  row: number,
  column: number,
): number | undefined {
  const value = worksheet?.getCell?.(row, column)?.custom?.[CENTER_ACROSS_END_KEY]
  return typeof value === 'number' && Number.isInteger(value) && value > column ? value : undefined
}

let installed = false

export function installCenterContinuousRender(): void {
  if (installed) return
  installed = true
  const skeletonProto = SpreadsheetSkeleton.prototype as any
  const origOverflow = skeletonProto.getOverflowPosition
  skeletonProto.getOverflowPosition = function (
    contentSize: unknown,
    horizontalAlign: unknown,
    row: number,
    column: number,
    columnCount: number,
  ) {
    const end = centerAcrossEnd(this.worksheet, row, column)
    // The run bounds the overflow on BOTH sides regardless of content width:
    // Excel clips at the run edge instead of spilling into neighbors.
    if (end !== undefined) {
      return { startColumn: column, endColumn: Math.min(end, columnCount - 1) }
    }
    return origOverflow.call(this, contentSize, horizontalAlign, row, column, columnCount)
  }
  const fontProto = Font.prototype as any
  // Both text paths center within renderFontCtx's cell box (Text.drawWith
  // width / documents.resize width); stretch that box to the run so the
  // centering midpoint is the run's, not the anchor cell's.
  for (const method of ['_renderText', '_renderDocuments']) {
    const orig = fontProto[method]
    fontProto[method] = function (
      ctx: unknown,
      row: number,
      col: number,
      renderFontCtx: {
        startX: number
        endX: number
        spreadsheetSkeleton?: {
          worksheet?: WorksheetLike
          getColumnCount?(): number
          getCellWithCoordByIndex?(row: number, column: number, header: boolean): { endX: number }
        }
      },
      overflowCache: unknown,
    ) {
      const skeleton = renderFontCtx?.spreadsheetSkeleton
      const end = skeleton ? centerAcrossEnd(skeleton.worksheet, row, col) : undefined
      if (end === undefined || !skeleton) {
        return orig.call(this, ctx, row, col, renderFontCtx, overflowCache)
      }
      const savedEndX = renderFontCtx.endX
      const lastColumn = Math.min(end, (skeleton.getColumnCount?.() ?? end + 1) - 1)
      const endCoord = skeleton.getCellWithCoordByIndex?.(row, lastColumn, false)
      // RTL mirroring swaps x coordinates; only widen a sane LTR box.
      if (endCoord && endCoord.endX > renderFontCtx.startX) renderFontCtx.endX = endCoord.endX
      try {
        return orig.call(this, ctx, row, col, renderFontCtx, overflowCache)
      } finally {
        renderFontCtx.endX = savedEndX
      }
    }
  }
}
