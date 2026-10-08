/**
 * Pixel ⇄ anchor conversion for installed floating objects: the arrange
 * commands work on sheet-pixel boxes and write back twoCellAnchor markers.
 */
import type { WorkbookVisualObject } from '../shared/desktop-api'
import type { ArrangeBox } from './visual-arrange'
import { EMU_PER_PIXEL, markerFrom, walkMarker, type InstalledVisualFrame } from './WorkbookVisuals'

export interface SheetMetrics {
  readonly columnWidth: (index: number) => number
  readonly rowHeight: (index: number) => number
  readonly maxColumn: number
  readonly maxRow: number
}

interface WorksheetLike {
  getColumnWidth(index: number): number
  getRowHeight(index: number): number
  getMaxColumns(): number
  getMaxRows(): number
  getSheet(): { getConfig(): { defaultColumnWidth: number; defaultRowHeight: number } }
}

const XLSX_MAX_COLUMN = 16383
const XLSX_MAX_ROW = 1048575

/// Past the data-sized grid the sheet defaults apply (install grows the
/// grid to a committed anchor with exactly those sizes).
export function sheetMetrics(worksheet: WorksheetLike): SheetMetrics {
  const config = worksheet.getSheet().getConfig()
  const gridColumns = worksheet.getMaxColumns()
  const gridRows = worksheet.getMaxRows()
  return {
    columnWidth: (index) =>
      index < gridColumns
        ? Math.max(worksheet.getColumnWidth(index), 1)
        : Math.max(config.defaultColumnWidth, 1),
    rowHeight: (index) =>
      index < gridRows
        ? Math.max(worksheet.getRowHeight(index), 1)
        : Math.max(config.defaultRowHeight, 1),
    maxColumn: XLSX_MAX_COLUMN,
    maxRow: XLSX_MAX_ROW,
  }
}

function originOf(index: number, sizeOf: (index: number) => number): number {
  let total = 0
  for (let at = 0; at < index; at += 1) total += sizeOf(at)
  return total
}

export function frameBox(frame: InstalledVisualFrame, metrics: SheetMetrics): ArrangeBox {
  return {
    id: frame.visual.id,
    x: originOf(frame.fromColumn, metrics.columnWidth) + frame.marginX,
    y: originOf(frame.fromRow, metrics.rowHeight) + frame.marginY,
    width: frame.width,
    height: frame.height,
  }
}

export function cellOrigin(
  row: number,
  column: number,
  metrics: SheetMetrics,
): { x: number; y: number } {
  return { x: originOf(column, metrics.columnWidth), y: originOf(row, metrics.rowHeight) }
}

/// twoCellAnchor markers for a sheet-pixel box (both markers walked from
/// cell A1, so the result is fully normalized).
export function anchorForBox(
  box: { x: number; y: number; width: number; height: number },
  metrics: SheetMetrics,
): WorkbookVisualObject['anchor'] {
  const fromX = walkMarker(
    markerFrom(0, 0),
    Math.max(0, box.x),
    metrics.columnWidth,
    metrics.maxColumn,
  )
  const fromY = walkMarker(markerFrom(0, 0), Math.max(0, box.y), metrics.rowHeight, metrics.maxRow)
  const toX = walkMarker(fromX, Math.max(1, box.width), metrics.columnWidth, metrics.maxColumn)
  const toY = walkMarker(fromY, Math.max(1, box.height), metrics.rowHeight, metrics.maxRow)
  return {
    fromRow: fromY.index,
    fromColumn: fromX.index,
    fromRowOffset: Math.round(fromY.offset * EMU_PER_PIXEL),
    fromColumnOffset: Math.round(fromX.offset * EMU_PER_PIXEL),
    toRow: toY.index,
    toColumn: toX.index,
    toRowOffset: Math.round(toY.offset * EMU_PER_PIXEL),
    toColumnOffset: Math.round(toX.offset * EMU_PER_PIXEL),
  }
}
