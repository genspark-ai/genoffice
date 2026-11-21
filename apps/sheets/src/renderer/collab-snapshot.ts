import { columnLabel, formatAddress } from '@genoffice/xlsx-gateway/domain/cell-address'
import type {
  CellFormatState,
  CellState,
  WorkbookSnapshot,
  WorksheetState,
} from '@genoffice/xlsx-gateway/domain/workbook.types'
import type { IStyleData } from '@univerjs/core'
import type { SheetVisual } from '@genoffice/xlsx-gateway/domain/chart-visual'

import type { WorkbookVisualObject } from '../shared/desktop-api'
import { toSelectionFormat } from './selection-format'
import type { UniverRuntime } from './univer-state'

/**
 * Capture the live workbook as the snapshot the room publishes.
 *
 * The file model (`InMemoryWorkbookAdapter`) cannot serve this purpose: it only
 * changes when an AI plan is committed, so typing in a cell — the whole point of
 * co-editing — would never reach the room. The Univer grid is the one place that
 * holds every edit (manual, AI and remote), so the snapshot is read from there.
 * Sheet ids come from the grid too, which is what makes the receiving side's
 * `getSheetBySheetId` lookup land on the same worksheet.
 *
 * Only fields the in-place apply path can replay are emitted. Charts are not
 * Univer cells — they live in the journal's `visualAdds` — so the caller passes
 * the live visuals in explicitly and they ride in `WorkbookSnapshot.visuals`.
 */

interface RawCell {
  readonly v?: string | number | boolean | null
  readonly f?: string
  readonly s?: string | IStyleData | null
}

interface RawRow {
  readonly h?: number
}

interface RawColumn {
  readonly w?: number
}

interface RawMerge {
  readonly startRow: number
  readonly startColumn: number
  readonly endRow: number
  readonly endColumn: number
}

/** Univer row heights are pixels; the snapshot's rowHeights are points. */
function pixelsToPoints(pixels: number): number {
  return Math.max(1, Math.round((pixels * 72) / 96))
}

/** Format echo → the resolved per-cell format the apply path understands. */
function toCellFormat(style: IStyleData, numberFormat: string): CellFormatState | null {
  const echo = toSelectionFormat(style, numberFormat)
  const format: Record<string, unknown> = {}
  if (echo.bold) format.bold = true
  if (echo.italic) format.italic = true
  if (echo.underline) format.underline = true
  if (echo.strike) format.strikethrough = true
  if (echo.fontFamily) format.fontFamily = echo.fontFamily
  if (echo.fontSize) format.fontSize = echo.fontSize
  if (echo.fontColor) format.fontColor = echo.fontColor
  if (echo.fillColor) format.fillColor = echo.fillColor
  if (echo.numberFormat && echo.numberFormat !== 'General') format.numberFormat = echo.numberFormat
  if (
    echo.horizontalAlignment === 'left' ||
    echo.horizontalAlignment === 'center' ||
    echo.horizontalAlignment === 'right'
  ) {
    format.horizontalAlign = echo.horizontalAlignment
  }
  if (
    echo.verticalAlignment === 'top' ||
    echo.verticalAlignment === 'center' ||
    echo.verticalAlignment === 'bottom'
  ) {
    format.verticalAlign = echo.verticalAlignment
  }
  if (echo.wrap) format.wrapText = true
  if (echo.textRotation !== null) format.textRotation = echo.textRotation
  if (echo.indent > 0) format.indent = echo.indent
  return Object.keys(format).length > 0 ? (format as CellFormatState) : null
}

export function snapshotLiveWorkbook(
  runtime: UniverRuntime | null,
  revision: number,
  visuals?: readonly WorkbookVisualObject[],
): WorkbookSnapshot | null {
  const workbook = runtime?.univerAPI.getActiveWorkbook()
  if (!workbook) return null
  const data = workbook.getSnapshot()
  const stylePool = (data.styles ?? {}) as Record<string, IStyleData | null | undefined>
  const sheets: WorksheetState[] = []

  for (const sheetId of data.sheetOrder ?? []) {
    const raw = data.sheets[sheetId]
    const worksheet = workbook.getSheetBySheetId(sheetId)
    if (!raw || !worksheet) continue

    const cells: Record<string, CellState> = {}
    const styles: Record<string, CellFormatState> = {}
    const cellData = (raw.cellData ?? {}) as Record<string, Record<string, RawCell | null>>
    for (const [rowKey, row] of Object.entries(cellData)) {
      for (const [columnKey, cell] of Object.entries(row ?? {})) {
        if (!cell) continue
        const address = formatAddress(Number(rowKey), Number(columnKey))
        if (cell.f) cells[address] = { value: null, formula: cell.f }
        else if (cell.v !== undefined && cell.v !== null && cell.v !== '')
          cells[address] = { value: cell.v }

        const style = typeof cell.s === 'string' ? stylePool[cell.s] : cell.s
        if (!style || typeof style !== 'object') continue
        let pattern = ''
        try {
          pattern = worksheet.getRange(address).getNumberFormat()
        } catch {
          // number-format resolution fails on untouched cells; General is fine
        }
        const format = toCellFormat(style, pattern)
        if (format) styles[address] = format
      }
    }

    const merges = ((raw.mergeData ?? []) as RawMerge[]).map(
      (merge) =>
        `${formatAddress(merge.startRow, merge.startColumn)}:${formatAddress(
          merge.endRow,
          merge.endColumn,
        )}`,
    )
    const rowHeights: Record<string, number> = {}
    for (const [key, row] of Object.entries((raw.rowData ?? {}) as Record<string, RawRow>)) {
      if (typeof row?.h === 'number') rowHeights[String(Number(key) + 1)] = pixelsToPoints(row.h)
    }
    const colWidths: Record<string, number> = {}
    for (const [key, column] of Object.entries((raw.columnData ?? {}) as Record<string, RawColumn>)) {
      if (typeof column?.w === 'number') colWidths[columnLabel(Number(key))] = Math.round(column.w)
    }

    sheets.push({
      id: sheetId,
      name: raw.name ?? worksheet.getSheetName(),
      cells,
      styles,
      merges,
      rowHeights,
      colWidths,
      gridRows: raw.rowCount ?? worksheet.getMaxRows(),
      gridColumns: raw.columnCount ?? worksheet.getMaxColumns(),
      // the snapshot schema carries charts; filter to that kind
      visuals: visuals?.filter(
        (visual) => visual.sheetId === sheetId && visual.kind === 'chart',
      ) as SheetVisual[] | undefined,
    })
  }

  return { revision, sheets }
}
