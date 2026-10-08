/**
 * Chunk-level wrap row auto-height: Univer fits a row by walking every column
 * through `calculateAutoHeightForCell` (cell, merge and composed-style lookups,
 * a document layout per wrap cell) one command per row run. A streamed chunk
 * already holds its cells, so its rows are fit from that payload and written
 * in one row-auto-height mutation.
 */
import type { IStyleData } from '@univerjs/core'
import {
  DEFAULT_PADDING_DATA,
  DocSimpleSkeleton,
  FontCache,
  getFontStyleString,
} from '@univerjs/engine-render'

import type { WorkbookCellStyle, WorkbookRangeResult } from '../shared/desktop-api'
import { excelWrapHeight } from './autofit-line-pitch'
import { AUTOFIT_EXTRA_PX } from './autofit-wrap-budget'

/// Univer's MAXIMUM_ROW_HEIGHT.
const MAX_ROW_HEIGHT_PX = 2000

export interface RowAutoHeight {
  row: number
  autoHeight: number
}

export interface WrapCellMeasurer {
  /// null: the cell needs the full Univer measure (rich text, rotation).
  cellHeight(
    cell: WorkbookRangeResult['cells'][number],
    style: WorkbookCellStyle | undefined,
    wraps: boolean,
    columnWidthPx: number,
  ): number | null
}

export interface BatchWrapHeightInput {
  cells: WorkbookRangeResult['cells']
  styles: readonly WorkbookCellStyle[]
  rows: readonly number[]
  merges: WorkbookRangeResult['merges'] | undefined
  inheritedWrap: (row: number, column: number) => boolean
  columnWidthPx: (column: number) => number
  defaultRowHeightPx: number
  currentAutoHeight: (row: number) => number | undefined
  keepTaller: boolean
}

export interface BatchWrapHeightResult {
  heights: RowAutoHeight[]
  fallbackRows: number[]
}

export function batchWrapRowHeights(
  input: BatchWrapHeightInput,
  measurer: WrapCellMeasurer,
): BatchWrapHeightResult {
  const wanted = new Set(input.rows)
  const heights = new Map<number, number>()
  const fallback = new Set<number>()
  for (const row of wanted) heights.set(row, input.defaultRowHeightPx)
  const merged = new Set<string>()
  for (const merge of input.merges ?? []) {
    for (let row = merge.startRow; row <= merge.endRow; row += 1) {
      if (!wanted.has(row)) continue
      for (let column = merge.startColumn; column <= merge.endColumn; column += 1) {
        merged.add(`${row}:${column}`)
      }
    }
  }
  for (const cell of input.cells) {
    if (!wanted.has(cell.row) || fallback.has(cell.row)) continue
    if (cell.value === undefined || cell.value === null || cell.value === '') continue
    if (merged.has(`${cell.row}:${cell.column}`)) continue
    const style = cell.styleIndex === undefined ? undefined : input.styles[cell.styleIndex]
    const wraps = style ? style.wrapText === true : input.inheritedWrap(cell.row, cell.column)
    const height = measurer.cellHeight(cell, style, wraps, input.columnWidthPx(cell.column))
    if (height === null) {
      fallback.add(cell.row)
      heights.delete(cell.row)
      continue
    }
    if (height > (heights.get(cell.row) as number)) heights.set(cell.row, height)
  }
  const result: RowAutoHeight[] = []
  for (const [row, height] of heights) {
    let autoHeight = Math.min(height, MAX_ROW_HEIGHT_PX)
    const current = input.keepTaller ? input.currentAutoHeight(row) : undefined
    if (current !== undefined && current > autoHeight) autoHeight = current
    result.push({ row, autoHeight })
  }
  result.sort((a, b) => a.row - b.row)
  return { heights: result, fallbackRows: [...fallback].sort((a, b) => a - b) }
}

interface FontEntry {
  style: IStyleData
  fontCache: string
  lineBoxPx: number
  paddingX: number
  paddingY: number
}

/// Univer's own measure minus the per-column walk; font metrics resolve once
/// per style object.
export function createUniverWrapMeasurer(
  toUniverStyle: (style: WorkbookCellStyle) => IStyleData,
): WrapCellMeasurer & { readonly fontEntries: number } {
  const entries = new Map<WorkbookCellStyle | null, FontEntry>()
  const entryFor = (style: WorkbookCellStyle | undefined): FontEntry => {
    const key = style ?? null
    let entry = entries.get(key)
    if (entry) return entry
    const univerStyle = style ? toUniverStyle(style) : {}
    const { fontCache } = getFontStyleString(univerStyle)
    const box = FontCache.getMeasureText('A', fontCache)
    const pd = univerStyle.pd
    entry = {
      style: univerStyle,
      fontCache,
      lineBoxPx: box.fontBoundingBoxAscent + box.fontBoundingBoxDescent,
      paddingX: (pd?.l ?? DEFAULT_PADDING_DATA.l) + (pd?.r ?? DEFAULT_PADDING_DATA.r),
      paddingY: (pd?.t ?? DEFAULT_PADDING_DATA.t) + (pd?.b ?? DEFAULT_PADDING_DATA.b),
    }
    entries.set(key, entry)
    return entry
  }
  return {
    get fontEntries() {
      return entries.size
    },
    cellHeight(cell, style, wraps, columnWidthPx) {
      if (cell.rich?.length || style?.textRotation) return null
      const entry = entryFor(style)
      if (!wraps || typeof cell.value !== 'string') return entry.lineBoxPx + entry.paddingY
      const skeleton = new DocSimpleSkeleton(
        cell.value,
        entry.fontCache,
        true,
        columnWidthPx - AUTOFIT_EXTRA_PX - entry.paddingX,
        Infinity,
      )
      skeleton.calculate()
      return excelWrapHeight(skeleton.getTotalHeight() + entry.paddingY, entry.style)
    },
  }
}
