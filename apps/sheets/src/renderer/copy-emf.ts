/**
 * Copy also lands the selection as an EMF on the Windows clipboard. CAD paste
 * dialogs (AutoCAD, ZWCAD PASTESPEC) list the formats the clipboard actually
 * carries, so without this a GenOffice copy only offers text; with it a
 * "Picture (Enhanced Metafile)" entry appears and the table lands in the
 * drawing as vector artwork (genoffice#1985 stage 1 — no OLE server, no
 * embedded-object activation).
 *
 * Failure is silent by contract: the text and HTML formats the grid wrote
 * stay untouched, and off Windows (or without the helper) nothing is sent.
 */
import { CellValueType, extractPureTextFromCell, numfmt, type ICellData } from '@univerjs/core'

import { FULL_LOAD_MAX_CELLS } from './app-constants'
import type { CopyEmfRequest } from '../shared/desktop-api'
import type { UniverRuntime } from './univer-state'

/// The same resident-selection budget the grid copy uses (COPY_MATERIALIZE_MAX_CELLS);
/// imported from its source to keep the two rules one constant without a cycle.
const COPY_EMF_MAX_CELLS = FULL_LOAD_MAX_CELLS

/// The worksheet surface copy-emf reads; structural methods mirror the facade
/// the print pipeline already uses (print-html.ts).
interface CopyEmfWorksheet {
  getCell(row: number, column: number): ICellData | null
  getColumnWidth(column: number): number
  getRowHeight(row: number): number
  getRowVisible(row: number): boolean
  getColVisible(column: number): boolean
  getMergeData(): Array<{
    getRow(): number
    getColumn(): number
    getHeight(): number
    getWidth(): number
  }>
}

/// The active range, as far as copy-emf reads it: the four coordinates.
type CopyEmfSelection = {
  getRow(): number
  getColumn(): number
  getHeight(): number
  getWidth(): number
} | null

/// Displayed text and alignment for one cell, matching what the TSV slice
/// carries: TRUE/FALSE centered, numbers through their number format
/// right-aligned, strings left-aligned, rich text collapsed to run text.
export function copyEmfCellDisplay(
  cell: ICellData | null,
): CopyEmfRequest['cells'][number][number] {
  if (!cell || cell.v === null || cell.v === undefined || cell.v === '') {
    return { v: '', a: 'l' }
  }
  if (cell.t === CellValueType.BOOLEAN) {
    return { v: Number(cell.v) === 0 ? 'FALSE' : 'TRUE', a: 'c' }
  }
  if (typeof cell.v === 'number') {
    const pattern = (cell.s as { n?: { pattern?: string } } | null)?.n?.pattern
    if (pattern && pattern !== 'General') {
      const text = numfmt.format(pattern, cell.v, { throws: false })
      if (typeof text === 'string') return { v: text, a: 'r' }
    }
    return { v: String(cell.v), a: 'r' }
  }
  const text = typeof cell.v === 'string' ? cell.v : extractPureTextFromCell(cell as ICellData)
  return { v: text, a: 'l' }
}

/// A merge clipped to the selection, as the helper consumes it: a relative
/// origin that stays inside the range and the visible span only. Returns
/// null for merges outside the range or degenerate single cells.
export function clippedCopyEmfMerge(
  merge: { getRow(): number; getColumn(): number; getHeight(): number; getWidth(): number },
  startRow: number,
  startColumn: number,
  rows: number,
  columns: number,
): CopyEmfRequest['merges'][number] | null {
  const offsetRow = merge.getRow() - startRow
  const offsetColumn = merge.getColumn() - startColumn
  const r = Math.max(offsetRow, 0)
  const c = Math.max(offsetColumn, 0)
  const rs = Math.min(offsetRow + merge.getHeight(), rows) - r
  const cs = Math.min(offsetColumn + merge.getWidth(), columns) - c
  // A merge wholly past the range clips to nothing; a 1×1 remainder is no
  // merge at all.
  if (rs <= 0 || cs <= 0) return null
  if (rs === 1 && cs === 1) return null
  return { r, c, rs, cs }
}

export function buildCopyEmfPayload(runtime: UniverRuntime): CopyEmfRequest | null {
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet() as CopyEmfWorksheet | null | undefined
  const selection = workbook?.getActiveRange() as CopyEmfSelection | undefined
  if (!worksheet || !selection) return null
  const startRow = selection.getRow()
  const startColumn = selection.getColumn()
  const rows = selection.getHeight()
  const columns = selection.getWidth()
  if (rows <= 0 || columns <= 0) return null
  if (rows * columns > COPY_EMF_MAX_CELLS) return null

  const colWidths: number[] = []
  for (let c = 0; c < columns; c += 1) {
    const column = startColumn + c
    // Hidden columns carry no artwork: a zero width draws nothing in the
    // metafile and the grid lines collapse onto one edge.
    const width = worksheet.getColVisible(column) ? worksheet.getColumnWidth(column) : 0
    colWidths.push(Math.max(width, 0))
  }
  const rowHeights: number[] = []
  for (let r = 0; r < rows; r += 1) {
    const row = startRow + r
    const height = worksheet.getRowVisible(row) ? worksheet.getRowHeight(row) : 0
    rowHeights.push(Math.max(height, 0))
  }

  const cells: CopyEmfRequest['cells'] = []
  for (let r = 0; r < rows; r += 1) {
    const row: CopyEmfRequest['cells'][number] = []
    for (let c = 0; c < columns; c += 1) {
      row.push(copyEmfCellDisplay(worksheet.getCell(startRow + r, startColumn + c)))
    }
    cells.push(row)
  }

  const merges: CopyEmfRequest['merges'] = []
  for (const merge of worksheet.getMergeData()) {
    const clipped = clippedCopyEmfMerge(merge, startRow, startColumn, rows, columns)
    if (clipped) merges.push(clipped)
  }

  return { colWidths, rowHeights, cells, merges }
}

/// Fire-and-forget: the helper appends the EMF beside the formats Univer's
/// copy already wrote, and every failure stays silent by design.
/** asked once, then cached: building the payload walks the whole selection */
let copyEmfAvailableCache: boolean | null = null

export function emitCopyEmf(runtime: UniverRuntime): void {
  void (async () => {
    try {
      if (copyEmfAvailableCache === null) {
        copyEmfAvailableCache = await window.desktopApi.isCopyEmfAvailable()
      }
      if (!copyEmfAvailableCache) return
      const payload = buildCopyEmfPayload(runtime)
      if (!payload) return
      await window.desktopApi.appendCopyEmf(payload)
    } catch {
      // The EMF is an enhancement; the copy's text formats already landed.
    }
  })()
}
