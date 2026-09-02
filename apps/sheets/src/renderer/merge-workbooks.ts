/**
 * Merge workbooks: pick one or more spreadsheet files and append every sheet
 * to the current workbook (alpha: Olivia, #genoffice-7).
 *
 * Each source opens as a secondary sidecar session (the tab's own session
 * stays untouched); cells arrive through the same readWorkbookRange pipeline
 * the streaming loader uses, and land via journaled commands — the sheet
 * addition and its content both persist through the existing save plan
 * (SheetEditPlan.additions + journaled cell edits). Formulas import as their
 * cached computed VALUES in this version: source references would point at
 * sheets/files that don't exist here, and cross-file reference resolution is
 * its own feature.
 */
import { CellValueType, ICommandService } from '@univerjs/core'
import type { ICellData, IRange, Nullable } from '@univerjs/core'
import { AddWorksheetMergeCommand, InsertSheetCommand } from '@univerjs/sheets'
import type { WorkbookFile } from '../shared/desktop-api'
import { t } from './i18n/locale'
import type { LazyWorkbookState, UniverRuntime } from './univer-state'
import { characterWidthToPixels, toUniverStyle, workbookStructureLocked } from './univer-sync'

/// Stay under the IPC schema's MAX_RANGE_CELLS (100k) with headroom.
const READ_CHUNK_CELLS = 90_000
const SET_RANGE_VALUES_COMMAND = 'sheet.command.set-range-values'
const SET_COL_WIDTH_COMMAND = 'sheet.command.set-worksheet-col-width'
const SET_ROW_HEIGHT_COMMAND = 'sheet.command.set-row-height'

export interface MergeWorkbooksDeps {
  runtime: UniverRuntime
  lazyWorkbookRef: { readonly current: LazyWorkbookState | null }
  setMessage: (message: string) => void
}

/** "Name" → "Name (2)" → "Name (3)" against the taken set (case-insensitive like Excel) */
export function dedupeSheetName(name: string, taken: ReadonlySet<string>): string {
  const lower = new Set([...taken].map((entry) => entry.toLowerCase()))
  if (!lower.has(name.toLowerCase())) return name
  for (let n = 2; ; n += 1) {
    const candidate = `${name} (${n})`
    if (!lower.has(candidate.toLowerCase())) return candidate
  }
}

/** rows per read chunk so rows × columns stays under the IPC cell cap */
export function chunkRowsFor(columnCount: number): number {
  return Math.max(1, Math.floor(READ_CHUNK_CELLS / Math.max(1, columnCount)))
}

type CellMatrix = Record<number, Record<number, ICellData>>

interface ImportedSheet {
  readonly name: string
  readonly rowCount: number
  readonly columnCount: number
  readonly matrix: CellMatrix
  readonly cellCount: number
  readonly merges: IRange[]
  readonly rowHeights: { row: number; height: number }[]
  readonly columnWidths: { startColumn: number; endColumn: number; width: number }[]
}

async function readSourceSheet(
  file: WorkbookFile,
  sheet: WorkbookFile['sheets'][number],
): Promise<ImportedSheet> {
  const matrix: CellMatrix = {}
  const merges: IRange[] = []
  const mergeKeys = new Set<string>()
  const rowHeights: { row: number; height: number }[] = []
  let cellCount = 0
  const columns = Math.max(1, sheet.columnCount)
  const step = chunkRowsFor(columns)
  for (let startRow = 0; startRow < sheet.rowCount; startRow += step) {
    const endRow = Math.min(sheet.rowCount - 1, startRow + step - 1)
    // The sidecar indexes asynchronously and returns whatever is ready —
    // poll until the chunk's rows are covered or the source stops indexing,
    // otherwise unindexed rows would be silently dropped.
    const deadline = Date.now() + 120_000
    let result = await window.desktopApi.readWorkbookRange({
      sessionId: file.sessionId,
      sheetId: sheet.id,
      range: { startRow, endRow, startColumn: 0, endColumn: columns - 1 },
    })
    while (
      !result.indexingComplete &&
      (result.indexedThroughRow === null || result.indexedThroughRow < endRow)
    ) {
      if (Date.now() > deadline) throw new Error(t('appMergeWorkbooksFailed'))
      await new Promise((resolve) => setTimeout(resolve, 400))
      result = await window.desktopApi.readWorkbookRange({
        sessionId: file.sessionId,
        sheetId: sheet.id,
        range: { startRow, endRow, startColumn: 0, endColumn: columns - 1 },
      })
    }
    for (const cell of result.cells) {
      const target: ICellData = {}
      const style = cell.styleIndex === undefined ? undefined : file.styles[cell.styleIndex]
      if (style) target.s = toUniverStyle(style) as ICellData['s']
      const value = cell.value
      if (typeof value === 'string') {
        target.v = value
        target.t = CellValueType.STRING
      } else if (value !== null && value !== undefined) {
        target.v = value as ICellData['v']
      }
      if (target.v === undefined && target.s === undefined) continue
      ;(matrix[cell.row] ??= {})[cell.column] = target
      cellCount += 1
    }
    for (const merge of result.merges) {
      // a merge spanning a chunk boundary comes back from both chunks
      const key = `${merge.startRow}:${merge.endRow}:${merge.startColumn}:${merge.endColumn}`
      if (mergeKeys.has(key)) continue
      mergeKeys.add(key)
      merges.push({
        startRow: merge.startRow,
        endRow: merge.endRow,
        startColumn: merge.startColumn,
        endColumn: merge.endColumn,
      })
    }
    for (const row of result.rows) {
      if (row.height !== undefined && row.customHeight === true) {
        rowHeights.push({ row: row.row, height: Math.round((row.height * 96) / 72) })
      }
    }
  }
  // file widths are Excel character units, heights are points — Univer
  // commands expect pixels (same conversions the workbook loader applies)
  const columnWidths = (sheet.columnWidths ?? [])
    .filter(
      (width): width is typeof width & { width: number } =>
        typeof width.width === 'number' && !width.hidden,
    )
    .map((width) => ({
      startColumn: width.startColumn,
      endColumn: width.endColumn,
      width: characterWidthToPixels(width.width),
    }))
  return {
    name: sheet.name,
    rowCount: sheet.rowCount,
    columnCount: sheet.columnCount,
    matrix,
    cellCount,
    merges,
    rowHeights,
    columnWidths,
  }
}

/** split the sheet's matrix into row bands so no single mutation is huge */
function matrixBands(matrix: CellMatrix, bandRows: number): CellMatrix[] {
  const bands = new Map<number, CellMatrix>()
  for (const [rowKey, rowValue] of Object.entries(matrix)) {
    const row = Number(rowKey)
    const band = Math.floor(row / bandRows)
    ;(bands.get(band) ?? bands.set(band, {}).get(band)!)[row] = rowValue
  }
  return [...bands.entries()].sort((a, b) => a[0] - b[0]).map(([, value]) => value)
}

function boundsOf(matrix: CellMatrix): IRange | null {
  let startRow = Infinity
  let endRow = -1
  let startColumn = Infinity
  let endColumn = -1
  for (const [rowKey, rowValue] of Object.entries(matrix)) {
    const row = Number(rowKey)
    startRow = Math.min(startRow, row)
    endRow = Math.max(endRow, row)
    for (const columnKey of Object.keys(rowValue)) {
      const column = Number(columnKey)
      startColumn = Math.min(startColumn, column)
      endColumn = Math.max(endColumn, column)
    }
  }
  if (endRow < 0) return null
  return { startRow, endRow, startColumn, endColumn }
}

export interface MergeSourcesResult {
  importedSheets: number
  files: number
  /** final (deduped) names of the sheets created in the current workbook */
  sheetNames: string[]
}

/** Core import loop over already-opened source sessions; closes them when done. */
export async function mergeSourcesIntoCurrent(
  deps: MergeWorkbooksDeps,
  sources: WorkbookFile[],
): Promise<MergeSourcesResult> {
  const { runtime, setMessage } = deps
  try {
    const workbook = runtime.univerAPI.getActiveWorkbook()
    if (!workbook) throw new Error(t('appMergeWorkbooksFailed'))
    const commandService = runtime.univer.__getInjector().get(ICommandService)
    const unitId = workbook.getId()
    const taken = new Set<string>(workbook.getSheets().map((sheet) => sheet.getSheetName()))
    const sheetNames: string[] = []
    let importedSheets = 0
    for (const file of sources) {
      for (const sheetMeta of file.sheets) {
        setMessage(t('appMergeWorkbooksReading', { file: file.name, sheet: sheetMeta.name }))
        const imported = await readSourceSheet(file, sheetMeta)
        const name = dedupeSheetName(imported.name, taken)
        taken.add(name)
        sheetNames.push(name)
        const sheetId = `merge-${Date.now().toString(36)}-${Math.floor(Math.random() * 1e9).toString(36)}`
        const inserted = await commandService.executeCommand(InsertSheetCommand.id, {
          unitId,
          index: workbook.getSheets().length,
          sheet: {
            id: sheetId,
            name,
            rowCount: Math.max(1000, imported.rowCount + 100),
            columnCount: Math.max(26, imported.columnCount + 5),
          },
        })
        if (!inserted) throw new Error(t('appMergeWorkbooksInsertFailed', { sheet: name }))
        const bandRows = chunkRowsFor(Math.max(1, imported.columnCount))
        for (const band of matrixBands(imported.matrix, bandRows)) {
          const bounds = boundsOf(band)
          if (!bounds) continue
          await commandService.executeCommand(SET_RANGE_VALUES_COMMAND, {
            unitId,
            subUnitId: sheetId,
            range: bounds,
            value: band as Nullable<Record<number, Record<number, ICellData>>>,
          })
        }
        for (const width of imported.columnWidths) {
          await commandService.executeCommand(SET_COL_WIDTH_COMMAND, {
            unitId,
            subUnitId: sheetId,
            ranges: [
              {
                startRow: 0,
                endRow: 0,
                startColumn: width.startColumn,
                endColumn: width.endColumn,
                rangeType: 2,
              },
            ],
            width: width.width,
          })
        }
        for (const rowHeight of imported.rowHeights) {
          await commandService.executeCommand(SET_ROW_HEIGHT_COMMAND, {
            unitId,
            subUnitId: sheetId,
            ranges: [
              {
                startRow: rowHeight.row,
                endRow: rowHeight.row,
                startColumn: 0,
                endColumn: 0,
                rangeType: 1,
              },
            ],
            value: rowHeight.height,
          })
        }
        if (imported.merges.length > 0) {
          await commandService.executeCommand(AddWorksheetMergeCommand.id, {
            unitId,
            subUnitId: sheetId,
            selections: imported.merges,
            defaultMerge: true,
          })
        }
        importedSheets += 1
      }
    }
    setMessage(t('appMergeWorkbooksDone', { sheets: importedSheets, files: sources.length }))
    return { importedSheets, files: sources.length, sheetNames }
  } finally {
    for (const file of sources) {
      void window.desktopApi.closeWorkbook(file.sessionId).catch(() => {})
    }
  }
}

/** Ribbon entry point: multi-file picker, then the shared import core. */
export async function mergeWorkbooksIntoCurrent(deps: MergeWorkbooksDeps): Promise<void> {
  const { lazyWorkbookRef, setMessage } = deps
  if (workbookStructureLocked(lazyWorkbookRef.current)) {
    setMessage(t('appMergeWorkbooksLocked'))
    return
  }
  setMessage(t('appMergeWorkbooksPicking'))
  try {
    const sources = await window.desktopApi.selectWorkbooksForMerge()
    if (!sources || sources.length === 0) {
      setMessage(t('appOpenCanceled'))
      return
    }
    await mergeSourcesIntoCurrent(deps, sources)
  } catch (error: unknown) {
    setMessage(error instanceof Error ? error.message : t('appMergeWorkbooksFailed'))
  }
}

/** AI entry point: open explicit attachment paths and run the shared core. */
export async function mergeAttachedWorkbooks(
  deps: MergeWorkbooksDeps,
  paths: string[],
): Promise<MergeSourcesResult> {
  if (workbookStructureLocked(deps.lazyWorkbookRef.current)) {
    throw new Error(t('appMergeWorkbooksLocked'))
  }
  const sources = await window.desktopApi.openWorkbooksForMerge(paths)
  if (!sources || sources.length === 0) throw new Error(t('appMergeWorkbooksFailed'))
  return mergeSourcesIntoCurrent(deps, sources)
}
