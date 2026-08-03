/**
 * Clipboard plain-text interop. Univer's copy joins raw cell text with
 * tabs — booleans land as 1/0 and embedded newlines go out unquoted (as bare
 * \r), so Excel/Numbers/scripts mis-parse the paste. Wrap the clipboard
 * service's content generation and rebuild the plain slice with Excel's TSV
 * conventions: TRUE/FALSE for booleans, CSV-style quoting for fields carrying
 * tabs/newlines/quotes, newlines normalized to \n.
 */
import {
  CellModeEnum,
  CellValueType,
  IUniverInstanceService,
  extractPureTextFromCell,
  type ICellData,
} from '@univerjs/core'
import { ISheetClipboardService } from '@univerjs/preset-sheets-core'

import type { UniverRuntime } from './univer-state'

type ClipboardCell = ICellData & { displayV?: string }

/// One TSV field: booleans as TRUE/FALSE, newlines normalized to \n, and
/// CSV-style quoting when the text carries tabs/newlines/quotes.
export function clipboardField(cell: ClipboardCell | null | undefined): string {
  if (!cell) return ''
  const text =
    cell.t === CellValueType.BOOLEAN
      ? Number(cell.v) === 0
        ? 'FALSE'
        : 'TRUE'
      : (cell.displayV ??
        // extractPureTextFromCell strips \r line breaks — keep string values
        // verbatim so embedded newlines survive to be quoted below
        (typeof cell.v === 'string' ? cell.v : extractPureTextFromCell(cell)))
  const normalized = text.replace(/\r\n|\r+/g, '\n').replace(/\n+$/, '')
  return /[\t\n"]/.test(normalized) ? `"${normalized.replace(/"/g, '""')}"` : normalized
}

/// The whole plain-text clipboard slice: fields per (row, column) of the
/// copied discrete range (absolute coordinates, filtered rows already gone).
export function plainTextFromCells(
  rows: readonly number[],
  columns: readonly number[],
  cellAt: (row: number, column: number) => ClipboardCell | null | undefined,
): string {
  return rows
    .map((row) => columns.map((column) => clipboardField(cellAt(row, column))).join('\t'))
    .join('\n')
}

interface CopyContent {
  plain: string
  discreteRange: { rows: number[]; cols: number[] }
}

export function installTsvClipboardFix(runtime: UniverRuntime): { dispose(): void } {
  const injector = runtime.univer.__getInjector()
  const instanceService = injector.get(IUniverInstanceService)
  const clipboardService = injector.get(ISheetClipboardService) as {
    generateCopyContent(
      workbookId: string,
      worksheetId: string,
      range: unknown,
      options?: { copyHookType?: string },
    ): CopyContent | null | undefined
  }
  const original = clipboardService.generateCopyContent.bind(clipboardService)
  clipboardService.generateCopyContent = (workbookId, worksheetId, range, options) => {
    const content = original(workbookId, worksheetId, range, options)
    // Special copies (formula-only) keep Univer's plain text.
    if (!content || (options?.copyHookType && options.copyHookType !== 'default-copy')) {
      return content
    }
    try {
      const worksheet = instanceService
        .getUniverSheetInstance(workbookId)
        ?.getSheetBySheetId(worksheetId)
      if (!worksheet) return content
      const { rows, cols } = content.discreteRange
      const bounds = {
        startRow: Math.min(...rows),
        endRow: Math.max(...rows),
        startColumn: Math.min(...cols),
        endColumn: Math.max(...cols),
      }
      const matrix = worksheet.getMatrixWithMergedCells(
        bounds.startRow,
        bounds.startColumn,
        bounds.endRow,
        bounds.endColumn,
        CellModeEnum.Both,
      )
      const plain = plainTextFromCells(
        rows,
        cols,
        (row, column) => matrix.getValue(row, column) ?? null,
      )
      return { ...content, plain }
    } catch {
      return content
    }
  }
  return {
    dispose: () => {
      clipboardService.generateCopyContent = original
    },
  }
}
