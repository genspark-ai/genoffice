/**
 * Excel repeat-paste for anchor-shaped targets.
 *
 * Copying C2:E2 and pasting into C2:C14 must repeat the row for each
 * selected row, spilling the source's full width — that's how Excel (and
 * Google Sheets) bulk-fill works. Univer tiles only when BOTH dimensions of
 * the target are exact multiples of the source; a 13×1 target over a 1×3
 * source falls through to a single anchored paste, so nothing repeats
 * (alpha feedback: "selecting C2:C14 does not bulk-copy").
 *
 * Before Univer computes the paste range, widen an anchor-shaped selection
 * (one axis a whole multiple with more than one repetition, the other axis
 * SMALLER than the source) to the source's extent on the smaller axis. The
 * selection then satisfies upstream's both-multiples rule and its own tiling,
 * merge guards, and filtered-row handling do the rest — and the widened
 * selection is exactly what Excel leaves selected after such a paste.
 * Cut-pastes move content and never repeat, so they are left untouched.
 */
import { ICommandService, IUniverInstanceService, RANGE_TYPE } from '@univerjs/core'
import {
  SelectionMoveType,
  SetSelectionsOperation,
  SheetsSelectionsService,
  getSheetCommandTarget,
} from '@univerjs/sheets'
import { COPY_TYPE, ISheetClipboardService } from '@univerjs/sheets-ui'

import type { UniverRuntime } from './univer-state'

interface MatrixLike {
  getDataRange(): { startRow: number; endRow: number; startColumn: number; endColumn: number }
}

interface SelectionRect {
  startRow: number
  endRow: number
  startColumn: number
  endColumn: number
  rangeType?: number
}

interface SheetBounds {
  rowCount: number
  columnCount: number
}

/**
 * The widened selection rectangle for an anchor-shaped paste target, or null
 * when upstream's behavior should stand. Whole-row/column/sheet selections
 * and expansions that would run past the sheet edge are left to upstream.
 */
export function anchorTileExpansion(
  selection: SelectionRect,
  sourceRows: number,
  sourceColumns: number,
  bounds?: SheetBounds,
): SelectionRect | null {
  if (sourceRows <= 0 || sourceColumns <= 0) return null
  if (selection.rangeType !== undefined && selection.rangeType !== RANGE_TYPE.NORMAL) return null
  const rows = selection.endRow - selection.startRow + 1
  const columns = selection.endColumn - selection.startColumn + 1
  let expanded: SelectionRect | null = null
  if (rows % sourceRows === 0 && rows > sourceRows && columns < sourceColumns) {
    expanded = {
      startRow: selection.startRow,
      endRow: selection.endRow,
      startColumn: selection.startColumn,
      endColumn: selection.startColumn + sourceColumns - 1,
      rangeType: RANGE_TYPE.NORMAL,
    }
  } else if (columns % sourceColumns === 0 && columns > sourceColumns && rows < sourceRows) {
    expanded = {
      startRow: selection.startRow,
      endRow: selection.startRow + sourceRows - 1,
      startColumn: selection.startColumn,
      endColumn: selection.endColumn,
      rangeType: RANGE_TYPE.NORMAL,
    }
  }
  if (
    expanded &&
    bounds &&
    (expanded.endRow >= bounds.rowCount || expanded.endColumn >= bounds.columnCount)
  ) {
    return null
  }
  return expanded
}

export function installClipboardAnchorTile(runtime: UniverRuntime): void {
  const injector = runtime.univer.__getInjector()
  // patching private methods of Univer's clipboard service singleton
  const service = injector.get(ISheetClipboardService) as unknown as {
    _pasteInternal: (copyId: string, pasteType: unknown) => Promise<unknown>
    _getPastedRange: (cellMatrix: MatrixLike) => unknown
    _copyContentCache?: { get(id: string): { copyType?: unknown } | undefined }
  }
  const originalPasteInternal = service._pasteInternal.bind(service)
  const originalGetPastedRange = service._getPastedRange.bind(service)
  // cut-pastes move content; only copy-pastes repeat
  let copyPaste = true
  service._pasteInternal = async function (copyId: string, pasteType: unknown): Promise<unknown> {
    copyPaste = service._copyContentCache?.get(copyId)?.copyType !== COPY_TYPE.CUT
    try {
      return await originalPasteInternal(copyId, pasteType)
    } finally {
      copyPaste = true
    }
  }
  service._getPastedRange = function (cellMatrix: MatrixLike): unknown {
    if (copyPaste) {
      const selection = injector.get(SheetsSelectionsService).getCurrentLastSelection()
      const target = getSheetCommandTarget(injector.get(IUniverInstanceService))
      if (selection?.range && target) {
        const data = cellMatrix.getDataRange()
        const expanded = anchorTileExpansion(
          selection.range,
          data.endRow - data.startRow + 1,
          data.endColumn - data.startColumn + 1,
          {
            rowCount: target.worksheet.getRowCount(),
            columnCount: target.worksheet.getColumnCount(),
          },
        )
        if (expanded) {
          injector.get(ICommandService).syncExecuteCommand(SetSelectionsOperation.id, {
            unitId: target.unitId,
            subUnitId: target.subUnitId,
            type: SelectionMoveType.MOVE_END,
            selections: [{ range: expanded, primary: selection.primary }],
          })
        }
      }
    }
    return originalGetPastedRange(cellMatrix)
  }
}
