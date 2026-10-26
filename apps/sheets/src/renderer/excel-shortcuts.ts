/**
 * Excel-standard shortcuts Univer doesn't ship.
 *
 * Univer's KeyCode enum stops at SPACE/arrows — Home, End, PageUp and
 * PageDown don't even have names, so nothing upstream can bind them. The
 * shortcut dispatcher matches on the raw browser keyCode, so registering
 * with the numeric codes works. Everything here reuses the
 * whenSheetEditorFocused precondition: with the cell editor or formula bar
 * active these keys keep their text-editing meaning.
 *
 * - Ctrl/Cmd+PgDn / PgUp: next / previous visible worksheet tab
 *   (mac also Option+←/→, Excel-for-mac parity)
 * - Ctrl/Cmd+Home: first unfrozen cell (A1 without frozen panes)
 * - Ctrl/Cmd+End: last used cell (content or formatting, like Excel)
 * - Home: start of the current row (first unfrozen column)
 * - Ctrl+Space / Shift+Space: select the whole column / row of the
 *   active cell (mac: ⌃Space — ⌘Space belongs to Spotlight)
 * - Ctrl+Shift+= / Ctrl+Numpad+ and Ctrl+- / Ctrl+Numpad-: insert / delete
 *   cells (whole rows/columns act directly, otherwise the dialog opens)
 * - Shift+Home / Ctrl+Shift+Home / Ctrl+Shift+End: extend the selection from
 *   the active cell to the row start / sheet origin / last used cell
 * - Ctrl+Shift+8 (*): select the current region around the selection
 * - Ctrl+2 / 3 / 4: bold / italic / underline (Excel's digit aliases)
 * - Ctrl+Shift+Enter: commit the formula into every selected cell (legacy
 *   array entry; results spill dynamically, so this is the Ctrl+Enter fill)
 * - mac editing keys Excel-for-Mac accepts next to Univer's own: ⌘Enter
 *   fills the selection like ⌃Enter, ⌃⌥Enter / ⌘⌥Enter break the line like
 *   ⌥Enter, ⌘T cycles the reference like F4
 */
import {
  CommandType,
  EDITOR_ACTIVATED,
  FOCUSING_EDITOR_INPUT_FORMULA,
  FOCUSING_EDITOR_STANDALONE,
  FOCUSING_SHEET,
  FOCUSING_UNIVER_EDITOR,
  ICommandService,
  IUniverInstanceService,
  RANGE_TYPE,
  Rectangle,
} from '@univerjs/core'
import type { IAccessor, ICommand, IRange, Worksheet } from '@univerjs/core'
import { IShortcutService, KeyCode, MetaKeys } from '@univerjs/ui'
import type { IShortcutItem } from '@univerjs/ui'
import {
  ScrollToCellOperation,
  SelectionMoveType,
  SetColHiddenCommand,
  SetRowHiddenCommand,
  SetSelectionsOperation,
  SetSpecificColsVisibleCommand,
  SetSpecificRowsVisibleCommand,
  SetWorksheetActivateCommand,
  SheetsSelectionsService,
  alignToMergedCellsBorders,
  getCellAtRowCol,
  getSheetCommandTarget,
} from '@univerjs/sheets'
import {
  SetCellEditVisibleOperation,
  SetRangeBoldCommand,
  SetRangeItalicCommand,
  SetRangeUnderlineCommand,
  whenSheetEditorActivated,
  whenSheetEditorFocused,
} from '@univerjs/sheets-ui'
import { BreakLineCommand } from '@univerjs/docs-ui'
import { ReferenceAbsoluteOperation } from '@univerjs/sheets-formula-ui'
import { DeviceInputEventType } from '@univerjs/engine-render'
import type { IContextService } from '@univerjs/core'
import type { UniverRuntime } from './univer-state'
import { requestCellsAction } from './insert-delete-cells'
import type { CellsMode } from './insert-delete-cells'
import { excelCellHasValue } from './excel-jump-nav'

// browser keycodes Univer's KeyCode enum doesn't name
const KEY_PAGE_UP = 33
const KEY_PAGE_DOWN = 34
const KEY_END = 35
const KEY_HOME = 36
const KEY_NUMPAD_ADD = 107
const KEY_NUMPAD_SUBTRACT = 109

const ACTIVATE_ADJACENT_SHEET_ID = 'genoffice.command.activate-adjacent-sheet'
const SELECT_SHEET_HOME_ID = 'genoffice.command.select-sheet-home'
const SELECT_SHEET_END_ID = 'genoffice.command.select-sheet-end'
const SELECT_ROW_HOME_ID = 'genoffice.command.select-row-home'
const SELECT_WHOLE_COLUMN_ID = 'genoffice.command.select-whole-column'
const SELECT_WHOLE_ROW_ID = 'genoffice.command.select-whole-row'
const HIDE_SELECTED_ROWS_ID = 'genoffice.command.hide-selected-rows'
const UNHIDE_SELECTED_ROWS_ID = 'genoffice.command.unhide-selected-rows'
const HIDE_SELECTED_COLS_ID = 'genoffice.command.hide-selected-cols'
const UNHIDE_SELECTED_COLS_ID = 'genoffice.command.unhide-selected-cols'
const INSERT_CELLS_ID = 'genoffice.command.insert-cells'
const DELETE_CELLS_ID = 'genoffice.command.delete-cells'
const EXPAND_ROW_HOME_ID = 'genoffice.command.expand-row-home'
const EXPAND_SHEET_HOME_ID = 'genoffice.command.expand-sheet-home'
const EXPAND_SHEET_END_ID = 'genoffice.command.expand-sheet-end'
const SELECT_CURRENT_REGION_ID = 'genoffice.command.select-current-region'

/** first visible line at or after `from` (hidden rows/columns are not landing spots) */
function firstVisible(worksheet: Worksheet, axis: 'row' | 'column', from: number): number {
  const count = axis === 'row' ? worksheet.getRowCount() : worksheet.getColumnCount()
  for (let index = from; index < count; index += 1) {
    const visible = axis === 'row' ? worksheet.getRowVisible(index) : worksheet.getColVisible(index)
    if (visible) return index
  }
  return from
}

/** first visible cell outside the frozen panes (Excel's Ctrl+Home target) */
function unfrozenOrigin(worksheet: Worksheet): { row: number; column: number } {
  const freeze = worksheet.getConfig().freeze
  const row = freeze && freeze.ySplit > 0 && freeze.startRow > 0 ? freeze.startRow : 0
  const column = freeze && freeze.xSplit > 0 && freeze.startColumn > 0 ? freeze.startColumn : 0
  return {
    row: firstVisible(worksheet, 'row', row),
    column: firstVisible(worksheet, 'column', column),
  }
}

/**
 * Excel's current region: grow the selection's bounding box until the rows
 * and columns bordering it are all empty.
 */
export function currentRegion(
  seed: IRange,
  hasValue: (row: number, column: number) => boolean,
  rowCount: number,
  columnCount: number,
): IRange {
  let { startRow, endRow, startColumn, endColumn } = seed
  const rowHasValue = (row: number): boolean => {
    for (let c = startColumn; c <= endColumn; c += 1) if (hasValue(row, c)) return true
    return false
  }
  const columnHasValue = (column: number): boolean => {
    for (let r = startRow; r <= endRow; r += 1) if (hasValue(r, column)) return true
    return false
  }
  for (let grew = true; grew;) {
    grew = false
    if (startRow > 0 && rowHasValue(startRow - 1)) {
      startRow -= 1
      grew = true
    }
    if (endRow < rowCount - 1 && rowHasValue(endRow + 1)) {
      endRow += 1
      grew = true
    }
    if (startColumn > 0 && columnHasValue(startColumn - 1)) {
      startColumn -= 1
      grew = true
    }
    if (endColumn < columnCount - 1 && columnHasValue(endColumn + 1)) {
      endColumn += 1
      grew = true
    }
  }
  return { startRow, endRow, startColumn, endColumn }
}

/**
 * Shift+navigation keeps the active cell as the anchor and moves the opposite
 * corner; `null` on an axis leaves that corner coordinate where it is.
 */
export function extendedRange(
  range: IRange,
  anchor: { row: number; column: number },
  corner: { row: number | null; column: number | null },
): IRange {
  // an untouched axis keeps its span; the anchor can sit inside it after a
  // current-region selection
  const span = (lo: number, hi: number, at: number, to: number | null): [number, number] =>
    to === null ? [lo, hi] : [Math.min(at, to), Math.max(at, to)]
  const [startRow, endRow] = span(range.startRow, range.endRow, anchor.row, corner.row)
  const [startColumn, endColumn] = span(
    range.startColumn,
    range.endColumn,
    anchor.column,
    corner.column,
  )
  return { startRow, endRow, startColumn, endColumn }
}

/**
 * The cell to reveal after an extension: the far end of each axis the key
 * moved; an untouched axis stays on the anchor so Shift+Home does not scroll
 * vertically.
 */
export function farCorner(
  range: IRange,
  anchor: { row: number; column: number },
  corner: { row: number | null; column: number | null },
): { row: number; column: number } {
  const far = (lo: number, hi: number, at: number, to: number | null): number =>
    to === null ? at : hi - at >= at - lo ? hi : lo
  return {
    row: far(range.startRow, range.endRow, anchor.row, corner.row),
    column: far(range.startColumn, range.endColumn, anchor.column, corner.column),
  }
}

function selectCell(accessor: IAccessor, row: number, column: number): Promise<boolean> {
  const target = getSheetCommandTarget(accessor.get(IUniverInstanceService))
  if (!target) return Promise.resolve(false)
  const { workbook, worksheet } = target
  const destRange = getCellAtRowCol(row, column, worksheet)
  return accessor.get(ICommandService).executeCommand(SetSelectionsOperation.id, {
    unitId: workbook.getUnitId(),
    subUnitId: worksheet.getSheetId(),
    selections: [
      {
        range: Rectangle.clone(destRange),
        primary: {
          startRow: destRange.startRow,
          startColumn: destRange.startColumn,
          endRow: destRange.endRow,
          endColumn: destRange.endColumn,
          actualRow: row,
          actualColumn: column,
          isMerged: destRange.isMerged,
          isMergedMainCell: destRange.startRow === row && destRange.startColumn === column,
        },
      },
    ],
    type: SelectionMoveType.MOVE_END,
    reveal: true,
  })
}

/**
 * `usedEndOf` supplies the file's last used cell for a sheet id (sidecar
 * dimensions) — on streamed workbooks Univer's cell matrix only holds the
 * loaded window, so Ctrl+End takes whichever reaches further per axis.
 */
export function registerExcelShortcuts(
  runtime: UniverRuntime,
  usedEndOf?: (subUnitId: string) => { row: number; column: number } | null,
  openCellsDialog: (mode: CellsMode) => void = () => {},
): void {
  const injector = runtime.univer.__getInjector()
  const commandService = injector.get(ICommandService)
  const shortcutService = injector.get(IShortcutService)

  const usedEnd = (worksheet: Worksheet): { row: number; column: number } => {
    const matrix = worksheet.getCellMatrix()
    const loaded = matrix.getLength() === 0 ? null : matrix.getDataRange()
    const fileEnd = usedEndOf?.(worksheet.getSheetId()) ?? null
    return {
      row: Math.max(loaded?.endRow ?? 0, fileEnd?.row ?? 0, 0),
      column: Math.max(loaded?.endColumn ?? 0, fileEnd?.column ?? 0, 0),
    }
  }

  const activateAdjacentSheet: ICommand<{ step: number }> = {
    id: ACTIVATE_ADJACENT_SHEET_ID,
    type: CommandType.COMMAND,
    handler: (accessor, params) => {
      if (!params) return false
      const target = getSheetCommandTarget(accessor.get(IUniverInstanceService))
      if (!target) return false
      const { workbook } = target
      const visible = workbook.getSheets().filter((sheet) => sheet.getConfig().hidden !== 1)
      const activeId = workbook.getActiveSheet()?.getSheetId()
      const index = visible.findIndex((sheet) => sheet.getSheetId() === activeId)
      if (index < 0) return false
      const next = visible[index + params.step]
      if (!next) return false // Excel doesn't wrap at either end
      // SetWorksheetActivateCommand's handler is async — syncExecuteCommand
      // throws on promise-returning handlers. Await it instead; real key
      // repeats flush microtasks between events, so the next press reads the
      // fresh active sheet.
      return accessor.get(ICommandService).executeCommand(SetWorksheetActivateCommand.id, {
        unitId: workbook.getUnitId(),
        subUnitId: next.getSheetId(),
      })
    },
  }

  const selectSheetHome: ICommand = {
    id: SELECT_SHEET_HOME_ID,
    type: CommandType.COMMAND,
    handler: (accessor) => {
      const target = getSheetCommandTarget(accessor.get(IUniverInstanceService))
      if (!target) return false
      const origin = unfrozenOrigin(target.worksheet)
      return selectCell(accessor, origin.row, origin.column)
    },
  }

  const selectSheetEnd: ICommand = {
    id: SELECT_SHEET_END_ID,
    type: CommandType.COMMAND,
    handler: (accessor) => {
      const target = getSheetCommandTarget(accessor.get(IUniverInstanceService))
      if (!target) return false
      const end = usedEnd(target.worksheet)
      return selectCell(accessor, end.row, end.column)
    },
  }

  const selectRowHome: ICommand = {
    id: SELECT_ROW_HOME_ID,
    type: CommandType.COMMAND,
    handler: (accessor) => {
      const target = getSheetCommandTarget(accessor.get(IUniverInstanceService))
      if (!target) return false
      const selection = accessor.get(SheetsSelectionsService).getCurrentLastSelection()
      if (!selection) return false
      const row = selection.primary?.actualRow ?? selection.range.startRow
      return selectCell(accessor, row, unfrozenOrigin(target.worksheet).column)
    },
  }

  const wholeRange = (accessor: IAccessor, kind: 'row' | 'column'): Promise<boolean> => {
    const target = getSheetCommandTarget(accessor.get(IUniverInstanceService))
    if (!target) return Promise.resolve(false)
    const { workbook, worksheet } = target
    const selection = accessor.get(SheetsSelectionsService).getCurrentLastSelection()
    if (!selection) return Promise.resolve(false)
    const { range, primary } = selection
    const destRange: IRange =
      kind === 'column'
        ? {
            startRow: 0,
            endRow: worksheet.getRowCount() - 1,
            startColumn: range.startColumn,
            endColumn: range.endColumn,
            rangeType: RANGE_TYPE.COLUMN,
          }
        : {
            startRow: range.startRow,
            endRow: range.endRow,
            startColumn: 0,
            endColumn: worksheet.getColumnCount() - 1,
            rangeType: RANGE_TYPE.ROW,
          }
    return accessor.get(ICommandService).executeCommand(SetSelectionsOperation.id, {
      unitId: workbook.getUnitId(),
      subUnitId: worksheet.getSheetId(),
      selections: [{ range: destRange, primary }],
      type: SelectionMoveType.MOVE_END,
    })
  }

  const selectWholeColumn: ICommand = {
    id: SELECT_WHOLE_COLUMN_ID,
    type: CommandType.COMMAND,
    handler: (accessor) => wholeRange(accessor, 'column'),
  }
  const selectWholeRow: ICommand = {
    id: SELECT_WHOLE_ROW_ID,
    type: CommandType.COMMAND,
    handler: (accessor) => wholeRange(accessor, 'row'),
  }

  /**
   * Excel's Ctrl+9/0 hides the rows/columns SPANNED by any selection —
   * Univer's own commands only act on header (ROW/COLUMN-type) selections,
   * so these wrappers convert the span and pass explicit ranges.
   */
  const axisRangeOf = (accessor: IAccessor, kind: 'row' | 'column'): IRange | null => {
    const target = getSheetCommandTarget(accessor.get(IUniverInstanceService))
    if (!target) return null
    const selection = accessor.get(SheetsSelectionsService).getCurrentLastSelection()
    if (!selection) return null
    const { range } = selection
    const { worksheet } = target
    return kind === 'row'
      ? {
          startRow: range.startRow,
          endRow: range.endRow,
          startColumn: 0,
          endColumn: worksheet.getColumnCount() - 1,
          rangeType: RANGE_TYPE.ROW,
        }
      : {
          startRow: 0,
          endRow: worksheet.getRowCount() - 1,
          startColumn: range.startColumn,
          endColumn: range.endColumn,
          rangeType: RANGE_TYPE.COLUMN,
        }
  }

  const hideOrUnhide = (
    accessor: IAccessor,
    kind: 'row' | 'column',
    action: 'hide' | 'unhide',
  ): Promise<boolean> => {
    const target = getSheetCommandTarget(accessor.get(IUniverInstanceService))
    const axisRange = axisRangeOf(accessor, kind)
    if (!target || !axisRange) return Promise.resolve(false)
    const { workbook, worksheet } = target
    if (action === 'hide') {
      // Excel refuses to hide every row/column; skip the whole-axis case
      const full =
        kind === 'row'
          ? axisRange.startRow === 0 && axisRange.endRow >= worksheet.getRowCount() - 1
          : axisRange.startColumn === 0 && axisRange.endColumn >= worksheet.getColumnCount() - 1
      if (full) return Promise.resolve(false)
    }
    const commandId =
      kind === 'row'
        ? action === 'hide'
          ? SetRowHiddenCommand.id
          : SetSpecificRowsVisibleCommand.id
        : action === 'hide'
          ? SetColHiddenCommand.id
          : SetSpecificColsVisibleCommand.id
    return accessor.get(ICommandService).executeCommand(commandId, {
      unitId: workbook.getUnitId(),
      subUnitId: worksheet.getSheetId(),
      ranges: [axisRange],
    })
  }

  const hideSelectedRows: ICommand = {
    id: HIDE_SELECTED_ROWS_ID,
    type: CommandType.COMMAND,
    handler: (accessor) => hideOrUnhide(accessor, 'row', 'hide'),
  }
  const unhideSelectedRows: ICommand = {
    id: UNHIDE_SELECTED_ROWS_ID,
    type: CommandType.COMMAND,
    handler: (accessor) => hideOrUnhide(accessor, 'row', 'unhide'),
  }
  const hideSelectedCols: ICommand = {
    id: HIDE_SELECTED_COLS_ID,
    type: CommandType.COMMAND,
    handler: (accessor) => hideOrUnhide(accessor, 'column', 'hide'),
  }
  const unhideSelectedCols: ICommand = {
    id: UNHIDE_SELECTED_COLS_ID,
    type: CommandType.COMMAND,
    handler: (accessor) => hideOrUnhide(accessor, 'column', 'unhide'),
  }

  const insertCells: ICommand = {
    id: INSERT_CELLS_ID,
    type: CommandType.COMMAND,
    handler: () => requestCellsAction(runtime, 'insert', openCellsDialog),
  }
  const deleteCells: ICommand = {
    id: DELETE_CELLS_ID,
    type: CommandType.COMMAND,
    handler: () => requestCellsAction(runtime, 'delete', openCellsDialog),
  }

  const setRange = async (
    accessor: IAccessor,
    next: IRange,
    reveal?: { row: number; column: number },
  ): Promise<boolean> => {
    const target = getSheetCommandTarget(accessor.get(IUniverInstanceService))
    const selection = accessor.get(SheetsSelectionsService).getCurrentLastSelection()
    if (!target || !selection) return false
    const { workbook, worksheet } = target
    const unitId = workbook.getUnitId()
    const subUnitId = worksheet.getSheetId()
    const range = alignToMergedCellsBorders(next, worksheet, true)
    const applied = await accessor.get(ICommandService).executeCommand(SetSelectionsOperation.id, {
      unitId,
      subUnitId,
      selections: [{ range, primary: selection.primary }],
      type: SelectionMoveType.MOVE_END,
      reveal: !reveal,
    })
    // `reveal` scrolls to the unmoved anchor; bring the moving corner in instead
    if (applied && reveal) {
      accessor.get(ICommandService).syncExecuteCommand(ScrollToCellOperation.id, {
        unitId,
        subUnitId,
        range: {
          startRow: reveal.row,
          endRow: reveal.row,
          startColumn: reveal.column,
          endColumn: reveal.column,
          rangeType: RANGE_TYPE.NORMAL,
        },
      })
    }
    return applied
  }

  const extendTo = (
    accessor: IAccessor,
    cornerOf: (worksheet: Worksheet) => { row: number | null; column: number | null },
  ): Promise<boolean> => {
    const target = getSheetCommandTarget(accessor.get(IUniverInstanceService))
    const selection = accessor.get(SheetsSelectionsService).getCurrentLastSelection()
    if (!target || !selection) return Promise.resolve(false)
    const { range, primary } = selection
    const anchor = {
      row: primary?.actualRow ?? range.startRow,
      column: primary?.actualColumn ?? range.startColumn,
    }
    const corner = cornerOf(target.worksheet)
    const next = extendedRange(range, anchor, corner)
    return setRange(accessor, next, farCorner(next, anchor, corner))
  }

  const expandRowHome: ICommand = {
    id: EXPAND_ROW_HOME_ID,
    type: CommandType.COMMAND,
    handler: (accessor) =>
      extendTo(accessor, (worksheet) => ({ row: null, column: unfrozenOrigin(worksheet).column })),
  }
  const expandSheetHome: ICommand = {
    id: EXPAND_SHEET_HOME_ID,
    type: CommandType.COMMAND,
    handler: (accessor) => extendTo(accessor, unfrozenOrigin),
  }
  const expandSheetEnd: ICommand = {
    id: EXPAND_SHEET_END_ID,
    type: CommandType.COMMAND,
    handler: (accessor) => extendTo(accessor, usedEnd),
  }

  const selectCurrentRegion: ICommand = {
    id: SELECT_CURRENT_REGION_ID,
    type: CommandType.COMMAND,
    handler: (accessor) => {
      const target = getSheetCommandTarget(accessor.get(IUniverInstanceService))
      const selection = accessor.get(SheetsSelectionsService).getCurrentLastSelection()
      if (!target || !selection) return false
      const { worksheet } = target
      const matrix = worksheet.getCellMatrix()
      return setRange(
        accessor,
        currentRegion(
          selection.range,
          (row, column) => excelCellHasValue(matrix.getValue(row, column) ?? null),
          worksheet.getRowCount(),
          worksheet.getColumnCount(),
        ),
      )
    },
  }

  for (const command of [
    expandRowHome,
    expandSheetHome,
    expandSheetEnd,
    selectCurrentRegion,
    activateAdjacentSheet,
    selectSheetHome,
    selectSheetEnd,
    selectRowHome,
    selectWholeColumn,
    selectWholeRow,
    hideSelectedRows,
    unhideSelectedRows,
    hideSelectedCols,
    unhideSelectedCols,
    insertCells,
    deleteCells,
  ]) {
    commandService.registerCommand(command)
  }

  const items: IShortcutItem[] = [
    {
      id: ACTIVATE_ADJACENT_SHEET_ID,
      binding: KEY_PAGE_DOWN | MetaKeys.CTRL_COMMAND,
      preconditions: whenSheetEditorFocused,
      staticParameters: { step: 1 },
    },
    {
      id: ACTIVATE_ADJACENT_SHEET_ID,
      binding: KEY_PAGE_UP | MetaKeys.CTRL_COMMAND,
      preconditions: whenSheetEditorFocused,
      staticParameters: { step: -1 },
    },
    // Excel-for-mac also switches tabs with Option+arrows; the zero base
    // binding makes these mac-only (the dispatcher skips falsy bindings)
    {
      id: ACTIVATE_ADJACENT_SHEET_ID,
      binding: 0,
      mac: KeyCode.ARROW_RIGHT | MetaKeys.ALT,
      preconditions: whenSheetEditorFocused,
      staticParameters: { step: 1 },
    },
    {
      id: ACTIVATE_ADJACENT_SHEET_ID,
      binding: 0,
      mac: KeyCode.ARROW_LEFT | MetaKeys.ALT,
      preconditions: whenSheetEditorFocused,
      staticParameters: { step: -1 },
    },
    {
      id: SELECT_SHEET_HOME_ID,
      binding: KEY_HOME | MetaKeys.CTRL_COMMAND,
      preconditions: whenSheetEditorFocused,
    },
    {
      id: SELECT_SHEET_END_ID,
      binding: KEY_END | MetaKeys.CTRL_COMMAND,
      preconditions: whenSheetEditorFocused,
    },
    {
      id: SELECT_ROW_HOME_ID,
      binding: KEY_HOME,
      preconditions: whenSheetEditorFocused,
    },
    // ⌘Space is Spotlight — Excel for mac uses real Ctrl+Space too
    {
      id: SELECT_WHOLE_COLUMN_ID,
      binding: KeyCode.SPACE | MetaKeys.CTRL_COMMAND,
      mac: KeyCode.SPACE | MetaKeys.MAC_CTRL,
      preconditions: whenSheetEditorFocused,
    },
    {
      id: SELECT_WHOLE_ROW_ID,
      binding: KeyCode.SPACE | MetaKeys.SHIFT,
      preconditions: whenSheetEditorFocused,
    },
    // Excel's hide/unhide: Ctrl+9/0 rows/columns, +Shift restores. Upstream
    // binds Ctrl+9 / Ctrl+Shift+0 to the header-selection-only commands
    // (no-ops on cell selections) and Ctrl+0 to zoom reset (priority 1) —
    // outrank them all; Excel semantics win on these keys.
    {
      id: HIDE_SELECTED_ROWS_ID,
      priority: 100,
      binding: KeyCode.Digit9 | MetaKeys.CTRL_COMMAND,
      preconditions: whenSheetEditorFocused,
    },
    {
      id: UNHIDE_SELECTED_ROWS_ID,
      priority: 100,
      binding: KeyCode.Digit9 | MetaKeys.CTRL_COMMAND | MetaKeys.SHIFT,
      preconditions: whenSheetEditorFocused,
    },
    {
      id: HIDE_SELECTED_COLS_ID,
      priority: 100,
      binding: KeyCode.Digit0 | MetaKeys.CTRL_COMMAND,
      preconditions: whenSheetEditorFocused,
    },
    {
      id: UNHIDE_SELECTED_COLS_ID,
      priority: 100,
      binding: KeyCode.Digit0 | MetaKeys.CTRL_COMMAND | MetaKeys.SHIFT,
      preconditions: whenSheetEditorFocused,
    },
    // Excel's Insert/Delete Cells. Upstream binds Ctrl+= / Ctrl+- to zoom
    // (priority 1); Ctrl+- goes to Excel semantics, plain Ctrl+= stays zoom-in.
    {
      id: INSERT_CELLS_ID,
      priority: 100,
      binding: KeyCode.EQUAL | MetaKeys.CTRL_COMMAND | MetaKeys.SHIFT,
      preconditions: whenSheetEditorFocused,
    },
    {
      id: INSERT_CELLS_ID,
      priority: 100,
      binding: KEY_NUMPAD_ADD | MetaKeys.CTRL_COMMAND,
      preconditions: whenSheetEditorFocused,
    },
    {
      id: DELETE_CELLS_ID,
      priority: 100,
      binding: KeyCode.MINUS | MetaKeys.CTRL_COMMAND,
      preconditions: whenSheetEditorFocused,
    },
    {
      id: DELETE_CELLS_ID,
      priority: 100,
      binding: KEY_NUMPAD_SUBTRACT | MetaKeys.CTRL_COMMAND,
      preconditions: whenSheetEditorFocused,
    },
    {
      id: EXPAND_ROW_HOME_ID,
      binding: KEY_HOME | MetaKeys.SHIFT,
      preconditions: whenSheetEditorFocused,
    },
    {
      id: EXPAND_SHEET_HOME_ID,
      binding: KEY_HOME | MetaKeys.CTRL_COMMAND | MetaKeys.SHIFT,
      preconditions: whenSheetEditorFocused,
    },
    {
      id: EXPAND_SHEET_END_ID,
      binding: KEY_END | MetaKeys.CTRL_COMMAND | MetaKeys.SHIFT,
      preconditions: whenSheetEditorFocused,
    },
    {
      id: SELECT_CURRENT_REGION_ID,
      binding: KeyCode.Digit8 | MetaKeys.CTRL_COMMAND | MetaKeys.SHIFT,
      preconditions: whenSheetEditorFocused,
    },
    {
      id: SetRangeBoldCommand.id,
      binding: KeyCode.Digit2 | MetaKeys.CTRL_COMMAND,
      preconditions: whenSheetEditorFocused,
    },
    {
      id: SetRangeItalicCommand.id,
      binding: KeyCode.Digit3 | MetaKeys.CTRL_COMMAND,
      preconditions: whenSheetEditorFocused,
    },
    {
      id: SetRangeUnderlineCommand.id,
      binding: KeyCode.Digit4 | MetaKeys.CTRL_COMMAND,
      preconditions: whenSheetEditorFocused,
    },
    {
      id: SetCellEditVisibleOperation.id,
      binding: KeyCode.ENTER | MetaKeys.CTRL_COMMAND | MetaKeys.SHIFT,
      preconditions: whenCellEditorNotFormulaActivated,
      staticParameters: {
        visible: false,
        eventType: DeviceInputEventType.Keyboard,
        keycode: KeyCode.ENTER | MetaKeys.CTRL_COMMAND,
      },
    },
    // Upstream's ⌃Enter precondition (whenEditorDidNotInputFormulaActivated)
    // is not exported; same context checks.
    {
      id: SetCellEditVisibleOperation.id,
      binding: 0,
      mac: KeyCode.ENTER | MetaKeys.CTRL_COMMAND,
      preconditions: whenCellEditorNotFormulaActivated,
      staticParameters: {
        visible: false,
        eventType: DeviceInputEventType.Keyboard,
        keycode: KeyCode.ENTER | MetaKeys.CTRL_COMMAND,
      },
    },
    {
      id: BreakLineCommand.id,
      binding: 0,
      mac: KeyCode.ENTER | MetaKeys.ALT | MetaKeys.MAC_CTRL,
      preconditions: whenSheetEditorActivated,
    },
    {
      id: BreakLineCommand.id,
      binding: 0,
      mac: KeyCode.ENTER | MetaKeys.ALT | MetaKeys.CTRL_COMMAND,
      preconditions: whenSheetEditorActivated,
    },
    {
      id: ReferenceAbsoluteOperation.id,
      binding: 0,
      mac: KeyCode.T | MetaKeys.CTRL_COMMAND,
      preconditions: whenSheetEditorActivated,
    },
  ]
  for (const item of items) shortcutService.registerShortcut(item)
}

function whenCellEditorNotFormulaActivated(contextService: IContextService): boolean {
  return (
    contextService.getContextValue(FOCUSING_SHEET) &&
    contextService.getContextValue(FOCUSING_UNIVER_EDITOR) &&
    contextService.getContextValue(EDITOR_ACTIVATED) &&
    !contextService.getContextValue(FOCUSING_EDITOR_INPUT_FORMULA) &&
    !contextService.getContextValue(FOCUSING_EDITOR_STANDALONE)
  )
}

/** exported for tests */
export const _shortcutInternals = {
  ACTIVATE_ADJACENT_SHEET_ID,
  SELECT_SHEET_HOME_ID,
  SELECT_SHEET_END_ID,
  SELECT_ROW_HOME_ID,
  SELECT_WHOLE_COLUMN_ID,
  SELECT_WHOLE_ROW_ID,
  KEY_PAGE_UP,
  KEY_PAGE_DOWN,
  KEY_END,
  KEY_HOME,
  KEY_NUMPAD_ADD,
  KEY_NUMPAD_SUBTRACT,
  INSERT_CELLS_ID,
  DELETE_CELLS_ID,
  EXPAND_ROW_HOME_ID,
  EXPAND_SHEET_HOME_ID,
  EXPAND_SHEET_END_ID,
  SELECT_CURRENT_REGION_ID,
  unfrozenOrigin,
}
