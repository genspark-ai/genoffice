/**
 * Excel-style worksheet protection in the editor. The file's
 * `<sheetProtection>` (or the session's desired state) gates Univer commands
 * before they run: content edits are blocked on locked cells, every other
 * command maps to one of the dialog's permission flags.
 */
import type { IRange } from '@univerjs/core'
import { SheetsSelectionsService } from '@univerjs/sheets'
import {
  DEFAULT_SHEET_PROTECTION_ALLOW,
  type SheetProtectionAllow,
  type SheetProtectionPermission,
} from '@genoffice/xlsx-gateway/gateway/xlsx-protection'
import type { SheetPasswordHash } from '@genoffice/xlsx-gateway/gateway/xlsx-protection-hash'

import { t } from './i18n/locale'
import { journalSuppression, type LazyWorkbookState, type UniverRuntime } from './univer-state'
import type { WorkbookRangeResult } from '../shared/desktop-api'

/// ICellData.custom keys carrying the xf protection flags.
export const CELL_LOCKED_KEY = 'xlsxLocked'
export const CELL_HIDDEN_KEY = 'xlsxHidden'

export interface SheetProtectionInfo {
  readonly protected: boolean
  readonly hasPassword: boolean
  readonly allow: SheetProtectionAllow
  readonly password: SheetPasswordHash | null
}

export const UNPROTECTED_SHEET: SheetProtectionInfo = {
  protected: false,
  hasPassword: false,
  allow: DEFAULT_SHEET_PROTECTION_ALLOW,
  password: null,
}

/// Normalizes the sidecar record (older binaries omit allow/password).
export function fileSheetProtection(
  raw: WorkbookRangeResult['sheetProtection'],
): SheetProtectionInfo {
  if (!raw) return UNPROTECTED_SHEET
  return {
    protected: raw.protected,
    hasPassword: raw.hasPassword,
    allow: { ...DEFAULT_SHEET_PROTECTION_ALLOW, ...(raw.allow ?? {}) },
    password: raw.password ?? null,
  }
}

/// Dialog/AI partial → full flag set (unlisted = not allowed, as in Excel).
export function mergeSheetProtectionAllow(
  partial: Partial<Record<SheetProtectionPermission, boolean | undefined>> | undefined,
): SheetProtectionAllow {
  const allow = { ...DEFAULT_SHEET_PROTECTION_ALLOW }
  for (const key of Object.keys(allow) as SheetProtectionPermission[]) {
    const value = partial?.[key]
    if (typeof value === 'boolean') allow[key] = value
  }
  return allow
}

/// Journal override, else the file; null while unknown (still indexing).
export function effectiveSheetProtection(
  state: LazyWorkbookState,
  sheetId: string,
): SheetProtectionInfo | null {
  const journaled = state.editJournal.sheetProtection.get(sheetId)
  if (journaled) return journaled
  const file = state.sheetProtections.get(sheetId)
  if (file) return file
  return state.editJournal.sheets.added.has(sheetId) ? UNPROTECTED_SHEET : null
}

export function cellIsLocked(
  cell: { custom?: unknown } | null | undefined,
  defaultLocked: boolean,
): boolean {
  const custom = cell?.custom as Record<string, unknown> | null | undefined
  const flag = custom?.[CELL_LOCKED_KEY]
  return typeof flag === 'boolean' ? flag : defaultLocked
}

/// 'edit' = cell content change (checked per target cell); everything else
/// is one of the dialog's permission flags.
export type GuardKind = 'edit' | SheetProtectionPermission

const EDIT_COMMANDS = [
  'sheet.operation.set-cell-edit-visible',
  'sheet.command.set-range-values',
  'sheet.command.clear-selection-content',
  'sheet.command.clear-selection-all',
  'sheet.command.auto-clear-content',
  'sheet.command.paste',
  'sheet.command.paste-by-short-key',
  'sheet.command.paste-value',
  'sheet.command.paste-formula',
  'sheet.command.paste-besides-border',
  'sheet.command.optional-paste',
  'sheet.command.cut',
  'sheet.command.auto-fill',
  'sheet.command.refill',
  'sheet.command.copy-down',
  'sheet.command.copy-right',
  'sheet.command.move-range',
  'sheet.command.split-text-to-columns',
  'sheet.command.text-to-number',
  'sheet.command.toggle-cell-checkbox',
  'sheet.command.set-range-custom-metadata',
]

const PERMISSION_COMMANDS: Record<SheetProtectionPermission, readonly string[]> = {
  selectLockedCells: [],
  selectUnlockedCells: [],
  formatCells: [
    'sheet.command.set-style',
    'sheet.command.set-range-bold',
    'sheet.command.set-range-italic',
    'sheet.command.set-range-underline',
    'sheet.command.set-range-stroke',
    'sheet.command.set-range-subscript',
    'sheet.command.set-range-superscript',
    'sheet.command.set-range-font-family',
    'sheet.command.set-range-fontsize',
    'sheet.command.set-range-font-increase',
    'sheet.command.set-range-font-decrease',
    'sheet.command.set-range-text-color',
    'sheet.command.reset-range-text-color',
    'sheet.command.set-bold',
    'sheet.command.set-italic',
    'sheet.command.set-underline',
    'sheet.command.set-stroke',
    'sheet.command.set-overline',
    'sheet.command.set-font-family',
    'sheet.command.set-font-size',
    'sheet.command.set-text-color',
    'sheet.command.reset-text-color',
    'sheet.command.set-background-color',
    'sheet.command.reset-background-color',
    'sheet.command.set-horizontal-text-align',
    'sheet.command.set-vertical-text-align',
    'sheet.command.set-text-wrap',
    'sheet.command.set-text-rotation',
    'sheet.command.set-border',
    'sheet.command.set-border-basic',
    'sheet.command.set-border-position',
    'sheet.command.set-border-style',
    'sheet.command.set-border-color',
    'sheet.command.clear-selection-format',
    'sheet.command.paste-format',
    'sheet.command.apply-format-painter',
    'sheet.command.add-worksheet-merge',
    'sheet.command.add-worksheet-merge-all',
    'sheet.command.add-worksheet-merge-horizontal',
    'sheet.command.add-worksheet-merge-vertical',
    'sheet.command.remove-worksheet-merge',
    'sheet.command.set-numfmt',
    'sheet.command.remove-numfmt',
    'sheet.command.add-conditional-rule',
    'sheet.command.set-conditional-rule',
    'sheet.command.delete-conditional-rule',
    'sheet.command.clear-range-conditional-rule',
    'sheet.command.clear-worksheet-conditional-rule',
  ],
  formatColumns: [
    'sheet.command.set-worksheet-col-width',
    'sheet.command.delta-column-width',
    'sheet.command.set-col-auto-width',
    'sheet.command.set-col-is-auto-width',
    'sheet.command.set-col-hidden',
    'sheet.command.hide-col-confirm',
    'sheet.command.set-col-visible-on-cols',
    'sheet.command.set-selected-cols-visible',
    'sheet.command.paste-col-width',
  ],
  formatRows: [
    'sheet.command.set-row-height',
    'sheet.command.delta-row-height',
    'sheet.command.set-row-is-auto-height',
    'sheet.command.set-rows-hidden',
    'sheet.command.hide-row-confirm',
    'sheet.command.set-specific-rows-visible',
    'sheet.command.set-selected-rows-visible',
  ],
  insertColumns: [
    'sheet.command.insert-col',
    'sheet.command.insert-col-before',
    'sheet.command.insert-col-after',
    'sheet.command.insert-col-by-range',
    'sheet.command.insert-multi-cols-before',
    'sheet.command.insert-multi-cols-right',
    'sheet.command.insert-range-move-right-confirm',
    'sheet.command.table-insert-col',
    'sheet.command.table-insert-column-at',
  ],
  insertRows: [
    'sheet.command.insert-row',
    'sheet.command.insert-row-before',
    'sheet.command.insert-row-after',
    'sheet.command.insert-row-by-range',
    'sheet.command.insert-multi-rows-above',
    'sheet.command.insert-multi-rows-after',
    'sheet.command.insert-range-move-down-confirm',
    'sheet.command.table-insert-row',
    'sheet.command.table-insert-row-at',
  ],
  insertHyperlinks: [
    'sheet.command.add-hyper-link',
    'sheet.command.update-hyper-link',
    'sheet.command.remove-hyper-link',
    'sheet.command.cancel-hyper-link',
  ],
  deleteColumns: [
    'sheet.command.remove-col',
    'sheet.command.remove-col-confirm',
    'sheet.command.remove-col-by-range',
    'sheet.command.delete-range-move-left-confirm',
    'sheet.command.table-remove-col',
    'sheet.command.table-remove-column-at',
  ],
  deleteRows: [
    'sheet.command.remove-row',
    'sheet.command.remove-row-confirm',
    'sheet.command.remove-row-by-range',
    'sheet.command.delete-range-move-up-confirm',
    'sheet.command.table-remove-row',
  ],
  sort: [
    'sheet.command.sort-range',
    'sheet.command.sort-range-asc',
    'sheet.command.sort-range-asc-ext',
    'sheet.command.sort-range-asc-ctx',
    'sheet.command.sort-range-asc-ext-ctx',
    'sheet.command.sort-range-desc',
    'sheet.command.sort-range-desc-ext',
    'sheet.command.sort-range-desc-ctx',
    'sheet.command.sort-range-desc-ext-ctx',
    'sheet.command.sort-range-custom',
    'sheet.command.sort-range-custom-ctx',
  ],
  autoFilter: [
    'sheet.command.smart-toggle-filter',
    'sheet.command.set-filter-range',
    'sheet.command.remove-sheet-filter',
    'sheet.command.set-filter-criteria',
    'sheet.command.clear-filter-criteria',
    'sheet.command.re-calc-filter',
    'sheet.operation.open-filter-panel',
  ],
  pivotTables: [],
  objects: [],
  scenarios: [],
}

const COMMAND_KINDS: ReadonlyMap<string, GuardKind> = (() => {
  const map = new Map<string, GuardKind>()
  for (const id of EDIT_COMMANDS) map.set(id, 'edit')
  for (const [permission, ids] of Object.entries(PERMISSION_COMMANDS)) {
    for (const id of ids) map.set(id, permission as SheetProtectionPermission)
  }
  return map
})()

export function commandGuardKind(commandId: string): GuardKind | undefined {
  return COMMAND_KINDS.get(commandId)
}

const CONTENT_KEYS = ['v', 'f', 'si', 'p', 't']

/// A set-range-values payload that only touches `s`/`custom` is formatting.
export function setRangeValuesChangesContent(value: unknown): boolean {
  if (typeof value !== 'object' || value === null) return false
  const cells: unknown[] = []
  const data = value as Record<string, unknown>
  if (CONTENT_KEYS.some((key) => key in data) || 's' in data || 'custom' in data) {
    cells.push(data)
  } else {
    for (const row of Object.values(data)) {
      if (Array.isArray(row)) cells.push(...row)
      else if (typeof row === 'object' && row !== null) cells.push(...Object.values(row))
    }
  }
  return cells.some(
    (cell) =>
      cell === null ||
      (typeof cell === 'object' &&
        CONTENT_KEYS.some((key) => key in (cell as Record<string, unknown>))),
  )
}

function isRange(value: unknown): value is IRange {
  if (typeof value !== 'object' || value === null) return false
  const range = value as Record<string, unknown>
  return (
    typeof range.startRow === 'number' &&
    typeof range.endRow === 'number' &&
    typeof range.startColumn === 'number' &&
    typeof range.endColumn === 'number'
  )
}

/// Ranges named in the params (range/ranges/targetRange/toRange…), else none.
export function rangesInParams(params: unknown): IRange[] {
  if (typeof params !== 'object' || params === null) return []
  const out: IRange[] = []
  for (const value of Object.values(params as Record<string, unknown>)) {
    if (isRange(value)) out.push(value)
    else if (Array.isArray(value)) for (const item of value) if (isRange(item)) out.push(item)
  }
  return out
}

/// Cells a set-range-values matrix addresses; null when it is not a matrix.
function matrixRanges(value: unknown): IRange[] | null {
  if (typeof value !== 'object' || value === null) return null
  const data = value as Record<string, unknown>
  if (CONTENT_KEYS.some((key) => key in data) || 's' in data || 'custom' in data) return null
  const out: IRange[] = []
  for (const [rowKey, row] of Object.entries(data)) {
    const rowIndex = Number(rowKey)
    if (!Number.isInteger(rowIndex) || typeof row !== 'object' || row === null) continue
    for (const columnKey of Object.keys(row as Record<string, unknown>)) {
      const columnIndex = Number(columnKey)
      if (!Number.isInteger(columnIndex)) continue
      out.push({
        startRow: rowIndex,
        endRow: rowIndex,
        startColumn: columnIndex,
        endColumn: columnIndex,
      })
    }
  }
  return out
}

function parseA1(ref: string): IRange | null {
  const match = /^([A-Z]{1,3})(\d+)(?::([A-Z]{1,3})(\d+))?$/.exec(
    ref.toUpperCase().replaceAll('$', ''),
  )
  if (!match) return null
  const column = (label: string) =>
    [...label].reduce((acc, char) => acc * 26 + (char.charCodeAt(0) - 64), 0) - 1
  const startColumn = column(match[1]!)
  const startRow = Number(match[2]) - 1
  const endColumn = match[3] ? column(match[3]) : startColumn
  const endRow = match[4] ? Number(match[4]) - 1 : startRow
  return { startRow, endRow, startColumn, endColumn }
}

function rangeContains(outer: IRange, row: number, column: number): boolean {
  return (
    row >= outer.startRow &&
    row <= outer.endRow &&
    column >= outer.startColumn &&
    column <= outer.endColumn
  )
}

const MAX_CELLS_CHECKED = 200_000

export interface LockedCellLookup {
  readonly getCell: (row: number, column: number) => { custom?: unknown } | null | undefined
  readonly defaultLocked: boolean
  /// Allow-edit ranges (A1 sqrefs, space separated) stay editable.
  readonly allowEditRefs: readonly string[]
  readonly maxRow: number
  readonly maxColumn: number
  /// Rows outside the resident window carry no row-xf flags yet; null = all
  /// rows resident. Frozen rows stay installed on their own.
  readonly residentRows: { readonly startRow: number; readonly endRow: number } | null
  readonly frozenRows: number
}

/// True when any cell of the ranges is locked. Whole-row/column selections
/// are clipped to the sheet; beyond the scan budget the range counts as
/// locked (fail closed).
export function rangesTouchLockedCell(
  ranges: readonly IRange[],
  lookup: LockedCellLookup,
): boolean {
  const allowEdit = lookup.allowEditRefs
    .flatMap((sqref) => sqref.split(/\s+/))
    .map(parseA1)
    .filter((range): range is IRange => range !== null)
  let budget = MAX_CELLS_CHECKED
  for (const range of ranges) {
    const endRow = Math.min(range.endRow, lookup.maxRow)
    const endColumn = Math.min(range.endColumn, lookup.maxColumn)
    for (let row = Math.max(0, range.startRow); row <= endRow; row += 1) {
      for (let column = Math.max(0, range.startColumn); column <= endColumn; column += 1) {
        if (budget-- <= 0) return true
        if (allowEdit.some((area) => rangeContains(area, row, column))) continue
        const resident =
          lookup.residentRows === null ||
          row < lookup.frozenRows ||
          (row >= lookup.residentRows.startRow && row <= lookup.residentRows.endRow)
        if (!resident) return true
        if (cellIsLocked(lookup.getCell(row, column), lookup.defaultLocked)) return true
      }
    }
  }
  return false
}

export interface SheetProtectionGuardDeps {
  readonly getState: () => LazyWorkbookState | null
  readonly notify: (message: string) => void
}

/// Vetoes commands the active sheet's protection forbids. Loader writes run
/// under journalSuppression and pass through.
export function installSheetProtectionGuard(
  runtime: UniverRuntime,
  deps: SheetProtectionGuardDeps,
): { dispose(): void } {
  const injector = runtime.univer.__getInjector()
  return runtime.univerAPI.addEvent(runtime.univerAPI.Event.BeforeCommandExecute, (event) => {
    const kind = commandGuardKind(event.id)
    if (!kind || journalSuppression.active) return
    const state = deps.getState()
    const workbook = runtime.univerAPI.getActiveWorkbook()
    const params = (event.params ?? {}) as Record<string, unknown>
    // Commands addressed to a background sheet are checked against it.
    const worksheet =
      typeof params.subUnitId === 'string'
        ? workbook?.getSheetBySheetId(params.subUnitId)
        : workbook?.getActiveSheet()
    if (!state || !worksheet) return
    const protection = effectiveSheetProtection(state, worksheet.getSheetId())
    // Unknown until the sheet is indexed: fail closed rather than let edits
    // land on what may turn out to be a protected sheet.
    if (!protection) {
      event.cancel = true
      deps.notify(t('appProtectionNeedsIndexed'))
      return
    }
    if (!protection.protected) return

    if (kind !== 'edit') {
      if (protection.allow[kind]) return
      event.cancel = true
      deps.notify(t('appProtectedSheetCommandBlocked'))
      return
    }

    if (event.id === 'sheet.operation.set-cell-edit-visible' && params.visible !== true) return
    let formattingOnly = false
    let ranges: IRange[] = []
    if (event.id === 'sheet.command.set-range-values') {
      formattingOnly = !setRangeValuesChangesContent(params.value)
      ranges = isRange(params.range) ? [params.range] : (matrixRanges(params.value) ?? [])
    } else if (event.id === 'sheet.command.set-range-custom-metadata') {
      formattingOnly = true
    } else {
      ranges = rangesInParams(params)
    }
    if (formattingOnly) {
      if (protection.allow.formatCells) return
      event.cancel = true
      deps.notify(t('appProtectedSheetCommandBlocked'))
      return
    }
    if (ranges.length === 0) {
      ranges = injector
        .get(SheetsSelectionsService)
        .getCurrentSelections()
        .map((selection) => selection.range)
    }
    if (ranges.length === 0) return
    const sheet = worksheet.getSheet()
    // Excel precedence for a cell without its own flag: row xf, column xf,
    // then the Normal style.
    const hasFlag = (data: unknown) =>
      typeof (data as { custom?: Record<string, unknown> | null } | null | undefined)?.custom?.[
        CELL_LOCKED_KEY
      ] === 'boolean'
    const locked = rangesTouchLockedCell(ranges, {
      getCell: (row, column) => {
        const cell = sheet.getCellRaw(row, column) ?? null
        if (hasFlag(cell)) return cell
        const rowData = sheet.getRowManager().getRow(row) ?? null
        if (hasFlag(rowData)) return rowData
        return sheet.getColumnManager().getColumn(column) ?? null
      },
      defaultLocked: state.file.styles[0]?.locked !== false,
      allowEditRefs: (state.sheetProtectedRanges.get(worksheet.getSheetId()) ?? []).map(
        (range) => range.sqref,
      ),
      maxRow: sheet.getRowCount() - 1,
      maxColumn: sheet.getColumnCount() - 1,
      residentRows:
        state.formulaMode && state.flags.preloadComplete
          ? null
          : (state.loadedRanges.get(worksheet.getSheetId()) ?? { startRow: -1, endRow: -1 }),
      frozenRows: sheet.getFreeze().ySplit,
    })
    if (!locked) return
    event.cancel = true
    deps.notify(t('appProtectedCellBlocked'))
  })
}
