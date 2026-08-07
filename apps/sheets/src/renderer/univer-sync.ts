/**
 * Univer runtime synchronization helpers for the sheets renderer.
 *
 * Module-level functions that translate between the workbook file model
 * (snapshots, edit journal, lazy streaming state) and the live Univer
 * spreadsheet instance. Extracted from App.tsx; they hold no React state.
 */
import {
  BooleanNumber,
  BorderStyleTypes,
  CellValueType,
  CommandType,
  HorizontalAlign,
  ICommandService,
  IUndoRedoService,
  VerticalAlign,
  WrapStrategy,
  type ICellData,
  type IRange,
  type IStyleData,
} from '@univerjs/core'
import { CFValueType, type IValueConfig } from '@univerjs/preset-sheets-conditional-formatting'

import type {
  AddConditionalFormatOperation,
  CellFormatPatch,
  SetDataValidationOperation,
  SetHyperlinkOperation,
} from '../domain/workbook-dsl'
import {
  columnIndex,
  columnLabel,
  parseAddress,
  parseRange,
  rangeCellCount,
} from '../domain/cell-address'
import { splitSheetRef, type CellBounds } from '../domain/chart-visual'
import { InMemoryWorkbookAdapter } from '../domain/in-memory-workbook'
import { WORST_FIRST_ICON_SETS } from '../gateway/xlsx-cf'
import type { CellFormatState, CellState, WorkbookSnapshot } from '../domain/workbook.types'
import type {
  WorkbookCellStyle,
  WorkbookCfState,
  WorkbookChartEdit,
  WorkbookDvState,
  WorkbookFile,
  WorkbookFilterState,
  WorkbookNoteState,
  WorkbookRangeResult,
  WorkbookRichRun,
  WorkbookVisualObject,
} from '../shared/desktop-api'
import {
  fromNeutralStyle,
  isSheetRemoved,
  journalEntriesInRange,
  recordHyperlinkEdit,
  recordSetRangeValues,
  toRecalcUserInput,
  type EditJournal,
  type VisualEditEntry,
} from './edit-journal'
import {
  cellKey,
  closureFetchRanges,
  computeFormulaClosure,
  recalcReadRanges,
  type ClosureSheetInput,
} from './formula-closure'
import { t } from './i18n/locale'
import { INDENT_STEP_PX } from './selection-format'
import {
  fileRangeToScreenRange,
  indexedThroughScreenRow,
  mapRangeResultToScreen,
  netAxisDelta,
  screenRangeToFileRange,
} from './view-transform'
import {
  buildCustomFilters,
  type AdvancedFilterColumn,
  type AdvancedFilterCondition,
} from './AdvancedFilterDialog'
import {
  installSparklines,
  installWorkbookVisuals,
  isChartEditorOpen,
  isVisualDragActive,
  type ChartEditData,
  type ChartVectorRead,
  type ShapeEditChanges,
  type SparklineGroupState,
} from './WorkbookVisuals'
import {
  BORDER_COMMAND_TYPES,
  CLOSURE_MAX_CELLS,
  journalSuppression,
  type LazyWorkbookState,
  type PinnedClosureCell,
  type UniverRuntime,
  type UniverWorksheet,
} from './univer-state'

export const MINIMUM_SHEET_ROW_COUNT = 1000
export const MINIMUM_SHEET_COLUMN_COUNT = 26

export function syncUniver(runtime: UniverRuntime | null, snapshot: WorkbookSnapshot): void {
  const workbook = runtime?.univerAPI.getActiveWorkbook()
  if (!workbook) return
  for (const sheet of snapshot.sheets) {
    const worksheet = workbook.getSheetBySheetId(sheet.id)
    if (!worksheet) continue
    worksheet.setName(sheet.name)
    for (const [address, cell] of Object.entries(sheet.cells)) {
      const range = worksheet.getRange(address)
      if (cell.formula) range.setFormula(cell.formula)
      else if (cell.value === null) range.clearContent()
      else range.setValue(cell.value)
    }
  }
}

export function loadSnapshotIntoUniver(
  runtime: UniverRuntime | null,
  snapshot: WorkbookSnapshot,
  workbookId: string,
  workbookName: string,
): void {
  if (!runtime) return
  const activeWorkbook = runtime.univerAPI.getActiveWorkbook()
  if (activeWorkbook) runtime.univerAPI.disposeUnit(activeWorkbook.getId())

  runtime.univerAPI.createWorkbook({
    id: workbookId,
    name: workbookName,
    sheetOrder: snapshot.sheets.map((sheet) => sheet.id),
    sheets: Object.fromEntries(
      snapshot.sheets.map((sheet) => {
        const cellData: Record<
          number,
          Record<number, { v?: string | number | boolean; f?: string }>
        > = {}
        let maximumRow = 0
        let maximumColumn = 0
        for (const [address, cell] of Object.entries(sheet.cells)) {
          const coordinates = parseAddress(address)
          maximumRow = Math.max(maximumRow, coordinates.row)
          maximumColumn = Math.max(maximumColumn, coordinates.column)
          const rowData = cellData[coordinates.row] ?? {}
          rowData[coordinates.column] = cell.formula
            ? { f: cell.formula }
            : cell.value === null
              ? {}
              : { v: cell.value }
          cellData[coordinates.row] = rowData
        }
        return [
          sheet.id,
          {
            id: sheet.id,
            name: sheet.name,
            rowCount: Math.max(MINIMUM_SHEET_ROW_COUNT, maximumRow + 100),
            columnCount: Math.max(MINIMUM_SHEET_COLUMN_COUNT, maximumColumn + 10),
            cellData,
          },
        ]
      }),
    ),
  })

  // Replay demo-mode formatting and layout after the rebuild (snapshot is
  // the source of truth; cellData above carries only values/formulas).
  const workbook = runtime.univerAPI.getActiveWorkbook()
  for (const sheet of snapshot.sheets) {
    const worksheet = workbook?.getSheetBySheetId(sheet.id)
    if (!worksheet) continue
    for (const [address, style] of Object.entries(sheet.styles ?? {})) {
      applyFormatPatchToRange(worksheet.getRange(address), style)
    }
    for (const merge of sheet.merges ?? []) {
      worksheet.getRange(merge).merge()
    }
    for (const [row, heightPoints] of Object.entries(sheet.rowHeights ?? {})) {
      worksheet.setRowHeights(Number(row) - 1, 1, Math.round((heightPoints * 96) / 72))
    }
    for (const [column, widthPx] of Object.entries(sheet.colWidths ?? {})) {
      worksheet.setColumnWidths(columnIndex(column), 1, Math.round(widthPx))
    }
  }
}

/// Excel's data-field caption prefixes for the baked pivot header row.
export const AGG_CAPTIONS: Record<'sum' | 'count' | 'average' | 'max' | 'min', string> = {
  sum: 'Sum',
  count: 'Count',
  average: 'Average',
  max: 'Max',
  min: 'Min',
}

/// Default session pivot names, mirroring nextSessionTableName.
export function nextSessionPivotName(journal: EditJournal): string {
  const taken = new Set(journal.pivotAdds.map((pivot) => pivot.name.toLowerCase()))
  let index = journal.pivotAdds.length + 1
  while (taken.has(`pivot${index}`)) index += 1
  return `Pivot${index}`
}

/// Default session table names: Table1, Table2, … skipping names the session
/// already used. Collisions with names in the file fail closed at save time.
export function nextSessionTableName(journal: EditJournal): string {
  const taken = new Set(journal.tableAdds.map((table) => table.name.toLowerCase()))
  let index = journal.tableAdds.length + 1
  while (taken.has(`table${index}`)) index += 1
  return `Table${index}`
}

/// Shared by demo replay and lazy Apply: pushes one format patch through the
/// same facade setters the ribbon uses. null (or a missing field in demo
/// CellFormatState) clears back to the default.
export function applyFormatPatchToRange(
  range: ReturnType<UniverWorksheet['getRange']>,
  format: CellFormatPatch | CellFormatState,
): void {
  const patch = format as CellFormatPatch
  if (patch.bold !== undefined) range.setFontWeight(patch.bold ? 'bold' : null)
  if (patch.italic !== undefined) range.setFontStyle(patch.italic ? 'italic' : null)
  if (patch.underline !== undefined) range.setFontLine(patch.underline ? 'underline' : null)
  if (patch.strikethrough !== undefined) {
    // setFontLine would overwrite the underline key; patch st directly.
    range.setValue({
      s: { st: patch.strikethrough ? { s: BooleanNumber.TRUE } : null },
    } as unknown as ICellData)
  }
  if (patch.fontFamily !== undefined) {
    if (patch.fontFamily === null) range.setValue({ s: { ff: null } } as unknown as ICellData)
    else range.setFontFamily(patch.fontFamily)
  }
  if (patch.fontSize !== undefined) {
    if (patch.fontSize === null) range.setValue({ s: { fs: null } } as unknown as ICellData)
    else range.setFontSize(patch.fontSize)
  }
  if (patch.fontColor !== undefined) range.setFontColor(patch.fontColor)
  if (patch.fillColor !== undefined) range.setBackground(patch.fillColor as unknown as string)
  if (patch.numberFormat !== undefined) range.setNumberFormat(patch.numberFormat ?? 'General')
  if (patch.horizontalAlign !== undefined) {
    range.setHorizontalAlignment(
      (patch.horizontalAlign ?? 'normal') as 'left' | 'center' | 'normal',
    )
  }
  if (patch.verticalAlign !== undefined) {
    if (patch.verticalAlign === null) range.setValue({ s: { vt: null } } as unknown as ICellData)
    else
      range.setVerticalAlignment(patch.verticalAlign === 'center' ? 'middle' : patch.verticalAlign)
  }
  if (patch.wrapText !== undefined) {
    if (patch.wrapText === null) range.setValue({ s: { tb: null } } as unknown as ICellData)
    else range.setWrap(patch.wrapText)
  }
  if (patch.textRotation !== undefined) {
    const rotation =
      patch.textRotation === null
        ? null
        : patch.textRotation === 'vertical'
          ? { v: BooleanNumber.TRUE }
          : { a: patch.textRotation }
    range.setValue({ s: { tr: rotation } } as unknown as ICellData)
  }
  if (patch.indent !== undefined) {
    // Indent renders as left padding (INDENT_STEP_PX per step); the journal
    // converts the padding back to OOXML indent steps on save.
    range.setValue({
      s: { pd: patch.indent ? { l: patch.indent * INDENT_STEP_PX } : null },
    } as unknown as ICellData)
  }
  if (patch.border !== undefined && patch.border !== null) {
    const type = BORDER_COMMAND_TYPES[patch.border.type]
    if (type) range.setBorder(type, BorderStyleTypes.THIN, patch.border.color ?? '#000000')
  }
}

export function loadWorkbookSkeleton(runtime: UniverRuntime | null, file: WorkbookFile): void {
  if (!runtime) return
  const activeWorkbook = runtime.univerAPI.getActiveWorkbook()
  if (activeWorkbook) runtime.univerAPI.disposeUnit(activeWorkbook.getId())
  runtime.univerAPI.createWorkbook({
    id: `file-${file.sha256}`,
    name: file.name,
    sheetOrder: file.sheets.map((sheet) => sheet.id),
    sheets: Object.fromEntries(
      file.sheets.map((sheet) => {
        const visuals = file.visuals.filter((visual) => visual.sheetId === sheet.id)
        const visualRowCount = visuals.reduce(
          (maximum, visual) => Math.max(maximum, visual.anchor.toRow + 1),
          0,
        )
        const visualColumnCount = visuals.reduce(
          (maximum, visual) => Math.max(maximum, visual.anchor.toColumn + 1),
          0,
        )
        return [
          sheet.id,
          {
            id: sheet.id,
            name: sheet.name,
            rowCount: Math.max(MINIMUM_SHEET_ROW_COUNT, sheet.rowCount, visualRowCount),
            columnCount: Math.max(MINIMUM_SHEET_COLUMN_COUNT, sheet.columnCount, visualColumnCount),
            hidden: sheet.hidden ? BooleanNumber.TRUE : BooleanNumber.FALSE,
            showGridlines: sheet.showGridLines ? BooleanNumber.TRUE : BooleanNumber.FALSE,
            ...(sheet.tabColor === null ? {} : { tabColor: sheet.tabColor }),
            ...(sheet.defaultRowHeight === null
              ? {}
              : { defaultRowHeight: (sheet.defaultRowHeight * 96) / 72 }),
            ...(sheet.defaultColumnWidth === null
              ? {}
              : { defaultColumnWidth: characterWidthToPixels(sheet.defaultColumnWidth) }),
            ...(sheet.freeze === null
              ? {}
              : {
                  freeze: {
                    xSplit: sheet.freeze.frozenColumns,
                    ySplit: sheet.freeze.frozenRows,
                    startRow: sheet.freeze.frozenRows,
                    startColumn: sheet.freeze.frozenColumns,
                  },
                }),
            columnData: createColumnData(sheet),
            cellData: {},
          },
        ]
      }),
    ),
  })
}

export function characterWidthToPixels(width: number): number {
  return width === 0 ? 0 : Math.floor(((256 * width + Math.floor(128 / 7)) / 256) * 7) + 5
}

/// Normalizes dialog input into the wire target format: '#Sheet!A1' for
/// internal anchors, a full URL otherwise. Bare domains get https://.
export function normalizeLinkTarget(input: string): string | null {
  const trimmed = input.trim()
  if (trimmed.length === 0 || trimmed.length > 2083) return null
  if (/^(https?:\/\/|mailto:)/i.test(trimmed)) return trimmed
  if (/^#?'?[^'!]+'?!\$?[A-Za-z]{1,3}\$?[0-9]+$/.test(trimmed)) {
    return trimmed.startsWith('#') ? trimmed : `#${trimmed}`
  }
  if (/^[\w-]+(\.[\w-]+)+(\/\S*)?$/.test(trimmed)) return `https://${trimmed}`
  return null
}

/// Shared by propose (fail early) and apply (fail closed) for protect_sheet.
export function protectSheetGuard(
  state: LazyWorkbookState,
  sheetId: string,
  nextProtected: boolean,
): string | null {
  if (isSheetRemoved(state.editJournal, sheetId)) return `Unknown sheet: ${sheetId}`
  const isAdded = state.editJournal.sheets.added.has(sheetId)
  const file = state.sheetProtections.get(sheetId)
  if (!file && !isAdded) {
    return t('appProtectionNeedsIndexed')
  }
  if (!nextProtected && file?.hasPassword) {
    return t('appProtectedWithPassword')
  }
  return null
}

/// Same journal write + link styling as the Insert Link menu action.
export function applyAiHyperlink(
  state: LazyWorkbookState,
  worksheet: UniverWorksheet,
  op: SetHyperlinkOperation,
): void {
  const sheetId = worksheet.getSheetId()
  const { row, column } = parseAddress(op.address)
  if (op.target === null) {
    recordHyperlinkEdit(state.editJournal, sheetId, row, column, null)
    state.hyperlinkTargets.get(sheetId)?.delete(`${row}:${column}`)
    worksheet.getRange(op.address).setValue({ s: { ul: null, cl: null } } as unknown as ICellData)
    return
  }
  const target = normalizeLinkTarget(op.target)
  if (target === null) {
    throw new Error(
      'set_hyperlink target must be a URL (https://…) or a sheet reference like Sheet1!A1.',
    )
  }
  recordHyperlinkEdit(state.editJournal, sheetId, row, column, target)
  worksheet.getRange(op.address).setValue({
    s: { cl: { rgb: '#0563C1' }, ul: { s: BooleanNumber.TRUE } },
  } as unknown as ICellData)
}

export function applyAiConditionalFormat(
  worksheet: UniverWorksheet,
  op: AddConditionalFormatOperation,
): void {
  const bounds = parseRange(op.range)
  const ranges: IRange[] = [
    {
      startRow: bounds.startRow,
      startColumn: bounds.startColumn,
      endRow: bounds.endRow,
      endColumn: bounds.endColumn,
    },
  ]
  const builder = worksheet.newConditionalFormattingRule()
  const rule = op.rule
  if (rule.kind === 'colorScale') {
    const stops = [
      rule.minColor,
      ...(rule.midColor === undefined ? [] : [rule.midColor]),
      rule.maxColor,
    ]
    const config = stops.map((color, index) => ({
      index,
      color,
      value:
        index === 0
          ? { type: CFValueType.min }
          : index === stops.length - 1
            ? { type: CFValueType.max }
            : { type: CFValueType.percentile, value: 50 },
    }))
    worksheet.addConditionalFormattingRule(
      builder
        .setColorScale(config as Parameters<typeof builder.setColorScale>[0])
        .setRanges(ranges)
        .build(),
    )
    return
  }
  if (rule.kind === 'dataBar') {
    worksheet.addConditionalFormattingRule(
      builder
        .setDataBar({
          min: { type: CFValueType.min },
          max: { type: CFValueType.max },
          positiveColor: rule.color ?? '#638EC6',
          nativeColor: '#FF555A',
          isShowValue: true,
        } as Parameters<typeof builder.setDataBar>[0])
        .setRanges(ranges)
        .build(),
    )
    return
  }
  let styled = buildAiHighlight(builder, rule)
  if (rule.format.fillColor !== undefined) styled = styled.setBackground(rule.format.fillColor)
  if (rule.format.fontColor !== undefined) styled = styled.setFontColor(rule.format.fontColor)
  if (rule.format.bold) styled = styled.setBold(true)
  if (rule.format.italic) styled = styled.setItalic(true)
  worksheet.addConditionalFormattingRule(styled.setRanges(ranges).build())
}

function buildAiHighlight(
  builder: ReturnType<UniverWorksheet['newConditionalFormattingRule']>,
  rule: Exclude<
    AddConditionalFormatOperation['rule'],
    { kind: 'colorScale' } | { kind: 'dataBar' }
  >,
): CfHighlightBuilder {
  switch (rule.kind) {
    case 'number':
      switch (rule.operator) {
        case 'greaterThan':
          return builder.whenNumberGreaterThan(rule.value)
        case 'greaterThanOrEqual':
          return builder.whenNumberGreaterThanOrEqualTo(rule.value)
        case 'lessThan':
          return builder.whenNumberLessThan(rule.value)
        case 'lessThanOrEqual':
          return builder.whenNumberLessThanOrEqualTo(rule.value)
        case 'equal':
          return builder.whenNumberEqualTo(rule.value)
        case 'notEqual':
          return builder.whenNumberNotEqualTo(rule.value)
        case 'between':
          return builder.whenNumberBetween(rule.value, rule.value2 ?? rule.value)
        case 'notBetween':
          return builder.whenNumberNotBetween(rule.value, rule.value2 ?? rule.value)
      }
      break
    case 'text':
      switch (rule.operator) {
        case 'contains':
          return builder.whenTextContains(rule.text)
        case 'notContains':
          return builder.whenTextDoesNotContain(rule.text)
        case 'beginsWith':
          return builder.whenTextStartsWith(rule.text)
        case 'endsWith':
          return builder.whenTextEndsWith(rule.text)
      }
      break
    case 'blank':
      return rule.blank ? builder.whenCellEmpty() : builder.whenCellNotEmpty()
    case 'duplicate':
      return rule.unique ? builder.setUniqueValues() : builder.setDuplicateValues()
    case 'top10':
      return builder.setRank({
        isBottom: rule.bottom === true,
        isPercent: rule.percent === true,
        value: rule.rank,
      })
    case 'formula':
      return builder.whenFormulaSatisfied(rule.formula)
  }
  throw new Error('Unsupported conditional-format rule.')
}

export function applyAiDataValidation(
  runtime: UniverRuntime,
  worksheet: UniverWorksheet,
  op: SetDataValidationOperation,
): void {
  const range = worksheet.getRange(op.range)
  if (op.validation === null) {
    range.setDataValidation(null)
    return
  }
  const rule = op.validation
  const builder = runtime.univerAPI.newDataValidation()
  const built =
    rule.kind === 'list'
      ? builder.requireValueInList([...rule.values], false, true)
      : rule.kind === 'listRef'
        ? builder.requireValueInRange(worksheet.getRange(rule.range), false, true)
        : rule.kind === 'numberBetween'
          ? builder.requireNumberBetween(rule.min, rule.max)
          : rule.kind === 'dateBetween'
            ? builder.requireDateBetween(new Date(rule.start), new Date(rule.end))
            : rule.kind === 'checkbox'
              ? builder.requireCheckbox()
              : builder.requireFormulaSatisfied(rule.formula)
  range.setDataValidation(built.build())
}

/// Jumps to an internal link target like `Sheet1!A1` or `'My Sheet'!B2`.
export function navigateToAnchor(
  runtime: UniverRuntime,
  location: string,
  setMessage: (message: string) => void,
): void {
  const match = /^'?([^'!]+)'?!(\$?[A-Z]+\$?[0-9]+)/.exec(location)
  const workbook = runtime.univerAPI.getActiveWorkbook()
  if (!match?.[1] || !match[2] || !workbook) {
    setMessage(t('appLinkInternal', { location }))
    return
  }
  const sheet = workbook.getSheets().find((candidate) => candidate.getSheetName() === match[1])
  if (!sheet) {
    setMessage(t('appLinkSheetNotFound', { name: match[1] }))
    return
  }
  try {
    workbook.setActiveSheet(sheet)
    const coordinates = parseAddress(match[2].replace(/\$/g, ''))
    sheet.scrollToCell(coordinates.row, coordinates.column)
  } catch {
    setMessage(t('appLinkJumpFailed', { location }))
  }
}

/// Installs the file's defined names into Univer with their scope. Names the
/// engine rejects go to `uninstalledDefinedNames`, which the declarative save
/// preserves verbatim. Installing must not mark the journal dirty.
export function applyDefinedNames(
  runtime: UniverRuntime | null,
  file: WorkbookFile,
  state: LazyWorkbookState,
): void {
  const workbook = runtime?.univerAPI.getActiveWorkbook()
  if (!workbook) return
  // Excel allows one definition per (name, scope) and resolves sheet-scope
  // first, workbook-scope second. Univer's name table is keyed by name alone
  // (first insert wins), so a #REF! sheet-scoped residue — Excel leaves those
  // behind when a sheet is deleted — appearing before the live workbook-level
  // definition used to shadow it for the whole book. Load each name's
  // live workbook-level definition first and push #REF! residues last. This
  // is a stopgap until the engine models (name, scope) pairs.
  const groups = new Map<string, typeof file.definedNames>()
  for (const defined of file.definedNames) {
    const list = groups.get(defined.name) ?? []
    list.push(defined)
    groups.set(defined.name, list)
  }
  const rank = (defined: (typeof file.definedNames)[number]): number =>
    defined.formula.includes('#REF!') ? 2 : defined.sheetIndex === undefined ? 0 : 1
  const ordered = [...groups.values()].flatMap((list) =>
    [...list].sort((a, b) => rank(a) - rank(b)),
  )
  journalSuppression.active = true
  try {
    for (const defined of ordered) {
      try {
        const localSheetId =
          defined.sheetIndex === undefined ? undefined : file.sheets[defined.sheetIndex]?.id
        if (defined.sheetIndex !== undefined && localSheetId === undefined) {
          throw new Error('Scope index out of range.')
        }
        const wb = workbook as unknown as {
          newDefinedNameBuilder(): {
            load(param: Record<string, unknown>): { build(): unknown }
          }
          insertDefinedNameBuilder(param: unknown): void
        }
        wb.insertDefinedNameBuilder(
          wb
            .newDefinedNameBuilder()
            .load({
              name: defined.name,
              formulaOrRefString: defined.formula,
              // Univer's workbook-scope sentinel; sheet scope carries the sheet id.
              localSheetId: localSheetId ?? 'AllDefaultWorkbook',
            })
            .build(),
        )
      } catch {
        // Names the engine can't model stay file-only; the save keeps them.
        state.uninstalledDefinedNames.add(defined.name)
      }
    }
  } finally {
    journalSuppression.active = false
  }
}

export function applyWorkbookNotes(runtime: UniverRuntime | null, file: WorkbookFile): void {
  const workbook = runtime?.univerAPI.getActiveWorkbook()
  if (!workbook) return
  // Installing the file's own notes must not mark their sheets note-dirty.
  journalSuppression.active = true
  try {
    applyWorkbookNotesInner(workbook, file)
  } finally {
    journalSuppression.active = false
  }
}

function applyWorkbookNotesInner(
  workbook: NonNullable<ReturnType<UniverRuntime['univerAPI']['getActiveWorkbook']>>,
  file: WorkbookFile,
): void {
  for (const sheet of file.sheets) {
    if (sheet.comments.length === 0) continue
    const worksheet = workbook.getSheetBySheetId(sheet.id)
    if (!worksheet) continue
    for (const comment of sheet.comments) {
      try {
        worksheet.getRange(comment.row, comment.column).createOrUpdateNote({
          id: `note-${sheet.id}-${comment.row}-${comment.column}`,
          row: comment.row,
          col: comment.column,
          width: 220,
          height: 90,
          note: comment.author ? `${comment.author}:\n${comment.text}` : comment.text,
        })
      } catch {
        // Notes are best-effort decoration.
      }
    }
  }
}

function createColumnData(
  sheet: WorkbookFile['sheets'][number],
): Record<number, { w?: number; hd?: BooleanNumber }> {
  const data: Record<number, { w?: number; hd?: BooleanNumber }> = {}
  for (const columnWidth of sheet.columnWidths) {
    const endColumn = Math.min(columnWidth.endColumn, sheet.columnCount - 1)
    // Outline-only <col> entries carry no width; leave the default width.
    const width = columnWidth.width
    const pixelWidth = width === undefined ? undefined : characterWidthToPixels(width)
    for (let column = columnWidth.startColumn; column <= endColumn; column += 1) {
      data[column] = {
        ...(pixelWidth !== undefined && ((width ?? 0) > 0 || !columnWidth.hidden)
          ? { w: pixelWidth }
          : {}),
        ...(columnWidth.hidden ? { hd: BooleanNumber.TRUE } : {}),
      }
    }
  }
  return data
}

export async function loadVisibleRange(
  runtime: UniverRuntime,
  lazyWorkbookRef: { current: LazyWorkbookState | null },
  worksheet: UniverWorksheet,
  setMessage: (message: string) => void,
  viewportStart?: { row: number; column: number },
): Promise<void> {
  const state = lazyWorkbookRef.current
  if (!state) return
  const sheetId = worksheet.getSheetId()
  const sheet = state.file.sheets.find((candidate) => candidate.id === sheetId)
  if (!sheet) return
  // Data bounds are screen-space: structural operations shift the extent.
  const ops = state.editJournal.structuralOps.get(sheetId) ?? []
  const screenRowCount = sheet.rowCount + netAxisDelta(ops, 'row')
  const screenColumnCount = sheet.columnCount + netAxisDelta(ops, 'column')
  if (screenRowCount <= 0 || screenColumnCount <= 0) return
  let visible: IRange | null = null
  try {
    visible = worksheet.getVisibleRange()
  } catch {
    // Univer can briefly have no scroll render controller while a workbook is
    // being replaced. The initial file range must still load.
  }
  // getVisibleRange lags the scroll by one render frame; a large jump
  // (name-box goto, hyperlink) produces a single Scroll event whose computed
  // range is the OLD spot — already loaded, so nothing loads and no later
  // event corrects it. Re-anchor at the actual scroll position.
  if (viewportStart) {
    visible = {
      startRow: viewportStart.row,
      endRow: viewportStart.row + (visible ? visible.endRow - visible.startRow : 79),
      startColumn: viewportStart.column,
      endColumn: viewportStart.column + (visible ? visible.endColumn - visible.startColumn : 15),
    }
  }
  const range = createBufferedRange(
    normalizeVisibleRange(visible, screenRowCount, screenColumnCount),
    screenRowCount,
    screenColumnCount,
  )
  await loadRange(runtime, lazyWorkbookRef, worksheet, range, setMessage)
  await loadFrozenColumnStrip(lazyWorkbookRef, worksheet, sheet, range)
}

/// Streams every sheet's formula list, computes the dependency closure, and
/// — when it fits the budget — installs and pins the closure cells so the
/// formula engine recalculates them live while the workbook keeps streaming.
export async function activateFormulaClosure(
  runtime: UniverRuntime,
  lazyWorkbookRef: { current: LazyWorkbookState | null },
  setMessage: (message: string) => void,
): Promise<void> {
  const state = lazyWorkbookRef.current
  if (!state || state.formulaMode || state.closure.status !== 'idle') return
  state.closure.status = 'pending'
  const giveUp = (): void => {
    if (lazyWorkbookRef.current === state) state.closure.status = 'unavailable'
  }

  const inputs: ClosureSheetInput[] = []
  for (const sheet of state.file.sheets) {
    const deadline = Date.now() + 180_000
    for (;;) {
      if (lazyWorkbookRef.current !== state) return
      let result
      try {
        result = await window.desktopApi.readWorkbookFormulas({
          sessionId: state.file.sessionId,
          sheetId: sheet.id,
        })
      } catch {
        return giveUp()
      }
      if (result.truncated) return giveUp()
      if (result.indexingComplete) {
        storeFormulaText(state, sheet.id, result.cells)
        inputs.push({
          id: sheet.id,
          name: sheet.name,
          rowCount: sheet.rowCount,
          columnCount: sheet.columnCount,
          formulas: result.cells.flatMap((cell) =>
            cell.formula ? [{ row: cell.row, column: cell.column, formula: cell.formula }] : [],
          ),
        })
        break
      }
      if (Date.now() > deadline) return giveUp()
      await new Promise((resolve) => setTimeout(resolve, 400))
    }
  }
  if (inputs.every((sheet) => sheet.formulas.length === 0)) return giveUp()
  const closure = computeFormulaClosure(inputs, CLOSURE_MAX_CELLS)
  if (!closure.ok) return giveUp()
  // Structural edits made while analyzing would shift the coordinates the
  // closure was computed in.
  if (state.editJournal.structuralOps.size > 0) return giveUp()
  const workbook = runtime.univerAPI.getActiveWorkbook()
  if (!workbook) return giveUp()

  for (const [sheetId, cells] of closure.cellsBySheet) {
    const worksheet = workbook.getSheetBySheetId(sheetId)
    const sheetMeta = state.file.sheets.find((candidate) => candidate.id === sheetId)
    if (!worksheet || !sheetMeta) continue
    const pinned = new Map<string, PinnedClosureCell>()
    for (const range of closureFetchRanges(cells)) {
      let result
      try {
        result = await window.desktopApi.readWorkbookRange({
          sessionId: state.file.sessionId,
          sheetId,
          range,
        })
      } catch {
        return giveUp()
      }
      if (lazyWorkbookRef.current !== state) return
      const wanted = result.cells.filter((cell) => cells.has(cellKey(cell.row, cell.column)))
      patchWorksheetRange(
        worksheet,
        undefined,
        range,
        wanted,
        state.file.styles,
        [],
        sheetMeta.tables,
        sheetMeta.freeze,
        true,
        state.editJournal,
      )
      for (const cell of wanted) {
        pinned.set(
          `${cell.row}:${cell.column}`,
          cell.formula ? { f: cell.formula, v: cell.value } : { v: cell.value },
        )
      }
    }
    state.closure.pinned.set(sheetId, pinned)
  }
  state.closure.status = 'active'
  setMessage(t('appClosureActive', { count: closure.formulaCount.toLocaleString() }))
}

interface MappedRangeRead {
  /// Result arrays translated into screen coordinates.
  readonly screen: Pick<WorkbookRangeResult, 'cells' | 'rows' | 'merges' | 'hyperlinks'>
  readonly raw: WorkbookRangeResult
  readonly indexedThroughScreen: number | null
  /// File-space end row actually requested; the indexing poll compares the
  /// raw cutoff against this.
  readonly fileEndRow: number
}

/// Reads a screen-space range, translating through the sheet's journaled
/// structural operations. Returns null when the range is entirely
/// journal-owned (inserted this session — nothing streams into it). A
/// request spanning deleted file rows can exceed the sidecar's per-read
/// cell budget, so file reads are split into row batches.
async function readSheetRangeMapped(
  state: LazyWorkbookState,
  sheetId: string,
  screenRange: IRange,
  sheet: WorkbookFile['sheets'][number],
): Promise<MappedRangeRead | null> {
  const ops = state.editJournal.structuralOps.get(sheetId) ?? []
  if (ops.length === 0) {
    const raw = await window.desktopApi.readWorkbookRange({
      sessionId: state.file.sessionId,
      sheetId,
      range: screenRange,
    })
    return {
      screen: raw,
      raw,
      indexedThroughScreen: raw.indexedThroughRow,
      fileEndRow: screenRange.endRow,
    }
  }
  const mappedRange = screenRangeToFileRange(ops, screenRange)
  if (
    !mappedRange ||
    mappedRange.startRow >= sheet.rowCount ||
    mappedRange.startColumn >= sheet.columnCount
  ) {
    return null
  }
  const fileRange: IRange = {
    startRow: mappedRange.startRow,
    endRow: Math.min(mappedRange.endRow, sheet.rowCount - 1),
    startColumn: mappedRange.startColumn,
    endColumn: Math.min(mappedRange.endColumn, sheet.columnCount - 1),
  }
  const width = fileRange.endColumn - fileRange.startColumn + 1
  const batchRows = Math.max(1, Math.floor(18_000 / width))
  const cells: WorkbookRangeResult['cells'] = []
  const rows: WorkbookRangeResult['rows'] = []
  const merges: WorkbookRangeResult['merges'] = []
  const hyperlinks: WorkbookRangeResult['hyperlinks'] = []
  let raw: WorkbookRangeResult | null = null
  for (let startRow = fileRange.startRow; startRow <= fileRange.endRow; startRow += batchRows) {
    const endRow = Math.min(startRow + batchRows - 1, fileRange.endRow)
    const batch = await window.desktopApi.readWorkbookRange({
      sessionId: state.file.sessionId,
      sheetId,
      range: { ...fileRange, startRow, endRow },
    })
    cells.push(...batch.cells)
    rows.push(...batch.rows)
    merges.push(...batch.merges)
    hyperlinks.push(...batch.hyperlinks)
    raw = batch
    // Later batches cannot have data before indexing reaches them; the
    // regular retry poll picks the rest up.
    if (batch.indexedThroughRow === null || batch.indexedThroughRow < endRow) break
  }
  if (!raw) return null
  return {
    screen: mapRangeResultToScreen(ops, { ...raw, cells, rows, merges, hyperlinks }),
    raw,
    indexedThroughScreen: indexedThroughScreenRow(ops, raw.indexedThroughRow),
    fileEndRow: fileRange.endRow,
  }
}

/// Reads a single-row/column vector for chart data-range edits. In lazy mode
/// the range may lie outside the loaded window, so values come from the file
/// (screen-mapped, journal edits overlaid) instead of the Univer model.
interface VisualUndoStep {
  undo(): void
  redo(): void
}

const VISUAL_UNDO_COMMAND_ID = 'sheets.mutation.visual-edit-step'
const visualUndoRegistry = new Map<number, VisualUndoStep>()
let visualUndoSequence = 0
const visualUndoRuntimes = new WeakSet<object>()

/// Interactive visual ops (chart edits, moves, deletes, inserts) enter
/// Univer's own undo stack as a custom mutation pair, so ⌘Z interleaves
/// them correctly with cell edits.
export function pushVisualUndo(runtime: UniverRuntime, step: VisualUndoStep): void {
  const unitId = runtime.univerAPI.getActiveWorkbook()?.getId()
  if (!unitId) return
  const injector = (
    runtime.univer as unknown as {
      __getInjector(): { get<T>(token: unknown): T }
    }
  ).__getInjector()
  if (!visualUndoRuntimes.has(runtime)) {
    visualUndoRuntimes.add(runtime)
    injector
      .get<{
        registerCommand(command: {
          id: string
          type: unknown
          handler: (
            accessor: unknown,
            params?: { token: number; direction: 'undo' | 'redo' },
          ) => boolean
        }): unknown
      }>(ICommandService)
      .registerCommand({
        id: VISUAL_UNDO_COMMAND_ID,
        type: CommandType.MUTATION,
        handler: (_accessor, params) => {
          const entry = params ? visualUndoRegistry.get(params.token) : undefined
          if (!entry || !params) return false
          if (params.direction === 'undo') entry.undo()
          else entry.redo()
          return true
        },
      })
  }
  const token = ++visualUndoSequence
  visualUndoRegistry.set(token, step)
  injector
    .get<{
      pushUndoRedo(item: {
        unitID: string
        undoMutations: { id: string; params: { token: number; direction: 'undo' | 'redo' } }[]
        redoMutations: { id: string; params: { token: number; direction: 'undo' | 'redo' } }[]
      }): void
    }>(IUndoRedoService)
    .pushUndoRedo({
      unitID: unitId,
      undoMutations: [{ id: VISUAL_UNDO_COMMAND_ID, params: { token, direction: 'undo' } }],
      redoMutations: [{ id: VISUAL_UNDO_COMMAND_ID, params: { token, direction: 'redo' } }],
    })
}

/// Bounding box of a set-range-values `cellValue` payload ({row: {col: …}}).
export function cellValueBounds(cellValue: unknown): CellBounds | null {
  if (typeof cellValue !== 'object' || cellValue === null) return null
  let bounds: CellBounds | null = null
  for (const [rowKey, rowValue] of Object.entries(cellValue)) {
    const row = Number(rowKey)
    if (!Number.isInteger(row) || row < 0) continue
    if (typeof rowValue !== 'object' || rowValue === null) continue
    for (const columnKey of Object.keys(rowValue as Record<string, unknown>)) {
      const column = Number(columnKey)
      if (!Number.isInteger(column) || column < 0) continue
      bounds =
        bounds === null
          ? { startRow: row, endRow: row, startColumn: column, endColumn: column }
          : {
              startRow: Math.min(bounds.startRow, row),
              endRow: Math.max(bounds.endRow, row),
              startColumn: Math.min(bounds.startColumn, column),
              endColumn: Math.max(bounds.endColumn, column),
            }
    }
  }
  return bounds
}

/// Demo counterpart of readChartRangeVector: the demo grid is fully loaded
/// in Univer, so the range reads straight off the worksheet.
export async function readDemoChartRangeVector(
  runtime: UniverRuntime,
  adapter: InMemoryWorkbookAdapter,
  visualId: string,
  rangeText: string,
): Promise<ChartVectorRead> {
  const visual = adapter.findVisual(visualId)
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const split = splitSheetRef(rangeText)
  const target = split
    ? workbook
        ?.getSheets()
        .find(
          (candidate) => candidate.getSheetName().toLowerCase() === split.sheetName.toLowerCase(),
        )
    : visual
      ? workbook?.getSheetBySheetId(visual.sheetId)
      : null
  if (!target) throw new Error('Unknown sheet for the chart data.')
  const range = (split?.range ?? rangeText).toUpperCase().replace(/\$/g, '')
  const bounds = parseRange(range)
  if (bounds.startRow !== bounds.endRow && bounds.startColumn !== bounds.endColumn) {
    throw new Error(t('appRangeMustBeVector', { range }))
  }
  if (rangeCellCount(bounds) > 1000)
    throw new Error(t('appRangeTooManyCells', { range, max: 1000 }))
  const vector = (
    target.getRange(range).getRawValues() as (string | number | boolean | null | undefined)[][]
  ).flat()
  return { vector, ref: absRangeRef(target.getSheetName(), range) }
}

/// Journal snapshot/restore for one visual, backing the undo closures.
interface VisualJournalSnapshot {
  readonly add: WorkbookVisualObject | undefined
  readonly edit: VisualEditEntry | undefined
  readonly chartEdit: Omit<WorkbookChartEdit, 'chartPath'> | undefined
}

export function captureVisualJournal(
  state: LazyWorkbookState,
  visualId: string,
  chartPath: string | undefined,
): VisualJournalSnapshot {
  return {
    add: state.editJournal.visualAdds.find((candidate) => candidate.id === visualId),
    edit: state.editJournal.visualEdits.get(visualId),
    chartEdit: chartPath === undefined ? undefined : state.editJournal.chartEdits.get(chartPath),
  }
}

export function restoreVisualJournal(
  state: LazyWorkbookState,
  visualId: string,
  chartPath: string | undefined,
  snapshot: VisualJournalSnapshot,
): void {
  const adds = state.editJournal.visualAdds
  const index = adds.findIndex((candidate) => candidate.id === visualId)
  if (snapshot.add === undefined) {
    if (index >= 0) adds.splice(index, 1)
  } else if (index >= 0) {
    adds[index] = snapshot.add
  } else {
    adds.push(snapshot.add)
  }
  if (snapshot.edit === undefined) state.editJournal.visualEdits.delete(visualId)
  else state.editJournal.visualEdits.set(visualId, snapshot.edit)
  if (chartPath !== undefined) {
    if (snapshot.chartEdit === undefined) state.editJournal.chartEdits.delete(chartPath)
    else state.editJournal.chartEdits.set(chartPath, snapshot.chartEdit)
  }
}

/// Values for a chart data range, robust to streaming: journal edits win,
/// then sidecar-mapped screen values; the Univer grid only when the workbook
/// is fully loaded (streamed-out cells read as empty there).
export async function readChartGridValues(
  state: LazyWorkbookState,
  runtime: UniverRuntime,
  sheetId: string,
  rangeText: string,
): Promise<(string | number | boolean | null | undefined)[][]> {
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const target = workbook?.getSheetBySheetId(sheetId)
  if (!target) throw new Error(`Unknown sheet: ${sheetId}`)
  const range = rangeText.toUpperCase().replace(/\$/g, '')
  const bounds = parseRange(range)
  if (rangeCellCount(bounds) > 2000)
    throw new Error(t('appRangeTooManyCells', { range, max: 2000 }))
  if (state.formulaMode) {
    // Raw model values: getValues() reads the view model, where numfmt and
    // formula-view interceptors have replaced numbers with display strings
    // ("12.5%", "=B5*C5"), which chartDataFromValues rejects.
    return target.getRange(range).getRawValues() as (
      string | number | boolean | null | undefined
    )[][]
  }
  const cells = new Map<string, string | number | boolean | null | undefined>()
  const sheetMeta = state.file.sheets.find((candidate) => candidate.id === sheetId)
  if (sheetMeta) {
    const mapped = await readSheetRangeMapped(state, sheetId, { ...bounds }, sheetMeta)
    if (
      mapped &&
      !mapped.raw.indexingComplete &&
      (mapped.indexedThroughScreen === null || mapped.indexedThroughScreen < bounds.endRow)
    ) {
      throw new Error(t('appSheetStillIndexing'))
    }
    for (const cell of mapped?.screen.cells ?? []) {
      cells.set(`${cell.row}:${cell.column}`, cell.value)
    }
  }
  for (const entry of journalEntriesInRange(state.editJournal, sheetId, bounds)) {
    if (entry.hasValue) cells.set(`${entry.row}:${entry.column}`, entry.value)
  }
  const grid: (string | number | boolean | null | undefined)[][] = []
  for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
    const line: (string | number | boolean | null | undefined)[] = []
    for (let column = bounds.startColumn; column <= bounds.endColumn; column += 1) {
      line.push(cells.get(`${row}:${column}`))
    }
    grid.push(line)
  }
  return grid
}

export async function readChartRangeVector(
  state: LazyWorkbookState,
  runtime: UniverRuntime,
  chartPath: string,
  rangeText: string,
): Promise<ChartVectorRead> {
  const workbook = runtime.univerAPI.getActiveWorkbook()
  // Chart `c:f` refs carry a sheet qualifier ("Data!$B$2:$B$5"); honor it,
  // falling back to the chart's own sheet for bare ranges.
  const split = splitSheetRef(rangeText)
  let sheetId: string | undefined
  if (split) {
    sheetId = state.file.sheets.find(
      (candidate) => candidate.name.toLowerCase() === split.sheetName.toLowerCase(),
    )?.id
  } else {
    sheetId = [...state.file.visuals, ...state.editJournal.visualAdds].find(
      (candidate) => candidate.chartPath === chartPath || candidate.id === chartPath,
    )?.sheetId
  }
  const target = sheetId === undefined ? null : workbook?.getSheetBySheetId(sheetId)
  if (sheetId === undefined || !target) throw new Error('Unknown sheet for the chart data.')
  const range = (split?.range ?? rangeText).toUpperCase().replace(/\$/g, '')
  const bounds = parseRange(range)
  if (bounds.startRow !== bounds.endRow && bounds.startColumn !== bounds.endColumn) {
    throw new Error(t('appRangeMustBeVector', { range }))
  }
  if (rangeCellCount(bounds) > 1000)
    throw new Error(t('appRangeTooManyCells', { range, max: 1000 }))
  const ref = absRangeRef(target.getSheetName(), range)
  if (state.formulaMode) {
    // Raw values: the view model may hold interceptor display strings.
    const vector = (
      target.getRange(range).getRawValues() as (string | number | boolean | null | undefined)[][]
    ).flat()
    return { vector, ref }
  }
  const cells = new Map<string, string | number | boolean | null | undefined>()
  const sheetMeta = state.file.sheets.find((candidate) => candidate.id === sheetId)
  if (sheetMeta) {
    const mapped = await readSheetRangeMapped(state, sheetId, { ...bounds }, sheetMeta)
    if (
      mapped &&
      !mapped.raw.indexingComplete &&
      (mapped.indexedThroughScreen === null || mapped.indexedThroughScreen < bounds.endRow)
    ) {
      throw new Error(t('appSheetStillIndexing'))
    }
    for (const cell of mapped?.screen.cells ?? []) {
      cells.set(`${cell.row}:${cell.column}`, cell.value)
    }
  }
  for (const entry of journalEntriesInRange(state.editJournal, sheetId, bounds)) {
    if (entry.hasValue) cells.set(`${entry.row}:${entry.column}`, entry.value)
  }
  const vector: (string | number | boolean | null | undefined)[] = []
  for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
    for (let column = bounds.startColumn; column <= bounds.endColumn; column += 1) {
      vector.push(cells.get(`${row}:${column}`))
    }
  }
  return { vector, ref }
}

const RECALC_DEBOUNCE_MS = 600
const RECALC_READ_BUDGET = 20_000
const RECALC_MAX_EDITS = 10_000
/// transient sidecar hiccups retry on the next edit; repeated rejection of
/// this workbook disables the fallback for the session
export const RECALC_MAX_FAILURES = 3

/// IronCalc fallback: when closure mode gave up on a streamed workbook, the
/// pending edits still recalculate — in the sidecar, against the on-disk
/// file — and the formula cells' engine values overlay the viewport.
export function queueFormulaRecalc(
  runtime: UniverRuntime,
  lazyWorkbookRef: { current: LazyWorkbookState | null },
  setMessage: (message: string) => void,
): void {
  const state = lazyWorkbookRef.current
  if (!state || state.formulaMode || state.closure.status !== 'unavailable') return
  if (state.recalc.failures >= RECALC_MAX_FAILURES) return
  // The engine loads the file from disk; session structural edits would
  // desync every coordinate — fail soft to cached values.
  if ([...state.editJournal.structuralOps.values()].some((ops) => ops.length > 0)) return
  if (state.recalc.timer) clearTimeout(state.recalc.timer)
  state.recalc.timer = setTimeout(() => {
    state.recalc.timer = null
    if (lazyWorkbookRef.current !== state) return
    void runFormulaRecalc(runtime, lazyWorkbookRef, state, setMessage)
  }, RECALC_DEBOUNCE_MS)
}

/// Formula-cell keys for one sheet, fetched once. A truncated list (>100k
/// formulas) caches as empty — unknown coverage would recalc the wrong set;
/// an incomplete index returns null so the next edit retries.
async function recalcFormulaCellKeys(
  state: LazyWorkbookState,
  sheetId: string,
): Promise<ReadonlySet<number> | null> {
  const cached = state.recalc.formulaCells.get(sheetId)
  if (cached) return cached
  const result = await window.desktopApi.readWorkbookFormulas({
    sessionId: state.file.sessionId,
    sheetId,
  })
  if (result.truncated) {
    state.recalc.formulaCells.set(sheetId, new Set())
    return null
  }
  if (!result.indexingComplete) return null
  storeFormulaText(state, sheetId, result.cells)
  const keys = new Set<number>()
  for (const cell of result.cells) keys.add(cellKey(cell.row, cell.column))
  state.recalc.formulaCells.set(sheetId, keys)
  return keys
}

/// Keep the formula text around for the formula bar — the closure may
/// still give up, and the recalc overlay only carries values.
function storeFormulaText(
  state: LazyWorkbookState,
  sheetId: string,
  cells: readonly { row: number; column: number; formula?: string | undefined }[],
): void {
  let bySheet = state.formulaText.get(sheetId)
  if (!bySheet) {
    bySheet = new Map()
    state.formulaText.set(sheetId, bySheet)
  }
  for (const cell of cells) {
    if (cell.formula) bySheet.set(`${cell.row}:${cell.column}`, cell.formula)
  }
}

async function runFormulaRecalc(
  runtime: UniverRuntime,
  lazyWorkbookRef: { current: LazyWorkbookState | null },
  state: LazyWorkbookState,
  setMessage: (message: string) => void,
): Promise<void> {
  const workbook = runtime.univerAPI.getActiveWorkbook()
  const worksheet = workbook?.getActiveSheet()
  if (!workbook || !worksheet) return
  const sheetId = worksheet.getSheetId()
  if (state.editJournal.sheets.added.has(sheetId)) return
  const edits: { sheetId: string; row: number; column: number; input: string }[] = []
  for (const [editSheetId, entries] of state.editJournal.cells) {
    if (isSheetRemoved(state.editJournal, editSheetId)) continue
    // A sheet added this session has no file part; formulas may reference
    // it, so the file-backed engine cannot represent this workbook.
    if (state.editJournal.sheets.added.has(editSheetId)) return
    for (const entry of entries.values()) {
      if (!entry.hasValue && !entry.formula) continue
      if (edits.length >= RECALC_MAX_EDITS) return
      edits.push({
        sheetId: editSheetId,
        row: entry.row,
        column: entry.column,
        input: toRecalcUserInput(entry),
      })
    }
  }
  const generation = ++state.recalc.generation
  try {
    const keys = await recalcFormulaCellKeys(state, sheetId)
    if (lazyWorkbookRef.current !== state || !keys || keys.size === 0) return
    const viewportStartRow = state.loadedRanges.get(sheetId)?.startRow ?? 0
    const reads = recalcReadRanges(keys, viewportStartRow, RECALC_READ_BUDGET)
    if (reads.length === 0) return
    const result = await window.desktopApi.recalcWorkbook({
      sessionId: state.file.sessionId,
      edits,
      reads: reads.map((range) => ({ sheetId, range })),
    })
    // A newer run superseded this one while the sidecar was evaluating.
    if (lazyWorkbookRef.current !== state || state.recalc.generation !== generation) return
    const overlay = new Map<string, PinnedClosureCell>()
    let unsupported = 0
    const journalCells = state.editJournal.cells.get(sheetId)
    for (const cell of result.cells) {
      if (cell.sheetId !== sheetId || !cell.isFormula) continue
      // The user's own journaled edits stay authoritative on screen.
      if (journalCells?.has(`${cell.row}:${cell.column}`)) continue
      // #NAME? flags a function IronCalc lacks; the file's cached value is
      // better — keep it.
      if (cell.formatted === '#NAME?') {
        unsupported += 1
        continue
      }
      overlay.set(`${cell.row}:${cell.column}`, { v: cell.number ?? cell.formatted })
    }
    state.recalc.overlay.set(sheetId, overlay)
    const loaded = state.loadedRanges.get(sheetId)
    if (loaded && overlay.size > 0) {
      journalSuppression.active = true
      try {
        applyPinnedOverlay(worksheet, overlay, undefined, loaded)
      } finally {
        journalSuppression.active = false
      }
    }
    state.recalc.failures = 0
    if (isActiveSheet(runtime, sheetId)) {
      setMessage(
        unsupported > 0
          ? t('appRecalcPartial', { count: unsupported })
          : t('appRecalcDone', { count: overlay.size }),
      )
    }
  } catch {
    // Fail soft: cached values stay on screen and the save still asks Excel
    // to recalculate on open. Repeated failures disable the fallback — but
    // only the current run may count (mirrors the success path's guard):
    // a superseded run failing after a newer success must not stack stale
    // failures toward the kill switch.
    if (lazyWorkbookRef.current !== state || state.recalc.generation !== generation) return
    state.recalc.failures += 1
  }
}

/// After horizontal scrolling the viewport range no longer covers frozen
/// columns; fetch that strip separately (patched without eviction).
async function loadFrozenColumnStrip(
  lazyWorkbookRef: { current: LazyWorkbookState | null },
  worksheet: UniverWorksheet,
  sheet: WorkbookFile['sheets'][number],
  viewportRange: IRange,
): Promise<void> {
  const state = lazyWorkbookRef.current
  const frozenColumns = sheet.freeze?.frozenColumns ?? 0
  if (!state || frozenColumns === 0 || frozenColumns > 8) return
  if (viewportRange.startColumn < frozenColumns) return
  const sheetId = worksheet.getSheetId()
  const stripRange: IRange = {
    startRow: viewportRange.startRow,
    endRow: viewportRange.endRow,
    startColumn: 0,
    endColumn: frozenColumns - 1,
  }
  const stripKey = `${sheetId}:${stripRange.startRow}:${stripRange.endRow}`
  if (state.frozenStripKeys.get(sheetId) === stripKey) return
  state.frozenStripKeys.set(sheetId, stripKey)
  try {
    const mapped = await readSheetRangeMapped(state, sheetId, stripRange, sheet)
    if (lazyWorkbookRef.current !== state || !mapped) {
      state.frozenStripKeys.delete(sheetId)
      return
    }
    const availableEndRow =
      mapped.indexedThroughScreen === null
        ? null
        : Math.min(mapped.indexedThroughScreen, stripRange.endRow)
    if (availableEndRow === null || availableEndRow < stripRange.startRow) {
      // Not indexed that far yet: without the rollback this strip would be
      // marked done and the frozen columns would stay blank forever.
      state.frozenStripKeys.delete(sheetId)
      return
    }
    patchWorksheetRange(
      worksheet,
      undefined,
      { ...stripRange, endRow: availableEndRow },
      mapped.screen.cells,
      state.file.styles,
      mapped.screen.hyperlinks,
      sheet.tables,
      sheet.freeze,
      state.formulaMode,
      state.editJournal,
      state.closure.pinned.get(sheetId),
      state.recalc.overlay.get(sheetId),
    )
  } catch {
    state.frozenStripKeys.delete(sheetId)
  }
}

async function loadRange(
  runtime: UniverRuntime,
  lazyWorkbookRef: { current: LazyWorkbookState | null },
  worksheet: UniverWorksheet,
  range: IRange,
  setMessage: (message: string) => void,
  isRetry = false,
  waitForRequestedRange = false,
  waitAttempt = 0,
): Promise<void> {
  const state = lazyWorkbookRef.current
  if (!state) return
  const sheetId = worksheet.getSheetId()
  const loaded = state.loadedRanges.get(sheetId)
  if (!isRetry && loaded && containsRange(loaded, range)) return
  const requestKey = `${range.startRow}:${range.endRow}:${range.startColumn}:${range.endColumn}`
  if (!isRetry && state.loadingKeys.get(sheetId) === requestKey) return
  const previousTimer = state.retryTimers.get(sheetId)
  if (previousTimer) clearTimeout(previousTimer)
  state.retryTimers.delete(sheetId)
  state.loadingKeys.set(sheetId, requestKey)

  try {
    const sheetMeta = state.file.sheets.find((candidate) => candidate.id === sheetId)
    if (!sheetMeta) return
    const mapped = await readSheetRangeMapped(state, sheetId, range, sheetMeta)
    if (lazyWorkbookRef.current !== state || state.loadingKeys.get(sheetId) !== requestKey) {
      return
    }
    if (!mapped) {
      // The whole range is journal-owned (inserted this session).
      state.loadedRanges.set(sheetId, range)
      return
    }
    const availableEndRow =
      mapped.indexedThroughScreen === null
        ? null
        : Math.min(mapped.indexedThroughScreen, range.endRow)
    if (availableEndRow !== null && availableEndRow >= range.startRow) {
      const availableRange = { ...range, endRow: availableEndRow }
      const alreadyLoaded = state.loadedRanges.get(sheetId)
      if (!alreadyLoaded || !containsRange(alreadyLoaded, availableRange)) {
        patchWorksheetRange(
          worksheet,
          alreadyLoaded,
          availableRange,
          mapped.screen.cells,
          state.file.styles,
          mapped.screen.hyperlinks,
          sheetMeta.tables,
          sheetMeta.freeze,
          state.formulaMode,
          state.editJournal,
          state.closure.pinned.get(sheetId),
          state.recalc.overlay.get(sheetId),
        )
        state.loadedRanges.set(sheetId, availableRange)
      }
    }
    const result = mapped.raw
    recordHyperlinks(state, sheetId, mapped.screen.hyperlinks)
    applyRowProperties(worksheet, state, sheetId, mapped.screen.rows)
    applyMerges(worksheet, state, sheetId, mapped.screen.merges)
    // Conditional formatting, filters, and validations install once with
    // file-space ranges; Univer shifts the installed models itself on later
    // structural edits, but a fresh install after a shift would be stale —
    // skip it (rare: the sheet was being edited before it first rendered).
    const hasStructuralOps = (state.editJournal.structuralOps.get(sheetId)?.length ?? 0) > 0
    if (!hasStructuralOps) {
      applyConditionalRules(worksheet, state, sheetId, result.conditionalRules)
      if (result.indexingComplete) {
        applySheetFilter(worksheet, state, sheetId, result.autoFilter)
        applyDataValidations(runtime, state, sheetId, result.dataValidations)
      }
    }
    if (result.indexingComplete && !state.sheetProtections.has(sheetId)) {
      state.sheetProtections.set(
        sheetId,
        result.sheetProtection ?? { protected: false, hasPassword: false },
      )
    }
    const sheet = sheetMeta
    if (!result.indexingComplete) {
      // Poll until the stream finishes: merged-cell ranges and trailing row
      // properties only become available at the end of the worksheet part.
      const indexedRows = (result.indexedThroughRow ?? -1) + 1
      if (
        isActiveSheet(runtime, sheetId) &&
        (result.indexedThroughRow === null || result.indexedThroughRow < mapped.fileEndRow)
      ) {
        setMessage(
          t('appIndexing', { name: sheet?.name ?? sheetId, rows: indexedRows.toLocaleString() }),
        )
      }
      if (
        waitForRequestedRange &&
        waitAttempt < 20 &&
        (availableEndRow === null || availableEndRow < range.endRow)
      ) {
        await new Promise((resolve) => setTimeout(resolve, 250))
        if (lazyWorkbookRef.current === state) {
          state.loadingKeys.delete(sheetId)
          await loadRange(
            runtime,
            lazyWorkbookRef,
            worksheet,
            range,
            setMessage,
            true,
            true,
            waitAttempt + 1,
          )
        }
      } else {
        const timer = setTimeout(() => {
          if (lazyWorkbookRef.current !== state) return
          state.loadingKeys.delete(sheetId)
          void loadRange(runtime, lazyWorkbookRef, worksheet, range, setMessage, true).then(() =>
            loadFrozenColumnStrip(lazyWorkbookRef, worksheet, sheet, range),
          )
        }, 250)
        state.retryTimers.set(sheetId, timer)
      }
    } else if (isActiveSheet(runtime, sheetId)) {
      setMessage(
        t('appStreamingRows', {
          name: state.file.name,
          rows: sheet?.rowCount.toLocaleString() ?? '?',
        }),
      )
    }
  } catch (error: unknown) {
    if (lazyWorkbookRef.current === state && isActiveSheet(runtime, sheetId)) {
      setMessage(error instanceof Error ? error.message : t('appLoadRangeFailed'))
    }
  } finally {
    if (state.loadingKeys.get(sheetId) === requestKey) {
      state.loadingKeys.delete(sheetId)
    }
  }
}

/// Loads an AI-requested range before its cells are read from Univer. Normal
/// viewport loading is intentionally fire-and-retry; AI reads instead wait
/// until the requested rows are indexed so unloaded cells cannot masquerade
/// as empty data.
export async function ensureLazyRangeLoaded(
  runtime: UniverRuntime,
  lazyWorkbookRef: { current: LazyWorkbookState | null },
  worksheet: UniverWorksheet,
  range: IRange,
  setMessage: (message: string) => void,
): Promise<boolean> {
  const initialState = lazyWorkbookRef.current
  if (!initialState) return false
  const sheet = initialState.file.sheets.find(
    (candidate) => candidate.id === worksheet.getSheetId(),
  )
  if (
    !sheet ||
    range.startRow < 0 ||
    range.startColumn < 0 ||
    range.endRow >= sheet.rowCount ||
    range.endColumn >= sheet.columnCount
  ) {
    return false
  }
  await loadRange(runtime, lazyWorkbookRef, worksheet, range, setMessage, false, true)
  const state = lazyWorkbookRef.current
  const loaded = state?.loadedRanges.get(worksheet.getSheetId())
  return state === initialState && loaded !== undefined && containsRange(loaded, range)
}

function applyRowProperties(
  worksheet: UniverWorksheet,
  state: LazyWorkbookState,
  sheetId: string,
  rows: WorkbookRangeResult['rows'],
): void {
  if (rows.length === 0) return
  let applied = state.appliedRowKeys.get(sheetId)
  if (!applied) {
    applied = new Set()
    state.appliedRowKeys.set(sheetId, applied)
  }
  journalSuppression.active = true
  try {
    for (const row of rows) {
      if (row.outlineLevel !== undefined || row.collapsed) {
        const rowsOutline = sheetOutline(state, sheetId).rows
        // Session group edits own the entry; file reads only seed it.
        if (!rowsOutline.has(row.row)) {
          rowsOutline.set(row.row, {
            level: row.outlineLevel ?? 0,
            collapsed: row.collapsed ?? false,
          })
        }
      }
      const key = `${row.row}:${row.height ?? ''}:${row.hidden}`
      if (applied.has(key)) continue
      applied.add(key)
      if (row.height !== undefined) {
        // The engine reports ht for every row that carries one, not just
        // customHeight="1" rows: Excel stores its laid-out height (auto-fit
        // included), and honoring it reproduces Excel's layout exactly.
        // Re-measuring instead with whatever fonts the host OS substitutes
        // clipped wrapped CJK rows on Windows. Forced = clip overflow and
        // skip auto-height, exactly like Excel renders a freshly opened file.
        worksheet.setRowHeightsForced(row.row, 1, Math.round((row.height * 96) / 72))
      }
      if (row.hidden) worksheet.hideRows(row.row, 1)
    }
  } finally {
    journalSuppression.active = false
  }
}

export function sheetOutline(
  state: LazyWorkbookState,
  sheetId: string,
): NonNullable<ReturnType<LazyWorkbookState['outline']['get']>> {
  let outline = state.outline.get(sheetId)
  if (!outline) {
    outline = { rows: new Map(), cols: new Map() }
    state.outline.set(sheetId, outline)
  }
  return outline
}

function applyMerges(
  worksheet: UniverWorksheet,
  state: LazyWorkbookState,
  sheetId: string,
  merges: WorkbookRangeResult['merges'],
): void {
  if (merges.length === 0) return
  let applied = state.appliedMerges.get(sheetId)
  if (!applied) {
    applied = new Set()
    state.appliedMerges.set(sheetId, applied)
  }
  journalSuppression.active = true
  try {
    for (const merge of merges) {
      const key = `${merge.startRow}:${merge.startColumn}:${merge.endRow}:${merge.endColumn}`
      if (applied.has(key)) continue
      applied.add(key)
      try {
        worksheet
          .getRange(
            merge.startRow,
            merge.startColumn,
            merge.endRow - merge.startRow + 1,
            merge.endColumn - merge.startColumn + 1,
          )
          .merge()
      } catch {
        // An overlapping merge from a previous partial pass is not fatal.
      }
    }
  } finally {
    journalSuppression.active = false
  }
}

function isActiveSheet(runtime: UniverRuntime, sheetId: string): boolean {
  return runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId() === sheetId
}

export function normalizeVisibleRange(
  visible: IRange | null | undefined,
  rowCount: number,
  columnCount: number,
): IRange {
  const fallback = {
    startRow: 0,
    endRow: Math.min(79, Math.max(0, rowCount - 1)),
    startColumn: 0,
    endColumn: Math.min(25, Math.max(0, columnCount - 1)),
  }
  if (
    !visible ||
    !Number.isFinite(visible.startRow) ||
    !Number.isFinite(visible.endRow) ||
    !Number.isFinite(visible.startColumn) ||
    !Number.isFinite(visible.endColumn) ||
    visible.startRow > visible.endRow ||
    visible.startColumn > visible.endColumn ||
    visible.startRow >= rowCount ||
    visible.startColumn >= columnCount
  ) {
    return fallback
  }
  return {
    startRow: Math.max(0, Math.trunc(visible.startRow)),
    endRow: Math.min(rowCount - 1, Math.max(0, Math.trunc(visible.endRow))),
    startColumn: Math.max(0, Math.trunc(visible.startColumn)),
    endColumn: Math.min(columnCount - 1, Math.max(0, Math.trunc(visible.endColumn))),
  }
}

function createBufferedRange(visible: IRange, rowCount: number, columnCount: number): IRange {
  const rowBuffer = 80
  const columnBuffer = 8
  // Frozen panes are protected from eviction in patchWorksheetRange rather
  // than folded into this range — startRow=0 at deep scroll would blow the
  // sidecar's 20k-cell request limit.
  return {
    startRow: Math.max(0, visible.startRow - rowBuffer),
    endRow: Math.min(rowCount - 1, visible.endRow + rowBuffer),
    startColumn: Math.max(0, visible.startColumn - columnBuffer),
    endColumn: Math.min(columnCount - 1, visible.endColumn + columnBuffer),
  }
}

function containsRange(container: IRange, requested: IRange): boolean {
  return (
    container.startRow <= requested.startRow &&
    container.endRow >= requested.endRow &&
    container.startColumn <= requested.startColumn &&
    container.endColumn >= requested.endColumn
  )
}

function patchWorksheetRange(
  worksheet: UniverWorksheet,
  previousRange: IRange | undefined,
  range: IRange,
  cells: WorkbookRangeResult['cells'],
  styles: readonly WorkbookCellStyle[],
  hyperlinks: WorkbookRangeResult['hyperlinks'],
  tables: WorkbookFile['sheets'][number]['tables'],
  freeze: WorkbookFile['sheets'][number]['freeze'],
  useFormulas = false,
  journal?: EditJournal,
  pinned?: ReadonlyMap<string, PinnedClosureCell>,
  recalcOverlay?: ReadonlyMap<string, PinnedClosureCell>,
  arrayFollowers?: ReadonlySet<string>,
): void {
  journalSuppression.active = true
  try {
    patchWorksheetRangeInner(
      worksheet,
      previousRange,
      range,
      cells,
      styles,
      hyperlinks,
      tables,
      freeze,
      useFormulas,
      arrayFollowers,
    )
    // Closure cells were just evicted or clobbered with static cached
    // values; re-pin them first — the journal overlay runs after so user
    // edits always win over pinned originals. Engine-recalculated values
    // sit between the two for the same reason.
    if (pinned?.size) applyPinnedOverlay(worksheet, pinned, previousRange, range)
    if (recalcOverlay?.size) applyPinnedOverlay(worksheet, recalcOverlay, previousRange, range)
    if (journal) applyJournalOverlay(worksheet, journal, range)
  } finally {
    journalSuppression.active = false
  }
}

function applyPinnedOverlay(
  worksheet: UniverWorksheet,
  pinned: ReadonlyMap<string, PinnedClosureCell>,
  previousRange: IRange | undefined,
  range: IRange,
): void {
  const covers = (candidate: IRange, row: number, column: number): boolean =>
    row >= candidate.startRow &&
    row <= candidate.endRow &&
    column >= candidate.startColumn &&
    column <= candidate.endColumn
  for (const [key, cell] of pinned) {
    const [rowText, columnText] = key.split(':')
    const row = Number(rowText)
    const column = Number(columnText)
    if (!covers(range, row, column) && !(previousRange && covers(previousRange, row, column))) {
      continue
    }
    worksheet
      .getRange(row, column, 1, 1)
      .setValues([
        [
          cell.f !== undefined
            ? cell.v === null || cell.v === undefined
              ? { f: cell.f }
              : { f: cell.f, v: cell.v }
            : typeof cell.v === 'string' && cell.v !== ''
              ? { v: cell.v, t: CellValueType.STRING }
              : { v: cell.v ?? null },
        ],
      ])
  }
}

function applyJournalOverlay(
  worksheet: UniverWorksheet,
  journal: EditJournal,
  range: IRange,
): void {
  for (const entry of journalEntriesInRange(journal, worksheet.getSheetId(), range)) {
    const cellRange = worksheet.getRange(entry.row, entry.column, 1, 1)
    if (entry.hasValue) {
      if (entry.formula) cellRange.setValues([[{ f: entry.formula }]])
      else if (entry.value === null) cellRange.clearContent()
      else if (entry.rich && typeof entry.value === 'string') {
        cellRange.setValues([[{ p: toRichTextDocument(entry.value, [...entry.rich]) }]])
      } else if (typeof entry.value === 'string' && entry.value.includes('\n')) {
        cellRange.setValues([[{ p: toRichTextDocument(entry.value) }]])
      } else cellRange.setValues([[{ v: entry.value }]])
    }
    // The set-range-values mutation merges style patches, so re-applying the
    // delta over the just-installed original reproduces the edited look.
    if (entry.styleReset) {
      cellRange.setValues([[{ s: null } as unknown as ICellData]])
    }
    if (entry.style) {
      cellRange.setValues([[{ s: fromNeutralStyle(entry.style) as IStyleData }]])
    }
  }
}

function patchWorksheetRangeInner(
  worksheet: UniverWorksheet,
  previousRange: IRange | undefined,
  range: IRange,
  cells: WorkbookRangeResult['cells'],
  styles: readonly WorkbookCellStyle[],
  hyperlinks: WorkbookRangeResult['hyperlinks'],
  tables: WorkbookFile['sheets'][number]['tables'],
  freeze: WorkbookFile['sheets'][number]['freeze'],
  useFormulas: boolean,
  arrayFollowers?: ReadonlySet<string>,
): void {
  if (previousRange) {
    // Frozen rows/columns stay visible while scrolling, so never evict them —
    // later viewport patches don't include them and they'd go blank.
    const clearStartRow = Math.max(previousRange.startRow, freeze?.frozenRows ?? 0)
    const clearStartColumn = Math.max(previousRange.startColumn, freeze?.frozenColumns ?? 0)
    if (clearStartRow <= previousRange.endRow && clearStartColumn <= previousRange.endColumn) {
      const previous = worksheet.getRange(
        clearStartRow,
        clearStartColumn,
        previousRange.endRow - clearStartRow + 1,
        previousRange.endColumn - clearStartColumn + 1,
      )
      previous.clearContent()
      previous.clearFormat()
    }
  }
  const linkedCells = new Set(hyperlinks.map((link) => `${link.row}:${link.column}`))
  const rows = range.endRow - range.startRow + 1
  const columns = range.endColumn - range.startColumn + 1
  const matrix: ICellData[][] = Array.from({ length: rows }, () =>
    Array.from({ length: columns }, () => ({})),
  )
  for (const cell of cells) {
    if (
      cell.row < range.startRow ||
      cell.row > range.endRow ||
      cell.column < range.startColumn ||
      cell.column > range.endColumn
    ) {
      continue
    }
    const displayValue = cell.value ?? cell.formula ?? ''
    const row = matrix[cell.row - range.startRow]
    const style = cell.styleIndex === undefined ? undefined : styles[cell.styleIndex]
    // CSE array follower: its dead cached value would block the master's
    // spill with #SPILL!; keep the style, let the engine fill the content.
    if (useFormulas && arrayFollowers?.has(`${cell.row}:${cell.column}`)) {
      if (row) {
        row[cell.column - range.startColumn] = style ? { s: toUniverStyle(style) } : {}
      }
      continue
    }
    const isLink = linkedCells.has(`${cell.row}:${cell.column}`)
    const multiline = typeof displayValue === 'string' && displayValue.includes('\n')
    if (row) {
      row[cell.column - range.startColumn] = {
        // Explicit string typing: bare `v` lets Univer coerce numeric-looking
        // text ("007", phone numbers) into numbers.
        ...(cell.rich && typeof displayValue === 'string'
          ? { p: toRichTextDocument(displayValue, cell.rich) }
          : useFormulas && cell.formula
            ? // No cached value: leave v unset so the engine computes instead
              // of showing the formula text as a literal.
              cell.value === null || cell.value === undefined
              ? { f: cell.formula }
              : { f: cell.formula, v: cell.value }
            : typeof displayValue === 'string' && multiline
              ? // Bare `v` renders only the first line; the doc model keeps all.
                { p: toRichTextDocument(displayValue) }
              : typeof displayValue === 'string' && displayValue !== ''
                ? { v: displayValue, t: CellValueType.STRING }
                : { v: displayValue }),
        ...(style || isLink || multiline
          ? {
              s: {
                // Link blue/underline is a fallback only: a colour or
                // underline the file specifies must win.
                ...(isLink ? { cl: { rgb: '#0563C1' }, ul: { s: BooleanNumber.TRUE } } : {}),
                ...(style ? toUniverStyle(style) : {}),
                // Excel shows manual line breaks even without wrapText.
                ...(multiline ? { tb: WrapStrategy.WRAP } : {}),
              },
            }
          : {}),
      }
    }
  }
  applyTableBanding(matrix, range, tables)
  worksheet.getRange(range.startRow, range.startColumn, rows, columns).setValues(matrix)
}

/// Approximates Excel table styles (header band + row stripes) for cells that
/// carry no explicit fill of their own.
function applyTableBanding(
  matrix: ICellData[][],
  range: IRange,
  tables: WorkbookFile['sheets'][number]['tables'],
): void {
  for (const table of tables) {
    const rowStart = Math.max(range.startRow, table.range.startRow)
    const rowEnd = Math.min(range.endRow, table.range.endRow)
    const columnStart = Math.max(range.startColumn, table.range.startColumn)
    const columnEnd = Math.min(range.endColumn, table.range.endColumn)
    if (rowStart > rowEnd || columnStart > columnEnd) continue
    // Colors are resolved sidecar-side from the workbook's real theme accents
    // (Light/Medium/Dark variant rules); the literals are a last-resort fallback.
    const headerFill = table.headerFill
    const headerFont = table.headerFontColor ?? '#FFFFFF'
    const stripeFill = table.stripeFill ?? '#D9E1F2'
    const dataStartRow = table.range.startRow + table.headerRowCount
    for (let row = rowStart; row <= rowEnd; row += 1) {
      const isHeader = row < dataStartRow
      const isStripe = !isHeader && table.showRowStripes && (row - dataStartRow) % 2 === 1
      if (!isHeader && !isStripe) continue
      for (let column = columnStart; column <= columnEnd; column += 1) {
        const cell = matrix[row - range.startRow]?.[column - range.startColumn]
        if (!cell) continue
        const style = (cell.s ?? {}) as IStyleData
        if (style.bg) continue
        cell.s = isHeader
          ? {
              ...style,
              ...(headerFill ? { bg: { rgb: headerFill } } : {}),
              cl: { rgb: headerFill ? headerFont : (table.headerFontColor ?? '#333333') },
              bl: BooleanNumber.TRUE,
            }
          : { ...style, bg: { rgb: stripeFill } }
      }
    }
  }
}

export function toRichTextDocument(
  text: string,
  runs: readonly WorkbookRichRun[] = [],
): ICellData['p'] {
  const textRuns = []
  let cursor = 0
  for (const run of runs) {
    const end = cursor + run.text.length
    textRuns.push({
      st: cursor,
      ed: end,
      ts: {
        ...(run.family ? { ff: run.family } : {}),
        ...(run.size ? { fs: run.size } : {}),
        ...(run.bold ? { bl: BooleanNumber.TRUE } : {}),
        ...(run.italic ? { it: BooleanNumber.TRUE } : {}),
        ...(run.underline ? { ul: { s: BooleanNumber.TRUE } } : {}),
        ...(run.strikethrough ? { st: { s: BooleanNumber.TRUE } } : {}),
        ...(run.color ? { cl: { rgb: run.color } } : {}),
      },
    })
    cursor = end
  }
  // Univer document streams use \r as paragraph break and \n as section
  // break; a raw \n would split the cell into sections and drop later lines.
  // 1:1 replacement, so textRun offsets stay valid.
  const dataStream = `${text.replace(/\n/g, '\r')}\r\n`
  const paragraphs: Array<{ startIndex: number }> = []
  for (let i = 0; i < dataStream.length; i += 1) {
    if (dataStream[i] === '\r') paragraphs.push({ startIndex: i })
  }
  return {
    id: 'rich-cell',
    body: {
      dataStream,
      textRuns,
      paragraphs,
      sectionBreaks: [{ startIndex: dataStream.length - 1 }],
    },
    documentStyle: {},
  }
}

/// Formula mode: pull every sheet block by block and patch cells with their
/// Record the follower cells of legacy CSE array formulas: every cell
/// a master's `<f t="array" ref>` covers except the master itself. Masters
/// sit at the range's top-left, so ascending row-block order sees each master
/// before its followers. Coordinates are screen-space (mapped through `ops`).
export function collectArrayFollowers(
  followers: Set<string>,
  cells: WorkbookRangeResult['cells'],
  ops: Parameters<typeof fileRangeToScreenRange>[0],
): void {
  for (const cell of cells) {
    if (!cell.arrayRef || !cell.formula) continue
    let bounds: IRange | null
    try {
      bounds = parseRange(cell.arrayRef)
    } catch {
      continue
    }
    if (rangeCellCount(bounds) > 100_000) continue
    if (ops.length > 0) bounds = fileRangeToScreenRange(ops, bounds)
    if (!bounds) continue
    for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
      for (let column = bounds.startColumn; column <= bounds.endColumn; column += 1) {
        if (row === cell.row && column === cell.column) continue
        followers.add(`${row}:${column}`)
      }
    }
  }
}

/// formulas so Univer's engine recalculates the whole workbook locally.
export async function preloadEntireWorkbook(
  runtime: UniverRuntime,
  lazyWorkbookRef: { current: LazyWorkbookState | null },
  setMessage: (message: string) => void,
): Promise<void> {
  const state = lazyWorkbookRef.current
  const workbook = runtime.univerAPI.getActiveWorkbook()
  if (!state || !workbook) return
  for (const sheet of state.file.sheets) {
    const worksheet = workbook.getSheetBySheetId(sheet.id)
    if (!worksheet) continue
    const sheetId = sheet.id
    const rowsPerBlock = Math.max(1, Math.floor(20_000 / sheet.columnCount))
    const arrayFollowers = new Set<string>()
    for (let startRow = 0; startRow < sheet.rowCount; startRow += rowsPerBlock) {
      if (lazyWorkbookRef.current !== state) return
      const range: IRange = {
        startRow,
        endRow: Math.min(sheet.rowCount - 1, startRow + rowsPerBlock - 1),
        startColumn: 0,
        endColumn: sheet.columnCount - 1,
      }
      let result
      try {
        result = await window.desktopApi.readWorkbookRange({
          sessionId: state.file.sessionId,
          sheetId,
          range,
        })
        let guard = 0
        while (
          !result.indexingComplete &&
          (result.indexedThroughRow === null || result.indexedThroughRow < range.endRow) &&
          guard < 400
        ) {
          await new Promise((resolve) => setTimeout(resolve, 150))
          guard += 1
          result = await window.desktopApi.readWorkbookRange({
            sessionId: state.file.sessionId,
            sheetId,
            range,
          })
        }
      } catch {
        return
      }
      if (lazyWorkbookRef.current !== state) return
      // Structural edits made while the preload runs shift screen positions;
      // install each block through the current mapping.
      const ops = state.editJournal.structuralOps.get(sheetId) ?? []
      const screenRange = ops.length === 0 ? range : fileRangeToScreenRange(ops, range)
      if (screenRange === null) continue
      const screen = ops.length === 0 ? result : mapRangeResultToScreen(ops, result)
      collectArrayFollowers(arrayFollowers, screen.cells, ops)
      patchWorksheetRange(
        worksheet,
        undefined,
        screenRange,
        screen.cells,
        state.file.styles,
        screen.hyperlinks,
        sheet.tables,
        sheet.freeze,
        true,
        state.editJournal,
        undefined,
        undefined,
        arrayFollowers,
      )
      recordHyperlinks(state, sheetId, screen.hyperlinks)
      applyRowProperties(worksheet, state, sheetId, screen.rows)
      applyMerges(worksheet, state, sheetId, screen.merges)
      if (result.indexingComplete && ops.length === 0) {
        applyConditionalRules(worksheet, state, sheetId, result.conditionalRules)
        applySheetFilter(worksheet, state, sheetId, result.autoFilter)
        applyDataValidations(runtime, state, sheetId, result.dataValidations)
      }
      if (result.indexingComplete && !state.sheetProtections.has(sheetId)) {
        state.sheetProtections.set(
          sheetId,
          result.sheetProtection ?? { protected: false, hasPassword: false },
        )
      }
    }
    const finalOps = state.editJournal.structuralOps.get(sheet.id) ?? []
    state.loadedRanges.set(sheetId, {
      startRow: 0,
      endRow: sheet.rowCount - 1 + netAxisDelta(finalOps, 'row'),
      startColumn: 0,
      endColumn: sheet.columnCount - 1 + netAxisDelta(finalOps, 'column'),
    })
  }
  if (lazyWorkbookRef.current === state) {
    state.flags.preloadComplete = true
    setMessage(t('appFullyLoaded'))
  }
}

function recordHyperlinks(
  state: LazyWorkbookState,
  sheetId: string,
  hyperlinks: WorkbookRangeResult['hyperlinks'],
): void {
  if (hyperlinks.length === 0) return
  let targets = state.hyperlinkTargets.get(sheetId)
  if (!targets) {
    targets = new Map()
    state.hyperlinkTargets.set(sheetId, targets)
  }
  for (const link of hyperlinks) {
    targets.set(`${link.row}:${link.column}`, link.target)
  }
}

function applySheetFilter(
  worksheet: UniverWorksheet,
  state: LazyWorkbookState,
  sheetId: string,
  autoFilter: WorkbookRangeResult['autoFilter'],
): void {
  if (state.appliedFilterSheets.has(sheetId)) return
  const sheet = state.file.sheets.find((candidate) => candidate.id === sheetId)
  // Excel allows one filter per sheet: worksheet autoFilter wins, else the
  // first table's own filter range.
  const area = autoFilter ?? sheet?.tables[0]?.range
  if (!area) return
  state.appliedFilterSheets.add(sheetId)
  const range: IRange = {
    startRow: area.startRow,
    startColumn: area.startColumn,
    endRow: area.endRow,
    endColumn: area.endColumn,
  }
  state.filterOrigins.set(sheetId, {
    origin: autoFilter ? 'worksheet' : 'table',
    range,
  })
  // Installing the file's own filter must not mark the sheet filter-dirty.
  journalSuppression.active = true
  try {
    worksheet
      .getRange(
        area.startRow,
        area.startColumn,
        area.endRow - area.startRow + 1,
        area.endColumn - area.startColumn + 1,
      )
      .createFilter()
  } catch {
    // A pre-existing filter is fine.
  } finally {
    journalSuppression.active = false
  }
}

/// Snapshots the live filter model of every filter-dirty sheet into the
/// declarative save payload. Color filters have no XLSX mapping here and
/// abort the save.
/// Snapshots the full CF rule set of every dirty sheet (Univer's model is
/// the wire format; the gateway maps it to OOXML and fails closed on shapes
/// it cannot represent).
export function collectCfStates(
  runtime: UniverRuntime | null,
  state: LazyWorkbookState,
): WorkbookCfState[] {
  const workbook = runtime?.univerAPI.getActiveWorkbook()
  if (!workbook) return []
  const states: WorkbookCfState[] = []
  for (const sheetId of state.editJournal.cfDirty) {
    if (isSheetRemoved(state.editJournal, sheetId)) continue
    const worksheet = workbook.getSheetBySheetId(sheetId)
    if (!worksheet) continue
    const rules = (
      worksheet as unknown as {
        getConditionalFormattingRules(): {
          ranges: IRange[]
          stopIfTrue?: boolean
          rule: Record<string, unknown>
        }[]
      }
    ).getConditionalFormattingRules()
    states.push({
      sheetId,
      rules: rules.map((rule) => ({
        ranges: rule.ranges.map((range) => ({
          startRow: range.startRow,
          endRow: range.endRow,
          startColumn: range.startColumn,
          endColumn: range.endColumn,
        })),
        stopIfTrue: rule.stopIfTrue === true,
        rule: rule.rule,
      })),
    })
  }
  return states
}

/// Snapshots the full data-validation rule set of every dirty sheet (same
/// recipe as CF: Univer's rule JSON is the wire format, mapped strictly by
/// the gateway, failing closed on unrepresentable shapes).
export function collectDvStates(
  runtime: UniverRuntime | null,
  state: LazyWorkbookState,
): WorkbookDvState[] {
  const workbook = runtime?.univerAPI.getActiveWorkbook()
  if (!workbook) return []
  const states: WorkbookDvState[] = []
  for (const sheetId of state.editJournal.dvDirty) {
    if (isSheetRemoved(state.editJournal, sheetId)) continue
    const worksheet = workbook.getSheetBySheetId(sheetId)
    if (!worksheet) continue
    const rules = (
      worksheet as unknown as {
        getDataValidations(): { rule: Record<string, unknown> & { ranges?: IRange[] } }[]
      }
    ).getDataValidations()
    states.push({
      sheetId,
      rules: rules.map(({ rule }) => {
        const { ranges, ...rest } = rule
        return {
          ranges: (ranges ?? []).map((range) => ({
            startRow: range.startRow,
            endRow: range.endRow,
            startColumn: range.startColumn,
            endColumn: range.endColumn,
          })),
          rule: rest,
        }
      }),
    })
  }
  return states
}

interface UniverDefinedName {
  getName(): string
  getFormulaOrRefString(): string
  getLocalSheetId(): string | undefined
  setName(name: string): void
  setRef(ref: string): void
  setScopeToWorkbook(): void
  delete(): void
}

export function univerDefinedNames(runtime: UniverRuntime | null): UniverDefinedName[] {
  const workbook = runtime?.univerAPI.getActiveWorkbook()
  if (!workbook) return []
  return (
    workbook as unknown as {
      getDefinedNames(): UniverDefinedName[]
    }
  ).getDefinedNames()
}

/// Snapshots the full defined-name model when it changed this session. Names
/// scoped to a sheet map back to the file's sheet order index.
export function collectDefinedNamesState(
  runtime: UniverRuntime | null,
  state: LazyWorkbookState,
): {
  names: { name: string; formula: string; sheetIndex?: number }[]
  preserveNames: string[]
} | null {
  if (!state.editJournal.definedNames.dirty) return null
  const names: { name: string; formula: string; sheetIndex?: number }[] = []
  for (const defined of univerDefinedNames(runtime)) {
    const localSheetId = defined.getLocalSheetId()
    // Univer reports workbook scope as the literal string 'AllDefaultWorkbook'.
    const scoped = localSheetId !== undefined && localSheetId !== 'AllDefaultWorkbook'
    const sheetIndex = scoped
      ? state.file.sheets.findIndex((sheet) => sheet.id === localSheetId)
      : -1
    if (scoped && sheetIndex === -1) {
      throw new Error(
        `The defined name "${defined.getName()}" is scoped to a sheet the file does ` +
          'not contain — it cannot be saved.',
      )
    }
    names.push({
      name: defined.getName(),
      formula: defined.getFormulaOrRefString().replace(/^=/, ''),
      ...(scoped ? { sheetIndex } : {}),
    })
  }
  return { names, preserveNames: [...state.uninstalledDefinedNames] }
}

/// Snapshots the live note set of every note-dirty sheet. Notes installed
/// from the file carry an "Author:\n" first line (see applyWorkbookNotes);
/// splitting it back keeps the author column on round-trip.
export function collectNoteStates(
  runtime: UniverRuntime | null,
  state: LazyWorkbookState,
): WorkbookNoteState[] {
  const workbook = runtime?.univerAPI.getActiveWorkbook()
  if (!workbook) return []
  const noteStates: WorkbookNoteState[] = []
  for (const sheetId of state.editJournal.noteDirty) {
    if (isSheetRemoved(state.editJournal, sheetId)) continue
    const worksheet = workbook.getSheetBySheetId(sheetId)
    if (!worksheet) continue
    const notes = worksheet.getNotes().map((note) => {
      const split = /^([^\n]{1,60}):\n([\s\S]*)$/.exec(note.note)
      return {
        row: note.row,
        column: note.col,
        author: split?.[1] ?? '',
        text: split?.[2] ?? note.note,
      }
    })
    noteStates.push({ sheetId, notes })
  }
  return noteStates
}

/// Shared landing path for column filter criteria: the AI op
/// `set_filter_criteria` and the Advanced Filter dialog both come through
/// here, so manual and AI edits hit the same facade command (and journal
/// through the same filter mutations). null clears the column's criteria.
export function applyFilterCriteria(
  worksheet: UniverWorksheet,
  column: string,
  criteria:
    | { readonly values: readonly string[] }
    | {
        readonly customs: {
          readonly and: boolean
          readonly filters: readonly AdvancedFilterCondition[]
        }
      }
    | null,
): void {
  const filter = worksheet.getFilter()
  if (!filter) throw new Error('This sheet has no auto-filter — set_filter first.')
  const filterColumn = columnIndex(column)
  const filterBounds = filter.getRange().getRange()
  if (filterColumn < filterBounds.startColumn || filterColumn > filterBounds.endColumn) {
    throw new Error(`Column ${column} is outside the auto-filter range.`)
  }
  if (criteria === null) {
    filter.removeColumnFilterCriteria(filterColumn)
    return
  }
  const colId = filterColumn - filterBounds.startColumn
  if ('values' in criteria) {
    filter.setColumnFilterCriteria(filterColumn, {
      colId,
      filters: { filters: [...criteria.values] },
    })
    return
  }
  filter.setColumnFilterCriteria(filterColumn, {
    colId,
    customFilters: buildCustomFilters(criteria.customs.and, criteria.customs.filters),
  })
}

/// Column choices for the Advanced Filter dialog: the filter range's header
/// row texts, falling back to the column letter for blank headers.
export function advancedFilterColumnOptions(
  worksheet: UniverWorksheet,
  filter: NonNullable<ReturnType<UniverWorksheet['getFilter']>>,
): AdvancedFilterColumn[] {
  const bounds = filter.getRange().getRange()
  const width = Math.min(bounds.endColumn - bounds.startColumn + 1, 26)
  const headerRow =
    worksheet.getRange(bounds.startRow, bounds.startColumn, 1, width).getValues()[0] ?? []
  return Array.from({ length: width }, (unused, offset) => {
    const header = headerRow[offset]
    return {
      colId: offset,
      label:
        header === null || header === undefined || header === ''
          ? t('appColumnLabel', { col: columnLabel(bounds.startColumn + offset) })
          : String(header),
    }
  })
}

export function collectFilterStates(
  runtime: UniverRuntime | null,
  state: LazyWorkbookState,
): WorkbookFilterState[] {
  const workbook = runtime?.univerAPI.getActiveWorkbook()
  if (!workbook) return []
  const filterStates: WorkbookFilterState[] = []
  for (const sheetId of state.editJournal.filterDirty) {
    if (isSheetRemoved(state.editJournal, sheetId)) continue
    const worksheet = workbook.getSheetBySheetId(sheetId)
    if (!worksheet) continue
    const origin = state.filterOrigins.get(sheetId)
    const filter = worksheet.getFilter()
    if (!filter) {
      // The user removed the filter; unhide what it was hiding.
      if (!origin) continue
      filterStates.push({
        sheetId,
        filter: null,
        hiddenRows: [],
        visibilityRange: toCellArea(origin.range),
      })
      continue
    }
    const filterRange = filter.getRange()
    const range: IRange = {
      startRow: filterRange.getRow(),
      startColumn: filterRange.getColumn(),
      endRow: filterRange.getRow() + filterRange.getHeight() - 1,
      endColumn: filterRange.getColumn() + filterRange.getWidth() - 1,
    }
    const columns: NonNullable<WorkbookFilterState['filter']>['columns'] = []
    for (let column = range.startColumn; column <= range.endColumn; column += 1) {
      const criteria = filter.getColumnFilterCriteria(column)
      if (!criteria) continue
      if (criteria.colorFilters) {
        throw new Error(t('appColorFiltersUnsaveable'))
      }
      if (!criteria.filters && !criteria.customFilters) continue
      columns.push({
        colId: column - range.startColumn,
        ...(criteria.filters?.filters ? { values: [...criteria.filters.filters] } : {}),
        ...(criteria.filters?.blank ? { blank: true } : {}),
        ...(criteria.customFilters
          ? {
              customs: {
                ...(criteria.customFilters.and ? { and: true } : {}),
                filters: criteria.customFilters.customFilters.map((custom) => ({
                  val: custom.val,
                  ...(custom.operator ? { operator: custom.operator } : {}),
                })),
              },
            }
          : {}),
      })
    }
    const visibilityRange = origin
      ? {
          startRow: Math.min(range.startRow, origin.range.startRow),
          startColumn: Math.min(range.startColumn, origin.range.startColumn),
          endRow: Math.max(range.endRow, origin.range.endRow),
          endColumn: Math.max(range.endColumn, origin.range.endColumn),
        }
      : range
    filterStates.push({
      sheetId,
      filter: { range: toCellArea(range), columns },
      hiddenRows: filter.getFilteredOutRows(),
      visibilityRange: toCellArea(visibilityRange),
    })
  }
  return filterStates
}

function toCellArea(range: IRange): WorkbookFilterState['visibilityRange'] {
  return {
    startRow: range.startRow,
    startColumn: range.startColumn,
    endRow: range.endRow,
    endColumn: range.endColumn,
  }
}

/// Journals every cell of a just-reordered (sorted) range straight from the
/// model, so the save writes the on-screen result.
/// Bounds of a Univer cell matrix (`{row: {col: cell}}`), for move-range
/// mutations that omit explicit from/to ranges.
export function matrixBounds(value: unknown): IRange | null {
  if (typeof value !== 'object' || value === null) return null
  let startRow = Number.POSITIVE_INFINITY
  let endRow = -1
  let startColumn = Number.POSITIVE_INFINITY
  let endColumn = -1
  for (const [rowKey, rowValue] of Object.entries(value)) {
    const row = Number(rowKey)
    if (!Number.isInteger(row) || typeof rowValue !== 'object' || rowValue === null) continue
    for (const columnKey of Object.keys(rowValue)) {
      const column = Number(columnKey)
      if (!Number.isInteger(column)) continue
      startRow = Math.min(startRow, row)
      endRow = Math.max(endRow, row)
      startColumn = Math.min(startColumn, column)
      endColumn = Math.max(endColumn, column)
    }
  }
  if (endRow < 0 || endColumn < 0) return null
  return { startRow, endRow, startColumn, endColumn }
}

export function journalRangeSnapshot(
  runtime: UniverRuntime,
  state: LazyWorkbookState,
  sheetId: string,
  range: IRange,
): void {
  const worksheet = runtime.univerAPI.getActiveWorkbook()?.getSheetBySheetId(sheetId)
  if (!worksheet) return
  const rows = range.endRow - range.startRow + 1
  const columns = range.endColumn - range.startColumn + 1
  if (rows <= 0 || columns <= 0 || rows * columns > 200_000) return
  const cellDatas = worksheet
    .getRange(range.startRow, range.startColumn, rows, columns)
    .getCellDatas()
  const cellValue: Record<number, Record<number, unknown>> = {}
  for (let rowOffset = 0; rowOffset < rows; rowOffset += 1) {
    const rowValues: Record<number, unknown> = {}
    for (let columnOffset = 0; columnOffset < columns; columnOffset += 1) {
      const cell = cellDatas[rowOffset]?.[columnOffset]
      const hasStyleObject = typeof cell?.s === 'object' && cell?.s !== null
      const hasContent =
        cell !== null &&
        cell !== undefined &&
        ('v' in cell ||
          (typeof cell.f === 'string' && cell.f.length > 0) ||
          cell.p !== undefined ||
          hasStyleObject)
      if (!hasContent) {
        rowValues[range.startColumn + columnOffset] = null
        continue
      }
      rowValues[range.startColumn + columnOffset] = {
        ...('v' in cell ? { v: cell.v } : {}),
        ...(typeof cell.f === 'string' && cell.f.length > 0 ? { f: cell.f } : {}),
        ...(cell.p !== undefined ? { p: cell.p } : {}),
        // Interned style ids can't be journaled; object styles (streamed
        // installs) re-apply as-is.
        ...(hasStyleObject ? { s: cell.s } : {}),
      }
    }
    cellValue[range.startRow + rowOffset] = rowValues
  }
  recordSetRangeValues(state.editJournal, sheetId, cellValue)
}

/// OOXML errorStyle ↔ Univer DataValidationErrorStyle (INFO=0, STOP=1,
/// WARNING=2). "stop" is the OOXML default.
const DV_ERROR_STYLES: Record<string, number> = { stop: 1, warning: 2, information: 0 }

/// Installs the file's validation rules verbatim into Univer's model — the
/// model is the wire format for the declarative save, so install fidelity IS
/// save fidelity. Only called once indexing completes; marking the sheet even
/// when it has no rules unlocks DV editing (the gate above).
function applyDataValidations(
  runtime: UniverRuntime,
  state: LazyWorkbookState,
  sheetId: string,
  rules: WorkbookRangeResult['dataValidations'],
): void {
  if (state.appliedDvSheets.has(sheetId)) return
  state.appliedDvSheets.add(sheetId)
  const unitId = `file-${state.file.sha256}`
  journalSuppression.active = true
  try {
    for (const [index, rule] of rules.entries()) {
      const mapped = toUniverDvRule(rule, `file-dv-${sheetId}-${index}`)
      if (!mapped) continue
      try {
        runtime.univerAPI.syncExecuteCommand('data-validation.mutation.addRule', {
          unitId,
          subUnitId: sheetId,
          rule: mapped,
        })
      } catch {
        // Unsupported validation shapes must not break streaming.
      }
    }
  } finally {
    journalSuppression.active = false
  }
}

/// File rule → Univer IDataValidationRule. Transformations are bijective with
/// the save-side mapping in xlsx-dv.ts: none↔any, list literal `"a,b"`↔`a,b`,
/// reference/custom formulas gain a leading `=`; everything else verbatim.
function toUniverDvRule(
  rule: WorkbookRangeResult['dataValidations'][number],
  uid: string,
): Record<string, unknown> | null {
  const type = rule.ruleType === 'none' ? 'any' : rule.ruleType
  if (!['any', 'whole', 'decimal', 'list', 'date', 'time', 'textLength', 'custom'].includes(type)) {
    return null
  }
  let formula1 = rule.formulas[0]
  const formula2 = rule.formulas[1]
  if (type === 'list' && formula1 !== undefined) {
    const literal = formula1.trim()
    // The insert-checkbox degrade writes list "1,0" (xlsx-dv.ts); restore it.
    if (literal === '"1,0"') {
      return {
        uid,
        type: 'checkbox',
        ranges: rule.ranges.map((area) => ({
          startRow: area.startRow,
          startColumn: area.startColumn,
          endRow: area.endRow,
          endColumn: area.endColumn,
        })),
        allowBlank: rule.allowBlank,
      }
    }
    formula1 =
      literal.startsWith('"') && literal.endsWith('"')
        ? literal.slice(1, -1)
        : `=${literal.replace(/^=/, '')}`
  } else if (type === 'custom' && formula1 !== undefined) {
    formula1 = `=${formula1.replace(/^=/, '')}`
  }
  const errorStyle = rule.errorStyle === undefined ? undefined : DV_ERROR_STYLES[rule.errorStyle]
  return {
    uid,
    type,
    ranges: rule.ranges.map((area) => ({
      startRow: area.startRow,
      startColumn: area.startColumn,
      endRow: area.endRow,
      endColumn: area.endColumn,
    })),
    allowBlank: rule.allowBlank,
    ...(rule.operator === undefined ? {} : { operator: rule.operator }),
    ...(formula1 === undefined ? {} : { formula1 }),
    ...(formula2 === undefined ? {} : { formula2 }),
    ...(type === 'list' ? { showDropDown: !rule.suppressDropdown } : {}),
    ...(rule.showInputMessage ? { showInputMessage: true } : {}),
    ...(rule.showErrorMessage ? { showErrorMessage: true } : {}),
    ...(errorStyle === undefined ? {} : { errorStyle }),
    ...(rule.errorTitle === undefined ? {} : { errorTitle: rule.errorTitle }),
    ...(rule.error === undefined ? {} : { error: rule.error }),
    ...(rule.promptTitle === undefined ? {} : { promptTitle: rule.promptTitle }),
    ...(rule.prompt === undefined ? {} : { prompt: rule.prompt }),
  }
}

function applyConditionalRules(
  worksheet: UniverWorksheet,
  state: LazyWorkbookState,
  sheetId: string,
  rules: WorkbookRangeResult['conditionalRules'],
): void {
  if (rules.length === 0 || state.appliedCfSheets.has(sheetId)) return
  state.appliedCfSheets.add(sheetId)
  // Lower xlsx priority number = higher precedence; Univer applies rules in
  // insertion order, so add the most important rules first. Installing the
  // file's own rules must not mark the sheet's CF as edited.
  const ordered = [...rules].sort((a, b) => a.priority - b.priority)
  journalSuppression.active = true
  try {
    for (const rule of ordered) {
      try {
        const built = buildConditionalRule(worksheet, state.file.dxfStyles, rule)
        if (built) worksheet.addConditionalFormattingRule(built)
      } catch {
        // An unsupported rule must not break the rest of the sheet.
      }
    }
  } finally {
    journalSuppression.active = false
  }
}

type CfHighlightBuilder = ReturnType<
  ReturnType<UniverWorksheet['newConditionalFormattingRule']>['whenCellNotEmpty']
>

function buildConditionalRule(
  worksheet: UniverWorksheet,
  dxfStyles: readonly WorkbookCellStyle[],
  rule: WorkbookRangeResult['conditionalRules'][number],
) {
  const ranges: IRange[] = rule.ranges.map((area) => ({
    startRow: area.startRow,
    startColumn: area.startColumn,
    endRow: area.endRow,
    endColumn: area.endColumn,
  }))
  const builder = worksheet.newConditionalFormattingRule()
  if (rule.ruleType === 'colorScale') {
    if (rule.colors.length < 2 || rule.cfvos.length !== rule.colors.length) return null
    return builder
      .setColorScale(
        rule.cfvos.map((cfvo, index) => ({
          index,
          color: rule.colors[index] ?? '#FFFFFF',
          value: toCfValue(cfvo),
        })),
      )
      .setRanges(ranges)
      .build()
  }
  if (rule.ruleType === 'dataBar') {
    const [min, max] = rule.cfvos
    if (!min || !max) return null
    return builder
      .setDataBar({
        min: toCfValue(min),
        max: toCfValue(max),
        positiveColor: rule.colors[0] ?? '#638EC6',
        nativeColor: rule.colors[1] ?? rule.colors[0] ?? '#FF0000',
        isShowValue: rule.showValue,
      })
      .setRanges(ranges)
      .build()
  }
  if (rule.ruleType === 'iconSet') {
    if (rule.cfvos.length < 3) return null
    const iconType = rule.iconSetName ?? '3TrafficLights1'
    const count = rule.cfvos.length
    // xlsx cfvos are ascending thresholds (first one is the catch-all minimum);
    // Univer wants a descending greaterThanOrEqual chain per icon.
    const worstFirst = WORST_FIRST_ICON_SETS.has(iconType)
    const configs = []
    for (let index = count - 1; index >= 0; index -= 1) {
      const cfvo = rule.cfvos[index]
      if (!cfvo) return null
      // The file's icon order runs worst-first; Univer's iconMap runs
      // best-first except for the rating sets.
      const fileIcon = rule.iconReverse ? count - 1 - index : index
      const iconIndex = worstFirst ? fileIcon : count - 1 - fileIcon
      configs.push({
        iconType,
        iconId: String(iconIndex),
        operator: index > 0 && cfvo.gte === false ? 'greaterThan' : 'greaterThanOrEqual',
        value: index === 0 ? { type: CFValueType.min } : toCfValue(cfvo),
      })
    }
    return builder
      .setIconSet({
        iconConfigs: configs as Parameters<typeof builder.setIconSet>[0]['iconConfigs'],
        isShowValue: rule.showValue,
      })
      .setRanges(ranges)
      .build()
  }
  const highlight = buildHighlightCondition(builder, rule)
  if (!highlight) return null
  return applyDxfFormat(highlight, dxfStyles, rule.dxfIndex).setRanges(ranges).build()
}

function toCfValue(cfvo: { kind: string; value?: string | undefined }): IValueConfig {
  switch (cfvo.kind) {
    case 'min':
      return { type: CFValueType.min }
    case 'max':
      return { type: CFValueType.max }
    case 'percent':
      return { type: CFValueType.percent, value: Number(cfvo.value ?? 0) }
    case 'percentile':
      return { type: CFValueType.percentile, value: Number(cfvo.value ?? 0) }
    case 'formula':
      return { type: CFValueType.formula, value: cfvo.value ?? '0' }
    default:
      return { type: CFValueType.num, value: Number(cfvo.value ?? 0) }
  }
}

function buildHighlightCondition(
  builder: ReturnType<UniverWorksheet['newConditionalFormattingRule']>,
  rule: WorkbookRangeResult['conditionalRules'][number],
): CfHighlightBuilder | null {
  const firstNumber = Number(rule.formulas[0])
  const secondNumber = Number(rule.formulas[1])
  switch (rule.ruleType) {
    case 'cellIs':
      if (!Number.isFinite(firstNumber)) return null
      switch (rule.operator) {
        case 'greaterThan':
          return builder.whenNumberGreaterThan(firstNumber)
        case 'greaterThanOrEqual':
          return builder.whenNumberGreaterThanOrEqualTo(firstNumber)
        case 'lessThan':
          return builder.whenNumberLessThan(firstNumber)
        case 'lessThanOrEqual':
          return builder.whenNumberLessThanOrEqualTo(firstNumber)
        case 'equal':
          return builder.whenNumberEqualTo(firstNumber)
        case 'notEqual':
          return builder.whenNumberNotEqualTo(firstNumber)
        case 'between':
          return Number.isFinite(secondNumber)
            ? builder.whenNumberBetween(firstNumber, secondNumber)
            : null
        case 'notBetween':
          return Number.isFinite(secondNumber)
            ? builder.whenNumberNotBetween(firstNumber, secondNumber)
            : null
        default:
          return null
      }
    case 'containsText':
      return rule.text ? builder.whenTextContains(rule.text) : null
    case 'notContainsText':
      return rule.text ? builder.whenTextDoesNotContain(rule.text) : null
    case 'beginsWith':
      return rule.text ? builder.whenTextStartsWith(rule.text) : null
    case 'endsWith':
      return rule.text ? builder.whenTextEndsWith(rule.text) : null
    case 'containsBlanks':
      return builder.whenCellEmpty()
    case 'notContainsBlanks':
      return builder.whenCellNotEmpty()
    case 'duplicateValues':
      return builder.setDuplicateValues()
    case 'uniqueValues':
      return builder.setUniqueValues()
    case 'top10':
      return rule.rank === undefined
        ? null
        : builder.setRank({
            isBottom: rule.bottom,
            isPercent: rule.percent,
            value: rule.rank,
          })
    case 'expression':
      return rule.formulas[0] ? builder.whenFormulaSatisfied(`=${rule.formulas[0]}`) : null
    default:
      return null
  }
}

function applyDxfFormat(
  highlight: CfHighlightBuilder,
  dxfStyles: readonly WorkbookCellStyle[],
  dxfIndex: number | undefined,
): CfHighlightBuilder {
  const dxf = dxfIndex === undefined ? undefined : dxfStyles[dxfIndex]
  if (!dxf) return highlight
  let styled = highlight
  if (dxf.fillColor) styled = styled.setBackground(dxf.fillColor)
  if (dxf.fontColor) styled = styled.setFontColor(dxf.fontColor)
  if (dxf.bold) styled = styled.setBold(true)
  if (dxf.italic) styled = styled.setItalic(true)
  if (dxf.underline) styled = styled.setUnderline(true)
  if (dxf.strikethrough) styled = styled.setStrikethrough(true)
  return styled
}

function toUniverStyle(style: WorkbookCellStyle): IStyleData {
  const diagonal = style.borderDiagonal ? toUniverBorder(style.borderDiagonal) : undefined
  const borders = {
    ...(style.borderTop ? { t: toUniverBorder(style.borderTop) } : {}),
    ...(style.borderBottom ? { b: toUniverBorder(style.borderBottom) } : {}),
    ...(style.borderLeft ? { l: toUniverBorder(style.borderLeft) } : {}),
    ...(style.borderRight ? { r: toUniverBorder(style.borderRight) } : {}),
    ...(diagonal && style.diagonalDown ? { tl_br: diagonal } : {}),
    ...(diagonal && style.diagonalUp ? { bl_tr: diagonal } : {}),
  }
  return {
    ...(style.fontFamily ? { ff: style.fontFamily } : {}),
    ...(style.fontSize ? { fs: style.fontSize } : {}),
    bl: style.bold ? BooleanNumber.TRUE : BooleanNumber.FALSE,
    it: style.italic ? BooleanNumber.TRUE : BooleanNumber.FALSE,
    ...(style.underline ? { ul: { s: BooleanNumber.TRUE } } : {}),
    ...(style.strikethrough ? { st: { s: BooleanNumber.TRUE } } : {}),
    ...(style.wrapText ? { tb: WrapStrategy.WRAP } : {}),
    ...(style.fontColor ? { cl: { rgb: style.fontColor } } : {}),
    ...(style.fillColor ? { bg: { rgb: style.fillColor } } : {}),
    ...(style.numberFormat ? { n: { pattern: style.numberFormat } } : {}),
    ...(Object.keys(borders).length > 0 ? { bd: borders } : {}),
    ...(mapHorizontalAlignment(style.horizontalAlignment) === undefined
      ? {}
      : { ht: mapHorizontalAlignment(style.horizontalAlignment) }),
    ...(mapVerticalAlignment(style.verticalAlignment) === undefined
      ? {}
      : { vt: mapVerticalAlignment(style.verticalAlignment) }),
    ...(style.indent ? { pd: { l: style.indent * INDENT_STEP_PX } } : {}),
  }
}

function toUniverBorder(edge: NonNullable<WorkbookCellStyle['borderTop']>): {
  s: BorderStyleTypes
  cl: { rgb: string }
} {
  return {
    s: mapBorderStyle(edge.style),
    cl: { rgb: edge.color ?? '#000000' },
  }
}

function mapBorderStyle(style: string): BorderStyleTypes {
  switch (style) {
    case 'hair':
      return BorderStyleTypes.HAIR
    case 'dotted':
      return BorderStyleTypes.DOTTED
    case 'dashed':
      return BorderStyleTypes.DASHED
    case 'dashDot':
      return BorderStyleTypes.DASH_DOT
    case 'dashDotDot':
      return BorderStyleTypes.DASH_DOT_DOT
    case 'double':
      return BorderStyleTypes.DOUBLE
    case 'medium':
      return BorderStyleTypes.MEDIUM
    case 'mediumDashed':
      return BorderStyleTypes.MEDIUM_DASHED
    case 'mediumDashDot':
      return BorderStyleTypes.MEDIUM_DASH_DOT
    case 'mediumDashDotDot':
      return BorderStyleTypes.MEDIUM_DASH_DOT_DOT
    case 'slantDashDot':
      return BorderStyleTypes.SLANT_DASH_DOT
    case 'thick':
      return BorderStyleTypes.THICK
    default:
      return BorderStyleTypes.THIN
  }
}

function mapHorizontalAlignment(value: string | undefined): HorizontalAlign | undefined {
  if (value === 'left') return HorizontalAlign.LEFT
  if (value === 'center') return HorizontalAlign.CENTER
  if (value === 'right') return HorizontalAlign.RIGHT
  if (value === 'justify') return HorizontalAlign.JUSTIFIED
  if (value === 'distributed') return HorizontalAlign.DISTRIBUTED
  return undefined
}

function mapVerticalAlignment(value: string | undefined): VerticalAlign | undefined {
  if (value === 'top') return VerticalAlign.TOP
  if (value === 'center') return VerticalAlign.MIDDLE
  if (value === 'bottom') return VerticalAlign.BOTTOM
  return undefined
}

export function disposeVisuals(disposables: { dispose(): void }[]): void {
  for (const disposable of disposables.splice(0)) disposable.dispose()
}

export function columnLetter(index: number): string {
  let label = ''
  for (let i = index; i >= 0; i = Math.floor(i / 26) - 1) {
    label = String.fromCharCode(65 + (i % 26)) + label
  }
  return label
}

/// Natural dimensions of an image data URL (fallback matches the picker).
export function measureImage(dataUrl: string): Promise<{ width: number; height: number }> {
  return new Promise((resolve) => {
    const image = new Image()
    image.onload = () => resolve({ width: image.naturalWidth, height: image.naturalHeight })
    image.onerror = () => resolve({ width: 480, height: 320 })
    image.src = dataUrl
  })
}

/// Absolute A1 ref for a rectangular range string, quoted-sheet prefixed.
export function absRangeRef(sheetName: string, range: string): string {
  const bounds = parseRange(range)
  const name = sheetName.replace(/'/g, "''")
  return (
    `'${name}'!$${columnLabel(bounds.startColumn)}$${bounds.startRow + 1}` +
    `:$${columnLabel(bounds.endColumn)}$${bounds.endRow + 1}`
  )
}

/// Absolute A1 ref over one column, rows in 0-based coordinates.
export function a1RangeRef(
  sheetName: string,
  column: number,
  fromRow: number,
  toRow: number,
): string {
  const col = columnLetter(column)
  const name = sheetName.replace(/'/g, "''")
  return `'${name}'!$${col}$${fromRow + 1}:$${col}$${toRow + 1}`
}

/// Absolute A1 ref over one row, columns in 0-based coordinates.
export function a1RowRangeRef(
  sheetName: string,
  row: number,
  fromColumn: number,
  toColumn: number,
): string {
  const name = sheetName.replace(/'/g, "''")
  return `'${name}'!$${columnLetter(fromColumn)}$${row + 1}:$${columnLetter(toColumn)}$${row + 1}`
}

export function queueVisualInstall(
  runtime: UniverRuntime,
  lazyWorkbookRef: { current: LazyWorkbookState | null },
  visualDisposablesRef: { current: { dispose(): void }[] },
  visualInstallTimerRef: { current: ReturnType<typeof setTimeout> | null },
  sheetId: string,
  chartEditRef?: { current: (chartPath: string, edit: ChartEditData) => void },
  chartVectorRef?: { current: (chartPath: string, range: string) => Promise<ChartVectorRead> },
  shapeEditRef?: { current: (visualId: string, changes: ShapeEditChanges) => void },
): void {
  const state = lazyWorkbookRef.current
  if (!state) return
  if (visualInstallTimerRef.current) clearTimeout(visualInstallTimerRef.current)
  visualInstallTimerRef.current = setTimeout(function install() {
    visualInstallTimerRef.current = null
    if (
      lazyWorkbookRef.current !== state ||
      runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId() !== sheetId
    ) {
      return
    }
    // Reinstalling mid-drag disposes the dragged node and kills its pointer
    // capture — hold off until the drop. Same for an open inline chart
    // editor, whose typed-in state lives in the float DOM.
    if (isVisualDragActive() || isChartEditorOpen()) {
      visualInstallTimerRef.current = setTimeout(install, 100)
      return
    }
    disposeVisuals(visualDisposablesRef.current)
    const addedVisuals = state.editJournal.visualAdds
    const visualEdits = state.editJournal.visualEdits
    // File visuals reflect pending edits: deleted ones disappear, moved ones
    // render at their journaled anchor.
    const fileVisuals =
      visualEdits.size === 0
        ? state.file.visuals
        : state.file.visuals
            .filter((visual) => !visualEdits.get(visual.id)?.remove)
            .map((visual) => {
              const anchor = visualEdits.get(visual.id)?.anchor
              return anchor ? { ...visual, anchor } : visual
            })
    const file =
      addedVisuals.length > 0 || fileVisuals !== state.file.visuals
        ? { ...state.file, visuals: [...fileVisuals, ...addedVisuals] }
        : state.file
    visualDisposablesRef.current = installWorkbookVisuals(
      runtime,
      file,
      sheetId,
      chartEditRef
        ? {
            edits: state.editJournal.chartEdits,
            onEdit: (chartPath, edit) => chartEditRef.current(chartPath, edit),
            ...(chartVectorRef
              ? { readVector: (chartPath, range) => chartVectorRef.current(chartPath, range) }
              : {}),
          }
        : undefined,
      shapeEditRef
        ? { onEdit: (visualId, changes) => shapeEditRef.current(visualId, changes) }
        : undefined,
    )
  }, 100)
}

/// Sparklines install separately from the floating visuals: dragging a
/// chart re-installs the visual pool, and rebuilding up to 200 sparkline
/// float DOMs with it would make every drag commit crawl.
export function queueSparklineInstall(
  runtime: UniverRuntime,
  lazyWorkbookRef: { current: LazyWorkbookState | null },
  sparklineDisposablesRef: { current: { dispose(): void }[] },
  sparklineTimerRef: { current: ReturnType<typeof setTimeout> | null },
  sheetId: string,
): void {
  const state = lazyWorkbookRef.current
  if (!state) return
  if (sparklineTimerRef.current) clearTimeout(sparklineTimerRef.current)
  sparklineTimerRef.current = setTimeout(() => {
    sparklineTimerRef.current = null
    if (
      lazyWorkbookRef.current !== state ||
      runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId() !== sheetId
    ) {
      return
    }
    disposeVisuals(sparklineDisposablesRef.current)
    const sheetMeta = state.file.sheets.find((sheet) => sheet.id === sheetId)
    const sparklines: SparklineGroupState[] = [
      ...(sheetMeta?.sparklines ?? []),
      ...state.editJournal.sparklineAdds
        .filter((entry) => entry.sheetId === sheetId)
        .map((entry) => ({
          type: entry.type,
          ...(entry.color === undefined ? {} : { color: entry.color }),
          cells: entry.cells,
        })),
    ]
    sparklineDisposablesRef.current =
      sparklines.length === 0 ? [] : installSparklines(runtime, sparklines, sheetId)
  }, 100)
}

export function clearLazyState(state: LazyWorkbookState | null): void {
  if (!state) return
  for (const timer of state.retryTimers.values()) clearTimeout(timer)
  state.retryTimers.clear()
  state.loadingKeys.clear()
  state.loadedRanges.clear()
}

/// Reads a cell's current content for AI previews and drift checks.
export function lazyCellReader(worksheet: UniverWorksheet): (address: string) => CellState {
  return (address) => {
    const range = worksheet.getRange(address)
    const formula = range.getFormula()
    const value = range.getValue() ?? null
    // Formula cells also carry their computed value (the AI needs to see results
    // and error values like #REF!/#DIV/0!; drift checks compare only formula
    // text for formula cells, see planStillMatches)
    if (formula) return { value, formula }
    return { value }
  }
}

/// Mirrors the BeforeSheetEditStart streaming guard for AI-planned cells.
export function lazyCellEditable(
  state: LazyWorkbookState,
  sheetId: string,
  row: number,
  column: number,
): boolean {
  if (state.flags.preloadComplete) return true
  const sheet = state.file.sheets.find((candidate) => candidate.id === sheetId)
  if (!sheet) return true
  if (row >= sheet.rowCount || column >= sheet.columnCount) return true
  const loaded = state.loadedRanges.get(sheetId)
  return (
    loaded !== undefined &&
    row >= loaded.startRow &&
    row <= loaded.endRow &&
    column >= loaded.startColumn &&
    column <= loaded.endColumn
  )
}
