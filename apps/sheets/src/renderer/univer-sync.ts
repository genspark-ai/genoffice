/**
 * Univer runtime synchronization helpers for the sheets renderer.
 *
 * Module-level functions that translate between the workbook file model
 * (snapshots, edit journal, lazy streaming state) and the live Univer
 * spreadsheet instance. Extracted from App.tsx; they hold no React state.
 */
import {
  BaselineOffset,
  BooleanNumber,
  BorderStyleTypes,
  CellValueType,
  CommandType,
  DataValidationRenderMode,
  HorizontalAlign,
  ICommandService,
  IUndoRedoService,
  LifecycleStages,
  VerticalAlign,
  WrapStrategy,
  type ICellData,
  type IRange,
  type IStyleData,
} from '@univerjs/core'
import { IFindReplaceService } from '@univerjs/preset-sheets-find-replace'
import { FontCache, getFontStyleString, IRenderManagerService } from '@univerjs/engine-render'
import { SheetSkeletonManagerService } from '@univerjs/sheets-ui'
import { CFValueType, type IValueConfig } from '@univerjs/preset-sheets-conditional-formatting'

import { CENTER_ACROSS_END_KEY } from './center-continuous'
import { notifyCfStreamWindow } from './cf-formula-fold'
import {
  THRESHOLD_RANGE_CELL_CAP,
  clampColorScaleStops,
  dataBarNeedsLayout,
  dataBarNeedsValues,
  defaultThreshold,
  evaluateThresholdFormula,
  isSelfContainedFormula,
  layoutDataBar,
  type ThresholdReader,
} from './cf-thresholds'

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
  escapeCssLeadingDigit,
  fromNeutralStyle,
  bulkConstantFillValueAt,
  journalCellContentAt,
  isSheetRemoved,
  journalEntriesInRange,
  ooxmlTextRotationToUniver,
  recordHyperlinkEdit,
  recordSetRangeValues,
  toRecalcUserInput,
  type EditJournal,
  type VisualEditEntry,
} from './edit-journal'
import {
  containsUnresolvedNames,
  cellKey,
  closureFetchRanges,
  computeFormulaClosure,
  keyColumn,
  keyRow,
  recalcReadRanges,
  type ClosureSheetInput,
} from './formula-closure'
import { isPlainArithmeticFormula } from './formula-cached-fallback'
import { degradeQuadraticFormulaCells } from './formula-cost'
import { extractFunctionNames } from './formula-functions'
import { DEFAULT_SHORT_DATE, setSystemShortDate } from '../shared/short-date'
import { getWorkbookMdw, setWorkbookMdw } from './app-constants'
import { EXCEL_DIGIT_PER_PT, fontAvailable } from './numfmt-fix'
import { t } from './i18n/locale'
import { mapProtectedRanges } from './protected-ranges'
import { INDENT_STEP_PX } from './selection-format'
import {
  fileRangeToScreenRange,
  fileRangeToScreenRanges,
  indexedThroughScreenRow,
  mapRangeResultToScreen,
  netAxisDelta,
  screenRangeToFileRange,
  screenToFile,
} from './view-transform'
import {
  buildCustomFilters,
  type AdvancedFilterColumn,
  type AdvancedFilterCondition,
} from './AdvancedFilterDialog'
import {
  installCellImages,
  installSparklines,
  installWorkbookVisuals,
  isChartEditorOpen,
  isVisualDragActive,
  type ChartEditData,
  type ChartVectorRead,
  type ShapeEditChanges,
  type SparklineGroupState,
} from './WorkbookVisuals'
import { VISUAL_UNDO_COMMAND_ID } from './undo-carry'
import {
  BORDER_COMMAND_TYPES,
  CLOSURE_MAX_CELLS,
  journalSuppression,
  lazySheetScreenExtent,
  loadAutoHeightSuppression,
  type ActiveWorkbook,
  type LazyWorkbookState,
  type PinnedClosureCell,
  type UniverRuntime,
  type UniverWorksheet,
} from './univer-state'

export const MINIMUM_SHEET_ROW_COUNT = 1000

/// Univer keys undo/redo stacks by unitId and keeps them across disposeUnit.
/// Both loaders reuse deterministic unitIds (`file-<sha>`, 'new-workbook'), so
/// without this reset a reopened workbook inherits the previous session's undo
/// steps and ⌘Z replays stale mutations onto the fresh content.
function clearUnitUndoHistory(runtime: UniverRuntime, unitId: string): void {
  ;(runtime.univer as unknown as { __getInjector(): { get<T>(token: unknown): T } })
    .__getInjector()
    .get<{ clearUndoRedo(unitId: string): void }>(IUndoRedoService)
    .clearUndoRedo(unitId)
}
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
  // A rebuild is a load, not an edit: suppress undo entries so the fresh
  // demo workbook starts with an empty stack (same convention as file opens)
  // and the QAT undo falls through to the adapter's revision history.
  journalSuppression.active = true
  try {
    loadSnapshotIntoUniverInner(runtime, snapshot, workbookId, workbookName)
  } finally {
    journalSuppression.active = false
  }
}

function loadSnapshotIntoUniverInner(
  runtime: UniverRuntime,
  snapshot: WorkbookSnapshot,
  workbookId: string,
  workbookName: string,
): void {
  const activeWorkbook = runtime.univerAPI.getActiveWorkbook()
  if (activeWorkbook) {
    clearUnitUndoHistory(runtime, activeWorkbook.getId())
    runtime.univerAPI.disposeUnit(activeWorkbook.getId())
  }
  // Rebuilds reuse the same unitId ('new-workbook'): drop the previous grid's
  // undo steps so ⌘Z falls through to the adapter's revision history instead
  // of replaying stale mutations onto the rebuilt content.
  clearUnitUndoHistory(runtime, workbookId)

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
            rowCount: Math.max(MINIMUM_SHEET_ROW_COUNT, maximumRow + 100, sheet.gridRows ?? 0),
            columnCount: Math.max(
              MINIMUM_SHEET_COLUMN_COUNT,
              maximumColumn + 10,
              sheet.gridColumns ?? 0,
            ),
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
  if (activeWorkbook) {
    clearUnitUndoHistory(runtime, activeWorkbook.getId())
    runtime.univerAPI.disposeUnit(activeWorkbook.getId())
  }
  // A load starts from a clean history even when the unitId was used before
  // (reopening the same unchanged file reuses `file-<sha>`).
  clearUnitUndoHistory(runtime, `file-${file.sha256}`)
  setWorkbookMdw(measureNormalFontMdw(file))
  setSystemShortDate(file.shortDateFormat ?? DEFAULT_SHORT_DATE)
  const created = runtime.univerAPI.createWorkbook({
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
        const columnCount = Math.max(
          MINIMUM_SHEET_COLUMN_COUNT,
          sheet.columnCount,
          visualColumnCount,
        )
        return [
          sheet.id,
          {
            id: sheet.id,
            name: sheet.name,
            rowCount: Math.max(MINIMUM_SHEET_ROW_COUNT, sheet.rowCount, visualRowCount),
            columnCount,
            hidden: sheet.hidden ? BooleanNumber.TRUE : BooleanNumber.FALSE,
            showGridlines: sheet.showGridLines ? BooleanNumber.TRUE : BooleanNumber.FALSE,
            ...(sheet.rightToLeft ? { rightToLeft: BooleanNumber.TRUE } : {}),
            ...(sheet.showRowColHeaders === false
              ? {
                  rowHeader: { width: 46, hidden: BooleanNumber.TRUE },
                  columnHeader: { height: 20, hidden: BooleanNumber.TRUE },
                }
              : {}),
            ...(sheet.tabColor === null ? {} : { tabColor: sheet.tabColor }),
            ...(sheet.defaultRowHeight === null
              ? {}
              : { defaultRowHeight: (sheet.defaultRowHeight * 96) / 72 }),
            // No defaultColWidth in the file ≠ Univer's 88px default: Excel
            // derives its built-in width from baseColWidth (default 8) plus
            // cell padding, snapped to 1/256 chars (~74px at MDW 8, live
            // Excel shows 8.0 chars). The narrower column matters: General
            // numbers switch to scientific when the digits stop fitting.
            defaultColumnWidth: characterWidthToPixels(
              sheet.defaultColumnWidth ?? paddedBaseColumnWidth(sheet.baseColumnWidth),
            ),
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
            // Excel restores the saved normal-view zoom on open; without it
            // a sheet authored at 70% opens cropped to the 100% viewport.
            ...(sheet.zoomScale === undefined ? {} : { zoomRatio: sheet.zoomScale / 100 }),
            columnData: createColumnData(sheet, file.styles, columnCount),
            cellData: {},
          },
        ]
      }),
    ),
  })
  // Excel opens on workbookView/@activeTab; Univer defaults to the first
  // visible sheet. Skip hidden targets (stale activeTab in the file).
  const activeMeta = [file.sheets[file.activeTab], ...file.sheets].find(
    (sheet) => sheet && !sheet.hidden,
  )
  if (created && activeMeta) {
    const sheet = created.getSheetBySheetId(activeMeta.id)
    if (sheet) created.setActiveSheet(sheet)
  }
}

export function characterWidthToPixels(width: number): number {
  const mdw = getWorkbookMdw()
  return width === 0 ? 0 : Math.floor(((256 * width + Math.floor(128 / mdw)) / 256) * mdw) + 5
}

/// Excel's built-in default column width when sheetFormatPr carries no
/// defaultColWidth: baseColWidth (default 8) plus cell padding, snapped to
/// the format's 1/256-char granularity (ECMA-376 §18.3.1.81).
export function paddedBaseColumnWidth(base: number | null | undefined): number {
  const chars = base ?? 8
  const mdw = getWorkbookMdw()
  return Math.trunc(((chars * mdw + 5) / mdw) * 256) / 256
}

const CJK_FAMILY_NAME =
  /[\u1100-\u11ff\u3000-\u303f\u3040-\u30ff\u3130-\u318f\u3400-\u4dbf\u4e00-\u9fff\uac00-\ud7af\uf900-\ufaff\uff00-\uffef]/

const clampMdw = (mdw: number): number => Math.max(4, Math.min(30, mdw))

/// The face Excel derives the column-width MDW from. For scheme fonts the
/// sidecar substitutes the theme latin face into styles[0], but Excel lays
/// CJK workbooks out per the Normal font's author-locale resolution — the
/// literal cached <name val> (e.g. MS PGothic under theme latin Calibri) —
/// with the theme's minor <a:ea> face as the fallback when no name is
/// cached. Only trust a differing literal that names a face we know.
export function resolveNormalMdwFamily(file: WorkbookFile): string {
  const normal = file.styles?.[0]
  const themeResolved = normal?.fontFamily ?? 'Calibri'
  // ja vertical-text names prefix the base face with '@'.
  let literal = (file.normalFontName ?? '').replace(/^@/, '')
  if (literal === '' && normal?.fontScheme === 'minor') {
    literal = file.themeFonts?.minorEa ?? ''
  }
  if (literal === '' || literal === themeResolved) return themeResolved
  if (
    EXCEL_DIGIT_PER_PT[literal] !== undefined ||
    CJK_FAMILY_NAME.test(literal) ||
    fontAvailable(literal)
  ) {
    return literal
  }
  return themeResolved
}

/// Excel's column-width unit is the Normal font's max digit width (MDW).
/// The hardcoded 7 only holds for Calibri 11; e.g. Verdana 10 workbooks use
/// MDW 8, and trusting 7 renders every column ~11% narrower than Excel,
/// wrapping text a line early. Known GDI digit widths win over canvas
/// measurement: alias-substituted faces (Aptos Narrow → Carlito, CJK names →
/// macOS faces) measure the substitute's digits, which is wrong in either
/// direction.
export function measureNormalFontMdw(file: WorkbookFile): number {
  const size = file.styles?.[0]?.fontSize ?? 11
  const family = resolveNormalMdwFamily(file)
  // Mac Excel (the fidelity reference) lays Calibri 11 out at MDW 8, not the
  // Windows GDI 7: live probes and ref print geometry both fit
  // floor((w+16/256)*8) within 1pt across corpora (prod_054 50.86ch→305pt,
  // prod_027 32.44ch→195pt, lo/built-in_ranges 13.71ch→82pt) while MDW 7
  // misses by 12%+, wrapping borderline text a line early.
  if (family === 'Calibri' && size === 11) return 8
  const perPt = EXCEL_DIGIT_PER_PT[family]
  if (perPt !== undefined) return clampMdw(Math.round(perPt * size))
  // CJK faces without a table entry: canvas would measure a macOS
  // substitute; every grounded legacy CJK face is em/2 → 8px at 11pt.
  if (CJK_FAMILY_NAME.test(family)) return clampMdw(Math.round((8 / 11) * size))
  if (typeof document === 'undefined') return 7
  try {
    const context = document.createElement('canvas').getContext('2d')
    if (!context) return 7
    // Quote the family: an unquoted leading digit or '@' silently no-ops
    // the whole ctx.font assignment and measures the 10px canvas default.
    context.font = `${(size * 96) / 72}px "${family.replace(/["\\]/g, '')}"`
    const width = context.measureText('0').width
    // Round with a 0.4 threshold: Arial 10 measures 7.42px and prints at
    // MDW 8 (plain round loses 12% of every column), while Arial 11 (8.16)
    // and Century Gothic 11 (8.13) print at 8 (ceil overshoots them 12% the
    // other way). Any threshold in [0.58, 0.84) fits all calibrated refs.
    return width > 0 ? clampMdw(Math.floor(width + 0.6)) : 7
  } catch {
    return 7
  }
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

/// Effective workbook structure lock: the session's desired state when
/// toggled, else the file's. Gates sheet add/remove/rename/reorder/hide.
export function workbookStructureLocked(state: LazyWorkbookState | null): boolean {
  if (!state) return false
  return (
    state.editJournal.workbookProtection.desired ??
    state.file.workbookProtection?.lockStructure ??
    false
  )
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
/**
 * scrollToCell computes its frozen-pane offset from DEFAULT row sizes, so
 * custom-height frozen rows hide the revealed cell underneath the pane
 * (alpha ledger r135: search/Go To landed the match behind the frozen
 * header). Iteratively correct: if the viewport starts past the target after
 * the scroll, re-scroll by the measured overshoot — converges immediately
 * for uniform-height regions and stops as soon as the target is visible.
 */
/** newest reveal wins: rapid Find Next/Previous must not let an older
 *  correction loop keep scrolling after a newer one started (bugbot) */
let revealGeneration = 0

function isRtlFacadeSheet(sheet: {
  getSheet?(): { getConfig(): { rightToLeft?: BooleanNumber } }
}): boolean {
  try {
    return sheet.getSheet?.().getConfig().rightToLeft === BooleanNumber.TRUE
  } catch {
    return false
  }
}

export async function revealCellBelowFreeze(
  sheet: {
    scrollToCell(row: number, column: number): unknown
    getVisibleRange(): IRange | null
    getSheet?(): { getConfig(): { rightToLeft?: BooleanNumber } }
  },
  row: number,
  column: number,
): Promise<void> {
  // aim one row above the target: the offset error is fractional rows, so a
  // position converged exactly on the target can still leave it half-sliced
  // under the pane — with the predecessor as the boundary row the target is
  // whole. Every scrollToCell call re-applies the same broken offset, so the
  // compensation folds into one running aim; the scroll and its viewport
  // update settle asynchronously, hence the waits between measurements.
  const generation = ++revealGeneration
  if (isRtlFacadeSheet(sheet)) {
    // On an RTL sheet scrollToCell parks the target at the viewport's
    // visually-left (max-column) edge — a full jump, not a reveal. Univer's
    // own minimal reveal (rtl-grid-mirror) has usually just made the target
    // visible; once that scroll has rendered, keep its position.
    await new Promise((resolve) => setTimeout(resolve, 60))
    if (generation !== revealGeneration) return
    let visible: IRange | null
    try {
      visible = sheet.getVisibleRange()
    } catch {
      return
    }
    if (
      visible &&
      row >= visible.startRow &&
      row <= visible.endRow &&
      column >= visible.startColumn &&
      column <= visible.endColumn
    )
      return
  }
  const aimRow = Math.max(0, row - 1)
  let scrollRow = aimRow
  let scrollColumn = column
  let lastStartRow = -1
  let lastStartColumn = -1
  sheet.scrollToCell(scrollRow, scrollColumn)
  for (let attempt = 0; attempt < 4; attempt += 1) {
    await new Promise((resolve) => setTimeout(resolve, 60))
    if (generation !== revealGeneration) return // superseded by a newer reveal
    let visible: IRange | null
    try {
      visible = sheet.getVisibleRange()
    } catch {
      return
    }
    if (!visible) return
    // over*: the viewport starts past the aim (target hidden under the pane).
    // under*: a correction clamped at 0 left the target below/right of the
    // viewport — the aim being visible is not enough, the target must be too.
    // The two cannot both be positive: overRows > 0 puts the whole viewport
    // at or past the target row.
    const overRows = Math.max(0, visible.startRow - aimRow)
    const underRows = Math.max(0, row - visible.endRow)
    const overColumns = Math.max(0, visible.startColumn - column)
    const underColumns = Math.max(0, column - visible.endColumn)
    if (overRows === 0 && overColumns === 0 && underRows === 0 && underColumns === 0) return
    // an aim inside the frozen pane (target on the first scrollable row) can
    // never be reached — the viewport clamps at the pane edge with the target
    // visible right below it, so a correction that moved nothing is done
    if (visible.startRow === lastStartRow && visible.startColumn === lastStartColumn) return
    lastStartRow = visible.startRow
    lastStartColumn = visible.startColumn
    scrollRow = Math.max(0, scrollRow - overRows + underRows)
    scrollColumn = Math.max(0, scrollColumn - overColumns + underColumns)
    sheet.scrollToCell(scrollRow, scrollColumn)
  }
}

/** The find bar's own reveal shares the broken freeze offset: re-reveal each
 *  navigated-to match (r135). The match position comes from the find service —
 *  the find bar highlights matches WITHOUT moving the active range. */
/**
 * redi's injector tracks resolution depth in a bare counter with no
 * try/finally: a construction that throws leaks one increment, and a retry
 * loop whose construction keeps throwing (each swallowed by a caller's
 * catch) walks the counter to MAX_RESOLUTIONS_QUEUED (300) — after which
 * EVERY facade call throws CircularDependencyError and the app is bricked
 * until reload (seen after add_table_column followed by a row delete).
 * Wrap the public entry points so an unwound exception restores the depth.
 */
export function installInjectorResolutionGuard(runtime: UniverRuntime): void {
  const injector = runtime.univer.__getInjector() as unknown as Record<string, unknown> & {
    resolutionOngoing?: number
  }
  for (const method of ['createInstance', 'get', 'invoke']) {
    const original = injector[method] as (...args: unknown[]) => unknown
    if (typeof original !== 'function') continue
    injector[method] = function guarded(this: unknown, ...args: unknown[]) {
      const depth = injector.resolutionOngoing ?? 0
      try {
        return original.apply(this, args)
      } catch (error) {
        if (typeof injector.resolutionOngoing === 'number') {
          injector.resolutionOngoing = depth
        }
        throw error
      }
    }
  }
}

export function installFindRevealFix(runtime: UniverRuntime): () => void {
  const injector = (
    runtime.univer as unknown as { __getInjector(): { get<T>(token: unknown): T } }
  ).__getInjector()
  let service: {
    currentMatch$: { subscribe(next: (match: unknown) => void): { unsubscribe(): void } }
  } | null
  try {
    service = injector.get(IFindReplaceService)
  } catch {
    return () => {} // find-replace not installed in this runtime
  }
  const subscription = service?.currentMatch$?.subscribe((match) => {
    const range = (match as { range?: { range?: IRange } } | null)?.range?.range
    if (!range) return
    // after the plugin's own (mis-offset) scroll settles
    window.setTimeout(() => {
      const sheet = runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
      if (!sheet) return
      void revealCellBelowFreeze(sheet, range.startRow, range.startColumn)
    }, 0)
  })
  return () => subscription?.unsubscribe()
}

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
    void revealCellBelowFreeze(sheet, coordinates.row, coordinates.column)
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
  styles: WorkbookFile['styles'],
  // The snapshot grid is padded past the used range (MINIMUM_SHEET_COLUMN_COUNT);
  // a workbook-wide <col min="1" max="16384"> must keep painting the padding.
  columnCount: number,
): Record<number, { w?: number; hd?: BooleanNumber; s?: IStyleData }> {
  const data: Record<number, { w?: number; hd?: BooleanNumber; s?: IStyleData }> = {}
  for (const columnWidth of sheet.columnWidths) {
    const endColumn = Math.min(columnWidth.endColumn, columnCount - 1)
    // Outline-only <col> entries carry no width; leave the default width.
    const width = columnWidth.width
    const pixelWidth = width === undefined ? undefined : characterWidthToPixels(width)
    // <col style=>: the default style for cells without one of their own.
    // Cell beats row beats column (OOXML order) — row-over-column needs the
    // isRowStylePrecedeColumnStyle preset flag set in App.tsx.
    const style = columnWidth.styleIndex === undefined ? undefined : styles[columnWidth.styleIndex]
    for (let column = columnWidth.startColumn; column <= endColumn; column += 1) {
      // Merge overlapping <col> spans: a later width-only span must not erase
      // an earlier span's style (and vice versa).
      data[column] = {
        ...data[column],
        ...(pixelWidth !== undefined && ((width ?? 0) > 0 || !columnWidth.hidden)
          ? { w: pixelWidth }
          : {}),
        ...(columnWidth.hidden ? { hd: BooleanNumber.TRUE } : {}),
        ...(style ? { s: toUniverStyle(style) } : {}),
      }
    }
  }
  return data
}

/**
 * The viewport's scroll anchor as scrollToCell arguments: the scroll state's
 * sheetViewStartRow/Column plus the freeze split round-trips exactly, and on
 * RTL sheets it is the only correct source — getVisibleRange().startColumn is
 * the logically LOWEST visible column there, while scrollToCell (and the lazy
 * streaming anchor) interpret the start column as the viewport's visually-left
 * (highest) logical index.
 */
export function getScrollAnchor(
  workbook: {
    getScrollStateBySheetId(
      sheetId: string,
    ): { sheetViewStartRow?: number; sheetViewStartColumn?: number } | null | undefined | void
  },
  sheet: {
    getSheetId(): string
    getSheet(): { getFreeze(): { xSplit: number; ySplit: number } }
  },
): { row: number; column: number } | null {
  try {
    const scroll = workbook.getScrollStateBySheetId(sheet.getSheetId())
    if (scroll?.sheetViewStartRow == null || scroll.sheetViewStartColumn == null) return null
    const freeze = sheet.getSheet().getFreeze()
    return {
      row: scroll.sheetViewStartRow + freeze.ySplit,
      column: scroll.sheetViewStartColumn + freeze.xSplit,
    }
  } catch {
    // No scroll render controller yet (still booting).
    return null
  }
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
    const columnSpan = visible ? visible.endColumn - visible.startColumn : 15
    // On an RTL sheet the scroll state's start column is the viewport's
    // visually-left column — the HIGHEST visible logical index.
    const startColumn = sheet.rightToLeft
      ? Math.max(0, viewportStart.column - columnSpan)
      : viewportStart.column
    visible = {
      startRow: viewportStart.row,
      endRow: viewportStart.row + (visible ? visible.endRow - visible.startRow : 79),
      startColumn,
      endColumn: startColumn + columnSpan,
    }
  }
  const range = createBufferedRange(
    normalizeVisibleRange(visible, screenRowCount, screenColumnCount),
    screenRowCount,
    screenColumnCount,
  )
  await loadRange(runtime, lazyWorkbookRef, worksheet, range, setMessage)
  await loadFrozenColumnStrip(lazyWorkbookRef, worksheet, sheet, range)
  await extendWindowPastHiddenRows(
    runtime,
    lazyWorkbookRef,
    worksheet,
    range,
    setMessage,
    screenRowCount,
  )
}

const HIDDEN_ROW_EXTEND_CELL_CAP = 2_000_000

/// A filtered sheet can hide nearly every physical row, so a physical-row
/// buffer covers almost no visible rows — and collapsing them fires no scroll
/// event, leaving the next visible rows as unloaded blanks (prod: autofilter
/// with 12 visible rows out of 11k). Grow the window geometrically until it
/// holds a viewport's worth of visible rows.
async function extendWindowPastHiddenRows(
  runtime: UniverRuntime,
  lazyWorkbookRef: { current: LazyWorkbookState | null },
  worksheet: UniverWorksheet,
  range: IRange,
  setMessage: (message: string) => void,
  screenRowCount: number,
): Promise<void> {
  const state = lazyWorkbookRef.current
  if (!state) return
  const sheetId = worksheet.getSheetId()
  const sheet = state.file.sheets.find((candidate) => candidate.id === sheetId)
  if (!sheet) return
  // Row properties stream in with indexing: right after open the hidden set
  // can still be empty even though the sheet is full of hidden rows.
  for (let attempt = 0; attempt < 20; attempt += 1) {
    if (lazyWorkbookRef.current !== state) return
    if ((state.hiddenFileRows.get(sheetId)?.size ?? 0) > 0) break
    if (!state.decorationsPendingSheets.has(sheetId)) return
    await new Promise((resolve) => setTimeout(resolve, 250))
  }
  const targetVisibleRows = range.endRow - range.startRow + 1
  let window = range
  while (lazyWorkbookRef.current === state && window.endRow < screenRowCount - 1) {
    const hidden = state.hiddenFileRows.get(sheetId)
    if (!hidden || hidden.size === 0) return
    const loadedEnd = state.loadedRanges.get(sheetId)?.endRow ?? -1
    const knownEnd = Math.min(window.endRow, loadedEnd)
    let visibleRows = 0
    for (let row = window.startRow; row <= knownEnd; row += 1) {
      if (!hidden.has(row)) visibleRows += 1
    }
    if (visibleRows >= targetVisibleRows) return
    // Budget by the column span the window actually loads, and clamp the
    // doubling to the remaining budget instead of aborting under it.
    const columnSpan = window.endColumn - window.startColumn + 1
    const maxEndRow = Math.min(
      screenRowCount - 1,
      window.startRow + Math.floor(HIDDEN_ROW_EXTEND_CELL_CAP / columnSpan) - 1,
    )
    const nextEnd = Math.min(maxEndRow, window.endRow + (window.endRow - window.startRow + 1))
    if (nextEnd <= window.endRow) return
    window = { ...window, endRow: nextEnd }
    await loadRange(runtime, lazyWorkbookRef, worksheet, window, setMessage)
    await loadFrozenColumnStrip(lazyWorkbookRef, worksheet, sheet, window)
  }
}

/// Over-budget file formulas never reach the engine (a foreign xlsx with a
/// distinct-count COUNTIF over 40k rows would freeze on open); the cell keeps
/// the file's cached value, like cache-only defined-name cells.
function degradeCostlyFormulas(
  state: LazyWorkbookState,
  sheetName: string,
  cells: WorkbookRangeResult['cells'],
): WorkbookRangeResult['cells'] {
  return degradeQuadraticFormulaCells(
    cells,
    sheetName,
    state.file.sheets.map((sheet) => ({
      name: sheet.name,
      rows: sheet.rowCount,
      columns: sheet.columnCount,
    })),
  )
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
    if (lazyWorkbookRef.current !== state) return
    state.closure.status = 'unavailable'
    // Cache-mode fallback starts NOW, not at the first edit: the file's
    // cached values may be stale or poisoned (saved by an earlier broken
    // session) and would otherwise display until the user edits something
    // (alpha ledger r141 reopen).
    queueFormulaRecalc(runtime, lazyWorkbookRef, setMessage)
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
      if (result.truncated) {
        state.recalc.engineOverBudget = true
        return giveUp()
      }
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
      // Same hazard as the pre-install guard: a structural edit landing while
      // the read was in flight shifts screen coordinates, so this file-space
      // result (and its pinned keys) would install at stale positions. Drop
      // pins from already-installed sheets too — a partial closure must not
      // keep re-applying on eviction.
      if (state.editJournal.structuralOps.size > 0) {
        state.closure.pinned.clear()
        return giveUp()
      }
      const picked = result.cells.filter((cell) => cells.has(cellKey(cell.row, cell.column)))
      recordCachedFormulaValues(state, sheetId, picked)
      const wanted = degradeCostlyFormulas(state, sheetMeta.name, picked)
      recordRowStyleKeys(state, sheetId, result.rows)
      recordHiddenFileRows(state, sheetId, result.rows, range)
      patchWorksheetRange(
        worksheet,
        undefined,
        range,
        wanted,
        state.file.styles,
        [],
        sheetMeta.tables,
        sheetMeta.pivotTables,
        sheetMeta.freeze,
        true,
        state.editJournal,
        undefined,
        undefined,
        undefined,
        sheetRowColStyleKeys(state, sheetId),
        inheritedWrapLookup(state.file.styles, result.rows, sheetMeta.columnWidths),
        hiddenRowsInfo(state, sheetId),
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

/// On-demand precedent pinning for formulas written onto a streamed
/// workbook: the referenced file-sheet cells are fetched (screen-mapped
/// through journaled structural ops), installed into the engine as cached
/// values, and added to the closure pinned map so viewport eviction
/// re-applies them — the new formula then computes correctly and STAYS
/// correct, exactly like closure-mode formulas. The session cell budget was
/// already checked at propose time (streamedPinBudgetError); returns false
/// when a sidecar read fails so the caller can abort before writing a
/// formula that would evaluate against partial data.
/// One source cell for a direct (grid-bypassing) copy read.
export interface RawCopyCell {
  v: string | number | boolean | null
  t: number | null
  s: Record<string, unknown> | null
  /// model-equivalent formula (a session edit); carried verbatim like grid f
  f: string | null
  /// the FILE's formula text for a value-only install (streamed mode) — the
  /// caller decides whether it can be carried (#1123 budget/pinning)
  fileFormula: string | null
}

/**
 * Streamed unfiltered copy_range source reader: consumes the sidecar's
 * mapped payload plus the journal overlay directly instead of installing
 * every source chunk into the Univer grid and reading it back — the install
 * plus triple facade read dominated bulk-copy time and transient renderer
 * memory. Value/type mapping mirrors patchWorksheetRangeInner ('' stays an
 * empty STRING, rich cells collapse to their full text, numbers/booleans
 * stay bare); a cell's own xf resolves through toUniverStyle exactly like
 * the copy's style-pool handling. Returns null when indexing does not catch
 * up in time — the caller falls back to the grid path.
 */
export async function readCopySourceDirect(
  lazyWorkbookRef: { current: LazyWorkbookState | null },
  sheetId: string,
  bounds: IRange,
  setMessage: (message: string) => void,
): Promise<RawCopyCell[][] | null> {
  const state = lazyWorkbookRef.current
  if (!state) return null
  const sheetMeta = state.file.sheets.find((candidate) => candidate.id === sheetId)
  if (!sheetMeta) return null
  const journal = state.editJournal
  const journalCells = journal.cells.get(sheetId)
  // File formula text is file-space; structural edits shift coordinates, so
  // carrying it would offset the wrong references (same refusal as
  // indexedFormulaText — the copy degrades to values, like the grid path).
  const fileFormulasReliable = (journal.structuralOps.get(sheetId)?.length ?? 0) === 0
  const width = bounds.endColumn - bounds.startColumn + 1
  const rows = bounds.endRow - bounds.startRow + 1
  const out: RawCopyCell[][] = Array.from({ length: rows }, () =>
    Array.from({ length: width }, () => ({
      v: null,
      t: null,
      s: null,
      f: null,
      fileFormula: null,
    })),
  )
  const chunkRows = Math.max(1, Math.floor(SIDECAR_RANGE_CELL_LIMIT / width))
  for (let startRow = bounds.startRow; startRow <= bounds.endRow; startRow += chunkRows) {
    if (lazyWorkbookRef.current !== state) return null
    const chunk: IRange = {
      startRow,
      endRow: Math.min(bounds.endRow, startRow + chunkRows - 1),
      startColumn: bounds.startColumn,
      endColumn: bounds.endColumn,
    }
    if (rows > chunkRows) {
      setMessage(
        t('appStreamingRows', {
          name: state.file.name,
          rows: `${(startRow - bounds.startRow).toLocaleString()} / ${rows.toLocaleString()}`,
        }),
      )
    }
    // Background indexing may not have reached these rows yet — poll like
    // the precedent pinner; on the deadline the caller falls back.
    const deadline = Date.now() + 120_000
    let mapped: MappedRangeRead | null
    for (;;) {
      try {
        mapped = await readSheetRangeMapped(state, sheetId, chunk, sheetMeta)
      } catch {
        return null
      }
      if (lazyWorkbookRef.current !== state) return null
      if (
        mapped === null ||
        mapped.raw.indexingComplete ||
        (mapped.raw.indexedThroughRow !== null && mapped.raw.indexedThroughRow >= mapped.fileEndRow)
      ) {
        break
      }
      if (Date.now() > deadline) return null
      await new Promise((resolve) => setTimeout(resolve, 400))
    }
    // mapped === null: the chunk is entirely journal-owned (inserted this
    // session) — the journal overlay below is the whole content.
    for (const cell of mapped?.screen.cells ?? []) {
      if (
        cell.row < chunk.startRow ||
        cell.row > chunk.endRow ||
        cell.column < chunk.startColumn ||
        cell.column > chunk.endColumn
      ) {
        continue
      }
      const target = out[cell.row - bounds.startRow]?.[cell.column - bounds.startColumn]
      if (!target) continue
      const style = cell.styleIndex === undefined ? undefined : state.file.styles[cell.styleIndex]
      if (style) target.s = toUniverStyle(style) as unknown as Record<string, unknown>
      if (cell.formula && fileFormulasReliable) target.fileFormula = cell.formula
      const value = cell.value
      if (typeof value === 'string') {
        target.v = value
        target.t = CellValueType.STRING
      } else if (value !== null && value !== undefined) {
        target.v = value
      }
    }
  }
  // The journal overlay wins over file content, exactly like the grid's
  // applyJournalOverlay: session values/formulas replace the record, and a
  // journaled style patches on top of the file style.
  if (journalCells || journal.bulkConstantFills.get(sheetId)?.length) {
    for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
      for (let column = bounds.startColumn; column <= bounds.endColumn; column += 1) {
        const entry = journalCells?.get(`${row}:${column}`)
        const target = out[row - bounds.startRow]?.[column - bounds.startColumn]
        if (!target) continue
        if (entry?.styleReset) target.s = null
        if (entry?.style) {
          target.s = {
            ...(target.s ?? {}),
            ...(fromNeutralStyle(entry.style) as Record<string, unknown>),
          }
        }
        const content = journalCellContentAt(journal, sheetId, row, column)
        if (!content.found) continue
        target.fileFormula = null
        if (content.formula) {
          target.f = content.formula
          target.v = null
          target.t = null
        } else {
          target.f = null
          const value = content.value
          if (typeof value === 'string') {
            target.v = value
            target.t = CellValueType.STRING
          } else {
            target.v = value ?? null
            target.t = null
          }
        }
      }
    }
  }
  return out
}

export async function pinStreamedPrecedents(
  runtime: UniverRuntime,
  lazyWorkbookRef: { current: LazyWorkbookState | null },
  needs: ReadonlyMap<string, ReadonlySet<number>>,
): Promise<boolean> {
  const state = lazyWorkbookRef.current
  if (!state || state.formulaMode) return true
  const workbook = runtime.univerAPI.getActiveWorkbook()
  if (!workbook) return false
  for (const [sheetId, keys] of needs) {
    const worksheet = workbook.getSheetBySheetId(sheetId)
    const sheetMeta = state.file.sheets.find((candidate) => candidate.id === sheetId)
    if (!worksheet || !sheetMeta || keys.size === 0) continue
    let pinned = state.closure.pinned.get(sheetId)
    if (!pinned) {
      pinned = new Map()
      state.closure.pinned.set(sheetId, pinned)
    }
    const missing = new Set<number>()
    for (const key of keys) {
      if (!pinned.has(`${keyRow(key)}:${keyColumn(key)}`)) missing.add(key)
    }
    if (missing.size === 0) continue
    for (const range of closureFetchRanges(missing)) {
      // Background indexing may not have reached these rows yet — pinning a
      // partial read would make the formula present a truncated result as
      // success, the exact failure this pinning exists to prevent. Poll
      // until the sidecar has indexed past the requested rows; fail closed
      // on the deadline (the caller aborts before writing the formula).
      const deadline = Date.now() + 60_000
      let mapped: MappedRangeRead | null
      for (;;) {
        try {
          mapped = await readSheetRangeMapped(state, sheetId, range, sheetMeta)
        } catch {
          return false
        }
        if (lazyWorkbookRef.current !== state) return false
        if (
          mapped === null ||
          mapped.raw.indexingComplete ||
          (mapped.raw.indexedThroughRow !== null &&
            mapped.raw.indexedThroughRow >= mapped.fileEndRow)
        ) {
          break
        }
        if (Date.now() > deadline) return false
        await new Promise((resolve) => setTimeout(resolve, 400))
      }
      // Entirely journal-owned region (inserted this session): the grid
      // already holds its cells — nothing streams in, nothing to pin.
      if (!mapped) continue
      recordRowStyleKeys(state, sheetId, mapped.screen.rows)
      recordHiddenFileRows(state, sheetId, mapped.screen.rows, range)
      // Cached values only (useFormulas=false), like every streamed install:
      // a precedent formula cell keeps its Excel-cached value — making it
      // live would require pinning ITS precedents transitively.
      patchWorksheetRange(
        worksheet,
        undefined,
        range,
        mapped.screen.cells,
        state.file.styles,
        mapped.screen.hyperlinks,
        sheetMeta.tables,
        sheetMeta.pivotTables,
        sheetMeta.freeze,
        false,
        state.editJournal,
        pinned,
        state.recalc.overlay.get(sheetId),
        undefined,
        sheetRowColStyleKeys(state, sheetId),
        inheritedWrapLookup(state.file.styles, mapped.screen.rows, sheetMeta.columnWidths),
        hiddenRowsInfo(state, sheetId),
      )
      for (const cell of mapped.screen.cells) {
        const key = cellKey(cell.row, cell.column)
        if (!missing.has(key)) continue
        pinned.set(`${cell.row}:${cell.column}`, { v: cell.value ?? null })
      }
    }
  }
  return true
}

export interface MappedRangeRead {
  /// Result arrays translated into screen coordinates.
  readonly screen: Pick<WorkbookRangeResult, 'cells' | 'rows' | 'merges' | 'hyperlinks'>
  readonly raw: WorkbookRangeResult
  readonly indexedThroughScreen: number | null
  /// File-space end row actually requested; the indexing poll compares the
  /// raw cutoff against this.
  readonly fileEndRow: number
}

/// Per-request cell budget for sidecar range reads; row batches are sized to
/// stay under it (just below MAX_RANGE_CELLS in shared/desktop-api.ts).
const SIDECAR_READ_BATCH_CELLS = 90_000

/// Reads a screen-space range, translating through the sheet's journaled
/// structural operations. Returns null when the range is entirely
/// journal-owned (inserted this session — nothing streams into it). A
/// request spanning deleted file rows — or simply a large one, like the
/// buffered viewport at far zoom-out — can exceed the sidecar's per-read
/// cell budget, so reads are split into row batches.
export async function readSheetRangeMapped(
  state: LazyWorkbookState,
  sheetId: string,
  screenRange: IRange,
  sheet: WorkbookFile['sheets'][number],
): Promise<MappedRangeRead | null> {
  const ops = state.editJournal.structuralOps.get(sheetId) ?? []
  if (ops.length === 0) {
    const width = screenRange.endColumn - screenRange.startColumn + 1
    const batchRows = Math.max(1, Math.floor(SIDECAR_READ_BATCH_CELLS / width))
    const cells: WorkbookRangeResult['cells'] = []
    const rows: WorkbookRangeResult['rows'] = []
    const merges: WorkbookRangeResult['merges'] = []
    const hyperlinks: WorkbookRangeResult['hyperlinks'] = []
    let raw: WorkbookRangeResult | null = null
    for (
      let startRow = screenRange.startRow;
      startRow <= screenRange.endRow;
      startRow += batchRows
    ) {
      const endRow = Math.min(startRow + batchRows - 1, screenRange.endRow)
      const batch = await window.desktopApi.readWorkbookRange({
        sessionId: state.file.sessionId,
        sheetId,
        range: { ...screenRange, startRow, endRow },
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
    if (!raw) {
      // Degenerate empty range (endRow < startRow): let the sidecar answer,
      // preserving the pre-batching behavior for out-of-contract input.
      raw = await window.desktopApi.readWorkbookRange({
        sessionId: state.file.sessionId,
        sheetId,
        range: screenRange,
      })
    }
    return {
      screen: { ...raw, cells, rows, merges, hyperlinks },
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
  const batchRows = Math.max(1, Math.floor(SIDECAR_READ_BATCH_CELLS / width))
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

// The command id lives in undo-carry.ts: cross-save carrying must truncate
// at these steps (their params are registry tokens resolving to closures
// over the pre-save session state).
const visualUndoRegistry = new Map<number, VisualUndoStep>()
let visualUndoSequence = 0
const visualUndoRuntimes = new WeakSet<object>()

/// Appends a registry step to the undo entry a Univer command just pushed, so
/// ONE ⌘Z reverts the whole user action (cells + shadow journal op) instead of
/// needing a second, visually-inert undo press — and no extra undo-carry
/// truncation point is created (alpha ledger r124 / bugbot). Falls back to a
/// standalone entry when the stack is empty.
/// The current top undo element (opaque identity token): callers snapshot it
/// BEFORE running a Univer command, so attachVisualUndoToLastStep can tell a
/// freshly pushed entry from a stale one (a no-op command pushes nothing, and
/// attaching to whatever was already on top would bind the step to an
/// unrelated edit — bugbot).
export function topUndoElement(runtime: UniverRuntime): unknown {
  try {
    const injector = (
      runtime.univer as unknown as {
        __getInjector(): { get<T>(token: unknown): T }
      }
    ).__getInjector()
    return injector.get<{ pitchTopUndoElement(): unknown }>(IUndoRedoService).pitchTopUndoElement()
  } catch {
    return null
  }
}

export function attachVisualUndoToLastStep(
  runtime: UniverRuntime,
  step: VisualUndoStep,
  /// topUndoElement() snapshot taken before the command this step shadows;
  /// when the top is unchanged the step gets its own standalone entry instead
  notThisElement?: unknown,
): void {
  const unitId = runtime.univerAPI.getActiveWorkbook()?.getId()
  if (!unitId) return
  const injector = (
    runtime.univer as unknown as {
      __getInjector(): { get<T>(token: unknown): T }
    }
  ).__getInjector()
  ensureVisualUndoCommand(injector, runtime)
  const token = ++visualUndoSequence
  visualUndoRegistry.set(token, step)
  const mutation = (direction: 'undo' | 'redo') => ({
    id: VISUAL_UNDO_COMMAND_ID,
    params: { token, direction },
  })
  const service = injector.get<{
    pitchTopUndoElement(): {
      unitID: string
      undoMutations: { id: string; params: unknown }[]
      redoMutations: { id: string; params: unknown }[]
    } | null
    pushUndoRedo(item: {
      unitID: string
      undoMutations: { id: string; params: unknown }[]
      redoMutations: { id: string; params: unknown }[]
    }): void
  }>(IUndoRedoService)
  const top = service.pitchTopUndoElement()
  if (top && top.unitID === unitId && top !== notThisElement) {
    top.undoMutations.push(mutation('undo'))
    top.redoMutations.push(mutation('redo'))
    return
  }
  service.pushUndoRedo({
    unitID: unitId,
    undoMutations: [mutation('undo')],
    redoMutations: [mutation('redo')],
  })
}

/// Interactive visual ops (chart edits, moves, deletes, inserts) enter
/// Univer's own undo stack as a custom mutation pair, so ⌘Z interleaves
/// them correctly with cell edits.
function ensureVisualUndoCommand(
  injector: { get<T>(token: unknown): T },
  runtime: UniverRuntime,
): void {
  if (visualUndoRuntimes.has(runtime)) return
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

export function pushVisualUndo(runtime: UniverRuntime, step: VisualUndoStep): void {
  const unitId = runtime.univerAPI.getActiveWorkbook()?.getId()
  if (!unitId) return
  const injector = (
    runtime.univer as unknown as {
      __getInjector(): { get<T>(token: unknown): T }
    }
  ).__getInjector()
  ensureVisualUndoCommand(injector, runtime)
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
  for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
    for (let column = bounds.startColumn; column <= bounds.endColumn; column += 1) {
      const fill = bulkConstantFillValueAt(state.editJournal, sheetId, row, column)
      if (fill.found) cells.set(`${row}:${column}`, fill.value)
    }
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
  for (let row = bounds.startRow; row <= bounds.endRow; row += 1) {
    for (let column = bounds.startColumn; column <= bounds.endColumn; column += 1) {
      const fill = bulkConstantFillValueAt(state.editJournal, sheetId, row, column)
      if (fill.found) cells.set(`${row}:${column}`, fill.value)
    }
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
/// a busy rejection re-queues after the in-flight import has had a chance
/// to move on; rejections are cheap, so a modest cadence is fine
const RECALC_BUSY_RETRY_MS = 5000
const RECALC_READ_BUDGET = 20_000
const RECALC_MAX_EDITS = 10_000
/// transient sidecar hiccups retry on the next edit; repeated rejection of
/// this workbook disables the fallback for the session
export const RECALC_MAX_FAILURES = 3
/// The fallback cold-imports the whole file into IronCalc at ~460 bytes per
/// grid cell of transient RSS (a 31MB / 8.7M-cell pure-data upload burned
/// ~4GB for minutes at open) — and the formula-count gate never fires on
/// such files. Cap by grid cell count (the dominant cost) and by compressed
/// size (catches string-heavy files with small grids): past either, cached
/// values stand for the session, same as engineOverBudget after a truncated
/// formula index. 2M cells ≈ 1GB import.
export const RECALC_MAX_FILE_BYTES = 64 * 1024 * 1024
export const RECALC_MAX_GRID_CELLS = 2_000_000

/// Decided once at open (engineOverBudget seed); the recalc queue re-checks
/// nothing size-related after that.
export function recalcOverBudgetAtOpen(
  fileBytes: number | undefined,
  gridCellCount: number,
): boolean {
  return (fileBytes ?? 0) > RECALC_MAX_FILE_BYTES || gridCellCount > RECALC_MAX_GRID_CELLS
}

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
  // minimal states (tests, partial teardown) may carry no recalc slot
  if (!state.recalc || state.recalc.failures >= RECALC_MAX_FAILURES) return
  if (state.recalc.engineOverBudget) return
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
    state.recalc.engineOverBudget = true
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
  // Viewport harvesting on a million-formula sheet would grow without bound;
  // dropping the store only costs formula-bar text for long-unseen cells.
  if (bySheet.size > 200_000) bySheet.clear()
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
  state.recalc.running = true
  state.recalc.lastRunAt = Date.now()
  try {
    const keys = await recalcFormulaCellKeys(state, sheetId)
    if (lazyWorkbookRef.current !== state) return
    if (!keys || keys.size === 0) {
      // nothing to recalculate on this sheet: mark it followed so the
      // finished-run chain below doesn't loop on it forever
      if (keys) state.recalc.follow.set(sheetId, { anchorRow: 0, complete: true })
      return
    }
    const viewportStartRow = state.loadedRanges.get(sheetId)?.startRow ?? 0
    const reads = recalcReadRanges(keys, viewportStartRow, RECALC_READ_BUDGET)
    if (reads.length === 0) {
      state.recalc.follow.set(sheetId, { anchorRow: viewportStartRow, complete: true })
      return
    }
    const windowComplete = reads.length === closureFetchRanges(keys).length
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
    const formulaTextBySheet = state.formulaText.get(sheetId)
    for (const cell of result.cells) {
      if (cell.sheetId !== sheetId || !cell.isFormula) continue
      // The user's own journaled edits stay authoritative on screen.
      if (journalCells?.has(`${cell.row}:${cell.column}`)) continue
      // #NAME? flags a function IronCalc lacks; the file's cached value is
      // better — keep it. Same for locale-dependent results the engine's en
      // locale can never reproduce.
      if (cell.formatted === '#NAME?') {
        unsupported += 1
        continue
      }
      const formulaText = formulaTextBySheet?.get(`${cell.row}:${cell.column}`)
      if (formulaText && usesLocaleDependentFunction(formulaText)) {
        unsupported += 1
        continue
      }
      overlay.set(`${cell.row}:${cell.column}`, { v: cell.number ?? cell.formatted })
    }
    state.recalc.overlay.set(sheetId, overlay)
    state.recalc.follow.set(sheetId, { anchorRow: viewportStartRow, complete: windowComplete })
    const loaded = state.loadedRanges.get(sheetId)
    if (loaded && overlay.size > 0) {
      journalSuppression.active = true
      loadAutoHeightSuppression.active = true
      try {
        applyPinnedOverlay(worksheet, overlay, undefined, loaded)
      } finally {
        journalSuppression.active = false
        loadAutoHeightSuppression.active = false
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
  } catch (error: unknown) {
    // Fail soft: cached values stay on screen and the save still asks Excel
    // to recalculate on open. Repeated failures disable the fallback — but
    // only the current run may count (mirrors the success path's guard):
    // a superseded run failing after a newer success must not stack stale
    // failures toward the kill switch. "Busy" is not a failure — another
    // recalculation holds the worker and this one simply retries later.
    if (lazyWorkbookRef.current !== state || state.recalc.generation !== generation) return
    const busy = error instanceof Error && error.message.includes('busy with another recalculation')
    if (!busy) state.recalc.failures += 1
    // The in-flight import that rejected this run may itself be superseded
    // (its success is discarded under the generation guard), so a busy pass
    // must re-queue or the workbook keeps cached values after the edit.
    if (busy) {
      setTimeout(() => {
        if (lazyWorkbookRef.current !== state) return
        queueFormulaRecalc(runtime, lazyWorkbookRef, setMessage)
      }, RECALC_BUSY_RETRY_MS)
    }
  } finally {
    if (state.recalc.generation === generation) {
      state.recalc.running = false
      // A sheet switched to while this run was in flight got its first-follow
      // queue attempt swallowed by the running/cooldown guards, and nothing
      // rechecks a fully loaded viewport — chain one run for it. STRICTLY a
      // different sheet: chaining on the run's own sheet would turn failures
      // and incomplete indexes into an immediate retry loop that burns the
      // RECALC_MAX_FAILURES budget before any edit (bugbot).
      try {
        const activeId = runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId()
        if (
          activeId &&
          activeId !== sheetId &&
          lazyWorkbookRef.current === state &&
          state.closure.status === 'unavailable' &&
          state.recalc.failures < RECALC_MAX_FAILURES &&
          !state.recalc.follow.has(activeId)
        ) {
          queueFormulaRecalc(runtime, lazyWorkbookRef, setMessage)
        }
      } catch {
        /* workbook mid-teardown */
      }
    }
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
    if (state.formulaMode) recordCachedFormulaValues(state, sheetId, mapped.screen.cells)
    const stripPatchRange = { ...stripRange, endRow: availableEndRow }
    recordRowStyleKeys(state, sheetId, mapped.screen.rows)
    recordHiddenFileRows(state, sheetId, mapped.screen.rows, stripPatchRange)
    patchWorksheetRange(
      worksheet,
      undefined,
      stripPatchRange,
      state.formulaMode
        ? degradeCostlyFormulas(state, sheet.name, mapped.screen.cells)
        : mapped.screen.cells,
      state.file.styles,
      mapped.screen.hyperlinks,
      sheet.tables,
      sheet.pivotTables,
      sheet.freeze,
      state.formulaMode,
      state.editJournal,
      state.closure.pinned.get(sheetId),
      state.recalc.overlay.get(sheetId),
      undefined,
      sheetRowColStyleKeys(state, sheetId),
      inheritedWrapLookup(state.file.styles, mapped.screen.rows, sheet.columnWidths),
      hiddenRowsInfo(state, sheetId),
    )
    const stripQualifying = wrapAutoFitRows(
      mapped.screen.cells,
      state.file.styles,
      mapped.screen.rows,
      sheet.columnWidths,
      sheet.defaultRowHeightFixed,
      sheet.defaultRowHeight,
      stripPatchRange,
      mapped.screen.merges,
    )
    measureWrapAutoFitRows(worksheet, stripQualifying)
    trackPreIndexMeasuredRows(`file-${state.file.sha256}:${sheetId}`, stripQualifying)
  } catch {
    state.frozenStripKeys.delete(sheetId)
  }
}

/// How many consecutive no-progress polls (250ms apart) a blocking AI load
/// tolerates before giving up. Progress resets the counter: a load keeps
/// waiting as long as background indexing is still advancing toward the
/// requested rows, however long that takes — only a stalled stream fails.
const INDEX_WAIT_STALL_LIMIT = 40

/// Pure wait policy for blocking loads on a still-indexing workbook: returns
/// the next stall count, or null to give up. Any forward movement of the
/// indexer resets the count; only consecutive stalls accumulate.
export function nextIndexWaitStall(
  stalls: number,
  lastIndexedRow: number | null,
  indexedThroughRow: number | null,
): number | null {
  const progressed =
    indexedThroughRow !== null && (lastIndexedRow === null || indexedThroughRow > lastIndexedRow)
  const next = progressed ? 0 : stalls + 1
  return next >= INDEX_WAIT_STALL_LIMIT ? null : next
}

async function loadRange(
  runtime: UniverRuntime,
  lazyWorkbookRef: { current: LazyWorkbookState | null },
  worksheet: UniverWorksheet,
  range: IRange,
  setMessage: (message: string) => void,
  isRetry = false,
  waitForRequestedRange = false,
  waitStalls = 0,
  lastIndexedRow: number | null = null,
  bulk = false,
): Promise<void> {
  const state = lazyWorkbookRef.current
  if (!state) return
  const sheetId = worksheet.getSheetId()
  const loaded = state.loadedRanges.get(sheetId)
  // A range loaded before the sidecar finished indexing came without the
  // sheet's decorations (conditional formats, filters, validations) — those
  // only exist post-indexing. Keep reading until a post-indexing result
  // lands, or the decorations starve forever behind this early return.
  if (
    !isRetry &&
    loaded &&
    containsRange(loaded, range) &&
    !state.decorationsPendingSheets.has(sheetId)
  ) {
    return
  }
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
      // The whole range is journal-owned (inserted this session): nothing
      // streams into it, but cells written there may have been evicted by an
      // earlier window move — replay them from the journal, or reads see
      // ghosts of empty cells where this session's edits live. Route this
      // through the normal window patcher so the previous chunk is evicted:
      // otherwise a whole-column fill leaves every visited 20k-row strip
      // resident in Univer and can grow the renderer past 2 GiB.
      patchWorksheetRange(
        worksheet,
        loaded,
        range,
        [],
        state.file.styles,
        [],
        sheetMeta.tables,
        sheetMeta.pivotTables,
        sheetMeta.freeze,
        state.formulaMode,
        state.editJournal,
        state.closure.pinned.get(sheetId),
        state.recalc.overlay.get(sheetId),
        undefined,
        undefined,
        undefined,
        hiddenRowsInfo(state, sheetId),
      )
      state.loadedRanges.set(sheetId, range)
      return
    }
    const availableEndRow =
      mapped.indexedThroughScreen === null
        ? null
        : Math.min(mapped.indexedThroughScreen, range.endRow)
    let patchedRange: IRange | undefined
    if (availableEndRow !== null && availableEndRow >= range.startRow) {
      const availableRange = { ...range, endRow: availableEndRow }
      const alreadyLoaded = state.loadedRanges.get(sheetId)
      if (!alreadyLoaded || !containsRange(alreadyLoaded, availableRange)) {
        if (state.formulaMode) recordCachedFormulaValues(state, sheetId, mapped.screen.cells)
        recordRowStyleKeys(state, sheetId, mapped.screen.rows)
        recordHiddenFileRows(state, sheetId, mapped.screen.rows, availableRange)
        patchWorksheetRange(
          worksheet,
          alreadyLoaded,
          availableRange,
          state.formulaMode
            ? degradeCostlyFormulas(state, sheetMeta.name, mapped.screen.cells)
            : mapped.screen.cells,
          state.file.styles,
          mapped.screen.hyperlinks,
          sheetMeta.tables,
          sheetMeta.pivotTables,
          sheetMeta.freeze,
          state.formulaMode,
          state.editJournal,
          state.closure.pinned.get(sheetId),
          state.recalc.overlay.get(sheetId),
          undefined,
          sheetRowColStyleKeys(state, sheetId),
          inheritedWrapLookup(state.file.styles, mapped.screen.rows, sheetMeta.columnWidths),
          hiddenRowsInfo(state, sheetId),
        )
        state.loadedRanges.set(sheetId, availableRange)
        patchedRange = availableRange
        // Partial recalc windows follow the user: far outside the anchored
        // window the grid would show raw (possibly poisoned) file cache —
        // re-run the sidecar recalc around the new viewport (r141 reopen).
        const recalc = state.recalc
        const follow = recalc.follow.get(sheetId)
        if (
          !state.formulaMode &&
          state.closure.status === 'unavailable' &&
          !recalc.running &&
          // a pending debounce counts as in flight: re-queuing would RESET
          // the timer and starve the open-time recalc while streaming keeps
          // patching every ~250ms (bugbot)
          recalc.timer === null &&
          Date.now() - recalc.lastRunAt > 3000 &&
          (!follow ||
            (!follow.complete && Math.abs(availableRange.startRow - follow.anchorRow) > 200))
        ) {
          queueFormulaRecalc(runtime, lazyWorkbookRef, setMessage)
        }
      }
    }
    const result = mapped.raw
    const hasStructuralOps = (state.editJournal.structuralOps.get(sheetId)?.length ?? 0) > 0
    // Every mode needs formulaText for the formula bar: formulaMode for
    // cache-only (defined-name) cells, value mode because streamed cells
    // carry no `f` in the grid at all — without the harvest a formula cell
    // shows only its cached value. Harvest from screen.cells — raw holds
    // only the last batch of an over-cap read — except under structural
    // ops, where the store must stay in file coordinates.
    storeFormulaText(state, sheetId, hasStructuralOps ? result.cells : mapped.screen.cells)
    recordHyperlinks(state, sheetId, mapped.screen.hyperlinks)
    keepActiveSheet(worksheet, () => {
      applyRowProperties(worksheet, state, sheetId, mapped.screen.rows)
      applyMerges(worksheet, state, sheetId, mapped.screen.merges)
    })
    // After merges (merged-only rows never auto-fit) and stored heights.
    // Bulk AI reads skip it: the measure is a real canvas text layout over
    // every wrap row × the full sheet width, and only non-final chunks load
    // as bulk — their windows are evicted by the next chunk anyway.
    // The stale-height reset must not hide behind the patch gate: the
    // post-indexing retry that finally carries the sheet's merges usually
    // finds its range already loaded and skips the patch entirely.
    const wrapWindow = patchedRange ?? (result.indexingComplete ? range : undefined)
    if (wrapWindow && !bulk) {
      const qualifying = wrapAutoFitRows(
        mapped.screen.cells,
        state.file.styles,
        mapped.screen.rows,
        sheetMeta.columnWidths,
        sheetMeta.defaultRowHeightFixed,
        sheetMeta.defaultRowHeight,
        wrapWindow,
        mapped.screen.merges,
      )
      if (patchedRange) measureWrapAutoFitRows(worksheet, qualifying)
      const sheetKey = `file-${state.file.sha256}:${sheetId}`
      if (result.indexingComplete) {
        const qualifyingWithoutMerges = wrapAutoFitRows(
          mapped.screen.cells,
          state.file.styles,
          mapped.screen.rows,
          sheetMeta.columnWidths,
          sheetMeta.defaultRowHeightFixed,
          sheetMeta.defaultRowHeight,
          wrapWindow,
        )
        resetStaleWrapAutoHeights(
          runtime,
          `file-${state.file.sha256}`,
          worksheet,
          takeContaminatedRows(sheetKey, qualifyingWithoutMerges, qualifying),
          mapped.screen.rows,
          sheetMeta.defaultRowHeight,
        )
      } else {
        trackPreIndexMeasuredRows(sheetKey, qualifying)
      }
    }
    // Conditional formatting, filters, and validations install once with
    // file-space ranges; Univer shifts the installed models itself on later
    // structural edits, but a fresh install after a shift would be stale —
    // skip it (rare: the sheet was being edited before it first rendered).
    if (!hasStructuralOps) {
      await applyConditionalRules(worksheet, state, sheetId, result.conditionalRules)
      if (result.indexingComplete) {
        applySheetFilter(worksheet, state, sheetId, result.autoFilter)
        applyDataValidations(runtime, state, sheetId, result.dataValidations)
        state.decorationsPendingSheets.delete(sheetId)
      } else {
        state.decorationsPendingSheets.add(sheetId)
      }
    } else {
      // Structurally-edited sheets never install decorations (see above), so
      // a pending entry would only force pointless sidecar re-reads.
      state.decorationsPendingSheets.delete(sheetId)
    }
    if (result.indexingComplete) captureSheetFileState(state, sheetId, result)
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
      const nextStalls =
        waitForRequestedRange && (availableEndRow === null || availableEndRow < range.endRow)
          ? nextIndexWaitStall(waitStalls, lastIndexedRow, result.indexedThroughRow)
          : null
      if (waitForRequestedRange && nextStalls !== null) {
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
            nextStalls,
            result.indexedThroughRow,
            bulk,
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
  bulk = false,
): Promise<boolean> {
  const initialState = lazyWorkbookRef.current
  if (!initialState) return false
  const extent = lazySheetScreenExtent(initialState, worksheet.getSheetId())
  if (
    !extent ||
    range.startRow < 0 ||
    range.startColumn < 0 ||
    range.endRow >= extent.rows ||
    range.endColumn >= extent.columns
  ) {
    return false
  }
  await loadRange(
    runtime,
    lazyWorkbookRef,
    worksheet,
    range,
    setMessage,
    false,
    true,
    0,
    null,
    bulk,
  )
  const state = lazyWorkbookRef.current
  const loaded = state?.loadedRanges.get(worksheet.getSheetId())
  return state === initialState && loaded !== undefined && containsRange(loaded, range)
}

/// readWorkbookRange's protocol cap (MAX_RANGE_CELLS in desktop-api.ts):
/// bigger chunks mean fewer sidecar round-trips and evict/reinstall cycles
/// for bulk ops; 100k keeps a chunk's JSON payload around 5MB.
const SIDECAR_RANGE_CELL_LIMIT = 100_000

/**
 * Applies a bulk edit (fill / large clear) over a possibly-unloaded range of
 * a streamed workbook in sidecar-request-sized chunks: each chunk is loaded
 * into the grid first — so the edit lands on real cells, goes through the
 * normal undoable commands, and the edit journal records it — then edited.
 * Loading a chunk evicts the previous window, but the already-written chunks
 * live in the journal and re-apply whenever their region streams back in.
 * Fully-preloaded workbooks take a single unchunked pass.
 *
 * `neighborColumns` widens each chunk load to the sheet's full width so
 * written formulas with row-local references (=B2*2 in column A) compute
 * against real neighbor values. Value-only edits (constant fills, clears,
 * text replaces) must pass false: loading only the target columns keeps the
 * chunk count minimal, and a target inside rows/columns inserted this
 * session becomes entirely journal-owned — loaded instantly with no sidecar
 * read, independent of how far background indexing has gotten.
 */
export async function applyRangeInLoadedChunks(
  runtime: UniverRuntime,
  lazyWorkbookRef: { current: LazyWorkbookState | null },
  worksheet: UniverWorksheet,
  bounds: IRange,
  applyChunk: (chunk: IRange) => void,
  setMessage: (message: string) => void,
  options?: { neighborColumns?: boolean },
): Promise<void> {
  const state = lazyWorkbookRef.current
  if (!state) throw new Error('No workbook is open.')
  const sheetId = worksheet.getSheetId()
  // Extents are screen-space: chunk coordinates already sit in the current
  // session's shifted space, so the clamp must too (a file-space clamp would
  // cut off / overshoot the last rows and columns after insert/delete ops).
  const extent = lazySheetScreenExtent(state, sheetId)
  const width = bounds.endColumn - bounds.startColumn + 1
  // Degenerate ultra-wide sheets fall back to the target columns even when
  // neighbors were requested, to keep the chunk count bounded.
  const fullWidth = (options?.neighborColumns ?? true) && extent !== null && extent.columns <= 200
  const loadStartColumn = fullWidth ? 0 : bounds.startColumn
  const loadEndColumn = fullWidth && extent !== null ? extent.columns - 1 : bounds.endColumn
  const loadWidth = Math.max(width, loadEndColumn - loadStartColumn + 1)
  const chunkRows = state.flags.preloadComplete
    ? bounds.endRow - bounds.startRow + 1
    : Math.max(1, Math.floor(SIDECAR_RANGE_CELL_LIMIT / loadWidth))
  const totalRows = bounds.endRow - bounds.startRow + 1
  // The final chunk's window stays resident after the loop (the viewport
  // refetch below early-returns when that window already contains the view),
  // so only non-final chunks may skip the wrap auto-fit measure (bugbot).
  const lastLoadRow = extent ? Math.min(bounds.endRow, extent.rows - 1) : bounds.endRow
  try {
    for (let startRow = bounds.startRow; startRow <= bounds.endRow; startRow += chunkRows) {
      const chunk: IRange = {
        startRow,
        endRow: Math.min(bounds.endRow, startRow + chunkRows - 1),
        startColumn: bounds.startColumn,
        endColumn: bounds.endColumn,
      }
      if (totalRows > chunkRows && isActiveSheet(runtime, sheetId)) {
        setMessage(
          t('appStreamingRows', {
            name: state.file.name,
            rows: `${(startRow - bounds.startRow).toLocaleString()} / ${totalRows.toLocaleString()}`,
          }),
        )
      }
      if (!state.flags.preloadComplete && extent) {
        // Clamp the load to the data extent — rows/columns beyond it have
        // nothing to stream in and ensureLazyRangeLoaded would reject them.
        const load: IRange = {
          startRow: chunk.startRow,
          endRow: Math.min(chunk.endRow, extent.rows - 1),
          startColumn: loadStartColumn,
          endColumn: Math.min(loadEndColumn, extent.columns - 1),
        }
        const inExtent = load.startRow <= load.endRow && load.startColumn <= load.endColumn
        const loaded = state.loadedRanges.get(sheetId)
        if (inExtent && (!loaded || !containsRange(loaded, load))) {
          const ok = await ensureLazyRangeLoaded(
            runtime,
            lazyWorkbookRef,
            worksheet,
            load,
            setMessage,
            load.endRow < lastLoadRow,
          )
          if (!ok) {
            throw new Error(
              'Part of the target range could not be loaded — retry after the workbook finishes indexing.',
            )
          }
        }
      }
      applyChunk(chunk)
    }
  } finally {
    // The chunked loads walked the streaming window strip by strip and left
    // it wherever the last chunk (or the failure) happened to be — the cells
    // the user is looking at were evicted along the way. Refetch the
    // viewport so the grid does not sit blank after a bulk edit. The window
    // record must survive: eviction keys off it, so dropping it would leave
    // the final strip resident and untracked. A mid-loop failure can leave
    // an unmeasured bulk window here — stale wrap heights until it scrolls
    // out, accepted over the leak.
    if (!state.flags.preloadComplete && lazyWorkbookRef.current === state) {
      const active = runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
      if (active?.getSheetId() === sheetId) {
        void loadVisibleRange(runtime, lazyWorkbookRef, active, setMessage)
      }
    }
  }
}

/// One text line at the sheet's default font — the smallest height an Excel
/// auto-fit can produce; stored heights below it are deliberate spacers.
/// Reads the file's own default (rounded to px like row heights are, so a
/// row at exactly the default compares equal); a workbook that omits it gets
/// Excel's factory 15pt, not Univer's taller UI default.
function defaultRowHeightPx(state: LazyWorkbookState, sheetId: string): number {
  const points = state.file.sheets.find((sheet) => sheet.id === sheetId)?.defaultRowHeight ?? 15
  return Math.round((points * 96) / 72)
}

/// Union of IStyleData keys the sheet's row/column default styles define.
/// Univer composes row/col styles into every cell per-property, but an OOXML
/// cell xf is complete — Excel never lets a row/column default show through a
/// cell that has its own xf — so styled cells override these keys explicitly.
export function sheetRowColStyleKeys(state: LazyWorkbookState, sheetId: string): Set<string> {
  let keys = state.rowColStyleKeys.get(sheetId)
  if (!keys) {
    keys = new Set()
    const sheet = state.file.sheets.find((candidate) => candidate.id === sheetId)
    for (const columnWidth of sheet?.columnWidths ?? []) {
      if (columnWidth.styleIndex === undefined) continue
      const style = state.file.styles[columnWidth.styleIndex]
      if (style) for (const key of Object.keys(toUniverStyle(style))) keys.add(key)
    }
    state.rowColStyleKeys.set(sheetId, keys)
  }
  return keys
}

/// Hidden-row knowledge for visible-order consumers: the set alone is not
/// enough — a row absent from it is only proven visible up to coveredThrough.
export interface HiddenRowsInfo {
  readonly rows: ReadonlySet<number>
  readonly coveredThrough: number
}

export function hiddenRowsInfo(
  state: LazyWorkbookState,
  sheetId: string,
): HiddenRowsInfo | undefined {
  const rows = state.hiddenFileRows.get(sheetId)
  if (!rows?.size) return undefined
  return { rows, coveredThrough: state.hiddenRowsCoveredThrough.get(sheetId) ?? -1 }
}

/// Must run before the range patch that covers these rows: filtered-table
/// stripe ranking needs the hidden set at banding time, but applyRowProperties
/// only lands after the patch. `range` is the row span this read covered —
/// coverage extends contiguously from row 0 (a batch with no row properties
/// still proves its rows visible; a gap-jumping read must not).
export function recordHiddenFileRows(
  state: LazyWorkbookState,
  sheetId: string,
  rows: WorkbookRangeResult['rows'],
  range: IRange,
): void {
  let hiddenSet = state.hiddenFileRows.get(sheetId)
  for (const row of rows) {
    if (!row.hidden) continue
    if (!hiddenSet) {
      hiddenSet = new Set()
      state.hiddenFileRows.set(sheetId, hiddenSet)
    }
    hiddenSet.add(row.row)
  }
  const covered = state.hiddenRowsCoveredThrough.get(sheetId) ?? -1
  if (range.startRow <= covered + 1 && range.endRow > covered) {
    state.hiddenRowsCoveredThrough.set(sheetId, range.endRow)
  }
}

/// Must run before the range patch that covers these rows: the patch bakes
/// the overrides into cell styles at apply time.
export function recordRowStyleKeys(
  state: LazyWorkbookState,
  sheetId: string,
  rows: WorkbookRangeResult['rows'],
): void {
  for (const row of rows) {
    if (row.styleIndex === undefined) continue
    const style = state.file.styles[row.styleIndex]
    if (!style) continue
    const keys = sheetRowColStyleKeys(state, sheetId)
    for (const key of Object.keys(toUniverStyle(style))) keys.add(key)
  }
}

/// Concrete "absent property" values per bleed key. Explicit nulls would be
/// the natural block, but Univer's set-range-values merge strips them
/// (Tools.removeNull), so the override must be a real value that renders
/// exactly like the missing property. pd/bd have no such value and their
/// (rare) bleed is left alone. Document data: hardcoded colors stay
/// theme-independent by design.
function bleedOverrideValue(key: string, normal: WorkbookCellStyle | undefined): unknown {
  switch (key) {
    case 'cl':
      return { rgb: '#000000' } // automatic font color
    case 'bg':
      return { rgb: '#FFFFFF' } // fillId 0 on the white page surface
    case 'ul':
    case 'st':
      return { s: BooleanNumber.FALSE }
    case 'ht':
      return HorizontalAlign.UNSPECIFIED
    case 'vt':
      return VerticalAlign.UNSPECIFIED
    case 'tr':
      return { a: 0 }
    case 'n':
      return { pattern: 'General' }
    case 'ff':
      return normal?.fontFamily === undefined ? 'Calibri' : escapeCssLeadingDigit(normal.fontFamily)
    case 'fs':
      return normal?.fontSize ?? 11
    default:
      return undefined
  }
}

/// Explicit stand-ins for every bleed key the cell xf leaves unset — an OOXML
/// cell xf is complete, so a row/column default must never show through it.
export function withRowColOverrides(
  style: IStyleData,
  bleedKeys: ReadonlySet<string> | undefined,
  normal: WorkbookCellStyle | undefined,
): IStyleData {
  if (bleedKeys?.size) {
    const record = style as Record<string, unknown>
    for (const key of bleedKeys) {
      if (record[key] !== undefined) continue
      const override = bleedOverrideValue(key, normal)
      if (override !== undefined) record[key] = override
    }
  }
  return style
}

export function applyRowProperties(
  worksheet: UniverWorksheet,
  state: LazyWorkbookState,
  sheetId: string,
  rows: WorkbookRangeResult['rows'],
): void {
  if (rows.length === 0) return
  recordRowStyleKeys(state, sheetId, rows)
  let applied = state.appliedRowKeys.get(sheetId)
  if (!applied) {
    applied = new Set()
    state.appliedRowKeys.set(sheetId, applied)
  }
  journalSuppression.active = true
  loadAutoHeightSuppression.active = true
  try {
    // Each of these commands runs the full Univer command pipeline. Exports
    // that stamp ht/customHeight on every row (common in ERP/BI files) would
    // otherwise cost one synchronous command per row — tens of thousands per
    // streamed chunk — so collect the rows and flush contiguous runs as
    // single ranged commands instead.
    const heights: { row: number; px: number }[] = []
    const autoRows: number[] = []
    const hiddenRows: number[] = []
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
      const key = `${row.row}:${row.height ?? ''}:${row.customHeight ?? false}:${row.hidden}:${row.styleIndex ?? ''}`
      if (applied.has(key)) continue
      applied.add(key)
      if (row.styleIndex !== undefined) {
        // <row s= customFormat>: the default style for cells in the row that
        // carry none of their own. Model-level write; the patch that follows
        // each chunk repaints the range.
        const style = state.file.styles[row.styleIndex]
        if (style) worksheet.getSheet().setRowStyle(row.row, toUniverStyle(style))
      }
      if (row.height !== undefined) {
        const px = Math.round((row.height * 96) / 72)
        // Paint the stored ht first; the open-time wrap measure
        // (wrapAutoFitRows) re-fits auto-mode wrap rows afterwards, matching
        // Excel's own open-time re-measure of cached heights.
        heights.push({ row: row.row, px })
        if (!row.customHeight && px >= defaultRowHeightPx(state, sheetId)) {
          // Without customHeight the row is still in Excel's auto mode: a
          // later USER edit in the row must re-fit it. setRowHeightsForced
          // locked ia=0; flip it back — the suppression flag raised above
          // keeps the command from measuring (and ballooning) the row now.
          // Sub-default heights stay locked: they are deliberate spacer rows
          // an auto-fit would balloon to a full text line.
          autoRows.push(row.row)
        }
      }
      if (row.hidden) hiddenRows.push(row.row)
    }
    if (hiddenRows.length > 0) {
      let hiddenSet = state.hiddenFileRows.get(sheetId)
      if (!hiddenSet) {
        hiddenSet = new Set()
        state.hiddenFileRows.set(sheetId, hiddenSet)
      }
      for (const row of hiddenRows) hiddenSet.add(row)
    }
    heights.sort((a, b) => a.row - b.row)
    let run: { start: number; count: number; px: number } | null = null
    for (const h of heights) {
      if (run && h.row === run.start + run.count && h.px === run.px) {
        run.count += 1
      } else {
        if (run) worksheet.setRowHeightsForced(run.start, run.count, run.px)
        run = { start: h.row, count: 1, px: h.px }
      }
    }
    if (run) worksheet.setRowHeightsForced(run.start, run.count, run.px)
    forEachRowRun(autoRows, (start, count) => worksheet.setRowAutoHeight(start, count))
    forEachRowRun(hiddenRows, (start, count) => worksheet.hideRows(start, count))
  } finally {
    journalSuppression.active = false
    loadAutoHeightSuppression.active = false
  }
}

function forEachRowRun(rows: number[], apply: (start: number, count: number) => void): void {
  rows.sort((a, b) => a - b)
  let start = -1
  let count = 0
  for (const row of rows) {
    if (count > 0 && row === start + count) {
      count += 1
    } else {
      if (count > 0) apply(start, count)
      start = row
      count = 1
    }
  }
  if (count > 0) apply(start, count)
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
  // .merge() re-selects each merged range (AddMergeRedoSelectionsOperation),
  // so a file load would leave a phantom selection highlight on whichever
  // merge happened to apply last. Remember the real selection and put it back.
  let activeBefore: string | null
  try {
    activeBefore = worksheet.getSelection()?.getActiveRange()?.getA1Notation() ?? null
  } catch {
    activeBefore = null
  }
  let appliedAny = false
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
        appliedAny = true
      } catch {
        // An overlapping merge from a previous partial pass is not fatal.
      }
    }
  } finally {
    journalSuppression.active = false
  }
  if (appliedAny && activeBefore) {
    try {
      worksheet.getRange(activeBefore).activate()
    } catch {
      // Selection restore is cosmetic; never fail the load over it.
    }
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

/// Remembers each formula cell's file-cached value before the engine gets a
/// chance to recalculate it, so the display fallback can restore Excel's
/// result when that recalculation errors. Only needed where formulas are
/// handed to the engine (useFormulas paths).
function recordCachedFormulaValues(
  state: LazyWorkbookState,
  sheetId: string,
  cells: WorkbookRangeResult['cells'],
): void {
  let cached = state.cachedFormulaValues.get(sheetId)
  for (const cell of cells) {
    if (!cell.formula || cell.value === null || cell.value === undefined) continue
    if (!cached) {
      cached = new Map()
      state.cachedFormulaValues.set(sheetId, cached)
    }
    cached.set(`${cell.row}:${cell.column}`, cell.value)
  }
}

/// Row/col property commands and SetRangeValuesCommand tail a selection op
/// onto the written sheet, and Univer's ActiveWorksheetController then
/// asynchronously activates whichever sheet the selection landed on.
/// Streaming file content into a background (even hidden) sheet must not
/// steal the active one. The activation runs after the command's promise
/// chain, so a synchronous restore alone loses the race — re-check across
/// the microtask and task queues too. Only a flip TO the patched sheet is
/// undone, so a genuine user sheet switch in the same window survives.
function keepActiveSheet<T>(worksheet: UniverWorksheet, run: () => T): T {
  const facade = worksheet as unknown as {
    getWorkbook?: () => {
      getActiveSheet(allowNull: true): { getSheetId(): string } | null
      setActiveSheet(sheet: unknown): void
    }
    _fWorkbook?: { setActiveSheet(sheetId: string): unknown }
  }
  const workbook = facade.getWorkbook?.()
  const before = workbook?.getActiveSheet(true)
  const patchedId = worksheet.getSheetId()
  const restore = (): void => {
    if (!workbook || !before || before.getSheetId() === patchedId) return
    const current = workbook.getActiveSheet(true)
    if (current && current !== before && current.getSheetId() === patchedId) {
      // Restore through the full SetWorksheetActiveOperation, not the bare
      // model setter: the stray activation also moved the render skeleton's
      // current sheet, and a model-only restore leaves canvas and model
      // pointing at different sheets — resolveRenderedSheetId then "heals"
      // the model back to the patched sheet, making the theft permanent.
      const fWorkbook = facade._fWorkbook
      if (fWorkbook) fWorkbook.setActiveSheet(before.getSheetId())
      else workbook.setActiveSheet(before)
    }
  }
  try {
    return run()
  } finally {
    restore()
    queueMicrotask(restore)
    setTimeout(restore, 0)
    setTimeout(restore, 60)
  }
}

function patchWorksheetRange(
  worksheet: UniverWorksheet,
  previousRange: IRange | undefined,
  range: IRange,
  cells: WorkbookRangeResult['cells'],
  styles: readonly WorkbookCellStyle[],
  hyperlinks: WorkbookRangeResult['hyperlinks'],
  tables: WorkbookFile['sheets'][number]['tables'],
  pivotTables: WorkbookFile['sheets'][number]['pivotTables'],
  freeze: WorkbookFile['sheets'][number]['freeze'],
  useFormulas = false,
  journal?: EditJournal,
  pinned?: ReadonlyMap<string, PinnedClosureCell>,
  recalcOverlay?: ReadonlyMap<string, PinnedClosureCell>,
  arrayFollowers?: ReadonlySet<string>,
  rowColStyleKeys?: ReadonlySet<string>,
  inheritedWrap?: (row: number, column: number) => boolean,
  hiddenRows?: HiddenRowsInfo,
): void {
  // Windowed CF formula registrations follow the loaded row window; report
  // it before the patch so rules installed right after see it. Best-effort:
  // without a sheet id (unit-test worksheet doubles) windowing simply stays
  // off and CF registrations keep their full folded ranges.
  const streamSheetId = (worksheet as { getSheet?: () => { getSheetId?: () => string } })
    .getSheet?.()
    ?.getSheetId?.()
  if (streamSheetId) notifyCfStreamWindow(streamSheetId, range.startRow, range.endRow)
  journalSuppression.active = true
  // Installing file content must keep every row at its stored height — Excel
  // does not re-measure on open — while leaving rows in auto mode for edits.
  loadAutoHeightSuppression.active = true
  try {
    keepActiveSheet(worksheet, () => {
      patchWorksheetRangeInner(
        worksheet,
        previousRange,
        range,
        cells,
        styles,
        tables,
        pivotTables,
        freeze,
        useFormulas,
        arrayFollowers,
        rowColStyleKeys,
        inheritedWrap,
        pinned?.size ? new Set(pinned.keys()) : undefined,
        hiddenRows,
      )
      // Closure cells are engine-owned: streaming skips them entirely (see
      // pinnedKeys in the inner patcher) instead of evict-and-re-pin — every
      // rewrite re-dirtied the whole dependency web, so each scroll chunk
      // re-ran thousands of formulas and painted mid-cascade values (an
      // incremental date chain recomputed from an emptied anchor shows year
      // 1900 — alpha ledger r141). The journal overlay still runs after so
      // user edits always win.
      if (recalcOverlay?.size) applyPinnedOverlay(worksheet, recalcOverlay, previousRange, range)
      if (journal) applyJournalOverlay(worksheet, journal, range)
    })
  } finally {
    journalSuppression.active = false
    loadAutoHeightSuppression.active = false
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
            ? cachedFormulaCellData(cell.f, cell.v)
            : typeof cell.v === 'string' && cell.v !== ''
              ? { v: cell.v, t: CellValueType.STRING }
              : { v: cell.v ?? null },
        ],
      ])
  }
}

export function applyJournalOverlay(
  worksheet: UniverWorksheet,
  journal: EditJournal,
  range: IRange,
): void {
  const sheetId = worksheet.getSheetId()
  const fills = journal.bulkConstantFills?.get(sheetId) ?? []
  for (const fill of fills) {
    const startRow = Math.max(range.startRow, fill.startRow)
    const endRow = Math.min(range.endRow, fill.endRow)
    const startColumn = Math.max(range.startColumn, fill.startColumn)
    const endColumn = Math.min(range.endColumn, fill.endColumn)
    if (startRow > endRow || startColumn > endColumn) continue
    const rows = endRow - startRow + 1
    const columns = endColumn - startColumn + 1
    const matrix = Array.from({ length: rows }, (_unused, rowOffset) =>
      Array.from({ length: columns }, (_unusedColumn, columnOffset) => {
        const result = bulkConstantFillValueAt(
          journal,
          sheetId,
          startRow + rowOffset,
          startColumn + columnOffset,
        )
        return result.found ? { v: result.value } : {}
      }),
    )
    worksheet.getRange(startRow, startColumn, rows, columns).setValues(matrix)
  }
  for (const entry of journalEntriesInRange(journal, sheetId, range)) {
    const cellRange = worksheet.getRange(entry.row, entry.column, 1, 1)
    if (entry.hasValue) {
      // Replayed rich/multiline docs need the same cell-font base as the
      // load path; the cell's composed style is already installed here.
      const baseFont = (): IStyleData =>
        fontTextStyleOf(worksheet.getSheet().getComposedCellStyle(entry.row, entry.column))
      if (entry.formula) cellRange.setValues([[{ f: entry.formula }]])
      else if (entry.value === null) cellRange.clearContent()
      else if (entry.rich && typeof entry.value === 'string') {
        cellRange.setValues([[{ p: toRichTextDocument(entry.value, [...entry.rich], baseFont()) }]])
      } else if (typeof entry.value === 'string' && entry.value.includes('\n')) {
        cellRange.setValues([[{ p: toRichTextDocument(entry.value, [], baseFont()) }]])
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

/// Formulas that use defined names (or external refs) recalculate as #NAME?
/// or blank while the engine races the name installation on open — Excel
/// shows the cached value instantly. Keep such cells cache-only (the formula
/// text still reaches the formula bar via formulaText).
const EXCEL_ERROR_LITERALS = new Set([
  '#NULL!',
  '#DIV/0!',
  '#VALUE!',
  '#REF!',
  '#NAME?',
  '#NUM!',
  '#N/A',
  '#SPILL!',
  '#CALC!',
])

/// A cached formula value installed as bare `{ f, v }` lets Univer coerce
/// numeric-looking text ("1") into a number, which then picks up the cell's
/// numeric format — a [h]:mm cell rendered 24:00 where Excel shows 1.
export function cachedFormulaCellData(
  formula: string,
  value: string | number | boolean | null | undefined,
): { f: string; v?: string | number | boolean; t?: CellValueType } {
  if (value === null || value === undefined) return { f: formula }
  return typeof value === 'string'
    ? { f: formula, v: value, t: CellValueType.STRING }
    : { f: formula, v: value }
}

/// Functions whose result is minted in the author's locale — NUMBERSTRING's
/// numeral script (ko/ja files cache hangul/kanji numerals, the engine emits
/// zh), DOLLAR's currency symbol (₩/¥ vs the engine's $). Re-evaluation is
/// always wrong for a file authored elsewhere, so the cached value stays.
const LOCALE_DEPENDENT_FUNCTION = /(?:^|[^A-Z0-9_.])(?:NUMBERSTRING|DOLLAR)\s*\(/
export function usesLocaleDependentFunction(formula: string): boolean {
  const segments = formula.split('"')
  for (let index = 0; index < segments.length; index += 2) {
    if (LOCALE_DEPENDENT_FUNCTION.test((segments[index] ?? '').toUpperCase())) return true
  }
  return false
}

interface KeepsCacheVerdict {
  /// Keep the cache whether or not the file has one: the engine's result is
  /// wrong by construction (locale, Google-only, unresolvable names).
  readonly always: boolean
  /// Keep the cache only when the file has one — with nothing to keep, the
  /// engine's attempt (#NAME?, or a wrapped fallback) beats a blank cell.
  readonly unsupportedFunction: boolean
}

const keepsCacheMemo = new Map<string, KeepsCacheVerdict>()

/// Window onto the running formula engine's function registry. Installed
/// once the Univer instance exists; without one every function counts as
/// supported, so the decision below degrades to its pre-probe behaviour.
export interface SupportedFunctionProbe {
  /// False until the engine has registered its builtins — the registry says
  /// nothing about support before that, so no verdict is memoised.
  readonly ready: () => boolean
  /// Whether the engine implements a function, by canonical name
  /// (uppercase, storage markers stripped).
  readonly supports: (name: string) => boolean
}

let supportedFunctionProbe: SupportedFunctionProbe | null = null

export function setSupportedFunctionProbe(probe: SupportedFunctionProbe | null): void {
  supportedFunctionProbe = probe
  // Verdicts memoised under the previous registry are stale.
  keepsCacheMemo.clear()
}

/// True when the formula calls a function the engine lacks. Recalculating
/// such a formula yields #NAME? — or, when the author wrapped the call in
/// IFERROR/ISERROR/IFNA/IFS/CHOOSE, the fallback literal (0, "") in place of
/// the true value. The file's cached <v> is Excel's own result: keep it.
function callsUnsupportedFunction(formula: string, probe: SupportedFunctionProbe): boolean {
  return extractFunctionNames(formula).some((name) => !probe.supports(name))
}

export function formulaKeepsCache(formula: string, hasCachedValue = true): boolean {
  let verdict = keepsCacheMemo.get(formula)
  if (verdict === undefined) {
    if (keepsCacheMemo.size > 20_000) keepsCacheMemo.clear()
    // Google Sheets exports unevaluable functions as
    // IFERROR(__xludf.DUMMYFUNCTION("..."), <literal>); recalculating turns
    // the float-repr literal (46235.0) into a string that numfmt skips.
    // The cached <v> IS the computed value — keep it, like Excel does.
    const always =
      formula.includes('__xludf.') ||
      usesLocaleDependentFunction(formula) ||
      containsUnresolvedNames(formula)
    const probe = supportedFunctionProbe
    const registryReady = probe === null || probe.ready()
    verdict = {
      always,
      unsupportedFunction:
        !always && probe !== null && registryReady && callsUnsupportedFunction(formula, probe),
    }
    // A verdict taken before the builtins landed would pin "supported" on
    // every function for the rest of the session.
    if (registryReady) keepsCacheMemo.set(formula, verdict)
  }
  return verdict.always || (hasCachedValue && verdict.unsupportedFunction)
}

/// A cached value worth keeping: present and not an error literal. An
/// error cache means the last writer could not evaluate the formula either
/// (LibreOffice caches #NAME? for functions it lacks) — nothing to preserve.
export function hasUsableCachedValue(value: string | number | boolean | null | undefined): boolean {
  if (value === null || value === undefined) return false
  return !(typeof value === 'string' && EXCEL_ERROR_LITERALS.has(value))
}

/// Excel never wraps non-text values: a too-wide number/date renders as a
/// one-line hash fill (####), so a wrap style on a numeric cell must not
/// change the line count or the row height. CLIP approximates that (the hash
/// fill itself is a separate numeric-overflow render feature). Formula cells
/// qualify on their cached numeric value: the engine could in principle
/// recompute one into text (which would then clip a line), but a wrapped
/// two-line number is the common, always-wrong case (prod ¥ amounts).
export function numericWrapOverride(value: unknown, wrapText: boolean | undefined): boolean {
  return typeof value === 'number' && wrapText === true
}

/// Excel renders a manual line break (Alt+Enter) as nothing when the cell
/// does not wrap: the lines join on one display line.
export function joinManualBreaks(value: string): string {
  return value.replace(/\r\n?|\n/g, '')
}

/// Effective wrapText for a cell without its own xf. A customFormat row xf is
/// the COMPLETE default xf for its unstyled cells — the column xf is not
/// consulted (same OOXML semantics the sidecar's row styleIndex guard
/// encodes) — otherwise the column xf applies.
export function inheritedWrapLookup(
  styles: readonly WorkbookCellStyle[],
  rows: WorkbookRangeResult['rows'],
  columnWidths: WorkbookFile['sheets'][number]['columnWidths'],
): (row: number, column: number) => boolean {
  const rowWrap = new Map<number, boolean>()
  for (const row of rows) {
    if (row.styleIndex === undefined) continue
    rowWrap.set(row.row, styles[row.styleIndex]?.wrapText === true)
  }
  const columnSpans = columnWidths
    .filter((span) => span.styleIndex !== undefined)
    .map((span) => ({
      start: span.startColumn,
      end: span.endColumn,
      wrap: styles[span.styleIndex as number]?.wrapText === true,
    }))
  return (row, column) => {
    const fromRow = rowWrap.get(row)
    if (fromRow !== undefined) return fromRow
    // Match createColumnData: overlapping <col> spans apply in file order, so
    // the last styled span covering the column decides.
    let wrap = false
    for (const span of columnSpans) {
      if (column >= span.start && column <= span.end) wrap = span.wrap
    }
    return wrap
  }
}

/// Rows Excel DOES auto-fit when opening a file: the row is in auto mode
/// (no customHeight — a cached ht alone does not opt out: Excel live-probes
/// re-fit ht="30" rows to 16pt on open, prod_054/prod_027), the sheet
/// default is not user-fixed (sheetFormatPr customHeight), the cached ht is
/// not a sub-default spacer, and at least one loaded cell wraps real text.
/// User-fixed heights keep their stored value verbatim (#884); numeric wrap
/// cells and no-wrap manual breaks never change the line count, so neither
/// counts.
export function wrapAutoFitRows(
  cells: WorkbookRangeResult['cells'],
  styles: readonly WorkbookCellStyle[],
  rows: WorkbookRangeResult['rows'],
  columnWidths: WorkbookFile['sheets'][number]['columnWidths'],
  defaultRowHeightFixed: boolean | undefined,
  defaultRowHeight: number | null | undefined,
  range: IRange,
  merges?: WorkbookRangeResult['merges'],
): number[] {
  if (defaultRowHeightFixed) return []
  const inheritedWrap = inheritedWrapLookup(styles, rows, columnWidths)
  // Excel never grows a row for a merged wrap cell, and measuring one at its
  // anchor column's width invents multi-line heights (a label anchored in a
  // hair-width column measures one character per line). Merge-covered cells
  // must not qualify their row.
  const mergedIntervals = new Map<number, Array<readonly [number, number]>>()
  for (const merge of merges ?? []) {
    for (let row = merge.startRow; row <= merge.endRow; row += 1) {
      if (row < range.startRow || row > range.endRow) continue
      let intervals = mergedIntervals.get(row)
      if (!intervals) {
        intervals = []
        mergedIntervals.set(row, intervals)
      }
      intervals.push([merge.startColumn, merge.endColumn])
    }
  }
  const inMerge = (row: number, column: number): boolean => {
    const intervals = mergedIntervals.get(row)
    if (!intervals) return false
    return intervals.some(([start, end]) => column >= start && column <= end)
  }
  const lockedHeights = new Set<number>()
  const defaultPt = defaultRowHeight ?? 15
  for (const row of rows) {
    if (row.customHeight || (row.height !== undefined && row.height < defaultPt)) {
      lockedHeights.add(row.row)
    }
  }
  const qualifying = new Set<number>()
  for (const cell of cells) {
    if (cell.row < range.startRow || cell.row > range.endRow) continue
    if (cell.column < range.startColumn || cell.column > range.endColumn) continue
    if (qualifying.has(cell.row) || lockedHeights.has(cell.row)) continue
    if (inMerge(cell.row, cell.column)) continue
    const style = cell.styleIndex === undefined ? undefined : styles[cell.styleIndex]
    const wraps = style ? style.wrapText === true : inheritedWrap(cell.row, cell.column)
    if (!wraps) continue
    const value = cell.value ?? ''
    if (typeof value !== 'string' || value === '') continue
    qualifying.add(cell.row)
  }
  return [...qualifying].sort((a, b) => a - b)
}

/// Univer's AutoHeightController registers the auto-height interceptor only
/// when the lifecycle reaches Rendered; a measure dispatched before that
/// silently yields nothing. Loads race that stage on open, so early measures
/// queue here and flush when the stage arrives.
export const wrapMeasureGate: {
  ready: boolean
  pending: Array<{ worksheet: UniverWorksheet; rows: readonly number[] }>
} = { ready: false, pending: [] }

export function installWrapMeasureLifecycle(runtime: UniverRuntime): { dispose(): void } {
  wrapMeasureGate.ready = false
  wrapMeasureGate.pending.length = 0
  preIndexMeasuredRows.clear()
  return runtime.univerAPI.addEvent(runtime.univerAPI.Event.LifeCycleChanged, (params) => {
    const { stage } = params as { stage: LifecycleStages }
    if (stage < LifecycleStages.Rendered || wrapMeasureGate.ready) return
    wrapMeasureGate.ready = true
    // Let the other lifecycle subscribers (the plugin hooks that create
    // AutoHeightController) run before the queued measures.
    setTimeout(() => {
      for (const item of wrapMeasureGate.pending.splice(0)) {
        measureWrapAutoFitRows(item.worksheet, item.rows)
      }
    }, 0)
  })
}

/// Rows measured while the sidecar was still indexing (merges incomplete),
/// keyed `unitId:sheetId`. Only these are reset candidates once indexing
/// completes — clearing any other non-qualifying row would wipe legitimate
/// heights (a user autofit, or wrap cells outside the window's columns).
const preIndexMeasuredRows = new Map<string, Set<number>>()

export function trackPreIndexMeasuredRows(key: string, rows: readonly number[]): void {
  if (rows.length === 0) return
  let tracked = preIndexMeasuredRows.get(key)
  if (!tracked) {
    tracked = new Set()
    preIndexMeasuredRows.set(key, tracked)
  }
  for (const row of rows) tracked.add(row)
}

/// Consumes the tracked rows visible in this window: rows that still qualify
/// were measured correctly and drop out; rows that qualified only without
/// merge knowledge come back as contamination to reset.
export function takeContaminatedRows(
  key: string,
  qualifyingWithoutMerges: readonly number[],
  qualifying: readonly number[],
): number[] {
  const tracked = preIndexMeasuredRows.get(key)
  if (!tracked || tracked.size === 0) return []
  const stillQualified = new Set(qualifying)
  const contaminated: number[] = []
  for (const row of qualifyingWithoutMerges) {
    if (!tracked.delete(row)) continue
    if (!stillQualified.has(row)) contaminated.push(row)
  }
  return contaminated
}

/// A wrap measure can run before the sheet's merges are known — the sidecar
/// publishes merges only when the whole part finishes parsing, while windowed
/// reads return as soon as their rows are indexed. Rows contaminated by such
/// a measure keep their absurd auto height forever: once the merges install,
/// the re-measure skips merged cells and yields nothing, so nothing corrects
/// the stored value. Once indexing is complete (merges final), clear the auto
/// height of every row in the window that no longer qualifies.
export function resetStaleWrapAutoHeights(
  runtime: UniverRuntime,
  unitId: string,
  worksheet: UniverWorksheet,
  candidates: readonly number[],
  rows: WorkbookRangeResult['rows'],
  defaultRowHeight: number | null | undefined,
): void {
  if (candidates.length === 0) return
  // Mirrors the lockedHeights rule in wrapAutoFitRows: a cached ht without
  // customHeight is still auto mode — such a row can carry a poisoned
  // measure too, and resets back to its stored file height. customHeight and
  // sub-default spacer heights stay verbatim.
  const defaultPt = defaultRowHeight ?? 15
  const locked = new Set<number>()
  for (const row of rows) {
    if (row.customHeight || (row.height !== undefined && row.height < defaultPt)) {
      locked.add(row.row)
    }
  }
  const resetRows = candidates.filter((row) => !locked.has(row))
  if (resetRows.length === 0) return
  // Purge these rows from measures still queued behind the Rendered gate:
  // the lifecycle flush would re-run the merge-less list and re-poison them.
  const sheetId = worksheet.getSheetId()
  const resetSet = new Set(resetRows)
  for (const [index, item] of wrapMeasureGate.pending.entries()) {
    if (item.worksheet.getSheetId() !== sheetId) continue
    wrapMeasureGate.pending[index] = {
      worksheet: item.worksheet,
      rows: item.rows.filter((row) => !resetSet.has(row)),
    }
  }
  const snapshot = worksheet.getSheet().getSnapshot()
  const rowData = snapshot.rowData as
    Record<number, { ah?: number; h?: number } | undefined> | undefined
  if (!rowData) return
  const stale: Array<{ row: number; autoHeight: undefined }> = []
  for (const row of resetRows) {
    if (rowData[row]?.ah === undefined) continue
    stale.push({ row, autoHeight: undefined })
  }
  if (stale.length === 0) return
  journalSuppression.active = true
  try {
    runtime.univerAPI.syncExecuteCommand('sheet.mutation.set-worksheet-row-auto-height', {
      unitId,
      subUnitId: worksheet.getSheetId(),
      rowsAutoHeightInfo: stale,
    })
  } finally {
    journalSuppression.active = false
  }
}

/// Runs the user-autofit measurement channel over the qualifying rows. The
/// load gate must be DOWN (this is the one load-time measure Excel really
/// does); the undo/journal suppression stays up so opening a file neither
/// pollutes undo nor dirties the document. Univer's measure never shrinks a
/// row below the sheet default, so single-line wrap rows are untouched.
export function measureWrapAutoFitRows(
  worksheet: UniverWorksheet,
  rowsToMeasure: readonly number[],
): void {
  if (rowsToMeasure.length === 0) return
  if (!wrapMeasureGate.ready) {
    wrapMeasureGate.pending.push({ worksheet, rows: rowsToMeasure })
    return
  }
  journalSuppression.active = true
  try {
    let start = rowsToMeasure[0] as number
    let previous = start
    for (const row of rowsToMeasure.slice(1)) {
      if (row === previous + 1) {
        previous = row
        continue
      }
      worksheet.setRowAutoHeight(start, previous - start + 1)
      start = row
      previous = row
    }
    worksheet.setRowAutoHeight(start, previous - start + 1)
  } finally {
    journalSuppression.active = false
  }
}

export function patchWorksheetRangeInner(
  worksheet: UniverWorksheet,
  previousRange: IRange | undefined,
  range: IRange,
  cells: WorkbookRangeResult['cells'],
  styles: readonly WorkbookCellStyle[],
  tables: WorkbookFile['sheets'][number]['tables'],
  pivotTables: WorkbookFile['sheets'][number]['pivotTables'],
  freeze: WorkbookFile['sheets'][number]['freeze'],
  useFormulas: boolean,
  arrayFollowers?: ReadonlySet<string>,
  rowColStyleKeys?: ReadonlySet<string>,
  inheritedWrap?: (row: number, column: number) => boolean,
  pinnedKeys?: ReadonlySet<string>,
  hiddenRows?: HiddenRowsInfo,
): void {
  if (previousRange) {
    // Frozen rows/columns stay visible while scrolling, so never evict them —
    // later viewport patches don't include them and they'd go blank.
    const clearStartRow = Math.max(previousRange.startRow, freeze?.frozenRows ?? 0)
    const clearStartColumn = Math.max(previousRange.startColumn, freeze?.frozenColumns ?? 0)
    if (clearStartRow <= previousRange.endRow && clearStartColumn <= previousRange.endColumn) {
      // Evict with a raw null-cell write, never clearContent()/clearFormat():
      // those facade calls run ClearSelectionContent/Format COMMANDS, whose
      // plugin interceptors apply user-level semantics to the CURRENT
      // SELECTION — conditional formatting strips the selected cells from
      // every rule, a rule-range rectangle decomposition that costs seconds
      // per eviction on wide rules and permanently fragments the rule
      // (genoffice#158). Eviction is internal bookkeeping; only cells move.
      // Engine-owned closure cells must survive the eviction ({} is a merge
      // no-op, so pinned cells stay untouched — clearing and re-installing a
      // formula re-dirties its whole dependency web, alpha ledger r141).
      const clearRows = previousRange.endRow - clearStartRow + 1
      const clearColumns = previousRange.endColumn - clearStartColumn + 1
      const wipe: ICellData[][] = Array.from({ length: clearRows }, (_unused, rowOffset) =>
        Array.from({ length: clearColumns }, (_unusedColumn, columnOffset) =>
          pinnedKeys?.has(`${clearStartRow + rowOffset}:${clearStartColumn + columnOffset}`)
            ? {}
            : { v: null, f: null, si: null, p: null, s: null, t: null, custom: null },
        ),
      )
      worksheet.getRange(clearStartRow, clearStartColumn, clearRows, clearColumns).setValues(wipe)
    }
  }
  const rows = range.endRow - range.startRow + 1
  const columns = range.endColumn - range.startColumn + 1
  const matrix: ICellData[][] = Array.from({ length: rows }, () =>
    Array.from({ length: columns }, () => ({})),
  )
  const overrideCells: Array<[number, number]> = []
  // row → column → "has content", for centerContinuous cells only.
  const centerAcross = new Map<number, Map<number, boolean>>()
  for (const cell of cells) {
    if (
      cell.row < range.startRow ||
      cell.row > range.endRow ||
      cell.column < range.startColumn ||
      cell.column > range.endColumn
    ) {
      continue
    }
    // Engine-owned closure cell: leave the {} merge no-op in the matrix —
    // rewriting the formula would re-dirty its dependency web (r141).
    if (pinnedKeys?.has(`${cell.row}:${cell.column}`)) continue
    const keepsCache =
      useFormulas && cell.formula
        ? formulaKeepsCache(cell.formula, hasUsableCachedValue(cell.value))
        : false
    // A formula cell with no cached value must show blank — falling back to
    // the formula string would print it as literal text (r141 reopen: cache
    // mode painted "=A750+1" walls). The formula-string fallback survives
    // only for engine-installed cells, where displayValue is never painted.
    const displayValue =
      cell.value ?? (cell.formula ? (useFormulas && !keepsCache ? cell.formula : '') : '') ?? ''
    const row = matrix[cell.row - range.startRow]
    const style = cell.styleIndex === undefined ? undefined : styles[cell.styleIndex]
    if (style?.horizontalAlignment === 'centerContinuous') {
      let rowRuns = centerAcross.get(cell.row)
      if (!rowRuns) centerAcross.set(cell.row, (rowRuns = new Map()))
      rowRuns.set(cell.column, (cell.value != null && cell.value !== '') || Boolean(cell.formula))
    }
    // CSE array follower: its dead cached value would block the master's
    // spill with #SPILL!; keep the style, let the engine fill the content.
    if (useFormulas && arrayFollowers?.has(`${cell.row}:${cell.column}`)) {
      if (row) {
        row[cell.column - range.startColumn] = style ? { s: toUniverStyle(style) } : {}
        if (style && rowColStyleKeys?.size) {
          overrideCells.push([cell.row - range.startRow, cell.column - range.startColumn])
        }
      }
      continue
    }
    const multiline = typeof displayValue === 'string' && displayValue.includes('\n')
    // Excel honors manual line breaks only when the cell wraps; with wrap off
    // it renders the lines joined on a single line at the stored row height.
    // A cell xf is complete, so it alone decides; only xf-less cells inherit
    // wrap from the row/column default.
    const wrapsEffective = style
      ? style.wrapText === true
      : (inheritedWrap?.(cell.row, cell.column) ?? false)
    const joinLines = multiline && !wrapsEffective
    // Wrap inherited from a row/column default counts too: Excel never
    // wraps numbers regardless of where the wrap flag comes from.
    const numericNoWrap = numericWrapOverride(displayValue, wrapsEffective)
    // shrinkToFit text cells: bake the reduced size into the cell style so
    // both the plain and rich render paths (and the #### rule) see it.
    // Formula cells are skipped in formula mode — their display text isn't
    // known until the engine evaluates.
    const shrinkScale =
      style?.shrinkToFit &&
      !style.wrapText &&
      typeof displayValue === 'string' &&
      displayValue !== '' &&
      !(useFormulas && cell.formula)
        ? shrinkScaleFor(
            worksheet,
            cell.row,
            cell.column,
            // Joined manual breaks display as one line, so fit that line —
            // the widest fragment alone would under-shrink.
            joinLines ? joinManualBreaks(displayValue) : displayValue,
            style,
            cell.rich,
          )
        : 1
    const effectiveStyle =
      shrinkScale >= 1 || !style
        ? style
        : { ...style, fontSize: scaledFontSize(style.fontSize ?? 11, shrinkScale) }
    if (row) {
      row[cell.column - range.startColumn] = {
        // Explicit string typing: bare `v` lets Univer coerce numeric-looking
        // text ("007", phone numbers) into numbers.
        ...(cell.rich && typeof displayValue === 'string'
          ? {
              p: toRichTextDocument(
                joinLines ? joinManualBreaks(displayValue) : displayValue,
                joinLines
                  ? cell.rich.map((run) => ({ ...run, text: joinManualBreaks(run.text) }))
                  : cell.rich,
                cellFontTextStyle(effectiveStyle),
                shrinkScale,
              ),
              // Excel joins manual breaks only visually — CHAR(10) stays in
              // the stored value; keep v raw so formulas and copies see it.
              ...(joinLines ? { v: displayValue, t: CellValueType.STRING } : {}),
            }
          : useFormulas && cell.formula && !keepsCache
            ? // No cached value: leave v unset so the engine computes instead
              // of showing the formula text as a literal. An error-literal
              // cache means the last writer could not evaluate the formula
              // (LibreOffice caches #NAME? for functions it lacks) — drop it
              // too, so the engine recomputes like Excel does on open.
              // Plain-arithmetic formulas also recompute: the engine handles
              // them at full parity, and Excel's own recalc-on-open shows
              // #VALUE! where a stale cache still holds a number (comma-
              // decimal text summed before the cells went text). CSE array
              // masters keep the cache — the engine would spill them.
              cell.value === null ||
              cell.value === undefined ||
              (typeof cell.value === 'string' && EXCEL_ERROR_LITERALS.has(cell.value)) ||
              (!cell.arrayRef && isPlainArithmeticFormula(cell.formula))
              ? { f: cell.formula }
              : cachedFormulaCellData(cell.formula, cell.value)
            : typeof displayValue === 'string' && joinLines
              ? // Excel joins manual breaks only visually — CHAR(10) stays in
                // the stored value. Render the joined line through the doc
                // model and keep v raw for formulas, edits, and copies.
                {
                  v: displayValue,
                  t: CellValueType.STRING,
                  p: toRichTextDocument(
                    joinManualBreaks(displayValue),
                    [],
                    cellFontTextStyle(effectiveStyle),
                  ),
                }
              : typeof displayValue === 'string' && multiline
                ? // Bare `v` renders only the first line; the doc model keeps all.
                  { p: toRichTextDocument(displayValue, [], cellFontTextStyle(effectiveStyle)) }
                : typeof displayValue === 'string' && displayValue !== ''
                  ? { v: displayValue, t: CellValueType.STRING }
                  : displayValue === ''
                    ? cell.value === ''
                      ? // A real empty-string value (shared string / inlineStr
                        // ''): keep it a STRING so copies and saves carry it —
                        // Excel keeps such cells non-blank (=ref returns '',
                        // =ref+1 is #VALUE!, COUNTA counts it).
                        { v: '', t: CellValueType.STRING }
                      : // Style-only cells (`<c r="B3" s="170"/>`, value null):
                        // `v: ''` would make them empty STRINGS, and a formula
                        // referencing one then returns '' where Excel returns 0
                        // (a serial-0 date shows 1900/1/0, a plain 0 shows 0 —
                        // not blank). `v: null` clears any earlier value and
                        // keeps them value-less.
                        { v: null }
                    : { v: displayValue }),
        ...(effectiveStyle || numericNoWrap
          ? {
              s: {
                // Hyperlinked cells keep the file's own font: Excel styles a
                // link via the cell xf, so injecting blue/underline here
                // overrode plain-styled links.
                ...(effectiveStyle ? toUniverStyle(effectiveStyle) : {}),
                ...(numericNoWrap ? { tb: WrapStrategy.CLIP } : {}),
              },
            }
          : {}),
      }
      // Only cells with their own xf override row/col defaults.
      if (effectiveStyle && rowColStyleKeys?.size) {
        overrideCells.push([cell.row - range.startRow, cell.column - range.startColumn])
      }
    }
  }
  // Excel centers a centerContinuous anchor's text across the run of trailing
  // BLANK cells sharing the format (the run ends at the next content cell or
  // differently-aligned cell) without merging; mark the run end so the render
  // patch (center-continuous.ts) can widen the centering/clip box.
  for (const [rowIndex, rowRuns] of centerAcross) {
    for (const [columnIndex, hasContent] of rowRuns) {
      if (!hasContent) continue
      let end = columnIndex
      while (rowRuns.get(end + 1) === false) end += 1
      if (end === columnIndex) continue
      const anchor = matrix[rowIndex - range.startRow]?.[columnIndex - range.startColumn]
      if (anchor) anchor.custom = { ...anchor.custom, [CENTER_ACROSS_END_KEY]: end }
    }
  }
  applyTableBanding(matrix, range, tables, hiddenRows)
  applyPivotStyling(matrix, range, pivotTables)
  // Bake the row/col-default overrides only after banding: the table/pivot
  // stripes treat any present bg/cl as the cell's own and must see the raw
  // xf, and a stripe fill they add counts as explicit (never overridden).
  for (const [rowIndex, columnIndex] of overrideCells) {
    const s = matrix[rowIndex]?.[columnIndex]?.s
    if (s && typeof s === 'object') {
      withRowColOverrides(s as IStyleData, rowColStyleKeys, styles[0])
    }
  }
  // No blanket auto-height sweep here: rows with a stored ht (or a fixed
  // sheet default) render verbatim like Excel; the caller raises
  // loadAutoHeightSuppression so the SetRangeValues interceptor stays quiet.
  // The one measure Excel really performs on open — auto rows with wrapped
  // text — runs afterwards via measureWrapAutoFitRows.
  worksheet.getRange(range.startRow, range.startColumn, rows, columns).setValues(matrix)
}

/// True only for a real fill: unfilled xfs carry the bg empty-rgb sentinel
/// (see toUniverStyle), which must not read as "baked fill".
function styleHasFill(style: IStyleData): boolean {
  return Boolean((style.bg as { rgb?: string } | null | undefined)?.rgb)
}

/// Approximates Excel table styles (header band + row stripes) for cells that
/// carry no explicit fill of their own.
export function applyTableBanding(
  matrix: ICellData[][],
  range: IRange,
  tables: WorkbookFile['sheets'][number]['tables'],
  hiddenRows?: HiddenRowsInfo,
): void {
  for (const table of tables) {
    const rowStart = Math.max(range.startRow, table.range.startRow)
    const rowEnd = Math.min(range.endRow, table.range.endRow)
    const columnStart = Math.max(range.startColumn, table.range.startColumn)
    const columnEnd = Math.min(range.endColumn, table.range.endColumn)
    if (rowStart > rowEnd || columnStart > columnEnd) continue
    // A name-less tableStyleInfo is Excel's style "None": paint nothing.
    if (!table.styleName && !table.headerFill && !table.headerFontColor && !table.stripeFill) {
      continue
    }
    // Colors are resolved sidecar-side from the workbook's real theme accents
    // (Light/Medium/Dark variant rules) or the file's custom <tableStyle>
    // dxfs; the literals are a last-resort fallback.
    const headerFill = table.headerFill
    const headerFont = table.headerFontColor ?? '#FFFFFF'
    // No stripe color means the style genuinely has none (custom styles
    // without band dxfs, Light 8-14) — a fallback would invent banding.
    const stripeFill = table.stripeFill
    const dataStartRow = table.range.startRow + table.headerRowCount
    const totalsStartRow = table.range.endRow - (table.totalsRowCount ?? 0) + 1
    // A live autoFilter makes Excel re-rank the stripes by VISIBLE row order
    // (a filtered Light19 ref alternates across the hidden gaps); manually
    // hidden rows keep the physical banding, so only filtered tables re-rank.
    // The scan needs full hidden knowledge from the table's top: past the
    // streamed coverage an absent row is merely unknown, not visible.
    let visibleParity: Map<number, number> | undefined
    const scanEnd = Math.min(rowEnd, totalsStartRow - 1)
    if (
      table.filterActive &&
      table.showRowStripes &&
      hiddenRows &&
      hiddenRows.rows.size > 0 &&
      scanEnd <= hiddenRows.coveredThrough
    ) {
      visibleParity = new Map()
      let ordinal = 0
      for (let row = dataStartRow; row <= scanEnd; row += 1) {
        if (hiddenRows.rows.has(row)) continue
        if (row >= rowStart) visibleParity.set(row, ordinal % 2)
        ordinal += 1
      }
    }
    for (let row = rowStart; row <= rowEnd; row += 1) {
      const isHeader = row < dataStartRow
      const isTotals = !isHeader && row >= totalsStartRow
      // Excel's firstRowStripe covers the FIRST data row (ref: Medium9 shades
      // data row 1 with #B8CCE4), then alternates with secondRowStripe.
      const rowParity = visibleParity?.get(row) ?? (row - dataStartRow) % 2
      const isStripe = !isHeader && !isTotals && table.showRowStripes && rowParity === 0
      const secondStripeFill =
        !isHeader && !isTotals && table.showRowStripes && rowParity === 1
          ? table.secondRowStripeFill
          : undefined
      for (let column = columnStart; column <= columnEnd; column += 1) {
        const cell = matrix[row - range.startRow]?.[column - range.startColumn]
        if (!cell) continue
        let style = (cell.s ?? {}) as IStyleData
        const hasCustomBorders =
          table.wholeTableBorderColor !== undefined ||
          table.innerHorizontalBorderColor !== undefined ||
          table.innerVerticalBorderColor !== undefined ||
          table.headerBottomBorderColor !== undefined
        if (table.borderColor || hasCustomBorders || (isTotals && table.totalRowBorderColor)) {
          const edges: IStyleData['bd'] = {}
          if (table.borderColor) {
            if (row === table.range.startRow) {
              edges.t = { s: BorderStyleTypes.MEDIUM, cl: { rgb: table.borderColor } }
            }
            if (isHeader && row === dataStartRow - 1) {
              edges.b = { s: BorderStyleTypes.THIN, cl: { rgb: table.borderColor } }
            }
            if (row === table.range.endRow) {
              edges.b = { s: BorderStyleTypes.MEDIUM, cl: { rgb: table.borderColor } }
            }
          }
          // Custom wholeTable dxf borders: inner grid first, the header rule
          // over it, the outline last — later edges win shared boundaries.
          if (table.innerHorizontalBorderColor && row < table.range.endRow) {
            edges.b = {
              s: mapBorderStyle(table.innerHorizontalBorderStyle ?? 'thin'),
              cl: { rgb: table.innerHorizontalBorderColor },
            }
          }
          if (table.innerVerticalBorderColor && column < table.range.endColumn) {
            edges.r = {
              s: mapBorderStyle(table.innerVerticalBorderStyle ?? 'thin'),
              cl: { rgb: table.innerVerticalBorderColor },
            }
          }
          if (table.headerBottomBorderColor && isHeader && row === dataStartRow - 1) {
            edges.b = {
              s: mapBorderStyle(table.headerBottomBorderStyle ?? 'thin'),
              cl: { rgb: table.headerBottomBorderColor },
            }
          }
          if (table.wholeTableBorderColor) {
            const outline = {
              s: mapBorderStyle(table.wholeTableBorderStyle ?? 'thin'),
              cl: { rgb: table.wholeTableBorderColor },
            }
            if (row === table.range.startRow) edges.t = outline
            if (row === table.range.endRow) edges.b = outline
            if (column === table.range.startColumn) edges.l = outline
            if (column === table.range.endColumn) edges.r = outline
          }
          if (isTotals && row === totalsStartRow && table.totalRowBorderColor) {
            edges.t = {
              s: mapBorderStyle(table.totalRowBorderStyle ?? 'thin'),
              cl: { rgb: table.totalRowBorderColor },
            }
          }
          if (edges.t || edges.b || edges.l || edges.r) {
            cell.s = { ...style, bd: { ...(style.bd ?? {}), ...edges } }
            style = cell.s as IStyleData
          }
        }
        if (isHeader) {
          const fontColor =
            column === table.range.startColumn && table.firstHeaderCellFontColor
              ? table.firstHeaderCellFontColor
              : headerFill
                ? headerFont
                : (table.headerFontColor ?? '#333333')
          if (styleHasFill(style)) {
            // Baked header fill: keep it, but a default-black font still takes
            // the style's header font (Excel lets table-style text win over
            // the automatic color).
            const cellFont = (style.cl as { rgb?: string } | undefined)?.rgb
            if (headerFill && (!cellFont || cellFont === '#000000')) {
              cell.s = { ...style, cl: { rgb: fontColor }, bl: BooleanNumber.TRUE }
            }
            continue
          }
          // An explicit non-automatic cell font color survives the table
          // style (Book1_custom's red "Names" header).
          const explicitFont = (style.cl as { rgb?: string } | undefined)?.rgb
          cell.s = {
            ...style,
            ...(headerFill ? { bg: { rgb: headerFill } } : {}),
            ...(explicitFont && explicitFont !== '#000000' ? {} : { cl: { rgb: fontColor } }),
            bl: BooleanNumber.TRUE,
          }
          continue
        }
        if (styleHasFill(style)) continue
        if (isTotals) {
          cell.s = {
            ...style,
            ...(table.totalRowFill ? { bg: { rgb: table.totalRowFill } } : {}),
            ...(table.totalRowFontColor ? { cl: { rgb: table.totalRowFontColor } } : {}),
            bl: BooleanNumber.TRUE,
          }
          continue
        }
        // Band precedence below header/totals: first/last column emphasis,
        // then row stripes, then column stripes, then the whole-table fill.
        const isFirstColumn = column === table.range.startColumn && table.firstColumnFill
        const isLastColumn = column === table.range.endColumn && table.lastColumnFill
        const columnStripeFill = table.showColumnStripes
          ? (column - table.range.startColumn) % 2 === 0
            ? table.columnStripeFill
            : table.secondColumnStripeFill
          : undefined
        const fill = isFirstColumn
          ? table.firstColumnFill
          : isLastColumn
            ? table.lastColumnFill
            : ((isStripe ? stripeFill : undefined) ??
              secondStripeFill ??
              columnStripeFill ??
              table.wholeTableFill)
        if (fill) {
          // Dark families set a body text color; a default-black font yields
          // to it (explicit cell colors survive, mirroring the header rule).
          const cellFont = (style.cl as { rgb?: string } | undefined)?.rgb
          const fontPatch =
            table.bodyFontColor && (!cellFont || cellFont === '#000000')
              ? { cl: { rgb: table.bodyFontColor } }
              : {}
          cell.s = { ...style, bg: { rgb: fill }, ...fontPatch }
        }
      }
    }
  }
}

/// One pivot style band: an absent value inherits the band below it.
interface PivotBand {
  fill?: string | undefined
  fontColor?: string | undefined
  bold?: boolean | undefined
}

/// Excel keeps pivot styling out of cell xfs entirely; paint the style bands
/// resolved sidecar-side from pivotTableStyleInfo (calibrated against Excel
/// for Mac: genoffice-sample/sheets/calib/pivot-style-truths.json).
///
/// Precedence, lowest first: wholeTable, row stripe, column stripe,
/// firstColumn (row-label columns), subheading / subtotal, header (+ the
/// first header cell), grand-total row. Row stripes count from firstDataRow
/// over every body row (subheadings and subtotals included; a band with its
/// own fill covers them); column stripes count from firstDataCol and skip
/// the header rows. Row kinds come from the sidecar's `rowKinds`; without
/// them the rows are plain data with the last one the grand total.
export function applyPivotStyling(
  matrix: ICellData[][],
  range: IRange,
  pivotTables: WorkbookFile['sheets'][number]['pivotTables'],
): void {
  for (const pivot of pivotTables) {
    // Fills imply a named style; the extra checks keep stale sidecars working.
    if (!pivot.styled && !pivot.headerFill && !pivot.wholeTableFill && !pivot.stripeFill) continue
    let bounds: ReturnType<typeof parseRange>
    try {
      bounds = parseRange(pivot.outputRef)
    } catch {
      continue
    }
    const headerEndRow = bounds.startRow + (pivot.firstDataRow ?? 1) - 1
    const firstDataColumn = bounds.startColumn + (pivot.firstDataCol ?? 1)
    const rowGrandTotals = pivot.rowGrandTotals ?? true
    const rowKinds = pivot.rowKinds ?? ''
    const rowStart = Math.max(range.startRow, bounds.startRow)
    const rowEnd = Math.min(range.endRow, bounds.endRow)
    const columnStart = Math.max(range.startColumn, bounds.startColumn)
    const columnEnd = Math.min(range.endColumn, bounds.endColumn)
    // An absent bold flag inherits (the calibrated palettes spell out every
    // band that is bold); only a style the sidecar could not resolve (custom
    // name, stale binary) falls back to bold header and grand-total rows.
    const resolved =
      pivot.headerBold !== undefined ||
      pivot.totalRowBold !== undefined ||
      pivot.headerFill !== undefined ||
      pivot.wholeTableFill !== undefined
    const fallbackBold = resolved ? undefined : true
    const wholeTable: PivotBand = {
      fill: pivot.wholeTableFill,
      fontColor: pivot.wholeTableFontColor,
    }
    const firstColumn: PivotBand = { fill: pivot.firstColumnFill, bold: pivot.firstColumnBold }
    const header: PivotBand = {
      fill: pivot.headerFill,
      fontColor: pivot.headerFontColor,
      bold: pivot.headerBold ?? fallbackBold,
    }
    const firstHeaderCell: PivotBand = {
      fontColor: pivot.firstHeaderCellFontColor,
      bold: pivot.firstHeaderCellBold,
    }
    const rowBands: Record<string, PivotBand | undefined> = {
      s: {
        fill: pivot.subheadingFill,
        fontColor: pivot.subheadingFontColor,
        bold: pivot.subheadingBold,
      },
      S: {
        fill: pivot.subheading2Fill,
        fontColor: pivot.subheading2FontColor,
        bold: pivot.subheading2Bold,
      },
      t: {
        fill: pivot.subtotalFill,
        fontColor: pivot.subtotalFontColor,
        bold: pivot.subtotalBold,
      },
      g: {
        fill: pivot.totalRowFill,
        fontColor: pivot.totalRowFontColor,
        bold: pivot.totalRowBold ?? fallbackBold,
      },
    }
    for (let row = rowStart; row <= rowEnd; row += 1) {
      const isHeader = row <= headerEndRow
      const rowOffset = row - headerEndRow - 1
      const kind = isHeader
        ? 'h'
        : (rowKinds[rowOffset] ?? (rowGrandTotals && row === bounds.endRow ? 'g' : 'd'))
      const rowStripe = isHeader
        ? undefined
        : rowOffset % 2 === 0
          ? pivot.stripeFill
          : pivot.secondRowStripeFill
      const rowBand = rowBands[kind]
      for (let column = columnStart; column <= columnEnd; column += 1) {
        const cell = matrix[row - range.startRow]?.[column - range.startColumn]
        if (!cell) continue
        const isFirstColumn = column < firstDataColumn
        const columnStripe =
          isHeader || isFirstColumn
            ? undefined
            : (column - firstDataColumn) % 2 === 0
              ? pivot.columnStripeFill
              : pivot.secondColumnStripeFill
        const layers: (PivotBand | undefined)[] = [
          wholeTable,
          { fill: rowStripe },
          { fill: columnStripe },
          isFirstColumn ? firstColumn : undefined,
          rowBand,
          isHeader ? header : undefined,
          isHeader && row === bounds.startRow && column === bounds.startColumn
            ? firstHeaderCell
            : undefined,
        ]
        let fill: string | undefined
        let fontColor: string | undefined
        let bold = false
        for (const layer of layers) {
          if (!layer) continue
          if (layer.fill !== undefined) fill = layer.fill
          if (layer.fontColor !== undefined) fontColor = layer.fontColor
          if (layer.bold !== undefined) bold = layer.bold
        }
        if (!fill && !fontColor && !bold) continue
        const style = (cell.s ?? {}) as IStyleData
        // Explicit cell fills and non-black font colors survive the band.
        const cellFont = (style.cl as { rgb?: string } | undefined)?.rgb
        const paintFill = fill && !styleHasFill(style)
        const paintFont = fontColor && (!cellFont || cellFont === '#000000')
        if (!paintFill && !paintFont && !bold) continue
        cell.s = {
          ...style,
          ...(paintFill ? { bg: { rgb: fill } } : {}),
          ...(paintFont ? { cl: { rgb: fontColor } } : {}),
          ...(bold ? { bl: BooleanNumber.TRUE } : {}),
        }
      }
    }
  }
}

/// The cell xf's font as a text style — the base a rich-text document
/// inherits. Univer applies a cell's `s` font only to plain `v` cells; a `p`
/// document renders purely from its textRuns, so runs without explicit
/// formatting must carry the cell font themselves (Excel semantics: a run
/// without rPr uses the cell font).
export function cellFontTextStyle(style: WorkbookCellStyle | undefined): IStyleData {
  if (!style) return {}
  return {
    ...(style.fontFamily ? { ff: escapeCssLeadingDigit(style.fontFamily) } : {}),
    ...(style.fontSize ? { fs: style.fontSize } : {}),
    ...(style.bold ? { bl: BooleanNumber.TRUE } : {}),
    ...(style.italic ? { it: BooleanNumber.TRUE } : {}),
    ...(style.underline ? { ul: { s: BooleanNumber.TRUE } } : {}),
    ...(style.strikethrough ? { st: { s: BooleanNumber.TRUE } } : {}),
    ...(style.fontColor ? { cl: { rgb: style.fontColor } } : {}),
  }
}

/// The font subset of an already-composed Univer style — the rich-document
/// base for journal-replayed cells, whose family is already CSS-escaped.
function fontTextStyleOf(s: IStyleData | null | undefined): IStyleData {
  if (!s) return {}
  return {
    ...(s.ff ? { ff: s.ff } : {}),
    ...(s.fs ? { fs: s.fs } : {}),
    ...(s.bl ? { bl: s.bl } : {}),
    ...(s.it ? { it: s.it } : {}),
    ...(s.ul?.s ? { ul: { s: BooleanNumber.TRUE } } : {}),
    ...(s.st?.s ? { st: { s: BooleanNumber.TRUE } } : {}),
    ...(s.cl?.rgb ? { cl: { rgb: s.cl.rgb } } : {}),
  }
}

/// Guards the float underflow in size * (shrunk / measured): 9 * (6/9)
/// is 5.999…, which must floor to 6, not 5.
function scaledFontSize(size: number, scale: number): number {
  return Math.max(1, Math.floor(size * scale + 1e-6))
}

/// Excel shrink-to-fit: scale the font down until the widest line fits the
/// column (never wraps, never enlarges). Integer result — Univer ceils
/// fractional font sizes, which would overflow again. Excel keeps shrinking
/// as far as needed; clamp at 1pt to stay renderable.
export function shrinkToFitFontSize(
  text: string,
  fontSize: number,
  availablePx: number,
  measure: (line: string) => number,
): number | null {
  if (!(availablePx > 0)) return null
  let widest = 0
  for (const line of text.split(/\r\n|[\r\n]/)) widest = Math.max(widest, measure(line))
  if (!(widest > availablePx)) return null
  return Math.max(1, Math.floor((fontSize * availablePx) / widest))
}

/// Shrink factor (≤1) for a shrinkToFit cell; 1 when the text already fits.
/// Rich runs are approximated with the cell font at the largest size in play
/// — per-run measurement isn't worth it for a fit heuristic.
function shrinkScaleFor(
  worksheet: UniverWorksheet,
  row: number,
  column: number,
  text: string,
  style: WorkbookCellStyle,
  runs?: readonly WorkbookRichRun[],
): number {
  const sheet = worksheet.getSheet()
  // A merged cell's budget is the merge span's total width; only the anchor
  // carries the text (covered cells have nothing to shrink).
  const merge = sheet.getMergedCell(row, column)
  if (merge && (row !== merge.startRow || column !== merge.startColumn)) return 1
  let cellWidth = 0
  for (let c = merge?.startColumn ?? column; c <= (merge?.endColumn ?? column); c += 1) {
    cellWidth += sheet.getColumnWidth(c)
  }
  const measureSize = Math.max(style.fontSize ?? 11, ...(runs ?? []).map((run) => run.size ?? 0))
  const { fontString } = getFontStyleString({ ...cellFontTextStyle(style), fs: measureSize })
  // 2+2 cell padding plus 1px blur offset (matches the #### rule), plus the
  // indent's left padding.
  const available = cellWidth - 5 - (style.indent ? style.indent * INDENT_STEP_PX : 0)
  const shrunk = shrinkToFitFontSize(text, measureSize, available, (line) =>
    line === '' ? 0 : FontCache.getMeasureText(line, fontString).width,
  )
  return shrunk === null ? 1 : shrunk / measureSize
}

/// A run with any explicit formatting came from an rPr, which is
/// authoritative for its boolean flags; a bare run inherits the cell font.
function runHasFormatting(run: WorkbookRichRun): boolean {
  return (
    run.bold ||
    run.italic ||
    run.underline ||
    run.strikethrough ||
    run.color !== undefined ||
    run.size !== undefined ||
    run.family !== undefined ||
    run.vertAlign !== undefined
  )
}

export function toRichTextDocument(
  text: string,
  runs: readonly WorkbookRichRun[] = [],
  // The cell font (cellFontTextStyle / fontTextStyleOf) — Univer applies a
  // cell's `s` font only to plain `v` cells, so a rich document's runs must
  // carry it themselves.
  base: IStyleData = {},
  // shrinkToFit factor (≤1): run-level sizes must scale along with the cell
  // font, or an rPr-sized run would keep overflowing the shrunken cell.
  fontScale = 1,
): ICellData['p'] {
  // Excel stores hard line breaks as CRLF (or bare \n); Univer's paragraph
  // break is a single \r — an unnormalized \r\n would leave a phantom empty
  // paragraph behind every break.
  const normalize = (value: string): string => value.replace(/\r\n?/g, '\n')
  const normalized = normalize(text)
  const hasBase = Object.keys(base).length > 0
  const textRuns = []
  let cursor = 0
  for (const run of runs) {
    const end = cursor + normalize(run.text).length
    textRuns.push({
      st: cursor,
      ed: end,
      ts: runHasFormatting(run)
        ? {
            ...base,
            ...(run.family ? { ff: escapeCssLeadingDigit(run.family) } : {}),
            ...(run.size ? { fs: scaledFontSize(run.size, fontScale) } : {}),
            bl: run.bold ? BooleanNumber.TRUE : BooleanNumber.FALSE,
            it: run.italic ? BooleanNumber.TRUE : BooleanNumber.FALSE,
            ul: { s: run.underline ? BooleanNumber.TRUE : BooleanNumber.FALSE },
            st: { s: run.strikethrough ? BooleanNumber.TRUE : BooleanNumber.FALSE },
            ...(run.color ? { cl: { rgb: run.color } } : {}),
            ...(run.vertAlign
              ? {
                  va:
                    run.vertAlign === 'subscript'
                      ? BaselineOffset.SUBSCRIPT
                      : BaselineOffset.SUPERSCRIPT,
                }
              : {}),
          }
        : base,
    })
    cursor = end
  }
  if (hasBase && cursor < normalized.length) {
    textRuns.push({ st: cursor, ed: normalized.length, ts: base })
  }
  // Univer document streams use \r as paragraph break and \n as section
  // break; a raw \n would split the cell into sections and drop later lines.
  // 1:1 replacement, so textRun offsets stay valid.
  const dataStream = `${normalized.replace(/\n/g, '\r')}\r\n`
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
    let bounds: IRange
    try {
      bounds = parseRange(cell.arrayRef)
    } catch {
      continue
    }
    if (rangeCellCount(bounds) > 100_000) continue
    // Exact rectangles, not the envelope: a follower marking blanks the cell,
    // so unrelated lines moved or left between the survivors must stay out.
    const rects = ops.length > 0 ? fileRangeToScreenRanges(ops, bounds) : [bounds]
    for (const rect of rects) {
      for (let row = rect.startRow; row <= rect.endRow; row += 1) {
        for (let column = rect.startColumn; column <= rect.endColumn; column += 1) {
          if (row === cell.row && column === cell.column) continue
          followers.add(`${row}:${column}`)
        }
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
    const rowsPerBlock = Math.max(1, Math.floor(SIDECAR_RANGE_CELL_LIMIT / sheet.columnCount))
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
      recordCachedFormulaValues(state, sheetId, screen.cells)
      const installable = degradeCostlyFormulas(state, sheet.name, screen.cells)
      collectArrayFollowers(arrayFollowers, installable, ops)
      recordRowStyleKeys(state, sheetId, screen.rows)
      recordHiddenFileRows(state, sheetId, screen.rows, screenRange)
      patchWorksheetRange(
        worksheet,
        undefined,
        screenRange,
        installable,
        state.file.styles,
        screen.hyperlinks,
        sheet.tables,
        sheet.pivotTables,
        sheet.freeze,
        true,
        state.editJournal,
        undefined,
        undefined,
        arrayFollowers,
        sheetRowColStyleKeys(state, sheetId),
        inheritedWrapLookup(state.file.styles, screen.rows, sheet.columnWidths),
        hiddenRowsInfo(state, sheetId),
      )
      if (state.formulaMode) storeFormulaText(state, sheetId, result.cells)
      recordHyperlinks(state, sheetId, screen.hyperlinks)
      keepActiveSheet(worksheet, () => {
        applyRowProperties(worksheet, state, sheetId, screen.rows)
        applyMerges(worksheet, state, sheetId, screen.merges)
      })
      const qualifying = wrapAutoFitRows(
        screen.cells,
        state.file.styles,
        screen.rows,
        sheet.columnWidths,
        sheet.defaultRowHeightFixed,
        sheet.defaultRowHeight,
        screenRange,
        screen.merges,
      )
      measureWrapAutoFitRows(worksheet, qualifying)
      const sheetKey = `file-${state.file.sha256}:${sheetId}`
      if (result.indexingComplete) {
        const qualifyingWithoutMerges = wrapAutoFitRows(
          screen.cells,
          state.file.styles,
          screen.rows,
          sheet.columnWidths,
          sheet.defaultRowHeightFixed,
          sheet.defaultRowHeight,
          screenRange,
        )
        resetStaleWrapAutoHeights(
          runtime,
          `file-${state.file.sha256}`,
          worksheet,
          takeContaminatedRows(sheetKey, qualifyingWithoutMerges, qualifying),
          screen.rows,
          sheet.defaultRowHeight,
        )
      } else {
        trackPreIndexMeasuredRows(sheetKey, qualifying)
      }
      if (ops.length === 0) {
        if (result.indexingComplete) {
          await applyConditionalRules(worksheet, state, sheetId, result.conditionalRules)
          applySheetFilter(worksheet, state, sheetId, result.autoFilter)
          applyDataValidations(runtime, state, sheetId, result.dataValidations)
          state.decorationsPendingSheets.delete(sheetId)
        } else {
          state.decorationsPendingSheets.add(sheetId)
        }
      } else {
        state.decorationsPendingSheets.delete(sheetId)
      }
      if (result.indexingComplete) captureSheetFileState(state, sheetId, result)
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

/// Records the sheet-wide, complete-only file state (protection, manual page
/// breaks, allow-edit ranges) the first time a sheet finishes indexing.
function captureSheetFileState(
  state: LazyWorkbookState,
  sheetId: string,
  result: WorkbookRangeResult,
): void {
  if (!state.sheetProtections.has(sheetId)) {
    state.sheetProtections.set(
      sheetId,
      result.sheetProtection ?? { protected: false, hasPassword: false },
    )
  }
  if (!state.sheetPageBreaks.has(sheetId)) {
    state.sheetPageBreaks.set(sheetId, {
      rowBreaks: [...result.rowBreaks],
      colBreaks: [...result.colBreaks],
    })
  }
  if (!state.sheetFilePageSetups.has(sheetId) && result.pageSetup != null) {
    state.sheetFilePageSetups.set(sheetId, result.pageSetup)
  }
  if (!state.sheetProtectedRanges.has(sheetId)) {
    // File coordinates → this session's screen space; later structural ops
    // remap the stored set incrementally (see the App structural listener).
    state.sheetProtectedRanges.set(
      sheetId,
      mapProtectedRanges(
        result.protectedRanges,
        state.editJournal.structuralOps.get(sheetId) ?? [],
      ),
    )
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
export function toUniverDvRule(
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
    // Univer overloads a list rule's formula2 as its per-item color list;
    // file list rules may carry a junk formula2 (LibreOffice writes "0"),
    // which would paint validated cells with that "color" (black).
    ...(formula2 === undefined || type === 'list' ? {} : { formula2 }),
    ...(type === 'list'
      ? {
          showDropDown: !rule.suppressDropdown,
          // Text mode preserves the workbook's normal cell appearance. The
          // app overlays only the small dropdown arrow on populated cells.
          renderMode: DataValidationRenderMode.TEXT,
        }
      : {}),
    ...(rule.showInputMessage ? { showInputMessage: true } : {}),
    ...(rule.showErrorMessage ? { showErrorMessage: true } : {}),
    ...(errorStyle === undefined ? {} : { errorStyle }),
    ...(rule.errorTitle === undefined ? {} : { errorTitle: rule.errorTitle }),
    ...(rule.error === undefined ? {} : { error: rule.error }),
    ...(rule.promptTitle === undefined ? {} : { promptTitle: rule.promptTitle }),
    ...(rule.prompt === undefined ? {} : { prompt: rule.prompt }),
  }
}

/// Excel paints only the highest-precedence rule of a "paint-once" type
/// (colorScale / dataBar / iconSet) on a cell; stacked rules of the same
/// type do not blend. Univer keeps whichever it evaluates last, so a
/// lower-precedence duplicate could win. Drop a rule whose every range is
/// fully covered by a higher-precedence same-type rule — the common shape
/// (tdf105272 carries three identical stacked scales plus the real one).
const PAINT_ONCE_TYPES = new Set(['colorScale', 'dataBar', 'iconSet'])

export function dropShadowedPaintOnceRules(
  rules: WorkbookRangeResult['conditionalRules'],
): WorkbookRangeResult['conditionalRules'] {
  const covers = (
    outer: (typeof rules)[number]['ranges'][number],
    inner: (typeof rules)[number]['ranges'][number],
  ): boolean =>
    outer.startRow <= inner.startRow &&
    outer.endRow >= inner.endRow &&
    outer.startColumn <= inner.startColumn &&
    outer.endColumn >= inner.endColumn
  return rules.filter(
    (rule) =>
      !PAINT_ONCE_TYPES.has(rule.ruleType) ||
      !rules.some(
        (other) =>
          other !== rule &&
          other.ruleType === rule.ruleType &&
          other.priority < rule.priority &&
          rule.ranges.every((range) => other.ranges.some((cover) => covers(cover, range))),
      ),
  )
}

async function applyConditionalRules(
  worksheet: UniverWorksheet,
  state: LazyWorkbookState,
  sheetId: string,
  rules: WorkbookRangeResult['conditionalRules'],
): Promise<void> {
  if (rules.length === 0 || state.appliedCfSheets.has(sheetId)) return
  state.appliedCfSheets.add(sheetId)
  // Lower xlsx priority number = higher precedence; Univer's addRule
  // unshifts, so the rule added LAST sits first and wins conflicts — add in
  // descending priority. Installing the file's own rules must not mark the
  // sheet's CF as edited.
  const ordered = [...dropShadowedPaintOnceRules(rules)].sort((a, b) => b.priority - a.priority)
  // Resolve name/reference cfvos to numbers BEFORE the suppression window —
  // the sidecar round-trips must not sit inside journalSuppression.
  const prepared = []
  for (const rule of ordered) {
    try {
      prepared.push(
        await resolveDataBarLayout(state, sheetId, await resolveRuleCfvos(state, sheetId, rule)),
      )
    } catch {
      // Resolution is best-effort: an odd reference must not cost the
      // sheet its conditional formatting (this sheet is already marked
      // applied, so a throw here would skip CF permanently).
      prepared.push(rule)
    }
  }
  journalSuppression.active = true
  try {
    for (const rule of prepared) {
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

/// Bar geometry Univer cannot express directly: x14 autoMin/autoMax bounds
/// (anchored at zero for one-signed data), 2006/x14 minLength/maxLength
/// extents, and the x14 cell-midpoint axis. Each becomes a pair of static
/// num bounds computed from the rule range's cached values — see
/// cf-thresholds.ts for the math and its limits.
async function resolveDataBarLayout(
  state: LazyWorkbookState,
  sheetId: string,
  rule: WorkbookRangeResult['conditionalRules'][number],
): Promise<WorkbookRangeResult['conditionalRules'][number]> {
  if (rule.ruleType !== 'dataBar' || !dataBarNeedsLayout(rule)) return rule
  const values = dataBarNeedsValues(rule)
    ? await readCachedValues(state, sheetId, rule.ranges, CF_AUTO_BOUNDS_CELL_CAP)
    : null
  const cfvos = layoutDataBar(rule, values)
  return cfvos === null ? rule : { ...rule, cfvos }
}

/// Scale cfvos (dataBar/colorScale/iconSet) whose value is a formula, a
/// defined name, or a cell reference: fold them to numbers against the
/// sidecar's cached cell values (Univer's formula registry only sees the
/// streamed window and knows no file table names). An unfoldable formula
/// takes the slot's default threshold rather than a stop Univer cannot
/// compute, which would leave the rule unpainted.
async function resolveRuleCfvos(
  state: LazyWorkbookState,
  sheetId: string,
  rule: WorkbookRangeResult['conditionalRules'][number],
): Promise<WorkbookRangeResult['conditionalRules'][number]> {
  if (!['dataBar', 'colorScale', 'iconSet'].includes(rule.ruleType) || rule.cfvos.length === 0) {
    return rule
  }
  const needsWork = rule.cfvos.some(
    (cfvo) =>
      cfvo.kind === 'formula' ||
      (cfvo.value !== undefined &&
        !Number.isFinite(Number(cfvo.value)) &&
        ['num', 'percent', 'percentile'].includes(cfvo.kind)),
  )
  let cfvos = rule.cfvos
  if (needsWork) {
    const reader = thresholdReader(state, sheetId)
    const resolvedCfvos: typeof cfvos = []
    for (const [index, cfvo] of rule.cfvos.entries()) {
      if (cfvo.kind === 'formula') {
        const resolved =
          cfvo.value === undefined ? null : await evaluateThresholdFormula(cfvo.value, reader)
        if (resolved !== null) {
          resolvedCfvos.push({ ...cfvo, kind: 'num', value: String(resolved) })
        } else if (cfvo.value !== undefined && isSelfContainedFormula(cfvo.value)) {
          // `TODAY()-30` and friends: no cell input, so Univer's own engine
          // computes them correctly whatever the streamed window holds.
          resolvedCfvos.push(cfvo)
        } else {
          resolvedCfvos.push({
            ...defaultThreshold(rule.ruleType, index, rule.cfvos.length),
            gte: cfvo.gte,
          })
        }
        continue
      }
      if (
        cfvo.value === undefined ||
        Number.isFinite(Number(cfvo.value)) ||
        !['num', 'percent', 'percentile'].includes(cfvo.kind)
      ) {
        resolvedCfvos.push(cfvo)
        continue
      }
      // Legacy writers put defined names / expressions into typed cfvos.
      const resolved = await evaluateThresholdFormula(cfvo.value, reader)
      resolvedCfvos.push(
        resolved === null
          ? { ...defaultThreshold(rule.ruleType, index, rule.cfvos.length), gte: cfvo.gte }
          : { ...cfvo, value: String(resolved) },
      )
    }
    cfvos = resolvedCfvos
  }
  if (rule.ruleType === 'colorScale') cfvos = clampColorScaleStops(cfvos)
  return cfvos === rule.cfvos ? rule : { ...rule, cfvos }
}

/// Sidecar-backed lookups for the threshold evaluator, scoped to the sheet
/// the rule lives on.
function thresholdReader(state: LazyWorkbookState, sheetId: string): ThresholdReader {
  const sheetByName = (name: string | null) =>
    name === null
      ? state.file.sheets.find((entry) => entry.id === sheetId)
      : state.file.sheets.find((entry) => entry.name.toLowerCase() === name.toLowerCase())
  return {
    async readValues(sheetName, range) {
      const sheet = sheetByName(sheetName)
      if (!sheet) return null
      // Clip to the sheet's extent: a `$B$2:$B$1000` over a 20-row sheet
      // must read the 19 rows, not fail the whole threshold.
      const clipped = {
        startRow: range.startRow,
        startColumn: range.startColumn,
        endRow: Math.min(range.endRow, sheet.rowCount - 1),
        endColumn: Math.min(range.endColumn, sheet.columnCount - 1),
      }
      if (clipped.endRow < clipped.startRow || clipped.endColumn < clipped.startColumn) return []
      return readCachedValues(state, sheet.id, [clipped], THRESHOLD_RANGE_CELL_CAP)
    },
    definedName(name) {
      const lower = name.toLowerCase()
      return (
        state.file.definedNames.find((entry) => entry.name.toLowerCase() === lower)?.formula ?? null
      )
    },
    tableColumn(tableName, columnName) {
      for (const sheet of state.file.sheets) {
        const table = sheet.tables.find((entry) => entry.name === tableName)
        if (!table) continue
        const columnIndex = table.columns?.indexOf(columnName) ?? -1
        if (columnIndex < 0) return null
        const column = table.range.startColumn + columnIndex
        const startRow = table.range.startRow + table.headerRowCount
        const endRow = table.range.endRow - (table.totalsRowCount ?? 0)
        if (endRow < startRow) return null
        return {
          sheetName: sheet.name,
          range: { startRow, endRow, startColumn: column, endColumn: column },
        }
      }
      return null
    },
  }
}

/// Numeric values of the populated cells of `ranges` (NaN for text), read
/// from the sidecar cache in preload-sized chunks; null when the ranges
/// exceed `cap` cells or a read fails.
async function readCachedValues(
  state: LazyWorkbookState,
  sheetId: string,
  ranges: readonly { startRow: number; endRow: number; startColumn: number; endColumn: number }[],
  cap: number,
): Promise<number[] | null> {
  const totalCells = ranges.reduce(
    (sum, area) =>
      sum + (area.endRow - area.startRow + 1) * (area.endColumn - area.startColumn + 1),
    0,
  )
  if (totalCells > cap) return null
  const values: number[] = []
  for (const area of ranges) {
    // The preload rejects reads above MAX_RANGE_CELLS — chunk by rows.
    const columns = area.endColumn - area.startColumn + 1
    const rowsPerChunk = Math.max(1, Math.floor(SIDECAR_RANGE_CELL_LIMIT / columns))
    for (let startRow = area.startRow; startRow <= area.endRow; startRow += rowsPerChunk) {
      const cells = await readCachedRange(state, sheetId, {
        startRow,
        endRow: Math.min(area.endRow, startRow + rowsPerChunk - 1),
        startColumn: area.startColumn,
        endColumn: area.endColumn,
      })
      if (cells === null) return null
      for (const value of cells) values.push(value)
    }
  }
  return values
}

async function readCachedRange(
  state: LazyWorkbookState,
  sheetId: string,
  range: { startRow: number; endRow: number; startColumn: number; endColumn: number },
): Promise<number[] | null> {
  // The referenced sheet may still be streaming when the rule's own sheet
  // renders (the budget's bar on sheet 1 sums a table on sheet 2) — an
  // incomplete read yields empty cells and a broken zero scale, so wait for
  // the index within a bounded window.
  const deadline = Date.now() + 15_000
  try {
    for (;;) {
      const result = await window.desktopApi.readWorkbookRange({
        sessionId: state.file.sessionId,
        sheetId,
        range,
      })
      // Row coverage is enough — matching how the streaming loader treats a
      // range as ready — so big sheets don't stall on full indexing.
      if (
        result.indexingComplete ||
        (result.indexedThroughRow !== null && result.indexedThroughRow >= range.endRow)
      ) {
        return result.cells.map((cell) => {
          const value = typeof cell.value === 'number' ? cell.value : Number(cell.value)
          return Number.isFinite(value) ? value : Number.NaN
        })
      }
      if (Date.now() > deadline) return null
      await new Promise((resolve) => setTimeout(resolve, 400))
    }
  } catch {
    return null
  }
}

type CfHighlightBuilder = ReturnType<
  ReturnType<UniverWorksheet['newConditionalFormattingRule']>['whenCellNotEmpty']
>

export function buildConditionalRule(
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
    const positive = rule.colors[0] ?? '#638EC6'
    return builder
      .setDataBar({
        min: toCfValue(min),
        max: toCfValue(max),
        positiveColor: positive,
        // Explicit x14 negative fill wins; an x14 twin flagged
        // same-as-positive reuses the positive fill; otherwise Excel's
        // default negative fill is red.
        nativeColor: rule.negativeColor ?? (rule.negativeSameAsPositive ? positive : '#FF0000'),
        isShowValue: rule.showValue,
        isGradient: rule.gradient ?? true,
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
  // Univer offsets relative CF formulas from the top-left-sorted first range,
  // not the file's sqref order.
  const first = [...ranges].sort(
    (a, b) => a.startRow - b.startRow || a.startColumn - b.startColumn,
  )[0]
  const anchor = first ? `${columnLetter(first.startColumn)}${first.startRow + 1}` : 'A1'
  const coveredCells = ranges.reduce(
    (sum, r) => sum + (r.endRow - r.startRow + 1) * (r.endColumn - r.startColumn + 1),
    0,
  )
  const highlight = buildHighlightCondition(builder, rule, anchor, coveredCells)
  if (!highlight) return null
  const built = applyDxfFormat(highlight, dxfStyles, rule.dxfIndex).setRanges(ranges).build()
  patchBuiltHighlightRule(
    built.rule as BuiltHighlightRule,
    rule.ruleType,
    rule.dxfIndex === undefined ? undefined : dxfStyles[rule.dxfIndex],
  )
  // The facade builder has no setter for cfRule/@stopIfTrue; Univer's
  // evaluator honors the flag on the installed rule (matched → skip
  // lower-priority rules). Only highlight rules carry it — Excel's UI does
  // not offer stop-if-true on color scales / data bars / icon sets.
  if (rule.stopIfTrue === true) {
    ;(built as { stopIfTrue?: boolean }).stopIfTrue = true
  }
  return built
}

export type BuiltHighlightRule = {
  operator?: string
  value?: string
  /// The builder's IStyleBase plus the patched-on keys.
  style?: {
    n?: { pattern: string }
    bd?: Partial<Record<'t' | 'b' | 'l' | 'r', ReturnType<typeof toUniverBorder>>>
    [key: string]: unknown
  }
}

/// IHighlightCell supports more than the facade builder exposes: style is an
/// IStyleBase (the paint path merges its n and bd verbatim), and the
/// error-presence text operators have no setter at all. Patch the built rule
/// for a dxf number format, dxf borders, and containsErrors /
/// notContainsErrors (built as a placeholder text condition in
/// buildHighlightCondition).
export function patchBuiltHighlightRule(
  target: BuiltHighlightRule,
  ruleType: string,
  dxf: WorkbookCellStyle | undefined,
): void {
  if (ruleType === 'containsErrors' || ruleType === 'notContainsErrors') {
    target.operator = ruleType
    delete target.value
  }
  if (!dxf) return
  // A numFmt-only dxf builds with no style object at all.
  if (dxf.numberFormat) {
    target.style = { ...target.style, n: { pattern: dxf.numberFormat } }
  }
  const bd = {
    ...(dxf.borderTop ? { t: toUniverBorder(dxf.borderTop) } : {}),
    ...(dxf.borderBottom ? { b: toUniverBorder(dxf.borderBottom) } : {}),
    ...(dxf.borderLeft ? { l: toUniverBorder(dxf.borderLeft) } : {}),
    ...(dxf.borderRight ? { r: toUniverBorder(dxf.borderRight) } : {}),
  }
  if (Object.keys(bd).length > 0) target.style = { ...target.style, bd }
}

function toCfValue(cfvo: { kind: string; value?: string | undefined }): IValueConfig {
  switch (cfvo.kind) {
    case 'min':
    case 'autoMin':
      return { type: CFValueType.min }
    case 'max':
    case 'autoMax':
      return { type: CFValueType.max }
    case 'percent':
      return { type: CFValueType.percent, value: Number(cfvo.value ?? 0) }
    case 'percentile':
      return { type: CFValueType.percentile, value: Number(cfvo.value ?? 0) }
    case 'formula':
      return { type: CFValueType.formula, value: toCfFormula(cfvo.value ?? '0') }
    default: {
      // Legacy writers put defined names / expressions into type="num" cfvos;
      // a NaN literal would collapse the scale to 0.
      const numeric = Number(cfvo.value ?? 0)
      return Number.isFinite(numeric)
        ? { type: CFValueType.num, value: numeric }
        : { type: CFValueType.formula, value: toCfFormula(cfvo.value ?? '0') }
    }
  }
}

/// Univer's CF formula service takes '='-prefixed formulas (same registry as
/// whenFormulaSatisfied); xlsx cfvo bodies come without the prefix.
function toCfFormula(body: string): string {
  return body.startsWith('=') ? body : `=${body}`
}

/// Formula CF costs one dependency tree per covered cell; above this, huge
/// (e.g. whole-column) rules keep the cheaper native condition.
const CELLIS_FORMULA_CELL_LIMIT = 20_000
/// Above this, auto dataBar bounds fall back to Univer's raw data extremes
/// instead of paying a chunked full-range read on open.
const CF_AUTO_BOUNDS_CELL_CAP = 512_000

/// Excel evaluates numeric cellIs rules on blank cells as 0; Univer's native
/// number conditions skip blanks (matching only notEqual/notBetween). True
/// when the two would paint blanks differently.
export function cellIsBlankDiverges(operator: string, first: number, second: number): boolean {
  let excelBlank: boolean
  switch (operator) {
    case 'greaterThan':
      excelBlank = 0 > first
      break
    case 'greaterThanOrEqual':
      excelBlank = 0 >= first
      break
    case 'lessThan':
      excelBlank = 0 < first
      break
    case 'lessThanOrEqual':
      excelBlank = 0 <= first
      break
    case 'equal':
      excelBlank = first === 0
      break
    case 'notEqual':
      excelBlank = first !== 0
      break
    case 'between':
      excelBlank = Math.min(first, second) <= 0 && 0 <= Math.max(first, second)
      break
    case 'notBetween':
      excelBlank = !(Math.min(first, second) <= 0 && 0 <= Math.max(first, second))
      break
    default:
      return false
  }
  const univerBlank = operator === 'notEqual' || operator === 'notBetween'
  return excelBlank !== univerBlank
}

function buildCellIsFormula(
  builder: ReturnType<UniverWorksheet['newConditionalFormattingRule']>,
  operator: string,
  anchor: string,
  first: number,
  second: number,
): CfHighlightBuilder | null {
  switch (operator) {
    case 'equal':
      return builder.whenFormulaSatisfied(`=${anchor}=${first}`)
    case 'notEqual':
      return builder.whenFormulaSatisfied(`=${anchor}<>${first}`)
    case 'greaterThan':
      return builder.whenFormulaSatisfied(`=${anchor}>${first}`)
    case 'greaterThanOrEqual':
      return builder.whenFormulaSatisfied(`=${anchor}>=${first}`)
    case 'lessThan':
      return builder.whenFormulaSatisfied(`=${anchor}<${first}`)
    case 'lessThanOrEqual':
      return builder.whenFormulaSatisfied(`=${anchor}<=${first}`)
    case 'between':
      return builder.whenFormulaSatisfied(
        `=AND(${anchor}>=${Math.min(first, second)},${anchor}<=${Math.max(first, second)})`,
      )
    case 'notBetween':
      return builder.whenFormulaSatisfied(
        `=NOT(AND(${anchor}>=${Math.min(first, second)},${anchor}<=${Math.max(first, second)}))`,
      )
    default:
      return null
  }
}

function buildHighlightCondition(
  builder: ReturnType<UniverWorksheet['newConditionalFormattingRule']>,
  rule: WorkbookRangeResult['conditionalRules'][number],
  anchor: string,
  coveredCells = 0,
): CfHighlightBuilder | null {
  const firstNumber = Number(rule.formulas[0])
  const secondNumber = Number(rule.formulas[1])
  switch (rule.ruleType) {
    case 'cellIs':
      if (!Number.isFinite(firstNumber)) return buildCellIsNonNumeric(builder, rule, anchor)
      if (
        (rule.operator === 'between' || rule.operator === 'notBetween') &&
        !Number.isFinite(secondNumber)
      ) {
        return null
      }
      if (
        rule.operator !== undefined &&
        coveredCells > 0 &&
        coveredCells <= CELLIS_FORMULA_CELL_LIMIT &&
        cellIsBlankDiverges(rule.operator, firstNumber, secondNumber)
      ) {
        return buildCellIsFormula(builder, rule.operator, anchor, firstNumber, secondNumber)
      }
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
    case 'notContainsText':
    case 'beginsWith':
    case 'endsWith':
      return buildTextCondition(builder, rule, anchor, coveredCells)
    case 'containsBlanks':
      return builder.whenCellEmpty()
    case 'notContainsBlanks':
      return builder.whenCellNotEmpty()
    case 'containsErrors':
    case 'notContainsErrors':
      // Univer's calculate unit evaluates these operators but the facade
      // builder has no setter; build a placeholder text condition and
      // re-target the operator on the built rule (patchBuiltHighlightRule).
      return builder.whenTextContains('')
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

/// Excel's text operators compare case-insensitively (containsText is
/// SEARCH-based), but Univer's evaluator is a raw indexOf — "SHORT" never
/// matches a containsText "Short" rule. Install the rule as a formula (the
/// file's own SEARCH/LEFT/RIGHT formula when present) and keep the cheaper
/// native condition for huge ranges, where the case gap is the lesser evil.
function buildTextCondition(
  builder: ReturnType<UniverWorksheet['newConditionalFormattingRule']>,
  rule: WorkbookRangeResult['conditionalRules'][number],
  anchor: string,
  coveredCells: number,
): CfHighlightBuilder | null {
  const text = rule.text
  if (!text) return null
  if (coveredCells > 0 && coveredCells <= CELLIS_FORMULA_CELL_LIMIT) {
    // Always synthesized: the file's own formula is anchored to its first
    // sqref cell, but Univer offsets from the top-left-sorted first range.
    const formula = synthesizeTextConditionFormula(rule.ruleType, text, anchor)
    if (formula) return builder.whenFormulaSatisfied(toCfFormula(formula))
  }
  switch (rule.ruleType) {
    case 'containsText':
      return builder.whenTextContains(text)
    case 'notContainsText':
      return builder.whenTextDoesNotContain(text)
    case 'beginsWith':
      return builder.whenTextStartsWith(text)
    case 'endsWith':
      return builder.whenTextEndsWith(text)
    default:
      return null
  }
}

function synthesizeTextConditionFormula(
  ruleType: string,
  text: string,
  anchor: string,
): string | null {
  const quoted = `"${text.replace(/"/g, '""')}"`
  switch (ruleType) {
    case 'containsText':
      return `NOT(ISERROR(SEARCH(${quoted},${anchor})))`
    case 'notContainsText':
      return `ISERROR(SEARCH(${quoted},${anchor}))`
    case 'beginsWith':
      return `LEFT(${anchor},LEN(${quoted}))=${quoted}`
    case 'endsWith':
      return `RIGHT(${anchor},LEN(${quoted}))=${quoted}`
    default:
      return null
  }
}

/// cellIs with a non-numeric operand: a quoted string (Excel compares text —
/// equality via the text builders, ordering via a formula rule) or a cell
/// reference / expression (always a formula rule). `anchor` is the relative
/// top-left of the rule's first range, the cell Excel evaluates against.
function buildCellIsNonNumeric(
  builder: ReturnType<UniverWorksheet['newConditionalFormattingRule']>,
  rule: WorkbookRangeResult['conditionalRules'][number],
  anchor: string,
): CfHighlightBuilder | null {
  const first = rule.formulas[0]
  const second = rule.formulas[1]
  if (!first) return null
  const quoted = /^"([\s\S]*)"$/.exec(first)
  switch (rule.operator) {
    case 'equal':
      return quoted
        ? builder.whenTextEqualTo(quoted[1]!.replace(/""/g, '"'))
        : builder.whenFormulaSatisfied(`=${anchor}=(${first})`)
    case 'notEqual':
      return builder.whenFormulaSatisfied(`=${anchor}<>${wrapOperand(first)}`)
    case 'greaterThan':
      return builder.whenFormulaSatisfied(`=${anchor}>${wrapOperand(first)}`)
    case 'greaterThanOrEqual':
      return builder.whenFormulaSatisfied(`=${anchor}>=${wrapOperand(first)}`)
    case 'lessThan':
      return builder.whenFormulaSatisfied(`=${anchor}<${wrapOperand(first)}`)
    case 'lessThanOrEqual':
      return builder.whenFormulaSatisfied(`=${anchor}<=${wrapOperand(first)}`)
    case 'between':
      return second
        ? builder.whenFormulaSatisfied(
            `=AND(${anchor}>=${wrapOperand(first)},${anchor}<=${wrapOperand(second)})`,
          )
        : null
    case 'notBetween':
      return second
        ? builder.whenFormulaSatisfied(
            `=NOT(AND(${anchor}>=${wrapOperand(first)},${anchor}<=${wrapOperand(second)}))`,
          )
        : null
    default:
      return null
  }
}

/// Quoted strings must stay verbatim; anything else gets parenthesized so
/// composite expressions keep their precedence inside the comparison.
function wrapOperand(operand: string): string {
  return /^"[\s\S]*"$/.test(operand) ? operand : `(${operand})`
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

export function toUniverStyle(style: WorkbookCellStyle): IStyleData {
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
    ...(style.fontFamily ? { ff: escapeCssLeadingDigit(style.fontFamily) } : {}),
    ...(style.fontSize ? { fs: style.fontSize } : {}),
    bl: style.bold ? BooleanNumber.TRUE : BooleanNumber.FALSE,
    it: style.italic ? BooleanNumber.TRUE : BooleanNumber.FALSE,
    ...(style.underline ? { ul: { s: BooleanNumber.TRUE } } : {}),
    ...(style.strikethrough ? { st: { s: BooleanNumber.TRUE } } : {}),
    // xf alignment is fully resolved in the file: an explicit non-wrap cell
    // must override a WRAP column/row style at compose time.
    // centerContinuous rides on the overflow machinery (center-continuous.ts),
    // so it must not wrap: Excel renders a wrapText+centerContinuous title on
    // one line across the run, not folded inside the anchor cell (prod ref).
    tb:
      style.wrapText && style.horizontalAlignment !== 'centerContinuous'
        ? WrapStrategy.WRAP
        : WrapStrategy.OVERFLOW,
    ...(style.fontColor ? { cl: { rgb: style.fontColor } } : {}),
    // Like tb above: an explicit xf without a fill must BLOCK a filled
    // column/row style at compose time — Univer merges styles by key, so a
    // missing bg lets a <col style=> fill bleed through explicitly-styled
    // cells (Excel treats each xf as complete, never a merge). bg: null
    // would be stripped by SetRangeValues' Tools.removeNull, so use an
    // empty-rgb sentinel: defined (blocks the compose fallthrough) but
    // falsy for every painter that checks bg.rgb.
    ...(style.fillColor ? { bg: { rgb: style.fillColor } } : { bg: { rgb: '' } }),
    ...(style.numberFormat ? { n: { pattern: style.numberFormat } } : {}),
    ...(Object.keys(borders).length > 0 ? { bd: borders } : {}),
    ...(mapHorizontalAlignment(style.horizontalAlignment) === undefined
      ? {}
      : { ht: mapHorizontalAlignment(style.horizontalAlignment) }),
    ...(mapVerticalAlignment(style.verticalAlignment) === undefined
      ? {}
      : { vt: mapVerticalAlignment(style.verticalAlignment) }),
    ...(style.indent ? { pd: { l: style.indent * INDENT_STEP_PX } } : {}),
    ...(style.textRotation
      ? { tr: ooxmlTextRotationToUniver(style.textRotation) ?? undefined }
      : {}),
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

export function mapBorderStyle(style: string): BorderStyleTypes {
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
  // The center-across-selection box is widened by center-continuous.ts.
  if (value === 'centerContinuous') return HorizontalAlign.CENTER
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

/// Magic-byte check for downloaded images: the ai:fetch-image handler labels
/// bytes from the Content-Type header (JPEG fallback), so a WebP or other
/// unsupported payload could otherwise land in the xlsx as a mislabeled media
/// part that Excel cannot display.
export function sniffImageMime(base64: string): 'image/png' | 'image/jpeg' | 'image/gif' | null {
  let bytes: Uint8Array
  try {
    bytes = Uint8Array.from(atob(base64.slice(0, 16)), (char) => char.charCodeAt(0))
  } catch {
    return null
  }
  if (bytes[0] === 0x89 && bytes[1] === 0x50 && bytes[2] === 0x4e && bytes[3] === 0x47) {
    return 'image/png'
  }
  if (bytes[0] === 0xff && bytes[1] === 0xd8 && bytes[2] === 0xff) return 'image/jpeg'
  if (bytes[0] === 0x47 && bytes[1] === 0x49 && bytes[2] === 0x46 && bytes[3] === 0x38) {
    return 'image/gif'
  }
  return null
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

/// The sheet the float installs should target, resolved at timer-fire time —
/// and from the RENDERED sheet (the skeleton the canvas shows), not the
/// workbook model. A tab click during load can leave the model pointing at
/// the file's stored activeTab while the canvas already shows the clicked
/// sheet; trusting the model then installs (and Univer's float-DOM service,
/// which checks the model, keeps painting) the wrong sheet's floats over the
/// visible grid. Re-activating the rendered sheet heals the model so the
/// float service accepts the install; visually it is already there.
export function resolveRenderedSheetId(runtime: UniverRuntime): string | undefined {
  const workbook = runtime.univerAPI.getActiveWorkbook()
  if (!workbook) return undefined
  const renderSheetId = (() => {
    try {
      return runtime.univer
        .__getInjector()
        .get(IRenderManagerService)
        .getRenderById(workbook.getId())
        ?.with(SheetSkeletonManagerService)
        .getCurrentParam()?.sheetId
    } catch {
      // Render modules not registered yet (redi throws): model fallback.
      return undefined
    }
  })()
  const sheetId = renderSheetId ?? workbook.getActiveSheet()?.getSheetId()
  if (!sheetId) return undefined
  if (workbook.getActiveSheet()?.getSheetId() !== sheetId) {
    const rendered = workbook.getSheetBySheetId(sheetId)
    if (rendered) workbook.setActiveSheet(rendered)
  }
  return sheetId
}

export function queueVisualInstall(
  runtime: UniverRuntime,
  lazyWorkbookRef: { current: LazyWorkbookState | null },
  visualDisposablesRef: { current: { dispose(): void }[] },
  visualInstallTimerRef: { current: ReturnType<typeof setTimeout> | null },
  chartEditRef?: { current: (chartPath: string, edit: ChartEditData) => void },
  chartVectorRef?: { current: (chartPath: string, range: string) => Promise<ChartVectorRead> },
  shapeEditRef?: { current: (visualId: string, changes: ShapeEditChanges) => void },
): void {
  const state = lazyWorkbookRef.current
  if (!state) return
  if (visualInstallTimerRef.current) clearTimeout(visualInstallTimerRef.current)
  visualInstallTimerRef.current = setTimeout(function install() {
    visualInstallTimerRef.current = null
    if (lazyWorkbookRef.current !== state) return
    const workbook = runtime.univerAPI.getActiveWorkbook()
    const workbookId = workbook?.getId()
    const render = workbookId
      ? runtime.univer.__getInjector().get(IRenderManagerService).getRenderById(workbookId)
      : null
    const renderMounted = (() => {
      try {
        return Boolean(render?.mainComponent && render.engine.getCanvasElement().isConnected)
      } catch {
        return false
      }
    })()
    if (!renderMounted) {
      visualInstallTimerRef.current = setTimeout(install, 100)
      return
    }
    const sheetId = resolveRenderedSheetId(runtime)
    if (!sheetId) return
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
              const edit = visualEdits.get(visual.id)
              if (!edit?.anchor && !edit?.frameSize) return visual
              return {
                ...visual,
                ...(edit.anchor ? { anchor: edit.anchor } : {}),
                // A resized rotated shape carries its new true frame; the
                // install's AABB re-anchoring must not rebuild the old one.
                ...(edit.frameSize
                  ? { frameWidth: edit.frameSize.width, frameHeight: edit.frameSize.height }
                  : {}),
              }
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
): void {
  const state = lazyWorkbookRef.current
  if (!state) return
  if (sparklineTimerRef.current) clearTimeout(sparklineTimerRef.current)
  sparklineTimerRef.current = setTimeout(() => {
    sparklineTimerRef.current = null
    if (lazyWorkbookRef.current !== state) return
    // Fire-time rendered sheet, for the same stale-float reason as
    // queueVisualInstall above.
    const sheetId = resolveRenderedSheetId(runtime)
    if (!sheetId) return
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
    // In-cell rich-value pictures share the per-cell float DOM lifecycle.
    const cellImages = sheetMeta?.cellImages ?? []
    sparklineDisposablesRef.current = [
      ...(sparklines.length === 0 ? [] : installSparklines(runtime, sparklines, sheetId)),
      ...(cellImages.length === 0
        ? []
        : installCellImages(runtime, state.file.sessionId, cellImages, sheetId)),
    ]
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
    // getValue() reads the view model, where numfmt interceptors have
    // replaced dates/formatted numbers with display strings; keep the raw
    // model value alongside for type-sensitive consumers (find_replace).
    // Rich-text cells have no plain v (and no display v either — rendering
    // reads the p document directly): take the dataStream, minus the
    // document-terminating \r\n that would break wholeCell matches.
    const rawCell = range.getCellDatas()[0]?.[0]
    const richStream = rawCell?.p?.body?.dataStream
    // paragraph breaks (\r) become \n, matching extractRichText and typed text
    const richText =
      typeof richStream === 'string' ? richStream.replace(/\r\n$/, '').replace(/\r/g, '\n') : null
    const rawValue = (rawCell?.v ?? richText) as CellState['rawValue']
    // Formula cells also carry their computed value (the AI needs to see results
    // and error values like #REF!/#DIV/0!; drift checks compare only formula
    // text for formula cells, see planStillMatches)
    if (formula) return { value, formula, rawValue }
    return { value, rawValue }
  }
}

/// Sheet-aware variant: operations carry their own sheetId, which may differ
/// from the active sheet — resolve (and cache) the target worksheet per read.
export function lazyWorkbookCellReader(
  workbook: ActiveWorkbook,
): (address: string, sheetId: string) => CellState {
  const readers = new Map<string, (address: string) => CellState>()
  return (address, sheetId) => {
    let reader = readers.get(sheetId)
    if (!reader) {
      const worksheet = workbook.getSheetBySheetId(sheetId)
      if (!worksheet) throw new Error(`Unknown sheet: ${sheetId}`)
      reader = lazyCellReader(worksheet)
      readers.set(sheetId, reader)
    }
    return reader(address)
  }
}

/// Range-level variant of lazyCellEditable for bulk ops (fill_range / large
/// clear_range): checking the four corners is not enough because a range can
/// straddle the loaded region and the beyond-extent area with unloaded rows
/// in between — clamp to the file extent first, then require containment.
export function lazyRangeEditable(
  state: LazyWorkbookState,
  sheetId: string,
  bounds: { startRow: number; endRow: number; startColumn: number; endColumn: number },
): boolean {
  if (state.flags.preloadComplete) return true
  const extent = lazySheetScreenExtent(state, sheetId)
  if (!extent) return true
  const inExtentEndRow = Math.min(bounds.endRow, extent.rows - 1)
  const inExtentEndColumn = Math.min(bounds.endColumn, extent.columns - 1)
  // Entirely beyond the data extent: nothing left to stream in.
  if (inExtentEndRow < bounds.startRow || inExtentEndColumn < bounds.startColumn) return true
  // Rows/columns inserted this session are journal-owned — nothing streams
  // into them, so only the file-backed remainder needs the loaded window.
  // (A range fully inside an inserted column, like a fill source written
  // this session, is editable regardless of where the window sits.)
  const ops = state.editJournal.structuralOps.get(sheetId) ?? []
  const fileBackedSpan = (
    axis: 'row' | 'column',
    start: number,
    end: number,
  ): { start: number; end: number } | null => {
    if (ops.length === 0) return { start, end }
    let first = -1
    let last = -1
    for (let position = start; position <= end; position += 1) {
      if (screenToFile(ops, axis, position) === null) continue
      if (first === -1) first = position
      last = position
    }
    return first === -1 ? null : { start: first, end: last }
  }
  const rows = fileBackedSpan('row', bounds.startRow, inExtentEndRow)
  const columns = fileBackedSpan('column', bounds.startColumn, inExtentEndColumn)
  if (rows === null || columns === null) return true
  const loaded = state.loadedRanges.get(sheetId)
  return (
    loaded !== undefined &&
    rows.start >= loaded.startRow &&
    rows.end <= loaded.endRow &&
    columns.start >= loaded.startColumn &&
    columns.end <= loaded.endColumn
  )
}

/// Mirrors the BeforeSheetEditStart streaming guard for AI-planned cells.
export function lazyCellEditable(
  state: LazyWorkbookState,
  sheetId: string,
  row: number,
  column: number,
): boolean {
  if (state.flags.preloadComplete) return true
  const extent = lazySheetScreenExtent(state, sheetId)
  if (!extent) return true
  if (row >= extent.rows || column >= extent.columns) return true
  const ops = state.editJournal.structuralOps.get(sheetId) ?? []
  const journalOwned =
    ops.length > 0 &&
    (screenToFile(ops, 'row', row) === null || screenToFile(ops, 'column', column) === null)
  if (journalOwned) return true
  const loaded = state.loadedRanges.get(sheetId)
  return (
    loaded !== undefined &&
    row >= loaded.startRow &&
    row <= loaded.endRow &&
    column >= loaded.startColumn &&
    column <= loaded.endColumn
  )
}
