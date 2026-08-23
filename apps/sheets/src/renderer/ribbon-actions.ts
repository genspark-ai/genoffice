/**
 * The ribbon command dispatcher. Extracted from App.tsx; the App component
 * passes a RibbonCommandContext built fresh per call so refs and state never
 * go stale. Page-layout, export and save flows stay in App and arrive here
 * as callbacks.
 */
import {
  BooleanNumber,
  BorderStyleTypes,
  WrapStrategy,
  type ICellData,
  type IStyleData,
} from '@univerjs/core'
import { IRenderManagerService, SHEET_VIEWPORT_KEY } from '@univerjs/engine-render'
import { SheetSkeletonManagerService } from '@univerjs/preset-sheets-core'
import { columnLabel, formatAddress } from '../domain/cell-address'
import { transposeChartSeries, type ChartSeriesVisualState } from '../domain/chart-visual'
import { applyFlashFillTemplate, inferFlashFillTemplate } from '../domain/flash-fill'
import type {
  WorkbookChartEdit,
  WorkbookStyleEdit,
  WorkbookVisualObject,
} from '../shared/desktop-api'
import {
  CELL_STYLE_PRESETS,
  CHART_LABEL_COMMANDS,
  CHART_LEGEND_COMMANDS,
  CHART_PALETTES,
  CHART_TYPE_COMMANDS,
} from './app-constants'
import {
  handleApplyFormula,
  handleCreateNamesFromSelection,
  handleFormatAsTable,
  handleImportCsv,
  handleOutline,
  openAdvancedFilterDialog,
  type DataToolsContext,
} from './data-tools-actions'
import { dedupeRows } from './dedupe'
import {
  isSheetRemoved,
  journalSize,
  recordHyperlinkEdit,
  recordNeutralStyleEdit,
  recordPageSetup,
  recordSheetProtection,
  recordSparklineAdd,
  recordWorkbookProtection,
  removeSparklineAdd,
} from './edit-journal'
import { applyShowFormulasView, formulaViewSheets } from './formula-view'
import { t } from './i18n/locale'
import {
  handleOpenSlicerPicker,
  handleOpenTimelinePicker,
  handleRefreshAllPivots,
  type PivotActionContext,
} from './pivot-actions'
import { INDENT_STEP_PX } from './selection-format'
import { collectDependents, collectPrecedents, installTraceArrows } from './trace-arrows'
import {
  absRangeRef,
  applyFormatPatchToRange,
  characterWidthToPixels,
  normalizeLinkTarget,
  pushVisualUndo,
  queueSparklineInstall,
  workbookStructureLocked,
} from './univer-sync'
import {
  BORDER_COMMAND_TYPES,
  type ActiveWorkbook,
  type LazyWorkbookState,
  type UniverRuntime,
} from './univer-state'
import {
  handleInsertChart,
  handleInsertPicture,
  handleInsertPivotChart,
  handleInsertShape,
  startShapeDraw,
  type VisualActionContext,
} from './visual-actions'
import type { ChartDialogKind, ChartEditData, ShapeEditChanges } from './WorkbookVisuals'

/** The App refs/state the ribbon dispatcher needs; built fresh per call. */
export interface RibbonCommandContext {
  univerRef: { readonly current: UniverRuntime | null }
  lazyWorkbookRef: { current: LazyWorkbookState | null }
  traceArrowsRef: { current: { disposables: { dispose(): void }[]; nextId: number } }
  sparklineDisposablesRef: { current: { dispose(): void }[] }
  sparklineTimerRef: { current: ReturnType<typeof setTimeout> | null }
  chartEditRef: { readonly current: (chartPath: string, edit: ChartEditData) => void }
  shapeEditRef: { readonly current: (visualId: string, changes: ShapeEditChanges) => void }
  refreshSelectionFormatRef: { readonly current: () => void }
  selectedVisual: WorkbookVisualObject | null
  selectedChart: {
    readonly isPie: boolean
    readonly canLabel: boolean
    readonly seriesCount: number
    readonly categoryCount: number
    /// Current series (pending edits applied) for Switch Row/Column.
    readonly series: readonly ChartSeriesVisualState[]
  } | null
  setMessage: (message: string) => void
  setChartDialog: (dialog: { kind: ChartDialogKind; editKey: string }) => void
  setSymbolDialogOpen: (open: boolean) => void
  setScreenshotDialogOpen: (open: boolean) => void
  setIconsDialogOpen: (open: boolean) => void
  setEquationDialogOpen: (open: boolean) => void
  openRecommendedCharts: () => void
  setPendingEdits: (count: number) => void
  visualContext: () => VisualActionContext
  dataToolsContext: () => DataToolsContext
  pivotContext: () => PivotActionContext
  handlePageLayoutCommand: (rest: string) => void
  handleExportPdf: () => Promise<void>
}

/// Resolves interned style references and merges row/col/sheet styles —
/// raw getCellData().s can be a style-id string with no fields on it.
function selectionStyle(
  range: NonNullable<ReturnType<ActiveWorkbook['getActiveRange']>>,
): IStyleData {
  return range.getCellStyleData() ?? {}
}

/// Commands whose wire format is `name:argument:extra` (three segments).
/// Every other command treats the whole remainder after the first colon as
/// the argument, because arguments like number-format patterns
/// ("h:mm:ss AM/PM") legitimately contain colons.
const EXTRA_SEGMENT_COMMANDS = new Set(['cellprot', 'sort-custom', 'border'])

/// Univer's built-in header sizes, restored when headings toggle back on.
const DEFAULT_ROW_HEADER_WIDTH = 46
const DEFAULT_COLUMN_HEADER_HEIGHT = 20

/// OOXML error literals (Formulas › Error Checking scans display values).
const ERROR_VALUE_PATTERN = /^#(DIV\/0!|N\/A|NAME\?|NULL!|NUM!|REF!|VALUE!|SPILL!|CALC!)/
/// Row-major ordinal for "next error after the active cell" (> max columns).
const COLUMN_LIMIT = 20_000

/// ST_BorderStyle names accepted by the Format Cells line-style picker.
const BORDER_LINE_STYLE_TYPES: Record<string, BorderStyleTypes> = {
  thin: BorderStyleTypes.THIN,
  medium: BorderStyleTypes.MEDIUM,
  thick: BorderStyleTypes.THICK,
  double: BorderStyleTypes.DOUBLE,
  hair: BorderStyleTypes.HAIR,
  dashed: BorderStyleTypes.DASHED,
  dotted: BorderStyleTypes.DOTTED,
}

/** Splits a `name:argument[:extra]` ribbon style command (see above). */
export function parseStyleCommand(command: string): {
  name: string
  argument: string
  extra: string
} {
  const nameEnd = command.indexOf(':')
  if (nameEnd === -1) return { name: command, argument: '', extra: '' }
  const name = command.slice(0, nameEnd)
  const rest = command.slice(nameEnd + 1)
  if (!EXTRA_SEGMENT_COMMANDS.has(name)) return { name, argument: rest, extra: '' }
  const argumentEnd = rest.indexOf(':')
  if (argumentEnd === -1) return { name, argument: rest, extra: '' }
  return { name, argument: rest.slice(0, argumentEnd), extra: rest.slice(argumentEnd + 1) }
}

export function handleRibbonCommand(ctx: RibbonCommandContext, command: string): void {
  const runtime = ctx.univerRef.current
  if (!runtime) return
  if (command === 'undo' || command === 'redo') {
    void runtime.univerAPI[command]()
    return
  }
  if (command.startsWith('error:')) {
    ctx.setMessage(command.slice('error:'.length))
    return
  }
  if (command === 'chart-select-data' || command === 'chart-format-pane') {
    const editKey = ctx.selectedVisual
      ? (ctx.selectedVisual.chartPath ?? ctx.selectedVisual.id)
      : undefined
    if (editKey) {
      ctx.setChartDialog({
        kind: command === 'chart-select-data' ? 'select-data' : 'format',
        editKey,
      })
    }
    return
  }
  // Chart Design tab commands act on the selected floating chart; file
  // charts key on their chart part path, in-memory ones on the visual id.
  if (command.startsWith('chart-') && command !== 'chart-delete') {
    const editKey = ctx.selectedVisual
      ? (ctx.selectedVisual.chartPath ?? ctx.selectedVisual.id)
      : undefined
    if (!editKey) return
    if (command.startsWith('chart-type-')) {
      const chartType = CHART_TYPE_COMMANDS.find((type) => command === `chart-type-${type}`)
      if (chartType) ctx.chartEditRef.current(editKey, { chartType })
    } else if (command.startsWith('chart-legend:')) {
      const legend = CHART_LEGEND_COMMANDS.find((pos) => command === `chart-legend:${pos}`)
      if (legend) ctx.chartEditRef.current(editKey, { legend })
    } else if (command.startsWith('chart-labels:')) {
      const labels = CHART_LABEL_COMMANDS.find((mode) => command === `chart-labels:${mode}`)
      // scatter/radar cannot save plot-level labels — fail here, not at ⌘S.
      if (labels && ctx.selectedChart?.canLabel) {
        ctx.chartEditRef.current(editKey, { dataLabels: labels })
      }
    } else if (command.startsWith('chart-grouping:')) {
      const grouping = (['clustered', 'stacked', 'percentStacked'] as const).find(
        (mode) => command === `chart-grouping:${mode}`,
      )
      if (grouping) ctx.chartEditRef.current(editKey, { grouping })
    } else if (command.startsWith('chart-colors:')) {
      const palette = CHART_PALETTES[command.slice('chart-colors:'.length)]
      if (!palette) return
      if (ctx.selectedChart?.isPie) {
        // Pie recolor varies by point, not by series.
        const count = Math.min(ctx.selectedChart.categoryCount, 64)
        if (count === 0) return
        ctx.chartEditRef.current(editKey, {
          pointColors: {
            '0': Object.fromEntries(
              Array.from({ length: count }, (_, index) => [
                String(index),
                palette[index % palette.length] ?? '#4472c4',
              ]),
            ),
          },
        })
        return
      }
      const count = Math.min(ctx.selectedChart?.seriesCount ?? 0, 24)
      if (count === 0) return
      ctx.chartEditRef.current(editKey, {
        seriesColors: Object.fromEntries(
          Array.from({ length: count }, (_, index) => [
            String(index),
            palette[index % palette.length] ?? '#4472c4',
          ]),
        ),
      })
    } else if (command.startsWith('chart-layout:')) {
      // Quick Layout presets: legend placement + data labels together,
      // labels picked per chart family (pie vs. axis charts).
      const labelsOn = ctx.selectedChart?.isPie ? ('category-percent' as const) : ('value' as const)
      const preset = {
        'chart-layout:1': { legend: 'right', dataLabels: labelsOn },
        'chart-layout:2': { legend: 'top', dataLabels: labelsOn },
        'chart-layout:3': { legend: 'bottom', dataLabels: 'none' },
        'chart-layout:4': { legend: 'none', dataLabels: labelsOn },
      }[command] as Pick<WorkbookChartEdit, 'legend' | 'dataLabels'> | undefined
      if (preset) {
        // Label-less families take the preset's legend only.
        const applied = ctx.selectedChart?.canLabel
          ? preset
          : { ...(preset.legend === undefined ? {} : { legend: preset.legend }) }
        ctx.chartEditRef.current(editKey, applied)
      }
    } else if (command.startsWith('chart-title:')) {
      ctx.chartEditRef.current(editKey, {
        title: command.slice('chart-title:'.length).slice(0, 255),
      })
    } else if (command.startsWith('chart-axis-cat:')) {
      const text = command.slice('chart-axis-cat:'.length).slice(0, 255)
      ctx.chartEditRef.current(editKey, { axisTitles: { category: text || null } })
    } else if (command.startsWith('chart-axis-val:')) {
      const text = command.slice('chart-axis-val:'.length).slice(0, 255)
      ctx.chartEditRef.current(editKey, { axisTitles: { value: text || null } })
    } else if (command === 'chart-switch-row-col') {
      const seriesSet = transposeChartSeries(ctx.selectedChart?.series ?? [], (n) =>
        t('appSeriesN', { n }),
      )
      if (seriesSet) ctx.chartEditRef.current(editKey, { seriesSet })
      else ctx.setMessage(t('appNoCategoriesToSwitch'))
    }
    return
  }
  if (command === 'chart-delete') {
    if (ctx.selectedVisual) ctx.shapeEditRef.current(ctx.selectedVisual.id, { remove: true })
    return
  }
  // Read-only-safe commands work on imported workbooks too.
  const worksheet = runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
  switch (command) {
    case 'find':
      void runtime.univerAPI.executeCommand('ui.operation.open-find-dialog')
      return
    case 'copy':
    case 'cut':
    case 'paste':
      // The sheet clipboard implementation registers on the shared UI
      // command ids; paste falls back to the internal copy cache when the
      // Clipboard API is unavailable.
      void runtime.univerAPI.executeCommand(`univer.command.${command}`)
      return
    case 'paste-special:value':
      // Excel's Paste Special: Univer ships each variant as its own sheet
      // command, so they undo exactly like a plain paste.
      void runtime.univerAPI.executeCommand('sheet.command.paste-value')
      return
    case 'paste-special:formula':
      void runtime.univerAPI.executeCommand('sheet.command.paste-formula')
      return
    case 'paste-special:format':
      void runtime.univerAPI.executeCommand('sheet.command.paste-format')
      return
    case 'paste-special:col-width':
      void runtime.univerAPI.executeCommand('sheet.command.paste-col-width')
      return
    case 'paste-special:besides-border':
      void runtime.univerAPI.executeCommand('sheet.command.paste-besides-border')
      return
    case 'insert-sheet': {
      const workbook = runtime.univerAPI.getActiveWorkbook()
      if (!workbook) return
      if (workbookStructureLocked(ctx.lazyWorkbookRef.current)) {
        ctx.setMessage(t('appWorkbookStructureLocked'))
        return
      }
      try {
        workbook.insertSheet()
        ctx.setMessage(t('appSheetAdded'))
      } catch (error: unknown) {
        ctx.setMessage(error instanceof Error ? error.message : t('appSheetAddFailed'))
      }
      return
    }
    case 'format-painter':
      // One-shot painter: the next selection receives the copied format.
      // Style deltas land as set-range-values mutations, so they journal
      // and save like any ribbon style edit.
      void runtime.univerAPI.executeCommand('sheet.command.set-once-format-painter')
      ctx.setMessage(t('appFormatCopied'))
      return
    case 'cf-open':
      // value 2 = the manage-rules list; other values preseed a new rule.
      void runtime.univerAPI.executeCommand('sheet.operation.open.conditional.formatting.panel', {
        value: 2,
      })
      return
    case 'dv-open':
      // Same trap as the CF panel: a missing params object silently no-ops.
      void runtime.univerAPI.executeCommand('data-validation.operation.open-validation-panel', {})
      return
    case 'sheet-protect': {
      const state = ctx.lazyWorkbookRef.current
      if (!state) {
        ctx.setMessage(t('appProtectionNeedsFile'))
        return
      }
      const sheetId = worksheet?.getSheetId()
      if (!sheetId || isSheetRemoved(state.editJournal, sheetId)) return
      const isAdded = state.editJournal.sheets.added.has(sheetId)
      const file = state.sheetProtections.get(sheetId)
      if (!file && !isAdded) {
        ctx.setMessage(t('appProtectionNeedsIndexed'))
        return
      }
      const original = file?.protected ?? false
      const current = state.editJournal.sheetProtection.get(sheetId) ?? original
      if (current && file?.hasPassword) {
        ctx.setMessage(t('appProtectedWithPassword'))
        return
      }
      recordSheetProtection(state.editJournal, sheetId, !current, original)
      ctx.setPendingEdits(journalSize(state.editJournal))
      ctx.setMessage(!current ? t('appProtectionWillWrite') : t('appProtectionWillRemove'))
      return
    }
    case 'workbook-protect': {
      const state = ctx.lazyWorkbookRef.current
      if (!state) {
        ctx.setMessage(t('appProtectionNeedsFile'))
        return
      }
      const file = state.file.workbookProtection
      const original = file?.lockStructure ?? false
      const current = state.editJournal.workbookProtection.desired ?? original
      // Locking would be as irreversible as unlocking is impossible: the
      // gateway refuses to touch a password-bearing element either way.
      if (file?.hasPassword) {
        ctx.setMessage(t('appWorkbookProtectedWithPassword'))
        return
      }
      recordWorkbookProtection(state.editJournal, !current, original)
      ctx.setPendingEdits(journalSize(state.editJournal))
      ctx.setMessage(
        !current ? t('appWorkbookProtectionWillWrite') : t('appWorkbookProtectionWillRemove'),
      )
      return
    }
    case 'outline-group:rows':
    case 'outline-group:cols':
    case 'outline-ungroup:rows':
    case 'outline-ungroup:cols':
    case 'outline-hide-detail:rows':
    case 'outline-hide-detail:cols':
    case 'outline-show-detail:rows':
    case 'outline-show-detail:cols': {
      const [action, axis] = command.slice('outline-'.length).split(':')
      handleOutline(
        ctx.dataToolsContext(),
        action as 'group' | 'ungroup' | 'hide-detail' | 'show-detail',
        axis as 'rows' | 'cols',
      )
      return
    }
    case 'fill-down':
    case 'fill-right':
      void runtime.univerAPI.executeCommand(
        command === 'fill-down' ? 'sheet.command.copy-down' : 'sheet.command.copy-right',
      )
      return
    case 'clear-contents':
      void runtime.univerAPI.executeCommand('sheet.command.clear-selection-content')
      return
    case 'clear-formats':
      void runtime.univerAPI.executeCommand('sheet.command.clear-selection-format')
      return
    case 'clear-all':
      void runtime.univerAPI.executeCommand('sheet.command.clear-selection-all')
      return
    case 'decimal-inc':
    case 'decimal-dec':
      // Univer's decimal commands read the displayed value's decimals for
      // General; SetNumfmt mutations journal like any format edit.
      void runtime.univerAPI.executeCommand(
        command === 'decimal-inc'
          ? 'sheet.command.numfmt.add.decimal.command'
          : 'sheet.command.numfmt.subtract.decimal.command',
      )
      return
    case 'filter-toggle':
      void runtime.univerAPI.executeCommand('sheet.command.smart-toggle-filter')
      return
    case 'refresh-all': {
      const error = handleRefreshAllPivots(ctx.pivotContext())
      if (error) ctx.setMessage(error)
      return
    }
    case 'error-checking': {
      const workbook = runtime.univerAPI.getActiveWorkbook()
      if (!workbook || !worksheet) return
      // A streamed workbook's cell matrix holds only loaded regions. Formula
      // -mode files finish preloading, so wait for that (same gate as
      // Replace); above-cap files never will — scan what is loaded and say so
      // in the result instead of under-reporting silently.
      const lazyState = ctx.lazyWorkbookRef.current
      const partialScan = Boolean(lazyState && !lazyState.flags.preloadComplete)
      if (partialScan && lazyState?.formulaMode) {
        ctx.setMessage(t('appErrorCheckNeedsFullLoad'))
        return
      }
      const errors: { row: number; column: number; value: string }[] = []
      worksheet
        .getSheet()
        .getCellMatrix()
        .forValue((row, column, cell) => {
          const value = cell?.v
          if (typeof value === 'string' && ERROR_VALUE_PATTERN.test(value)) {
            errors.push({ row, column, value })
          }
          return undefined
        })
      if (errors.length === 0) {
        ctx.setMessage(t(partialScan ? 'appNoErrorsFoundPartial' : 'appNoErrorsFound'))
        return
      }
      // Step to the first error after the active cell, wrapping around, so
      // repeated clicks cycle through all of them.
      const active = workbook.getActiveRange()
      const afterActive = active ? active.getRow() * COLUMN_LIMIT + active.getColumn() : -1
      const next =
        errors.find((error) => error.row * COLUMN_LIMIT + error.column > afterActive) ?? errors[0]
      if (!next) return
      worksheet.getRange(next.row, next.column, 1, 1).activate()
      // Programmatic selection emits no SelectionChanged; refresh the echo.
      ctx.refreshSelectionFormatRef.current()
      void runtime.univerAPI.executeCommand('sheet.command.scroll-to-cell', {
        range: {
          startRow: next.row,
          endRow: next.row,
          startColumn: next.column,
          endColumn: next.column,
        },
      })
      ctx.setMessage(
        t(partialScan ? 'appErrorsFoundPartial' : 'appErrorsFound', {
          count: errors.length,
          cell: formatAddress(next.row, next.column),
          value: next.value,
        }),
      )
      return
    }
    case 'filter-reapply':
      void runtime.univerAPI.executeCommand('sheet.command.re-calc-filter')
      return
    case 'filter-clear':
      void runtime.univerAPI.executeCommand('sheet.command.clear-filter-criteria')
      return
    case 'filter-advanced':
      if (!worksheet) return
      if (!worksheet.getFilter()) {
        // No auto-filter yet: create one over the smart selection first
        // (the filter-toggle equivalent), then open on the fresh range.
        void runtime.univerAPI
          .executeCommand('sheet.command.smart-toggle-filter')
          .then(() => openAdvancedFilterDialog(ctx.dataToolsContext()))
        return
      }
      openAdvancedFilterDialog(ctx.dataToolsContext())
      return
    case 'insert-symbol':
      ctx.setSymbolDialogOpen(true)
      return
    case 'insert-screenshot':
      ctx.setScreenshotDialogOpen(true)
      return
    case 'recommended-charts-open':
      ctx.openRecommendedCharts()
      return
    case 'insert-icons':
      ctx.setIconsDialogOpen(true)
      return
    case 'insert-equation':
      ctx.setEquationDialogOpen(true)
      return
    case 'import-csv':
      handleImportCsv(ctx.dataToolsContext())
      return
    case 'insert-checkbox': {
      // The DV command gate in App cancels this (with its own message) while
      // the sheet's file rules are still streaming in.
      const active = runtime.univerAPI.getActiveWorkbook()?.getActiveRange()
      if (!active) {
        ctx.setMessage(t('appSelectCellFirst'))
        return
      }
      active.setDataValidation(runtime.univerAPI.newDataValidation().requireCheckbox().build())
      return
    }
    case 'note-open':
      // Opens the note editor popup at the primary selected cell; the
      // journal snapshots the sheet's notes and ⌘S writes legacy comments.
      void runtime.univerAPI.executeCommand('sheet.operation.add-note-popup')
      return
    case 'note-delete':
      void runtime.univerAPI.executeCommand('sheet.command.delete-note')
      return
    case 'note-prev':
    case 'note-next': {
      const workbook = runtime.univerAPI.getActiveWorkbook()
      const sheet = workbook?.getActiveSheet()
      const active = workbook?.getActiveRange()
      if (!sheet || !active) return
      const notes = [...sheet.getNotes()].sort((a, b) => a.row - b.row || a.col - b.col)
      if (notes.length === 0) {
        ctx.setMessage(t('appNoNotesOnSheet'))
        return
      }
      const row = active.getRow()
      const column = active.getColumn()
      // Reading order with wrap-around, like Excel's Previous/Next Comment.
      const target =
        command === 'note-next'
          ? (notes.find((note) => note.row > row || (note.row === row && note.col > column)) ??
            notes[0])
          : ([...notes]
              .reverse()
              .find((note) => note.row < row || (note.row === row && note.col < column)) ??
            notes[notes.length - 1])
      if (!target) return
      sheet.getRange(target.row, target.col, 1, 1).activate()
      sheet.scrollToCell(target.row, target.col)
      void runtime.univerAPI.executeCommand('sheet.operation.add-note-popup')
      return
    }
    case 'note-show-toggle':
      // Pins/unpins the popup of the note at the selection (no-op elsewhere).
      void runtime.univerAPI.executeCommand('sheet.command.toggle-note-popup')
      return
    case 'zoom-in':
    case 'zoom-out':
    case 'zoom-reset': {
      if (!worksheet) return
      const next =
        command === 'zoom-reset'
          ? 1
          : Math.min(4, Math.max(0.5, worksheet.getZoom() + (command === 'zoom-in' ? 0.1 : -0.1)))
      worksheet.zoom(Number(next.toFixed(2)))
      ctx.setMessage(t('appZoom', { percent: Math.round(next * 100) }))
      return
    }
    case 'zoom-to-selection': {
      const workbook = runtime.univerAPI.getActiveWorkbook()
      const selection = workbook?.getActiveRange()?.getRange()
      if (!workbook || !worksheet || !selection) {
        ctx.setMessage(t('appSelectCellFirst'))
        return
      }
      const render = runtime.univer
        .__getInjector()
        .get(IRenderManagerService)
        .getRenderById(workbook.getId())
      const skeleton = render?.with(SheetSkeletonManagerService).getCurrentSkeleton()
      if (!render || !skeleton) return
      const rows = skeleton.rowHeightAccumulation
      const columns = skeleton.columnWidthAccumulation
      // Accumulation arrays are unscaled; entry i is the bottom/right edge of
      // line i, so the selection's extent is end-edge minus start-1-edge.
      const top = selection.startRow > 0 ? (rows[selection.startRow - 1] ?? 0) : 0
      const bottom = rows[Math.min(selection.endRow, rows.length - 1)] ?? 0
      const left = selection.startColumn > 0 ? (columns[selection.startColumn - 1] ?? 0) : 0
      const right = columns[Math.min(selection.endColumn, columns.length - 1)] ?? 0
      const viewWidth = render.engine.width - skeleton.rowHeaderWidthAndMarginLeft
      const viewHeight = render.engine.height - skeleton.columnHeaderHeightAndMarginTop
      if (right <= left || bottom <= top || viewWidth <= 0 || viewHeight <= 0) return
      const ratio = Math.min(
        4,
        Math.max(0.5, Math.min(viewWidth / (right - left), viewHeight / (bottom - top))),
      )
      // Scroll only after the zoom lands: the scroll target clamps against
      // the viewport extents, and at the old ratio a selection near the grid
      // edge clamps short and ends up off-screen.
      void runtime.univerAPI
        .executeCommand('sheet.command.set-zoom-ratio', {
          unitId: workbook.getId(),
          subUnitId: worksheet.getSheetId(),
          zoomRatio: Number(ratio.toFixed(2)),
        })
        .then(() =>
          runtime.univerAPI.executeCommand('sheet.command.scroll-to-cell', {
            range: selection,
            forceTop: true,
            forceLeft: true,
          }),
        )
      ctx.setMessage(t('appZoom', { percent: Math.round(ratio * 100) }))
      return
    }
    case 'toggle-headings': {
      const workbook = runtime.univerAPI.getActiveWorkbook()
      if (!workbook || !worksheet) return
      const sheetId = worksheet.getSheetId()
      const config = worksheet.getSheet().getConfig()
      const nextHidden = config.rowHeader.hidden !== BooleanNumber.TRUE
      // The skeleton reads these on every rebuild (sheet switches, reloads);
      // the two size commands below repaint the current view.
      config.rowHeader.hidden = nextHidden ? BooleanNumber.TRUE : BooleanNumber.FALSE
      config.columnHeader.hidden = nextHidden ? BooleanNumber.TRUE : BooleanNumber.FALSE
      const unitId = workbook.getId()
      // Univer's set-row-header-width infers the horizontal shift from the
      // corner viewport's width with a `|| 46` fallback; after hiding, that
      // reads 46 either way (stale value, or the fallback swallowing 0), so
      // re-showing computes a zero shift and leaves the grid under the
      // row-header strip. Record the expected viewport lefts and correct
      // after the commands — a no-op whenever Univer shifts correctly.
      const render = runtime.univer.__getInjector().get(IRenderManagerService).getRenderById(unitId)
      const scene = render?.scene
      const skeleton = render?.with(SheetSkeletonManagerService).getCurrentSkeleton()
      const viewMain = scene?.getViewport(SHEET_VIEWPORT_KEY.VIEW_MAIN)
      const viewColumnRight = scene?.getViewport(SHEET_VIEWPORT_KEY.VIEW_COLUMN_RIGHT)
      const nextWidth = nextHidden ? 0 : DEFAULT_ROW_HEADER_WIDTH
      const shift = skeleton ? nextWidth - skeleton.rowHeaderWidth : 0
      const mainLeft = (viewMain?.left ?? 0) + shift
      const columnRightLeft = (viewColumnRight?.left ?? 0) + shift
      void Promise.all([
        runtime.univerAPI.executeCommand('sheet.command.set-row-header-width', {
          unitId,
          subUnitId: sheetId,
          size: nextWidth,
        }),
        runtime.univerAPI.executeCommand('sheet.command.set-col-header-height', {
          unitId,
          subUnitId: sheetId,
          size: nextHidden ? 0 : DEFAULT_COLUMN_HEADER_HEIGHT,
        }),
      ]).then(() => {
        if (!viewMain || viewMain.left === mainLeft) return
        viewMain.left = mainLeft
        viewColumnRight?.setViewportSize({ left: columnRightLeft })
        scene?.makeDirty(true)
      })
      const state = ctx.lazyWorkbookRef.current
      if (state && !isSheetRemoved(state.editJournal, sheetId)) {
        recordPageSetup(state.editJournal, sheetId, { showHeadings: !nextHidden })
        ctx.setPendingEdits(journalSize(state.editJournal))
      }
      ctx.setMessage(t(nextHidden ? 'appHeadingsHidden' : 'appHeadingsShown'))
      return
    }
    case 'freeze-top-row':
      // Freeze/gridline changes journal from their mutations (App listener),
      // so Univer's undo/redo keeps the journal in step.
      if (worksheet) {
        worksheet.setFreeze({ startRow: 1, startColumn: -1, xSplit: 0, ySplit: 1 })
        ctx.setMessage(t('appTopRowFrozen'))
      }
      return
    case 'freeze-first-col':
      if (worksheet) {
        worksheet.setFreeze({ startRow: -1, startColumn: 1, xSplit: 1, ySplit: 0 })
        ctx.setMessage(t('appFirstColFrozen'))
      }
      return
    case 'replace':
      // Replacing over a partially streamed workbook would silently skip
      // unloaded regions; fully-loaded ones are safe.
      if (ctx.lazyWorkbookRef.current && !ctx.lazyWorkbookRef.current.flags.preloadComplete) {
        ctx.setMessage(t('appReplaceNeedsFullLoad'))
        return
      }
      void runtime.univerAPI.executeCommand('ui.operation.open-replace-dialog')
      return
    case 'toggle-gridlines': {
      if (!worksheet) return
      const nextHidden = !worksheet.hasHiddenGridLines()
      worksheet.setHiddenGridlines(nextHidden)
      // The display toggle persists as sheetView@showGridLines on save.
      // The message also forces the render that refreshes the ribbon
      // checkbox (journalSize may not change here).
      const state = ctx.lazyWorkbookRef.current
      const sheetId = worksheet.getSheetId()
      if (state && !isSheetRemoved(state.editJournal, sheetId)) {
        ctx.setMessage(nextHidden ? t('appGridlinesHiddenSave') : t('appGridlinesShownSave'))
      } else {
        ctx.setMessage(nextHidden ? t('appGridlinesHidden') : t('appGridlinesShown'))
      }
      return
    }
    case 'toggle-show-formulas': {
      // Formula view is per-sheet state (sheetView/@showFormulas): remember it
      // for sheet switches and persist it on save.
      const sheetId = worksheet?.getSheetId()
      if (!sheetId) return
      const state = ctx.lazyWorkbookRef.current
      const sheets = formulaViewSheets(state)
      const next = !sheets.has(sheetId)
      if (next) sheets.add(sheetId)
      else sheets.delete(sheetId)
      applyShowFormulasView(runtime, state, sheetId)
      if (state && !isSheetRemoved(state.editJournal, sheetId)) {
        recordPageSetup(state.editJournal, sheetId, { showFormulas: next })
        ctx.setPendingEdits(journalSize(state.editJournal))
      }
      ctx.setMessage(next ? t('appShowingFormulas') : t('appShowingValues'))
      return
    }
    case 'trace-precedents':
    case 'trace-dependents': {
      const workbook = runtime.univerAPI.getActiveWorkbook()
      const active = workbook?.getActiveRange()
      if (!workbook || !worksheet || !active) {
        ctx.setMessage(t('appSelectCellFirst'))
        return
      }
      const row = active.getRow()
      const column = active.getColumn()
      const idPrefix = `trace-${ctx.traceArrowsRef.current.nextId}`
      ctx.traceArrowsRef.current.nextId += 1
      if (command === 'trace-precedents') {
        const formula = worksheet.getRange(row, column).getFormula()
        if (!formula) {
          ctx.setMessage(t('appTraceNoFormula'))
          return
        }
        const result = collectPrecedents(formula, worksheet.getSheetName(), {
          lastRow: worksheet.getLastRow(),
          lastColumn: worksheet.getLastColumn(),
          maxRows: worksheet.getMaxRows(),
          maxColumns: worksheet.getMaxColumns(),
        })
        if (result.areas.length === 0 && result.offSheet === 0) {
          ctx.setMessage(t('appTraceNoRefs'))
          return
        }
        ctx.traceArrowsRef.current.disposables.push(
          ...installTraceArrows(
            runtime,
            worksheet,
            result.areas.map((area) => ({
              from: { row: area.startRow, column: area.startColumn },
              to: { row, column },
            })),
            result.areas,
            idPrefix,
          ),
        )
        ctx.setMessage(
          t('appTracedPrecedents', { count: result.areas.length }) +
            (result.offSheet > 0
              ? ` ${t('appTraceOffSheetRefs', { count: result.offSheet })}`
              : ''),
        )
        return
      }
      const result = collectDependents(workbook.getSnapshot(), worksheet.getSheetName(), {
        row,
        column,
      })
      const loadedNote =
        ctx.lazyWorkbookRef.current && !ctx.lazyWorkbookRef.current.flags.preloadComplete
          ? t('appLoadedRegionOnly')
          : ''
      if (result.cells.length === 0 && result.offSheet === 0) {
        ctx.setMessage(t('appTraceNoDependents', { note: loadedNote }))
        return
      }
      ctx.traceArrowsRef.current.disposables.push(
        ...installTraceArrows(
          runtime,
          worksheet,
          result.cells.map((cell) => ({ from: { row, column }, to: cell })),
          [],
          idPrefix,
        ),
      )
      ctx.setMessage(
        t('appTracedDependents', { count: result.cells.length, note: loadedNote }) +
          (result.offSheet > 0 ? ` ${t('appTraceOffSheetDeps', { count: result.offSheet })}` : ''),
      )
      return
    }
    case 'remove-arrows': {
      const traced = ctx.traceArrowsRef.current.disposables
      ctx.traceArrowsRef.current.disposables = []
      for (const disposable of traced) {
        try {
          disposable.dispose()
        } catch {
          // The float layer already died with a closed workbook.
        }
      }
      ctx.setMessage(traced.length > 0 ? t('appTraceArrowsRemoved') : t('appNoTraceArrows'))
      return
    }
    case 'workbook-statistics': {
      const workbook = runtime.univerAPI.getActiveWorkbook()
      if (!workbook) return
      const snapshot = workbook.getSnapshot()
      let cells = 0
      let formulas = 0
      for (const sheet of Object.values(snapshot.sheets)) {
        for (const row of Object.values(sheet.cellData ?? {})) {
          for (const cell of Object.values(row ?? {}) as (ICellData | null | undefined)[]) {
            if (!cell) continue
            if (cell.v !== undefined && cell.v !== null && cell.v !== '') cells += 1
            if (typeof cell.f === 'string' && cell.f.length > 0) formulas += 1
          }
        }
      }
      const sheetCount = Object.keys(snapshot.sheets).length
      const loadedNote =
        ctx.lazyWorkbookRef.current && !ctx.lazyWorkbookRef.current.flags.preloadComplete
          ? t('appLoadedRegionOnly')
          : ''
      ctx.setMessage(
        t('appWorkbookStats', {
          sheets: sheetCount,
          cells: cells.toLocaleString(),
          formulas: formulas.toLocaleString(),
          note: loadedNote,
        }),
      )
      return
    }
    case 'freeze-here': {
      const active = runtime.univerAPI.getActiveWorkbook()?.getActiveRange()
      if (worksheet && active) {
        worksheet.setFreeze({
          startRow: active.getRow(),
          startColumn: active.getColumn(),
          xSplit: active.getColumn(),
          ySplit: active.getRow(),
        })
        ctx.setMessage(t('appFrozenAtSelection'))
      }
      return
    }
    case 'unfreeze':
      if (worksheet) worksheet.cancelFreeze()
      return
    case 'insert-row-here':
    case 'delete-row-here':
    case 'insert-col-here':
    case 'delete-col-here': {
      const active = runtime.univerAPI.getActiveWorkbook()?.getActiveRange()
      if (!worksheet || !active) {
        ctx.setMessage(t('appSelectCellFirst'))
        return
      }
      try {
        // BeforeCommandExecute gating decides whether the workbook allows it.
        if (command === 'insert-row-here') worksheet.insertRowsBefore(active.getRow(), 1)
        else if (command === 'delete-row-here') worksheet.deleteRows(active.getRow(), 1)
        else if (command === 'insert-col-here') worksheet.insertColumnsBefore(active.getColumn(), 1)
        else worksheet.deleteColumns(active.getColumn(), 1)
      } catch {
        // A canceled structural command already surfaced its own message.
      }
      return
    }
    default:
      break
  }
  if (command.startsWith('format-as-table:')) {
    handleFormatAsTable(ctx.dataToolsContext(), command.slice('format-as-table:'.length))
    return
  }
  if (command.startsWith('use-in-formula:')) {
    // Unlike Excel, this commits immediately instead of opening the editor,
    // so a non-empty cell (say, the header a Create-from-Selection left
    // selected) would silently lose its formula or value — refuse instead.
    const workbook = runtime.univerAPI.getActiveWorkbook()
    const active = workbook?.getActiveRange()
    if (!workbook || !worksheet || !active) {
      ctx.setMessage(t('appSelectCellFirst'))
      return
    }
    const cell = worksheet.getRange(active.getRow(), active.getColumn(), 1, 1)
    const value = cell.getValue()
    if (cell.getFormula() || (value != null && value !== '')) {
      ctx.setMessage(t('appUseInFormulaNeedsEmptyCell'))
      return
    }
    // Defined names contain no colons, so the remainder is the whole name.
    const error = handleApplyFormula(
      ctx.dataToolsContext(),
      `=${command.slice('use-in-formula:'.length)}`,
    )
    if (error) ctx.setMessage(error)
    return
  }
  if (command === 'create-names:top' || command === 'create-names:left') {
    handleCreateNamesFromSelection(
      ctx.dataToolsContext(),
      command === 'create-names:top' ? 'top' : 'left',
    )
    return
  }
  if (command.startsWith('cell-style:')) {
    const presets = CELL_STYLE_PRESETS[command.slice('cell-style:'.length)]
    const range = runtime.univerAPI.getActiveWorkbook()?.getActiveRange()
    if (!presets || !range) {
      ctx.setMessage(t('appSelectCellsFirst'))
      return
    }
    for (const patch of presets) applyFormatPatchToRange(range, patch)
    ctx.setMessage(t('appCellStyleApplied'))
    return
  }
  if (command.startsWith('insert-chart:')) {
    void handleInsertChart(ctx.visualContext(), command.slice('insert-chart:'.length))
    return
  }
  if (command.startsWith('insert-pivot-chart:')) {
    void handleInsertPivotChart(ctx.visualContext(), command.slice('insert-pivot-chart:'.length))
    return
  }
  if (command === 'slicer-open') {
    handleOpenSlicerPicker(ctx.pivotContext())
    return
  }
  if (command === 'timeline-open') {
    handleOpenTimelinePicker(ctx.pivotContext())
    return
  }
  if (command.startsWith('insert-shape:')) {
    // Excel parity: gallery pick arms the crosshair draw mode (click = default
    // size, drag = custom, Shift = square, Esc = cancel)
    startShapeDraw(ctx.visualContext(), command.slice('insert-shape:'.length))
    return
  }
  if (command === 'insert-textbox') {
    handleInsertShape(ctx.visualContext(), 'rect', true)
    return
  }
  if (command === 'insert-picture') {
    handleInsertPicture(ctx.visualContext())
    return
  }
  if (command === 'format-as-table') {
    handleFormatAsTable(ctx.dataToolsContext(), 'TableStyleMedium2')
    return
  }
  if (command.startsWith('page-layout:')) {
    ctx.handlePageLayoutCommand(command.slice('page-layout:'.length))
    return
  }
  if (command === 'export-pdf') {
    void ctx.handleExportPdf()
    return
  }
  if (command.startsWith('row-height:') || command.startsWith('col-width:')) {
    const active = runtime.univerAPI.getActiveWorkbook()?.getActiveRange()
    if (!worksheet || !active) {
      ctx.setMessage(t('appSelectCellFirst'))
      return
    }
    const isRow = command.startsWith('row-height:')
    const value = Number(command.slice(command.indexOf(':') + 1))
    // file units: points for rows (max 409.5), character width for columns (max 255)
    if (!Number.isFinite(value) || value < 0 || value > (isRow ? 409.5 : 255)) return
    const area = active.getRange()
    if (isRow) {
      worksheet.setRowHeightsForced(
        area.startRow,
        area.endRow - area.startRow + 1,
        Math.round((value * 96) / 72),
      )
    } else {
      worksheet.setColumnWidths(
        area.startColumn,
        area.endColumn - area.startColumn + 1,
        Math.round(characterWidthToPixels(value)),
      )
    }
    return
  }
  if (command.startsWith('zoom:')) {
    const percent = Number(command.slice('zoom:'.length))
    if (worksheet && Number.isInteger(percent) && percent >= 25 && percent <= 400) {
      worksheet.zoom(percent / 100)
      ctx.setMessage(t('appZoom', { percent }))
    }
    return
  }
  const range = runtime.univerAPI.getActiveWorkbook()?.getActiveRange()
  if (!range) {
    ctx.setMessage(t('appSelectRangeFirst'))
    return
  }
  const { name, argument, extra } = parseStyleCommand(command)
  try {
    const style = selectionStyle(range)
    switch (name) {
      case 'bold':
        // An explicit :on/:off argument sets absolutely (Format Cells
        // dialog); no argument keeps ribbon toggle behavior.
        range.setFontWeight(
          argument
            ? argument === 'on'
              ? 'bold'
              : null
            : style.bl === BooleanNumber.TRUE
              ? null
              : 'bold',
        )
        break
      case 'italic':
        range.setFontStyle(
          argument
            ? argument === 'on'
              ? 'italic'
              : null
            : style.it === BooleanNumber.TRUE
              ? null
              : 'italic',
        )
        break
      case 'underline': {
        if (argument === 'double') {
          // Toggle: Excel's double-underline button removes an existing
          // double underline. setValue merges the style patch range-wide.
          const current = style.ul as { s?: number; t?: number } | undefined
          const isDouble = current?.s === BooleanNumber.TRUE && current.t === 10
          range.setValue({
            s: { ul: isDouble ? null : { s: BooleanNumber.TRUE, t: 10 } },
          } as unknown as ICellData)
          break
        }
        range.setFontLine(
          argument
            ? argument === 'on'
              ? 'underline'
              : null
            : style.ul?.s === BooleanNumber.TRUE
              ? null
              : 'underline',
        )
        break
      }
      case 'strike':
        range.setFontLine(
          argument
            ? argument === 'on'
              ? 'line-through'
              : null
            : style.st?.s === BooleanNumber.TRUE
              ? null
              : 'line-through',
        )
        break
      case 'align':
        // justify/distributed set the raw style key (same mutation shape,
        // journals identically). The facade helper only accepts
        // 'left' | 'center' | 'normal' — and 'normal' (confusingly) maps to
        // RIGHT — so translate 'right' before calling it; anything else
        // would hit its throwing default branch.
        if (argument === 'justify' || argument === 'distributed') {
          range.setValue({
            s: { ht: argument === 'justify' ? 4 : 6 },
          } as unknown as ICellData)
        } else {
          range.setHorizontalAlignment(
            argument === 'right' ? 'normal' : (argument as 'left' | 'center'),
          )
        }
        break
      case 'valign':
        range.setVerticalAlignment(argument as 'top' | 'middle' | 'bottom')
        break
      case 'wrap':
        // An explicit :on/:off sets absolutely (Format Cells dialog).
        range.setWrap(argument ? argument === 'on' : style.tb !== WrapStrategy.WRAP)
        break
      case 'indent': {
        // Renders as left padding (pd) and journals through the normal
        // set-range-values channel, so it is visible and undoable.
        const steps = Number(argument)
        if (!Number.isInteger(steps) || steps < 0 || steps > 250) return
        range.setValue({
          s: { pd: steps === 0 ? null : { l: steps * INDENT_STEP_PX } },
        } as unknown as ICellData)
        break
      }
      case 'cellprot': {
        // No Univer model for cell protection — journal the neutral delta
        // directly. File-only: not rendered, not undoable.
        const state = ctx.lazyWorkbookRef.current
        if (!state) {
          ctx.setMessage(t('appSettingNeedsFile'))
          return
        }
        const sheetId = worksheet?.getSheetId()
        if (!sheetId || isSheetRemoved(state.editJournal, sheetId)) return
        if (range.getHeight() * range.getWidth() > 10_000) {
          ctx.setMessage(t('appTooManyCellsForSetting'))
          return
        }
        const delta: WorkbookStyleEdit = {}
        if (argument === 'locked-on') delta.protectionLocked = true
        else if (argument === 'locked-off') delta.protectionLocked = false
        if (extra === 'hidden-on') delta.protectionHidden = true
        else if (extra === 'hidden-off') delta.protectionHidden = false
        if (Object.keys(delta).length === 0) return
        for (let row = range.getRow(); row < range.getRow() + range.getHeight(); row += 1) {
          for (
            let column = range.getColumn();
            column < range.getColumn() + range.getWidth();
            column += 1
          ) {
            recordNeutralStyleEdit(state.editJournal, sheetId, row, column, delta)
          }
        }
        ctx.setPendingEdits(journalSize(state.editJournal))
        ctx.setMessage(t('appProtectionFlagsRecorded'))
        return
      }
      case 'merge':
        if (argument === 'unmerge') {
          range.breakApart()
        } else if (argument === 'across') {
          range.mergeAcross()
        } else {
          range.merge()
          if (argument === 'center') range.setHorizontalAlignment('center')
        }
        break
      case 'format':
        range.setNumberFormat(argument)
        break
      case 'font-family':
        range.setFontFamily(argument)
        break
      case 'font-size':
        range.setFontSize(Number(argument))
        break
      case 'fill':
        // The facade types demand a string, but null is the documented way
        // to clear the background (mirrors setFontWeight(null) above).
        range.setBackground((argument === 'none' ? null : argument) as unknown as string)
        break
      case 'font-color':
        // 'auto' clears the explicit color back to the default (Excel's
        // Automatic); null is the documented reset, same as setBackground
        range.setFontColor((argument === 'auto' ? null : argument) as unknown as string)
        break
      case 'sort': {
        if (range.getHeight() < 2) {
          ctx.setMessage(t('appSortSelectRows'))
          return
        }
        // Set the outcome message first: if the streaming gate cancels the
        // sort command, its explanation lands afterwards and wins.
        ctx.setMessage(argument === 'asc' ? t('appSortedAsc') : t('appSortedDesc'))
        range.sort({ column: 0, ascending: argument === 'asc' })
        return
      }
      case 'autofn': {
        if (!worksheet) return
        if (range.getHeight() < 2) {
          ctx.setMessage(t('appAutofnSelectCells', { fn: argument }))
          return
        }
        const startRow = range.getRow()
        const startColumn = range.getColumn()
        const lastRow = startRow + range.getHeight() - 1
        // Same protection as manual edits: never write into a row whose
        // original content has not streamed in yet (see BeforeSheetEditStart).
        const state = ctx.lazyWorkbookRef.current
        if (state && !state.flags.preloadComplete) {
          const sheetId = worksheet.getSheetId()
          const sheet = state.file.sheets.find((candidate) => candidate.id === sheetId)
          const loaded = state.loadedRanges.get(sheetId)
          const targetInData = sheet !== undefined && lastRow + 1 < sheet.rowCount
          const targetLoaded =
            loaded !== undefined && lastRow + 1 >= loaded.startRow && lastRow + 1 <= loaded.endRow
          if (targetInData && !targetLoaded) {
            ctx.setMessage(t('appRowBelowStreaming'))
            return
          }
        }
        for (let column = startColumn; column < startColumn + range.getWidth(); column += 1) {
          const letter = columnLabel(column)
          worksheet
            .getRange(lastRow + 1, column)
            .setFormula(`=${argument}(${letter}${startRow + 1}:${letter}${lastRow + 1})`)
        }
        ctx.setMessage(t('appAutofnInserted', { fn: argument }))
        return
      }
      case 'sort-custom': {
        // `argument` = header flag, `extra` = "12a,3d" (colIndex + order).
        if (!worksheet) return
        if (range.getHeight() < 2) {
          ctx.setMessage(t('appSortSelectRows'))
          return
        }
        const orderRules = extra.split(',').flatMap((token) => {
          const parsed = /^([0-9]+)([ad])$/.exec(token)
          if (!parsed?.[1]) return []
          return [{ type: parsed[2] === 'd' ? 'desc' : 'asc', colIndex: Number(parsed[1]) }]
        })
        if (orderRules.length === 0) return
        ctx.setMessage(t('appSortedCustom'))
        void runtime.univerAPI.executeCommand('sheet.command.sort-range', {
          unitId: runtime.univerAPI.getActiveWorkbook()?.getId(),
          subUnitId: worksheet.getSheetId(),
          range: {
            startRow: range.getRow(),
            endRow: range.getRow() + range.getHeight() - 1,
            startColumn: range.getColumn(),
            endColumn: range.getColumn() + range.getWidth() - 1,
          },
          orderRules,
          hasTitle: argument === '1',
        })
        return
      }
      case 'remove-duplicates': {
        if (!worksheet) return
        const state = ctx.lazyWorkbookRef.current
        const sheetId = worksheet.getSheetId()
        // Same data-dependency gate as sorting: partially streamed rows
        // would dedupe against data the user never saw.
        if (
          state &&
          !state.editJournal.sheets.added.has(sheetId) &&
          (!state.formulaMode || !state.flags.preloadComplete)
        ) {
          ctx.setMessage(t('appDedupeNeedsFullLoad'))
          return
        }
        if (range.getHeight() < 2) {
          ctx.setMessage(t('appDedupeSelectRows'))
          return
        }
        const values = range.getValues().map((row) => row.map((value) => value ?? null))
        const { rows, removed } = dedupeRows(values, argument === '1')
        if (removed === 0) {
          ctx.setMessage(t('appNoDuplicates'))
          return
        }
        while (rows.length < values.length) {
          rows.push(Array.from({ length: range.getWidth() }, () => null))
        }
        const startRow = range.getRow()
        const startColumn = range.getColumn()
        // Only rewrite rows that actually change, so formulas in rows that
        // stay put survive; moved rows land as their computed values.
        for (const [offset, row] of rows.entries()) {
          const current = values[offset]
          if (current && row.every((value, index) => value === current[index])) continue
          worksheet
            .getRange(startRow + offset, startColumn, 1, range.getWidth())
            // null clears the cell (the facade type omits it, the runtime
            // treats it as a full delete — same as Clear All).
            .setValues([[...row]] as unknown as ICellData[][])
        }
        ctx.setMessage(t('appDuplicatesRemoved', { count: removed }))
        return
      }
      case 'link-set':
      case 'link-remove': {
        const state = ctx.lazyWorkbookRef.current
        if (!state || !worksheet) {
          ctx.setMessage(t('appLinksNeedFile'))
          return
        }
        const sheetId = worksheet.getSheetId()
        const row = range.getRow()
        const column = range.getColumn()
        if (name === 'link-remove') {
          recordHyperlinkEdit(state.editJournal, sheetId, row, column, null)
          state.hyperlinkTargets.get(sheetId)?.delete(`${row}:${column}`)
          // Excel's Remove Hyperlink also clears the link look.
          range.setValue({ s: { ul: null, cl: null } } as unknown as ICellData)
          ctx.setMessage(t('appLinkRemoved'))
        } else {
          const target = normalizeLinkTarget(decodeURIComponent(argument))
          if (target === null) {
            ctx.setMessage(t('appLinkInvalid'))
            return
          }
          recordHyperlinkEdit(state.editJournal, sheetId, row, column, target)
          range.setValue({
            s: { cl: { rgb: '#0563C1' }, ul: { s: BooleanNumber.TRUE } },
          } as unknown as ICellData)
          ctx.setMessage(t('appLinkSaved'))
        }
        ctx.setPendingEdits(journalSize(state.editJournal))
        ctx.refreshSelectionFormatRef.current()
        return
      }
      case 'rotate': {
        // Wire values map to Univer's set-text-rotation: a number of
        // degrees (counterclockwise positive) or 'v' for stacked text.
        const value = argument === 'vertical' ? 'v' : Number(argument)
        if (value !== 'v' && !Number.isFinite(value)) return
        void runtime.univerAPI.executeCommand('sheet.command.set-text-rotation', { value })
        break
      }
      case 'flash-fill': {
        if (!worksheet) return
        const targetColumn = range.getColumn()
        if (targetColumn === 0) {
          ctx.setMessage(t('appFlashFillNeedsLeft'))
          return
        }
        const startRow = range.getRow()
        let endRow = startRow + range.getHeight() - 1
        if (range.getHeight() === 1) {
          // Single-cell selection: follow the column to the left downward.
          const probeRows = Math.min(1000, worksheet.getMaxRows() - startRow)
          const probe = worksheet
            .getRange(startRow, targetColumn - 1, probeRows, 1)
            .getValues() as (string | number | boolean | null | undefined)[][]
          let last = startRow
          for (const [offset, rowValues] of probe.entries()) {
            const value = rowValues[0]
            if (value === null || value === undefined || String(value) === '') break
            last = startRow + offset
          }
          endRow = last
        }
        const rowCount = endRow - startRow + 1
        if (rowCount < 2) {
          ctx.setMessage(t('appFlashFillNeedsRows'))
          return
        }
        const firstSource = Math.max(0, targetColumn - 6)
        const grid = worksheet
          .getRange(startRow, firstSource, rowCount, targetColumn - firstSource + 1)
          .getValues() as (string | number | boolean | null | undefined)[][]
        const sourceOf = (
          row: readonly (string | number | boolean | null | undefined)[],
        ): string[] => row.slice(0, -1).map((value) => String(value ?? ''))
        const examples: { source: string[]; output: string }[] = []
        for (const row of grid) {
          const target = row[row.length - 1]
          if (target !== null && target !== undefined && String(target) !== '') {
            examples.push({ source: sourceOf(row), output: String(target) })
            if (examples.length >= 3) break
          }
        }
        if (examples.length === 0) {
          ctx.setMessage(t('appFlashFillNeedsExamples'))
          return
        }
        const template = inferFlashFillTemplate(examples)
        if (!template) {
          ctx.setMessage(t('appFlashFillNoPattern'))
          return
        }
        let filled = 0
        const matrix: (string | number | boolean)[][] = grid.map((row) => {
          const current = row[row.length - 1]
          if (current !== null && current !== undefined && String(current) !== '') {
            return [current]
          }
          const source = sourceOf(row)
          if (source.every((value) => value.trim() === '')) return ['']
          filled += 1
          return [applyFlashFillTemplate(template, source)]
        })
        if (filled === 0) {
          ctx.setMessage(t('appFlashFillNothingToFill'))
          return
        }
        worksheet.getRange(startRow, targetColumn, rowCount, 1).setValues(matrix)
        ctx.setMessage(t('appFlashFillDone', { count: filled }))
        return
      }
      case 'sparkline': {
        if (!worksheet) return
        const state = ctx.lazyWorkbookRef.current
        if (!state) {
          ctx.setMessage(t('appSparklineNeedsFile'))
          return
        }
        const type =
          argument === 'column'
            ? ('column' as const)
            : argument === 'stacked'
              ? ('stacked' as const)
              : ('line' as const)
        if (range.getWidth() < 2) {
          ctx.setMessage(t('appSparklineNeedsCols'))
          return
        }
        const sheetId = worksheet.getSheetId()
        const sheetName = worksheet.getSheetName()
        const startRow = range.getRow()
        const startColumn = range.getColumn()
        const endColumn = startColumn + range.getWidth() - 1
        const targetColumn = endColumn + 1
        if (targetColumn >= worksheet.getMaxColumns()) {
          ctx.setMessage(t('appSparklineNoSpace'))
          return
        }
        const rows = Math.min(range.getHeight(), 200)
        const cells = Array.from({ length: rows }, (_, offset) => ({
          cell: `${columnLabel(targetColumn)}${startRow + offset + 1}`,
          sourceRef: absRangeRef(
            sheetName,
            `${columnLabel(startColumn)}${startRow + offset + 1}` +
              `:${columnLabel(endColumn)}${startRow + offset + 1}`,
          ),
        }))
        const entry = {
          id: `sparkline-${Date.now().toString(36)}-${state.editJournal.sparklineAdds.length + 1}`,
          sheetId,
          type,
          cells,
        }
        recordSparklineAdd(state.editJournal, entry)
        const refreshSparklinesHere = (): void => {
          ctx.setPendingEdits(journalSize(state.editJournal))
          queueSparklineInstall(
            runtime,
            ctx.lazyWorkbookRef,
            ctx.sparklineDisposablesRef,
            ctx.sparklineTimerRef,
            sheetId,
          )
        }
        pushVisualUndo(runtime, {
          undo: () => {
            removeSparklineAdd(state.editJournal, entry.id)
            refreshSparklinesHere()
          },
          redo: () => {
            recordSparklineAdd(state.editJournal, entry)
            refreshSparklinesHere()
          },
        })
        refreshSparklinesHere()
        ctx.setMessage(t('appSparklinesInserted', { count: cells.length }))
        return
      }
      case 'text-to-columns': {
        if (!worksheet) return
        if (range.getWidth() > 1) {
          ctx.setMessage(t('appTextToColsSelectOne'))
          return
        }
        const delimiter = Number(argument)
        if (!Number.isInteger(delimiter) || delimiter <= 0) return
        ctx.setMessage(t('appSplitIntoColumns'))
        void runtime.univerAPI.executeCommand('sheet.command.split-text-to-columns', {
          range: {
            startRow: range.getRow(),
            endRow: range.getRow() + range.getHeight() - 1,
            startColumn: range.getColumn(),
            endColumn: range.getColumn(),
          },
          delimiter,
        })
        return
      }
      case 'border': {
        const type = BORDER_COMMAND_TYPES[argument]
        if (!type) return
        // `extra` is "<#color>[:<line-style>]" — the ribbon sends only the
        // color, the Format Cells dialog appends an ST_BorderStyle name
        const [colorPart = '', stylePart = ''] = extra.split(':')
        const borderStyle =
          BORDER_LINE_STYLE_TYPES[stylePart] ??
          (argument === 'thick-outer' ? BorderStyleTypes.MEDIUM : BorderStyleTypes.THIN)
        range.setBorder(
          type,
          borderStyle,
          /^#[0-9a-fA-F]{6}$/.test(colorPart) ? colorPart : '#000000',
        )
        break
      }
      default:
        return
    }
    ctx.setMessage(t('appAppliedToSelection'))
  } catch (error: unknown) {
    ctx.setMessage(error instanceof Error ? error.message : t('appCommandFailed'))
  }
}
