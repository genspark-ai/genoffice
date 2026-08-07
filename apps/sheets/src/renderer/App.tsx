import {
  absRangeRef,
  activateFormulaClosure,
  applyAiConditionalFormat,
  applyAiDataValidation,
  applyAiHyperlink,
  applyDefinedNames,
  applyFilterCriteria,
  applyFormatPatchToRange,
  applyWorkbookNotes,
  cellValueBounds,
  clearLazyState,
  columnLetter,
  disposeVisuals,
  ensureLazyRangeLoaded,
  journalRangeSnapshot,
  lazyCellReader,
  loadSnapshotIntoUniver,
  loadVisibleRange,
  loadWorkbookSkeleton,
  matrixBounds,
  measureImage,
  navigateToAnchor,
  preloadEntireWorkbook,
  protectSheetGuard,
  queueFormulaRecalc,
  queueSparklineInstall,
  RECALC_MAX_FAILURES,
  queueVisualInstall,
  sheetOutline,
  syncUniver,
  univerDefinedNames,
} from './univer-sync'
import {
  journalSuppression,
  type ActiveWorkbook,
  type LazyWorkbookState,
  type UniverRuntime,
} from './univer-state'
import {
  applyAiPivotAdd,
  applyAiTableAdd,
  applyAiTableColumnAdd,
  applyAiTableColumnDelete,
  applyAiTableRowAdd,
  applyAiTableRowDelete,
  renameChartRefsForSheet,
} from './workbook-ops'
import {
  proposeOperations as proposeOperationsImpl,
  runDeterministicPlan as runDeterministicPlanImpl,
  type PlanContext,
} from './plan-operations'
import { isNumericIdentifierText } from './cell-warning'
import { useCallback, useEffect, useRef, useState, useSyncExternalStore } from 'react'

import {
  CellValueType,
  getNumfmtParseValueFilter,
  InterceptorEffectEnum,
  isRealNum,
  IUndoRedoService,
  LocaleType,
  mergeLocales,
  type ICellData,
  type IRange,
  type IStyleData,
} from '@univerjs/core'
import { UniverSheetsConditionalFormattingPreset } from '@univerjs/preset-sheets-conditional-formatting'
import UniverPresetSheetsConditionalFormattingEnUS from '@univerjs/preset-sheets-conditional-formatting/locales/en-US'
import '@univerjs/preset-sheets-conditional-formatting/lib/index.css'
import {
  INTERCEPTOR_POINT,
  SheetInterceptorService,
  UniverSheetsCorePreset,
} from '@univerjs/preset-sheets-core'
import UniverPresetSheetsCoreEnUS from '@univerjs/preset-sheets-core/locales/en-US'
import { UniverSheetsDataValidationPreset } from '@univerjs/preset-sheets-data-validation'
import UniverPresetSheetsDataValidationEnUS from '@univerjs/preset-sheets-data-validation/locales/en-US'
import '@univerjs/preset-sheets-data-validation/lib/index.css'
import { UniverSheetsDrawingPreset } from '@univerjs/preset-sheets-drawing'
import '@univerjs/preset-sheets-drawing/lib/index.css'
import { UniverSheetsFindReplacePreset } from '@univerjs/preset-sheets-find-replace'
import UniverPresetSheetsFindReplaceEnUS from '@univerjs/preset-sheets-find-replace/locales/en-US'
import '@univerjs/preset-sheets-find-replace/lib/index.css'
import { UniverSheetsFilterPreset } from '@univerjs/preset-sheets-filter'
import UniverPresetSheetsFilterEnUS from '@univerjs/preset-sheets-filter/locales/en-US'
import '@univerjs/preset-sheets-filter/lib/index.css'
import { UniverSheetsNotePreset } from '@univerjs/preset-sheets-note'
import UniverPresetSheetsNoteEnUS from '@univerjs/preset-sheets-note/locales/en-US'
import '@univerjs/preset-sheets-note/lib/index.css'
import { UniverSheetsSortPreset } from '@univerjs/preset-sheets-sort'
import UniverPresetSheetsSortEnUS from '@univerjs/preset-sheets-sort/locales/en-US'
import '@univerjs/preset-sheets-sort/lib/index.css'
import { UniverSheetsTablePreset } from '@univerjs/preset-sheets-table'
import UniverPresetSheetsTableEnUS from '@univerjs/preset-sheets-table/locales/en-US'
import '@univerjs/preset-sheets-table/lib/index.css'
import { greenTheme } from '@univerjs/themes'
import { createUniver } from './create-univer'

import {
  AgentLoop,
  COMPLETED_VIA_TOOLS_TEXT,
  composeSkills,
  type AgentImage,
} from '@genoffice/agent-core'
import type { AiSettings } from '@genoffice/ai-provider'
import { type WorkbookOperation } from '../domain/workbook-dsl'
import { columnIndex, columnLabel, parseAddress, parseRange } from '../domain/cell-address'
import {
  applyChartStateEdit,
  chartSupportsDataLabels,
  chartSupportsSeriesReplace,
  withDefaultBarLabels,
  type CellBounds,
} from '../domain/chart-visual'
import { InMemoryWorkbookAdapter } from '../domain/in-memory-workbook'
import { cfRuleUnsaveableReason, iconSetSaveable } from '../gateway/xlsx-cf'
import type { ApplyOutcome, ChangePlan } from '../domain/workbook.types'
import { createElectronTransport } from './ai/transport'
import type { ActiveSheetInfo, SheetsSkillDeps } from './ai/tools'
import type { AiChatMessage } from './ai/AiChatPanel'
import { createWorkbookSkill } from './ai/workbook-skill'
import { createFilesSkill } from './ai/files-skill'
import { createSearchSkill } from './ai/search-skill'
import { ATTACHMENT_IMAGE_EXTS } from '../shared/desktop-api'
import type {
  AttachmentAddResult,
  AttachmentMeta,
  MenuAction,
  WorkbookFile,
  WorkbookVisualObject,
} from '../shared/desktop-api'
import type { PageSetupJournalState, StructuralJournalOp } from './edit-journal'
import {
  AUTO_FILL_COMMAND,
  AXIS_ATTR_MUTATIONS,
  BLOCKED_COMMAND_PATTERN,
  CF_MUTATIONS,
  CF_RULE_COMMAND_PATTERN,
  CHAT_STORAGE_KEY,
  COPY_SHEET_COMMAND,
  DEFINED_NAME_MUTATIONS,
  DV_EDIT_COMMAND_PATTERN,
  DV_MUTATIONS,
  EMPTY_CHART_EDITS,
  FILTER_COMMAND_PATTERN,
  FILTER_MUTATIONS,
  FORMULA_MODE_MAX_CELLS,
  initialSnapshot,
  MERGE_MUTATIONS,
  MOVE_RANGE_COMMAND,
  MOVE_ROWS_COMMAND,
  MOVE_ROWS_MUTATION,
  MOVE_RANGE_MUTATION,
  NOTE_MUTATIONS,
  PERSIST_TOOL_FIELD_MAX,
  pixelsToCharacterWidth,
  REMOVE_NUMFMT_MUTATION,
  REORDER_RANGE_MUTATION,
  ROW_COLUMN_MUTATIONS,
  safeJsonInput,
  SET_FROZEN_MUTATION,
  SET_NUMFMT_MUTATION,
  TOGGLE_GRIDLINES_MUTATION,
  SET_RANGE_VALUES_MUTATION,
  SHEET_LIFECYCLE_MUTATIONS,
  SORT_COMMAND_PATTERN,
  STRUCTURAL_EDIT_COMMAND_PATTERN,
} from './app-constants'
import {
  getActiveSheetInfo as getActiveSheetInfoImpl,
  readCells as readCellsImpl,
  readFormats as readFormatsImpl,
  readSheetFeatures as readSheetFeaturesImpl,
  type WorkbookReadContext,
} from './ai/workbook-readers'
import {
  getSourceRange as getSourceRangeImpl,
  handleCreatePivot as handleCreatePivotImpl,
  handleCreateSlicer as handleCreateSlicerImpl,
  handleEditPivotApply as handleEditPivotApplyImpl,
  handleRefreshPivot as handleRefreshPivotImpl,
  handleCreateTimeline as handleCreateTimelineImpl,
  handleRemoveSlicer as handleRemoveSlicerImpl,
  handleRemoveTimeline as handleRemoveTimelineImpl,
  handleSlicerSelectAll as handleSlicerSelectAllImpl,
  handleSlicerToggle as handleSlicerToggleImpl,
  handleTimelineRange as handleTimelineRangeImpl,
  isSelectionInPivot as isSelectionInPivotImpl,
  pivotEditInitial as pivotEditInitialImpl,
  pivotFieldOptions as pivotFieldOptionsImpl,
  refreshPivotTables as refreshPivotTablesImpl,
  type PivotActionContext,
  type PivotEditContext,
  type SlicerPickerState,
  type TimelinePickerState,
} from './pivot-actions'
import type { ChartRecommendations } from '../domain/chart-recommend'
import {
  applyAiShapeEdit as applyAiShapeEditImpl,
  buildAiChartEdit as buildAiChartEditImpl,
  handleInsertChart as handleInsertChartImpl,
  handleInsertEquation as handleInsertEquationImpl,
  handleInsertIcon as handleInsertIconImpl,
  handleInsertScreenshot,
  handleRecommendedCharts as handleRecommendedChartsImpl,
  insertAiChartVisual as insertAiChartVisualImpl,
  insertAiImageVisual as insertAiImageVisualImpl,
  insertAiShapeVisual as insertAiShapeVisualImpl,
  type VisualActionContext,
} from './visual-actions'
import {
  activeCellLabel as activeCellLabelImpl,
  consolidateDefaultReference as consolidateDefaultReferenceImpl,
  goToReference as goToReferenceImpl,
  handleApplyAdvancedFilter as handleApplyAdvancedFilterImpl,
  handleApplyFormula as handleApplyFormulaImpl,
  handleCreateConsolidate as handleCreateConsolidateImpl,
  handleCreateSubtotal as handleCreateSubtotalImpl,
  handleInsertSymbol as handleInsertSymbolImpl,
  listDefinedNames as listDefinedNamesImpl,
  type DataToolsContext,
} from './data-tools-actions'
import { installTsvClipboardFix } from './clipboard-tsv'
import { installFilteredCopyHook } from './filtered-copy'
import {
  applyShowFormulasView,
  installFormulaTextInterceptor,
  installFormulaViewInterceptor,
} from './formula-view'
import { installCellFilenameFunction } from './cell-function'
import { installFormulaLexerFix } from './formula-lexer-fix'
import { installSheetRenameFix } from './sheet-rename-fix'
import { installSelectionWrapGuard } from './selection-wrap-fix'
import { installCopyMaterialize } from './copy-materialize'
import { applyUniverLocale } from './univer-locales'
import { installRuleDetail } from './univer-rule-detail'
import { installFormulaNullResultFix } from './formula-null-result'
import { installNumberFormatFix } from './numfmt-fix'
import { installRateFallback } from './rate-function'
import {
  handleRibbonCommand as handleRibbonCommandImpl,
  type RibbonCommandContext,
} from './ribbon-actions'
import {
  handleApplyHeaderFooter as handleApplyHeaderFooterImpl,
  handleExportPdf as handleExportPdfImpl,
  handlePageLayoutCommand as handlePageLayoutCommandImpl,
  type PageLayoutContext,
} from './page-layout-actions'
import { handleSave as handleSaveImpl, type SaveContext } from './save-actions'
import {
  applyChartEdit as applyChartEditImpl,
  applyShapeEdit as applyShapeEditImpl,
  queueChartDataSync as queueChartDataSyncImpl,
  readChartVector as readChartVectorImpl,
  type VisualSyncContext,
} from './visual-edit-sync'

import {
  createEditJournal,
  isSheetRemoved,
  hyperlinkEditAt,
  journalSize,
  recordCfChange,
  recordPageSetup,
  recordDefinedNamesChange,
  recordDvChange,
  recordNoteChange,
  recordSheetProtection,
  recordFilterChange,
  recordSetNumfmt,
  recordSetRangeValues,
  recordSheetHidden,
  recordSheetDuplicate,
  recordSheetInsert,
  recordSheetOrderChange,
  recordSheetRemove,
  recordSheetRename,
  recordStructuralOp,
  shiftVisualForStructuralOp,
  removeTableAdd,
  recordSparklineAdd,
} from './edit-journal'
import { shiftPinnedCells } from './formula-closure'
import { getLang, t, aiLangDirective } from './i18n/locale'
import { planStillMatches } from './lazy-plan'
import { netAxisDelta, screenToFile } from './view-transform'
import { selectionFormatEquals, toSelectionFormat, type SelectionFormat } from './selection-format'
import { ExcelShell } from './ExcelShell'
import { ToastHost } from './toast'
import { AdvancedFilterDialog, type AdvancedFilterColumn } from './AdvancedFilterDialog'
import { EquationDialog } from './EquationDialog'
import { IconsDialog } from './IconsDialog'
import { RecommendedChartsDialog } from './RecommendedChartsDialog'
import { ScreenshotDialog } from './ScreenshotDialog'
import { SymbolDialog } from './SymbolDialog'
import { SlicerFieldPicker, SlicerPanels, type SlicerUiState } from './SlicerPanel'
import { TimelineFieldPicker, TimelinePanels, type TimelineUiState } from './TimelinePanel'
import type { DefinedNameAction, DefinedNameRow } from './NameManagerDialog'
import {
  clearVisualSelection,
  convertibleType,
  getChartElementSelection,
  installWorkbookVisuals,
  isVisualDragActive,
  setChartDialogListener,
  setVisualSelectionListener,
  subscribeChartElementSelection,
  type ChartDialogKind,
  type ChartEditData,
  type ChartVectorRead,
  type ShapeEditChanges,
} from './WorkbookVisuals'
import { ChartFormatPane, SelectDataDialog } from './ChartPanels'

// Source sheet id of an in-flight copy-sheet command; the next insert-sheet
// mutation is that copy and must journal as a duplicate, not a blank add.
let pendingCopySource: string | undefined

export function App(): React.JSX.Element {
  const adapterRef = useRef(new InMemoryWorkbookAdapter(initialSnapshot))
  const univerRef = useRef<UniverRuntime | null>(null)
  const lazyWorkbookRef = useRef<LazyWorkbookState | null>(null)
  /// True while Univer's in-cell editor is open (AutoSave must not save-reload then).
  const editingCellRef = useRef(false)
  const visualDisposablesRef = useRef<{ dispose(): void }[]>([])
  const traceArrowsRef = useRef<{ disposables: { dispose(): void }[]; nextId: number }>({
    disposables: [],
    nextId: 0,
  })
  const visualInstallTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null)
  const sparklineDisposablesRef = useRef<{ dispose(): void }[]>([])
  const sparklineTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null)
  const visualViewportKeyRef = useRef('')
  const demoVisualDisposablesRef = useRef<{ dispose(): void }[]>([])
  const demoVisualInstallTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null)
  const [prompt, setPrompt] = useState('')
  const [preview, setPreview] = useState<ChangePlan | null>(null)
  const [_revision, setRevision] = useState(0)
  const [workbookFile, setWorkbookFile] = useState<WorkbookFile | null>(null)
  const [pendingEdits, setPendingEdits] = useState(0)
  /// Whether any cell in the workbook has content — the ribbon's one-click AI
  /// action buttons are greyed out on a fully empty sheet.
  const [sheetHasContent, setSheetHasContent] = useState(false)
  const recomputeSheetContent = useCallback(() => {
    setSheetHasContent(() => {
      // A file opened from disk always counts as having content.
      if (lazyWorkbookRef.current) return true
      const workbook = univerRef.current?.univerAPI.getActiveWorkbook()
      if (!workbook) return false
      const snapshot = workbook.getSnapshot()
      for (const sheet of Object.values(snapshot.sheets ?? {})) {
        for (const row of Object.values(sheet.cellData ?? {})) {
          for (const cell of Object.values(row ?? {}) as (ICellData | null | undefined)[]) {
            if (!cell) continue
            if (cell.f || cell.p) return true
            if (cell.v !== undefined && cell.v !== null && String(cell.v) !== '') return true
          }
        }
      }
      return false
    })
  }, [])
  // Open/close of a real file swaps the whole data source; re-evaluate then too.
  useEffect(() => {
    recomputeSheetContent()
  }, [workbookFile, recomputeSheetContent])
  // The close guard lives in the main process; keep it fed with the badge count.
  useEffect(() => {
    window.desktopApi?.notifyPendingEdits?.(pendingEdits)
  }, [pendingEdits])
  const [autoSave, setAutoSave] = useState(
    () => localStorage.getItem('ai-sheets-auto-save') === '1',
  )
  // Ref mirror for callbacks captured when an AI run starts
  const autoSaveRef = useRef(autoSave)
  autoSaveRef.current = autoSave
  useEffect(() => {
    localStorage.setItem('ai-sheets-auto-save', autoSave ? '1' : '0')
  }, [autoSave])
  // AutoSave tick (docs/slides parity): every 30 s and on window blur, flush
  // pending edits of the open workbook. The journal is read at tick time so
  // the interval stays stable; demo mode has no backing file and is skipped.
  useEffect(() => {
    if (!autoSave) return
    let saving = false
    const tick = () => {
      const state = lazyWorkbookRef.current
      if (saving || !state || journalSize(state.editJournal) === 0) return
      // Never while the in-cell editor is open (saving reloads the workbook
      // and would wipe the edit), and never for converted .xls/.csv imports
      // whose first save opens a Save As dialog.
      if (editingCellRef.current || state.file.needsSaveAs) return
      saving = true
      void handleSaveRef.current('save', true).finally(() => {
        saving = false
      })
    }
    const id = window.setInterval(tick, 30_000)
    window.addEventListener('blur', tick)
    return () => {
      window.clearInterval(id)
      window.removeEventListener('blur', tick)
    }
  }, [autoSave])

  // Crash-recovery copy: independent of the AutoSave pill — a dirty
  // workbook gets a real .xlsx copy under userData every 30 s, so a force-quit or a
  // renderer crash no longer costs everything since the last manual save. A normal
  // save removes the copy; reopening a file whose copy is newer offers Restore.
  useEffect(() => {
    let writing = false
    const tick = () => {
      const state = lazyWorkbookRef.current
      if (writing || !state || journalSize(state.editJournal) === 0) return
      // The in-cell editor's pending text is not in the journal yet, and a
      // converted import has no original file to recover into.
      if (editingCellRef.current || state.file.needsSaveAs) return
      writing = true
      void handleSaveRef.current('recovery').finally(() => {
        writing = false
      })
    }
    const id = window.setInterval(tick, 30_000)
    return () => window.clearInterval(id)
  }, [])
  const [message, setMessage] = useState(t('appReadyInitial'))
  /// Zoom of the active sheet in percent, echoed by the status-bar slider.
  const [zoomPercent, setZoomPercent] = useState(100)
  const [selectionFormat, setSelectionFormat] = useState<SelectionFormat | null>(null)
  /// A1 label of the active cell, echoed live by the Name Box. Updated from
  /// the same SelectionChanged refresh that keeps selectionFormat current.
  const [activeCellA1, setActiveCellA1] = useState('')
  /// Non-null while the Advanced Filter dialog is open: the column choices
  /// sampled from the active filter range's header row.
  const [advancedFilterColumns, setAdvancedFilterColumns] = useState<
    readonly AdvancedFilterColumn[] | null
  >(null)
  /// True while the Insert → Symbol dialog is open.
  const [symbolDialogOpen, setSymbolDialogOpen] = useState(false)
  const [screenshotDialogOpen, setScreenshotDialogOpen] = useState(false)
  const [iconsDialogOpen, setIconsDialogOpen] = useState(false)
  const [equationDialogOpen, setEquationDialogOpen] = useState(false)
  const [recommendedCharts, setRecommendedCharts] = useState<ChartRecommendations | null>(null)
  /// The focused floating visual (chart/shape/image); charts surface a
  /// contextual Chart Design ribbon tab while selected.
  const [selectedVisual, setSelectedVisual] = useState<WorkbookVisualObject | null>(null)
  /// Chart panels (Select Data dialog / format task pane), opened from the
  /// ribbon or the chart context menu, keyed like chart edits.
  const [chartDialog, setChartDialog] = useState<{ kind: ChartDialogKind; editKey: string } | null>(
    null,
  )
  const chartElement = useSyncExternalStore(
    subscribeChartElementSelection,
    getChartElementSelection,
  )
  /// Bumped on every visual/chart edit: journal edits merge in place, so the
  /// journal size (pendingEdits) alone misses same-target re-edits and the
  /// ribbon echo would go stale.
  const [, setVisualEditTick] = useState(0)
  /// In-session slicers (OOXML slicer part persistence: see the TODO in
  /// SlicerPanel).
  const [slicers, setSlicers] = useState<readonly SlicerUiState[]>([])
  /// Non-null while the "Insert Slicer" field picker is open.
  const [slicerPicker, setSlicerPicker] = useState<SlicerPickerState | null>(null)
  /// In-session timelines (same session-only model as slicers).
  const [timelines, setTimelines] = useState<readonly TimelineUiState[]>([])
  /// Non-null while the "Insert Timeline" field picker is open.
  const [timelinePicker, setTimelinePicker] = useState<TimelinePickerState | null>(null)
  const menuActionRef = useRef<(action: MenuAction) => void>(() => {})
  /// Fresh handleSave for the AutoSave tick (assigned each render, like
  /// menuActionRef, so the interval closure never goes stale).
  const handleSaveRef = useRef<
    (mode: 'save' | 'save-as' | 'recovery', quiet?: boolean) => Promise<void>
  >(() => Promise.resolve())
  const closeSaveRef = useRef<() => Promise<void>>(() => Promise.resolve())
  const refreshSelectionFormatRef = useRef<() => void>(() => {})
  const chartEditRef = useRef<(chartPath: string, edit: ChartEditData) => void>(() => {})
  const chartVectorRef = useRef<(chartPath: string, range: string) => Promise<ChartVectorRead>>(
    () => Promise.reject(new Error('Workbook not ready.')),
  )
  const shapeEditRef = useRef<(visualId: string, changes: ShapeEditChanges) => void>(() => {})
  /// A3 editing of an existing pivot: context locked when the dialog opens, used
  /// on Apply.
  const pivotEditContextRef = useRef<PivotEditContext | null>(null)
  const lazyPreviewRef = useRef<{
    sessionId: string
    sheetId: string
    plan: ChangePlan
  } | null>(null)

  /** App-scope state bundle for the extracted plan builders (plan-operations.ts). */
  function planContext(): PlanContext {
    return { adapterRef, univerRef, lazyWorkbookRef, lazyPreviewRef, setPreview, autoApplySafePlan }
  }

  /** App-scope refs/state bundle for the extracted pivot actions (pivot-actions.ts). */
  function pivotContext(): PivotActionContext {
    return {
      univerRef,
      lazyWorkbookRef,
      pivotEditContextRef,
      slicers,
      slicerPicker,
      setSlicers,
      setSlicerPicker,
      timelines,
      timelinePicker,
      setTimelines,
      setTimelinePicker,
      setMessage,
      setPendingEdits,
    }
  }

  /** App-scope refs/state bundle for the extracted visual-insert actions (visual-actions.ts). */
  function visualContext(): VisualActionContext {
    return {
      adapterRef,
      univerRef,
      lazyWorkbookRef,
      visualDisposablesRef,
      visualInstallTimerRef,
      chartEditRef,
      chartVectorRef,
      shapeEditRef,
      setMessage,
      setRevision,
      setPreview,
      setPendingEdits,
      pivotContext,
      queueDemoVisualInstall,
      refreshLazyVisuals,
    }
  }

  /** App-scope refs/state bundle for the extracted data-tool actions (data-tools-actions.ts). */
  function dataToolsContext(): DataToolsContext {
    return { univerRef, lazyWorkbookRef, setMessage, setPendingEdits, setAdvancedFilterColumns }
  }

  function pageLayoutContext(): PageLayoutContext {
    return { univerRef, lazyWorkbookRef, setMessage, setPendingEdits }
  }

  function saveContext(): SaveContext {
    return { univerRef, lazyWorkbookRef, setMessage, openLazyWorkbook }
  }

  function visualSyncContext(): VisualSyncContext {
    return {
      adapterRef,
      univerRef,
      lazyWorkbookRef,
      chartSyncRef,
      setMessage,
      refreshLazyVisuals,
      refreshDemoVisuals,
    }
  }

  function proposeOperations(
    operations: readonly WorkbookOperation[],
    summary: string,
  ): { ok: true; plan: ChangePlan } | { ok: false; error: string } {
    return proposeOperationsImpl(planContext(), operations, summary)
  }

  function runDeterministicPlan(instruction: string): { text: string; isError?: boolean } {
    return runDeterministicPlanImpl(planContext(), instruction)
  }

  // ---- AI: real LLM agent (falls back to the deterministic planner above
  // when no provider is configured — see isAgentConfigured/handleSend) ----
  const [aiSettings, setAiSettingsState] = useState<AiSettings | null>(null)
  const aiSettingsRef = useRef<AiSettings | null>(null)
  aiSettingsRef.current = aiSettings
  const [aiBusy, setAiBusy] = useState(false)
  // Display history survives restarts via localStorage; the AgentLoop's model
  // context does not, so restored turns are read-only transcript.
  const [chat, setChat] = useState<readonly AiChatMessage[]>([])
  /** History loaded from project-store (read-only transcript, not fed to the model) */
  const [historicChat, setHistoricChat] = useState<readonly AiChatMessage[]>([])
  // ── Chat attachments (same structure as docs/slides: text types go through the
  // read_attachment tool, images go multimodal) ──
  const [attachments, setAttachments] = useState<readonly AttachmentMeta[]>([])
  const [attachNotice, setAttachNotice] = useState<string | null>(null)
  const attachmentsRef = useRef(attachments)
  attachmentsRef.current = attachments
  /** Synchronous re-entrancy guard between runAgent trigger and loop.run
   * (loop.busy is still false while attachment images load asynchronously) */
  const runStartingRef = useRef(false)
  /** The shell can repeat its queued-open nudge while the renderer starts.
   * Only one picker/open request may own the workbook session at a time. */
  const workbookOpeningRef = useRef(false)
  /** Current session's projectId/chatId (resolved when the workbook opens) */
  const chatRefIdsRef = useRef<{ projectId: string; chatId: string } | null>(null)

  // File renamed externally (in the shell Home list) → sync the title-bar file
  // name (the save path is synced by the main process)
  useEffect(
    () =>
      window.desktopApi?.onWorkbookRenamed?.((newName) => {
        setWorkbookFile((prev) => (prev ? { ...prev, name: newName } : prev))
      }) ?? (() => undefined),
    [],
  )

  useEffect(() => {
    setVisualSelectionListener({
      select: (visual) =>
        setSelectedVisual((current) => (current?.id === visual.id ? current : visual)),
      deselect: () => setSelectedVisual(null),
    })
    setChartDialogListener((editKey, dialog) => setChartDialog({ kind: dialog, editKey }))
    return () => {
      setVisualSelectionListener(null)
      setChartDialogListener(null)
    }
  }, [])

  // The format pane follows the chart selection; deselecting closes it.
  useEffect(() => {
    if (!selectedVisual) setChartDialog(null)
  }, [selectedVisual])

  // ── One-time migration: import legacy localStorage history into project-store ──
  useEffect(() => {
    const api = (window as Window & { projectApi?: typeof window.projectApi }).projectApi
    if (!api) return
    const raw = localStorage.getItem(CHAT_STORAGE_KEY)
    if (!raw) return
    try {
      const parsed: unknown = JSON.parse(raw)
      if (!Array.isArray(parsed) || parsed.length === 0) {
        localStorage.removeItem(CHAT_STORAGE_KEY)
        return
      }
      const msgs = parsed.filter(
        (
          e,
        ): e is {
          role: 'user' | 'assistant'
          text: string
          tools?: Array<{ summary: string; isError?: boolean }>
        } =>
          !!e &&
          typeof e === 'object' &&
          ((e as { role: string }).role === 'user' ||
            (e as { role: string }).role === 'assistant') &&
          typeof (e as { text: string }).text === 'string',
      )
      if (msgs.length === 0) {
        localStorage.removeItem(CHAT_STORAGE_KEY)
        return
      }
      // Get the default project's chatId (unsaved-0 marks the file-less default chat)
      const tempChatId = 'unsaved-legacy'
      void api
        .resolveChat({ filePath: null, tempChatId })
        .then(async (ids) => {
          for (const m of msgs) {
            const appendArgs: Parameters<typeof api.appendChat>[0] = {
              projectId: ids.projectId,
              chatId: ids.chatId,
              role: m.role,
              text: m.text,
            }
            if (m.tools && m.tools.length > 0) {
              appendArgs.tools = m.tools.map((t) => ({
                name: '',
                summary: t.summary,
                isError: !!t.isError,
              }))
            }
            await api.appendChat(appendArgs)
          }
          localStorage.removeItem(CHAT_STORAGE_KEY)
        })
        .catch(() => {
          // A failed migration doesn't affect normal use; retried on next launch
        })
    } catch {
      localStorage.removeItem(CHAT_STORAGE_KEY)
    }
  }, [])

  // ── project-store: resolve chatId and load history when a workbook opens ──
  useEffect(() => {
    const api = (window as Window & { projectApi?: typeof window.projectApi }).projectApi
    if (!api) return
    // Reset (new workbook or new session)
    chatRefIdsRef.current = null
    setHistoricChat([])
    const tempChatId = `unsaved-${Date.now()}`
    const sessionId = workbookFile?.sessionId
    const resolveArgs: Parameters<typeof api.resolveChat>[0] = { filePath: null, tempChatId }
    if (sessionId !== undefined) resolveArgs.sessionId = sessionId
    void api
      .resolveChat(resolveArgs)
      .then(async (ids) => {
        chatRefIdsRef.current = ids
        const msgs = await api.loadChat({
          projectId: ids.projectId,
          chatId: ids.chatId,
          limit: 200,
        })
        if (msgs.length === 0) return
        setHistoricChat(
          msgs.map((m) => ({
            role: m.role,
            text: m.text,
            tools:
              m.tools?.map((t) => ({
                summary: t.summary,
                isError: !!t.isError,
                ...(t.name ? { name: t.name } : {}),
                ...(t.output ? { output: t.output.slice(0, 2000) } : {}),
              })) ?? [],
          })),
        )
        // Restore model context: follow-ups after reopening the file continue the
        // previous conversation (only when the loop is idle and has no history)
        agentLoopRef.current?.restore(msgs.map((m) => ({ role: m.role, text: m.text })))
      })
      .catch(() => {
        /* silent */
      })
  }, [workbookFile?.sessionId])

  const persistChatMessage = (
    role: 'user' | 'assistant',
    text: string,
    tools?: Array<{
      name?: string
      summary: string
      isError?: boolean
      input?: string
      output?: string
    }>,
  ) => {
    const ids = chatRefIdsRef.current
    const api = (window as Window & { projectApi?: typeof window.projectApi }).projectApi
    if (!ids || !api) return
    void api
      .appendChat({
        projectId: ids.projectId,
        chatId: ids.chatId,
        role,
        text,
        ...(tools && tools.length > 0
          ? { tools: tools.map((t) => ({ ...t, name: t.name ?? '' })) }
          : {}),
      })
      .catch(() => {
        /* silent */
      })
  }

  function appendChat(entry: AiChatMessage): void {
    setChat((previous) => [...previous, entry])
  }

  function patchLastAssistant(patch: (entry: AiChatMessage) => AiChatMessage): void {
    setChat((previous) => {
      const index = previous.length - 1
      const last = previous[index]
      if (!last || last.role !== 'assistant') return previous
      const next = previous.slice()
      next[index] = patch(last)
      return next
    })
  }

  /** Tool activity for the whole run (args/output included, accumulated across
   * turns) — for full transcript persistence */
  const runToolsRef = useRef<
    Array<{ name: string; summary: string; isError?: boolean; input?: string; output?: string }>
  >([])
  /** AI plans apply asynchronously after propose_operations returns. Run
   * completion waits for these before doing the run's single auto-save. */
  const aiApplyPromisesRef = useRef<Promise<boolean>[]>([])
  /** Last non-empty streamed text of the run: a final empty turn falls back to
   * it instead of wiping the model's own summary from the tool-call turn. */
  const runLastTextRef = useRef('')
  /** true once any tool of the run mutated the workbook */
  const runMutatedRef = useRef(false)

  const agentLoopRef = useRef<AgentLoop | null>(null)
  if (!agentLoopRef.current) {
    agentLoopRef.current = new AgentLoop({
      transport: createElectronTransport(() => aiSettingsRef.current!),
      systemSuffix: aiLangDirective,
      skill: composeSkills('sheets+files', '', [
        createWorkbookSkill(sheetsSkillDeps()),
        createFilesSkill(() => attachmentsRef.current),
        createSearchSkill(),
      ]),
      // guide loading adds a tool round; the default 8 cuts off multi-step work
      maxTurns: 24,
      events: {
        onText: (text) => {
          if (text) runLastTextRef.current = text
          setMessage(text || t('appAiThinking'))
          // When the model retries successfully and keeps streaming after a
          // mid-run failure (e.g. one apply error), clear the error flag —
          // otherwise the whole successful message stays rendered in red.
          patchLastAssistant((entry) => ({ ...entry, text, isError: false }))
        },
        onToolStart: (call) => {
          // Live "running" chip: replaced in place by onToolExecuted
          patchLastAssistant((entry) => ({
            ...entry,
            tools: [
              ...entry.tools,
              {
                summary: call.name.replace(/[_-]+/g, ' '),
                isError: false,
                name: call.name,
                running: true,
              },
            ],
          }))
        },
        onToolExecuted: ({ call, execution }) => {
          if (execution.mutated) runMutatedRef.current = true
          const input = safeJsonInput(call.input)
          const output = execution.output
            ? execution.output.slice(0, PERSIST_TOOL_FIELD_MAX)
            : undefined
          runToolsRef.current.push({
            name: call.name,
            summary: execution.summary,
            isError: !!execution.isError,
            ...(input !== undefined ? { input } : {}),
            ...(output !== undefined ? { output } : {}),
          })
          patchLastAssistant((entry) => {
            // Swap out the running placeholder pushed by onToolStart (parse-fail calls have none)
            const tools = [...entry.tools]
            if (tools.at(-1)?.running) tools.pop()
            return {
              ...entry,
              tools: [
                ...tools,
                {
                  summary: execution.summary,
                  isError: !!execution.isError,
                  name: call.name,
                  ...(execution.output ? { output: execution.output.slice(0, 2000) } : {}),
                },
              ],
            }
          })
        },
        onDone: ({ text, cancelled, turnLimit }) => {
          // Prefer tool summaries when the model finished via tools with no prose
          // (agent-core fills history with COMPLETED_VIA_TOOLS_TEXT so follow-ups
          // stay provider-safe; the UI can show the real work that ran).
          const toolSummaries = (() => {
            const lines: string[] = []
            const seen = new Set<string>()
            for (const tool of runToolsRef.current) {
              if (!tool.summary || tool.isError || seen.has(tool.summary)) continue
              seen.add(tool.summary)
              lines.push(tool.summary)
              if (lines.length >= 8) break
            }
            return lines.join('\n')
          })()
          // A cancelled run must keep the "stopped" notice: earlier narration or
          // tool summaries would make an aborted run read as completed.
          const prose =
            text && text !== COMPLETED_VIA_TOOLS_TEXT
              ? text
              : cancelled
                ? ''
                : runLastTextRef.current || toolSummaries || text
          // A final empty turn must not claim completion: reuse the model's last
          // streamed text; with none, only a mutating run gets the "done" phrasing.
          const fallback = cancelled
            ? t('appAiStopped')
            : runLastTextRef.current ||
              toolSummaries ||
              (runMutatedRef.current ? t('appAiNoSummary') : t('appAiNoAction'))
          const finalText = turnLimit
            ? [prose, t('appAiTurnLimit')].filter(Boolean).join('\n\n')
            : prose || fallback
          setMessage(finalText)
          patchLastAssistant((entry) => ({
            ...entry,
            text: finalText,
            streaming: false,
            isError: false,
            // A stop mid-tool can leave a running placeholder behind — drop it
            tools: entry.tools.filter((tl) => !tl.running),
          }))
          // Persist the assistant message (side effect outside the updater;
          // tools stores the run's complete activity)
          if (!cancelled && finalText) {
            persistChatMessage('assistant', finalText, runToolsRef.current)
          }
          void autoSaveCompletedAiRun().finally(() => setAiBusy(false))
        },
        onError: (error) => {
          setMessage(error)
          setChat((previous) => {
            const next = [...previous]
            // the loop rolled this run's user message out of the model context — surface that
            for (let i = next.length - 1; i >= 0; i--) {
              const entry = next[i]!
              if (entry.role === 'user') {
                next[i] = { ...entry, undelivered: true }
                break
              }
            }
            const last = next.at(-1)
            if (last?.role === 'assistant') {
              next[next.length - 1] = {
                ...last,
                text: error,
                isError: true,
                streaming: false,
                tools: last.tools.filter((tl) => !tl.running),
              }
            }
            return next
          })
          // Signed-out failures get an inline sign-in button; detected via
          // gsk status rather than matching the localized error text
          void window.desktopApi
            .aiGskStatus()
            .then((status) => {
              if (status.loggedIn) return
              setChat((previous) => {
                const next = [...previous]
                const last = next.at(-1)
                if (last?.role === 'assistant' && last.isError) {
                  next[next.length - 1] = { ...last, loginRequired: true }
                }
                return next
              })
            })
            .catch(() => {})
          void autoSaveCompletedAiRun().finally(() => setAiBusy(false))
        },
      },
    })
  }

  function isAgentConfigured(): boolean {
    const settings = aiSettingsRef.current
    if (!settings) return false
    const config = settings.providers[settings.provider]
    if (!config?.model) return false
    // Genspark's key never lands in the settings file; the main process injects
    // it from the gsk login state. When logged out, requests return an error
    // guiding sign-in — not intercepted here.
    return settings.provider === 'genspark' || !!config.apiKey
  }

  /** Image attachments read as base64 and sent multimodal with this user message
   * (≤5MB each, max 20; same structure as docs/slides) */
  const MAX_IMAGES_PER_MESSAGE = 20
  async function collectImageAttachments(): Promise<AgentImage[]> {
    const imageAtts = attachmentsRef.current.filter((a) => ATTACHMENT_IMAGE_EXTS.has(a.ext))
    const images: AgentImage[] = []
    const failures: string[] = []
    for (const att of imageAtts.slice(0, MAX_IMAGES_PER_MESSAGE)) {
      const result = await window.desktopApi.readAttachmentImage(att.path)
      if (result.ok && result.base64 && result.mime) {
        images.push({ base64: result.base64, mime: result.mime })
      } else {
        failures.push(result.error ?? t('appAttachmentReadFailed', { name: att.name }))
      }
    }
    if (imageAtts.length > MAX_IMAGES_PER_MESSAGE) {
      failures.push(t('appTooManyImages', { max: MAX_IMAGES_PER_MESSAGE }))
    }
    if (failures.length > 0) {
      setAttachNotice(failures.join('；'))
      window.setTimeout(() => setAttachNotice(null), 5000)
    }
    return images
  }

  function runAgent(instruction: string): void {
    const loop = agentLoopRef.current
    if (!instruction.trim() || !loop || loop.busy || runStartingRef.current) return
    runStartingRef.current = true
    aiApplyPromisesRef.current = []
    runLastTextRef.current = ''
    runMutatedRef.current = false
    setAiBusy(true)
    setMessage(t('appAiThinking'))
    appendChat({ role: 'assistant', text: '', tools: [], streaming: true })
    void collectImageAttachments()
      .then((images) => {
        runStartingRef.current = false
        loop.run(instruction, images)
      })
      .catch(() => {
        runStartingRef.current = false
        loop.run(instruction)
      })
  }

  const mergeAttachments = (result: AttachmentAddResult | null): void => {
    if (!result) return
    if (result.accepted.length > 0) {
      setAttachments((prev) => {
        const seen = new Set(prev.map((a) => a.path))
        return [...prev, ...result.accepted.filter((a) => !seen.has(a.path))]
      })
    }
    if (result.rejected.length > 0) {
      setAttachNotice(result.rejected.join('；'))
      window.setTimeout(() => setAttachNotice(null), 5000)
    }
  }

  async function handlePickAttachments(): Promise<void> {
    mergeAttachments(await window.desktopApi.pickAttachments())
  }

  async function handleAddAttachmentPaths(paths: readonly string[]): Promise<void> {
    if (paths.length === 0) return
    mergeAttachments(await window.desktopApi.addAttachmentPaths([...paths]))
  }

  async function handleAddPastedImage(data: ArrayBuffer, ext: string): Promise<void> {
    mergeAttachments(await window.desktopApi.addPastedImage(data, ext))
  }

  function handleRemoveAttachment(path: string): void {
    setAttachments((prev) => prev.filter((a) => a.path !== path))
  }

  function handleStopAgent(): void {
    agentLoopRef.current?.cancel()
  }

  function handleNewChat(): void {
    agentLoopRef.current?.reset()
    setAiBusy(false)
    setChat([])
    setHistoricChat([])
    setPreview(null)
    lazyPreviewRef.current = null
    setMessage(t('appNewConversation'))
  }

  /** DSL context the AgentSkill reads/writes through — reuses the exact same
   * preview-then-apply path handlePlan/handleLazyPlan already exercise. */
  /** App-scope refs bundle for the extracted workbook readers (ai/workbook-readers.ts). */
  function readContext(): WorkbookReadContext {
    return { univerRef, lazyWorkbookRef, adapterRef }
  }

  function getActiveSheetInfo(): ActiveSheetInfo {
    return getActiveSheetInfoImpl(readContext())
  }

  function sheetsSkillDeps(): SheetsSkillDeps {
    return {
      getActiveSheetInfo,
      ensureRangeLoaded: async (range) => {
        const state = lazyWorkbookRef.current
        if (!state) return true
        const runtime = univerRef.current
        const worksheet = runtime?.univerAPI.getActiveWorkbook()?.getActiveSheet()
        if (!runtime || !worksheet) return false
        return ensureLazyRangeLoaded(runtime, lazyWorkbookRef, worksheet, range, setMessage)
      },
      readCells: (addresses) => readCellsImpl(readContext(), addresses),
      readFormats: (addresses) => readFormatsImpl(readContext(), addresses),
      readSheetFeatures: (sheetId) => readSheetFeaturesImpl(readContext(), sheetId),
      proposeOperations,
    }
  }

  useEffect(() => {
    void window.desktopApi.getAiSettings().then(setAiSettingsState)
  }, [])

  useEffect(() => {
    const runtime = createUniver({
      // green selection/highlight instead of Univer's default blue
      theme: greenTheme,
      locale: LocaleType.EN_US,
      locales: {
        [LocaleType.EN_US]: mergeLocales(
          UniverPresetSheetsCoreEnUS,
          UniverPresetSheetsConditionalFormattingEnUS,
          UniverPresetSheetsFilterEnUS,
          UniverPresetSheetsDataValidationEnUS,
          UniverPresetSheetsNoteEnUS,
          UniverPresetSheetsFindReplaceEnUS,
          UniverPresetSheetsSortEnUS,
          UniverPresetSheetsTableEnUS,
          // sheets-ui code in 0.25.1 references these two keys, but the language
          // pack shipped without the entries — unless patched, the raw
          // "sheets-ui.info.forceStringInfo" pops up for users.
          // mergeLocales shallow-merges namespaces, so the existing entries must
          // be spread; otherwise the whole sheets-ui namespace gets overwritten
          // (the sheet-tab context menu turns into bare keys).
          {
            'sheets-ui': {
              ...(UniverPresetSheetsCoreEnUS as Record<string, Record<string, unknown>>)[
                'sheets-ui'
              ],
              info: {
                ...(
                  UniverPresetSheetsCoreEnUS as Record<
                    string,
                    Record<string, Record<string, string>>
                  >
                )['sheets-ui']?.info,
                error: 'Number stored as text',
                forceStringInfo:
                  'The value in this cell is stored as text — it will not be treated as a ' +
                  'number in formulas.',
              },
            },
          },
        ),
      },
      presets: [
        UniverSheetsCorePreset({
          container: 'univer-container',
          // header: true + toolbar: false renders only the name box + formula
          // bar (the Univer ribbon needs both flags).
          header: true,
          toolbar: false,
          contextMenu: true,
          formulaBar: true,
          footer: {
            sheetBar: true,
            statisticBar: true,
            menus: true,
            // zoom lives in the custom full-width status bar (unified with docs/slides)
            zoomSlider: false,
          },
          statusBarStatistic: true,
        }),
        UniverSheetsDrawingPreset(),
        UniverSheetsConditionalFormattingPreset(),
        UniverSheetsFilterPreset(),
        UniverSheetsDataValidationPreset(),
        UniverSheetsNotePreset(),
        UniverSheetsFindReplacePreset(),
        UniverSheetsSortPreset(),
        UniverSheetsTablePreset(),
      ],
    })
    loadSnapshotIntoUniver(runtime, initialSnapshot, 'new-workbook', 'Untitled')
    univerRef.current = runtime
    // The window always starts blank now; still consume the one-shot
    // new-blank flag so it doesn't leak into the next workbook open.
    void window.desktopApi?.consumeNewBlankWorkbook?.()
    // Pull any shell-queued workbook ourselves: the shell's 'open' nudge loop
    // gives up after 30s, and on slow dev cold starts Univer mounts later than
    // that — the tab would strand as a blank in-memory workbook (no save, no
    // shapes) with the queued file silently never opened.
    void window.desktopApi?.hasQueuedWorkbook?.().then((queued) => {
      if (queued) void handleInspectWorkbook()
    })
    // Univer 0.25.1 also badges text parseable as date/time, phone numbers, and
    // other long numeric identifiers with "Number stored as text". Those values
    // should remain text, so clear the view type before the built-in marker
    // interceptor (priority 10). Short numeric text ("007", "20%") keeps its
    // warning.
    const dateTextDisposable = runtime.univer
      .__getInjector()
      .get(SheetInterceptorService)
      .intercept(INTERCEPTOR_POINT.CELL_CONTENT, {
        priority: 11,
        effect: InterceptorEffectEnum.Style,
        handler: (cell, _position, next) => {
          if (cell?.t === CellValueType.STRING && typeof cell.v === 'string') {
            if (isNumericIdentifierText(cell.v)) return next({ ...cell, t: undefined })
            if (isRealNum(cell.v)) return next(cell)
            const parsed = getNumfmtParseValueFilter(cell.v)
            if (parsed?.z && /[ymdhs]/i.test(parsed.z)) return next({ ...cell, t: undefined })
          }
          return next(cell)
        },
      })
    // Copying in a filtered sheet must skip hidden rows;
    // see filtered-copy.ts for why the built-in hook is not enough.
    const filteredCopyDisposable = installFilteredCopyHook(runtime)
    // Excel-compatible TSV plain text (TRUE/FALSE, quoted newlines).
    const tsvClipboardDisposable = installTsvClipboardFix(runtime)
    // Formula view: swap formula cells to their formula text per sheet.
    const formulaViewDisposable = installFormulaViewInterceptor(runtime, lazyWorkbookRef)
    // Formula bar shows harvested formula text on streamed workbooks whose
    // closure gave up; display-only, the engine never sees it.
    const formulaTextDisposable = installFormulaTextInterceptor(runtime, lazyWorkbookRef)
    // Excel-parity number-format display: empty sections, text section,
    // _/* padding, General digit fitting.
    const numberFormatFixDisposable = installNumberFormatFix(runtime)
    // CELL("filename") resolves the session's on-disk path; converted
    // imports (needsSaveAs) count as never-saved, like Excel.
    const cellFilenameDisposable = installCellFilenameFunction(runtime, () => {
      const file = lazyWorkbookRef.current?.file
      return file && !file.needsSaveAs ? (file.path ?? null) : null
    })
    // RATE converges near -100% via bisection instead of erroring.
    const rateFallbackDisposable = installRateFallback(runtime)
    // Escaped quotes ("") no longer shift lexer indices and silently
    // rewrite committed formulas.
    const formulaLexerFixDisposable = installFormulaLexerFix(runtime)
    // Renaming a sheet to a case variant of itself is not a duplicate.
    const sheetRenameFixDisposable = installSheetRenameFix()
    // Arrow keys stop at the sheet edge instead of wrapping to the far side.
    const selectionWrapGuardDisposable = installSelectionWrapGuard(runtime)
    // Empty-value formula results (IFERROR/IF/CHOOSE over blank refs)
    // display as 0 like Excel.
    const nullResultDisposable = installFormulaNullResultFix(runtime)
    // Copy/cut load their selection into the lazy window first so streamed
    // workbooks don't serialize blanks for never-viewed rows.
    const copyMaterializeDisposable = installCopyMaterialize(runtime, lazyWorkbookRef, setMessage)
    // Univer's own UI (rule-management panels, dialogs) follows the app
    // language instead of hard-coded English.
    void applyUniverLocale(runtime, getLang())
    // Rule-management panels show what each rule actually does: list options /
    // source range, CF formula text, ⚠ on #REF! dead rules.
    const ruleDetailDisposable = installRuleDetail(runtime)
    const scrollDisposable = runtime.univerAPI.addEvent(
      runtime.univerAPI.Event.Scroll,
      (params) => {
        const { worksheet } = params
        // The event carries the true post-scroll position; getVisibleRange
        // inside loadVisibleRange lags a frame.
        const eventStart = params as { sheetViewStartRow?: number; sheetViewStartColumn?: number }
        void loadVisibleRange(
          runtime,
          lazyWorkbookRef,
          worksheet,
          setMessage,
          typeof eventStart.sheetViewStartRow === 'number' &&
            typeof eventStart.sheetViewStartColumn === 'number'
            ? { row: eventStart.sheetViewStartRow, column: eventStart.sheetViewStartColumn }
            : undefined,
        )
        let visible: ReturnType<typeof worksheet.getVisibleRange>
        try {
          visible = worksheet.getVisibleRange()
        } catch {
          // The lazy loader falls back to the top-left viewport while Univer
          // replaces its scroll controller; visual installation can wait.
          return
        }
        const viewportKey = visible
          ? `${worksheet.getSheetId()}:${visible.startRow}:${visible.endRow}:${visible.startColumn}:${visible.endColumn}`
          : worksheet.getSheetId()
        if (visualViewportKeyRef.current === viewportKey) return
        visualViewportKeyRef.current = viewportKey
        queueVisualInstall(
          runtime,
          lazyWorkbookRef,
          visualDisposablesRef,
          visualInstallTimerRef,
          worksheet.getSheetId(),
          chartEditRef,
          chartVectorRef,
          shapeEditRef,
        )
        queueSparklineInstall(
          runtime,
          lazyWorkbookRef,
          sparklineDisposablesRef,
          sparklineTimerRef,
          worksheet.getSheetId(),
        )
      },
    )
    const zoomDisposable = runtime.univerAPI.addEvent(
      runtime.univerAPI.Event.SheetZoomChanged,
      ({ worksheet }) => {
        setZoomPercent(Math.round(worksheet.getZoom() * 100))
      },
    )
    // In-cell editor open/closed, read by the AutoSave tick: saving reloads
    // the workbook and would wipe an in-progress edit.
    const editStartDisposable = runtime.univerAPI.addEvent(
      runtime.univerAPI.Event.SheetEditStarted,
      () => {
        editingCellRef.current = true
      },
    )
    const editEndDisposable = runtime.univerAPI.addEvent(
      runtime.univerAPI.Event.SheetEditEnded,
      () => {
        editingCellRef.current = false
      },
    )
    const sheetDisposable = runtime.univerAPI.addEvent(
      runtime.univerAPI.Event.ActiveSheetChanged,
      ({ activeSheet }) => {
        void loadVisibleRange(runtime, lazyWorkbookRef, activeSheet, setMessage)
        // formula view is per-sheet (sheetView/@showFormulas)
        applyShowFormulasView(runtime, lazyWorkbookRef.current, activeSheet.getSheetId())
        // zoom is per-sheet state; echo the new sheet's level
        setZoomPercent(Math.round(activeSheet.getZoom() * 100))
        refreshSelectionFormatRef.current()
        visualViewportKeyRef.current = ''
        if (!lazyWorkbookRef.current) {
          queueDemoVisualInstall(runtime, activeSheet.getSheetId())
        }
        queueVisualInstall(
          runtime,
          lazyWorkbookRef,
          visualDisposablesRef,
          visualInstallTimerRef,
          activeSheet.getSheetId(),
          chartEditRef,
          chartVectorRef,
          shapeEditRef,
        )
        queueSparklineInstall(
          runtime,
          lazyWorkbookRef,
          sparklineDisposablesRef,
          sparklineTimerRef,
          activeSheet.getSheetId(),
        )
      },
    )
    const editDisposable = runtime.univerAPI.addEvent(
      runtime.univerAPI.Event.BeforeSheetEditStart,
      (event) => {
        const state = lazyWorkbookRef.current
        if (!state) return
        const sheetId = event.worksheet.getSheetId()
        const sheet = state.file.sheets.find((candidate) => candidate.id === sheetId)
        if (!sheet) return
        // Pivot output is baked into the worksheet; editing it would corrupt
        // the file's pivot semantics. Protected in every load mode.
        if (
          sheet.pivotRanges.some(
            (range) =>
              event.row >= range.startRow &&
              event.row <= range.endRow &&
              event.column >= range.startColumn &&
              event.column <= range.endColumn,
          )
        ) {
          event.cancel = true
          setMessage(t('appPivotCellNoEdit'))
          return
        }
        // Fully-loaded workbooks have nothing left to stream in.
        if (state.flags.preloadComplete) return
        // Editing a cell whose original content hasn't streamed in yet would
        // silently overwrite data the user never saw. Beyond the file's used
        // range every cell is genuinely empty, so those edits are safe — and
        // so are rows/columns inserted this session (journal-owned, nothing
        // streams into them). Bounds are screen-space: structural ops shift
        // the data extent.
        const ops = state.editJournal.structuralOps.get(sheetId) ?? []
        const beyondData =
          event.row >= sheet.rowCount + netAxisDelta(ops, 'row') ||
          event.column >= sheet.columnCount + netAxisDelta(ops, 'column')
        const journalOwned =
          ops.length > 0 &&
          (screenToFile(ops, 'row', event.row) === null ||
            screenToFile(ops, 'column', event.column) === null)
        const loaded = state.loadedRanges.get(sheetId)
        const inLoaded =
          loaded !== undefined &&
          event.row >= loaded.startRow &&
          event.row <= loaded.endRow &&
          event.column >= loaded.startColumn &&
          event.column <= loaded.endColumn
        const inFrozen =
          loaded !== undefined &&
          (event.row < (sheet.freeze?.frozenRows ?? 0) ||
            event.column < (sheet.freeze?.frozenColumns ?? 0))
        if (!beyondData && !journalOwned && !inLoaded && !inFrozen) {
          event.cancel = true
          setMessage(t('appAreaStreaming'))
        }
      },
    )
    const journalDisposable = runtime.univerAPI.addEvent(
      runtime.univerAPI.Event.CommandExecuted,
      (event) => {
        if (journalSuppression.active) return
        // The formula engine re-applies cached results with these execution
        // options; they are derived state, never user edits.
        const options = event.options as { fromFormula?: boolean } | undefined
        if (options?.fromFormula) return
        // The copy finished (or failed); a stale source must not claim a
        // later, unrelated insert-sheet.
        if (event.id === COPY_SHEET_COMMAND) {
          pendingCopySource = undefined
          return
        }
        const rowColumn = ROW_COLUMN_MUTATIONS[event.id]
        const merge = MERGE_MUTATIONS[event.id]
        const axisAttr = AXIS_ATTR_MUTATIONS[event.id]
        if (
          event.id !== SET_RANGE_VALUES_MUTATION &&
          event.id !== SET_NUMFMT_MUTATION &&
          !rowColumn &&
          !merge &&
          !axisAttr &&
          !SHEET_LIFECYCLE_MUTATIONS.has(event.id) &&
          event.id !== REORDER_RANGE_MUTATION &&
          !FILTER_MUTATIONS.has(event.id) &&
          !CF_MUTATIONS.has(event.id) &&
          !DV_MUTATIONS.has(event.id) &&
          !DEFINED_NAME_MUTATIONS.has(event.id) &&
          !NOTE_MUTATIONS.has(event.id) &&
          event.id !== MOVE_RANGE_MUTATION &&
          event.id !== MOVE_ROWS_MUTATION &&
          event.id !== SET_FROZEN_MUTATION &&
          event.id !== TOGGLE_GRIDLINES_MUTATION
        ) {
          return
        }
        const state = lazyWorkbookRef.current
        if (!state) {
          // Demo mode journals nothing, but chart↔data sync still applies.
          if (event.id === SET_RANGE_VALUES_MUTATION) {
            const demoParams = event.params as
              { subUnitId?: string; cellValue?: unknown } | undefined
            const bounds = cellValueBounds(demoParams?.cellValue)
            if (demoParams?.subUnitId && bounds) queueChartDataSync(demoParams.subUnitId, bounds)
          }
          return
        }
        const params = event.params as
          | {
              unitId?: string
              subUnitId?: string
              cellValue?: unknown
              range?: IRange
              ranges?: IRange[]
              name?: string
              sheet?: { id?: string; name?: string }
            }
          | undefined
        if (params?.unitId !== `file-${state.file.sha256}`) return
        if (SHEET_LIFECYCLE_MUTATIONS.has(event.id)) {
          if (event.id === 'sheet.mutation.insert-sheet') {
            const { id, name } = params.sheet ?? {}
            if (typeof id === 'string' && typeof name === 'string') {
              if (pendingCopySource !== undefined) {
                recordSheetDuplicate(state.editJournal, id, name, pendingCopySource)
                pendingCopySource = undefined
              } else {
                recordSheetInsert(state.editJournal, id, name)
              }
            }
          } else if (event.id === 'sheet.mutation.remove-sheet') {
            if (typeof params.subUnitId === 'string') {
              recordSheetRemove(state.editJournal, params.subUnitId)
            }
          } else if (event.id === 'sheet.mutation.set-worksheet-order') {
            recordSheetOrderChange(state.editJournal)
          } else if (event.id === 'sheet.mutation.set-worksheet-hidden') {
            const hidden = (params as { hidden?: number | boolean }).hidden
            if (typeof params.subUnitId === 'string' && hidden !== undefined) {
              const originallyHidden =
                state.file.sheets.find((sheet) => sheet.id === params.subUnitId)?.hidden ?? false
              recordSheetHidden(
                state.editJournal,
                params.subUnitId,
                hidden === true || hidden === 1,
                originallyHidden,
              )
            }
          } else if (typeof params.subUnitId === 'string' && typeof params.name === 'string') {
            const originalName = state.file.sheets.find(
              (sheet) => sheet.id === params.subUnitId,
            )?.name
            // The live sync matches series refs against live sheet names, so
            // in-memory refs must follow the rename (the file's own c:f refs
            // are rewritten independently at save time).
            const previousName =
              state.editJournal.sheets.renamed.get(params.subUnitId) ??
              state.editJournal.sheets.added.get(params.subUnitId)?.name ??
              originalName
            recordSheetRename(state.editJournal, params.subUnitId, params.name, originalName)
            if (previousName !== undefined && previousName !== params.name) {
              renameChartRefsForSheet(state, previousName, params.name)
            }
          }
          setPendingEdits(journalSize(state.editJournal))
          return
        }
        // Move-range carries its sheet ids inside from/to, not at top level.
        if (event.id === MOVE_RANGE_MUTATION) {
          const move = event.params as
            | {
                from?: { subUnitId?: string; value?: unknown }
                to?: { subUnitId?: string; value?: unknown }
                fromRange?: IRange
                toRange?: IRange
              }
            | undefined
          const fromSheet = move?.from?.subUnitId ?? params.subUnitId
          const toSheet = move?.to?.subUnitId ?? params.subUnitId
          const fromRange = move?.fromRange ?? matrixBounds(move?.from?.value)
          const toRange = move?.toRange ?? matrixBounds(move?.to?.value)
          if (fromSheet && fromRange) journalRangeSnapshot(runtime, state, fromSheet, fromRange)
          if (toSheet && toRange) journalRangeSnapshot(runtime, state, toSheet, toRange)
          // Moved cells feed charts too, same as value mutations.
          if (fromSheet && fromRange) queueChartDataSync(fromSheet, fromRange)
          if (toSheet && toRange) queueChartDataSync(toSheet, toRange)
          setPendingEdits(journalSize(state.editJournal))
          return
        }
        // Workbook-level: defined-name mutations carry no subUnitId.
        if (DEFINED_NAME_MUTATIONS.has(event.id)) {
          recordDefinedNamesChange(state.editJournal)
          setPendingEdits(journalSize(state.editJournal))
          return
        }
        if (!params.subUnitId) return
        if (axisAttr) {
          const attrParams = event.params as {
            ranges?: IRange[]
            rowHeight?: number | Record<number, number>
            colWidth?: number | Record<number, number>
            autoHeightInfo?: number | Record<number, number>
          }
          const uniform = axisAttr.axis === 'row' ? attrParams.rowHeight : attrParams.colWidth
          const toFileSize = (pixels: number): number =>
            axisAttr.axis === 'row'
              ? Math.round(pixels * 0.75 * 100) / 100
              : pixelsToCharacterWidth(pixels)
          for (const range of attrParams.ranges ?? []) {
            const start = axisAttr.axis === 'row' ? range.startRow : range.startColumn
            const end = axisAttr.axis === 'row' ? range.endRow : range.endColumn
            if (!Number.isInteger(start) || !Number.isInteger(end) || end < start) continue
            if (end - start >= 100_000) continue
            const sizeKind = axisAttr.axis === 'row' ? 'set-row-size' : 'set-col-size'
            if (axisAttr.kind === 'hidden') {
              recordStructuralOp(state.editJournal, params.subUnitId, {
                kind: axisAttr.axis === 'row' ? 'set-rows-hidden' : 'set-cols-hidden',
                start,
                end,
                hidden: axisAttr.hidden === true,
              })
            } else if (axisAttr.kind === 'auto-size') {
              // Setting an explicit height ALSO emits this mutation with
              // autoHeightInfo=0 (auto off) — only auto ON resets the height.
              const info = attrParams.autoHeightInfo
              if (typeof info === 'number') {
                if (info === 1) {
                  recordStructuralOp(state.editJournal, params.subUnitId, {
                    kind: sizeKind,
                    start,
                    end,
                    size: null,
                  })
                }
              } else if (info && typeof info === 'object') {
                for (let line = start; line <= end; line += 1) {
                  if (info[line] !== 1) continue
                  recordStructuralOp(state.editJournal, params.subUnitId, {
                    kind: sizeKind,
                    start: line,
                    end: line,
                    size: null,
                  })
                }
              }
            } else if (typeof uniform === 'number') {
              recordStructuralOp(state.editJournal, params.subUnitId, {
                kind: sizeKind,
                start,
                end,
                size: toFileSize(uniform),
              })
            } else if (uniform && typeof uniform === 'object') {
              // Per-line sizes (undo restores): one op per line in the range.
              for (let line = start; line <= end; line += 1) {
                const pixels = uniform[line]
                if (typeof pixels !== 'number') continue
                recordStructuralOp(state.editJournal, params.subUnitId, {
                  kind: sizeKind,
                  start: line,
                  end: line,
                  size: toFileSize(pixels),
                })
              }
            }
          }
          setPendingEdits(journalSize(state.editJournal))
          return
        }
        if (FILTER_MUTATIONS.has(event.id)) {
          recordFilterChange(state.editJournal, params.subUnitId)
          setPendingEdits(journalSize(state.editJournal))
          return
        }
        if (CF_MUTATIONS.has(event.id)) {
          recordCfChange(state.editJournal, params.subUnitId)
          setPendingEdits(journalSize(state.editJournal))
          return
        }
        if (DV_MUTATIONS.has(event.id)) {
          if (params.subUnitId) {
            recordDvChange(state.editJournal, params.subUnitId)
            setPendingEdits(journalSize(state.editJournal))
          }
          return
        }
        if (event.id === SET_FROZEN_MUTATION) {
          // Recording from the mutation (not the ribbon handler) keeps the
          // journal in step with Univer's undo/redo of the freeze.
          const freeze = event.params as
            { subUnitId?: string; ySplit?: number; xSplit?: number } | undefined
          if (freeze?.subUnitId && !isSheetRemoved(state.editJournal, freeze.subUnitId)) {
            recordPageSetup(state.editJournal, freeze.subUnitId, {
              frozenRows: Math.max(0, freeze.ySplit ?? 0),
              frozenColumns: Math.max(0, freeze.xSplit ?? 0),
            })
            setPendingEdits(journalSize(state.editJournal))
          }
          return
        }
        if (event.id === TOGGLE_GRIDLINES_MUTATION) {
          const gridlines = event.params as
            { subUnitId?: string; showGridlines?: number } | undefined
          if (
            gridlines?.subUnitId &&
            gridlines.showGridlines !== undefined &&
            !isSheetRemoved(state.editJournal, gridlines.subUnitId)
          ) {
            recordPageSetup(state.editJournal, gridlines.subUnitId, {
              showGridlines: gridlines.showGridlines === 1,
            })
            setPendingEdits(journalSize(state.editJournal))
          }
          return
        }
        if (NOTE_MUTATIONS.has(event.id)) {
          if (params.subUnitId) {
            recordNoteChange(state.editJournal, params.subUnitId)
            setPendingEdits(journalSize(state.editJournal))
          }
          return
        }
        if (event.id === REORDER_RANGE_MUTATION) {
          if (params.range) {
            journalRangeSnapshot(runtime, state, params.subUnitId, params.range)
            // Sorted cells feed charts too, same as value mutations.
            queueChartDataSync(params.subUnitId, params.range)
            setPendingEdits(journalSize(state.editJournal))
          }
          return
        }
        if (rowColumn || event.id === MOVE_ROWS_MUTATION) {
          let structuralOp: StructuralJournalOp
          if (rowColumn) {
            const range = params.range
            if (!range) return
            const index = rowColumn.axis === 'row' ? range.startRow : range.startColumn
            const count =
              rowColumn.axis === 'row'
                ? range.endRow - range.startRow + 1
                : range.endColumn - range.startColumn + 1
            if (count <= 0) return
            structuralOp = { kind: rowColumn.kind, index, count }
          } else {
            const move = event.params as { sourceRange?: IRange; targetRange?: IRange } | undefined
            if (!move?.sourceRange || !move.targetRange) return
            const index = move.sourceRange.startRow
            const count = move.sourceRange.endRow - move.sourceRange.startRow + 1
            const before = move.targetRange.startRow
            if (count <= 0 || (before >= index && before <= index + count)) return
            structuralOp = { kind: 'move-rows', index, count, before }
          }
          const structuralSheetId = params.subUnitId
          // Refs are matched by live sheet name (they follow renames).
          const structuralSheetName =
            runtime.univerAPI
              .getActiveWorkbook()
              ?.getSheetBySheetId(structuralSheetId)
              ?.getSheetName() ??
            state.file.sheets.find((sheet) => sheet.id === structuralSheetId)?.name
          recordStructuralOp(
            state.editJournal,
            structuralSheetId,
            structuralOp,
            structuralSheetName,
          )
          // File visuals shift on-screen too (the save shifts the file's own
          // anchors and c:f refs independently); keeping the in-memory copy in
          // the new space keeps the preview and the live data sync honest.
          state.file.visuals.forEach((visual, at) => {
            state.file.visuals[at] = shiftVisualForStructuralOp(
              visual,
              structuralSheetId,
              structuralSheetName,
              structuralOp,
            )
          })
          refreshLazyVisuals(state)
          // Univer shifted its installed cells itself, but the loaded-range
          // bookkeeping and frozen strip are now stale — refetch the viewport
          // through the updated coordinate mapping. Moves are exempt: they
          // are gated to fully loaded sheets, and the refetch would re-install
          // cells through undoable commands, burying the move's undo entry.
          if (structuralOp.kind !== 'move-rows') {
            state.loadedRanges.delete(params.subUnitId)
            state.frozenStripKeys.delete(params.subUnitId)
          }
          // Pinned closure values shift with the model; pinned formulas are
          // dropped — Univer rewrote their references in the model, so a
          // stale snapshot must not be re-applied after eviction.
          const pinnedClosure = state.closure.pinned.get(params.subUnitId)
          if (pinnedClosure && 'index' in structuralOp) {
            const shifted = shiftPinnedCells(pinnedClosure, structuralOp)
            for (const [key, cell] of [...shifted]) {
              if (cell.f !== undefined) shifted.delete(key)
            }
            state.closure.pinned.set(params.subUnitId, shifted)
          }
          // The recalc fallback reads the on-disk file; structural edits
          // desync every coordinate, so its overlays must not re-apply.
          state.recalc.overlay.clear()
          state.recalc.formulaCells.clear()
          const activeSheet = runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()
          if (activeSheet?.getSheetId() === params.subUnitId) {
            void loadVisibleRange(runtime, lazyWorkbookRef, activeSheet, setMessage)
          }
          setPendingEdits(journalSize(state.editJournal))
          return
        }
        if (merge) {
          for (const range of params.ranges ?? []) {
            if (range.endRow < range.startRow || range.endColumn < range.startColumn) continue
            recordStructuralOp(state.editJournal, params.subUnitId, {
              kind: merge,
              range: {
                startRow: range.startRow,
                endRow: range.endRow,
                startColumn: range.startColumn,
                endColumn: range.endColumn,
              },
            })
          }
          setPendingEdits(journalSize(state.editJournal))
          return
        }
        // Copy-sheet batches a large source's cellData into follow-up chunk
        // mutations; the save clones the worksheet part, so journaling them
        // as edits would duplicate (and re-encode) content the clone covers.
        if ((event.params as { __splitChunk__?: boolean } | undefined)?.__splitChunk__) return
        const recorded =
          event.id === SET_NUMFMT_MUTATION
            ? recordSetNumfmt(state.editJournal, params.subUnitId, params)
            : recordSetRangeValues(state.editJournal, params.subUnitId, params.cellValue)
        if (recorded.length === 0) return
        setPendingEdits(journalSize(state.editJournal))
        const contentEdited = recorded.some(
          (entry) => entry.hasValue || entry.formula !== undefined,
        )
        if (contentEdited) {
          const valueEntries = recorded.filter(
            (entry) => entry.hasValue || entry.formula !== undefined,
          )
          queueChartDataSync(params.subUnitId, {
            startRow: Math.min(...valueEntries.map((entry) => entry.row)),
            endRow: Math.max(...valueEntries.map((entry) => entry.row)),
            startColumn: Math.min(...valueEntries.map((entry) => entry.column)),
            endColumn: Math.max(...valueEntries.map((entry) => entry.column)),
          })
          // Sparkline values read live from the grid — re-render them too.
          if (
            state.editJournal.sparklineAdds.length > 0 ||
            state.file.sheets.some((sheet) => sheet.sparklines.length > 0)
          ) {
            queueSparklineInstall(
              runtime,
              lazyWorkbookRef,
              sparklineDisposablesRef,
              sparklineTimerRef,
              params.subUnitId,
            )
          }
        }
        if (
          !state.formulaMode &&
          contentEdited &&
          state.closure.status === 'unavailable' &&
          state.recalc.failures < RECALC_MAX_FAILURES
        ) {
          queueFormulaRecalc(runtime, lazyWorkbookRef, setMessage)
        } else if (
          !state.formulaMode &&
          state.closure.status !== 'active' &&
          recorded.some((entry) => entry.formula)
        ) {
          setMessage(t('appFormulaRecordedPartial'))
        }
      },
    )
    const structuralDisposable = runtime.univerAPI.addEvent(
      runtime.univerAPI.Event.BeforeCommandExecute,
      (event) => {
        const state = lazyWorkbookRef.current
        if (journalSuppression.active || !state) return
        if (CF_RULE_COMMAND_PATTERN.test(event.id)) {
          // The Univer panel offers rules base OOXML cannot hold (x14-only
          // icon sets, date-occurring, equal/notEqual average, …); block them
          // here instead of failing the whole save later.
          const rule = (
            event.params as
              | {
                  rule?: { rule?: Record<string, unknown> & { type?: string; config?: unknown } }
                }
              | undefined
          )?.rule?.rule
          if (rule?.type === 'iconSet' && !iconSetSaveable(rule.config)) {
            event.cancel = true
            setMessage(t('appIconSetUnsupported'))
            return
          }
          if (rule && cfRuleUnsaveableReason(rule) !== null) {
            event.cancel = true
            setMessage(t('appCfRuleUnsaveable'))
          }
          return
        }
        if (STRUCTURAL_EDIT_COMMAND_PATTERN.test(event.id)) {
          // Row/column inserts/removals and merges are allowed in every load
          // mode: viewport reads translate screen ↔ file coordinates through
          // the journaled operation stream (view-transform.ts), and the save
          // replays the same stream against the file. Sheets carrying pivot
          // tables are the exception — a shift would desync the baked pivot
          // output from its definition.
          const subUnitId =
            (event.params as { subUnitId?: string } | undefined)?.subUnitId ??
            runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId()
          const sheet = state.file.sheets.find((candidate) => candidate.id === subUnitId)
          if (sheet && sheet.pivotRanges.length > 0) {
            event.cancel = true
            setMessage(t('appPivotSheetNoStructural'))
          }
          return
        }
        if (
          SORT_COMMAND_PATTERN.test(event.id) ||
          FILTER_COMMAND_PATTERN.test(event.id) ||
          event.id === MOVE_RANGE_COMMAND ||
          event.id === MOVE_ROWS_COMMAND
        ) {
          const subUnitId =
            (event.params as { subUnitId?: string } | undefined)?.subUnitId ??
            runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId()
          const isAddedSheet =
            subUnitId !== undefined && state.editJournal.sheets.added.has(subUnitId)
          // Sorting, filtering, and range moves read/rewrite model content,
          // so partially streamed data would silently produce wrong results.
          if (!isAddedSheet && (!state.formulaMode || !state.flags.preloadComplete)) {
            event.cancel = true
            setMessage(t('appNeedFullLoadSort'))
            return
          }
          if (
            (event.id === MOVE_RANGE_COMMAND || event.id === MOVE_ROWS_COMMAND) &&
            state.file.sheets.find((candidate) => candidate.id === subUnitId)?.pivotRanges.length
          ) {
            event.cancel = true
            setMessage(t('appPivotSheetNoMove'))
            return
          }
          if (
            FILTER_COMMAND_PATTERN.test(event.id) &&
            subUnitId !== undefined &&
            state.filterOrigins.get(subUnitId)?.origin === 'table'
          ) {
            event.cancel = true
            setMessage(t('appTableFilterNoEdit'))
          }
          return
        }
        if (event.id === AUTO_FILL_COMMAND && !state.flags.preloadComplete) {
          const target = (event.params as { targetRange?: IRange } | undefined)?.targetRange
          const subUnitId =
            (event.params as { subUnitId?: string } | undefined)?.subUnitId ??
            runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId()
          const sheet = state.file.sheets.find((candidate) => candidate.id === subUnitId)
          const loaded = subUnitId === undefined ? undefined : state.loadedRanges.get(subUnitId)
          const ops =
            subUnitId === undefined ? [] : (state.editJournal.structuralOps.get(subUnitId) ?? [])
          const beyondRow = sheet === undefined ? 0 : sheet.rowCount + netAxisDelta(ops, 'row')
          const beyondColumn =
            sheet === undefined ? 0 : sheet.columnCount + netAxisDelta(ops, 'column')
          const covered =
            target !== undefined &&
            (target.startRow >= beyondRow ||
              target.startColumn >= beyondColumn ||
              (loaded !== undefined &&
                target.startRow >= loaded.startRow &&
                target.endRow <= loaded.endRow &&
                target.startColumn >= loaded.startColumn &&
                target.endColumn <= loaded.endColumn))
          if (!covered) {
            event.cancel = true
            setMessage(t('appAutofillStreaming'))
          }
          return
        }
        if (DV_EDIT_COMMAND_PATTERN.test(event.id)) {
          const subUnitId =
            (event.params as { subUnitId?: string } | undefined)?.subUnitId ??
            runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId()
          const isAddedSheet =
            subUnitId !== undefined && state.editJournal.sheets.added.has(subUnitId)
          // The save rewrites the whole section from Univer's model, so the
          // file's own rules must be in the model before any edit.
          if (!isAddedSheet && (subUnitId === undefined || !state.appliedDvSheets.has(subUnitId))) {
            event.cancel = true
            setMessage(t('appDvNeedsIndexed'))
          }
          return
        }
        if (event.id === COPY_SHEET_COMMAND) {
          const subUnitId =
            (event.params as { subUnitId?: string } | undefined)?.subUnitId ??
            runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId()
          const isAddedSheet =
            subUnitId !== undefined && state.editJournal.sheets.added.has(subUnitId)
          // The Univer-side copy clones the model, so a partially streamed
          // source would produce a copy with silently missing data.
          if (!isAddedSheet && (!state.formulaMode || !state.flags.preloadComplete)) {
            event.cancel = true
            setMessage(t('appDuplicateNeedsFullLoad'))
            return
          }
          const sheet = state.file.sheets.find((candidate) => candidate.id === subUnitId)
          if (sheet && sheet.pivotRanges.length > 0) {
            event.cancel = true
            setMessage(t('appPivotSheetNoDuplicate'))
            return
          }
          if (subUnitId !== undefined) pendingCopySource = subUnitId
          return
        }
        if (BLOCKED_COMMAND_PATTERN.test(event.id)) {
          event.cancel = true
          setMessage(t('appMoveRowsColsUnsaved'))
        }
      },
    )
    // File-menu accelerators (⌘O/⌘S/⇧⌘S) arrive from the main process.
    const unsubscribeMenu =
      window.desktopApi?.onMenuAction((action) => menuActionRef.current(action)) ??
      (() => undefined)
    // Close guard chose Save: run the journal save and report the outcome.
    const unsubscribeCloseSave =
      window.desktopApi?.onCloseSaveRequest?.(() => void closeSaveRef.current()) ??
      (() => undefined)
    const selectionDisposable = runtime.univerAPI.addEvent(
      runtime.univerAPI.Event.SelectionChanged,
      () => {
        refreshSelectionFormatRef.current()
        // A grid click ends any floating-visual selection.
        clearVisualSelection()
      },
    )
    // Style edits (ribbon, dialog, undo/redo, AI apply) all land as these
    // mutations; re-reading the selection keeps the ribbon echo current.
    const formatEchoDisposable = runtime.univerAPI.addEvent(
      runtime.univerAPI.Event.CommandExecuted,
      ({ id }) => {
        if (
          id === SET_RANGE_VALUES_MUTATION ||
          id === SET_NUMFMT_MUTATION ||
          id === REMOVE_NUMFMT_MUTATION
        ) {
          refreshSelectionFormatRef.current()
        }
      },
    )
    const clickDisposable = runtime.univerAPI.addEvent(
      runtime.univerAPI.Event.CellClicked,
      ({ worksheet, row, column }) => {
        const state = lazyWorkbookRef.current
        if (!state) return
        // A journaled link edit (set, changed, or removed) wins over the
        // file's streamed target.
        const journaled = hyperlinkEditAt(state.editJournal, worksheet.getSheetId(), row, column)
        const target =
          journaled !== undefined
            ? journaled
            : state.hyperlinkTargets.get(worksheet.getSheetId())?.get(`${row}:${column}`)
        if (target?.startsWith('http')) {
          void window.desktopApi.openExternal(target)
        } else if (target?.startsWith('#')) {
          navigateToAnchor(runtime, target.slice(1), setMessage)
        }
      },
    )
    // Track "any cell has content" for the ribbon's AI action buttons.
    // Mutations fire in bursts (paste, AI plans), so recompute on a short
    // trailing debounce; the scan itself early-exits on the first value.
    let contentTimer: ReturnType<typeof setTimeout> | null = null
    const contentDisposable = runtime.univerAPI.addEvent(
      runtime.univerAPI.Event.CommandExecuted,
      (event) => {
        if (!event.id.includes('mutation')) return
        if (contentTimer) clearTimeout(contentTimer)
        contentTimer = setTimeout(recomputeSheetContent, 200)
      },
    )
    return () => {
      unsubscribeMenu()
      unsubscribeCloseSave()
      dateTextDisposable.dispose()
      filteredCopyDisposable.dispose()
      tsvClipboardDisposable.dispose()
      formulaViewDisposable.dispose()
      formulaTextDisposable.dispose()
      numberFormatFixDisposable.dispose()
      cellFilenameDisposable.dispose()
      rateFallbackDisposable.dispose()
      formulaLexerFixDisposable.dispose()
      sheetRenameFixDisposable.dispose()
      selectionWrapGuardDisposable.dispose()
      nullResultDisposable.dispose()
      copyMaterializeDisposable.dispose()
      ruleDetailDisposable()
      scrollDisposable.dispose()
      zoomDisposable.dispose()
      editStartDisposable.dispose()
      editEndDisposable.dispose()
      sheetDisposable.dispose()
      editDisposable.dispose()
      journalDisposable.dispose()
      structuralDisposable.dispose()
      selectionDisposable.dispose()
      formatEchoDisposable.dispose()
      clickDisposable.dispose()
      if (contentTimer) clearTimeout(contentTimer)
      contentDisposable.dispose()
      if (visualInstallTimerRef.current) clearTimeout(visualInstallTimerRef.current)
      disposeVisuals(visualDisposablesRef.current)
      visualViewportKeyRef.current = ''
      const lazyState = lazyWorkbookRef.current
      lazyWorkbookRef.current = null
      clearLazyState(lazyState)
      if (lazyState) {
        void window.desktopApi.closeWorkbook(lazyState.file.sessionId)
      }
      runtime.univer.dispose()
      univerRef.current = null
    }
  }, [])

  function handleSend(overrideInstruction?: string): void {
    const instruction = (overrideInstruction ?? prompt).trim()
    if (!instruction || aiBusy) return
    runToolsRef.current = []
    appendChat({ role: 'user', text: instruction, tools: [] })
    persistChatMessage('user', instruction)
    if (!overrideInstruction) setPrompt('')
    // real LLM configured → let the agent read context and propose operations;
    // otherwise fall back to the local, deterministic regex planner
    // (kept for offline use and for the fixed micro-DSL it still supports).
    if (isAgentConfigured()) {
      runAgent(instruction)
      return
    }
    const outcome = runDeterministicPlan(instruction)
    setMessage(outcome.text)
    appendChat({ role: 'assistant', text: outcome.text, tools: [], isError: outcome.isError })
    persistChatMessage('assistant', outcome.text)
  }

  /// AI edits on imported workbooks preview against the live sheet, then
  /// apply through Univer commands so they enter the edit journal exactly
  /// like manual edits (and save with ⌘S).

  /// A generated workbook shouldn't keep the pristine default sheet: when an
  /// AI plan adds its own sheet(s) and never touches Sheet1, drop the empty
  /// leftover so the first sheet the user sees is the generated one.
  /// Returns the post-prune revision, or null when nothing was pruned.
  function pruneEmptyDefaultSheet(plan: ChangePlan): number | null {
    const addedSheet = plan.structuralChanges.some((change) => change.op.op === 'add_sheet')
    if (!addedSheet) return null
    const snapshot = adapterRef.current.getSnapshot()
    if (snapshot.sheets.length < 2) return null
    const defaultSheet = snapshot.sheets.find((sheet) => sheet.id === 'sheet-1')
    if (!defaultSheet || defaultSheet.name !== 'Sheet1') return null
    if (Object.keys(defaultSheet.cells).length > 0) return null
    if ((defaultSheet.visuals?.length ?? 0) > 0) return null
    const touched =
      plan.cellChanges.some((change) => change.sheetId === 'sheet-1') ||
      plan.formatChanges.some((change) => change.sheetId === 'sheet-1') ||
      plan.sheetRenames.some((rename) => rename.sheetId === 'sheet-1') ||
      plan.structuralChanges.some(
        (change) => 'sheetId' in change.op && change.op.sheetId === 'sheet-1',
      )
    if (touched) return null
    try {
      const prunePlan = adapterRef.current.plan({
        dslVersion: 1,
        transactionId: `prune-default-sheet-${crypto.randomUUID()}`,
        baseRevision: snapshot.revision,
        summary: 'Remove the empty default sheet',
        operations: [{ op: 'delete_sheet', sheetId: 'sheet-1' }],
      })
      return adapterRef.current.apply(prunePlan).revision
    } catch {
      // Best-effort cleanup: keep the empty sheet if the delete is rejected.
      return null
    }
  }

  /// Demo-mode counterpart of queueVisualInstall: charts live in the adapter
  /// snapshot, so every grid rebuild (Apply/undo) and sheet switch re-installs
  /// them from there.
  function queueDemoVisualInstall(runtime: UniverRuntime, sheetId: string): void {
    if (demoVisualInstallTimerRef.current) clearTimeout(demoVisualInstallTimerRef.current)
    demoVisualInstallTimerRef.current = setTimeout(function install() {
      demoVisualInstallTimerRef.current = null
      if (lazyWorkbookRef.current) return
      if (runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId() !== sheetId) return
      if (isVisualDragActive()) {
        demoVisualInstallTimerRef.current = setTimeout(install, 100)
        return
      }
      disposeVisuals(demoVisualDisposablesRef.current)
      const visuals = adapterRef.current
        .getSnapshot()
        .sheets.flatMap((sheet) => sheet.visuals ?? [])
      demoVisualDisposablesRef.current =
        visuals.length === 0
          ? []
          : installWorkbookVisuals(
              runtime,
              { sessionId: 'demo-workbook', visuals },
              sheetId,
              {
                edits: EMPTY_CHART_EDITS,
                onEdit: (editKey, edit) => chartEditRef.current(editKey, edit),
                readVector: (editKey, range) => chartVectorRef.current(editKey, range),
              },
              { onEdit: (visualId, changes) => shapeEditRef.current(visualId, changes) },
            )
    }, 100)
  }

  function queueDemoVisualInstallForActiveSheet(): void {
    const runtime = univerRef.current
    const sheetId = runtime?.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId()
    if (runtime && sheetId) queueDemoVisualInstall(runtime, sheetId)
  }

  /** Default worksheet names carry no content signal, so they never name the file. */
  const DEFAULT_SHEET_NAME_RE = /^(sheet|工作表|ワークシート|シート)\s*\d*$/i

  /** Waits for every plan submitted during one AI run, then persists all
   * successful writes in one save. A canceled/failed Save As leaves both the
   * journal and inline undo available. */
  async function autoSaveCompletedAiRun(): Promise<void> {
    const applies = aiApplyPromisesRef.current
    aiApplyPromisesRef.current = []
    if (applies.length === 0) return
    const results = await Promise.all(applies)
    if (!results.some(Boolean)) return
    const state = lazyWorkbookRef.current
    if (!state || journalSize(state.editJournal) === 0) return
    // AutoSave off = the user decides when the file is written: the
    // run's edits stay pending in the journal, so the offered Undo / ⌘Z keeps
    // working (saving would reopen the session and reset the undo stack).
    if (!autoSaveRef.current) {
      setMessage(t('appAiChangesNotSaved'))
      return
    }
    // AutoSave-driven write after an AI run: silent like the interval autosave.
    await handleSave('save', true)
    const after = lazyWorkbookRef.current
    if (after && journalSize(after.editJournal) === 0) {
      // Saving reopens the sidecar session and resets Univer's undo stack.
      patchLastAssistant(({ autoApplied: _autoApplied, ...entry }) => entry)
      // Sheets' analog of slides' deckName: propose the first AI-named sheet as
      // the file name. The main process no-ops unless the file still carries the
      // shell's auto-created untitled name, so user-chosen names are never touched.
      const candidate = after.file.sheets
        .map((sheet) => sheet.name.trim())
        .find((name) => name.length > 0 && !DEFAULT_SHEET_NAME_RE.test(name))
      if (candidate) {
        try {
          await window.desktopApi.autoRenameWorkbook(after.file.sessionId, candidate)
        } catch {
          // naming is best-effort; the save itself already succeeded
        }
      }
    }
  }

  /**
   * Auto-apply a just-proposed plan without the manual Apply click.
   *
   * All plans (content, format, and structural) commit immediately for a
   * smoother, Google-Sheets-like flow — AI edits share the ribbon's command
   * channel + edit journal, so undo (⌘Z / inline button) covers everything.
   *
   * The CAS/planStillMatches guards inside plan()/handleLazyApply are preserved
   * — auto-apply never bypasses the "workbook changed since preview" check.
   * When apply fails, the preview card stays up as a manual fallback.
   */
  function autoApplySafePlan(plan: ChangePlan): Promise<ApplyOutcome> {
    const opCount =
      plan.cellChanges.length +
      plan.formatChanges.length +
      plan.sheetRenames.length +
      plan.structuralChanges.length
    const state = lazyWorkbookRef.current
    if (state) {
      // Lazy path reads lazyPreviewRef (a ref, already set by the caller) —
      // safe to invoke synchronously right after propose.
      const apply = handleLazyApply(state).then((outcome) => {
        if (outcome.ok) {
          // Patch last assistant message with inline undo button.
          patchLastAssistant((entry) => ({ ...entry, autoApplied: { opCount } }))
        } else {
          // No manual-apply entry point: the failure reason is already in the
          // chat/status bar, and the preview card just collapses.
          lazyPreviewRef.current = null
          setPreview(null)
        }
        return outcome
      })
      aiApplyPromisesRef.current.push(apply.then((outcome) => outcome.ok))
      return apply
    }
    // Non-lazy path: apply the passed plan directly (setPreview is async, so we
    // cannot rely on the preview state within the same tick).
    try {
      const receipt = adapterRef.current.apply(plan)
      const prunedRevision = pruneEmptyDefaultSheet(plan)
      // Row/column shifts and new sheets can't be patched cell-by-cell into
      // the existing Univer grid — rebuild the demo workbook from the snapshot.
      if (plan.structuralChanges.length > 0 || prunedRevision !== null) {
        loadSnapshotIntoUniver(
          univerRef.current,
          adapterRef.current.getSnapshot(),
          'new-workbook',
          'Untitled',
        )
        queueDemoVisualInstallForActiveSheet()
      } else {
        syncUniver(univerRef.current, adapterRef.current.getSnapshot())
        const workbook = univerRef.current?.univerAPI.getActiveWorkbook()
        for (const formatChange of plan.formatChanges) {
          const worksheet = workbook?.getSheetBySheetId(formatChange.sheetId)
          if (worksheet)
            applyFormatPatchToRange(worksheet.getRange(formatChange.range), formatChange.format)
        }
      }
      const revision = prunedRevision ?? receipt.revision
      setRevision(revision)
      setPreview(null)
      setMessage(t('appAppliedRevision', { revision }))
      // Patch last assistant message to show inline undo button.
      patchLastAssistant((entry) => ({ ...entry, autoApplied: { opCount } }))
      return Promise.resolve({ ok: true })
    } catch (error: unknown) {
      // Fall back to leaving the preview up so the user can Apply manually.
      const reason = error instanceof Error ? error.message : t('appApplyTxFailed')
      setMessage(reason)
      return Promise.resolve({ ok: false, reason })
    }
  }

  async function handleLazyApply(state: LazyWorkbookState): Promise<ApplyOutcome> {
    const stored = lazyPreviewRef.current
    const runtime = univerRef.current
    if (!stored || !runtime) return { ok: false, reason: t('appApplyTxFailed') }
    if (stored.sessionId !== state.file.sessionId) {
      lazyPreviewRef.current = null
      setPreview(null)
      setMessage(t('appPreviewOtherWorkbook'))
      return { ok: false, reason: t('appPreviewOtherWorkbook') }
    }
    const worksheet = runtime.univerAPI.getActiveWorkbook()?.getSheetBySheetId(stored.sheetId)
    if (!worksheet) {
      lazyPreviewRef.current = null
      setPreview(null)
      setMessage(t('appPreviewSheetGone'))
      return { ok: false, reason: t('appPreviewSheetGone') }
    }
    // Image bytes load BEFORE the drift check and the (synchronous) mutation
    // loop, so a slow disk read can never interleave with edits.
    const imageData = new Map<
      string,
      { dataUrl: string; mediaType: string; width: number; height: number }
    >()
    try {
      for (const structural of stored.plan.structuralChanges) {
        if (structural.op.op !== 'add_image' || imageData.has(structural.op.path)) continue
        const image = await window.desktopApi.readLocalImage({ path: structural.op.path })
        const dataUrl = `data:${image.mediaType};base64,${image.base64}`
        const size = await measureImage(dataUrl)
        imageData.set(structural.op.path, { dataUrl, mediaType: image.mediaType, ...size })
      }
    } catch (error: unknown) {
      const reason = error instanceof Error ? error.message : t('appCannotReadImage')
      setMessage(reason)
      patchLastAssistant((entry) => ({
        ...entry,
        text: `${entry.text}\n\n${t('appApplyFailed', { reason })}`,
        isError: true,
      }))
      return { ok: false, reason }
    }
    if (lazyPreviewRef.current !== stored || lazyWorkbookRef.current !== state) {
      return { ok: false, reason: t('appApplyTxFailed') }
    }
    if (!planStillMatches(stored.plan, lazyCellReader(worksheet))) {
      const reason = t('appWorkbookChangedSincePreview')
      setMessage(reason)
      patchLastAssistant((entry) => ({
        ...entry,
        text: `${entry.text}\n\n${t('appApplyFailed', { reason })}`,
        isError: true,
      }))
      return { ok: false, reason }
    }
    // All commands of one propose merge into a single undo item (⌘Z / [Undo]
    // rolls back the whole batch in one step)
    const batchUnitId = runtime.univerAPI.getActiveWorkbook()?.getId()
    const undoBatching = batchUnitId
      ? runtime.univer.__getInjector().get(IUndoRedoService).__tempBatchingUndoRedo(batchUnitId)
      : null
    try {
      // Structural and layout changes go through the same facade commands as
      // the ribbon, so BeforeCommandExecute gating and the edit journal apply.
      const workbook = runtime.univerAPI.getActiveWorkbook()
      const sheetById = (id: string): typeof worksheet => {
        const found = workbook?.getSheetBySheetId(id)
        if (!found) throw new Error(`Unknown sheet: ${id}`)
        return found
      }
      for (const structural of stored.plan.structuralChanges) {
        const op = structural.op
        if (op.op === 'insert_rows') worksheet.insertRowsBefore(op.row - 1, op.count)
        else if (op.op === 'delete_rows') worksheet.deleteRows(op.row - 1, op.count)
        else if (op.op === 'insert_cols')
          worksheet.insertColumnsBefore(columnIndex(op.column), op.count)
        else if (op.op === 'delete_cols') worksheet.deleteColumns(columnIndex(op.column), op.count)
        else if (op.op === 'add_sheet') workbook?.insertSheet(op.name)
        else if (op.op === 'delete_sheet') workbook?.deleteSheet(op.sheetId)
        else if (op.op === 'merge_cells') worksheet.getRange(op.range).merge()
        else if (op.op === 'unmerge_cells') worksheet.getRange(op.range).breakApart()
        else if (op.op === 'set_row_height') {
          worksheet.setRowHeights(op.row - 1, op.count, Math.round((op.heightPoints * 96) / 72))
        } else if (op.op === 'set_col_width') {
          worksheet.setColumnWidths(columnIndex(op.column), op.count, Math.round(op.widthPx))
        } else if (op.op === 'set_rows_hidden') {
          if (op.hidden) sheetById(op.sheetId).hideRows(op.row - 1, op.count)
          else sheetById(op.sheetId).showRows(op.row - 1, op.count)
        } else if (op.op === 'set_cols_hidden') {
          if (op.hidden) sheetById(op.sheetId).hideColumns(columnIndex(op.column), op.count)
          else sheetById(op.sheetId).showColumns(columnIndex(op.column), op.count)
        } else if (op.op === 'duplicate_sheet') {
          if (!workbook) throw new Error(t('appNoWorkbookOpen'))
          const copy = workbook.duplicateSheet(sheetById(op.sheetId))
          if (op.name) copy.setName(op.name)
        } else if (op.op === 'set_sheet_hidden') {
          if (op.hidden) sheetById(op.sheetId).hideSheet()
          else sheetById(op.sheetId).showSheet()
        } else if (op.op === 'move_sheet') {
          if (!workbook) throw new Error(t('appNoWorkbookOpen'))
          workbook.moveSheet(sheetById(op.sheetId), op.position - 1)
        } else if (op.op === 'add_sparkline') {
          const target = workbook?.getSheetBySheetId(op.sheetId)
          if (!target) throw new Error(`Unknown sheet: ${op.sheetId}`)
          const bounds = parseRange(op.dataRange)
          const rows = Math.min(bounds.endRow - bounds.startRow + 1, 200)
          const base =
            op.targetCell === undefined
              ? { row: bounds.startRow, column: bounds.endColumn + 1 }
              : parseAddress(op.targetCell)
          const sheetName = target.getSheetName()
          const cells = Array.from({ length: rows }, (_, offset) => ({
            cell: `${columnLabel(base.column)}${base.row + offset + 1}`,
            sourceRef: absRangeRef(
              sheetName,
              `${columnLabel(bounds.startColumn)}${bounds.startRow + offset + 1}` +
                `:${columnLabel(bounds.endColumn)}${bounds.startRow + offset + 1}`,
            ),
          }))
          recordSparklineAdd(state.editJournal, {
            id: `sparkline-${Date.now().toString(36)}-${state.editJournal.sparklineAdds.length + 1}`,
            sheetId: op.sheetId,
            type: op.type,
            ...(op.color === undefined ? {} : { color: op.color }),
            cells,
          })
          setPendingEdits(journalSize(state.editJournal))
          queueSparklineInstall(
            runtime,
            lazyWorkbookRef,
            sparklineDisposablesRef,
            sparklineTimerRef,
            op.sheetId,
          )
        } else if (op.op === 'delete_visual') {
          const visual = [...state.file.visuals, ...state.editJournal.visualAdds].find(
            (candidate) => candidate.id === op.visualId || candidate.chartPath === op.visualId,
          )
          if (!visual) throw new Error(`Unknown visual: ${op.visualId}`)
          shapeEditRef.current(visual.id, { remove: true })
        } else if (op.op === 'delete_table') {
          if (!removeTableAdd(state.editJournal, op.sheetId, op.tableName)) {
            throw new Error(t('appTableNotDeletable', { name: op.tableName }))
          }
          setPendingEdits(journalSize(state.editJournal))
        } else if (op.op === 'add_chart') {
          await insertAiChartVisualImpl(visualContext(), runtime, state, op)
        } else if (op.op === 'add_shape') {
          insertAiShapeVisualImpl(visualContext(), runtime, state, op)
        } else if (op.op === 'edit_shape') {
          applyAiShapeEditImpl(visualContext(), runtime, state, op)
        } else if (op.op === 'add_image') {
          const image = imageData.get(op.path)
          if (!image) throw new Error(t('appImageNotLoaded', { path: op.path }))
          insertAiImageVisualImpl(visualContext(), runtime, state, op, image)
        } else if (op.op === 'add_table') {
          applyAiTableAdd(runtime, state, op)
        } else if (op.op === 'add_table_row') {
          applyAiTableRowAdd(runtime, state, op)
        } else if (op.op === 'add_table_column') {
          applyAiTableColumnAdd(runtime, state, op)
        } else if (op.op === 'delete_table_row') {
          applyAiTableRowDelete(runtime, state, op)
        } else if (op.op === 'delete_table_column') {
          applyAiTableColumnDelete(runtime, state, op)
        } else if (op.op === 'add_pivot') {
          applyAiPivotAdd(runtime, state, op)
        } else if (op.op === 'set_hyperlink') {
          applyAiHyperlink(state, sheetById(op.sheetId), op)
        } else if (op.op === 'protect_sheet') {
          const guard = protectSheetGuard(state, op.sheetId, op.protected)
          if (guard) throw new Error(guard)
          const original = state.sheetProtections.get(op.sheetId)?.protected ?? false
          recordSheetProtection(state.editJournal, op.sheetId, op.protected, original)
        } else if (op.op === 'set_filter') {
          const target = sheetById(op.sheetId)
          target.getFilter()?.remove()
          if (!target.getRange(op.range).createFilter()) {
            throw new Error(t('appAutoFilterCreateFailed'))
          }
        } else if (op.op === 'clear_filter') {
          sheetById(op.sheetId).getFilter()?.remove()
        } else if (op.op === 'set_filter_criteria') {
          applyFilterCriteria(
            sheetById(op.sheetId),
            op.column,
            op.values === null ? null : { values: op.values },
          )
        } else if (op.op === 'add_conditional_format') {
          applyAiConditionalFormat(sheetById(op.sheetId), op)
        } else if (op.op === 'clear_conditional_formats') {
          const target = sheetById(op.sheetId)
          for (const rule of target.getConditionalFormattingRules()) {
            if (rule.cfId) target.deleteConditionalFormattingRule(rule.cfId)
          }
        } else if (op.op === 'set_data_validation') {
          applyAiDataValidation(runtime, sheetById(op.sheetId), op)
        } else if (op.op === 'add_defined_name') {
          if (!workbook) throw new Error(t('appNoWorkbookOpen'))
          workbook.insertDefinedName(op.name, op.ref)
        } else if (op.op === 'delete_defined_name') {
          if (!workbook) throw new Error(t('appNoWorkbookOpen'))
          workbook.deleteDefinedName(op.name)
        } else if (op.op === 'set_page_setup') {
          sheetById(op.sheetId)
          const prior = state.editJournal.pageSetup.get(op.sheetId) ?? {}
          const patch: PageSetupJournalState = {}
          if (op.orientation !== undefined) patch.orientation = op.orientation
          if (op.paperSize !== undefined) patch.paperSize = op.paperSize
          if (op.margins !== undefined) patch.margins = op.margins
          if (op.printGridlines !== undefined) patch.printGridlines = op.printGridlines
          if (op.printHeadings !== undefined) patch.printHeadings = op.printHeadings
          if (op.printArea !== undefined) patch.printArea = op.printArea
          // Scale and fit-to-page are exclusive; whichever the op sets wins,
          // and a fit on one axis keeps the other axis' prior value.
          if (op.scale !== undefined) {
            patch.scale = op.scale
            patch.fitToPage = false
          } else if (op.fitToWidth !== undefined || op.fitToHeight !== undefined) {
            patch.fitToWidth = op.fitToWidth ?? prior.fitToWidth ?? 0
            patch.fitToHeight = op.fitToHeight ?? prior.fitToHeight ?? 0
            patch.fitToPage = patch.fitToWidth > 0 || patch.fitToHeight > 0
          }
          recordPageSetup(state.editJournal, op.sheetId, patch)
        } else if (op.op === 'set_freeze') {
          // Journaled by the set-frozen mutation listener.
          const target = sheetById(op.sheetId)
          if (op.rows === 0 && op.columns === 0) {
            target.cancelFreeze()
          } else {
            target.setFreeze({
              startRow: op.rows > 0 ? op.rows : -1,
              startColumn: op.columns > 0 ? op.columns : -1,
              xSplit: op.columns,
              ySplit: op.rows,
            })
          }
        } else if (op.op === 'refresh_pivot') {
          refreshPivotTablesImpl(pivotContext(), op.sheetId)
        } else if (op.op === 'set_note') {
          const target = sheetById(op.sheetId)
          const noteRange = target.getRange(op.address)
          if (op.text === null) {
            noteRange.deleteNote()
          } else {
            const cell = parseAddress(op.address)
            noteRange.createOrUpdateNote({
              id: `note-${op.sheetId}-${cell.row}-${cell.column}`,
              row: cell.row,
              col: cell.column,
              width: 220,
              height: 90,
              note: op.text,
            })
          }
        } else {
          chartEditRef.current(
            op.chartPath,
            await buildAiChartEditImpl(visualContext(), state, workbook, op),
          )
        }
      }
      setPendingEdits(journalSize(state.editJournal))
      for (const change of stored.plan.cellChanges) {
        const range = worksheet.getRange(change.address)
        if (change.after.formula) range.setFormula(change.after.formula)
        else if (change.after.value === null) range.clearContent()
        // Explicit f/si null mirrors the cell editor: overwriting a formula
        // cell with a value must clear the formula (in Univer and journal).
        else range.setValues([[{ v: change.after.value, f: null, si: null }]])
      }
      // Same facade setters as the ribbon, so the edit journal records them
      // (indent included — it lands as a pd patch in set-range-values).
      for (const formatChange of stored.plan.formatChanges) {
        applyFormatPatchToRange(worksheet.getRange(formatChange.range), formatChange.format)
      }
      for (const rename of stored.plan.sheetRenames) {
        worksheet.setName(rename.after)
      }
      lazyPreviewRef.current = null
      setPreview(null)
      setMessage(t('appAppliedJournaled'))
      return { ok: true }
    } catch (error: unknown) {
      const reason = error instanceof Error ? error.message : t('appApplyTxFailed')
      setMessage(reason)
      // The chat answer already promised the change — surface the failure
      // there too, or it silently never lands on the canvas.
      patchLastAssistant((entry) =>
        entry.text.includes(reason)
          ? entry
          : {
              ...entry,
              text: `${entry.text}\n\n${t('appApplyFailed', { reason })}`,
              isError: true,
            },
      )
      return { ok: false, reason }
    } finally {
      undoBatching?.dispose()
    }
  }

  function handleUndo(): void {
    if (lazyWorkbookRef.current) {
      void univerRef.current?.univerAPI.undo()
      return
    }
    try {
      const receipt = adapterRef.current.undo()
      // Rebuild instead of patching: undo can remove cells and reverse
      // structural changes, neither of which syncUniver can express.
      loadSnapshotIntoUniver(
        univerRef.current,
        adapterRef.current.getSnapshot(),
        'new-workbook',
        'Untitled',
      )
      queueDemoVisualInstallForActiveSheet()
      setRevision(receipt.revision)
      setPreview(null)
      setMessage(t('appUndoCommitted', { revision: receipt.revision }))
    } catch (error: unknown) {
      setMessage(error instanceof Error ? error.message : t('appUndoFailed'))
    }
  }

  /// QAT Redo: workbook history via Univer, same path as the app menu's ⇧⌘Z
  /// (the demo adapter has no redo, matching the menu's behavior).
  function handleRedo(): void {
    void univerRef.current?.univerAPI.redo()
  }

  /** App-scope refs/state bundle for the extracted ribbon dispatcher (ribbon-actions.ts). */
  function ribbonContext(): RibbonCommandContext {
    return {
      univerRef,
      lazyWorkbookRef,
      traceArrowsRef,
      sparklineDisposablesRef,
      sparklineTimerRef,
      chartEditRef,
      shapeEditRef,
      refreshSelectionFormatRef,
      selectedVisual,
      selectedChart,
      setMessage,
      setChartDialog,
      setSymbolDialogOpen,
      setScreenshotDialogOpen,
      setIconsDialogOpen,
      setEquationDialogOpen,
      openRecommendedCharts: () => {
        void handleRecommendedChartsImpl(visualContext()).then((result) => {
          if (result) setRecommendedCharts(result)
        })
      },
      setPendingEdits,
      visualContext,
      dataToolsContext,
      pivotContext,
      handlePageLayoutCommand: (rest) => handlePageLayoutCommandImpl(pageLayoutContext(), rest),
      handleExportPdf: () => handleExportPdfImpl(pageLayoutContext()),
    }
  }

  function handleRibbonCommand(command: string): void {
    handleRibbonCommandImpl(ribbonContext(), command)
  }

  function selectionStyle(
    range: NonNullable<ReturnType<ActiveWorkbook['getActiveRange']>>,
  ): IStyleData {
    // Resolves interned style references and merges row/col/sheet styles —
    // raw getCellData().s can be a style-id string with no fields on it.
    return range.getCellStyleData() ?? {}
  }

  function anchorCellValue(): number | string | null {
    try {
      const value = univerRef.current?.univerAPI.getActiveWorkbook()?.getActiveRange()?.getValue()
      return typeof value === 'number' || typeof value === 'string' ? value : null
    } catch {
      return null
    }
  }

  refreshSelectionFormatRef.current = () => {
    const range = univerRef.current?.univerAPI.getActiveWorkbook()?.getActiveRange()
    if (!range) {
      setSelectionFormat(null)
      setActiveCellA1('')
      return
    }
    setActiveCellA1(`${columnLetter(range.getColumn())}${range.getRow() + 1}`)
    let pattern: string
    try {
      pattern = range.getNumberFormat()
    } catch {
      // A disposing workbook can race the read; keep the last echo.
      return
    }
    const next = toSelectionFormat(selectionStyle(range), pattern, selectionLinkTarget(range))
    setSelectionFormat((previous) => (selectionFormatEquals(previous, next) ? previous : next))
  }

  function selectionLinkTarget(
    range: NonNullable<ReturnType<ActiveWorkbook['getActiveRange']>>,
  ): string | null {
    const state = lazyWorkbookRef.current
    const sheetId = univerRef.current?.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId()
    if (!state || !sheetId) return null
    const row = range.getRow()
    const column = range.getColumn()
    const journaled = hyperlinkEditAt(state.editJournal, sheetId, row, column)
    if (journaled !== undefined) return journaled
    return state.hyperlinkTargets.get(sheetId)?.get(`${row}:${column}`) ?? null
  }

  function openLazyWorkbook(opened: WorkbookFile): void {
    const selected: WorkbookFile = {
      ...opened,
      visuals: opened.visuals.map((visual) =>
        visual.kind === 'chart' && visual.chart !== undefined
          ? { ...visual, chart: withDefaultBarLabels(visual.chart) }
          : visual,
      ),
    }
    setWorkbookFile(selected)
    const previous = lazyWorkbookRef.current
    if (previous) {
      clearLazyState(previous)
      void window.desktopApi.closeWorkbook(previous.file.sessionId).catch(() => undefined)
    }
    if (demoVisualInstallTimerRef.current) {
      clearTimeout(demoVisualInstallTimerRef.current)
      demoVisualInstallTimerRef.current = null
    }
    disposeVisuals(demoVisualDisposablesRef.current)
    demoVisualDisposablesRef.current = []
    const state: LazyWorkbookState = {
      file: selected,
      generation: Date.now(),
      loadedRanges: new Map(),
      loadingKeys: new Map(),
      retryTimers: new Map(),
      appliedMerges: new Map(),
      appliedRowKeys: new Map(),
      sheetProtections: new Map(),
      uninstalledDefinedNames: new Set(),
      appliedCfSheets: new Set(),
      appliedFilterSheets: new Set(),
      appliedDvSheets: new Set(),
      hyperlinkTargets: new Map(),
      frozenStripKeys: new Map(),
      filterOrigins: new Map(),
      showFormulaSheets: new Set(
        selected.sheets.filter((sheet) => sheet.showFormulas).map((sheet) => sheet.id),
      ),
      formulaMode:
        selected.sheets.reduce((sum, sheet) => sum + sheet.rowCount * sheet.columnCount, 0) <=
        FORMULA_MODE_MAX_CELLS,
      editJournal: createEditJournal(),
      flags: { preloadComplete: false },
      closure: { status: 'idle', pinned: new Map() },
      formulaText: new Map(),
      pivotDefinitions: new Map(),
      outline: new Map(),
      recalc: {
        timer: null,
        generation: 0,
        failures: 0,
        formulaCells: new Map(),
        overlay: new Map(),
      },
    }
    // Column outline levels arrive with the sheet metadata; seed them now.
    for (const sheet of selected.sheets) {
      for (const columnWidth of sheet.columnWidths) {
        if (columnWidth.outlineLevel === undefined && !columnWidth.collapsed) continue
        const cols = sheetOutline(state, sheet.id).cols
        const endColumn = Math.min(columnWidth.endColumn, sheet.columnCount - 1)
        for (let column = columnWidth.startColumn; column <= endColumn; column += 1) {
          cols.set(column, {
            level: columnWidth.outlineLevel ?? 0,
            collapsed: columnWidth.collapsed ?? false,
          })
        }
      }
    }
    lazyWorkbookRef.current = state
    // Pivot definitions load eagerly so refresh (a synchronous apply step)
    // never waits on IPC. Best effort: a failed parse just disables refresh.
    for (const sheet of selected.sheets) {
      for (const pivot of sheet.pivotTables) {
        if (pivot.cachePath === null) continue
        void window.desktopApi
          .readPivotDefinition({
            sessionId: selected.sessionId,
            path: pivot.path,
            cachePath: pivot.cachePath,
          })
          .then((definition) => {
            if (lazyWorkbookRef.current === state) {
              state.pivotDefinitions.set(pivot.path, definition)
            }
          })
          .catch(() => undefined)
      }
    }
    // Dev-only diagnosis hooks: e2e drivers dump journal state and dispatch
    // Univer commands (drag interactions are hard to synthesize over CDP).
    if (import.meta.env.DEV) {
      ;(window as unknown as Record<string, unknown>).__journal = state.editJournal
      ;(window as unknown as Record<string, unknown>).__univerAPI = univerRef.current?.univerAPI
    }
    setRevision(0)
    setPreview(null)
    lazyPreviewRef.current = null
    setPendingEdits(0)
    // Slicers/timelines belong to the previous workbook's session only;
    // switching files invalidates them.
    setSlicers([])
    setSlicerPicker(null)
    setTimelines([])
    setTimelinePicker(null)
    disposeVisuals(visualDisposablesRef.current)
    loadWorkbookSkeleton(univerRef.current, selected)
    applyWorkbookNotes(univerRef.current, selected)
    applyDefinedNames(univerRef.current, selected, state)
    const runtime = univerRef.current
    if (runtime) {
      requestAnimationFrame(() => {
        const workbook = runtime.univerAPI.getActiveWorkbook()
        if (!workbook) return
        // Register existing file tables so Univer renders filter dropdowns
        // and banding. This is visual-only (the journal is empty for file
        // tables), so failures are swallowed — the data is still usable.
        for (const sheet of selected.sheets) {
          if (sheet.tables.length === 0) continue
          const ws = workbook.getSheetBySheetId(sheet.id)
          if (!ws) continue
          for (let index = 0; index < sheet.tables.length; index += 1) {
            const table = sheet.tables[index]!
            const tableId = `file-table-${sheet.id}-${index}`
            const tableName = `Table${index + 1}_${sheet.id.slice(0, 6)}`
            try {
              void ws.addTable(tableName, table.range, tableId)
            } catch {
              // Best-effort: skip if Univer rejects (e.g. overlapping ranges)
            }
          }
        }
        const worksheet = workbook.getActiveSheet()
        if (!worksheet) return
        // apply the opening sheet's formula view (sheetView/@showFormulas)
        applyShowFormulasView(runtime, state, worksheet.getSheetId())
        queueVisualInstall(
          runtime,
          lazyWorkbookRef,
          visualDisposablesRef,
          visualInstallTimerRef,
          worksheet.getSheetId(),
          chartEditRef,
          chartVectorRef,
          shapeEditRef,
        )
        try {
          worksheet.scrollToCell(0, 0)
        } catch {
          // A workbook opened during startup (the shell's queued-open nudge)
          // can land before Univer's Rendered lifecycle registers the scroll
          // render controller, and the facade then throws a redi
          // QuantityCheckError. The fresh view is already at the origin, so
          // skipping the reset is harmless.
        }
        void loadVisibleRange(runtime, lazyWorkbookRef, worksheet, setMessage)
        if (state.formulaMode) {
          void preloadEntireWorkbook(runtime, lazyWorkbookRef, setMessage)
        } else {
          // Deferred so first paint and initial streaming win the sidecar.
          setTimeout(() => {
            void activateFormulaClosure(runtime, lazyWorkbookRef, setMessage)
          }, 1500)
        }
      })
    }
  }

  async function handleInspectWorkbook(): Promise<void> {
    if (workbookOpeningRef.current) return
    workbookOpeningRef.current = true
    try {
      if (!window.desktopApi) {
        throw new Error(t('appBridgeUnavailable'))
      }
      const selected = await window.desktopApi.selectWorkbook()
      if (!selected) {
        setMessage(t('appOpenCanceled'))
        return
      }
      openLazyWorkbook(selected)
      setMessage(t('appOpened', { name: selected.name }))
    } catch (error: unknown) {
      setMessage(error instanceof Error ? error.message : t('appOpenFailed'))
    } finally {
      workbookOpeningRef.current = false
    }
  }

  async function handleSave(mode: 'save' | 'save-as' | 'recovery', quiet = false): Promise<void> {
    return handleSaveImpl(saveContext(), mode, quiet)
  }
  closeSaveRef.current = async () => {
    const state = lazyWorkbookRef.current
    if (!state || journalSize(state.editJournal) === 0) {
      window.desktopApi?.reportCloseSaveResult?.(true)
      return
    }
    await handleSave('save')
    // handleSave swallows errors into the status bar; a drained journal
    // (fresh state after openLazyWorkbook) is the success signal.
    const after = lazyWorkbookRef.current
    window.desktopApi?.reportCloseSaveResult?.(
      after === null || journalSize(after.editJournal) === 0,
    )
  }
  function sortColumnOptions(): { label: string; colIndex: number }[] {
    const range = univerRef.current?.univerAPI.getActiveWorkbook()?.getActiveRange()
    if (!range) return []
    const start = range.getColumn()
    const width = Math.min(range.getWidth(), 26)
    return Array.from({ length: width }, (_, offset) => ({
      label: t('appColumnLabel', { col: columnLabel(start + offset) }),
      colIndex: start + offset,
    }))
  }

  menuActionRef.current = (action) => {
    if (action === 'open') {
      void handleInspectWorkbook()
    } else if (action === 'export-pdf') {
      void handleExportPdfImpl(pageLayoutContext())
    } else if (action === 'undo' || action === 'redo') {
      // The shell's own text fields (AI prompt, dialog inputs) keep native
      // text undo; everywhere else ⌘Z means workbook history.
      const active = document.activeElement
      if (active instanceof HTMLTextAreaElement || active instanceof HTMLInputElement) {
        document.execCommand(action)
      } else if (action === 'undo') {
        void univerRef.current?.univerAPI.undo()
      } else {
        void univerRef.current?.univerAPI.redo()
      }
    } else {
      void handleSave(action)
    }
  }
  handleSaveRef.current = handleSave
  /// Re-renders the floating visuals after a journal mutation (edits and
  /// their undo/redo closures share it).
  function refreshLazyVisuals(state: LazyWorkbookState): void {
    const runtime = univerRef.current
    if (!runtime || lazyWorkbookRef.current !== state) return
    setPendingEdits(journalSize(state.editJournal))
    setVisualEditTick((tick) => tick + 1)
    const sheetId = runtime.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId()
    if (sheetId) {
      queueVisualInstall(
        runtime,
        lazyWorkbookRef,
        visualDisposablesRef,
        visualInstallTimerRef,
        sheetId,
        chartEditRef,
        chartVectorRef,
        shapeEditRef,
      )
    }
  }

  function refreshDemoVisuals(): void {
    if (lazyWorkbookRef.current) return
    setVisualEditTick((tick) => tick + 1)
    queueDemoVisualInstallForActiveSheet()
  }

  const chartSyncRef = useRef<{
    timer: ReturnType<typeof setTimeout> | null
    dirty: Map<string, CellBounds>
  }>({ timer: null, dirty: new Map() })

  function queueChartDataSync(sheetId: string, bounds: CellBounds): void {
    queueChartDataSyncImpl(visualSyncContext(), sheetId, bounds)
  }

  chartEditRef.current = (editKey, edit) => applyChartEditImpl(visualSyncContext(), editKey, edit)
  chartVectorRef.current = (editKey, rangeText) =>
    readChartVectorImpl(visualSyncContext(), editKey, rangeText)
  shapeEditRef.current = (visualId, changes) =>
    applyShapeEditImpl(visualSyncContext(), visualId, changes)

  // Ribbon echo of the selected chart; a live lookup so deletion, file
  // switches, and pending type edits reflect without extra bookkeeping.
  const selectedChart = (() => {
    if (!selectedVisual || selectedVisual.kind !== 'chart') return null
    const state = lazyWorkbookRef.current
    const live: WorkbookVisualObject | undefined = state
      ? [...state.file.visuals, ...state.editJournal.visualAdds].find(
          (candidate) => candidate.id === selectedVisual.id,
        )
      : (adapterRef.current.findVisual(selectedVisual.id) ?? undefined)
    if (!live?.chart || (state && state.editJournal.visualEdits.get(live.id)?.remove)) return null
    const pending =
      state && live.chartPath ? state.editJournal.chartEdits.get(live.chartPath) : undefined
    const currentChart = pending ? applyChartStateEdit(live.chart, pending) : live.chart
    const convertible = convertibleType(live.chart)
    const currentType = pending?.chartType ?? convertible
    const isPie =
      currentType !== null
        ? currentType === 'pie' || currentType === 'doughnut'
        : live.chart.chartTypes.some((type) => type.includes('pie') || type.includes('doughnut'))
    return {
      title: pending?.title ?? live.chart.title,
      convertible,
      currentType,
      canEdit: !state || live.chartPath !== undefined || live.id.startsWith('added-'),
      isPie,
      hasAxes: currentType !== null && !isPie,
      // A pending conversion always lands on a labelable family.
      canLabel: pending?.chartType !== undefined || chartSupportsDataLabels(live.chart.chartTypes),
      seriesCount: currentChart.series.length,
      categoryCount: currentChart.series[0]?.categories.length ?? 0,
      series: currentChart.series,
      legend: pending?.legend ?? live.chart.legend,
      axisTitles: { ...live.chart.axisTitles, ...pending?.axisTitles },
      dataLabels: pending?.dataLabels ?? live.chart.dataLabels,
      grouping: pending?.grouping ?? live.chart.grouping,
    }
  })()

  const activePageLayout = (() => {
    const worksheet = univerRef.current?.univerAPI.getActiveWorkbook()?.getActiveSheet()
    const journalState = worksheet
      ? (lazyWorkbookRef.current?.editJournal.pageSetup.get(worksheet.getSheetId()) ?? {})
      : {}
    return {
      ...journalState,
      showGridlines:
        journalState.showGridlines ?? (worksheet ? !worksheet.hasHiddenGridLines() : true),
    }
  })()

  // Chart panels resolve their chart live (pending edits applied), so every
  // control reflects the state the next save would write.
  const chartDialogTarget = (() => {
    if (!chartDialog) return null
    const state = lazyWorkbookRef.current
    const live: WorkbookVisualObject | undefined = state
      ? [...state.file.visuals, ...state.editJournal.visualAdds].find(
          (candidate) =>
            candidate.chartPath === chartDialog.editKey || candidate.id === chartDialog.editKey,
        )
      : (adapterRef.current.findVisual(chartDialog.editKey) ?? undefined)
    // A visual pending removal must not keep a live panel producing edits.
    if (!live?.chart || (state && state.editJournal.visualEdits.get(live.id)?.remove)) return null
    const pending =
      state && live.chartPath ? state.editJournal.chartEdits.get(live.chartPath) : undefined
    return {
      visualId: live.id,
      chart: applyChartStateEdit(live.chart, pending),
      supported: chartSupportsSeriesReplace(live.chart.chartTypes),
    }
  })()

  return (
    <>
      <ToastHost />
      {chartDialog && chartDialogTarget && chartDialog.kind === 'format' && (
        <ChartFormatPane
          chart={chartDialogTarget.chart}
          element={
            chartElement?.visualId === chartDialogTarget.visualId ? chartElement.element : null
          }
          onEdit={(edit) => chartEditRef.current(chartDialog.editKey, edit)}
          onClose={() => setChartDialog(null)}
        />
      )}
      {chartDialog && chartDialogTarget && chartDialog.kind === 'select-data' && (
        <SelectDataDialog
          chart={chartDialogTarget.chart}
          supported={chartDialogTarget.supported}
          readVector={(range) => chartVectorRef.current(chartDialog.editKey, range)}
          onApply={(edit) => chartEditRef.current(chartDialog.editKey, edit)}
          onClose={() => setChartDialog(null)}
        />
      )}
      <ExcelShell
        prompt={prompt}
        preview={preview}
        sheetHasContent={sheetHasContent}
        pageLayout={activePageLayout}
        selectionFormat={selectionFormat}
        statusMessage={message}
        aiBusy={aiBusy}
        chat={chat}
        historicChat={historicChat}
        attachments={attachments}
        attachNotice={attachNotice}
        onPickAttachments={() => void handlePickAttachments()}
        onAddAttachmentPaths={(paths) => void handleAddAttachmentPaths(paths)}
        onAddPastedImage={(data, ext) => void handleAddPastedImage(data, ext)}
        onRemoveAttachment={handleRemoveAttachment}
        onPromptChange={setPrompt}
        onSend={handleSend}
        onStop={handleStopAgent}
        onNewChat={handleNewChat}
        onUndo={handleUndo}
        onCommand={handleRibbonCommand}
        zoomPercent={zoomPercent}
        canSave={pendingEdits > 0}
        onSave={() => void handleSave('save')}
        onRedo={handleRedo}
        autoSave={autoSave}
        onAutoSaveChange={setAutoSave}
        selectedChart={selectedChart}
        onGetSortColumns={sortColumnOptions}
        onGetSheetProtection={sheetProtectionEcho}
        onGetDefinedNames={definedNameRows}
        onDefinedNameAction={handleDefinedNameAction}
        onGetPivotFields={() => pivotFieldOptionsImpl(pivotContext())}
        onGetSourceRange={() => getSourceRangeImpl(pivotContext())}
        onCreatePivot={(config) => handleCreatePivotImpl(pivotContext(), config)}
        onGetPivotEditSeed={() => pivotEditInitialImpl(pivotContext())}
        onEditPivot={(config) => handleEditPivotApplyImpl(pivotContext(), config)}
        onRefreshPivot={() => handleRefreshPivotImpl(pivotContext())}
        onIsSelectionInPivot={() => isSelectionInPivotImpl(pivotContext())}
        onGetActiveCell={() => activeCellLabelImpl(dataToolsContext())}
        onGetAnchorValue={anchorCellValue}
        activeCellA1={activeCellA1}
        onGoToReference={(ref) => goToReferenceImpl(dataToolsContext(), ref)}
        onListDefinedNames={() => listDefinedNamesImpl(dataToolsContext())}
        onApplyFormula={(formula) => handleApplyFormulaImpl(dataToolsContext(), formula)}
        onCreateSubtotal={(config) => handleCreateSubtotalImpl(dataToolsContext(), config)}
        onCreateConsolidate={(config) => handleCreateConsolidateImpl(dataToolsContext(), config)}
        onGetConsolidateDefault={() => consolidateDefaultReferenceImpl(dataToolsContext())}
        onApplyHeaderFooter={(result) => handleApplyHeaderFooterImpl(pageLayoutContext(), result)}
      />
      {advancedFilterColumns !== null && (
        <AdvancedFilterDialog
          columns={advancedFilterColumns}
          onApply={(criteria) => handleApplyAdvancedFilterImpl(dataToolsContext(), criteria)}
          onClose={() => setAdvancedFilterColumns(null)}
        />
      )}
      {symbolDialogOpen && (
        <SymbolDialog
          onInsert={(char) => handleInsertSymbolImpl(dataToolsContext(), char)}
          onClose={() => setSymbolDialogOpen(false)}
        />
      )}
      {screenshotDialogOpen && (
        <ScreenshotDialog
          onInsert={(dataUrl, width, height) =>
            handleInsertScreenshot(visualContext(), dataUrl, width, height)
          }
          onClose={() => setScreenshotDialogOpen(false)}
        />
      )}
      {iconsDialogOpen && (
        <IconsDialog
          onInsert={(dataUrl, size, name) =>
            handleInsertIconImpl(visualContext(), dataUrl, size, name)
          }
          onClose={() => setIconsDialogOpen(false)}
        />
      )}
      {equationDialogOpen && (
        <EquationDialog
          onInsert={(dataUrl, width, height) =>
            handleInsertEquationImpl(visualContext(), dataUrl, width, height)
          }
          onClose={() => setEquationDialogOpen(false)}
        />
      )}
      {recommendedCharts !== null && (
        <RecommendedChartsDialog
          recommendations={recommendedCharts}
          onPick={(kind) => void handleInsertChartImpl(visualContext(), kind)}
          onClose={() => setRecommendedCharts(null)}
        />
      )}
      {slicerPicker !== null && (
        <SlicerFieldPicker
          fields={slicerPicker.fields}
          onPick={(field) => handleCreateSlicerImpl(pivotContext(), field)}
          onClose={() => setSlicerPicker(null)}
        />
      )}
      <SlicerPanels
        slicers={slicers}
        onToggle={(slicerId, member) => handleSlicerToggleImpl(pivotContext(), slicerId, member)}
        onSelectAll={(slicerId) => handleSlicerSelectAllImpl(pivotContext(), slicerId)}
        onRemove={(slicerId) => handleRemoveSlicerImpl(pivotContext(), slicerId)}
      />
      {timelinePicker !== null && (
        <TimelineFieldPicker
          fields={timelinePicker.fields}
          onPick={(field) => handleCreateTimelineImpl(pivotContext(), field)}
          onClose={() => setTimelinePicker(null)}
        />
      )}
      <TimelinePanels
        timelines={timelines}
        onRange={(timelineId, start, end) =>
          handleTimelineRangeImpl(pivotContext(), timelineId, { start, end })
        }
        onClear={(timelineId) => handleTimelineRangeImpl(pivotContext(), timelineId, null)}
        onRemove={(timelineId) => handleRemoveTimelineImpl(pivotContext(), timelineId)}
      />
    </>
  )

  function definedNameRows(): {
    names: DefinedNameRow[]
    sheets: { id: string; name: string }[]
  } {
    const workbook = univerRef.current?.univerAPI.getActiveWorkbook()
    const sheets =
      workbook?.getSheets().map((sheet) => ({
        id: sheet.getSheetId(),
        name: sheet.getSheetName(),
      })) ?? []
    const sheetNames = new Map(sheets.map((sheet) => [sheet.id, sheet.name]))
    const names = univerDefinedNames(univerRef.current).map((defined) => {
      const localSheetId = defined.getLocalSheetId()
      const scoped = localSheetId !== undefined && localSheetId !== 'AllDefaultWorkbook'
      return {
        name: defined.getName(),
        ref: defined.getFormulaOrRefString(),
        scopeSheetId: scoped ? localSheetId : null,
        scopeLabel: scoped ? (sheetNames.get(localSheetId) ?? localSheetId) : t('appScopeWorkbook'),
      }
    })
    return { names, sheets }
  }

  function handleDefinedNameAction(action: DefinedNameAction): string | null {
    const runtime = univerRef.current
    const workbook = runtime?.univerAPI.getActiveWorkbook()
    if (!workbook || !lazyWorkbookRef.current) {
      return t('appNamesNeedFile')
    }
    try {
      if (action.kind === 'add') {
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
              name: action.name,
              formulaOrRefString: action.ref.replace(/^=/, ''),
              localSheetId: action.sheetId ?? 'AllDefaultWorkbook',
            })
            .build(),
        )
      } else {
        const target = univerDefinedNames(runtime).find((defined) => {
          const localSheetId = defined.getLocalSheetId()
          const scoped = localSheetId !== undefined && localSheetId !== 'AllDefaultWorkbook'
          const scopeSheetId = scoped ? localSheetId : null
          const originalName = action.kind === 'update' ? action.originalName : action.name
          return defined.getName() === originalName && scopeSheetId === action.scopeSheetId
        })
        if (!target) return t('appNameGone')
        if (action.kind === 'remove') {
          target.delete()
        } else {
          if (action.name !== action.originalName) target.setName(action.name)
          target.setRef(action.ref.replace(/^=/, ''))
        }
      }
    } catch (error: unknown) {
      return error instanceof Error ? error.message : t('appNameApplyFailed')
    }
    setMessage(t('appNamesUpdated'))
    return null
  }

  /// Effective protection of the active sheet: journal override, else file
  /// state; null while unknown (still indexing) or in the demo workbook.
  function sheetProtectionEcho(): boolean | null {
    const state = lazyWorkbookRef.current
    const sheetId = univerRef.current?.univerAPI.getActiveWorkbook()?.getActiveSheet()?.getSheetId()
    if (!state || !sheetId) return null
    const journaled = state.editJournal.sheetProtection.get(sheetId)
    if (journaled !== undefined) return journaled
    const file = state.sheetProtections.get(sheetId)
    if (file) return file.protected
    return state.editJournal.sheets.added.has(sheetId) ? false : null
  }
}
